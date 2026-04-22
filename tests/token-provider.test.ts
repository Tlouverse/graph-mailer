import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { GraphAuthError } from '../src/errors.js';
import { TokenProvider } from '../src/auth/token-provider.js';

const TOKEN_RESPONSE = {
  access_token: 'test-token-123',
  expires_in: 3600,
};

function mockFetch(response: { ok: boolean; status: number; body: string }) {
  return vi.fn().mockResolvedValue({
    ok: response.ok,
    status: response.status,
    text: () => Promise.resolve(response.body),
  });
}

beforeEach(() => {
  vi.useFakeTimers();
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.useRealTimers();
});

describe('TokenProvider', () => {
  it('fetches and returns a token', async () => {
    vi.stubGlobal('fetch', mockFetch({ ok: true, status: 200, body: JSON.stringify(TOKEN_RESPONSE) }));

    const provider = new TokenProvider('tenant', 'client', 'secret');
    const token = await provider.getToken();

    expect(token).toBe('test-token-123');
    expect(fetch).toHaveBeenCalledOnce();
  });

  it('caches the token and does not re-fetch while valid', async () => {
    vi.stubGlobal('fetch', mockFetch({ ok: true, status: 200, body: JSON.stringify(TOKEN_RESPONSE) }));

    const provider = new TokenProvider('tenant', 'client', 'secret');
    await provider.getToken();
    await provider.getToken();

    expect(fetch).toHaveBeenCalledOnce();
  });

  it('refreshes the token after the 60s expiry window', async () => {
    vi.stubGlobal('fetch', mockFetch({ ok: true, status: 200, body: JSON.stringify(TOKEN_RESPONSE) }));

    const provider = new TokenProvider('tenant', 'client', 'secret');
    await provider.getToken();

    // advance past expires_in - 60 seconds
    vi.advanceTimersByTime((TOKEN_RESPONSE.expires_in - 60 + 1) * 1000);

    await provider.getToken();

    expect(fetch).toHaveBeenCalledTimes(2);
  });

  it('deduplicates concurrent token requests into a single fetch', async () => {
    vi.stubGlobal('fetch', mockFetch({ ok: true, status: 200, body: JSON.stringify(TOKEN_RESPONSE) }));

    const provider = new TokenProvider('tenant', 'client', 'secret');
    const [t1, t2, t3] = await Promise.all([
      provider.getToken(),
      provider.getToken(),
      provider.getToken(),
    ]);

    expect(fetch).toHaveBeenCalledOnce();
    expect(t1).toBe('test-token-123');
    expect(t2).toBe('test-token-123');
    expect(t3).toBe('test-token-123');
  });

  it('throws GraphAuthError on HTTP error', async () => {
    vi.stubGlobal(
      'fetch',
      mockFetch({ ok: false, status: 401, body: '{"error":"invalid_client"}' }),
    );

    const provider = new TokenProvider('tenant', 'client', 'secret');
    await expect(provider.getToken()).rejects.toThrow(GraphAuthError);

    try {
      const provider2 = new TokenProvider('tenant', 'client', 'secret');
      await provider2.getToken();
    } catch (err) {
      expect(err).toBeInstanceOf(GraphAuthError);
      const authErr = err as GraphAuthError;
      expect(authErr.details?.status).toBe(401);
      expect(authErr.details?.body).toBe('{"error":"invalid_client"}');
    }
  });
});

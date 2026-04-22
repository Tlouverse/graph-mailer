import { beforeEach, describe, expect, it, vi } from 'vitest';
import { GraphMailError } from '../src/errors.js';
import { GraphMailer } from '../src/mailer.js';

const { postMock, apiMock } = vi.hoisted(() => {
  const postMock = vi.fn().mockResolvedValue(undefined);
  const apiMock = vi.fn().mockReturnValue({ post: postMock });
  return { postMock, apiMock };
});

vi.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    initWithMiddleware: vi.fn().mockReturnValue({ api: apiMock }),
  },
}));

vi.mock('../src/auth/token-provider.js', () => ({
  TokenProvider: vi.fn().mockImplementation(() => ({
    getToken: vi.fn().mockResolvedValue('mock-token'),
  })),
}));

beforeEach(() => {
  vi.clearAllMocks();
  postMock.mockResolvedValue(undefined);
  apiMock.mockReturnValue({ post: postMock });
});

describe('GraphMailer.send', () => {
  it('uses options.from when provided', async () => {
    const mailer = new GraphMailer({ tenantId: 't', clientId: 'c', clientSecret: 's' });
    await mailer.send({ from: 'sender@example.com', to: 'a@example.com', html: 'x', subject: 'Hi' });

    expect(apiMock).toHaveBeenCalledWith('/users/sender%40example.com/sendMail');
  });

  it('falls back to defaultFrom when from is not in options', async () => {
    const mailer = new GraphMailer({
      tenantId: 't',
      clientId: 'c',
      clientSecret: 's',
      defaultFrom: 'default@example.com',
    });
    await mailer.send({ to: 'a@example.com', html: 'x', subject: 'Hi' });

    expect(apiMock).toHaveBeenCalledWith('/users/default%40example.com/sendMail');
  });

  it('throws GraphMailError when no from is available', async () => {
    const mailer = new GraphMailer({ tenantId: 't', clientId: 'c', clientSecret: 's' });
    await expect(mailer.send({ to: 'a@example.com', html: 'x', subject: 'Hi' })).rejects.toThrow(
      GraphMailError,
    );
  });

  it('sends with saveToSentItems false by default', async () => {
    const mailer = new GraphMailer({
      tenantId: 't',
      clientId: 'c',
      clientSecret: 's',
      defaultFrom: 'sender@example.com',
    });
    await mailer.send({ to: 'a@example.com', html: 'x', subject: 'Hi' });

    expect(postMock).toHaveBeenCalledWith(expect.objectContaining({ saveToSentItems: false }));
  });

  it('allows overriding saveToSentItems to true', async () => {
    const mailer = new GraphMailer({
      tenantId: 't',
      clientId: 'c',
      clientSecret: 's',
      defaultFrom: 'sender@example.com',
    });
    await mailer.send({ to: 'a@example.com', html: 'x', subject: 'Hi', saveToSentItems: true });

    expect(postMock).toHaveBeenCalledWith(expect.objectContaining({ saveToSentItems: true }));
  });

  it('wraps SDK errors in GraphMailError', async () => {
    const sdkError = Object.assign(new Error('Graph error'), {
      statusCode: 403,
      body: JSON.stringify({ error: { code: 'Forbidden' } }),
    });
    postMock.mockRejectedValueOnce(sdkError);

    const mailer = new GraphMailer({
      tenantId: 't',
      clientId: 'c',
      clientSecret: 's',
      defaultFrom: 'sender@example.com',
    });

    await expect(mailer.send({ to: 'a@example.com', html: 'x', subject: 'Hi' })).rejects.toThrow(
      GraphMailError,
    );
  });
});

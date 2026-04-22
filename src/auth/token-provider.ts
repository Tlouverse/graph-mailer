import { GraphAuthError } from '../errors.js';

interface TokenCache {
  accessToken: string;
  expiresAt: number;
}

export class TokenProvider {
  private cache: TokenCache | null = null;
  private inFlight?: Promise<string>;

  constructor(
    private readonly tenantId: string,
    private readonly clientId: string,
    private readonly clientSecret: string,
  ) {}

  async getToken(): Promise<string> {
    if (this.cache && Date.now() < this.cache.expiresAt) {
      return this.cache.accessToken;
    }

    this.inFlight ??= this.fetchToken().finally(() => {
      this.inFlight = undefined;
    });

    return this.inFlight;
  }

  private async fetchToken(): Promise<string> {
    const url = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: this.clientId,
      client_secret: this.clientSecret,
      scope: 'https://graph.microsoft.com/.default',
    });

    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });

    const responseBody = await response.text();

    if (!response.ok) {
      throw new GraphAuthError(
        `Token request failed with status ${response.status}`,
        { status: response.status, body: responseBody },
      );
    }

    const json = JSON.parse(responseBody) as { access_token: string; expires_in: number };
    const expiresAt = Date.now() + (json.expires_in - 60) * 1000;

    this.cache = { accessToken: json.access_token, expiresAt };
    return json.access_token;
  }
}

import type { AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import type { TokenProvider } from './token-provider.js';

export class ClientCredentialsAuthProvider implements AuthenticationProvider {
  constructor(private readonly tokenProvider: TokenProvider) {}

  async getAccessToken(): Promise<string> {
    return this.tokenProvider.getToken();
  }
}

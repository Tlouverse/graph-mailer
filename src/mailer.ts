import { Client } from '@microsoft/microsoft-graph-client';
import { ClientCredentialsAuthProvider } from './auth/auth-provider.js';
import { TokenProvider } from './auth/token-provider.js';
import { GraphMailError } from './errors.js';
import { buildMessage } from './payload.js';
import type { GraphMailerConfig, SendMailOptions } from './types.js';

export class GraphMailer {
  private readonly client: Client;
  private readonly defaultFrom?: string;

  constructor(config: GraphMailerConfig) {
    const tokenProvider = new TokenProvider(config.tenantId, config.clientId, config.clientSecret);
    const authProvider = new ClientCredentialsAuthProvider(tokenProvider);
    this.client = Client.initWithMiddleware({ authProvider });
    this.defaultFrom = config.defaultFrom;
  }

  async send(options: SendMailOptions): Promise<void> {
    const from = options.from ?? this.defaultFrom;
    if (!from) {
      throw new GraphMailError(
        'No sender specified: provide `from` in send() or `defaultFrom` in GraphMailerConfig',
      );
    }

    const message = buildMessage(options);
    const saveToSentItems = options.saveToSentItems ?? false;

    const endpoint = `/users/${encodeURIComponent(from)}/sendMail`;

    try {
      await this.client.api(endpoint).post({ message, saveToSentItems });
    } catch (err) {
      throw GraphMailError.fromGraphError(err);
    }
  }
}

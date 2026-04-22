function asObject(val: unknown): Record<string, unknown> | undefined {
  return val !== null && typeof val === 'object' ? (val as Record<string, unknown>) : undefined;
}

function tryParseJson(val: unknown): Record<string, unknown> | undefined {
  if (typeof val !== 'string') return undefined;
  try {
    return asObject(JSON.parse(val));
  } catch {
    return undefined;
  }
}

export class GraphMailError extends Error {
  readonly name = 'GraphMailError';

  constructor(
    message: string,
    readonly details?: {
      status?: number;
      code?: string;
      requestId?: string;
      raw?: unknown;
    },
  ) {
    super(message);
    Object.setPrototypeOf(this, new.target.prototype);
  }

  static fromGraphError(err: unknown): GraphMailError {
    const e = asObject(err);
    if (!e) {
      return new GraphMailError('Microsoft Graph request failed', { raw: err });
    }

    const message = typeof e['message'] === 'string' ? e['message'] : 'Microsoft Graph request failed';
    const status = typeof e['statusCode'] === 'number' ? e['statusCode'] : undefined;

    const body = asObject(e['body']) ?? tryParseJson(e['body']);
    const code = typeof body?.['error'] === 'object'
      ? (asObject(body['error'])?.['code'] as string | undefined)
      : undefined;

    const headers = asObject(e['headers']);
    const requestId = typeof headers?.['request-id'] === 'string' ? headers['request-id'] : undefined;

    return new GraphMailError(message, { status, code, requestId, raw: err });
  }
}

export class GraphAuthError extends Error {
  readonly name = 'GraphAuthError';

  constructor(
    message: string,
    readonly details?: { status?: number; body?: string },
  ) {
    super(message);
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

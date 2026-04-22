/**
 * An email address, either as a plain string or an object with an optional display name.
 *
 * @example
 * 'alice@example.com'
 * { email: 'alice@example.com', name: 'Alice' }
 */
export type Address = string | { email: string; name?: string };

/** Credentials and defaults passed once when constructing a {@link GraphMailer}. */
export interface GraphMailerConfig {
  /** Azure AD / Entra ID tenant ID (GUID). */
  tenantId: string;

  /** Client ID of the app registration that holds the `Mail.Send` application permission. */
  clientId: string;

  /** Client secret associated with the app registration. */
  clientSecret: string;

  /**
   * Fallback sender address used when `from` is omitted in {@link SendMailOptions}.
   * Must be a mailbox the app is authorised to send from.
   */
  defaultFrom?: string;
}

/**
 * A file to attach to the email.
 *
 * @remarks
 * The Graph API supports inline attachments up to 4 MB. For larger files
 * an upload session is required, which is not handled by this package.
 */
export interface Attachment {
  /** Filename as it will appear in the email (e.g. `'report.pdf'`). */
  name: string;

  /** MIME type of the file (e.g. `'application/pdf'`, `'image/png'`). */
  contentType: string;

  /** File content as a `Buffer` or a base64-encoded string. */
  content: Buffer | string;
}

/** Options passed to {@link GraphMailer.send} for a single email. */
export interface SendMailOptions {
  /**
   * Sender address for this message.
   * Overrides `defaultFrom` when provided; required if `defaultFrom` is not set on the mailer.
   */
  from?: string;

  /** Primary recipient(s). Accepts a single address or an array of addresses. */
  to: Address | Address[];

  /** Carbon-copy recipient(s). */
  cc?: Address | Address[];

  /** Blind carbon-copy recipient(s). */
  bcc?: Address | Address[];

  /** Reply-to address(es). Defaults to the sender when omitted. */
  replyTo?: Address | Address[];

  /** Email subject line. */
  subject: string;

  /**
   * HTML body. Takes priority over `text` when both are supplied.
   * At least one of `html` or `text` must be provided.
   */
  html?: string;

  /**
   * Plain-text body. Used only when `html` is not provided.
   * At least one of `html` or `text` must be provided.
   */
  text?: string;

  /**
   * Whether to save the sent message in the sender's Sent Items folder.
   * @default false
   */
  saveToSentItems?: boolean;

  /** Files to attach to the email. Each attachment must be under 4 MB. */
  attachments?: Attachment[];
}

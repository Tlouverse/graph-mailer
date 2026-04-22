import type { Message } from '@microsoft/microsoft-graph-types';
import { GraphMailError } from './errors.js';
import type { Address, Attachment, SendMailOptions } from './types.js';

function toRecipients(
  input: Address | Address[] | undefined,
): { emailAddress: { address: string; name?: string } }[] | undefined {
  if (input === undefined) return undefined;

  const arr = Array.isArray(input) ? input : [input];
  return arr.map((a) => {
    if (typeof a === 'string') {
      return { emailAddress: { address: a } };
    }
    return { emailAddress: { address: a.email, ...(a.name ? { name: a.name } : {}) } };
  });
}

export function buildMessage(options: SendMailOptions): Message {
  if (!options.html && !options.text) {
    throw new GraphMailError('Either html or text must be provided');
  }

  const body: Message['body'] = options.html
    ? { contentType: 'html', content: options.html }
    : { contentType: 'text', content: options.text! };

  const message: Message = {
    subject: options.subject,
    body,
    toRecipients: toRecipients(options.to),
  };

  const cc = toRecipients(options.cc);
  if (cc) message.ccRecipients = cc;

  const bcc = toRecipients(options.bcc);
  if (bcc) message.bccRecipients = bcc;

  const replyTo = toRecipients(options.replyTo);
  if (replyTo) message.replyTo = replyTo;

  if (options.attachments?.length) {
    message.attachments = options.attachments.map((a: Attachment) => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.name,
      contentType: a.contentType,
      contentBytes: Buffer.isBuffer(a.content) ? a.content.toString('base64') : a.content,
    }));
  }

  return message;
}

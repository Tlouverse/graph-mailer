import { describe, expect, it } from 'vitest';
import { GraphMailError } from '../src/errors.js';
import { buildMessage } from '../src/payload.js';

describe('buildMessage', () => {
  it('builds a message with HTML body', () => {
    const msg = buildMessage({ subject: 'Hi', to: 'a@example.com', html: '<p>Hello</p>' });
    expect(msg.body).toEqual({ contentType: 'html', content: '<p>Hello</p>' });
  });

  it('builds a message with text body', () => {
    const msg = buildMessage({ subject: 'Hi', to: 'a@example.com', text: 'Hello' });
    expect(msg.body).toEqual({ contentType: 'text', content: 'Hello' });
  });

  it('prefers HTML when both html and text are provided', () => {
    const msg = buildMessage({ subject: 'Hi', to: 'a@example.com', html: '<b>Hi</b>', text: 'Hi' });
    expect(msg.body?.contentType).toBe('html');
    expect(msg.body?.content).toBe('<b>Hi</b>');
  });

  it('throws GraphMailError when neither html nor text is provided', () => {
    expect(() => buildMessage({ subject: 'Hi', to: 'a@example.com' })).toThrow(GraphMailError);
  });

  it('maps a string recipient', () => {
    const msg = buildMessage({ subject: 'Hi', to: 'a@example.com', html: 'x' });
    expect(msg.toRecipients).toEqual([{ emailAddress: { address: 'a@example.com' } }]);
  });

  it('maps an object recipient with name', () => {
    const msg = buildMessage({
      subject: 'Hi',
      to: { email: 'b@example.com', name: 'Bob' },
      html: 'x',
    });
    expect(msg.toRecipients).toEqual([{ emailAddress: { address: 'b@example.com', name: 'Bob' } }]);
  });

  it('maps a mixed array of recipients', () => {
    const msg = buildMessage({
      subject: 'Hi',
      to: ['a@example.com', { email: 'b@example.com', name: 'Bob' }],
      html: 'x',
    });
    expect(msg.toRecipients).toEqual([
      { emailAddress: { address: 'a@example.com' } },
      { emailAddress: { address: 'b@example.com', name: 'Bob' } },
    ]);
  });

  it('omits ccRecipients when cc is not provided', () => {
    const msg = buildMessage({ subject: 'Hi', to: 'a@example.com', html: 'x' });
    expect(msg.ccRecipients).toBeUndefined();
  });

  it('omits bccRecipients when bcc is not provided', () => {
    const msg = buildMessage({ subject: 'Hi', to: 'a@example.com', html: 'x' });
    expect(msg.bccRecipients).toBeUndefined();
  });

  it('omits replyTo when replyTo is not provided', () => {
    const msg = buildMessage({ subject: 'Hi', to: 'a@example.com', html: 'x' });
    expect(msg.replyTo).toBeUndefined();
  });

  it('maps cc, bcc, and replyTo when provided', () => {
    const msg = buildMessage({
      subject: 'Hi',
      to: 'a@example.com',
      cc: 'c@example.com',
      bcc: 'b@example.com',
      replyTo: 'r@example.com',
      html: 'x',
    });
    expect(msg.ccRecipients).toEqual([{ emailAddress: { address: 'c@example.com' } }]);
    expect(msg.bccRecipients).toEqual([{ emailAddress: { address: 'b@example.com' } }]);
    expect(msg.replyTo).toEqual([{ emailAddress: { address: 'r@example.com' } }]);
  });
});

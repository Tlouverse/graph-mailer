# graph-mailer — Agent & contributor guide

## Commands

```bash
npm run build        # compile to dist/ (ESM + CJS) via tsup
npm run typecheck    # tsc --noEmit, no emit
npm test             # vitest run
```

Always run `typecheck` and `test` before considering a change done.

## Structure

```
src/
├── index.ts              # public exports only — nothing else goes here
├── types.ts              # public types (Address, GraphMailerConfig, SendMailOptions)
├── errors.ts             # GraphMailError, GraphAuthError + private helpers
├── mailer.ts             # GraphMailer class
├── payload.ts            # buildMessage() — SendMailOptions → Graph Message
└── auth/
    ├── token-provider.ts # token fetch + in-memory cache + deduplication
    └── auth-provider.ts  # AuthenticationProvider adapter for the Graph SDK

tests/
├── token-provider.test.ts
├── payload.test.ts
└── mailer.test.ts
```

## Invariants to preserve

- **Zero `process.env` reads.** The consumer passes everything explicitly.
- **`src/index.ts` exports exactly:** `GraphMailer`, `GraphMailerConfig`, `SendMailOptions`, `Address`, `Attachment`, `GraphMailError`, `GraphAuthError`. No internals.
- **Single runtime dependency:** `@microsoft/microsoft-graph-client`. Do not add others without discussion.
- **Node 18+** — native `fetch` and `URLSearchParams` are available, no polyfills.
- **Token deduplication** in `TokenProvider.getToken()` — concurrent calls must share one in-flight promise. Do not break the `??=` pattern.
- **No email validation** in `buildMessage()` — no regex, no split on `,`/`;`. The consumer is responsible for clean data.

## Adding a feature

1. Types first — update `types.ts` if the public API changes.
2. Implementation — keep internal helpers private (not exported from `index.ts`).
3. Tests — cover the new behaviour in the relevant test file. Mock `fetch` globally for `TokenProvider`, mock `Client` from the SDK for `GraphMailer`.
4. JSDoc — public types and their properties must have JSDoc. Implementation files do not need comments unless the why is non-obvious.
5. README — update both the EN and FR sections if the public API or behaviour changed.

## What is out of scope (V1)

Do not implement without explicit instruction:

- Attachments larger than 4 MB (requires upload session)
- Custom `internetMessageHeaders`
- Retry logic (the SDK middleware handles basics)
- Injectable logger
- Framework integrations (NestJS module, Next.js helper)
- Non-client-credentials auth strategies
- Draft creation (`/createDraft`)

## Publishing

```bash
npm version patch   # or minor / major
npm run build
npm publish         # runs prepublishOnly = build + test
```

The package is published as `@tlouverse/graph-mailer` on the npm public registry under the `tlouverse` org.

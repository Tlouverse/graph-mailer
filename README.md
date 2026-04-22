# @tlouverse/graph-mailer

[![npm version](https://img.shields.io/npm/v/@tlouverse/graph-mailer)](https://www.npmjs.com/package/@tlouverse/graph-mailer)
[![license](https://img.shields.io/npm/l/@tlouverse/graph-mailer)](./LICENSE)

**[English](#english) · [Français](#français)**

---

## English

Send emails through the Microsoft Graph API using app-only (client credentials) authentication.

### Installation

```bash
npm install @tlouverse/graph-mailer
```

### Quick start

```ts
import { GraphMailer } from '@tlouverse/graph-mailer';

const mailer = new GraphMailer({
  tenantId: 'your-tenant-id',
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret',
  defaultFrom: 'noreply@example.com',
});

await mailer.send({
  to: ['alice@example.com', { email: 'bob@example.com', name: 'Bob' }],
  cc: 'carol@example.com',
  subject: 'Hello',
  html: '<p>Hi there!</p>',
});
```

Reuse the same `GraphMailer` instance across your application — it handles token caching and renewal automatically.

### API

#### `new GraphMailer(config)`

| Option | Type | Required | Description |
|---|---|---|---|
| `tenantId` | `string` | ✓ | Azure AD / Entra ID tenant GUID |
| `clientId` | `string` | ✓ | App registration client ID |
| `clientSecret` | `string` | ✓ | App registration client secret |
| `defaultFrom` | `string` | — | Fallback sender address |

#### `mailer.send(options)` → `Promise<void>`

| Option | Type | Required | Description |
|---|---|---|---|
| `to` | `Address \| Address[]` | ✓ | Primary recipient(s) |
| `subject` | `string` | ✓ | Subject line |
| `html` | `string` | ✓* | HTML body |
| `text` | `string` | ✓* | Plain-text body |
| `from` | `string` | — | Sender (overrides `defaultFrom`) |
| `cc` | `Address \| Address[]` | — | CC recipient(s) |
| `bcc` | `Address \| Address[]` | — | BCC recipient(s) |
| `replyTo` | `Address \| Address[]` | — | Reply-to address(es) |
| `saveToSentItems` | `boolean` | — | Default: `false` |
| `attachments` | `Attachment[]` | — | Files to attach (max 4 MB each) |

*At least one of `html` or `text` is required. `html` takes priority when both are provided.

#### `Address`

```ts
type Address = string | { email: string; name?: string };
```

#### `Attachment`

```ts
interface Attachment {
  name: string;        // filename as it appears in the email
  contentType: string; // MIME type, e.g. 'application/pdf'
  content: Buffer | string; // file content as a Buffer or base64 string
}
```

#### Errors

| Class | Thrown when |
|---|---|
| `GraphMailError` | Missing sender, missing body, or Graph API error |
| `GraphAuthError` | Token acquisition failed (bad credentials, network issue) |

Both extend `Error` and expose a `details` property with the HTTP status and error context for debugging.

### Azure prerequisites

1. **App registration** — Create an app in [Entra ID](https://entra.microsoft.com).
2. **API permission** — Add the **`Mail.Send`** Microsoft Graph permission as an **Application permission** (not Delegated), then grant admin consent.
3. **Application Access Policy** — Without this, your app can send from *any* mailbox in the tenant. Restrict it with Exchange Online PowerShell:

   ```powershell
   New-ApplicationAccessPolicy `
     -AppId <clientId> `
     -PolicyScopeGroupId <mailEnabledGroupOrMailbox> `
     -AccessRight RestrictAccess `
     -Description "Restrict app to authorised mailboxes"
   ```

   > ⚠️ This step is critical. Skip it and the app has unrestricted send-as access across your entire tenant.

---

## Français

Envoi d'emails via l'API Microsoft Graph avec une authentification applicative (client credentials).

### Installation

```bash
npm install @tlouverse/graph-mailer
```

### Démarrage rapide

```ts
import { GraphMailer } from '@tlouverse/graph-mailer';

const mailer = new GraphMailer({
  tenantId: 'votre-tenant-id',
  clientId: 'votre-client-id',
  clientSecret: 'votre-client-secret',
  defaultFrom: 'noreply@example.com',
});

await mailer.send({
  to: ['alice@example.com', { email: 'bob@example.com', name: 'Bob' }],
  cc: 'carol@example.com',
  subject: 'Bonjour',
  html: '<p>Bonjour à tous !</p>',
});
```

Réutilisez la même instance `GraphMailer` dans toute votre application — elle gère le cache et le renouvellement du token automatiquement.

### API

#### `new GraphMailer(config)`

| Option | Type | Requis | Description |
|---|---|---|---|
| `tenantId` | `string` | ✓ | GUID du tenant Azure AD / Entra ID |
| `clientId` | `string` | ✓ | Client ID de l'inscription d'application |
| `clientSecret` | `string` | ✓ | Secret client de l'inscription d'application |
| `defaultFrom` | `string` | — | Adresse expéditeur par défaut |

#### `mailer.send(options)` → `Promise<void>`

| Option | Type | Requis | Description |
|---|---|---|---|
| `to` | `Address \| Address[]` | ✓ | Destinataire(s) principal(aux) |
| `subject` | `string` | ✓ | Objet du message |
| `html` | `string` | ✓* | Corps HTML |
| `text` | `string` | ✓* | Corps en texte brut |
| `from` | `string` | — | Expéditeur (remplace `defaultFrom`) |
| `cc` | `Address \| Address[]` | — | Destinataire(s) en copie |
| `bcc` | `Address \| Address[]` | — | Destinataire(s) en copie cachée |
| `replyTo` | `Address \| Address[]` | — | Adresse(s) de réponse |
| `saveToSentItems` | `boolean` | — | Défaut : `false` |
| `attachments` | `Attachment[]` | — | Fichiers joints (4 Mo max chacun) |

*Au moins `html` ou `text` est requis. `html` est prioritaire si les deux sont fournis.

#### `Address`

```ts
type Address = string | { email: string; name?: string };
```

#### `Attachment`

```ts
interface Attachment {
  name: string;        // nom du fichier tel qu'il apparaît dans l'email
  contentType: string; // type MIME, ex. 'application/pdf'
  content: Buffer | string; // contenu en Buffer ou en chaîne base64
}
```

#### Erreurs

| Classe | Levée quand |
|---|---|
| `GraphMailError` | Expéditeur manquant, corps manquant, ou erreur de l'API Graph |
| `GraphAuthError` | Échec d'acquisition du token (mauvaises credentials, problème réseau) |

Les deux étendent `Error` et exposent une propriété `details` contenant le status HTTP et le contexte de l'erreur pour faciliter le débogage.

### Prérequis Azure

1. **Inscription d'application** — Créer une application dans [Entra ID](https://entra.microsoft.com).
2. **Permission API** — Ajouter la permission **`Mail.Send`** de Microsoft Graph en tant que permission **Application** (et non Déléguée), puis accorder le consentement administrateur.
3. **Application Access Policy** — Sans cette étape, l'application peut envoyer depuis *n'importe quelle* boîte aux lettres du tenant. Restreindre via Exchange Online PowerShell :

   ```powershell
   New-ApplicationAccessPolicy `
     -AppId <clientId> `
     -PolicyScopeGroupId <groupeOuBoiteAutorisé> `
     -AccessRight RestrictAccess `
     -Description "Restreindre l'application aux boîtes autorisées"
   ```

   > ⚠️ Cette étape est critique. Sans elle, l'application dispose d'un accès en envoi illimité à l'ensemble des boîtes du tenant.

---

## License

MIT

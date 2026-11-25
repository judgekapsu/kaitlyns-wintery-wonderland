# Kaitlyn's Wintery Wonderland

A family wishlist web app using Google Apps Script and a private Google Sheet. This version supports shared login, wishlist submissions, and browsing only. Price tracking, email alerts, and purchase actions are not implemented.

## Files

```text
apps-script/       Apps Script backend, manifest, and HTML website
 tools/             Local password helper
.clasp.json.example clasp configuration template
README.md           Setup guide
```

The Apps Script project should be bound to the target spreadsheet. `clasp` uses `apps-script` as its root and pushes the manifest and app files from that directory.

## Configure privately

Create a spreadsheet-bound Apps Script project by opening the Sheet and choosing **Extensions → Apps Script**. Configure these property names under **Project Settings → Script Properties**; never put the values in GitHub, source files, or the Sheet:

- `FAMILY_USERNAME`: choose your own shared username.
- `FAMILY_PASSWORD_HASH`: generated verifier record, steps below.
- `SPREADSHEET_ID`: spreadsheet ID (not the whole URL).
- `OWNER_EMAIL`: owner verification and existing-item association.
- `OWNER_NAME`: name used if an owner People record must be created.

Generate the verifier locally in a VS Code terminal with `node .\\tools\\hash-password.cjs`. Enter your own password at the hidden prompt. Copy the output record into Script Property `FAMILY_PASSWORD_HASH`. The password itself is not printed or saved. Password input is UTF-8 without normalization and limited to 256 Unicode code points; username is limited to 128 characters.

The record format is `pbkdf2-sha256-client$600000$<16-byte salt as 32 hex chars>$<SHA-256 of derived key as 64 hex chars>`. The 600,000 PBKDF2-HMAC-SHA-256 iterations happen in the browser using Web Crypto. The browser sends the derived key—not the plaintext password—to Apps Script. Apps Script hashes that key once and compares the verifier cheaply. The salt is public and included in the login response. Do not change the iteration count in only one component.

## Local source and clasp

From the repository root in a Windows 11 VS Code terminal:

```powershell
npm install --global @google/clasp
clasp login
Copy-Item .clasp.json.example .clasp.json
code .clasp.json
```

Set the spreadsheet-bound Script ID and keep `rootDir` as `apps-script`. `.clasp.json` is ignored by Git. Inspect before upload with `clasp status`, then `clasp push` to the intended script project. Do not overwrite a project containing work you need.

## Preserve existing wishlist items

The owner-only setup checks the active Google account against `OWNER_EMAIL`. Reload the spreadsheet after setting properties and choose **Wishlist admin → Initialize database and connect existing items**. The setup creates missing tabs and headers, stops on header conflicts, and never clears rows. It changes only `PersonID` cells exactly equal to `OWNER-SETUP`; the existing 13 item records and their other columns are preserved.

If the menu does not appear: (a) reload the Sheet; (b) open **Extensions → Apps Script**, select and run `runOwnerSetup` (the same owner check still applies); (c) if `Session.getActiveUser().getEmail()` is still blank, create an installable trigger in the Apps Script Triggers panel: add an `onOpen` trigger, event source **From spreadsheet**, event type **On open**, then reload. Do not remove or bypass the owner check. If the email remains unavailable, stop and ask your Workspace administrator whether identity is available to the script.

## Tests

Run local helper checks:

```powershell
node .\\tools\\hash-password.cjs --self-test
node .\\tools\\hash-password.cjs --test-vector
```

The first checks the RFC PBKDF2 vector, Node WebCrypto versus `pbkdf2Sync` at 600,000 iterations, and wrong-password derivation. `--test-vector` emits fixed public test data and its derived key for manual client/server comparison; it is not a real credential. In Apps Script, use **Wishlist admin → Run compatibility checks** to test the server's SHA-256 verifier comparison and URL cases. That server test uses a deterministic verifier fixture; it does not redo PBKDF2 or access your password.

## Authentication limitations

A signed, expiring login challenge and a 10-bit proof-of-work add friction before authentication; this is not denial-of-service protection and cannot prevent Apps Script quota exhaustion. Failed-login throttling uses Google's temporary per-user key when available. If unavailable, there is no reliable per-caller server-side rate limit. There is deliberately no family-wide lockout.

Sessions use random tokens; only token hashes are cached server-side. The token exists in browser page memory only, not URL or browser storage. Reloading logs out. Cache eviction may expire sessions early, and someone with access to an open browser/devtools may access that active session.

This is one shared login, not individual identity verification. Email addresses are not verified, and anyone with shared credentials can submit under another person's name/email. Keep the spreadsheet private and avoid sensitive information.

Deploy as a web app executing as the owner. Choose the most restrictive audience that works for your family. Complete Google authorization yourself; this repository setup does not access your Google account or deploy anything.

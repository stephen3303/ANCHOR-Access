# ANCHOR web front end (prototype)

A browser UI for ANCHOR that reads and writes the same Access data file
(`ANCHOR_BE.accdb`) as the Access front end. The relational database stays
the one source of truth: tables, relationships, indexes and the Access app all
keep working, and both front ends can be open at once.

```
 Browser (React, one HTML file)
      |  HTTP + JSON, Windows sign-in (Negotiate/NTLM)
      v
 web/server/AnchorWeb.ps1   (Windows PowerShell 5.1, on a Windows PC or server)
      |  ACE OLE DB, parameterized SQL, transactions
      v
 ANCHOR_BE.accdb  <---- Access front end (ANCHOR.accdb) linked as today
```

## Why it is built this way

- **A browser can't open an .accdb.** Something on Windows has to hold the
  Access engine and answer the browser. That is `AnchorWeb.ps1`.
- **Nothing to install.** The server is Windows PowerShell (built into every
  Windows 10/11 PC) and the Access engine that Office already installs. The UI
  is built once into `web/dist/index.html` and committed, so the machine that
  runs it needs no Node, no .NET SDK and no admin rights for localhost use.
- **Windows sign-in, no passwords.** The server asks the browser for the
  person's Windows account (the same idea as the Access app's sign-in), then
  looks it up in the shared Personnel directory (`kit_person.windows_account`)
  to get the `person_id` ANCHOR's tables use. No account, or no active ANCHOR
  profile: no access.
- **Same rules as the Access app.** Writes are copies of the Access code paths:
  - Acknowledge (`modMessages.Acknowledge`): the receipt and its audit row are
    one transaction; if the audit row can't be written, nothing is saved.
  - Opening a message (`modMessages.MarkRead`): first open records the read
    and audits the view.
  - Audit rows are sealed onto the same SHA-256 hash chain
    (`modAudit.AuditSealPending`), so **Verify on the Access Audit page still
    passes after web activity**, and the web Audit page verifies Access rows.
- **Safe by default.**
  - Listens on `localhost` only unless started with `-Lan`.
  - Every value goes to the engine as a parameter, never pasted into SQL.
  - Writes need a per-run token that only the page itself can read, so another
    web site can't make the browser acknowledge something on your behalf.
  - Only its own host names are answered (stops DNS-rebinding tricks).
  - The page is self-contained with a strict Content-Security-Policy: no
    scripts, fonts or images from anywhere else.
  - Classification markings other than UNCLASSIFIED show the demonstration
    banner (Stephen, 2026-11-05).

## Recommended setup: one copy per PC, one data file on the share

This is the same split as the Access app today. Each PC gets this `web` folder
(in place of a local ANCHOR.accdb front end), and every copy points at the one
`ANCHOR_BE.accdb` on the network drive.

```
 PC 1: browser -> localhost server ─┐
 PC 2: browser -> localhost server ─┼──> S:\Apps\ANCHOR\ANCHOR_BE.accdb  <── Access front ends
 PC 3: browser -> localhost server ─┘
```

- Each server serves only its own user, so the one-request-at-a-time limit
  never matters.
- Concurrency is the same as N Access front ends: the Access engine's lock file
  on the share keeps everyone apart, and web and Access users can mix.
- Localhost only: no firewall rule, no admin step, nothing reachable from the
  network.

To set up a PC: copy the `web` folder (only `Start-AnchorWeb.cmd`,
`Start-AnchorWeb.ps1`, `server\` and `dist\` are needed), copy
`anchor-web.config.example.json` to `anchor-web.config.json` and set the two
paths, then double-click `Start-AnchorWeb.cmd`. It starts the server hidden
(picking the PowerShell that matches Office's bitness), waits for it, and opens
the browser. Running it again just opens the browser. `Start-AnchorWeb.ps1 -Stop`
stops it. The log is `%LOCALAPPDATA%\ANCHOR\web.log`.

Each PC needs Access or the free Access Database Engine, and the person needs
read/write rights on the share (the same as for the Access front end).

## Run it by hand

From the `access` folder, with a seeded data file (see `backend\new-backend.ps1`
and `backend\seed-backend.ps1`):

```powershell
# Real sign-in: Windows account -> Personnel directory
..\web\server\AnchorWeb.ps1 -DataFile build\ANCHOR_BE.accdb -Personnel S:\Apps\Personnel\Personnel_be.accdb

# Review/demo: pick who to act as (anyone who reaches the page can be anyone)
..\web\server\AnchorWeb.ps1 -DataFile build\ANCHOR_BE.accdb -DemoSignIn
```

Then open <http://localhost:8787/> in Edge or Chrome. Ctrl+C stops the server.

Bitness: the PowerShell that runs the server must match Office. With 64-bit
Office (the usual case) use the normal `powershell.exe`. With 32-bit Office use
`C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`. On a PC without
Office, install the free Microsoft Access Database Engine redistributable.

To serve other PCs on the network (`-Lan`), an admin reserves the address once
and opens the port:

```powershell
netsh http add urlacl url=http://+:8787/ user=DOMAIN\serviceaccount
New-NetFirewallRule -DisplayName "ANCHOR web" -Direction Inbound -Protocol TCP -LocalPort 8787 -Action Allow
```

## Change the UI

```bash
cd web/ui
npm install
npm run mock     # stand-in API over access/backend/seed, port 8788 (no Access needed)
npm run dev      # Vite dev server, proxies /api to the mock
npm run build    # writes web/dist/index.html; commit it
```

## What's in the prototype

| Page | Reads | Writes |
|---|---|---|
| Home | acknowledgements due, unread count, latest traffic | Acknowledge |
| Messages | list with search, precedence, unread and ack-due filters; message view with sections, attachments, receipts, official text, integrity | mark read (first open audited), Acknowledge |
| Audit | last 200 entries, full chain verify (Recompute) | none |

API (all JSON): `GET /api/me`, `GET /api/summary`, `GET /api/messages?q=&precedence=&unread=1&ackDue=1`,
`GET /api/messages/{id}`, `POST /api/messages/{id}/ack`, `GET /api/audit`, and with `-DemoSignIn` `GET /api/people`.

## Limits to know before going further

- **Access file concurrency.** The Access engine is a file-sharing database:
  every user (Access or web) locks pages of the same file over the network.
  It's comfortable for tens of people and fine for this server, which is one
  process making short requests one at a time. For hundreds of concurrent
  users, or a server farm, move the tables to SQL Server (Express is free) and
  keep Access as a linked front end; the API and UI stay the same.
- **One request at a time.** PowerShell's listener here is single-threaded.
  With one copy per PC that never matters. Only the shared-server option
  (`-Lan`) would feel it; if that's ever wanted at scale, the same endpoints
  port directly to a small C# service (one .exe, Windows sign-in, OLE DB).
- **The data file's network share** must be reachable by the account that runs
  the server, with the same read/write rights the Access front end needs.
- **Business rules live in two places** (VBA and the server) for the actions
  the web does. Keep each web write a line-for-line copy of its VBA routine and
  name that routine in a comment, as the prototype does.

# Web front end on Access data files (prototype)

One browser front end for all the apps. Each app keeps its own tables in its
own Access data file (`.accdb`) on the share; there are no Access front ends.
Access is used only as the database engine, in place of a SQL Server.

```
 Browser (React, one HTML file): app switcher, one shell, every app's pages
      |  HTTP + JSON, Windows sign-in (Negotiate/NTLM), localhost only
      v
 kit/server/KitHost.ps1 + KitWeb.ps1   (Windows PowerShell 5.1, this PC)
      |  apps/anchor/api.ps1   apps/examgen/api.ps1   ...  (one module per app)
      |  ACE OLE DB, parameterized SQL, transactions
      v
 S:\Apps\ANCHOR\ANCHOR_BE.accdb   S:\Apps\ExamGen\ExamGen_BE.accdb   ...
 S:\Apps\Personnel\Personnel_be.accdb   (who is who: shared by every app)
```

## Layout

```
web/
  kit/                       shared core, nothing app-specific (see kit/README.md)
    server/KitWeb.ps1        sign-in, app registry, routing, data layer, security
    server/KitHost.ps1       reads web.config.json, registers the apps, starts
    Start-KitWeb.ps1         per-PC launcher
    ui/                      React shell, components, API client, palettes
  apps/
    anchor/api.ps1           ANCHOR: profile rule, routes, audit chain
    anchor/ui/               ANCHOR pages (Home, Messages, Audit)
    examgen/api.ps1          Exam Generator: profile rule (roles), Overview route
    examgen/ui/              Exam Generator Overview page
  ui/                        Vite project: src/main.tsx lists the app modules
  dist/index.html            the built UI (committed)
  test/test-web.ps1          end-to-end check on real data files
  web.config.example.json    which apps, which data files
  Start-Web.cmd              double-click launcher
```

Adding an app: an `apps/<key>/` folder (module plus pages), one line in
`ui/src/main.tsx`, one entry in `web.config.json`. Nothing in `kit/` changes.

## Permissions and audit stay per app

- **One identity, many permissions.** The server reads the Windows account
  from the browser and looks it up once in the Personnel directory
  (`kit_person.windows_account`) to get the `person_id`. Whether that person
  may use an app, and as what, comes only from **that app's own data file**
  (ANCHOR: an active `tblPeople` row; Exam Generator: `tblPeople` plus
  `tblPersonRoles`). The check runs on every request, so removing someone
  from an app's people list takes effect on their next click.
- **The shell shows only the apps you have a profile in.** An app whose data
  file is unreachable is shown as unavailable; the others keep working.
- **A route sees only its own app's file.** `/api/anchor/...` gets a connection
  to the ANCHOR file and nothing else. A bug in one module cannot read or
  change another app's tables through the kit.
- **Each audit trail lives in its own app's file**, written in the same
  transaction as the change it records. ANCHOR's hash chain (`modAudit`:
  SHA-256 over the previous hash and the row) is sealed by the ANCHOR module
  only; another app with its own audit table keeps its own chain. Chains are
  never mixed, so one app's file can be verified, archived or handed to an
  inspector alone.

## Recommended setup: one copy per PC, data files on the share

```
 PC 1: browser -> localhost server ─┐      S:\Apps\ANCHOR\ANCHOR_BE.accdb
 PC 2: browser -> localhost server ─┼──>   S:\Apps\ExamGen\ExamGen_BE.accdb
 PC 3: browser -> localhost server ─┘      S:\Apps\Personnel\Personnel_be.accdb
```

- Each server serves only its own user, so its one-request-at-a-time listener
  never matters.
- Localhost only: no firewall rule, no admin step, nothing reachable from the
  network.
- Each PC needs Windows PowerShell (built in) and the Access engine: the free
  Microsoft Access Database Engine redistributable is enough. No Access, no
  Node, no .NET SDK, no admin rights.

To set up a PC: copy the `web` folder (only `Start-Web.cmd`, `kit\`, `apps\`
and `dist\` are needed), copy `web.config.example.json` to `web.config.json`
and set the paths, then double-click `Start-Web.cmd`. It starts the server
hidden (picking the PowerShell that matches the engine's bitness), waits for
it, and opens the browser. Running it again just opens the browser.
`kit\Start-KitWeb.ps1 -Config web.config.json -Stop` stops it. The log is
`%LOCALAPPDATA%\<title>\web.log`.

Tip: put the `web` folder itself on the share and give each PC a shortcut to
`Start-Web.cmd`; one update then reaches every PC.

## Run it by hand, and test it

```powershell
# Real sign-in: Windows account -> Personnel directory
web\kit\server\KitHost.ps1 -Config web\web.config.json

# Review/demo: pick who to act as (anyone who reaches the page can be anyone)
web\kit\server\KitHost.ps1 -Config web\web.config.json -DemoSignIn

# End-to-end check on scratch copies (it writes one read, one ack and audit rows)
web\test\test-web.ps1 -DataFile access\build\ANCHOR_BE.accdb -ExamGenDataFile <eg>\access\build\ExamGen_BE.accdb
```

Then open <http://localhost:8787/>. The test also checks the per-app rules: the
same person is refused by Exam Generator without a profile there while ANCHOR
still answers, then allowed once a profile row exists.

## Change the UI

```bash
cd web/ui
npm install
npm run mock     # stand-in API on port 8788 (ANCHOR over access/backend/seed, invented Exam Generator data)
npm run dev      # Vite dev server, proxies /api to the mock
npm run build    # writes web/dist/index.html; commit it
```

## What's in the prototype

| App | Page | Reads | Writes |
|---|---|---|---|
| (shell) | Apps | your apps and profiles | none |
| ANCHOR | Home | acknowledgements due, unread, latest traffic | Acknowledge |
| ANCHOR | Messages | search and filters; message view with sections, attachments, receipts, integrity | mark read (first open audited), Acknowledge |
| ANCHOR | Audit | last 200 entries, full chain verify | none |
| Exam Generator | Overview | questions and exams by status, topics, templates, lookups | none |

API: `GET /api/me` (you, your apps, token), `/api/people` (demo only), and per
app under `/api/<key>/`: ANCHOR `summary`, `messages`, `messages/{id}`,
`POST messages/{id}/ack`, `audit`; Exam Generator `overview`.

## Concurrency: what Access as the database can and can't do

Stated plainly, for this setup (one localhost server per PC, files on an SMB
share):

| Limit | What it means here |
|---|---|
| **255 connections per file** (engine maximum) | Each server opens a file only for the length of one request, so even 100 PCs rarely hold more than a handful at once. Not the limit you will hit first. |
| **Locking is file-sharing, not a server** | Every PC's engine locks pages in the file over the network and coordinates through the `.laccdb` lock file. Two people writing the same area at the same moment wait on each other; the server runs a locked request again up to 4 times (150 to 450 ms apart). Two PCs inserting the same row at once is handled the same way: the loser rolls back, runs again and sees the winner's row (the two-server race test: the same person acknowledging on two servers at once must get one 200 and one 409, one receipt, one audit row). |
| **Comfortable range** | Tens of people actively using one app's file, mostly reading, on a wired LAN. Writes such as acknowledgements are short and fine. Separate files per app help: ANCHOR traffic never locks Exam Generator's file. |
| **2 GB per file** | Including attachments and bloat between compactions. Separate files per app spread this out. Keep large attachments as files on the share, not in tables. |
| **Network quality matters** | A PC dropping off Wi-Fi or VPN in the middle of a write can corrupt the file. Wired LAN only for anyone who writes. |
| **No recovery log** | There is no point-in-time restore. Back up by copying each file on a schedule while nobody is connected (a copy taken mid-write can be inconsistent), and keep several days. Compact and repair needs everyone out. |
| **Each request opens the file** | Over SMB that costs tens of milliseconds per request on a good LAN, more on a slow one. Acceptable for a page at a time; if it shows, hold one connection per app open in each server. |

Time to move an app's tables to SQL Server (Express is free) when you see any
of: regular "record is locked" messages, a file past about 1 GB, people who must
work over VPN or Wi-Fi, a need for point-in-time restore or round-the-clock
use, or more than roughly 50 people writing to one app daily. The move is per
app: one app can be on SQL Server while the rest stay on Access files.

## Keeping a SQL Server move cheap

- **One seam.** Routes reach the database only through `Open-KitDb`,
  `Invoke-KitQuery`, `Invoke-KitNonQuery`, `Use-KitTransaction` and two
  error tests (`Test-KitLocked`, `Test-KitDuplicateKey`) (see `kit/README.md`). The SQL Server OLE DB driver
  (MSOLEDBSQL) takes the same positional `?` parameters, so routes keep their
  parameter lists.
- **What changes in an app's SQL** when it moves: `IIf(...)` to `CASE WHEN`,
  `Now()` to `SYSDATETIME()`, Yes/No values (-1) to `bit` (1),
  `AUTOINCREMENT`/`MEMO` in DDL to `IDENTITY`/`nvarchar(max)`, string `&` to `+`.
  Today the ANCHOR module has 6 `IIf` and 5 `Now()`; Exam Generator has none.
  New routes should keep to SQL that both engines accept where it's easy
  (`CASE` isn't available in Access, so `IIf` stays for now).
- **The schema is untouched.** Existing `.accdb` schemas are read as they are;
  nothing in the web layer needs a column added. Each file keeps its tables,
  relationships and indexes, and SQL Server's upsizing tools can copy them.

## Without Access front ends: what has to be ported

With no Access apps, every business rule the VBA modules carry today must live
in the server module (`apps/<key>/api.ps1`) with a page for it. Layout-only
modules (`modBuild*`, themes, palettes, UI helpers) are not ported: the React
shell replaces them. Line counts are the VBA today, as a size guide.

| App | Logic to port (VBA module, lines) | Notes |
|---|---|---|
| **ANCHOR** | Drafting `modDrafting` (969), routing and release `modRouting` (656), naval message format `modNavalFormat` (730), templates `modTemplates` (622), subscriptions and approvals `modSubscriptions` (396), briefings `modBriefings` (312), watch turnover `modWatch` (272), search `modSearch` (263), reports `modReports` (399), admin `modAdmin` (377), sign-in and profile `modSignIn` (358). Done in the prototype: read/acknowledge (`modMessages`, `modMessageView`) and audit (`modAudit`). | Mail ingest `modAnchorIngest` (1,309) drives Outlook through COM; a localhost server can do the same as the signed-in user, but it is the riskiest piece. Message Generator's drafting is absorbed into ANCHOR. |
| **Exam Generator** | Question workflow and versioning `modQuestions` (1,181), exam build and question swap `modExams` (1,049), question approvals `modApprovals` (383), exam sign-off `modExamApprovals` (608), notifications and tasks `modNotify` (568), admin `modAdmin` (490), sign-in and roles `modSignIn` (640), stats `modStats` (169), logs `modLogs` (156). Done: read-only Overview. | Exam output `modExamDoc` (561) drives Word through COM; on the web the exam can be produced as print-ready HTML/PDF from the browser instead, or the server can keep using Word on the PC. |
| **SPM** | Programs `modPrograms` (729), people `modPeople` (524), Personnel import/export `modPersonnelIO` (466), views `modViews` (390), page 13 `modPage13` (172), command types `modCommandType` (267), roles `modRoles` (257), access rules `modAccess` (236), records `modRecords` (112), document types `modDocTypes` (193), sign-in `modSignIn` (406). | Dialog and form-library modules (`modDialogs`, `modFormLibrary`, `modPills`) are UI and become React pages. |
| **SABER** | Inspections `modInspections` (578), grading `modGrading` (308), roll-up `modRollup` (427), comments `modComments` (478), CAP library `modCapLibrary` (297), team and roster `modTeam`, `modRoster` (383), reports `modReports` (573), sign-in `modAuth` (157). | Excel export `modXlsx`/`modZip` (440) writes the file itself (no Excel COM), so it ports directly. |
| **Message Generator** | Absorbed by ANCHOR: its drafting, routing, approvals, revision and formatting rules (`modMGDrafts`, `modMGRoute`, `modMGApprovals`, `modMGRevision`, `modMGFormat`) map onto ANCHOR's. | Port only what ANCHOR lacks; its own data file can be retired once migrated. |
| **Shared** | Personnel directory admin and sign-in (`modPersonnel`, `modKitUser`), back-end creation and upgrades (`modBuildBackEnd` in each app). | Back-end builders run inside Access today. Without Access they become PowerShell scripts that run the same DDL through the engine, so a new or upgraded data file needs no Access either. |

Rule while porting: each route is a line-for-line copy of its VBA routine,
named in a comment, with the same transaction boundaries and audit rows; then
the VBA is retired. The prototype's acknowledge and mark-read routes show the
pattern.

## Port or bridge? (spike, 2026-11-06)

The alternative to porting: keep the VBA, and have the server call it in a
hidden Access instance over COM (`kit/server/KitBridge.ps1`,
`apps/anchor/bridge/`). Measured on review VM #1 with ANCHOR's real
`modMessages.Acknowledge`, `MarkRead` and `modAudit.AuditVerify`:

| | Full Access (automation) | Access `/runtime` |
|---|---|---|
| Hidden Access start | 0.9 s | 3.4 s |
| Memory | 138 to 145 MB | 94 to 101 MB |
| Acknowledge inside Access | 37 ms | similar |
| Untrapped VBA error | error box caught, Access restarted in under 1 s | came back as a COM error in 2 ms |
| Stray MsgBox | caught, restarted | caught, restarted |
| Hang | stopped at the 8 s timeout, restarted | same, restart 3.5 s |
| Audit Verify afterwards | passes | passes |

It works, and the VBA runs unchanged behind a small wrapper per entry point.
What it costs:

- **Every PC needs Access or the Access Runtime**, matched to Office's
  bitness. The free Runtime would not install beside Click-to-Run Office on
  the VM, so a PC with Office relies on its full Access.
- **Each PC needs a Trusted Location** for the logic file, and the logic file
  is rebuilt whenever the VBA changes.
- **A dialog nobody catches stops that PC's server until the timeout.** Every
  routine called must be checked for `MsgBox`, forms and prompts, because it
  was written for a person sitting at Access. Full and runtime Access behave
  differently here (above).
- **A hidden process to look after**: killed servers leave it running (now
  cleaned up at the next start and by `-Stop`), and its memory is on top of
  the server's.
- **VBA stays the logic language indefinitely**, which is the opposite of
  "no Access apps".

**Recommendation: port.** Make the PowerShell module the home of each app's
rules, routine by routine, as the acknowledge route does. Use the bridge only
as a stopgap, behind the same API so pages don't change, for a large routine
that is already free of UI and isn't worth porting yet (for example ANCHOR's
naval message formatting), and retire it as those routines are ported.

## What moves into the kit repo

- `web/kit/` as a whole, versioned (`KIT_WEB_VERSION.txt`) and copied into the
  web front end the way the Access kit is copied into each app.
- `make_web_palettes.py` reads the kit's own `make_themes.py`, so a palette
  added there reaches the web on the next copy.
- A template app folder (`apps/_template/`: module with a profile rule and one
  route, one page) so a new app starts from a copy.
- Worth adding later: a generic list/detail page component, a generic
  hash-chained audit module (today in ANCHOR's module), and the test harness
  split into kit checks and per-app checks.

Where the web front end itself lives (a new repo, the kit repo, or here) is
still Stephen's call; it is kept here until then.

## Security defaults

- Listens on `localhost` only unless started with `-Lan`.
- Every value goes to the engine as a parameter, never pasted into SQL.
- Writes need a per-run token only the page itself can read, so another site
  can't make the browser act for you.
- Only its own host names are answered (stops DNS-rebinding tricks).
- The page is self-contained with a strict Content-Security-Policy.
- Classification markings other than UNCLASSIFIED show the demonstration
  banner (Stephen, 2026-11-05).

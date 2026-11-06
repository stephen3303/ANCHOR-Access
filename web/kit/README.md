# Kit web (vendored copy)

The shared part of the web front end: one server and one browser shell that
host every app, each app on its own Access data file. This folder is meant to
live in the kit repo and be copied into the web front end's `kit/`, the same way
the Access kit is copied into each app's `access/kit/`. Nothing in it is
specific to one app.

| Part | What it does |
|---|---|
| `server/KitWeb.ps1` | The server core: HTTP listener, Windows sign-in mapped through the shared Personnel directory, app registry (`Use-KitApp`), per-app routing under `/api/<key>/`, per-app profile check, the data-layer seam (below), record-lock retry, CSRF token, Host allow-list, CSP. |
| `server/KitHost.ps1` | Reads the settings file (`web.config.json`), registers each listed app with its data file, dot-sources its module, starts the server. |
| `Start-KitWeb.ps1` | The per-PC launcher: starts `KitHost.ps1` hidden with the PowerShell that matches Office's bitness, waits, opens the browser; `-Stop` stops it. |
| `ui/` | The React shell (`KitHost`: sign-in, demo picker, app switcher, apps page, sidebar, top bar, palette picker, toasts), the API client (`appApi(key)`), shared components (Card, Stat, Pill, ClassificationBanner, dates), `base.css`. Apps import it as `@kit`. |
| `ui/palettes.css`, `ui/palettes.ts` | Generated: never edit. The Access kit's named palettes as CSS variables (`--clr-primary` = `CLR_PRIMARY`). |
| `tools/make_web_palettes.py` | Regenerates the two palette files from `access/tools/make_themes.py`. |

## The data-layer seam

Routes never touch OLE DB directly. Everything goes through five functions:

| Function | Does |
|---|---|
| `Open-KitDb path` | Opens a connection (ACE OLE DB today). `$r.Conn` is already open for a route. |
| `Invoke-KitQuery conn sql params [tx]` | Rows as hashtables. |
| `Invoke-KitNonQuery conn sql params [tx]` | Rows affected. |
| `Use-KitTransaction conn { param($tx) ... }` | All or nothing. |
| `Test-KitLocked err`, `Test-KitDuplicateKey err` | Another PC holds the record, or inserted the same key a moment earlier: the request is run again (up to 4 times). |

Moving one app to SQL Server means changing `Open-KitDb` (connection string per
app) and the two error tests, then fixing that app's SQL where Access's dialect
differs (see the main README). Routes, UI and the other apps don't change.

## An app supplies

- `apps/<key>/api.ps1`: its profile rule (`Set-KitProfileResolver`), optional
  demo people list and start-up step, its routes (`Add-KitRoute`), and the
  business rules those routes perform. Prefix every function with the app name
  (`Get-AnchorSetting`): all modules share one scope.
- `apps/<key>/ui/index.tsx`: a `KitAppModule` (key, name, mark, nav, `page(path)`)
  and its pages; `ui/src/main.tsx` lists the modules.
- One entry in `web.config.json`: key, name, data file, module.

## Rules for routes

- Every value goes in as a `?` parameter; only integers the server resolved itself may be pasted into SQL.
- Bracket column names that are Access reserved words (`[Action]`, `[Note]`, `[Date]`, `[Name]`, `[Value]`).
- A write that has an audit row or touches a second table runs inside `Use-KitTransaction`.
- A route sees only its own app's data file and never opens another app's.
  A page that shows two apps' data calls both apps' APIs from the browser, so
  each app's own permission check still applies.
- Write as if another PC is doing the same thing at the same moment: an
  update that changes state carries the old state in its `WHERE`
  (`... AND AckOn Is Null`) and checks the row count.
- Return through `$r.Json(status, object)`, once, after the last write (a
  request that is retried runs the route again). Errors thrown are logged and answered with a 500 that hides the detail.

# Kit web (vendored copy)

The shared part of a kit app's web front end. This folder is meant to live in
the kit repo (`access-modern-ui`, as `web/`) and be copied into each app's
`web/kit/`, the same way the Access kit is copied into each app's `access/kit/`.
Nothing in it is specific to one app.

| Part | What it does |
|---|---|
| `server/KitWeb.ps1` | The server core: HTTP listener, Windows sign-in mapped through the shared Personnel directory, Access engine (ACE OLE DB) connection per request, parameterized query helpers, transactions, record-lock retry, CSRF token, Host allow-list, CSP. Apps add routes with `Add-KitRoute` and two hooks (`Set-KitProfileResolver`, `Set-KitPeopleLister`). |
| `Start-KitWeb.ps1` | The per-PC launcher: reads the app's `*-web.config.json`, starts the app's server hidden with the PowerShell that matches Office's bitness, opens the browser, `-Stop` stops it. |
| `ui/` | The React shell (`KitApp`: sign-in, demo picker, sidebar, top bar, palette picker, toasts), the API client, shared components (Card, Stat, Pill, ClassificationBanner, dates), `base.css`. Apps import it as `@kit`. |
| `ui/palettes.css`, `ui/palettes.ts` | Generated: never edit. The Access kit's named palettes as CSS variables (`--clr-primary` = `CLR_PRIMARY`). |
| `tools/make_web_palettes.py` | Regenerates the two palette files from `access/tools/make_themes.py`, so web and Access share one palette source. |

## An app supplies

- `web/server/<App>Web.ps1`: about 20 lines. Dot-sources the core, calls `Initialize-KitWeb`, dot-sources its module, calls `Start-KitWeb`.
- `web/server/<App>Api.ps1`: its profile rule, its routes, and copies of the VBA routines those routes perform (name each routine in a comment).
- `web/ui/src/`: its pages and `main.tsx` with its `KitAppConfig` (name, nav, routes).
- `web/<app>-web.config.example.json` and `Start-<App>Web.cmd`.

## Rules for routes

- Every value goes in as a `?` parameter; only integers the server resolved itself may be pasted into SQL.
- Bracket column names that are Access reserved words (`[Action]`, `[Note]`, `[Date]`, `[Name]`, `[Value]`).
- A write that has an audit or second table runs inside `Use-KitTransaction`, so it is all or nothing.
- Return through `$r.Json(status, object)`. Errors thrown are logged and answered with a 500 that hides the detail.

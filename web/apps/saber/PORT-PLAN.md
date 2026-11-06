# SABER on the web platform: inventory and port plan

Source: `stephen3303/SABER-Access` at `1ba13e0` (Release v1.5.0, schema version 8).
Target: this web platform (`web/kit` core, `web/apps/saber` module). Access stays
only as the database engine: SABER keeps its own `SABER_be.accdb`; there are no
Access forms. Each VBA routine is ported into `apps/saber/api.ps1` with the same
checks and the same rows written, named in a comment, then retired.

## Inventory

### Tables (SABER_be.accdb, 25 tables, modSchemaDDL.LoadSpecs)

| Area | Tables | Web status |
|---|---|---|
| Accounts | `command`, `team_member`, `system_setting` | read (phase 1); admin edits phase 4 |
| Publication library | `instruction_reference` | phase 4 |
| Approved capabilities | `approved_macc`, `approved_capability`, `approved_attribute` | phase 4 |
| Inspection setup | `inspection_event`, `inspection_role_definition`, `inspection_role`, `inspection_grade_band`, `inspection_hierarchy_weight` | read (phase 1); edits phase 3 |
| Schema | `inspection_mission_area`, `inspection_group`, `inspection_capability`, `inspection_attribute` | read (phase 1); import/export phase 3 |
| Grading | `grade`, `caucus_grade`, `grade_history` | **grade + history written (phase 1)**; caucus phase 2 |
| Comments | `comment`, `comment_version`, `workflow_event` | phase 2 |
| Reports and archive | `report_artifact`, `export_package` | phase 4 |
| Errors | `error_log` | phase 4 (the web server logs to its own log today) |

Shared: the Personnel file (`kit_person`, `kit_person_app` with app code `SABER`)
decides who may open SABER; `team_member.person_id` links the two.

SABER's audit is not a hash chain like ANCHOR's: it is `grade_history` (every
grade and caucus change: who, when, from, to) and `workflow_event` plus
`comment_version` (every comment action and edit). The web writes those same
rows, in the same transaction as the change, so either front end reads one
history.

### Business-logic modules (to port)

| Module | Lines | What | Phase |
|---|---|---|---|
| modAuth | 157 | sign-in, session, review stage (UserStage, MaxStage, StageLabel) | 1 (stage, profile); sign-in is the kit's |
| modTeam | 232 | team_member kept in line with Personnel (sync, EnsureTeamMember, first-admin bridge) | 1 (link, add, refresh; no Personnel writes) |
| modGrading | 308 | SaveGrade, ClearGrade, capability N/E, caucus overrides and locks, LogGradeChange | 1 (grade, clear, history); 2 (N/E, caucus) |
| modRollup | 427 | weighted rollup Attribute > Capability > Group > Mission Area > Overall, N/E rule, bands | 2 |
| modComments | 478 | comments, versions, submit/approve/return workflow, past-comment search | 2 |
| modInspections | 578 | create/delete inspection, lifecycle (activate/close/archive/reopen), roles, schema xlsx in/out | 3 |
| modCapLibrary | 297 | approved MACC list, xlsx import (checked row by row) and export | 4 |
| modReports | 573 | PDF reports (versioned), Excel exports, export package (ZIP + CSV + manifest + SHA-256), verify/import, library CSV | 4 |
| modXlsx, modZip | 440 | read xlsx without Excel; ZIP with optional ZipCrypto password | 3, 4 (port directly: no Office COM) |
| modVersion, modRoster | 274 | front-end build checks, Signed In page | dropped / phase 4 (one web build; roster becomes "seen recently") |
| modErrors | 219 | error log table, event log | 4 |
| modHelp | 182 + help.txt | per-page help | 4 (help.txt reused as page help) |
| modGlobals | 277 | constants, session, SQL helpers | constants ported where used |
| modSchemaDDL, modSeed | 862 | back-end creation, upgrades to schema 8, seed | 5: PowerShell DDL script so a new data file needs no Access |

Not ported (the React shell replaces them): modBuild, modBuildUI, modBuildPages,
modBuildPages2, modBuildAppearance, modUI, modWrap, modPalette, modTheme, and the
kit's Access UI modules.

### Forms (40) and reports (2)

| Access form | Web page | Phase |
|---|---|---|
| frmSplash, frmStartup, frmLogin, dlgSignIn, dlgPassword (sign-in parts) | kit sign-in (Windows) | 1 |
| frmMain (AppShell: sidebar, context, classification) | kit shell + SABER nav | 1 |
| pgDashboard + subEventList (Inspections) | **Inspections** | 1 |
| pgGrading + subTreeGrading, dlgBandPick, dlgGradeHistory | **Inspection > Grading** (tree, score, band, history) | 1 (rollup pills phase 2) |
| pgComments, dlgComment, dlgCommentSearch, dlgHistory, subCommentList | Comments | 2 |
| pgReview, dlgReturn | Review board | 2 |
| pgProgress + subTreeProgress | Progress | 2 |
| pgCaucus + subTreeCaucus | Caucus / Results | 2 |
| pgDetail (Overview, Roles, Schema, Grade Bands) + subRoles, subRoleDefs, subTreeSchema, dlgNewInspection | Inspection detail tabs, New inspection | 3 |
| pgReports, rptGrades, rptComments | Reports & Export (print-ready HTML to PDF in the browser) | 4 |
| pgAdmin (Team, Commands, System, Library, Capabilities, Import package, Error Log, Versions, Signed In) + subCommands, subLibrary, dlgUser, dlgPerson, dlgPersonnel, dlgPersonnelSetup | Admin | 4 |
| dlgMessage, dlgConfirm, dlgInput, dlgToast, dlgBusy, dlgMenu, dlgPopover, dlgFilter, dlgDatePick, dlgAppearance, dlgHelp | kit web components (toast, confirm, inputs, palette picker, help) | as each page needs them |

## Phases

1. **Sign-in, Inspections, grading with history** (this PR). Profile from
   `team_member` via the Personnel person; Personnel sync at start; Inspections
   list with status cards, search and year; the inspection's grading tree; grade
   by score, band or Not Evaluated, clear, per-attribute history; refusals for
   no role, not Active, bad input; two PCs grading one attribute at once.
2. **Rollups, comments and review.** modRollup.Compute (weighted rollup and N/E
   rule) feeding Grading, Progress and Caucus; capability N/E; caucus
   override/lock; comments with versions and the submit/approve/return workflow
   (Comments and Review pages). Sidebar rules from AppShell (Review hidden for
   level 1, closed inspections show only Results and Reports).
3. **Inspection set-up.** New inspection, lifecycle (activate needs a schema;
   close checks ClosureBlockers and locks the caucus; archive purges history;
   reopen admin-only), roles and role definitions, grade bands, weights, schema
   xlsx import/export (modXlsx ported to .NET ZIP + XML).
4. **Reports, packages, admin.** Grades and Comments reports as printable pages,
   Excel exports, export package create/verify/import, library and approved
   capabilities, team members and commands, error log, help.
5. **No Access anywhere.** Back-end creation and schema upgrades as a PowerShell
   DDL script; the Access edition retired for SABER.

Bridge (hidden Access over COM): not used. Every SABER routine above is plain
data work with no Office automation, so each ports directly; the PR 6 spike's
advice (port, bridge only as a stopgap) holds.

## Known gaps after phase 1

- Rollup scores on capability, group, mission area and overall are not shown yet
  (the Access edition shows them on every level); the tree shows graded counts.
- Members with a SABER password and no Personnel person are linked when the
  Personnel file has them, but not added to it (the web never writes another
  app's file); they need the Access edition or the phase 4 Admin page once.
- Personnel-mode sync is written but the automated test runs in demo sign-in;
  it needs one run against a scratch Personnel file.
- `team_member.last_login_at` and the Signed In roster are not updated by the web.

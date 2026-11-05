Option Compare Database
Option Explicit

' ANCHOR (Access edition) - BuildAll, the entry point build.ps1 calls, the
' shell (frmMain) and its sidebar, and the helpers every page builder shares.
'
' The front end is a build output: every form is drawn with the kit's
' builders (modUIKit) and rebuilt by build.ps1, never laid out by hand. The
' shared data file is made separately (backend\new-backend.ps1); BuildAll
' links this front end to it when build.ps1 is given its path, so UI work
' never touches data.
'
' Steps, in order: link, user interface (subforms, then pages, then dialogs,
' then the shell last), help, startup. A failed build returns
' "ERROR during <step>: <number> <description> [<last form traced>]".
'
' Each page has its own builder module, modBuild<Page> (Build_pg<Page>, and
' Build<Page>Grids for the tables it hosts), so pages are built and reviewed
' one at a time. A page not built yet is a placeholder that says so
' (BuildPlaceholderPage).
'
' Pages (the web app's routes): Home (/), Messages (/messages), the Message
' page (/messages/[id], opened from any list, not in the sidebar), Search,
' Subscriptions, Briefings, Watch turnover, Reports, Admin, Audit and
' Profile, plus Drafting & Release, Templates and Admin: Routing, which
' came over from Message Generator (ANCHOR absorbs it). The web app's
' simulators, topology and offline pages are network demos with no place in
' an Access file.
'
' Owner: the shell and the build. Depends on: modBuild* (one per page),
' modBuildAppearance, modErrors, modGlobals; the kit's modDarkMode,
' modGridCols, modGridFilter, modKitBackend, modKitHelp, modKitKeys,
' modKitPalette, modKitShell, modMenu, modPersonnelUI, modTrust, modUIKit.

Public Const APP_VERSION As String = "0.1.0"

Public Const PAGE_W As Long = 15840             ' the page area, 1056 px
Public Const PAGE_H As Long = 9600              ' 640 px
Public Const PAGE_PAD As Long = 480             ' page padding, 32 px

Private m_step As String

' Content width inside the page padding.
Public Function ContentW() As Long
    ContentW = PAGE_W - 2 * PAGE_PAD
End Function

' Build entry point (build.ps1). bePath is the shared data file to link to,
' "" for none (a front end with no data, for UI work). Returns "OK" or the
' step that failed.
Public Function BuildAll(ByVal srcRoot As String, ByVal outDir As String, ByVal uiOnly As Boolean, _
        ByVal bePath As String) As String
    On Error GoTo Fail
    g_Silent = True
    UseBuildPalette
    g_SrcRoot = srcRoot
    If Not uiOnly And Len(bePath) > 0 Then
        m_step = "link": LinkBackEnd bePath
    End If
    m_step = "user interface": BuildUserInterface
    m_step = "help": LoadHelpTopics srcRoot & "\help\help.txt"
    m_step = "startup": SetStartupOptions uiOnly, bePath
    g_Silent = False
    BuildAll = "OK"
    Exit Function
Fail:
    BuildAll = "ERROR during " & m_step & ": " & Err.Number & " " & Err.Description & " [" & BuildTrace() & "]"
    g_Silent = False
End Function

' ---------------------------------------------------------------------------
' Link
' ---------------------------------------------------------------------------

' Link every table in the shared data file (all but Access's system tables),
' replacing any links of the same names.
Private Sub LinkBackEnd(ByVal bePath As String)
    Dim be As DAO.Database, td As DAO.TableDef, names As String
    KitCheckBackendPath bePath
    Set be = DBEngine.OpenDatabase(bePath, False, True)
    For Each td In be.TableDefs
        If Left$(td.Name, 4) <> "MSys" And Left$(td.Name, 4) <> "USys" And Left$(td.Name, 1) <> "~" Then
            names = names & IIf(Len(names) > 0, vbTab, "") & td.Name
        End If
    Next
    be.Close
    If Len(names) = 0 Then Err.Raise 5, , "The data file has no tables: " & bePath
    KitLinkTables bePath, Split(names, vbTab)
End Sub

' ---------------------------------------------------------------------------
' User interface
' ---------------------------------------------------------------------------

' The pages (with the tables they host), the app's dialogs, the kit's
' dialogs, the splash, and the shell last.
Private Sub BuildUserInterface()
    BuildPages
    BuildAppDialogs
    Build_dlgAppearance            ' palette picker (modBuildAppearance)
    BuildMessageDialog             ' the kit's dlgMessage, for KitMsgBox
    BuildInputDialog               ' dlgInput, for KitInputBox
    BuildConfirmDialog             ' dlgConfirm, for KitConfirmTyped
    BuildDatePicker                ' dlgDatePick, for DateBox
    BuildToast                     ' dlgToast, for KitToast
    BuildBusy                      ' dlgBusy, for KitBusy
    BuildColumnsDialog             ' dlgColumns, for ChooseGridColumns
    BuildFilterDialog              ' dlgFilter, for the header funnels (GridFilter)
    BuildMenuDialog                ' dlgMenu and dlgPopover, for KitMenu and KitPopover
    BuildPopoverDialog
    BuildPickDialog                ' dlgPick, for PickBox and PersonPickBox
    BuildHelpDialog                ' dlgHelp, for KitHelp (Help in the sidebar and F1)
    BuildPaletteDialog             ' dlgPalette, for Ctrl+K (KitPalette)
    BuildPersonnelDialogs          ' dlgPersonnelSetup, dlgSignIn, dlgPersonnel, dlgPerson (KitPersonnelStart)
    BuildSplash "frmMain", "Message handling", ChrW(&HE715)
    Build_frmMain
End Sub

' The pages, in sidebar order, each module building its tables first.
Private Sub BuildPages()
    BuildHomePage                  ' modBuildHome
    BuildMessagesPages             ' modBuildMessages: Messages and the Message page
    BuildSearchPage                ' modBuildSearch
    BuildSubscriptionsPage         ' modBuildSubscriptions
    BuildBriefingsPage             ' modBuildBriefings
    BuildWatchPage                 ' modBuildWatch
    BuildReportsPage               ' modBuildReports
    BuildDraftingPages             ' modBuildDrafting: Drafting, Templates, Admin: Routing
    BuildAdminPage                 ' modBuildAdmin
    BuildAuditPage                 ' modBuildAudit
    BuildProfilePage               ' modBuildProfile (a sidebar link)
End Sub

' The dialogs pages open, after the pages (each page module adds its own here).
Private Sub BuildAppDialogs()
    BuildMessagesDialogs           ' modBuildMessages
    BuildSubscriptionsDialogs      ' modBuildSubscriptions
    Build_rptBriefing              ' modBuildBriefings: the printed briefing (report)
    BuildDraftingDialogs           ' modBuildDrafting
End Sub

' A page not built yet: its header and a note that says so, so the sidebar,
' Ctrl+K and the help already work for it.
Public Sub BuildPlaceholderPage(ByVal formName As String, ByVal title As String, ByVal subtitle As String, ByVal glyph As String)
    NewForm "", 0, PAGE_W, PAGE_H
    PageHeader title, subtitle, PAGE_PAD
    EmptyState "emptyPage", PAGE_PAD, px(120), ContentW(), glyph, "This page is on its way", _
        "It is being built from the ANCHOR web app's " & title & " page and arrives in a later build."
    SaveForm formName
End Sub

' Grouping the controls a page shows and hides together (one tab's): call
' NamesNow before building them and GroupSince after.

' The names of the form's controls now.
Public Function NamesNow() As Object
    Dim c As Control, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    For Each c In F.Controls
        d(c.Name) = True
    Next
    Set NamesNow = d
End Function

' List the controls built since before that show (Visible) in a hidden
' label nm, "a;b;c", for the page to show and hide as a group.
Public Sub GroupSince(before As Object, ByVal nm As String)
    Dim c As Control, s As String
    For Each c In F.Controls
        If Not before.Exists(c.Name) Then
            If c.Visible Then s = s & ";" & c.Name
        End If
    Next
    If Len(s) > 2048 Then Err.Raise 5, , "Group " & nm & " lists too many controls for a label caption"
    Set c = Lbl(nm, Mid$(s, 2), 0, 0, px(8), px(8))
    c.Visible = False
End Sub

' 8pt bold muted uppercase caption above a field.
Public Sub FieldLbl(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long)
    Lbl nm, cap, x, y, w, px(16), 8, True, CLR_MUTED_FG
End Sub

' The shell: the sidebar of pages and the page host sfPage. At sign-in the
' pages a person's roles don't allow are hidden (ShowPagesForRoles).
' Profile is a bottom link; it, the Message page, Templates (from Drafting)
' and Routing (from Admin), which no sidebar item opens, come from frmMain's
' PageForm.
Private Sub Build_frmMain()
    ShellBuild APP_NAME, "Message handling", _
        Array("home|Home|pgHome|E80F", "messages|Messages|pgMessages|E715", "drafting|Drafting|pgDrafting|E70F", _
            "search|Search|pgSearch|E721", _
            "subscriptions|Subscriptions|pgSubscriptions|EA8F", "briefings|Briefings|pgBriefings|E8A5", _
            "watch|Watch turnover|pgWatchTurnover|E916", "reports|Reports|pgReports|E9D2", _
            "admin|Admin|pgAdmin|E716", "audit|Audit|pgAudit|E81C"), _
        Array("help|Help|E897|=KitHelp()", "profile|Profile|E77B|=NavTo(""profile"")"), _
        PAGE_W, PAGE_H
    SaveForm "frmMain"
End Sub

' ---------------------------------------------------------------------------
' Startup
' ---------------------------------------------------------------------------

' Set database property name, adding it when the front end hasn't got it
' yet (error 3270). Any other failure, and a failed add, stops the build.
Private Sub SetProp(ByVal name As String, ByVal kind As Integer, ByVal value As Variant)
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error GoTo Missing
    db.Properties(name) = value
    Exit Sub
Missing:
    If Err.Number <> 3270 Then RaiseAgain "modBuild.SetProp", name
    db.Properties.Append db.CreateProperty(name, kind, value)   ' an error here goes up to BuildAll
End Sub

' Open on the splash screen, which opens frmMain, without Access's own
' splash, the navigation pane or document tabs, and record when this front
' end was built (error-log entries and the splash show it). BuildTrustCheck
' opens the splash at startup (in place of StartUpForm) and shows the kit's
' trust page when Access would ask to enable content.
' Shift at opening (AllowBypassKey) and F11 (AllowSpecialKeys) stay on in
' this .accdb, the developer's copy; build.ps1 -Release turns them off in
' the locked .accde users get.
' HasDataFile records whether this front end was built linked to a data file,
' so a copy that loses its tblSettings link isn't taken for a UI-only build
' (modSignIn.DataLinked). -UiOnly keeps the value the full build set.
Private Sub SetStartupOptions(ByVal uiOnly As Boolean, ByVal bePath As String)
    If Not uiOnly Then SetProp "HasDataFile", dbBoolean, (Len(bePath) > 0)
    SetProp "AppTitle", dbText, APP_NAME
    SetProp "StartUpShowDBWindow", dbBoolean, False
    SetProp "ShowDocumentTabs", dbBoolean, False
    SetProp "Perform Name AutoCorrect", dbLong, 0
    SetProp "Track Name AutoCorrect Info", dbLong, 0
    SetProp "AppVersion", dbText, APP_VERSION
    SetProp "AppBuild", dbText, Format$(Now(), "yyyy-mm-dd hh:nn")
    BuildShortcutKeys               ' Ctrl+S, Ctrl+N, Ctrl+F, Ctrl+K and F1, the AutoKeys macro
    HideAccessSplash                ' it writes a file, so a failure leaves the settings above in place
    BuildTrustCheck "frmSplash"     ' after any StartUpForm setting: it removes StartUpForm
End Sub

Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Watch turnover page (/watch-turnover): watch notes, ongoing topics and the turnover checklist.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, BuildPlaceholderPage); the kit's modUIKit.

Public Sub BuildWatchPage()
    BuildPlaceholderPage "pgWatchTurnover", "Watch turnover", "Notes, ongoing topics and the checklist for handing over the watch.", "E916"
End Sub

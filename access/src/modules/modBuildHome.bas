Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Home page (the web app's /): mandatory acknowledgements, status changes and the latest traffic. Built by modBuild.BuildPages.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, BuildPlaceholderPage); the kit's modUIKit.

Public Sub BuildHomePage()
    BuildPlaceholderPage "pgHome", "Home", "What needs you now: acknowledgements due, status changes and the latest traffic.", "E80F"
End Sub

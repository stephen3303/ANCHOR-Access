Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Search page (/search): basic, advanced and saved searches.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, BuildPlaceholderPage); the kit's modUIKit.

Public Sub BuildSearchPage()
    BuildPlaceholderPage "pgSearch", "Search", "Find messages by words, reference, originator, type, precedence and dates.", "E721"
End Sub

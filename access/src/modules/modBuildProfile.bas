Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Profile page (/profile): your details and roles, colors and Sign out.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, BuildPlaceholderPage); the kit's modUIKit.

Public Sub BuildProfilePage()
    BuildPlaceholderPage "pgProfile", "Profile", "Your details, your roles, your colors and signing out.", "E77B"
End Sub

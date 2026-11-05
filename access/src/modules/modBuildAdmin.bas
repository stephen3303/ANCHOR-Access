Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Admin page (/admin): subscription approvals, mandatory alert rules, users and command access.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, BuildPlaceholderPage); the kit's modUIKit.

Public Sub BuildAdminPage()
    BuildPlaceholderPage "pgAdmin", "Admin", "Approve subscriptions, review alert rules, and manage users and command access.", "E716"
End Sub

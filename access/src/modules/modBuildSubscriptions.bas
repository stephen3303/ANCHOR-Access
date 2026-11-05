Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Subscriptions page (/subscriptions) and the request dialog.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, BuildPlaceholderPage); the kit's modUIKit.

Public Sub BuildSubscriptionsPage()
    BuildPlaceholderPage "pgSubscriptions", "Subscriptions", "The traffic you follow. Mandatory subscriptions are assigned by your command.", "EA8F"
End Sub

' The dialogs the Subscriptions page opens (modBuild.BuildAppDialogs).
Public Sub BuildSubscriptionsDialogs()
End Sub

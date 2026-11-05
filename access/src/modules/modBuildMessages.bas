Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Messages page (/messages: the list with its filters and preview) and the Message page (/messages/[id]: the official text, the AI summary apart, integrity, lifecycle and Acknowledge).
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, BuildPlaceholderPage); the kit's modUIKit.

Public Sub BuildMessagesPages()
    BuildPlaceholderPage "pgMessages", "Messages", "Every message received, newest first, with filters and a preview.", "E715"
    BuildPlaceholderPage "pgMessage", "Message", "The official text as received, its integrity data and your acknowledgement.", "E715"
End Sub

' The dialogs the message pages open (modBuild.BuildAppDialogs).
Public Sub BuildMessagesDialogs()
End Sub

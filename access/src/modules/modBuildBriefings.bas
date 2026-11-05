Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Briefings page (/briefings and /briefings/generated): the builder and the generated briefing.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, BuildPlaceholderPage); the kit's modUIKit.

Public Sub BuildBriefingsPage()
    BuildPlaceholderPage "pgBriefings", "Briefings", "Build a briefing from the traffic of a period and print it.", "E8A5"
End Sub

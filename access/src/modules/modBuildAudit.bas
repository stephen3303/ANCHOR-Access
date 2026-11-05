Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the Audit page (/audit): the append-only audit trail and its hash chain check.
'
' Owner: this page's build task. Depends on: modBuild (PAGE_W, PAGE_H,
' PAGE_PAD, ContentW, BuildPlaceholderPage); the kit's modUIKit.

Public Sub BuildAuditPage()
    BuildPlaceholderPage "pgAudit", "Audit", "Every audited action, append-only, with a check that no entry was changed.", "E81C"
End Sub

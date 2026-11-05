Option Compare Database
Option Explicit

' ANCHOR (Access edition) - message templates: the web app's Templates page
' (/templates), with Message Generator's template rules (modMGText,
' modMGTemplates), brought over when ANCHOR took over drafting (Stephen,
' 2026-11-05).
'
' A template (tblTemplates) is a draft's header (type, class, precedence,
' classification, TO, INFO, subject) and text, with {{ Name }} fields in the
' subject, the addressees and the text. Each field has its rules in
' tblTemplateFields: required, a maximum length, a default and a hint for
' drafters (Message Generator's field rules). The fields follow the tags:
' saving a template adds a row for a new tag, keeps the rules of one still
' there and drops the row of one gone.
'
' Versions, as in Message Generator: a template name has numbered versions.
' A new version is Inactive, and only an Inactive version can change.
' Activating one retires the name's active version, so a name has at most
' one Active version, the one drafts start from. Retired and Active versions
' never change; New version copies one to edit. Drafts keep the TemplateID
' they started from.
'
' The fixed text (outside the {{tags}}) goes into every draft, whatever its
' classification, so it may carry no classification marking
' (FixedTextMarking, Message Generator's MG_FixedTextMarking). Activation
' also refuses a template whose tags don't parse.
'
' Tag grammar (the web app's extractVariables): {{ Name }}, the name 1 to 64
' characters starting with a letter, then letters, digits, spaces, "_" or
' "-"; the spaces inside the braces are trimmed and names match whatever
' their case, so a repeated tag shares one value.
'
' Editing is for ANCHOR admins (RequireAnchorAdmin at the write); anyone who
' can use the app starts a draft from an Active template.
'
' Owner: drafting and release. Depends on: modAudit, modDrafting, modErrors,
' modGlobals, modSignIn; the kit's modKitLocks (KitExecute), modKitUI
' (KitInputBox, KitMsgBox).

' The Templates page's chips.
Public Const TPL_CHIPS As String = "All|Active|Inactive|Retired"

' The longest hint for drafters (Message Generator's MG_HINT_MAX).
Public Const TPL_HINT_MAX As Long = 100

Private Const TS_INACTIVE As String = "Inactive"
Private Const TS_ACTIVE As String = "Active"
Private Const TS_RETIRED As String = "Retired"

' ---------------------------------------------------------------------------
' Tags
' ---------------------------------------------------------------------------

' The field names text's {{ tags }} hold, in the order they first appear,
' each once (names match whatever their case; the first spelling is kept).
'   text   a template's subject, addressees or text
' Errors: 5 with the reason (to show) for a malformed tag: an unmatched or
'   nested brace, or a name the grammar doesn't allow.
Public Function TemplateTags(ByVal text As String) As Collection
    Dim c As New Collection, seen As Object, pos As Long, opening As Long, closing As Long, stray As Long, nm As String
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    pos = 1
    Do
        opening = InStr(pos, text, "{{", vbBinaryCompare)
        stray = InStr(pos, text, "}}", vbBinaryCompare)
        If stray > 0 And (opening = 0 Or stray < opening) Then Err.Raise 5, "modTemplates.TemplateTags", _
            "A closing }} has no opening {{. Write each field as {{ Name }}."
        If opening = 0 Then Exit Do
        closing = InStr(opening + 2, text, "}}", vbBinaryCompare)
        If closing = 0 Then Err.Raise 5, "modTemplates.TemplateTags", "A {{ is never closed. Write each field as {{ Name }}."
        nm = Mid$(text, opening + 2, closing - opening - 2)
        If InStr(nm, "{") > 0 Or InStr(nm, "}") > 0 Then Err.Raise 5, "modTemplates.TemplateTags", _
            "Fields can't be nested or use extra braces. Write each field as {{ Name }}."
        nm = Trim$(nm)
        If Not TagNameOk(nm) Then Err.Raise 5, "modTemplates.TemplateTags", "{{" & ClipText(nm, 40) & "}}: a field name starts with " & _
            "a letter and holds only letters, digits, spaces, ""_"" or ""-"", up to 64 characters."
        If Not seen.Exists(nm) Then seen.Add nm, True: c.Add nm
        pos = closing + 2
    Loop
    Set TemplateTags = c
End Function

Private Function TagNameOk(ByVal nm As String) As Boolean
    Dim i As Long, ch As String
    If Len(nm) = 0 Or Len(nm) > 64 Then Exit Function
    For i = 1 To Len(nm)
        ch = Mid$(nm, i, 1)
        If Not (ch Like "[A-Za-z]" Or (i > 1 And (ch Like "[0-9]" Or ch = " " Or ch = "_" Or ch = "-"))) Then Exit Function
    Next
    TagNameOk = True
End Function

' Why text's tags don't parse, "" when they do (TemplateTags' reason).
Public Function TemplateTagProblem(ByVal text As String) As String
    Dim c As Collection
    On Error GoTo Bad
    Set c = TemplateTags(text)
    Exit Function
Bad:
    If Err.Number <> 5 Then RaiseAgain "modTemplates.TemplateTagProblem"
    TemplateTagProblem = Err.Description
End Function

' text with each {{ tag }} replaced by its value in values (a dictionary by
' field name, matched whatever the case); a tag with no value shows as
' [Name], as previews show it.
' Errors: TemplateTags' for a malformed tag.
Public Function FillTemplateText(ByVal text As String, ByVal values As Object) As String
    Dim pos As Long, opening As Long, closing As Long, nm As String, s As String, v As Variant
    Dim c As Collection
    Set c = TemplateTags(text)                 ' raises on a malformed tag
    pos = 1
    Do
        opening = InStr(pos, text, "{{", vbBinaryCompare)
        If opening = 0 Then Exit Do
        closing = InStr(opening + 2, text, "}}", vbBinaryCompare)
        nm = Trim$(Mid$(text, opening + 2, closing - opening - 2))
        v = Null
        If Not values Is Nothing Then
            If values.Exists(nm) Then v = values(nm)
        End If
        If IsNull(v) Then
            s = s & Mid$(text, pos, opening - pos) & "[" & nm & "]"
        Else
            s = s & Mid$(text, pos, opening - pos) & CStr(v)
        End If
        pos = closing + 2
    Loop
    FillTemplateText = s & Mid$(text, pos)
End Function

' The first classification marking in text's fixed text (outside the
' {{tags}}), "" when there is none: a portion mark such as (S), (C), (TS),
' (S//NF), a classification word (CONFIDENTIAL, SECRET, also letter-spaced),
' or a control marking (NOFORN, ORCON, REL TO). (U) and words such as
' ITEM(S) are fine; a one-letter (C) or (S) at the start of a line is a
' subparagraph marker. Upper case only, as messages are written. Message
' Generator's MG_FixedTextMarking, unchanged.
Public Function FixedTextMarking(ByVal text As String) As String
    Dim fixed As String, opening As Long, closing As Long, p As Long, q As Long, inner As String
    Dim lineStart As Long, w As Variant
    ' Each tag becomes " # ": the text either side stays apart, and a mark after
    ' a tag ("{{para}} (S) ...") isn't taken for one at the start of a line.
    fixed = text
    Do
        opening = InStr(1, fixed, "{{", vbBinaryCompare)
        If opening = 0 Then Exit Do
        closing = InStr(opening + 2, fixed, "}}", vbBinaryCompare)
        If closing = 0 Then Exit Do
        fixed = Left$(fixed, opening - 1) & " # " & Mid$(fixed, closing + 2)
    Loop
    fixed = vbLf & Replace(Replace(Replace(fixed, vbCrLf, vbLf), vbCr, vbLf), vbTab, " ") & vbLf
    For Each w In Array("NOFORN", "ORCON", "REL TO", "CONFIDENTIAL", "SECRET", "C O N F I D E N T I A L", "S E C R E T")
        p = InStr(1, fixed, CStr(w), vbBinaryCompare)
        Do While p > 0
            If Not IsLetter(Mid$(fixed, p - 1, 1)) And Not IsLetter(Mid$(fixed, p + Len(w), 1)) Then
                FixedTextMarking = CStr(w): Exit Function
            End If
            p = InStr(p + 1, fixed, CStr(w), vbBinaryCompare)
        Loop
    Next
    p = InStr(1, fixed, "(", vbBinaryCompare)
    Do While p > 0
        q = InStr(p + 1, fixed, ")", vbBinaryCompare)
        If q = 0 Then Exit Do
        inner = Mid$(fixed, p + 1, q - p - 1)
        If InStr(1, " ." & vbLf, Mid$(fixed, p - 1, 1), vbBinaryCompare) > 0 And _
            InStr(1, " " & vbLf, Mid$(fixed, q + 1, 1), vbBinaryCompare) > 0 Then
            If StrComp(inner, "TS", vbBinaryCompare) = 0 Or StrComp(Left$(inner, 2), "C/", vbBinaryCompare) = 0 Or _
                    StrComp(Left$(inner, 2), "S/", vbBinaryCompare) = 0 Or StrComp(Left$(inner, 3), "TS/", vbBinaryCompare) = 0 Then
                FixedTextMarking = "(" & inner & ")": Exit Function
            ElseIf StrComp(inner, "C", vbBinaryCompare) = 0 Or StrComp(inner, "S", vbBinaryCompare) = 0 Then
                lineStart = InStrRev(fixed, vbLf, p - 1, vbBinaryCompare)
                If Len(Trim$(Mid$(fixed, lineStart + 1, p - lineStart - 1))) > 0 Then
                    FixedTextMarking = "(" & inner & ")": Exit Function
                End If
            End If
        End If
        p = InStr(p + 1, fixed, "(", vbBinaryCompare)
    Loop
End Function

' An upper-case letter only, as Message Generator compared (Option Compare
' Binary there): "a" next to SECRET is a letter too.
Private Function IsLetter(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsLetter = (AscW(ch) >= 65 And AscW(ch) <= 90) Or (AscW(ch) >= 97 And AscW(ch) <= 122)
End Function

' ---------------------------------------------------------------------------
' Lists
' ---------------------------------------------------------------------------

Private Function TplWhereSql(ByVal chip As Long, ByVal search As String) As String
    Dim w As String
    Select Case chip
        Case 1: w = "t.Status='" & TS_ACTIVE & "'"
        Case 2: w = "t.Status='" & TS_INACTIVE & "'"
        Case 3: w = "t.Status='" & TS_RETIRED & "'"
    End Select
    search = Trim$(search)
    If Len(search) > 0 Then
        w = w & IIf(Len(w) > 0, " AND ", "") & "(t.TemplateName & ' ' & Nz(t.MsgType,'') & ' ' & Nz(t.Description,'')) Like " & _
            Q("*" & LikeLiteral(search) & "*")
    End If
    If Len(w) > 0 Then TplWhereSql = " WHERE " & w
End Function

' Rows of the Templates page's list: id, the name with its version, status,
' type, how many fields, when a draft last started from it.
'   chip    TPL_CHIPS: 0 All, 1 Active, 2 Inactive, 3 Retired
'   search  words in the name, type or description
Public Function TemplateListSql(ByVal chip As Long, ByVal search As String) As String
    TemplateListSql = "SELECT t.TemplateID AS id, t.TemplateName & ' v' & t.VersionNo AS tpl, t.Status AS status, t.MsgType, " & _
        "(SELECT Count(*) FROM tblTemplateFields AS f WHERE f.TemplateID=t.TemplateID) AS fields_n, " & _
        "IIf(t.LastUsedOn Is Null,'Never',Format(t.LastUsedOn,'dd mmm yyyy')) AS used_txt, t.LastUsedOn, t.TemplateName, t.VersionNo " & _
        "FROM tblTemplates AS t" & TplWhereSql(chip, search) & " ORDER BY t.TemplateName, t.VersionNo DESC"
End Function

' How many templates a chip lists (no search).
Public Function TemplateCount(ByVal chip As Long) As Long
    Dim rs As DAO.Recordset
    If Not DataLinked() Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM tblTemplates AS t" & TplWhereSql(chip, ""), dbOpenSnapshot)
    TemplateCount = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' The template's fields in order: id (FieldID), the name, required, max
' length, default and hint.
Public Function TemplateFieldsSql(ByVal tid As Long) As String
    TemplateFieldsSql = "SELECT FieldID AS id, FieldName, IIf(IsRequired,'Required','Optional') AS req_txt, " & _
        "IIf(Nz(MaxLen,0)=0,'No limit',CStr(MaxLen)) AS max_txt, DefaultText, Hint FROM tblTemplateFields " & _
        "WHERE TemplateID=" & tid & " ORDER BY SortOrder"
End Function

' "NAME v2", as the lists name a template version.
Public Function TemplateLabel(ByVal tid As Long) As String
    TemplateLabel = Nz(DLookup("TemplateName & ' v' & VersionNo", "tblTemplates", "TemplateID=" & tid), "template " & tid)
End Function

' ---------------------------------------------------------------------------
' A template's values
' ---------------------------------------------------------------------------

' Template tid's values: keys name, description, msg_type, msg_class,
' precedence, classification, to_line, info_line, subject, body, and id,
' version, status, updated_on, updated_by, last_used_on. Nothing when it
' isn't there.
' Errors: the engine's.
Public Function TemplateValues(ByVal tid As Long) As Object
    Dim rs As DAO.Recordset, d As Object
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblTemplates WHERE TemplateID=" & tid, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    d("id") = tid
    d("name") = Nz(rs!TemplateName, "")
    d("version") = Nz(rs!VersionNo, 1)
    d("status") = Nz(rs!Status, TS_INACTIVE)
    d("description") = Nz(rs!Description, "")
    d("msg_type") = Nz(rs!MsgType, "")
    d("msg_class") = Nz(rs!MsgClass, "A")
    d("precedence") = Nz(rs!Precedence, "ROUTINE")
    d("classification") = Nz(rs!Classification, "UNCLASSIFIED")
    d("to_line") = Nz(rs!ToLine, "")
    d("info_line") = Nz(rs!InfoLine, "")
    d("subject") = Nz(rs!Subject, "")
    d("body") = Nz(rs!BodyText, "")
    d("updated_on") = rs!UpdatedOn
    d("updated_by") = Nz(rs!UpdatedBy, "")
    d("last_used_on") = rs!LastUsedOn
    rs.Close
    Set TemplateValues = d
End Function

' Every tag in a template's subject, addressees and text, in that order.
' Errors: TemplateTags'.
Private Function AllTags(ByVal d As Object) As Collection
    Dim c As New Collection, seen As Object, part As Variant, nm As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For Each part In Array(d("subject"), d("to_line"), d("info_line"), d("body"))
        For Each nm In TemplateTags(CStr(part))
            If Not seen.Exists(nm) Then seen.Add nm, True: c.Add nm
        Next
    Next
    Set AllTags = c
End Function

' Why the signed-in person can't change template tid, "" when they can.
Public Function TemplateEditProblem(ByVal tid As Long) As String
    Dim st As String
    st = Nz(DLookup("Status", "tblTemplates", "TemplateID=" & tid), "")
    If Len(st) = 0 Then TemplateEditProblem = "This template is no longer in the data file.": Exit Function
    If Not IsAnchorAdmin() Then TemplateEditProblem = "Only an ANCHOR admin changes templates. Start a draft from it instead.": Exit Function
    If st <> TS_INACTIVE Then TemplateEditProblem = TemplateLabel(tid) & " is " & LCase$(st) & " and doesn't change. " & _
        "New version makes a copy to edit."
End Function

' Why values d (TemplateValues' keys) can't be saved, "" when they can.
Private Function TplValuesProblem(ByVal d As Object) As String
    Dim why As Variant
    If Len(Trim$(d("name"))) = 0 Then TplValuesProblem = "Give the template a name.": Exit Function
    If Len(d("name")) > 100 Then TplValuesProblem = "The name is " & Len(d("name")) & " characters; keep it to 100.": Exit Function
    If Len(d("description")) > 255 Then TplValuesProblem = "The description is " & Len(d("description")) & " characters; keep it to 255.": Exit Function
    If Len(d("subject")) > 255 Then TplValuesProblem = "The subject is " & Len(d("subject")) & " characters; keep it to 255.": Exit Function
    If Len(d("to_line")) > 255 Or Len(d("info_line")) > 255 Then TplValuesProblem = "Keep each addressee line to 255 characters.": Exit Function
    If Len(d("msg_type")) = 0 Then TplValuesProblem = "Pick the message type.": Exit Function
    If InStr(";ROUTINE;PRIORITY;IMMEDIATE;FLASH;", ";" & d("precedence") & ";") = 0 Then TplValuesProblem = "Pick the precedence.": Exit Function
    If InStr(";UNCLASSIFIED;CUI;CONFIDENTIAL;SECRET;", ";" & d("classification") & ";") = 0 Then
        TplValuesProblem = "Pick the classification.": Exit Function
    End If
    If DCount("*", "tblMessageClasses", "ClassCode=" & Q(d("msg_class"))) = 0 Then TplValuesProblem = "Pick the message class (A to E).": Exit Function
    For Each why In Array(TemplateTagProblem(d("subject")), TemplateTagProblem(d("to_line")), TemplateTagProblem(d("info_line")), _
            TemplateTagProblem(d("body")))
        If Len(why) > 0 Then TplValuesProblem = why: Exit Function
    Next
End Function

' ---------------------------------------------------------------------------
' Changing templates (ANCHOR admins)
' ---------------------------------------------------------------------------

' Add an Inactive template, version 1, named "New template" (with a number
' when that is taken), with the command's defaults. Returns its id.
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function NewTemplate() As Long
    Dim nm As String, n As Long, d As Object, id As Long
    RequireAnchorAdmin
    nm = "New template"
    Do While DCount("*", "tblTemplates", "TemplateName=" & Q(nm)) > 0
        n = n + 1: nm = "New template " & n
    Loop
    Set d = NewDraftValues()
    id = InsertRow("INSERT INTO tblTemplates (TemplateName, VersionNo, Description, MsgType, MsgClass, Precedence, Classification, " & _
        "ToLine, InfoLine, Subject, BodyText, Status, CreatedOn, CreatedBy, UpdatedOn, UpdatedBy) VALUES (" & Q(nm) & ", 1, '', " & _
        Q(d("msg_type")) & ", 'A', 'ROUTINE', 'UNCLASSIFIED', '', '', '', '', '" & TS_INACTIVE & "', Now(), " & _
        Q(Left$(SignedInName(), 100)) & ", Now(), " & Q(Left$(SignedInName(), 100)) & ")")
    AuditLog "Create template", "Template (" & nm & " v1)", "New template, inactive."
    NewTemplate = id
End Function

' Save values d (TemplateValues' keys) to the Inactive template tid and
' bring its field rows in line with its tags. Returns "" when saved, else
' why not (to show).
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function SaveTemplate(ByVal tid As Long, ByVal d As Object) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, ver As Long, tags As Collection, nm As String
    why = TemplateEditProblem(tid)
    If Len(why) = 0 Then why = TplValuesProblem(d)
    If Len(why) > 0 Then SaveTemplate = why: Exit Function
    nm = Trim$(d("name"))
    ver = Nz(DLookup("VersionNo", "tblTemplates", "TemplateID=" & tid), 1)
    If DCount("*", "tblTemplates", "TemplateName=" & Q(nm) & " AND VersionNo=" & ver & " AND TemplateID<>" & tid) > 0 Then
        SaveTemplate = "There is already a " & nm & " v" & ver & ". Pick another name.": Exit Function
    End If
    Set tags = AllTags(d)
    If tags.Count > 100 Then SaveTemplate = "A template holds at most 100 fields; this one has " & tags.Count & ".": Exit Function
    RequireAnchorAdmin
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If KitExecute("UPDATE tblTemplates SET TemplateName=" & Q(nm) & ", Description=" & Q(Trim$(d("description"))) & _
            ", MsgType=" & Q(d("msg_type")) & ", MsgClass=" & Q(d("msg_class")) & ", Precedence=" & Q(d("precedence")) & _
            ", Classification=" & Q(d("classification")) & ", ToLine=" & Q(d("to_line")) & ", InfoLine=" & Q(d("info_line")) & _
            ", Subject=" & Q(d("subject")) & ", BodyText=" & Q(d("body")) & ", UpdatedOn=Now(), UpdatedBy=" & Q(Left$(SignedInName(), 100)) & _
            " WHERE TemplateID=" & tid & " AND Status='" & TS_INACTIVE & "'") = 0 Then
        ws.Rollback: inTrans = False
        SaveTemplate = "Someone activated or removed this template first. Nothing was saved."
        Exit Function
    End If
    SyncFields tid, tags
    AuditLog "Change template", "Template (" & nm & " v" & ver & ")", "Saved, " & Plural(tags.Count, "field") & "."
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modTemplates.SaveTemplate", tid
End Function

' Bring template tid's field rows in line with tags: keep a field still
' there (its rules, in its new place), add a new one (required, no limit),
' remove one gone.
Private Sub SyncFields(ByVal tid As Long, ByVal tags As Collection)
    Dim keep As Object, rs As DAO.Recordset, nm As Variant, n As Long, fid As Variant
    Set keep = CreateObject("Scripting.Dictionary")
    keep.CompareMode = vbTextCompare
    For Each nm In tags
        keep(nm) = True
    Next
    Set rs = CurrentDb.OpenRecordset("SELECT FieldID, FieldName FROM tblTemplateFields WHERE TemplateID=" & tid, dbOpenSnapshot)
    Do While Not rs.EOF
        If Not keep.Exists(CStr(rs!FieldName)) Then KitExecute "DELETE FROM tblTemplateFields WHERE FieldID=" & rs!FieldID
        rs.MoveNext
    Loop
    rs.Close
    For Each nm In tags
        n = n + 1
        fid = DLookup("FieldID", "tblTemplateFields", "TemplateID=" & tid & " AND FieldName=" & Q(nm))
        If IsNull(fid) Then
            KitExecute "INSERT INTO tblTemplateFields (TemplateID, FieldName, SortOrder, IsRequired, MaxLen, DefaultText, Hint) VALUES (" & _
                tid & ", " & Q(nm) & ", " & n & ", True, 0, '', '')"
        Else
            KitExecute "UPDATE tblTemplateFields SET FieldName=" & Q(nm) & ", SortOrder=" & n & " WHERE FieldID=" & fid
        End If
    Next
End Sub

' Set the rules of field fieldId on an Inactive template (dlgTemplateField).
' Returns "" when saved, else why not: a max length that isn't a whole
' number up to 65535, a default longer than it, a hint over TPL_HINT_MAX.
'   maxText  the typed max length, "" or 0 for no limit
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function SaveTemplateField(ByVal fieldId As Long, ByVal required As Boolean, ByVal maxText As String, _
        ByVal defaultText As String, ByVal hint As String) As String
    Dim tid As Long, maxLen As Long, why As String, nm As String
    tid = Nz(DLookup("TemplateID", "tblTemplateFields", "FieldID=" & fieldId), 0)
    If tid = 0 Then SaveTemplateField = "This field is no longer on the template.": Exit Function
    why = TemplateEditProblem(tid)
    If Len(why) > 0 Then SaveTemplateField = why: Exit Function
    hint = Trim$(Replace(Replace(hint, vbCr, " "), vbLf, " "))
    If Len(hint) > TPL_HINT_MAX Then SaveTemplateField = "The hint is " & Len(hint) & " characters; keep it to " & TPL_HINT_MAX & ".": Exit Function
    maxText = Trim$(maxText)
    If Len(maxText) = 0 Then maxText = "0"
    If Len(maxText) > 5 Or maxText Like "*[!0-9]*" Then SaveTemplateField = "Type the max length as a whole number, or 0 for no limit.": Exit Function
    maxLen = CLng(maxText)
    If maxLen > 65535 Then SaveTemplateField = "The max length can be at most 65535.": Exit Function
    If Len(defaultText) > 255 Then SaveTemplateField = "Keep the default to 255 characters.": Exit Function
    If maxLen > 0 And Len(defaultText) > maxLen Then
        SaveTemplateField = "The default is " & Len(defaultText) & " characters, longer than the max length.": Exit Function
    End If
    nm = Nz(DLookup("FieldName", "tblTemplateFields", "FieldID=" & fieldId), "")
    RequireAnchorAdmin
    KitExecute "UPDATE tblTemplateFields SET IsRequired=" & IIf(required, "True", "False") & ", MaxLen=" & maxLen & _
        ", DefaultText=" & Q(defaultText) & ", Hint=" & Q(hint) & " WHERE FieldID=" & fieldId
    AuditLog "Change template", "Template (" & TemplateLabel(tid) & ")", "Field " & nm & ": " & IIf(required, "required", "optional") & _
        ", " & IIf(maxLen = 0, "no limit", "at most " & maxLen & " characters") & "."
End Function

' Why template tid can't be activated, "" when it can: it is Inactive and
' saved with a name, its tags parse and its fixed text carries no
' classification marking.
Public Function TemplateActivateProblem(ByVal tid As Long) As String
    Dim d As Object, why As String, part As Variant, mk As String
    why = TemplateEditProblem(tid)
    If Len(why) > 0 Then TemplateActivateProblem = why: Exit Function
    Set d = TemplateValues(tid)
    why = TplValuesProblem(d)
    If Len(why) > 0 Then TemplateActivateProblem = why: Exit Function
    If Len(Trim$(d("body"))) = 0 Then TemplateActivateProblem = "The template has no text yet.": Exit Function
    For Each part In Array(d("subject"), d("to_line"), d("info_line"), d("body"))
        mk = FixedTextMarking(CStr(part))
        If Len(mk) > 0 Then
            TemplateActivateProblem = "Its fixed text carries the marking " & mk & ". A template's fixed text goes into every draft, " & _
                "whatever the draft's classification, so leave markings to the drafter (a field, or the draft itself)."
            Exit Function
        End If
    Next
End Function

' Activate the Inactive template tid: the name's Active version, if any, is
' retired in the same step. Returns "" when done, else why not.
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function ActivateTemplate(ByVal tid As Long) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, why As String, nm As String, n As Long
    why = TemplateActivateProblem(tid)
    If Len(why) > 0 Then ActivateTemplate = why: Exit Function
    nm = DLookup("TemplateName", "tblTemplates", "TemplateID=" & tid)
    RequireAnchorAdmin
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    n = KitExecute("UPDATE tblTemplates SET Status='" & TS_RETIRED & "', UpdatedOn=Now(), UpdatedBy=" & Q(Left$(SignedInName(), 100)) & _
        " WHERE TemplateName=" & Q(nm) & " AND Status='" & TS_ACTIVE & "'")
    If KitExecute("UPDATE tblTemplates SET Status='" & TS_ACTIVE & "', UpdatedOn=Now(), UpdatedBy=" & Q(Left$(SignedInName(), 100)) & _
            " WHERE TemplateID=" & tid & " AND Status='" & TS_INACTIVE & "'") = 0 Then
        ws.Rollback: inTrans = False
        ActivateTemplate = "Someone changed this template first. Nothing was changed."
        Exit Function
    End If
    AuditLog "Activate template", "Template (" & TemplateLabel(tid) & ")", "Activated" & IIf(n > 0, "; the earlier active version was retired.", ".")
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modTemplates.ActivateTemplate", tid
End Function

' Retire the Active template tid: no new draft starts from it. Returns ""
' when done, else why not.
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function RetireTemplate(ByVal tid As Long) As String
    If Not IsAnchorAdmin() Then RetireTemplate = "Only an ANCHOR admin changes templates.": Exit Function
    RequireAnchorAdmin
    If KitExecute("UPDATE tblTemplates SET Status='" & TS_RETIRED & "', UpdatedOn=Now(), UpdatedBy=" & Q(Left$(SignedInName(), 100)) & _
            " WHERE TemplateID=" & tid & " AND Status='" & TS_ACTIVE & "'") = 0 Then
        RetireTemplate = "Only an active template is retired, and this one isn't active now."
        Exit Function
    End If
    AuditLog "Retire template", "Template (" & TemplateLabel(tid) & ")", "Retired: no new draft starts from it."
End Function

' Copy template tid as the name's next version, Inactive, with its field
' rules, to edit. Returns the new id, or 0 with why in problem.
' Errors: ERR_NOT_ALLOWED for a non-admin; the engine's.
Public Function NewTemplateVersion(ByVal tid As Long, ByRef problem As String) As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, nm As String, ver As Long, id As Long, other As Variant
    problem = ""
    If Not IsAnchorAdmin() Then problem = "Only an ANCHOR admin changes templates.": Exit Function
    nm = Nz(DLookup("TemplateName", "tblTemplates", "TemplateID=" & tid), "")
    If Len(nm) = 0 Then problem = "This template is no longer in the data file.": Exit Function
    other = DLookup("VersionNo", "tblTemplates", "TemplateName=" & Q(nm) & " AND Status='" & TS_INACTIVE & "'")
    If Not IsNull(other) Then problem = nm & " v" & other & " is already being edited. Finish and activate it, or edit it further.": Exit Function
    ver = Nz(DMax("VersionNo", "tblTemplates", "TemplateName=" & Q(nm)), 0) + 1
    RequireAnchorAdmin
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    id = InsertRow("INSERT INTO tblTemplates (TemplateName, VersionNo, Description, MsgType, MsgClass, Precedence, Classification, " & _
        "ToLine, InfoLine, Subject, BodyText, Status, CreatedOn, CreatedBy, UpdatedOn, UpdatedBy) SELECT TemplateName, " & ver & _
        ", Description, MsgType, MsgClass, Precedence, Classification, ToLine, InfoLine, Subject, BodyText, '" & TS_INACTIVE & _
        "', Now(), " & Q(Left$(SignedInName(), 100)) & ", Now(), " & Q(Left$(SignedInName(), 100)) & " FROM tblTemplates WHERE TemplateID=" & tid)
    KitExecute "INSERT INTO tblTemplateFields (TemplateID, FieldName, SortOrder, IsRequired, MaxLen, DefaultText, Hint) SELECT " & id & _
        ", FieldName, SortOrder, IsRequired, MaxLen, DefaultText, Hint FROM tblTemplateFields WHERE TemplateID=" & tid
    AuditLog "Create template", "Template (" & nm & " v" & ver & ")", "New version, copied from v" & (ver - 1) & "; inactive until activated."
    ws.CommitTrans: inTrans = False
    NewTemplateVersion = id
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modTemplates.NewTemplateVersion", tid
End Function

' ---------------------------------------------------------------------------
' Preview and starting a draft
' ---------------------------------------------------------------------------

' Draft values (modDrafting.NewDraftValues' keys) for template values t with
' field values vals filled in (a dictionary by field name; Nothing shows
' each field as [Name]).
' Errors: TemplateTags' for a malformed tag.
Public Function TemplateDraftValues(ByVal t As Object, ByVal vals As Object) As Object
    Dim d As Object
    Set d = NewDraftValues()
    d("subject") = FillTemplateText(t("subject"), vals)
    d("msg_type") = t("msg_type")
    d("msg_class") = t("msg_class")
    d("precedence") = t("precedence")
    d("classification") = t("classification")
    d("to_line") = FillTemplateText(t("to_line"), vals)
    d("info_line") = FillTemplateText(t("info_line"), vals)
    d("paragraphs") = FillTemplateText(t("body"), vals)
    Set TemplateDraftValues = d
End Function

' The message template values t would make, as it would transmit, with
' each field as its default or [Name], or why it can't be shown.
Public Function TemplatePreviewText(ByVal t As Object) As String
    Dim why As String, vals As Object, rs As DAO.Recordset
    why = TplValuesProblem(t)
    If Len(why) > 0 Then TemplatePreviewText = "No preview yet: " & why: Exit Function
    Set vals = CreateObject("Scripting.Dictionary")
    vals.CompareMode = vbTextCompare
    If t.Exists("id") Then
        Set rs = CurrentDb.OpenRecordset("SELECT FieldName, DefaultText FROM tblTemplateFields WHERE TemplateID=" & t("id"), dbOpenSnapshot)
        Do While Not rs.EOF
            If Len(Nz(rs!DefaultText, "")) > 0 Then vals(CStr(rs!FieldName)) = rs!DefaultText
            rs.MoveNext
        Loop
        rs.Close
    End If
    TemplatePreviewText = DraftPreviewText(TemplateDraftValues(t, vals))
End Function

' Why a value typed for a field is refused, "" when it is fine.
Private Function FieldValueProblem(ByVal nm As String, ByVal v As String, ByVal required As Boolean, ByVal maxLen As Long) As String
    If required And Len(Trim$(v)) = 0 Then FieldValueProblem = nm & " is required.": Exit Function
    If maxLen > 0 And Len(v) > maxLen Then FieldValueProblem = nm & " is " & Len(v) & " characters; keep it to " & maxLen & ".": Exit Function
    If InStr(v, "{{") > 0 Or InStr(v, "}}") > 0 Then FieldValueProblem = nm & " can't hold {{ or }}.": Exit Function
End Function

' Start a draft for the signed-in person from the Active template tid:
' asks for each field's value in turn (its hint, default and rules),
' then adds the draft (modDrafting.CreateDraft). Returns the draft's id, 0
' when the person cancelled, or 0 with why in problem.
' Errors: CreateDraft's; the engine's.
Public Function StartDraftFromTemplate(ByVal tid As Long, ByRef problem As String) As Long
    Dim t As Object, rs As DAO.Recordset, vals As Object, v As Variant, nm As String, why As String, ask As String, id As Long
    problem = ""
    If Nz(DLookup("Status", "tblTemplates", "TemplateID=" & tid), "") <> TS_ACTIVE Then
        problem = "Drafts start from an active template, and " & TemplateLabel(tid) & " isn't active.": Exit Function
    End If
    Set t = TemplateValues(tid)
    why = TplValuesProblem(t)
    If Len(why) > 0 Then problem = why: Exit Function
    Set vals = CreateObject("Scripting.Dictionary")
    vals.CompareMode = vbTextCompare
    Set rs = CurrentDb.OpenRecordset("SELECT FieldName, IsRequired, MaxLen, DefaultText, Hint FROM tblTemplateFields WHERE TemplateID=" & _
        tid & " ORDER BY SortOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        nm = rs!FieldName
        v = Nz(rs!DefaultText, "")
        ask = nm & IIf(rs!IsRequired, " (required)", " (optional)") & IIf(Nz(rs!MaxLen, 0) > 0, ", up to " & rs!MaxLen & " characters", "") & _
            IIf(Len(Nz(rs!Hint, "")) > 0, vbCrLf & rs!Hint, "")
        Do
            v = KitInputBox(ask, "Start a draft: " & ClipText(t("name"), 40), CStr(v))
            If IsNull(v) Then rs.Close: Exit Function
            why = FieldValueProblem(nm, CStr(v), rs!IsRequired, Nz(rs!MaxLen, 0))
            If Len(why) = 0 Then Exit Do
            KitMsgBox why, vbExclamation, "Start a draft"
        Loop
        vals(nm) = CStr(v)
        rs.MoveNext
    Loop
    rs.Close
    id = CreateDraft(TemplateDraftValues(t, vals), tid)
    KitExecute "UPDATE tblTemplates SET LastUsedOn=Now() WHERE TemplateID=" & tid
    StartDraftFromTemplate = id
End Function

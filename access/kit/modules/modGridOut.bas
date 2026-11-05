Option Compare Database
Option Explicit

' Access UI kit - print and export a BuildGrid table as the user sees it: the
' visible columns with their on-screen captions, and the rows its record
' source, filter and sort show. PrintGrid builds a report (zz_rptGrid,
' replaced each time) and opens its preview. ExportGrid writes CSV through
' modCsv, or an Excel workbook. Neither leaves a query behind.
'
' Both read the grid form: pass the subform's form, e.g. Me.sfList.Form. A
' filter must name fields without a table prefix ("status='Open'").

Private Const RPT_NAME As String = "zz_rptGrid"
Private Const TMP_QUERY As String = "zz_qryGridExport"
Private Const LANDSCAPE_W As Long = 14400      ' 11 in less two 0.5 in margins (twips)

' Print preview of the grid's rows: title and subtitle at the top, the
' columns' captions, the rows, and page numbers. The columns are scaled to
' the page, which turns landscape when they are wider than portrait.
Public Sub PrintGrid(gridForm As Form, ByVal title As String, Optional ByVal subtitle As String = "")
    Dim cols As Collection, col As Variant, sql As String, total As Long, pageW As Long, ratio As Double
    Dim x As Long, cw As Long, i As Long, c As Control, n As Long, d As String, s As String
    On Error GoTo Fail
    Set cols = GridColumns(gridForm)
    For Each col In cols
        total = total + col(2)
    Next
    pageW = IIf(total > REPORT_W, LANDSCAPE_W, REPORT_W)
    ratio = pageW / total
    sql = GridSql(gridForm, cols, False)
    On Error Resume Next
    DoCmd.Close acReport, RPT_NAME, acSaveNo
    On Error GoTo Fail
    Application.Echo False
    NewReport sql, pageW
    F.HasModule = False
    If pageW = LANDSCAPE_W Then F.Printer.Orientation = acPRORLandscape
    F.Caption = title
    F.Section(acDetail).Height = px(24)
    F.Section(acDetail).BackColor = PRINT_PAPER
    F.Section(acDetail).AlternateBackColor = PRINT_ROW1
    F.Section(acPageHeader).Height = px(80)
    F.Section(acPageFooter).Height = px(26)
    Lbl "lblTitle", title, 0, 0, pageW, px(28), 14, True, PRINT_TITLE, acPageHeader
    Set c = Lbl("lblSub", subtitle, 0, px(28), pageW - px(160), px(18), 9, False, PRINT_MUTED, acPageHeader)
    Set c = RO("txtDate", "=Format(Date(),""dd mmm yyyy"")", pageW - px(160), px(28), px(160), px(18), 9, False, PRINT_MUTED, acPageHeader)
    c.TextAlign = 3
    Box "boxLine", 0, px(76), pageW, px(1), PRINT_FG, -1, acPageHeader
    For Each col In cols
        cw = CLng(col(2) * ratio)
        Set c = Lbl("lblCol" & i, UCase$(col(1)), x + px(4), px(56), cw - px(8), px(16), 8, True, PRINT_MUTED, acPageHeader)
        c.TextAlign = col(4)
        Set c = RO("txtCol" & i, "col_" & i, x + px(4), px(3), cw - px(8), px(18), 9, False, PRINT_FG, acDetail)
        c.TextAlign = col(4)
        c.Format = col(3)
        c.CanGrow = True
        x = x + cw
        i = i + 1
    Next
    F.Section(acDetail).CanGrow = True
    Set c = RO("txtPage", "=""Page "" & [Page] & "" of "" & [Pages]", 0, px(6), pageW, px(18), 8, False, PRINT_MUTED, acPageFooter)
    c.TextAlign = 3
    SaveForm RPT_NAME
    Application.Echo True
    DoCmd.OpenReport RPT_NAME, acViewPreview
    Exit Sub
Fail:
    n = Err.Number: d = Err.Description: s = Err.Source
    Application.Echo True
    DiscardReport
    Err.Number = n: Err.Description = d: Err.Source = s
    RaiseAgain "modGridOut.PrintGrid", title
End Sub

' Write the grid's rows to path, with the captions as column names: an Excel
' workbook when path ends in .xlsx, else CSV. Replaces an existing file.
' Returns the number of rows written.
Public Function ExportGrid(gridForm As Form, ByVal path As String) As Long
    Dim sql As String, n As Long, d As String, s As String
    KitBusy "Exporting..."
    On Error GoTo Fail
    sql = GridSql(gridForm, GridColumns(gridForm), True)
    If FileExists(path) Then Kill path
    If LCase$(Right$(path, 5)) = ".xlsx" Then
        DropQuery
        CurrentDb.CreateQueryDef TMP_QUERY, sql
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, TMP_QUERY, path, True
        ExportGrid = DCount("*", TMP_QUERY)
        DropQuery
    Else
        ExportGrid = ExportSqlToCsv(sql, path)
    End If
    KitBusyEnd
    Exit Function
Fail:
    ' Clean-up resets Err, so keep the error and put it back.
    n = Err.Number: d = Err.Description: s = Err.Source
    DropQuery
    KitBusyEnd
    Err.Number = n: Err.Description = d: Err.Source = s
    RaiseAgain "modGridOut.ExportGrid", "file " & path
End Function

' The grid's visible columns, left to right: Array(source, caption, width,
' format, text align), read from the columns BuildGrid made (txtCol0, ...).
' Ordered by position, since the column chooser (modGridCols) can move them.
Private Function GridColumns(gridForm As Form) As Collection
    Dim out As New Collection, lefts As New Collection, c As Control, i As Long, k As Long, t() As String, src As String
    Do
        Set c = Nothing
        On Error Resume Next
        Set c = gridForm.Controls("txtCol" & i)
        On Error GoTo 0
        If c Is Nothing Then Exit Do
        If c.Visible Then
            t = Split(c.Tag & "||", "|")
            src = c.ControlSource
            ' A search column (BuildGrid option "s") shows the marked text of
            ' hidden txtSrc<i>: use that control's field instead.
            If Left$(src, 10) = "=GridMark(" Then src = gridForm.Controls("txtSrc" & i).ControlSource
            For k = 1 To lefts.Count
                If lefts(k) > c.Left Then Exit For
            Next
            If k > lefts.Count Then
                out.Add Array(src, IIf(Len(t(0)) > 0, t(0), src), c.Width + px(8), c.Format, c.TextAlign)
                lefts.Add c.Left
            Else
                out.Add Array(src, IIf(Len(t(0)) > 0, t(0), src), c.Width + px(8), c.Format, c.TextAlign), , k
                lefts.Add c.Left, , k
            End If
        End If
        i = i + 1
    Loop
    If out.Count = 0 Then Err.Raise vbObjectError + 1, "modGridOut", "The table " & gridForm.Name & " has no columns to print or export."
    Set GridColumns = out
End Function

' SQL for the rows the grid shows: its record source, filter and sort, with
' a column per grid column, named by its caption (named) or col_<n>.
Private Function GridSql(gridForm As Form, cols As Collection, ByVal named As Boolean) As String
    Dim src As String, sel As String, col As Variant, i As Long, pos As Long, ord As String, nm As String
    Dim extra As Variant, f As String
    src = Trim$(gridForm.RecordSource)
    If Right$(src, 1) = ";" Then src = Trim$(Left$(src, Len(src) - 1))
    If UCase$(Left$(src, 7)) = "SELECT " Then
        ' The sort moves to the outer query, which is the only place Access
        ' keeps it. Sort terms the query doesn't return are added to its
        ' select list as zz_sort<n>.
        pos = InStrRev(UCase$(src), " ORDER BY ")
        If pos > 0 Then
            ord = OuterOrder(Mid$(src, pos + 10), gridForm, extra)
            src = Left$(src, pos - 1)
            If Len(extra) > 0 Then
                pos = InStr(UCase$(src), " FROM ")
                If pos > 0 Then src = Left$(src, pos - 1) & extra & Mid$(src, pos) Else ord = ""
            End If
        End If
        src = "(" & src & ") AS q"
    Else
        src = "[" & src & "] AS q"
    End If
    For Each col In cols
        If named Then nm = Replace(Replace(col(1), "[", "("), "]", ")") Else nm = "col_" & i
        If Left$(col(0), 1) = "=" Then
            sel = sel & ", (" & Mid$(col(0), 2) & ") AS [" & nm & "]"
        Else
            ' Qualified, because a caption may repeat a field's name. A source
            ' given as "[field]" (BuildGrid "[SHA-256]|...") loses its brackets first.
            f = col(0)
            If Left$(f, 1) = "[" And Right$(f, 1) = "]" Then f = Mid$(f, 2, Len(f) - 2)
            sel = sel & ", q.[" & f & "] AS [" & nm & "]"
        End If
        i = i + 1
    Next
    GridSql = "SELECT " & Mid$(sel, 3) & " FROM " & src
    If gridForm.FilterOn And Len(gridForm.Filter) > 0 Then GridSql = GridSql & " WHERE " & gridForm.Filter
    If gridForm.OrderByOn And Len(gridForm.OrderBy) > 0 Then ord = OuterOrder(gridForm.OrderBy, gridForm)
    If Len(ord) > 0 Then GridSql = GridSql & " ORDER BY " & ord
End Function

' A sort clause for the outer query: "t.name DESC, [due]" -> "q.[name] DESC, q.[due]".
' A term that isn't a field the grid's record source returns (a hidden field
' such as "ORDER BY due_date" with only Format(due_date) AS due selected, or
' an expression) is added to extra as ", <term> AS [zz_sort<n>]" for the
' inner select list, and sorted by that alias. Without a ByRef extra (the
' form's own OrderBy, which names returned fields), such a term is dropped.
Private Function OuterOrder(ByVal clause As String, gridForm As Form, Optional ByRef extra As Variant) As String
    Dim term As Variant, t As String, f As String, dirn As String, out As String, n As Long
    If Not IsMissing(extra) Then extra = ""
    For Each term In Split(clause, ",")
        t = Trim$(term)
        dirn = ""
        If UCase$(Right$(t, 5)) = " DESC" Then dirn = " DESC": t = Trim$(Left$(t, Len(t) - 5))
        If UCase$(Right$(t, 4)) = " ASC" Then t = Trim$(Left$(t, Len(t) - 4))
        f = ""
        If InStr(t, "(") = 0 Then
            f = t
            If InStrRev(f, ".") > 0 Then f = Mid$(f, InStrRev(f, ".") + 1)
            f = Replace(Replace(f, "[", ""), "]", "")
            If Not HasField(gridForm, f) Then f = ""
        End If
        If Len(f) > 0 Then
            out = out & ", q.[" & f & "]" & dirn
        ElseIf Not IsMissing(extra) Then
            extra = extra & ", " & t & " AS [zz_sort" & n & "]"
            out = out & ", q.[zz_sort" & n & "]" & dirn
            n = n + 1
        End If
    Next
    OuterOrder = Mid$(out, 3)
End Function

' True when the grid's record source returns a field named nm.
Private Function HasField(gridForm As Form, ByVal nm As String) As Boolean
    Dim fld As Object
    For Each fld In gridForm.RecordsetClone.Fields
        If StrComp(fld.Name, nm, vbTextCompare) = 0 Then HasField = True: Exit Function
    Next
End Function

' Delete the export query if it is there.
Private Sub DropQuery()
    On Error Resume Next
    CurrentDb.QueryDefs.Delete TMP_QUERY
End Sub

' Close a report PrintGrid was building when it failed, without saving.
Private Sub DiscardReport()
    On Error Resume Next
    If Not F Is Nothing Then DoCmd.Close acReport, F.Name, acSaveNo
    Set F = Nothing
End Sub

Option Compare Database
Option Explicit

' Access UI kit - responsive tables, counted filters and opt-in footer paging.
' Table controls carry Tag "Pagination:on" when a footer is wanted. The caller
' supplies a unique key field, and report/export callers use TableFullFilter.
Private m_pages As Object
Private m_geometry As Object
Private m_busy As Boolean

Public Function CountedCaption(ByVal caption As String, ByVal count As Long) As String
    CountedCaption = caption & " (" & count & ")"
End Function

' Finite filter groups use matching widths and one consistent caption/count line.
Public Sub LayoutFilterRow(ByVal host As Form, ByVal names As Variant, ByVal left As Long, ByVal top As Long, ByVal available As Long, Optional ByVal size As Long = 9)
    Dim name As Variant, width As Long, measured As Long, count As Long, c As Control, x As Long
    count = UBound(names) - LBound(names) + 1
Measure:
    width = 0
    For Each name In names
        measured = PillTextWidth(host.Controls(CStr(name)).Caption, size, px(16))
        If measured > width Then width = measured
    Next
    If width * count + Sp(SP_ITEM_GAP) * (count - 1) > available Then
        If size > 7 Then size = size - 1: GoTo Measure
        Err.Raise 5, , "Filter captions need a wider row"
    End If
    x = left
    For Each name In names
        Set c = host.Controls(CStr(name))
        c.Left = x: c.Top = top: c.Width = width: c.Height = px(30)
        c.FontSize = size: c.TextAlign = 2
        x = x + width + Sp(SP_ITEM_GAP)
    Next
End Sub

' Numbered timelines keep circles fixed; only spacing and connector lengths grow.
' Controls are prefixMarkerN/NumberN/LabelN/DateN/ConnectorN, N starting at zero.
Public Sub LayoutNumberedTimeline(ByVal host As Form, ByVal prefix As String, ByVal count As Long, ByVal left As Long, ByVal width As Long, Optional ByVal markerSize As Long = 420)
    Dim slot As Long, center As Long, i As Long
    If count < 1 Or width < count * (markerSize + 60) Then Err.Raise 5, , "Timeline needs more space"
    slot = width \ count
    For i = 0 To count - 1
        center = left + i * slot + slot \ 2
        With host.Controls(prefix & "Marker" & i)
            .Left = center - markerSize \ 2: .Width = markerSize: .Height = markerSize
        End With
        With host.Controls(prefix & "Number" & i)
            .Left = center - markerSize \ 2: .Width = markerSize: .Height = markerSize: .TextAlign = 2
        End With
        host.Controls(prefix & "Label" & i).Left = left + i * slot
        host.Controls(prefix & "Label" & i).Width = slot - 60
        host.Controls(prefix & "Date" & i).Left = left + i * slot
        host.Controls(prefix & "Date" & i).Width = slot - 60
        If i < count - 1 Then
            host.Controls(prefix & "Connector" & i).Left = center + markerSize \ 2 + 30
            host.Controls(prefix & "Connector" & i).Width = slot - markerSize - 60
        End If
    Next
End Sub

' All categorical tags in a row or column use the widest measured label.
Public Function TagWidth(ByVal values As Variant, Optional ByVal size As Long = 9) As Long
    Dim value As Variant, w As Long
    For Each value In values
        w = PillTextWidth(value, size)
        If w > TagWidth Then TagWidth = w
    Next
End Function

Public Function TableTagWidth(ByVal grid As Form, ByVal field As String, Optional ByVal size As Long = 9) As Long
    Dim rows As DAO.Recordset, w As Long
    Set rows = grid.Recordset.Clone
    If Not rows.EOF Then rows.MoveFirst
    Do While Not rows.EOF
        w = PillTextWidth(rows.Fields(field).Value, size)
        If w > TableTagWidth Then TableTagWidth = w
        rows.MoveNext
    Loop
    rows.Close
End Function

' Fixed-height rows show one centered line; the editor retains the full value.
Public Sub CenterTableCell(ByVal grid As Form, ByVal cell As Control)
    If Not cell.Visible Then Exit Sub
    If cell.Name = "txtBg" Or cell.Name = "txtFocusSink" Or cell.Name = "txtSel" Then Exit Sub
    Select Case cell.ControlType
        Case acTextBox, acComboBox
            cell.Height = CLng(cell.FontSize) * 27
            cell.Top = (grid.Section(acDetail).Height - cell.Height) \ 2
            cell.TopMargin = 0: cell.BottomMargin = 0: cell.LeftMargin = 90: cell.RightMargin = 90
        Case acImage, acCommandButton
            cell.Top = (grid.Section(acDetail).Height - cell.Height) \ 2
    End Select
End Sub

' Headers use the actual runtime positions, widths and text justification.
' Header controls are prefix0, prefix1, ...; cells lists their matching names.
Public Sub AlignGridHeaders(ByVal host As Form, ByVal subName As String, ByVal prefix As String, ByVal cells As Variant)
    Dim box As Control, cell As Control, heading As Control, i As Long
    Set box = host.Controls(subName)
    For i = LBound(cells) To UBound(cells)
        Set cell = box.Form.Controls(CStr(cells(i)))
        Set heading = host.Controls(prefix & i)
        heading.Left = box.Left + cell.Left: heading.Width = cell.Width
        heading.TextAlign = cell.TextAlign: heading.Visible = cell.Visible
        heading.LeftMargin = cell.LeftMargin: heading.RightMargin = cell.RightMargin
    Next
End Sub

' Explicit columns: field|caption|minimum px|options, as in BuildGrid.
' Extra width goes primarily to long-text fields (option t); minimum tag widths
' should be supplied from TagWidth/TableTagWidth. Headers follow these cells.
Public Sub ResizeGrid(ByVal grid As Form, ByVal width As Long, ByVal cols As Variant)
    Dim i As Long, p() As String, minimum As Long, extra As Long, weights As Long
    Dim x As Long, w As Long, cell As Control
    For i = LBound(cols) To UBound(cols)
        p = Split(cols(i), "|")
        minimum = minimum + px(Val(p(2)))
        weights = weights + IIf(InStr(p(3), "t") > 0, 3, 1)
    Next
    If width < minimum + px(16) Then width = minimum + px(16): grid.ScrollBars = 3 Else grid.ScrollBars = 2
    extra = width - px(16) - minimum
    If extra < 0 Then extra = 0
    x = px(8)
    For i = LBound(cols) To UBound(cols)
        p = Split(cols(i), "|")
        Set cell = grid.Controls("txtCol" & i)
        w = px(Val(p(2))) + extra * IIf(InStr(p(3), "t") > 0, 3, 1) \ weights
        If i = UBound(cols) Then w = width - px(8) - x
        If w < px(8) Then w = px(8)
        cell.Left = x: cell.Width = w - px(8)
        CenterTableCell grid, cell
        x = x + w
    Next
    grid.Width = IIf(x > width, x, width)
    grid.txtBg.Width = grid.Width: grid.boxRowLine.Width = grid.Width
End Sub

Public Function TablePaginationEnabled(ByVal box As Control) As Boolean
    TablePaginationEnabled = InStr(1, ";" & box.Tag & ";", ";Pagination:on;", vbTextCompare) > 0
End Function

' Changes only the pagination token, preserving all other control tags.
Public Sub SetTablePagination(ByVal box As Control, ByVal enabled As Boolean)
    Dim tag As String
    tag = Replace(";" & box.Tag & ";", ";Pagination:on;", ";", 1, -1, vbTextCompare)
    Do While Left$(tag, 1) = ";": tag = Mid$(tag, 2): Loop
    Do While Len(tag) > 0 And Right$(tag, 1) = ";": tag = Left$(tag, Len(tag) - 1): Loop
    If enabled Then tag = tag & IIf(Len(tag) > 0, ";", "") & "Pagination:on"
    box.Tag = tag
End Sub

' Use instead of grid.Filter when exporting or printing a paged table.
Public Function TableFullFilter(ByVal grid As Form) As String
    Dim key As Variant, state As Object
    If Not m_pages Is Nothing Then
        For Each key In m_pages.Keys
            Set state = m_pages(key)
            If state("Grid") Is grid Then
                If state("Last") = grid.Filter Then TableFullFilter = state("Base"): Exit Function
            End If
        Next
    End If
    If grid.FilterOn Then TableFullFilter = grid.Filter
End Function

Private Function TableKey(ByVal host As Form, ByVal subName As String) As String
    TableKey = host.Name & "." & host.hWnd & "." & subName
End Function

Private Function KeyLiteral(ByVal value As Variant) As String
    Select Case VarType(value)
        Case vbString: KeyLiteral = "'" & Replace(value, "'", "''") & "'"
        Case vbDate: KeyLiteral = "#" & Format$(value, "yyyy-mm-dd hh:nn:ss") & "#"
        Case Else: KeyLiteral = Trim$(Str$(value))
    End Select
End Function

' Footer controls are prefixCount/Previous/Page0/Page1/Page2/Next. This call is
' harmless on unmarked tables: it hides every footer, restores height and scope.
' The record source must have a stable ORDER BY and a non-null unique key.
Public Sub LayoutTableFooter(ByVal host As Form, ByVal subName As String, ByVal prefix As String, ByVal keyField As String)
    On Error GoTo Failed
    Dim box As Control, grid As Form, c As Control, state As Object, key As String
    Dim originalScroll As Long, fullHeight As Long, size As Long, count As Long, pages As Long, skip As Long
    Dim rows As DAO.Recordset, ids As String, shown As Long, i As Long, number As Long, first As Long
    If m_busy Then Exit Sub
    Set box = host.Controls(subName): key = TableKey(host, subName)
    For Each c In host.Controls
        If c.Tag = "TableFooter:" & prefix Then c.Visible = box.Visible And TablePaginationEnabled(box)
    Next
    If Not TablePaginationEnabled(box) Then
        If Not m_geometry Is Nothing Then
            If m_geometry.Exists(key) Then
                If box.Height = m_geometry(key)(1) Then box.Height = m_geometry(key)(0)
                If Len(box.SourceObject) > 0 Then box.Form.ScrollBars = m_geometry(key)(2)
                m_geometry.Remove key
            End If
        End If
        If Not m_pages Is Nothing Then
            If m_pages.Exists(key) Then
                Set state = m_pages(key)
                If Len(box.SourceObject) > 0 Then
                    Set grid = box.Form
                    If grid.Filter = state("Last") Then grid.Filter = state("Base"): grid.FilterOn = Len(state("Base")) > 0
                End If
                m_pages.Remove key
            End If
        End If
        Exit Sub
    End If
    If Not box.Visible Or Len(box.SourceObject) = 0 Then Exit Sub
    Set grid = box.Form
    If grid.DefaultView <> 1 Then Exit Sub
    m_busy = True
    If m_geometry Is Nothing Then Set m_geometry = CreateObject("Scripting.Dictionary")
    If m_pages Is Nothing Then Set m_pages = CreateObject("Scripting.Dictionary")
    fullHeight = box.Height: originalScroll = grid.ScrollBars
    If m_geometry.Exists(key) Then
        originalScroll = m_geometry(key)(2)
        If fullHeight = m_geometry(key)(1) Then fullHeight = m_geometry(key)(0)
    End If
    If fullHeight < grid.Section(acDetail).Height + px(40) Then fullHeight = grid.Section(acDetail).Height + px(40)
    box.Height = fullHeight - px(40): m_geometry(key) = Array(fullHeight, box.Height, originalScroll)
    grid.ScrollBars = IIf(grid.ScrollBars = 1 Or grid.ScrollBars = 3, 1, 0)
    size = box.Height \ grid.Section(acDetail).Height: If size < 1 Then size = 1
    If Not m_pages.Exists(key) Then
        Set state = CreateObject("Scripting.Dictionary")
        state.Add "Page", 1: state.Add "Base", IIf(grid.FilterOn, grid.Filter, "")
        state.Add "Last", grid.Filter: state.Add "Source", grid.RecordSource
        Set state("Grid") = grid: m_pages.Add key, state
    Else
        Set state = m_pages(key)
        If Not (state("Grid") Is grid) Or state("Source") <> grid.RecordSource Or state("Last") <> grid.Filter Then
            state("Page") = 1: state("Base") = IIf(grid.FilterOn, grid.Filter, "")
            state("Source") = grid.RecordSource: Set state("Grid") = grid
        End If
    End If
    state("Prefix") = prefix: state("KeyField") = keyField   ' for RestartTablePages
    grid.Filter = state("Base"): grid.FilterOn = Len(state("Base")) > 0
    Set rows = grid.Recordset.Clone
    If Not rows.EOF Then rows.MoveLast: count = rows.RecordCount: rows.MoveFirst
    pages = (count + size - 1) \ size: If pages < 1 Then pages = 1
    If state("Page") > pages Then state("Page") = pages
    If state("Page") < 1 Then state("Page") = 1
    skip = (state("Page") - 1) * size
    If count > size Then
        For i = 1 To skip: rows.MoveNext: Next
        Do While Not rows.EOF And shown < size
            If IsNull(rows.Fields(keyField).Value) Then Err.Raise 5, , "Pagination needs a non-null unique key"
            ids = ids & IIf(Len(ids) > 0, ",", "") & KeyLiteral(rows.Fields(keyField).Value)
            shown = shown + 1: rows.MoveNext
        Loop
        grid.Filter = IIf(Len(state("Base")) > 0, "(" & state("Base") & ") AND ", "") & "[" & keyField & "] IN (" & ids & ")"
        grid.FilterOn = True
    Else
        shown = count
    End If
    rows.Close: Set rows = Nothing
    state("Last") = grid.Filter: state("Pages") = pages
    With host.Controls(prefix & "Count")
        .Left = box.Left: .Top = box.Top + box.Height + px(4): .Width = box.Width - px(192)
        .Caption = IIf(count = 0, "No records", "Showing " & skip + 1 & "-" & skip + shown & " of " & count)
    End With
    first = state("Page") - 1
    If first > pages - 2 Then first = pages - 2
    If first < 1 Then first = 1
    For i = 0 To 4
        Set c = host.Controls(prefix & Choose(i + 1, "Previous", "Page0", "Page1", "Page2", "Next"))
        c.Left = box.Left + box.Width - px(186) + i * px(36)
        c.Top = box.Top + box.Height + px(4): c.Width = px(30): c.Height = px(26)
        If i = 0 Then c.Enabled = state("Page") > 1
        If i = 4 Then c.Enabled = state("Page") < pages
        If i > 0 And i < 4 Then
            number = first + i - 1: c.Caption = CStr(number): c.Visible = number <= pages
            c.BackColor = Th(IIf(number = state("Page"), CLR_PRIMARY, CLR_OUT_BG))
            c.ForeColor = Th(IIf(number = state("Page"), CLR_WHITE, CLR_OUT_FG))
        End If
    Next
    m_busy = False: Exit Sub
Failed:
    Dim num As Long, description As String, source As String
    num = Err.Number: description = Err.Description: source = Err.source
    On Error Resume Next
    If Not rows Is Nothing Then rows.Close
    m_busy = False
    On Error GoTo 0
    Err.Raise num, source, description
End Sub

' A paginated table's sort changed (GridSort): go back to page 1 and pick
' its rows again in the new order, since the page shown is the old order's
' rows. Does nothing for a table without pagination or not yet laid out.
Public Sub RestartTablePages(ByVal host As Form, ByVal subName As String)
    Dim key As String, state As Object
    If Not TablePaginationEnabled(host.Controls(subName)) Then Exit Sub
    If m_pages Is Nothing Then Exit Sub
    key = TableKey(host, subName): If Not m_pages.Exists(key) Then Exit Sub
    Set state = m_pages(key)
    state("Page") = 1
    LayoutTableFooter host, subName, state("Prefix"), state("KeyField")
End Sub

Public Function MoveTablePage(ByVal host As Form, ByVal subName As String, ByVal prefix As String, ByVal keyField As String, ByVal delta As Long, Optional ByVal slot As Long = -1) As Boolean
    Dim key As String, state As Object, page As Long
    If Not TablePaginationEnabled(host.Controls(subName)) Then Exit Function
    If m_pages Is Nothing Then Exit Function
    key = TableKey(host, subName): If Not m_pages.Exists(key) Then Exit Function
    Set state = m_pages(key)
    If slot >= 0 Then page = CLng(host.Controls(prefix & "Page" & slot).Caption) Else page = state("Page") + delta
    If page < 1 Or page > state("Pages") Then Exit Function
    state("Page") = page: LayoutTableFooter host, subName, prefix, keyField
    MoveTablePage = True
End Function

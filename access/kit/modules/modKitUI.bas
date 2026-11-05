Option Compare Database
Option Explicit

' Access UI kit - run-time helpers for pages built with modUIKit: tables that
' scale their columns (GridFit), list/detail pages (SelectFirstRow,
' ShowDetail, ClearDetail), search matches marked in tables (GridHighlight),
' tabs and filter chips with counts (SetCount, StyleChip), focus rings on
' inputs (FocusRing), toggle switches (ToggleSwitch, ToggleFocus), segmented
' controls (SegmentClick, SetSegment, SegmentValue), the splash
' screen's status line (SplashStatus) and messages and questions in the kit's
' look (KitMsgBox, KitInputBox, and KitToast for ones that fade on their own).
' Colors set here go through Th() so they follow the palette.

' Width the table's vertical scroll bar takes (twips).
Private Const SCROLL_W As Long = 300
' Widest a form can be: 22 inches (twips).
Private Const MAX_FORM_W As Long = 31680
' Arrows a sorted GridHeaders caption shows (GridSort): up ascending, down
' descending (Unicode code points, for ChrW).
Private Const SORT_UP As Long = &H25B2
Private Const SORT_DOWN As Long = &H25BC
' Design width of each GridFit host, by its window and name, and the design
' Left and Width of its anchored controls, read on its first fit before any
' caption has widened it (HostDesignW).
Private m_hostW As Collection
' True while GridFit runs: narrowing the host can fire its Form_Resize again.
Private m_fitting As Boolean
' The DateBox the date picker (dlgDatePick) is open for (PickDate).
Private m_dateTarget As Control
' Dates the app marks on the date picker (MarkDate): label by "d" & date serial.
Private m_marks As Collection
' The message the message dialog (dlgMessage) is showing (KitMsgBox), and the
' button chosen in it.
Private m_msgOpen As Boolean
Private m_msgText As String
Private m_msgStyle As Long
Private m_msgTitle As String
Private m_msgAnswer As Long
' The splash screen BuildSplash builds (SplashStatus, CloseSplash).
Private Const SPLASH_FORM As String = "frmSplash"
' Ends frmSplash's Tag when it opens its next form hidden (BuildSplash
' openHidden).
Public Const SPLASH_HIDDEN As String = "|hidden"
' The field showing the focus ring (FocusRing): the control, its form's
' window and its name, and the border color the ring covers.
Private m_ringCtl As Control
Private m_ringHwnd As Long
Private m_ringName As String
Private m_ringWas As Long
' The question the input dialog (dlgInput) is asking (KitInputBox), and the
' text typed (Null until OK is clicked).
Private m_inOpen As Boolean
Private m_inPrompt As String
Private m_inTitle As String
Private m_inDefault As String
Private m_inMasked As Boolean
Private m_inAnswer As Variant
' KitBusy: the messages waiting for their KitBusyEnd, newest last; whether
' this busy work dimmed the app for dlgBusy; whether the status bar shows
' the message (no dlgBusy, or it failed to open).
Private m_busy As Collection
Private m_busyPopup As Boolean
Private m_busyStatus As Boolean
' A BeforeUpdate refused the save and told the user why (KitRefuseSave).
' Access then raises an error that depends on where the focus is (2465 when
' it is still in the subform), so the save routines go by this, not the number.
Private m_refused As Boolean

' ---------------------------------------------------------------------------
' Tables
' ---------------------------------------------------------------------------

' Proportional fit for a table built with BuildGrid(..., fit:=True): scale
' every column of the table shown in subform control sf to the control's
' current width, and move the captions GridHeaders drew on the parent with
' prefix. Columns never get narrower than designed. Call it from the host's
' Form_Resize.
Public Sub GridFit(sf As Control, Optional ByVal prefix As String = "")
    Dim host As Form, g As Form, c As Control, i As Long, t() As String
    Dim design As Long, avail As Long, ratio As Double, newW As Long
    Set host = OwnerForm(sf)
    Set g = sf.Form
    If Left$(Nz(g.Tag, ""), 5) <> "grid:" Then Exit Sub
    If m_fitting Then Exit Sub
    m_fitting = True
    On Error GoTo Fail
    ' Narrowing and then widening the host again would flash.
    host.Painting = False
    design = Val(Mid$(g.Tag, 6))
    ' A caption moved past the host's right edge widens the host, and Access
    ' never narrows it again: the page and its anchored controls would stay
    ' wide after the window is narrowed. So start each fit from the host's
    ' design width (ShrinkHost).
    ShrinkHost host, prefix
    ' Resizing the host can leave the reference to the table's form dead.
    Set g = sf.Form
    ' An anchored control's Width keeps its design value, so add how much
    ' wider than designed the host is.
    avail = sf.Width
    If sf.HorizontalAnchor = 2 Then avail = avail + host.InsideWidth - host.Width
    ratio = (avail - SCROLL_W) / (design - SCROLL_W)
    If ratio < 1 Then ratio = 1
    If design * ratio > MAX_FORM_W Then ratio = MAX_FORM_W / design
    newW = CLng(design * ratio)
    ' Access ignores g.Width = n on a form shown in a subform: the form only
    ' grows when a control's right edge pushes it, and moving a column past
    ' its edge raises error 5. So widen the row background first (it pushes
    ' the form out), place the columns, then shrink the background.
    ' Resizing the form can leave an earlier reference to it dead (error 5 on
    ' any use), so it is read from sf again before each use.
    If newW > g.Width Then
        If HasCtl(g, "txtBg") Then g.Controls("txtBg").Width = newW
        Set g = sf.Form
        If HasCtl(g, "boxRowLine") Then g.Controls("boxRowLine").Width = newW
    End If
    Do
        Set g = sf.Form
        Set c = Nothing
        On Error Resume Next
        Set c = g.Controls("txtCol" & i)
        On Error GoTo Fail
        If c Is Nothing Then Exit Do
        t = Split(c.Tag, "|")
        If UBound(t) >= 2 Then
            FitCtl c, CLng(Val(t(1)) * ratio), CLng(Val(t(2)) * ratio) - px(8)
            Set g = sf.Form
            Set c = g.Controls("txtCol" & i)
            If HasCtl(g, c.Name & "_bg") Then FitCtl g.Controls(c.Name & "_bg"), c.Left, c.Width
        End If
        PlaceCaption host, prefix, i, ratio
        i = i + 1
    Loop
    Set g = sf.Form
    If HasCtl(g, "txtBg") Then FitCtl g.Controls("txtBg"), 0, newW
    Set g = sf.Form
    If HasCtl(g, "boxRowLine") Then FitCtl g.Controls("boxRowLine"), 0, newW
    Set g = sf.Form
    ' Narrowing the form is cosmetic (no empty strip to scroll to), and
    ' Access may ignore it.
    If newW < g.Width Then
        On Error Resume Next
        g.Width = newW
        On Error GoTo Fail
    End If
    host.Painting = True
    m_fitting = False
    Exit Sub
Fail:
    Dim errNo As Long, errSrc As String, errDesc As String
    errNo = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    m_fitting = False
    On Error Resume Next
    host.Painting = True
    On Error GoTo 0
    Err.Raise errNo, errSrc, errDesc
End Sub

' Put GridFit's host back at its design width when a caption has widened it.
' When the host widens, Access widens its anchored controls' design sizes
' with it (Width of a control anchored left and right, Left of one anchored
' right), by more than the host over several widenings. So they go back to
' the sizes HostDesignW recorded, and the captions to their design places,
' before the host's Width; otherwise Access refuses the narrower Width.
' Setting a narrower Width can land short of it (1027px for 1056 on the demo
' Tables page), and Access ignores a wider Width on a form in a subform. A
' control's right edge does push it out, so the first caption is stretched to
' the design width and back, then the anchored controls, which grew with it,
' are put back again.
Private Sub ShrinkHost(host As Form, ByVal prefix As String)
    Dim hostW As Long, c As Control, i As Long, w As Long
    hostW = HostDesignW(host, prefix)
    If host.Width = hostW Then
        ' The detail section can be left short after the host's width came
        ' back (RestoreAnchored).
        If host.Section(acDetail).Height <> CLng(m_hostW(HostKey(host) & "|*detail")) Then RestoreAnchored host
        Exit Sub
    End If
    If host.Width > hostW Then
        If Len(prefix) > 0 Then
            Do While HasCtl(host, prefix & i)
                PlaceCaption host, prefix, i, 1
                i = i + 1
            Loop
        End If
        RestoreAnchored host
        host.Width = hostW
    End If
    If host.Width < hostW And Len(prefix) > 0 Then
        If HasCtl(host, prefix & "0") Then
            Set c = host.Controls(prefix & "0")
            w = c.Width
            c.Width = hostW - c.Left
            c.Width = w
            RestoreAnchored host
        End If
    End If
End Sub

' Place caption prefix & i (GridHeaders, Tag "x|offset|width|...") over its
' column scaled by ratio, with its filter funnel if it has one.
Private Sub PlaceCaption(host As Form, ByVal prefix As String, ByVal i As Long, ByVal ratio As Double)
    Dim t() As String, capX As Long, capW As Long
    If Len(prefix) = 0 Then Exit Sub
    If Not HasCtl(host, prefix & i) Then Exit Sub
    t = Split(host.Controls(prefix & i).Tag, "|")
    If UBound(t) < 2 Then Exit Sub
    capX = CLng(Val(t(0)) + Val(t(1)) * ratio)
    capW = CLng((Val(t(2)) + px(8)) * ratio) - px(8)
    ' A filter funnel (GridHeaders filterSf) keeps the right end of the
    ' caption's space; the text stops short of it.
    If HasCtl(host, prefix & "Filter" & i) Then
        FitCtl host.Controls(prefix & "Filter" & i), MaxL(capX + capW - FILTER_W, 0), FILTER_W
        capW = capW - FILTER_W - FILTER_GAP
    End If
    FitCtl host.Controls(prefix & i), capX, capW
End Sub

' Put each anchored control of GridFit's host back at the design Left,
' Width, Top and Height HostDesignW recorded, and the detail section at its
' design height. The width change shifts the vertical anchoring too: after a
' widen and narrow the Tables page's detail section read 8160 twips for 8400,
' and its tables reached 16px lower. A taller section goes first (a control
' can't reach past it), a shorter one after the controls.
Private Sub RestoreAnchored(host As Form)
    Dim c As Control, k As String, v() As String, secH As Long
    k = HostKey(host)
    secH = CLng(m_hostW(k & "|*detail"))
    If host.Section(acDetail).Height < secH Then host.Section(acDetail).Height = secH
    For Each c In host.Controls
        If Anchored(c) Then
            v = Split(m_hostW(k & "|" & c.Name), "|")
            If c.Width > CLng(v(1)) Then c.Width = CLng(v(1))
            If c.Left <> CLng(v(0)) Then c.Left = CLng(v(0))
            If c.Width <> CLng(v(1)) Then c.Width = CLng(v(1))
            If c.Height > CLng(v(3)) Then c.Height = CLng(v(3))
            If c.Top <> CLng(v(2)) Then c.Top = CLng(v(2))
            If c.Height <> CLng(v(3)) Then c.Height = CLng(v(3))
        End If
    Next
    If host.Section(acDetail).Height > secH Then host.Section(acDetail).Height = secH
End Sub

' True when c is anchored right, bottom or both ways. A tab control's pages
' have no anchoring of their own.
Private Function Anchored(c As Control) As Boolean
    If c.ControlType = acPage Then Exit Function
    Anchored = c.HorizontalAnchor <> 0 Or c.VerticalAnchor <> 0
End Function

' Key of a GridFit host in m_hostW.
Private Function HostKey(host As Form) As String
    HostKey = host.Hwnd & "|" & host.Name
End Function

' Design width of GridFit's host: the form width GridHeaders saved in the
' first caption's Tag (part "fw=<twips>"). Captions drawn before that existed
' fall back to the host's Width on its first fit, which can be short of the
' design width while the page is still opening. Kept per window (and name,
' in case a window handle is reused) so a reopened page starts again. The
' first fit also records each anchored control's design "Left|Width|Top|
' Height" (RestoreAnchored), under the host's key, "|" and the control's
' name, and the detail section's height under the key and "|*detail".
Private Function HostDesignW(host As Form, ByVal prefix As String) As Long
    Dim k As String, c As Control, found As Boolean, p As Variant
    If m_hostW Is Nothing Then Set m_hostW = New Collection
    k = HostKey(host)
    On Error Resume Next
    HostDesignW = m_hostW(k)
    found = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If found Then Exit Function
    HostDesignW = host.Width
    If Len(prefix) > 0 And HasCtl(host, prefix & "0") Then
        For Each p In Split(host.Controls(prefix & "0").Tag, "|")
            If Left$(p, 3) = "fw=" And Val(Mid$(p, 4)) > 0 Then HostDesignW = CLng(Val(Mid$(p, 4)))
        Next
    End If
    m_hostW.Add HostDesignW, k
    m_hostW.Add host.Section(acDetail).Height, k & "|*detail"
    For Each c In host.Controls
        If Anchored(c) Then m_hostW.Add c.Left & "|" & c.Width & "|" & c.Top & "|" & c.Height, k & "|" & c.Name
    Next
End Function

' The larger of a and b.
Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    MaxL = IIf(a > b, a, b)
End Function

' Move and size a control horizontally.
Private Sub FitCtl(c As Control, ByVal x As Long, ByVal w As Long)
    If w < 15 Then w = 15
    If c.Left <> x Then c.Left = x
    If c.Width <> w Then c.Width = w
End Sub

' True when the form has a control named nm.
Private Function HasCtl(f As Form, ByVal nm As String) As Boolean
    Dim c As Control
    If Len(nm) = 0 Then Exit Function
    On Error Resume Next
    Set c = f.Controls(nm)
    HasCtl = Not c Is Nothing
End Function

' Show a table's vertical scroll bar only when its rows don't all fit:
' Access draws the bar on a continuous form whatever its content. A BuildGrid
' table calls this from its OnCurrent, which fires whenever its record source
' is set or requeried (and once before the host sets it: no record source,
' no rows). Rows are counted only until one more than fits, so a large linked
' table isn't read to the end on every row change. A failure is logged, not
' shown: it only leaves the bar.
Public Function GridScroll(g As Form)
    Dim n As Long, sb As Integer, rowH As Long, fits As Long
    On Error GoTo Fail
    rowH = g.Section(acDetail).Height
    If Len(g.RecordSource) > 0 And rowH > 0 Then
        fits = g.InsideHeight \ rowH
        With g.RecordsetClone
            If Not (.BOF And .EOF) Then .MoveFirst   ' a new clone's position isn't defined
            Do While Not .EOF And n <= fits
                n = n + 1
                .MoveNext
            Loop
        End With
    End If
    ' Only the vertical bar (2) changes; a horizontal one (1, set by ResizeGrid) stays.
    sb = (g.ScrollBars And 1) Or IIf(n * rowH > g.InsideHeight, 2, 0)
    If g.ScrollBars <> sb Then g.ScrollBars = sb
    Exit Function
Fail:
    If Err.Number = 7951 Then Exit Function     ' no recordset yet: nothing to scroll
    LogError "modKitUI.GridScroll", g.Name
End Function

' GridScroll for every BuildGrid table on frm and on the forms inside it.
' Call it from the window's Form_Resize: a table anchored to the window
' changes height without its OnCurrent firing. A subform still loading (the
' window resizes as it opens) is skipped; GridScroll logs its own failures.
Public Sub GridScrollAll(frm As Form)
    Dim c As Control
    On Error Resume Next
    For Each c In frm.Controls
        If c.ControlType = acSubform Then
            If Len(c.SourceObject) > 0 Then
                If c.Form.Tag Like "grid:*" Then GridScroll c.Form Else GridScrollAll c.Form
            End If
        End If
    Next
End Sub

' Caption click on a sortable GridHeaders table (its sortSf): sort the table
' in subform control sfName on host by caption prefix & col's sort field,
' ascending, or descending when that caption already shows the ascending
' arrow. The sorted caption shows the SORT_UP or SORT_DOWN arrow (on its
' left when the caption is right-aligned, so the text stays over the
' numbers) and the others lose theirs. The arrow is the sort state: a host
' that sorts the table in code should do it through GridSort.
' The sort is the form's OrderBy, which PrintGrid and ExportGrid follow. A
' paginated table (GridFooter) goes back to its first page in the new order.
' Ties keep id order. The selected row stays selected (txtSel keeps its id).
' A failure is shown (HandleError), as for any click; code that sorts a
' table and must know it failed should check the table's OrderBy after.
Public Function GridSort(host As Form, ByVal sfName As String, ByVal prefix As String, ByVal col As Long)
    Dim g As Form, cap As Control, t() As String, key As String, desc As Boolean, i As Long
    On Error GoTo Fail
    Set cap = host.Controls(prefix & col)
    t = Split(cap.Tag & "|||", "|")
    key = Trim$(t(3))
    If Len(key) = 0 Then Exit Function
    If Left$(key, 1) <> "[" Then key = "[" & key & "]"
    desc = InStr(cap.Caption, ChrW(SORT_UP)) > 0
    Set g = host.Controls(sfName).Form
    g.OrderBy = key & IIf(desc, " DESC", "") & IIf(key = "[id]", "", ", [id]")
    g.OrderByOn = True
    RestartTablePages host, sfName      ' a paginated table shows the old order's page
    Do While HasCtl(host, prefix & i)
        SortArrow host.Controls(prefix & i), IIf(i = col, IIf(desc, ChrW(SORT_DOWN), ChrW(SORT_UP)), "")
        i = i + 1
    Loop
    Exit Function
Fail:
    HandleError "modKitUI.GridSort", prefix & col
End Function

' Put arrow (or none, "") on a GridHeaders caption in place of the one it has.
Private Sub SortArrow(cap As Control, ByVal arrow As String)
    Dim s As String
    s = Trim$(Replace(Replace(cap.Caption, ChrW(SORT_UP), ""), ChrW(SORT_DOWN), ""))
    If Len(arrow) > 0 Then
        If cap.TextAlign = 3 Then s = arrow & " " & s Else s = s & " " & arrow
    End If
    If cap.Caption <> s Then cap.Caption = s
End Sub

' Select the first row of a BuildGrid table and tell the host, as a click
' would (its GridClick grid, False). Returns the row's id, or Null when the
' table is empty (then nothing is called). Call it after the list loads or
' its filter changes.
Public Function SelectFirstRow(sf As Control) As Variant
    Dim g As Form
    SelectFirstRow = Null
    Set g = sf.Form
    If g.RecordsetClone.RecordCount = 0 Then
        g.txtSel = Null
        Exit Function
    End If
    g.Recordset.MoveFirst
    g.txtSel = g!id
    g.Recalc
    SelectFirstRow = g.txtSel
    OwnerForm(sf).GridClick sf.Name, False
End Function

' Mark term in the search columns (BuildGrid option "s") of the table shown
' in subform control sf. Call it with the page's search box text whenever
' that changes; "" clears the marks. Text matches literally, ignoring case.
' A table built before the "s" option (no txtFind) is left as it is.
Public Sub GridHighlight(sf As Control, ByVal term As String)
    Dim g As Form
    Set g = sf.Form
    If Not HasCtl(g, "txtFind") Then Exit Sub
    If StrComp(Nz(g.txtFind, ""), term, vbBinaryCompare) = 0 Then Exit Sub
    g.txtFind = IIf(Len(term) > 0, term, Null)
    g.Recalc
End Sub

' A search column's cell: value as rich text, with each match of term shown
' in the highlight colors (CLR_WARNING_FG on CLR_WARNING_BG, which keep their
' contrast on the selected row and in every palette). Access calls it for
' every row it paints, so it never raises: on a failure it shows the value
' unmarked.
Public Function GridMark(ByVal v As Variant, ByVal term As Variant) As Variant
    Dim s As String, t As String, i As Long, j As Long, out As String, mark As String
    On Error GoTo Fail
    If IsNull(v) Then GridMark = Null: Exit Function
    s = CStr(v)
    t = Nz(term, "")
    i = 1
    If Len(t) > 0 Then
        mark = "<font color=""" & HtmlColor(Th(CLR_WARNING_FG)) & """ style=""BACKGROUND-COLOR:" & _
            HtmlColor(Th(CLR_WARNING_BG)) & """>"
        Do
            j = InStr(i, s, t, vbTextCompare)
            If j = 0 Then Exit Do
            out = out & HtmlText(Mid$(s, i, j - i), i = 1) & mark & HtmlText(Mid$(s, j, Len(t)), j = 1) & "</font>"
            i = j + Len(t)
        Loop
        s = Mid$(s, i)
    End If
    GridMark = out & HtmlText(s, i = 1)
    Exit Function
Fail:
    Resume Unmarked                         ' leave the handler, so a second error is trapped too
Unmarked:
    On Error Resume Next
    GridMark = HtmlText(CStr(v), True)
End Function

' Text for a rich text box: the characters HTML reserves escaped, line breaks
' as <br>, and runs of spaces and tabs kept (HTML would collapse them).
' cellStart = True when s begins the cell, where a single space is kept too.
Private Function HtmlText(ByVal s As String, Optional ByVal cellStart As Boolean = False) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, vbTab, "    ")
    s = Replace(s, "  ", " &nbsp;")
    If cellStart And Left$(s, 1) = " " Then s = "&nbsp;" & Mid$(s, 2)
    s = Replace(s, vbCrLf, "<br>")
    HtmlText = Replace(s, vbLf, "<br>")
End Function

' A color as HTML "#rrggbb".
Private Function HtmlColor(ByVal c As Long) As String
    HtmlColor = "#" & Right$("0" & Hex$(c And &HFF&), 2) & Right$("0" & Hex$((c \ &H100&) And &HFF&), 2) & _
        Right$("0" & Hex$((c \ &H10000) And &HFF&), 2)
End Function

' ---------------------------------------------------------------------------
' List and detail pages
' ---------------------------------------------------------------------------

' Show the records matching where in the detail panel: subform control
' detailCtl on host, showing form formName. A changed record in the panel is
' saved first, and the empty-state control emptyCtl is hidden. The panel's
' form is only swapped when formName changes. Swapping a subform while one of
' its own events is running is unsafe, so call this from the list (GridClick)
' or the page, never from inside the detail form. Returns False, and changes
' nothing, when the panel's form refuses to save its changed record (its
' BeforeUpdate cancels, for example CheckRequired found an empty field): the
' panel keeps the record so the user can finish it.
Public Function ShowDetail(host As Form, ByVal detailCtl As String, ByVal formName As String, ByVal where As String, _
                      Optional ByVal emptyCtl As String = "lblDetailEmpty") As Boolean
    Dim sf As Control
    On Error GoTo Fail
    Set sf = host.Controls(detailCtl)
    If Not SaveDetailCtl(sf) Then Exit Function
    If sf.SourceObject <> formName And sf.SourceObject <> "Form." & formName Then sf.SourceObject = formName
    sf.Form.Filter = where
    sf.Form.FilterOn = True
    SetVisible sf, True
    If HasCtl(host, emptyCtl) Then SetVisible host.Controls(emptyCtl), False
    ShowDetail = True
    Exit Function
Fail:
    RaiseAgain "modKitUI.ShowDetail", formName & " where " & where
End Function

' Empty the detail panel (saving a changed record first) and show the
' empty-state control instead. Returns False, and changes nothing, when the
' panel's form refuses to save (as ShowDetail).
Public Function ClearDetail(host As Form, ByVal detailCtl As String, Optional ByVal emptyCtl As String = "lblDetailEmpty") As Boolean
    Dim sf As Control
    On Error GoTo Fail
    Set sf = host.Controls(detailCtl)
    If Not SaveDetailCtl(sf) Then Exit Function
    SetVisible sf, False
    sf.SourceObject = ""
    If HasCtl(host, emptyCtl) Then SetVisible host.Controls(emptyCtl), True
    ClearDetail = True
    Exit Function
Fail:
    RaiseAgain "modKitUI.ClearDetail", detailCtl
End Function

' Save the changed record in the detail panel (subform control detailCtl on
' host). Returns False when the panel's form refuses (as ShowDetail). Call
' it before a page change that doesn't go through ShowDetail, such as
' applying a filter chip, and stop when it returns False.
Public Function SaveDetail(host As Form, ByVal detailCtl As String) As Boolean
    On Error GoTo Fail
    SaveDetail = SaveDetailCtl(host.Controls(detailCtl))
    Exit Function
Fail:
    RaiseAgain "modKitUI.SaveDetail", detailCtl
End Function

' Save the record shown in a detail panel if it changed. False when the
' form's BeforeUpdate cancelled the save. Which error Access raises for a
' cancelled Dirty = False depends on what the BeforeUpdate did: 2101, 2115
' (a control's own BeforeUpdate) or 2455 (it opened a message, as
' CheckRequired does). Any other error, such as the database engine's own
' validation, is raised again so the user still sees it.
Private Function SaveDetailCtl(sf As Control) As Boolean
    Dim errNo As Long, errDesc As String, errSrc As String
    SaveDetailCtl = True
    If Len(sf.SourceObject) = 0 Then Exit Function
    If Not sf.Form.Dirty Then Exit Function
    KitTakeRefused
    On Error Resume Next
    sf.Form.Dirty = False
    errNo = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    On Error GoTo 0
    If errNo = 0 Then Exit Function
    If KitTakeRefused() Or errNo = 2101 Or errNo = 2115 Or errNo = 2455 Then
        SaveDetailCtl = False
    Else
        Err.Raise errNo, errSrc, errDesc
    End If
End Function

' ---------------------------------------------------------------------------
' Date fields
' ---------------------------------------------------------------------------

' Click on a DateBox's calendar button: open the date picker for the box nm
' and wait for it to close. Does nothing while the box is disabled or locked,
' or its form is read-only (AllowEdits off, unless on a record being added).
Public Function PickDate(frm As Form, ByVal nm As String)
    Dim c As Control
    On Error GoTo Fail
    Set c = frm.Controls(nm)
    If Not c.Enabled Or c.Locked Then Exit Function
    If Not frm.AllowEdits And Not (frm.NewRecord And frm.AllowAdditions) Then Exit Function
    c.SetFocus
    Set m_dateTarget = c
    DoCmd.OpenForm "dlgDatePick", WindowMode:=acDialog
    Set m_dateTarget = Nothing
    Exit Function
Fail:
    Set m_dateTarget = Nothing
    HandleError "modKitUI.PickDate", nm
End Function

' The open date picker's box: its value (Null when none is open), and
' setting it. Setting a value does not fire the box's own update events; the
' record's BeforeUpdate still checks it when it is saved.
Public Function DateTargetValue() As Variant
    DateTargetValue = Null
    If Not m_dateTarget Is Nothing Then DateTargetValue = m_dateTarget.Value
End Function

Public Sub SetDateTarget(ByVal v As Variant)
    If m_dateTarget Is Nothing Then Exit Sub
    m_dateTarget.Value = v
End Sub

' Mark day d on the date picker with label lbl (a deadline, a drill, a pay
' day). A second label on the same day is added after "; ". Marks last until
' ClearMarkedDates or the app closes, so set them before the picker opens, for
' example at startup from the app's own deadlines table.
Public Sub MarkDate(ByVal d As Date, ByVal lbl As String)
    Dim k As String, cur As String
    If m_marks Is Nothing Then Set m_marks = New Collection
    k = "d" & CLng(Int(d))
    cur = MarkedDate(d)
    If Len(cur) > 0 Then
        m_marks.Remove k
        lbl = cur & "; " & lbl
    End If
    m_marks.Add lbl, k
End Sub

' Remove every MarkDate mark.
Public Sub ClearMarkedDates()
    Set m_marks = Nothing
End Sub

' The app's label for day d (MarkDate), or "" when it has none.
Public Function MarkedDate(ByVal d As Date) As String
    If m_marks Is Nothing Then Exit Function
    On Error Resume Next
    MarkedDate = m_marks("d" & CLng(Int(d)))
    Err.Clear
End Function

' The US federal holiday on day d (5 U.S.C. 6103), or "" when there is none.
' Fixed-date holidays are given on the day they are observed: a Saturday
' holiday on the Friday before, a Sunday one on the Monday after, with
' " (observed)" added. So New Year's Day can fall on 31 December. These are
' the rules since 1978 (Veterans Day back on 11 November); earlier dates
' return "".
Public Function FederalHoliday(ByVal d As Date) As String
    Dim y As Integer, nm As String
    d = Int(d)
    y = Year(d)
    If y < 1978 Then Exit Function
    nm = ObservedHoliday(d, DateSerial(y + 1, 1, 1), "New Year's Day")
    If Len(nm) = 0 Then nm = ObservedHoliday(d, DateSerial(y, 1, 1), "New Year's Day")
    If Len(nm) = 0 And y >= 2021 Then nm = ObservedHoliday(d, DateSerial(y, 6, 19), "Juneteenth National Independence Day")
    If Len(nm) = 0 Then nm = ObservedHoliday(d, DateSerial(y, 7, 4), "Independence Day")
    If Len(nm) = 0 Then nm = ObservedHoliday(d, DateSerial(y, 11, 11), "Veterans Day")
    If Len(nm) = 0 Then nm = ObservedHoliday(d, DateSerial(y, 12, 25), "Christmas Day")
    If Len(nm) > 0 Then FederalHoliday = nm: Exit Function
    ' The Monday and Thursday holidays never move.
    If d = NthWeekday(y, 1, vbMonday, 3) And y >= 1986 Then
        nm = "Birthday of Martin Luther King, Jr."
    ElseIf d = NthWeekday(y, 2, vbMonday, 3) Then
        nm = "Washington's Birthday"
    ElseIf d = NthWeekday(y, 6, vbMonday, 1) - 7 Then
        nm = "Memorial Day"                     ' the last Monday in May
    ElseIf d = NthWeekday(y, 9, vbMonday, 1) Then
        nm = "Labor Day"
    ElseIf d = NthWeekday(y, 10, vbMonday, 2) Then
        nm = "Columbus Day"
    ElseIf d = NthWeekday(y, 11, vbThursday, 4) Then
        nm = "Thanksgiving Day"
    End If
    FederalHoliday = nm
End Function

' nm when d is the day holiday hol is observed on, else "".
Private Function ObservedHoliday(ByVal d As Date, ByVal hol As Date, ByVal nm As String) As String
    Dim obs As Date
    obs = hol
    Select Case Weekday(hol, vbSunday)
        Case vbSaturday: obs = hol - 1
        Case vbSunday: obs = hol + 1
    End Select
    If d = obs Then ObservedHoliday = nm & IIf(obs <> hol, " (observed)", "")
End Function

' The n-th weekday wd (vbMonday ...) of month m in year y.
Private Function NthWeekday(ByVal y As Integer, ByVal m As Integer, ByVal wd As VbDayOfWeek, ByVal n As Integer) As Date
    Dim first As Date
    first = DateSerial(y, m, 1)
    NthWeekday = first + ((wd - Weekday(first, vbSunday) + 7) Mod 7) + (n - 1) * 7
End Function

' ---------------------------------------------------------------------------
' Sensitive fields
' ---------------------------------------------------------------------------

' Click on a SecretBox's eye button: show the box nm's text if it is hidden,
' hide it if it is shown, and put the focus back in the box with the caret
' at the end, so typing carries on rather than replacing the value (not
' while the box is disabled or hidden).
Public Function ToggleSecret(frm As Form, ByVal nm As String)
    Dim c As Control
    On Error GoTo Fail
    Set c = frm.Controls(nm)
    ShowSecret frm, nm, (c.InputMask = SECRET_MASK)
    If c.Enabled And c.Visible Then
        c.SetFocus
        ' The caret is a nicety, so it must never cost an error message. Access can
        ' refuse the box's Text (error 2196) just after SetFocus, before the focus
        ' has really arrived; the saved Value is the same text then.
        On Error Resume Next
        c.SelStart = Len(Nz(c.Text, ""))
        If Err.Number <> 0 Then
            Err.Clear
            c.SelStart = Len(Nz(c.Value, ""))
        End If
        On Error GoTo Fail
    End If
    Exit Function
Fail:
    HandleError "modKitUI.ToggleSecret", nm
End Function

' Hide the text of every SecretBox on frm. Call it from the form's
' Form_Current, so a field shown on one record is hidden again on the next.
' A button named eye_<x> with no box x beside it is not a SecretBox's and is
' left alone.
Public Sub HideSecrets(frm As Form)
    Dim c As Control
    On Error GoTo Fail
    For Each c In frm.Controls
        If c.ControlType = acCommandButton Then
            If Left$(c.Name, 4) = "eye_" Then
                If HasCtl(frm, Mid$(c.Name, 5)) Then ShowSecret frm, Mid$(c.Name, 5), False
            End If
        End If
    Next
    Exit Sub
Fail:
    RaiseAgain "modKitUI.HideSecrets", frm.Name
End Sub

' Show (reveal True) or hide the text of SecretBox nm, and set its eye
' button's glyph and tip to the next action. On a click the eye button has
' the focus, not the box, so the mask never changes under text being typed.
Private Sub ShowSecret(frm As Form, ByVal nm As String, ByVal reveal As Boolean)
    Dim c As Control, b As Control
    Set c = frm.Controls(nm)
    Set b = frm.Controls("eye_" & nm)
    c.InputMask = IIf(reveal, "", SECRET_MASK)
    b.Caption = ChrW(IIf(reveal, EYE_HIDE, EYE_SHOW))
    b.ControlTipText = IIf(reveal, "Hide", "Show")
End Sub

' ---------------------------------------------------------------------------
' Focus ring
' ---------------------------------------------------------------------------

' A kit input's GotFocus (isOn True) and LostFocus (False), set by Txt and
' Cbo: paint control nm's border CLR_PRIMARY while it has the focus, then give
' back the color it had. That color is whatever was there on GotFocus, not
' always CLR_INPUT_BORDER, so a caller's own border or a red outline from
' CheckRequired survives tabbing through the field. If something else
' repaints the border while the field has the focus, it is left alone.
' Only the border color changes (no width, size or margin), so the text never
' moves. No ring on a locked field or a read-only form (nothing to type), or
' on a continuous form's rows, where one control's border is every row's.
' One field is ringed at a time: each GotFocus first restores the last one,
' because Access fires no LostFocus for a field inside a subform when the
' focus moves out to the main form (FocusRingClear covers the subform's Exit
' and the kit's buttons, check boxes, list boxes and focus sink). A failure is logged but not shown, like ThemeForm:
' the field keeps working without the ring.
Public Function FocusRing(frm As Form, ByVal nm As String, ByVal isOn As Boolean)
    Dim c As Control
    On Error GoTo Fail
    Set c = frm.Controls(nm)
    If Not isOn Then
        If m_ringHwnd = frm.Hwnd And m_ringName = nm Then ClearRing
        Exit Function
    End If
    ClearRing
    If c.Locked Then Exit Function
    If Not (frm.AllowEdits Or (frm.NewRecord And frm.AllowAdditions)) Then Exit Function
    If frm.DefaultView <> 0 And c.Section = acDetail Then Exit Function
    m_ringWas = c.BorderColor
    c.BorderColor = Th(CLR_PRIMARY)
    Set m_ringCtl = c
    m_ringHwnd = frm.Hwnd
    m_ringName = nm
    Exit Function
Fail:
    LogError "modKitUI.FocusRing", frm.Name & "." & nm
End Function

' Take the ring off the field showing it. Set by SubF as the subform
' control's Exit (the field inside gets no LostFocus when the focus leaves the
' subform), and by Btn, Chk, Lst and NewForm's focus sink as their GotFocus:
' a subform's Enter re-rings its last field before the focus lands, so a click
' on its Save button would otherwise leave that field ringed.
Public Function FocusRingClear()
    On Error GoTo Fail
    ClearRing
    Exit Function
Fail:
    LogError "modKitUI.FocusRingClear"
End Function

' A kit subform control's Enter (set by SubF): when the focus comes back to
' the field that last had it inside subform control nm on frm, Access fires
' no GotFocus for it, so ring it here. When the click lands on another
' control, that control's GotFocus (FocusRing or FocusRingClear) follows and
' moves or clears the ring before Access repaints.
Public Function FocusRingEnter(frm As Form, ByVal nm As String)
    Dim inner As Form, c As Control
    On Error Resume Next
    Set inner = frm.Controls(nm).Form      ' error 2467 while it shows no form
    On Error GoTo Fail
    If inner Is Nothing Then Exit Function
    Do
        Set c = Nothing
        On Error Resume Next
        Set c = inner.ActiveControl          ' error 2474 when none has the focus yet
        On Error GoTo Fail
        If c Is Nothing Then Exit Function
        If c.ControlType <> acSubform Then Exit Do
        Set inner = c.Form                   ' a subform inside the subform
    Loop
    If c.ControlType <> acTextBox And c.ControlType <> acComboBox Then Exit Function
    If Left$(Nz(c.OnGotFocus, ""), 11) = "=FocusRing(" Then FocusRing inner, c.Name, True
    Exit Function
Fail:
    LogError "modKitUI.FocusRingEnter", frm.Name & "." & nm
End Function

' Give the ringed field its own border back, if it still shows the ring. Its
' form may have closed since, which makes the saved reference unusable.
Private Sub ClearRing()
    If m_ringCtl Is Nothing Then Exit Sub
    On Error Resume Next
    If m_ringCtl.BorderColor = Th(CLR_PRIMARY) Then m_ringCtl.BorderColor = m_ringWas
    On Error GoTo 0
    Set m_ringCtl = Nothing
    m_ringHwnd = 0
    m_ringName = ""
End Sub

' ---------------------------------------------------------------------------
' Splash screen
' ---------------------------------------------------------------------------

' Show msg on the splash screen's status line (frmSplash, BuildSplash) while
' the app starts, for example "Opening the data file...". Does nothing once
' the splash has closed, so start-up code can call it either way.
Public Sub SplashStatus(ByVal msg As String)
    If Not SplashOpen() Then Exit Sub
    Forms(SPLASH_FORM).Controls("lblSplashStatus").Caption = msg
    Forms(SPLASH_FORM).Repaint
End Sub

' Close the splash screen if it is open. KitMsgBox calls it: the splash is a
' pop-up, so a message opened during start-up could sit behind it. Access can
' refuse (2585) while the splash's own Timer is still running; the splash then
' closes itself when that event ends, so the refusal is ignored.
Public Sub CloseSplash()
    If Not SplashOpen() Then Exit Sub
    On Error Resume Next
    DoCmd.Close acForm, SPLASH_FORM, acSaveNo
    On Error GoTo 0
End Sub

' True when the kit's splash is open in Form view: not one open in Design
' view, and not an app's own form that happens to be called frmSplash (it
' has no lblSplashStatus). A missing control is expected there, so its error
' leaves the answer False.
Private Function SplashOpen() As Boolean
    Dim frm As Form
    If Not IsFormOpen(SPLASH_FORM) Then Exit Function
    On Error Resume Next
    Set frm = Forms(SPLASH_FORM)
    SplashOpen = (frm.CurrentView = 1) And (frm.Controls("lblSplashStatus").Name = "lblSplashStatus")
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------------
' Toggle switches
' ---------------------------------------------------------------------------

' Click, or Space, on a Toggle switch nm (modUIKit.Toggle): flips its value
' (<nm>_val, and so its field), redraws the track and calls the form's
' onChange sub if the switch has one. Does nothing while the value is locked
' or disabled, or the form is read-only (AllowEdits off, unless on a record
' being added), like PickDate.
Public Function ToggleSwitch(frm As Form, ByVal nm As String)
    Dim v As Control, onChange As String
    On Error GoTo Fail
    Set v = frm.Controls(nm & "_val")
    If Not v.Enabled Or v.Locked Then Exit Function
    ' Nested, so an unbound settings form never reads NewRecord.
    If Not frm.AllowEdits Then
        If Not (frm.NewRecord And frm.AllowAdditions) Then Exit Function
    End If
    v.Value = Not CBool(Nz(v.Value, False))
    frm.Controls(nm & "_bg").Requery            ' an unbound image recalculates
    onChange = Nz(frm.Controls(nm).Tag, "")
    If Len(onChange) > 0 Then CallByName frm, onChange, VbMethod
    Exit Function
Fail:
    HandleError "modKitUI.ToggleSwitch", nm
End Function

' Focus moving onto (hasFocus True) or off a Toggle switch nm: shows or hides
' its focus ring, since a transparent button draws no focus of its own, and
' takes an input's ring off (FocusRing) like the kit's buttons do. A failure
' is logged, not shown, so moving the focus never stops on a message.
Public Function ToggleFocus(frm As Form, ByVal nm As String, ByVal hasFocus As Boolean)
    On Error GoTo Fail
    If hasFocus Then ClearRing
    SetVisible frm.Controls(nm & "_ring"), hasFocus
    Exit Function
Fail:
    LogError "modKitUI.ToggleFocus", nm
End Function

' A KeyButton's focus (set by the builder): label nm in CLR_PRIMARY,
' underlined unless it is an MDL2 glyph, while its button has the focus, and
' offColor (a build-palette color, shown through Th) when it leaves. An
' input's focus ring comes off, as for any button.
Public Function FocusMark(frm As Form, ByVal nm As String, ByVal hasFocus As Boolean, ByVal offColor As Long)
    Dim c As Control
    On Error GoTo Fail
    If hasFocus Then FocusRingClear
    Set c = frm.Controls(nm)
    c.ForeColor = IIf(hasFocus, Th(CLR_PRIMARY), Th(offColor))
    If c.FontName <> "Segoe MDL2 Assets" Then c.FontUnderline = hasFocus
    Exit Function
Fail:
    LogError "modKitUI.FocusMark", nm
End Function

' A Segmented option's click (set by the builder): pick option key, move the
' thumb and bold its caption, then call the control's onChange. Does nothing
' when key is already picked, while the value is disabled or locked, or on a
' read-only form.
Public Function SegmentClick(frm As Form, ByVal nm As String, ByVal key As String)
    Dim v As Control, onChange As String
    On Error GoTo Fail
    Set v = frm.Controls(nm & "_val")
    If Not v.Enabled Or v.Locked Then Exit Function
    ' Nested, so an unbound settings form never reads NewRecord.
    If Not frm.AllowEdits Then
        If Not (frm.NewRecord And frm.AllowAdditions) Then Exit Function
    End If
    If Nz(v.Value, "") & "" = key Then Exit Function
    v.Value = key
    SegmentPaint frm, nm
    onChange = Nz(frm.Controls(nm & "_b0").Tag, "")
    If Len(onChange) > 0 Then CallByName frm, onChange, VbMethod
    Exit Function
Fail:
    HandleError "modKitUI.SegmentClick", nm
End Function

' Pick option key of Segmented control nm in code, without calling its
' onChange (as a value set in code fires no AfterUpdate). Raises error 5 for a
' key the control doesn't have. A bound control follows its field on its own:
' call this only to change the value.
Public Sub SetSegment(frm As Form, ByVal nm As String, ByVal key As String)
    Dim v As Control
    Set v = frm.Controls(nm & "_val")
    If InStr(";" & v.Tag & ";", ";" & key & ";") = 0 Then Err.Raise 5, , "Segmented " & nm & " has no option " & key
    v.Value = key
    SegmentPaint frm, nm
End Sub

' The picked key of Segmented control nm, as text ("" when none is).
Public Function SegmentValue(frm As Form, ByVal nm As String) As String
    SegmentValue = Nz(frm.Controls(nm & "_val").Value, "") & ""
End Function

' Redraw a Segmented control after its value changed in code: an unbound
' image and a calculated caption only recalculate when asked.
Private Sub SegmentPaint(frm As Form, ByVal nm As String)
    Dim i As Long, n As Long
    n = UBound(Split(frm.Controls(nm & "_val").Tag, ";")) + 1
    For i = 0 To n - 1
        frm.Controls(nm & "_th" & i).Requery
        frm.Controls(nm & "_t" & i).Requery
    Next
End Sub

' Show Alert nm with message msg (and title, "" for none) in variant vrnt
' (info, success, warning, destructive), in the current palette. A title
' moves the message below it; no title moves it up to the top. Raises error 5
' for an unknown variant.
Public Sub ShowAlert(frm As Form, ByVal nm As String, ByVal msg As String, _
        Optional ByVal vrnt As String = "info", Optional ByVal title As String = "")
    Dim k As Variant, pnl As Control, t As Control, m As Control
    k = AlertColors(vrnt)
    If IsEmpty(k) Then Err.Raise 5, , "Alert " & nm & " has an unknown variant: " & vrnt
    Set pnl = frm.Controls(nm): Set t = frm.Controls(nm & "_title"): Set m = frm.Controls(nm & "_msg")
    pnl.BackColor = Th(k(0))
    pnl.Tag = LCase$(vrnt)
    frm.Controls(nm & "_bar").BackColor = Th(k(1))
    frm.Controls(nm & "_icon").ForeColor = Th(k(1))
    frm.Controls(nm & "_icon").Caption = ChrW(k(2))
    t.Caption = title
    m.Caption = msg
    m.Top = pnl.Top + AlertMsgTop(pnl.Height, Len(title) > 0)
    m.Height = AlertMsgHeight(pnl.Height, Len(title) > 0)
    AlertParts frm, nm, True
    t.Visible = (Len(title) > 0)
End Sub

' Hide Alert nm, every part of it (an alert's close mark calls it).
Public Function HideAlert(frm As Form, ByVal nm As String)
    On Error GoTo Fail
    AlertParts frm, nm, False
    Exit Function
Fail:
    HandleError "modKitUI.HideAlert", nm
End Function

' True while Alert nm shows.
Public Function AlertShown(frm As Form, ByVal nm As String) As Boolean
    AlertShown = frm.Controls(nm).Visible
End Function

Private Sub AlertParts(frm As Form, ByVal nm As String, ByVal v As Boolean)
    Dim s As Variant, c As Control
    For Each s In Array("", "_bar", "_icon", "_title", "_msg", "_x", "_xk")
        Set c = Nothing
        On Error Resume Next
        Set c = frm.Controls(nm & s)
        On Error GoTo 0
        If Not c Is Nothing Then SetVisible c, v
    Next
End Sub

' Show path items ("key|Caption", root first) on Breadcrumb prefix: the last
' is the current page. Raises error 5 for more items than the trail has slots.
Public Sub SetBreadcrumb(frm As Form, ByVal prefix As String, ParamArray items() As Variant)
    Dim path As Variant
    path = items
    LayoutCrumbs frm, prefix, path
End Sub

' A breadcrumb's click (set by the builder): call the trail's handler with the
' crumb's key, unless it is the current page (the last crumb).
Public Function CrumbClick(frm As Form, ByVal prefix As String, ByVal i As Long)
    Dim t() As String
    On Error GoTo Fail
    t = Split(frm.Controls(prefix).Tag, "|")
    If i >= CLng(t(1)) - 1 Then Exit Function
    CallByName frm, t(0), VbMethod, CStr(frm.Controls(prefix & i).Tag)
    Exit Function
Fail:
    HandleError "modKitUI.CrumbClick", prefix & i
End Function

' Lay a breadcrumb path out on its slots, in a form open in design view (the
' builder) or form view (SetBreadcrumb). Crumbs are as wide as their text;
' when the path is wider than the frame, crumbs from the second on fold into
' one ellipsis slot, as few as make it fit (the first and last always show).
Public Sub LayoutCrumbs(frm As Object, ByVal prefix As String, path As Variant)
    Dim frame As Control, c As Control, t() As String, p() As String
    Dim n As Long, slots As Long, i As Long, j As Long, fold As Long, x As Long, sepW As Long
    Dim total As Long, keys() As String, caps() As String, widths() As Long, folded As String, ELLIPSIS As String
    ELLIPSIS = ChrW(&H2026)
    Set frame = frm.Controls(prefix)
    t = Split(frame.Tag, "|")
    slots = CLng(t(2))
    n = 0
    If IsArray(path) Then n = UBound(path) - LBound(path) + 1
    If n > slots Then Err.Raise 5, , "Breadcrumb " & prefix & " has " & slots & " slots, not " & n
    frame.Tag = t(0) & "|" & n & "|" & slots
    sepW = px(16)
    If n > 0 Then
        ReDim keys(n - 1): ReDim caps(n - 1): ReDim widths(n - 1)
        For i = 0 To n - 1
            p = Split(CStr(path(LBound(path) + i)), "|")
            keys(i) = p(0)
            caps(i) = IIf(UBound(p) >= 1, p(UBound(p)), p(0))
            widths(i) = CrumbWidth(caps(i))
            total = total + widths(i) + IIf(i > 0, sepW, 0)
        Next
    End If
    ' Fold crumbs 1..fold into slot fold, shown as an ellipsis, until the path fits.
    If total > frame.Width And n > 2 Then
        For fold = 1 To n - 2
            total = widths(0) + CrumbWidth(ELLIPSIS) + 2 * sepW
            For j = fold + 1 To n - 1
                total = total + widths(j) + sepW
            Next
            If total <= frame.Width Then Exit For
        Next
        If fold > n - 2 Then fold = n - 2
    End If
    x = frame.Left
    For i = 0 To slots - 1
        Set c = frm.Controls(prefix & i)
        If i >= n Or (i >= 1 And i < fold) Then
            c.Visible = False
            ShowKey frm.Controls(prefix & i & "_k"), False
            If i > 0 Then frm.Controls(prefix & "_s" & i).Visible = False
        Else
            If i > 0 Then
                With frm.Controls(prefix & "_s" & i)
                    .Left = x: .Top = frame.Top: .Visible = True
                End With
                x = x + sepW
            End If
            c.Tag = keys(i)
            If i = fold And fold > 0 Then
                folded = ""
                For j = 1 To fold
                    folded = folded & IIf(j > 1, " > ", "") & caps(j)
                Next
                c.Caption = ELLIPSIS
                c.ControlTipText = folded
                c.Width = CrumbWidth(ELLIPSIS)
            Else
                c.Caption = Replace(caps(i), "&", "&&")
                c.ControlTipText = ""
                c.Width = widths(i)
            End If
            c.Left = x: c.Top = frame.Top
            c.ForeColor = Th(IIf(i = n - 1, CLR_FG, CLR_MUTED_FG))
            c.FontUnderline = False
            c.Visible = True
            ' The current page's crumb is no link, so its KeyButton is hidden.
            With frm.Controls(prefix & i & "_k")
                .Left = c.Left: .Top = c.Top: .Width = c.Width
                .Caption = IIf(i = fold And fold > 0, folded, caps(i))
            End With
            ShowKey frm.Controls(prefix & i & "_k"), (i < n - 1)
            x = x + c.Width
        End If
    Next
End Sub

' Show or hide a crumb's KeyButton. The one just pressed has the focus when
' its handler sets a new path, and Access refuses to hide a focused control
' (error 2165), so SetVisible moves the focus first; in design view (the
' builder) nothing has the focus.
Private Sub ShowKey(c As Control, ByVal v As Boolean)
    On Error Resume Next
    c.Visible = v
    If Err.Number = 0 Then Exit Sub
    On Error GoTo 0
    SetVisible c, v
End Sub

' Width of a crumb with caption s (9pt, with a little room for Access's margins).
Private Function CrumbWidth(ByVal s As String) As Long
    CrumbWidth = MeasureText(s, 9, 400) + px(6)
End Function

' Make step current (1-based; steps + 1 shows every step done) on Stepper
' prefix. Raises error 5 for a step out of range.
Public Sub SetStep(frm As Form, ByVal prefix As String, ByVal current As Long)
    Dim t() As String
    t = Split(frm.Controls(prefix).Tag, "|")
    If current < 1 Or current > CLng(t(2)) + 1 Then Err.Raise 5, , "Stepper " & prefix & " has no step " & current
    frm.Controls(prefix).Tag = t(0) & "|" & current & "|" & t(2)
    StepPaint frm, prefix
End Sub

' The current step of Stepper prefix.
Public Function CurrentStep(frm As Form, ByVal prefix As String) As Long
    CurrentStep = CLng(Split(frm.Controls(prefix).Tag, "|")(1))
End Function

' A step's click (set by the builder): call the handler with a done step's
' number. The current step and later ones do nothing.
Public Function StepClick(frm As Form, ByVal prefix As String, ByVal i As Long)
    Dim t() As String
    On Error GoTo Fail
    t = Split(frm.Controls(prefix).Tag, "|")
    If Len(t(0)) = 0 Or i >= CLng(t(1)) Then Exit Function
    CallByName frm, t(0), VbMethod, i
    Exit Function
Fail:
    HandleError "modKitUI.StepClick", prefix & i
End Function

' Paint a Stepper's steps for its current step, in a form open in design view
' (the builder) or form view (SetStep).
Public Sub StepPaint(frm As Object, ByVal prefix As String)
    Dim t() As String, n As Long, cur As Long, i As Long, done As Boolean, numLbl As Control, capLbl As Control
    t = Split(frm.Controls(prefix).Tag, "|")
    cur = CLng(t(1)): n = CLng(t(2))
    For i = 1 To n
        done = (i < cur)
        Set numLbl = frm.Controls(prefix & "_n" & i)
        Set capLbl = frm.Controls(prefix & "_l" & i)
        frm.Controls(prefix & "_f" & i).Visible = (i <= cur)
        frm.Controls(prefix & "_r" & i).Visible = (i > cur)
        If done Then
            numLbl.FontName = "Segoe MDL2 Assets": numLbl.Caption = ChrW(&HE73E)
        Else
            numLbl.FontName = FONT_UI: numLbl.Caption = CStr(i)
        End If
        numLbl.ForeColor = Th(IIf(i <= cur, CLR_BTN_FG, CLR_MUTED_FG))
        capLbl.ForeColor = Th(IIf(i = cur, CLR_FG, CLR_MUTED_FG))
        capLbl.FontWeight = IIf(i = cur, 600, 400)
        If i < n Then frm.Controls(prefix & "_c" & i).BackColor = Th(IIf(done, CLR_PRIMARY, CLR_INPUT_BORDER))
        With frm.Controls(prefix & "_k" & i)
            .TabStop = done And Len(t(0)) > 0
            .Caption = "Step " & i & " of " & n & ", " & capLbl.Caption & IIf(done, ", done", IIf(i = cur, ", current", ""))
        End With
    Next
End Sub

' An Accordion header's click (set by the builder): open the section if
' closed, close it if open (in a single accordion, opening one closes the
' others).
Public Function AccordionClick(frm As Form, ByVal prefix As String, ByVal i As Long)
    On Error GoTo Fail
    OpenSection frm, prefix, i, Not SectionOpen(frm, prefix, i)
    Exit Function
Fail:
    HandleError "modKitUI.AccordionClick", prefix & i
End Function

' Open or close section i of Accordion prefix and move the sections below.
Public Sub OpenSection(frm As Form, ByVal prefix As String, ByVal i As Long, ByVal isOpen As Boolean)
    Dim t() As String, j As Long, n As Long
    t = Split(frm.Controls(prefix).Tag, "|")
    n = UBound(Split(t(1), ",")) + 1
    If i < 1 Or i > n Then Err.Raise 5, , "Accordion " & prefix & " has no section " & i
    If isOpen And t(0) = "single" Then
        For j = 1 To n
            If j <> i Then SetVisible frm.Controls(prefix & "_b" & j), False
        Next
    End If
    ' A body that has the focus can't be hidden: SetVisible moves it first.
    SetVisible frm.Controls(prefix & "_b" & i), isOpen
    AccordionLayout frm, prefix
End Sub

' True while section i of Accordion prefix is open.
Public Function SectionOpen(frm As Form, ByVal prefix As String, ByVal i As Long) As Boolean
    SectionOpen = frm.Controls(prefix & "_b" & i).Visible
End Function

' Stack an Accordion's sections from its top: each header, then its body when
' open, with 8px below it; the chevron points down on an open section. Works
' in design view (the builder) and form view (OpenSection).
Public Sub AccordionLayout(frm As Object, ByVal prefix As String)
    Dim frame As Control, hs() As String, i As Long, y As Long, b As Control
    Set frame = frm.Controls(prefix)
    hs = Split(Split(frame.Tag, "|")(1), ",")
    y = frame.Top
    For i = 1 To UBound(hs) + 1
        Set b = frm.Controls(prefix & "_b" & i)
        frm.Controls(prefix & "_h" & i).Top = y
        frm.Controls(prefix & "_t" & i).Top = y
        frm.Controls(prefix & "_v" & i).Top = y
        frm.Controls(prefix & "_v" & i).Caption = ChrW(IIf(b.Visible, &HE70D, &HE76C))
        frm.Controls(prefix & "_h" & i).Caption = frm.Controls(prefix & "_t" & i).Caption & IIf(b.Visible, ", expanded", ", collapsed")
        frm.Controls(prefix & "_u" & i).Top = y + px(40)
        y = y + px(41)
        If b.Visible Then
            b.Top = y + px(7)
            y = y + px(7) + CLng(hs(i - 1)) + px(8)
        End If
    Next
    frame.Height = y - frame.Top
End Sub

' Show or hide EmptyState nm (every part), with a new title and text when
' given ("" keeps the built ones).
Public Sub ShowEmpty(frm As Form, ByVal nm As String, ByVal isShown As Boolean, _
        Optional ByVal title As String = "", Optional ByVal text As String = "")
    Dim s As Variant, c As Control
    If Len(title) > 0 Then frm.Controls(nm & "_t").Caption = Replace(title, "&", "&&")
    If Len(text) > 0 Then frm.Controls(nm & "_m").Caption = Replace(text, "&", "&&")
    For Each s In Array("", "_bg", "_ico", "_t", "_m", "_b")
        Set c = Nothing
        On Error Resume Next
        Set c = frm.Controls(nm & s)
        On Error GoTo 0
        If Not c Is Nothing Then SetVisible c, isShown
    Next
End Sub

' A SearchBox's typing (set by the builder): show the clear mark once there
' is text, and call the handler with it.
Public Function SearchChanged(frm As Form, ByVal nm As String)
    On Error GoTo Fail
    SearchNotify frm, nm, Nz(frm.Controls(nm).Text, "")
    Exit Function
Fail:
    HandleError "modKitUI.SearchChanged", nm
End Function

' A SearchBox's clear mark (set by the builder): empty the box, put the focus
' back in it and call the handler with "".
Public Function SearchClear(frm As Form, ByVal nm As String)
    On Error GoTo Fail
    frm.Controls(nm).Value = Null
    frm.Controls(nm).SetFocus
    SearchNotify frm, nm, ""
    Exit Function
Fail:
    HandleError "modKitUI.SearchClear", nm
End Function

' A SearchBox's focus (set by the builder): the focus ring is on its frame.
Public Function SearchFocus(frm As Form, ByVal nm As String, ByVal hasFocus As Boolean)
    On Error GoTo Fail
    If hasFocus Then FocusRingClear
    frm.Controls(nm & "_f").BorderColor = Th(IIf(hasFocus, CLR_PRIMARY, CLR_INPUT_BORDER))
    Exit Function
Fail:
    LogError "modKitUI.SearchFocus", nm
End Function

' A SearchBox's text: what is typed so far while it has the focus, else its value.
Public Function SearchText(frm As Form, ByVal nm As String) As String
    Dim c As Control, act As String
    Set c = frm.Controls(nm)
    On Error Resume Next
    act = frm.ActiveControl.Name            ' error 2474 when none has the focus
    On Error GoTo 0
    If act = nm Then SearchText = Nz(c.Text, "") Else SearchText = Nz(c.Value, "")
End Function

' Put text in a SearchBox in code and call its handler, as typing would.
Public Sub SetSearch(frm As Form, ByVal nm As String, ByVal text As String)
    frm.Controls(nm).Value = IIf(Len(text) = 0, Null, text)
    SearchNotify frm, nm, text
End Sub

Private Sub SearchNotify(frm As Form, ByVal nm As String, ByVal text As String)
    Dim handler As String
    SetVisible frm.Controls(nm & "_x"), Len(text) > 0
    SetVisible frm.Controls(nm & "_xk"), Len(text) > 0
    handler = Nz(frm.Controls(nm & "_f").Tag, "")
    If Len(handler) > 0 Then CallByName frm, handler, VbMethod, text
End Sub

' A Rating star's click (set by the builder): set the value to star i, or to
' 0 when it is already i, then call onChange. Not while the value is disabled
' or locked, or on a read-only form.
Public Function RatingClick(frm As Form, ByVal nm As String, ByVal i As Long)
    Dim v As Control, onChange As String
    On Error GoTo Fail
    Set v = frm.Controls(nm & "_val")
    If Not v.Enabled Or v.Locked Then Exit Function
    If Not frm.AllowEdits Then
        If Not (frm.NewRecord And frm.AllowAdditions) Then Exit Function
    End If
    v.Value = IIf(Nz(v.Value, 0) = i, 0, i)
    RatingPaint frm, nm
    onChange = Nz(frm.Controls(nm & "_b1").Tag, "")
    If Len(onChange) > 0 Then CallByName frm, onChange, VbMethod
    Exit Function
Fail:
    HandleError "modKitUI.RatingClick", nm
End Function

' Set Rating nm to v (0 to its stars) in code, without calling onChange.
Public Sub SetRating(frm As Form, ByVal nm As String, ByVal v As Long)
    If v < 0 Or v > CLng(frm.Controls(nm & "_val").Tag) Then Err.Raise 5, , "Rating " & nm & " has no value " & v
    frm.Controls(nm & "_val").Value = v
    RatingPaint frm, nm
End Sub

' The value of Rating nm (0 when empty).
Public Function RatingValue(frm As Form, ByVal nm As String) As Long
    RatingValue = Nz(frm.Controls(nm & "_val").Value, 0)
End Function

Private Sub RatingPaint(frm As Form, ByVal nm As String)
    Dim i As Long
    For i = 1 To CLng(frm.Controls(nm & "_val").Tag)
        frm.Controls(nm & "_s" & i).Requery
    Next
End Sub

' A Slider step's click (set by the builder): move to step i, then call
' onChange (not when it is already there, while the value is disabled or
' locked, or on a read-only form).
Public Function SliderClick(frm As Form, ByVal nm As String, ByVal i As Long)
    Dim t() As String, v As Control, nv As Double, onChange As String
    On Error GoTo Fail
    Set v = frm.Controls(nm & "_val")
    If Not v.Enabled Or v.Locked Then Exit Function
    If Not frm.AllowEdits Then
        If Not (frm.NewRecord And frm.AllowAdditions) Then Exit Function
    End If
    t = Split(v.Tag, "|")
    nv = Val(t(0)) + i * Val(t(2))
    If Not IsNull(v.Value) Then If Abs(CDbl(v.Value) - nv) < Val(t(2)) / 1000 Then Exit Function
    SliderSet frm, nm, nv
    onChange = Nz(frm.Controls(nm & "_b0").Tag, "")
    If Len(onChange) > 0 Then CallByName frm, onChange, VbMethod
    Exit Function
Fail:
    HandleError "modKitUI.SliderClick", nm
End Function

' Move Slider nm by delta steps (from the form's KeyDown: -1 for Left, 1 for
' Right), kept in its range, and call onChange when it moved.
Public Sub SliderStep(frm As Form, ByVal nm As String, ByVal delta As Long)
    Dim t() As String, i As Long, n As Long
    t = Split(frm.Controls(nm & "_val").Tag, "|")
    n = CLng((Val(t(1)) - Val(t(0))) / Val(t(2)))
    i = CLng((Nz(frm.Controls(nm & "_val").Value, Val(t(0))) - Val(t(0))) / Val(t(2))) + delta
    If i < 0 Then i = 0
    If i > n Then i = n
    SliderClick frm, nm, i
End Sub

' Set Slider nm to v in code (rounded to its nearest step, kept in its range),
' without calling onChange.
Public Sub SetSlider(frm As Form, ByVal nm As String, ByVal v As Double)
    Dim t() As String, i As Long, n As Long
    t = Split(frm.Controls(nm & "_val").Tag, "|")
    n = CLng((Val(t(1)) - Val(t(0))) / Val(t(2)))
    i = CLng((v - Val(t(0))) / Val(t(2)))
    If i < 0 Then i = 0
    If i > n Then i = n
    SliderSet frm, nm, Val(t(0)) + i * Val(t(2))
End Sub

' The value of Slider nm.
Public Function SliderValue(frm As Form, ByVal nm As String) As Double
    SliderValue = Nz(frm.Controls(nm & "_val").Value, Val(Split(frm.Controls(nm & "_val").Tag, "|")(0)))
End Function

Private Sub SliderSet(frm As Form, ByVal nm As String, ByVal v As Double)
    Dim t() As String, i As Long
    frm.Controls(nm & "_val").Value = v
    frm.Controls(nm & "_bg").Requery
    ' The current step takes the Tab stop, so Tab lands on the knob.
    t = Split(frm.Controls(nm & "_val").Tag, "|")
    For i = 0 To CLng((Val(t(1)) - Val(t(0))) / Val(t(2)))
        frm.Controls(nm & "_b" & i).TabStop = (Abs(Val(t(0)) + i * Val(t(2)) - v) < Val(t(2)) / 1000)
    Next
    On Error Resume Next
    frm.Controls(nm & "_lbl").Requery
End Sub

' ---------------------------------------------------------------------------
' Messages
' ---------------------------------------------------------------------------

' MsgBox in the kit's look: opens the message dialog (dlgMessage, built by
' BuildMessageDialog) and waits for a button. Windows centres a MsgBox on the
' monitor, so on a large or second screen it opens outside Access and Access
' waits behind it; dlgMessage is centred over Access like the app's other
' dialogs. Takes the buttons vbOKOnly, vbOKCancel, vbYesNo and vbYesNoCancel,
' an icon (vbInformation, vbExclamation, vbQuestion, vbCritical) and
' vbDefaultButton2/3, and returns vbOK, vbCancel, vbYes or vbNo like MsgBox.
' Any other button set, a message while one is already showing, or a dialog
' that can't open (not built yet) uses MsgBox itself.
Public Function KitMsgBox(ByVal msg As String, Optional ByVal style As VbMsgBoxStyle = vbOKOnly, _
        Optional ByVal title As String = "") As VbMsgBoxResult
    Dim btns As Long
    btns = style And 7
    CloseSplash                    ' a start-up message must not open behind the splash
    If m_msgOpen Then GoTo Plain
    If Not (btns = vbOKOnly Or btns = vbOKCancel Or btns = vbYesNo Or btns = vbYesNoCancel) Then GoTo Plain
    On Error GoTo Failed
    m_msgText = msg
    m_msgStyle = style
    m_msgTitle = title
    m_msgAnswer = 0
    m_msgOpen = True
    DimApp                          ' the app behind is dimmed, as for RunDialog
    DoCmd.OpenForm "dlgMessage", WindowMode:=acDialog
    UndimApp
    m_msgOpen = False
    If m_msgAnswer = 0 Then        ' closed some other way: as Esc
        m_msgAnswer = Choose(btns + 1, vbOK, vbCancel, vbCancel, vbCancel, vbNo, vbCancel)
    End If
    KitMsgBox = m_msgAnswer
    Exit Function
Failed:
    ' Only this message's own dialog failed to open; a message already showing
    ' (Plain from the first line) keeps m_msgOpen set until it closes.
    ' m_msgOpen is set just before DimApp (which never raises), so it also says
    ' this message's dim is still on.
    If m_msgOpen Then UndimApp
    m_msgOpen = False
Plain:
    KitMsgBox = MsgBox(msg, style, title)
End Function

' What the open message dialog shows, and its answer (dlgMessage calls these).
Public Function KitMsgText() As String
    KitMsgText = m_msgText
End Function

Public Function KitMsgStyle() As Long
    KitMsgStyle = m_msgStyle
End Function

Public Function KitMsgTitle() As String
    KitMsgTitle = m_msgTitle
End Function

Public Sub SetKitMsgAnswer(ByVal answer As Long)
    m_msgAnswer = answer
End Sub

' A message that doesn't interrupt: the toast (dlgToast, built by BuildToast)
' shows msg in the bottom-right corner of frmMain for ms milliseconds, then
' fades and closes. It never takes the focus, so use it to confirm something
' the user just did ("Saved"); anything they must read or answer belongs in
' KitMsgBox. vrnt picks the icon and accent color: "success", "info" (the
' default), "warning" or "destructive". A toast while one is showing
' replaces it. The toast shows up to 3 lines, so keep the message short.
' If the toast can't open (not built yet), the message goes to the status
' bar and the failure is logged, never shown.
Public Function KitToast(ByVal msg As String, Optional ByVal vrnt As String = "info", Optional ByVal ms As Long = 3000)
    On Error GoTo Fail
    If ms <= 0 Then ms = 3000
    If Not IsFormOpen("dlgToast") Then DoCmd.OpenForm "dlgToast", WindowMode:=acHidden
    Forms("dlgToast").ShowToast msg, vrnt, ms
    Exit Function
Fail:
    On Error Resume Next
    LogError "modKitUI.KitToast", Left$(msg, 60)
    SysCmd acSysCmdSetStatus, Left$(msg, 255)
End Function

' Busy pop-up (dlgBusy, built by BuildBusy): KitBusy msg dims the app and
' shows msg in a small pop-up in the middle of it, with a bar sliding under
' the text, until KitBusyEnd. Use it around work the user could take for a
' hang: relinking to a data file on a share, waiting out another user's
' lock, an import or export. Pair every KitBusy with one KitBusyEnd, on
' the error path too.
'   KitBusy "Exporting the list..."
'   On Error GoTo Fail
'   ... the work ...
'   KitBusyEnd
'   Exit Sub
' Fail:
'   num = Err.Number: d = Err.Description: src = Err.Source
'   KitBusyEnd                      ' resets Err, so it is read first
'   Err.Number = num: Err.Description = d: Err.Source = src
'   HandleError "pgItems.Export"
' Busy work inside busy work shares the pop-up: the inner message shows
' until its KitBusyEnd, then the outer one comes back. The pop-up never
' takes the focus. The bar moves only while the work lets Access run
' (DoEvents); in work that doesn't, the message still shows. An app that
' hasn't built dlgBusy gets msg in the status bar instead, with no dim;
' if dlgBusy fails to open, the same, and the failure is logged.
Public Sub KitBusy(ByVal msg As String)
    If m_busy Is Nothing Then Set m_busy = New Collection
    m_busy.Add msg
    On Error GoTo Fail
    If m_busy.Count = 1 Then
        m_busyPopup = BusyFormBuilt()       ' without it, the status bar, quietly: the kit calls KitBusy itself
        If m_busyPopup Then DimApp          ' never raises
    End If
    If Not m_busyPopup Then GoTo StatusBar
    If Not IsFormOpen("dlgBusy") Then DoCmd.OpenForm "dlgBusy", WindowMode:=acHidden
    Forms("dlgBusy").ShowBusy msg
    Exit Sub
Fail:
    LogError "modKitUI.KitBusy", Left$(msg, 60)
StatusBar:
    On Error Resume Next
    SysCmd acSysCmdSetStatus, Left$(msg, 255)
    m_busyStatus = True
End Sub

' End the latest KitBusy: the message before it comes back, or after the
' last one the pop-up closes and the dim goes. Never raises, so it is safe
' in an error handler; it does reset Err, so read Err first.
Public Sub KitBusyEnd()
    Dim msg As String
    On Error Resume Next
    If KitBusyDepth() = 0 Then Exit Sub
    m_busy.Remove m_busy.Count
    If m_busy.Count > 0 Then
        msg = m_busy(m_busy.Count)
        If IsFormOpen("dlgBusy") Then Forms("dlgBusy").ShowBusy msg
        If m_busyStatus Then SysCmd acSysCmdSetStatus, Left$(msg, 255)
        Exit Sub
    End If
    If IsFormOpen("dlgBusy") Then DoCmd.Close acForm, "dlgBusy", acSaveNo
    If m_busyStatus Then SysCmd acSysCmdClearStatus
    If m_busyPopup Then UndimApp
    m_busyStatus = False
    m_busyPopup = False
End Sub

' True when the app has a dlgBusy form (BuildBusy).
Private Function BusyFormBuilt() As Boolean
    On Error Resume Next
    BusyFormBuilt = (Len(CurrentProject.AllForms("dlgBusy").Name) > 0)
End Function

' KitBusy calls still waiting for their KitBusyEnd. dlgBusy closes itself
' when this is 0, so a VBA reset during busy work can't leave it showing.
Public Function KitBusyDepth() As Long
    If Not m_busy Is Nothing Then KitBusyDepth = m_busy.Count
End Function

' InputBox in the kit's look: opens the input dialog (dlgInput, built by
' BuildInputDialog), centred over Access like dlgMessage, and waits. Returns
' the text typed when OK is clicked (a String, "" if the box was left empty)
' or Null on Cancel, Esc or any other close, so an empty answer and a cancel
' can be told apart. masked:=True shows each typed character as * (the
' SECRET_MASK input mask) for a PIN or password; the text returned is the real
' one.
' A question while one is already showing, or a dialog that can't open (not
' built yet), uses InputBox itself, with Null for its Cancel. A masked
' question never does: InputBox would show the secret, so it returns Null
' instead (logged when the dialog failed to open).
Public Function KitInputBox(ByVal prompt As String, Optional ByVal title As String = "", _
        Optional ByVal defaultText As String = "", Optional ByVal masked As Boolean = False) As Variant
    Dim s As String
    If m_inOpen Then GoTo Plain
    On Error GoTo Failed
    m_inPrompt = prompt
    m_inTitle = title
    m_inDefault = defaultText
    m_inMasked = masked
    m_inAnswer = Null
    m_inOpen = True
    CloseSplash                    ' a start-up question must not open behind the splash
    DimApp                          ' the app behind is dimmed, as for KitMsgBox
    DoCmd.OpenForm "dlgInput", WindowMode:=acDialog
    UndimApp
    m_inOpen = False
    KitInputBox = m_inAnswer
    ClearInput                     ' don't keep a typed PIN or password in memory
    Exit Function
Failed:
    ' Only this question's own dialog failed to open; a question already
    ' showing (Plain from the first line) keeps m_inOpen set until it closes.
    ' m_inOpen is set just before DimApp (which never raises), so it also says
    ' this question's dim is still on. Log first: UndimApp may clear Err.
    If masked Then LogError "modKitUI.KitInputBox", "masked input dialog did not open"
    If m_inOpen Then UndimApp
    m_inOpen = False
    ClearInput
    If masked Then
        KitInputBox = Null
        Exit Function
    End If
Plain:
    If masked Then          ' asked while another question is open: no error to log
        KitInputBox = Null
        Exit Function
    End If
    s = InputBox(prompt, IIf(Len(title) > 0, title, APP_NAME), defaultText)
    If StrPtr(s) = 0 Then KitInputBox = Null Else KitInputBox = s   ' StrPtr 0: Cancel
End Function

' What the open input dialog asks, and its answer (dlgInput calls these).
Public Function KitInputPrompt() As String
    KitInputPrompt = m_inPrompt
End Function

Public Function KitInputTitle() As String
    KitInputTitle = m_inTitle
End Function

Public Function KitInputDefault() As String
    KitInputDefault = m_inDefault
End Function

Public Function KitInputMasked() As Boolean
    KitInputMasked = m_inMasked
End Function

Public Sub SetKitInputAnswer(ByVal answer As Variant)
    m_inAnswer = answer
End Sub

' Forget the last question and its answer once KitInputBox has returned it.
Private Sub ClearInput()
    m_inPrompt = "": m_inTitle = "": m_inDefault = "": m_inMasked = False
    m_inAnswer = Null
End Sub

' ---------------------------------------------------------------------------
' Required fields
' ---------------------------------------------------------------------------

' Save check for a form's required inputs: every empty one (Null, or only
' spaces) is outlined in CLR_DESTRUCTIVE, all of them are listed in one
' KitMsgBox, and the focus goes to the first. fields lists the inputs in tab
' order as "control=Caption" pairs separated by ";", for example
' "txtName=Item;txtOwner=Owner". Returns True when all are filled. Call it at
' the top of the Save button's handler and stop when it returns False:
'     If Not CheckRequired(Me, "txtName=Item;txtOwner=Owner") Then Exit Sub
' and in the form's Form_BeforeUpdate, so every other save is checked too:
'     Cancel = Not CheckRequired(Me, "txtName=Item;txtOwner=Owner")
' (during an update the focus can't move, so it stays where it is).
Public Function CheckRequired(frm As Form, ByVal fields As String) As Boolean
    Dim missing As String, first As String
    On Error GoTo Fail
    missing = RequiredMissing(frm, fields, first)
    If Len(missing) = 0 Then
        CheckRequired = True
        Exit Function
    End If
    KitRefuseSave
    KitMsgBox "Please fill in:" & vbCrLf & missing, vbExclamation, "Required fields"
    ' Moving the focus is a courtesy: a control that can't take it (error
    ' 2110) or a form that isn't active any more is left as it is.
    On Error Resume Next
    frm.Controls(first).SetFocus
    Exit Function
Fail:
    HandleError "modKitUI.CheckRequired", frm.Name & ": " & fields
End Function

' A form's BeforeUpdate that cancels a save and shows its own message calls
' this first, so SaveDetail and KitSaveRecord take the error Access raises
' next as a refusal it has explained (CheckRequired does it for you).
Public Sub KitRefuseSave()
    m_refused = True
End Sub

' True once after KitRefuseSave (it clears the note); the save routines call
' it before a save, to clear a stale note, and after a failed one.
Public Function KitTakeRefused() As Boolean
    KitTakeRefused = m_refused
    m_refused = False
End Function

' CheckRequired without the message: outline the empty required inputs and
' return their captions, one per line after a bullet (an empty string when
' all are filled), and the first one's control name in first. A filled input
' gets its normal border back, so a field fixed since the last Save stops
' being red. Hidden or disabled inputs are skipped: the user can't fill them.
' Raises on a control name the form doesn't have.
Public Function RequiredMissing(frm As Form, ByVal fields As String, Optional ByRef first As String) As String
    Dim p As Variant, nm As String, cap As String, c As Control, out As String, blank As Boolean
    first = ""
    For Each p In Split(fields, ";")
        If Len(Trim$(p)) > 0 Then
            nm = Trim$(p): cap = nm
            If InStr(nm, "=") > 0 Then
                cap = Trim$(Mid$(nm, InStr(nm, "=") + 1))
                nm = Trim$(Left$(nm, InStr(nm, "=") - 1))
            End If
            Set c = frm.Controls(nm)
            blank = False
            If c.Visible And c.Enabled Then blank = (Len(Trim$(Nz(c.Value, ""))) = 0)
            If blank Then
                c.BorderColor = Th(CLR_DESTRUCTIVE)
                If Len(first) = 0 Then first = nm
                out = out & IIf(Len(out) > 0, vbCrLf, "") & ChrW(&H2022) & " " & cap
            Else
                ' The border the kit's builders give the input (Txt, DateBox, Cbo; Lst).
                c.BorderColor = Th(IIf(c.ControlType = acListBox, CLR_BORDER, CLR_INPUT_BORDER))
            End If
        End If
    Next
    RequiredMissing = out
End Function

' ---------------------------------------------------------------------------
' Tabs and filter chips
' ---------------------------------------------------------------------------

' Show a count after a tab or chip caption, "Name (n)"; Null shows the name
' alone. The caption without the count is kept in the control's Tag, so the
' count can change any number of times.
Public Sub SetCount(c As Control, ByVal n As Variant)
    Dim cap As String
    If Len(c.Tag) = 0 Then
        cap = c.Caption
        If Right$(cap, 1) = ")" And InStrRev(cap, " (") > 0 Then cap = Left$(cap, InStrRev(cap, " (") - 1)
        c.Tag = cap
    End If
    c.Caption = c.Tag & IIf(IsNull(n), "", " (" & n & ")")
End Sub

' Filter chip (a label): filled CLR_PRIMARY with white text when on,
' transparent with muted text when off. n, when given, is its count (SetCount).
Public Sub StyleChip(lbl As Control, ByVal isOn As Boolean, Optional ByVal n As Variant)
    lbl.BackStyle = IIf(isOn, 1, 0)
    lbl.BackColor = Th(CLR_PRIMARY)
    lbl.ForeColor = IIf(isOn, Th(CLR_WHITE), Th(CLR_MUTED_FG))
    If Not IsMissing(n) Then SetCount lbl, n
End Sub

' Paint ChipRow prefix's chips: chip i on (StyleChip), the others off; -1
' turns them all off.
Public Sub ChipOn(frm As Form, ByVal prefix As String, ByVal i As Long)
    Dim n As Long, j As Long
    n = CLng(Split(frm.Controls(prefix).Tag, "|")(1))
    For j = 0 To n - 1
        StyleChip frm.Controls(prefix & "_" & j), j = i
    Next
End Sub

' Show tab key of the form's TabStrip as the open one: its caption CLR_FG
' with its underline, the other captions CLR_MUTED_FG without. Call it on
' load (the builder shows every underline) and from the strip's handler.
Public Sub SelectTab(frm As Form, ByVal key As String)
    Dim c As Control
    For Each c In frm.Controls
        If c.Name Like "tab_*" Then c.ForeColor = Th(IIf(StrComp(Mid$(c.Name, 5), key, vbTextCompare) = 0, CLR_FG, CLR_MUTED_FG))
        If c.Name Like "ul_*" Then c.Visible = (StrComp(Mid$(c.Name, 4), key, vbTextCompare) = 0)
    Next
End Sub

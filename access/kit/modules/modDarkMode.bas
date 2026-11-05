Option Compare Database
Option Explicit

' Access UI kit - color palettes (light, dark and the others in
' modPalette). Forms are built once in the light palette (modTheme), where
' every color role has its own value. Every form calls ThemeForm from its Open
' event; for any other palette that repaints the form's colors, conditional
' formats and pill images through modPalette.PaletteOf. Runtime code wraps
' palette colors in Th(). On Windows 11, ThemeTitleBar also colors pop-up
' dialogs' title bars and borders.
'
' The choice is kept per Windows user (registry, HKCU\Software\VB and VBA
' Program Settings\<APP_NAME>\UI, value Theme = the palette key). Until a user
' picks one, the app follows the Windows "app mode" setting (Navy or Midnight).

Private Const PROPS As String = "BackColor,ForeColor,BorderColor,HoverColor,PressedColor,HoverForeColor,PressedForeColor"
Private m_pal As Long               ' 0 = not read yet, else the palette index (1 = light)
' The OnOpen every kit form is built with (modUIKit.NewForm).
Public Const THEME_HOOK As String = "=ThemeForm([Form])"

' Windows 11 lets an app color a window's title bar and border (DWM). Pop-up
' forms are real top-level windows, so dialogs can take the palette instead
' of showing a white title bar over a dark dialog. Older Windows rejects the
' attributes and keeps its own title bar.
Private Declare PtrSafe Function DwmSetWindowAttribute Lib "dwmapi" ( _
    ByVal hwnd As LongPtr, ByVal dwAttribute As Long, ByRef pvAttribute As Long, _
    ByVal cbAttribute As Long) As Long

Private Const DWMWA_USE_IMMERSIVE_DARK_MODE As Long = 20   ' light or dark caption buttons
Private Const DWMWA_BORDER_COLOR As Long = 34
Private Const DWMWA_CAPTION_COLOR As Long = 35
Private Const DWMWA_TEXT_COLOR As Long = 36

' Index of the current palette. Read once per session, then cached.
Public Function CurrentPalette() As Long
    Dim s As String
    If m_pal = 0 Then
        s = GetSetting(APP_NAME, "UI", "Theme", "")
        If Len(s) = 0 Then s = IIf(WindowsDark(), "dark", "light")
        m_pal = PaletteIndex(s)
    End If
    CurrentPalette = m_pal
End Function

' Index of the palette saved as key (1, the light palette, for an unknown key).
Private Function PaletteIndex(ByVal key As String) As Long
    Dim i As Long
    PaletteIndex = 1
    For i = 1 To PALETTE_COUNT
        If PaletteKey(i) = key Then PaletteIndex = i
    Next
End Function

' True when the current palette has dark surfaces.
Public Function IsDark() As Boolean
    IsDark = PaletteIsDark(CurrentPalette())
End Function

' Build time: colors written into form designs are always the light palette.
Public Sub UseBuildPalette()
    m_pal = 1
End Sub

' Palette color for the current palette.
Public Function Th(ByVal c As Long) As Long
    Th = PaletteOf(c, CurrentPalette())
End Function

' Windows "Choose your app mode" setting (AppsUseLightTheme = 0 means dark).
Private Function WindowsDark() As Boolean
    On Error Resume Next
    WindowsDark = (CreateObject("WScript.Shell").RegRead( _
        "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme") = 0)
End Function

' Sidebar "Appearance": pick a palette in dlgAppearance and apply it.
Public Function ShowAppearance()
    On Error GoTo Err_Handler
    TempVars("PalettePick") = 0
    If RunDialog("dlgAppearance") Then SetPalette CLng(TempVars("PalettePick"))
    Exit Function
Err_Handler:
    HandleError "modDarkMode.ShowAppearance"
End Function

' Switch palettes and redraw the app: the main window is reopened on the same
' page, which rebuilds every form from its light design.
Public Sub SetPalette(ByVal i As Long)
    Dim pg As String, shownTopic As String, reopenHelp As Boolean
    Dim frm As Form, help As Object
    If i < 1 Or i > PALETTE_COUNT Or i = CurrentPalette() Then Exit Sub
    m_pal = i
    SaveSetting APP_NAME, "UI", "Theme", PaletteKey(i)
    ' The help window (modKitHelp, when the app has it) stays open beside the
    ' page, so it is closed and opened again on the same topic in the new
    ' colors. Its form is read late-bound, so apps without help still compile.
    If IsFormOpen("dlgHelp") Then
        Set help = Forms("dlgHelp")
        shownTopic = help.HelpTopic
        Set help = Nothing
        DoCmd.Close acForm, "dlgHelp", acSaveNo
        reopenHelp = True
    End If
    ' A toast showing now is closed: the main window reopening below would
    ' leave it at the top left with its placeholder text and no fade.
    If IsFormOpen("dlgToast") Then DoCmd.Close acForm, "dlgToast", acSaveNo
    ' Other pop-ups still open keep their contents until reopened, but their
    ' title bars follow the new palette now (forms opened later get it in
    ' ThemeForm).
    For Each frm In Forms
        ThemeTitleBar frm
    Next
    If IsFormOpen("frmMain") Then
        pg = Forms("frmMain").PageKey
        Forms("frmMain").Reopen pg
    End If
    If reopenHelp Then DoCmd.OpenForm "dlgHelp", OpenArgs:=shownTopic
End Sub

' Form Open hook (each form's OnOpen is =ThemeForm([Form])).
' A theming problem is logged but not shown: the form still opens, just with
' some light colors.
' A form must be themed only once: Pal maps light colors, and some palette
' colors are also light colors (Midnight's CLR_FG is the light CLR_ROW1), so a
' second pass turns them into other colors. Access can run a form's OnOpen
' again on the colors it already has: a table widened by GridFit opens again
' with its text dark on Midnight's dark rows. So once the form is themed its
' hook is taken off; a form opened afresh from its design still has it.
Public Function ThemeForm(frm As Object)
    Dim c As Control, i As Long
    On Error GoTo Fail
    ThemeTitleBar frm                   ' every palette, light included
    If CurrentPalette() = 1 Then Exit Function
    For i = 0 To 4
        If HasSection(frm, i) Then
            frm.Section(i).BackColor = Pal(frm.Section(i).BackColor)
            frm.Section(i).AlternateBackColor = Pal(frm.Section(i).AlternateBackColor)
        End If
    Next
    For Each c In frm.Controls
        ThemeControl c
    Next
    If frm.OnOpen = THEME_HOOK Then frm.OnOpen = ""
    Exit Function
Fail:
    LogError "modDarkMode.ThemeForm", "form " & frm.Name
End Function

' Color a pop-up form's title bar and border with the current palette: the
' caption matches the dialog background (CLR_DIALOG), with CLR_FG text and a
' CLR_BORDER edge. Forms inside the Access window (frmMain, pages, subforms)
' have no title bar of their own and are skipped. Before Windows 11 the color
' calls fail with an HRESULT, which is ignored; Windows 10 (20H1 and later)
' still honors the dark-mode flag, so dark palettes get its plain dark title
' bar. Any other problem is logged, never shown.
Public Sub ThemeTitleBar(frm As Object)
    Dim h As LongPtr
    On Error GoTo Fail
    If Not frm.PopUp Then Exit Sub
    h = frm.hwnd
    If h = 0 Then Exit Sub
    SetDwmLong h, DWMWA_USE_IMMERSIVE_DARK_MODE, IIf(IsDark(), 1, 0)
    SetDwmLong h, DWMWA_CAPTION_COLOR, Th(CLR_DIALOG)
    SetDwmLong h, DWMWA_TEXT_COLOR, Th(CLR_FG)
    SetDwmLong h, DWMWA_BORDER_COLOR, Th(CLR_BORDER)
    Exit Sub
Fail:
    LogError "modDarkMode.ThemeTitleBar", "form " & frm.Name
End Sub

' One DWM attribute holding a 4-byte value (a COLORREF, which has the same
' layout as a VBA color, or a BOOL). The HRESULT is ignored on purpose: a
' failure means this Windows can't do it, and the default title bar is fine.
Private Sub SetDwmLong(ByVal h As LongPtr, ByVal attr As Long, ByVal v As Long)
    DwmSetWindowAttribute h, attr, v, 4
End Sub

' Forms only have the sections they were built with (detail, maybe header/footer).
Private Function HasSection(frm As Object, ByVal i As Long) As Boolean
    Dim h As Long
    On Error Resume Next
    h = frm.Section(i).Height
    HasSection = (Err.Number = 0)
End Function

' Repaint one control: every color property, conditional formats, the colors
' inside rounded image expressions, and dark-only extras (Tag "darkonly").
Private Sub ThemeControl(c As Control)
    Dim p As Variant, fc As Object, src As String
    On Error GoTo Fail
    For Each p In Split(PROPS, ",")
        RepaintProp c, CStr(p)
    Next
    Select Case c.ControlType
        Case acTextBox, acComboBox
            For Each fc In c.FormatConditions
                If Pal(fc.BackColor) <> fc.BackColor Then fc.BackColor = Pal(fc.BackColor)
                If Pal(fc.ForeColor) <> fc.ForeColor Then fc.ForeColor = Pal(fc.ForeColor)
            Next
        Case acImage
            ' Rounded images (modPng): =PillPng(Switch(cond,color,...),w,h),
            ' =CardPng(fill,border,back,r,corner), =ProgressPng(pct,fill,track,w,h)
            If c.Tag = "RoundedCard" Then ResizeRoundedCard c, c.Width, c.Height: Exit Sub
            src = Nz(c.Properties("ControlSource"), "")
            If InStr(src, "Png(") > 0 Then c.Properties("ControlSource") = PaletteColorsIn(src)
    End Select
    If c.Tag = "darkonly" Then c.Visible = IsDark()
    Exit Sub
Fail:
    ' Log and carry on, so one control can't leave the rest of the form light.
    LogError "modDarkMode.ThemeControl", "control " & c.Name
End Sub

' Repaint one color property. Each control type has only some of PROPS, so a
' missing property is expected and skipped.
Private Sub RepaintProp(c As Control, ByVal prop As String)
    Dim v As Long
    On Error Resume Next
    v = c.Properties(prop)
    If Err.Number <> 0 Then Exit Sub
    If Pal(v) <> v Then c.Properties(prop) = Pal(v)
End Sub

' Replace every palette color written as a number in an expression.
Private Function PaletteColorsIn(ByVal s As String) As String
    Dim i As Long, j As Long, out As String, n As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 1) Like "#" And Not (Mid$(s, IIf(i > 1, i - 1, 1), 1) Like "[A-Za-z_#0-9]" And i > 1) Then
            j = i
            Do While j <= Len(s) And Mid$(s, j, 1) Like "#"
                j = j + 1
            Loop
            n = Mid$(s, i, j - i)
            If Len(n) >= 5 And Len(n) <= 8 Then n = CStr(Pal(CLng(n)))
            out = out & n
            i = j
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    PaletteColorsIn = out
End Function

' Current palette's color for a build color.
Private Function Pal(ByVal c As Long) As Long
    Pal = PaletteOf(c, m_pal)
End Function

Option Compare Database
Option Explicit

' Palette picker (sidebar Appearance). Each row shows the palette's sidebar,
' accent and page colors; the current palette is highlighted and checked.
' Result in TempVars("PalettePick"): the palette index (modDarkMode.ShowAppearance
' applies it).

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim i As Long, k As Long, isCur As Boolean
    For i = 1 To PALETTE_COUNT
        For k = 1 To 3
            Me("sw" & i & "_" & k).BackColor = PaletteSwatch(i, k)
        Next
        isCur = (i = CurrentPalette())
        Me("lblChk" & i).Visible = isCur
        Me("lblName" & i).FontWeight = IIf(isCur, 600, 400)
        If isCur Then
            Me("boxRow" & i).BackColor = Th(CLR_PRIMARY_SOFT)
            Me("boxRow" & i).BorderColor = Th(CLR_PRIMARY)
        End If
    Next
    CenterOnApp Me                  ' over Access, on its monitor (modKitUtil)
    Exit Sub
Err_Handler:
    HandleError "dlgAppearance.Form_Load"
End Sub

' Row click: return the palette and close.
Public Function Pick(ByVal i As Long)
    On Error GoTo Err_Handler
    TempVars("PalettePick") = i
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Function
Err_Handler:
    HandleError "dlgAppearance.Pick"
End Function

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgAppearance.btnCancel_Click"
End Sub

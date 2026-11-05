Option Compare Database
Option Explicit

' Form designer page (BuildDesigner, modKitDesigner): every event goes to the
' Designer* routines there.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    DesignerLoad Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler
    KeyCode = DesignerKeyDown(Me, KeyCode, Shift)
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.Form_KeyDown"
End Sub

' The shell asks before showing another page: unsaved changes are saved or dropped first.
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = DesignerCanLeave(Me)
    Exit Function
Err_Handler:
    HandleError "pgDesigner.CanLeave"
End Function

' The canvas: select (Shift-click adds, a box selects what it encloses),
' drag, resize, and a double-click for the element's settings.
Private Sub dzGlass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Handler
    DesignerMouse Me, 1, Button, X, Y, Shift
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzGlass_MouseDown"
End Sub

Private Sub dzGlass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Handler
    DesignerMouse Me, 2, Button, X, Y, Shift
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzGlass_MouseMove"
End Sub

Private Sub dzGlass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Handler
    DesignerMouse Me, 3, Button, X, Y, Shift
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzGlass_MouseUp"
End Sub

Private Sub dzGlass_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler
    DesignerEdit Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzGlass_DblClick"
End Sub

' The palette: one transparent button over every tile.
Private Sub dzPal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Handler
    DesignerToolMouse Me, 1, Button, X, Y
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzPal_MouseDown"
End Sub

Private Sub dzPal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Handler
    DesignerToolMouse Me, 2, Button, X, Y
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzPal_MouseMove"
End Sub

Private Sub dzPal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Handler
    DesignerToolMouse Me, 3, Button, X, Y
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.dzPal_MouseUp"
End Sub

' The selected element and the layout.
Private Sub btnDzEdit_Click()
    On Error GoTo Err_Handler
    DesignerEdit Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzEdit_Click"
End Sub

Private Sub btnDzDelete_Click()
    On Error GoTo Err_Handler
    DesignerDelete Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzDelete_Click"
End Sub

Private Sub btnDzArrange_Click()
    On Error GoTo Err_Handler
    DesignerArrange Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzArrange_Click"
End Sub

Private Sub btnDzUndo_Click()
    On Error GoTo Err_Handler
    DesignerUndo Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzUndo_Click"
End Sub

Private Sub btnDzRedo_Click()
    On Error GoTo Err_Handler
    DesignerRedo Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzRedo_Click"
End Sub

Private Sub cboDzOpen_AfterUpdate()
    On Error GoTo Err_Handler
    DesignerOpen Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.cboDzOpen_AfterUpdate"
End Sub

Private Sub txtDzName_AfterUpdate()
    On Error GoTo Err_Handler
    DesignerRenamed
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.txtDzName_AfterUpdate"
End Sub

Private Sub btnDzForm_Click()
    On Error GoTo Err_Handler
    DesignerForm Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzForm_Click"
End Sub

Private Sub btnDzTabs_Click()
    On Error GoTo Err_Handler
    DesignerTabs Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzTabs_Click"
End Sub

' New, Import and Export.
Private Sub btnDzMore_Click()
    On Error GoTo Err_Handler
    DesignerMore Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzMore_Click"
End Sub

Private Sub btnDzSave_Click()
    On Error GoTo Err_Handler
    DesignerSave Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzSave_Click"
End Sub

Private Sub btnDzGenerate_Click()
    On Error GoTo Err_Handler
    DesignerGenerate Me
    Exit Sub
Err_Handler:
    HandleError "pgDesigner.btnDzGenerate_Click"
End Sub

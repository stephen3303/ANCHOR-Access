Option Compare Database
Option Explicit

' Access UI kit - keyboard shortcuts that work the same in every app:
'   Ctrl+S  save        Ctrl+N  new         Ctrl+F  find        F1  help
'   Ctrl+K  the command palette (KitPalette; nothing without dlgPalette)
' BuildShortcutKeys writes them as the AutoKeys macro (it replaces any the
' database has, BuildHelpKey's included, and keeps F1 for KitHelp). Each key
' runs KitShortcut with its action ("save", "new", "find"), which asks the
' forms the user is in, innermost first (a page in frmMain's sfPage before
' frmMain, a subform before its parent), whether one of them does it: a form
' takes part by having, in its code-behind,
'   Public Function Shortcut(ByVal action As String) As Boolean
'       Select Case action
'           Case "new": NewItem: Shortcut = True
'       End Select
'   End Function
' returning True when it did the action. When no form does:
'   save  saves the changed record of the innermost bound form that has one,
'         with KitSaveRecord (lock retries and the kit's write-conflict prompt)
'   find  puts the cursor in the innermost form's search box (txtSearch)
'   new   does nothing
' Esc isn't an AutoKeys key; dialogs close on Esc through their Cancel
' button. Show the keys in tooltips: Save's ControlTipText "Save (Ctrl+S)".

Public Const SHORTCUT_SAVE As String = "save"
Public Const SHORTCUT_NEW As String = "new"
Public Const SHORTCUT_FIND As String = "find"

' Build step: write the AutoKeys macro for the shortcuts. withHelp:=False
' leaves F1 to Access (an app without the kit's help).
Public Sub BuildShortcutKeys(Optional ByVal withHelp As Boolean = True)
    Dim text As String
    text = "Version =196611" & vbCrLf & "ColumnsShown =1" & vbCrLf
    ' Single quotes: double quotes inside a macro argument don't survive LoadFromText.
    text = text & KeyMacro("^S", "KitShortcut('" & SHORTCUT_SAVE & "')")
    text = text & KeyMacro("^N", "KitShortcut('" & SHORTCUT_NEW & "')")
    text = text & KeyMacro("^F", "KitShortcut('" & SHORTCUT_FIND & "')")
    text = text & KeyMacro("^K", "KitPalette()")
    If withHelp Then text = text & KeyMacro("{F1}", "KitHelp()")
    WriteAutoKeys text
End Sub

' Do a shortcut's action in the forms the user is in (see the top of the
' module). Returns True when something did it. AutoKeys calls it.
Public Function KitShortcut(ByVal action As String) As Boolean
    Dim chain As Collection, i As Long
    On Error GoTo Fail
    Set chain = FocusChain()
    For i = 1 To chain.Count
        If AskShortcut(chain(i), action) Then KitShortcut = True: Exit Function
    Next
    Select Case action
        Case SHORTCUT_SAVE
            For i = 1 To chain.Count
                If IsDirty(chain(i)) Then
                    KitSaveRecord chain(i)          ' it tells the user when the save can't happen
                    KitShortcut = True
                    Exit Function
                End If
            Next
        Case SHORTCUT_FIND
            For i = 1 To chain.Count
                If HasSearch(chain(i)) Then
                    chain(i).Controls("txtSearch").SetFocus
                    KitShortcut = True
                    Exit Function
                End If
            Next
    End Select
    Exit Function
Fail:
    Select Case Err.Number
        Case 2101, 2115, 2169, 3020     ' a save the form refused: it has said why
        Case Else: HandleError "modKitKeys.KitShortcut", action
    End Select
End Function

' The forms the user is in, innermost first: the active form, then down
' through each subform control that has the focus, reversed.
Public Function FocusChain() As Collection
    Dim outer As Collection, frm As Object, c As Object, i As Long
    Set outer = New Collection
    Set FocusChain = New Collection
    On Error Resume Next
    Set frm = Screen.ActiveForm
    On Error GoTo 0
    Do While Not frm Is Nothing
        outer.Add frm
        Set c = Nothing
        On Error Resume Next
        Set c = frm.ActiveControl
        On Error GoTo 0
        Set frm = Nothing
        If Not c Is Nothing Then
            If c.ControlType = acSubform Then
                On Error Resume Next
                Set frm = c.Form
                On Error GoTo 0
            End If
        End If
    Loop
    For i = outer.Count To 1 Step -1
        FocusChain.Add outer(i)
    Next
End Function

' Replace the AutoKeys macro with text (the SaveAsText form of a macro).
Public Sub WriteAutoKeys(ByVal text As String)
    Dim path As String, f As Integer
    path = Environ$("TEMP") & "\kit_autokeys_" & Format$(Now(), "yyyymmddhhnnss") & ".txt"
    f = FreeFile
    Open path For Output As #f
    Print #f, text;
    Close #f
    On Error Resume Next
    DoCmd.DeleteObject acMacro, "AutoKeys"
    On Error GoTo Fail
    Application.LoadFromText acMacro, "AutoKeys", path
    Kill path
    Exit Sub
Fail:
    On Error Resume Next
    Kill path
    On Error GoTo 0
    RaiseAgain "modKitKeys.WriteAutoKeys"
End Sub

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

Private Function KeyMacro(ByVal key As String, ByVal fn As String) As String
    KeyMacro = "Begin" & vbCrLf & "    MacroName =""" & key & """" & vbCrLf & "    Action =""RunCode""" & vbCrLf & _
        "    Argument =""" & Replace(fn, """", """""") & """" & vbCrLf & "End" & vbCrLf
End Function

' frm's Shortcut(action), or False when it has none (438, or 2465 for a
' form whose module lacks it or that has no module).
Private Function AskShortcut(frm As Object, ByVal action As String) As Boolean
    On Error GoTo NoHandler
    AskShortcut = frm.Shortcut(action)
    Exit Function
NoHandler:
    Select Case Err.Number
        Case 438, 2465: Err.Clear
        Case Else: RaiseAgain "modKitKeys.AskShortcut", frm.Name & " " & action
    End Select
End Function

Private Function IsDirty(frm As Object) As Boolean
    On Error Resume Next                        ' an unbound form has no Dirty
    IsDirty = frm.Dirty
    Err.Clear                                   ' else RunCode reports the key's action as failed
End Function

Private Function HasSearch(frm As Object) As Boolean
    Dim c As Control
    On Error Resume Next
    Set c = frm.Controls("txtSearch")
    If Not c Is Nothing Then HasSearch = c.Visible And c.Enabled
    Err.Clear                                   ' a form without one: not a failure
End Function

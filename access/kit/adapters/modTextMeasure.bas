Attribute VB_Name = "modTextMeasure"
Option Compare Database
Option Explicit

' Optional Windows GDI MeasureText hook. Import only when the app has no hook.
' Segoe UI point sizes are rendered at a canonical 96 DPI and returned in twips.

#If VBA7 Then
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal dc As LongPtr) As Long
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal height As Long, ByVal width As Long, ByVal escapement As Long, ByVal orientation As Long, ByVal weight As Long, ByVal italic As Long, ByVal underline As Long, ByVal strikeout As Long, ByVal charset As Long, ByVal outPrecision As Long, ByVal clipPrecision As Long, ByVal quality As Long, ByVal pitch As Long, ByVal face As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal dc As LongPtr, ByVal obj As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal obj As LongPtr) As Long
Private Declare PtrSafe Function GetTextExtentPoint32W Lib "gdi32" (ByVal dc As LongPtr, ByVal text As LongPtr, ByVal count As Long, ByRef extent As TextExtent) As Long
#Else
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal dc As Long) As Long
Private Declare Function CreateFontW Lib "gdi32" (ByVal height As Long, ByVal width As Long, ByVal escapement As Long, ByVal orientation As Long, ByVal weight As Long, ByVal italic As Long, ByVal underline As Long, ByVal strikeout As Long, ByVal charset As Long, ByVal outPrecision As Long, ByVal clipPrecision As Long, ByVal quality As Long, ByVal pitch As Long, ByVal face As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal dc As Long, ByVal obj As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal obj As Long) As Long
Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal dc As Long, ByVal text As Long, ByVal count As Long, ByRef extent As TextExtent) As Long
#End If
Private Type TextExtent
    x As Long
    y As Long
End Type

Public Function MeasureText(ByVal text As String, ByVal size As Long, ByVal weight As Long) As Long
#If VBA7 Then
    Dim dc As LongPtr, font As LongPtr, previous As LongPtr
#Else
    Dim dc As Long, font As Long, previous As Long
#End If
    Dim extent As TextExtent, face As String
    If Len(text) = 0 Then Exit Function
    face = "Segoe UI": dc = GetDC(0)
    If dc = 0 Then Err.Raise 5, , "Text measurement unavailable"
    font = CreateFontW(-CLng(size * 96 / 72), 0, 0, 0, weight, 0, 0, 0, 1, 0, 0, 4, 0, StrPtr(face))
    If font = 0 Then ReleaseDC 0, dc: Err.Raise 5, , "Text font unavailable"
    previous = SelectObject(dc, font)
    Dim success As Long
    success = GetTextExtentPoint32W(dc, StrPtr(text), Len(text), extent)
    SelectObject dc, previous: DeleteObject font: ReleaseDC 0, dc
    If success = 0 Then Err.Raise 5, , "Text width unavailable"
    MeasureText = extent.x * 15
End Function

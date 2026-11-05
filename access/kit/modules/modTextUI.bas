Option Compare Database
Option Explicit

' Standalone text presentation helpers. The app supplies MeasureText(text, size,
' weight), returning Segoe UI text width in twips. No table or palette dependency.

' Segoe UI, matching the kit controls. Width excludes the control's margins.
' Presentation only: keep the original bound value available in the editor.
Public Function FitSingleLineText(ByVal value As Variant, ByVal width As Long, Optional ByVal size As Long = 9, Optional ByVal weight As Long = 400) As String
    Dim text As String, suffix As String, lo As Long, hi As Long, middle As Long
    text = Replace(Replace(Replace(Nz(value, ""), vbCrLf, " "), vbCr, " "), vbLf, " ")
    If width <= 0 Or Len(text) = 0 Then Exit Function
    If MeasureText(text, size, weight) <= width Then FitSingleLineText = text: Exit Function
    suffix = ChrW(&H2026)
    If MeasureText(suffix, size, weight) > width Then Exit Function
    hi = Len(text)
    Do While lo < hi
        middle = (lo + hi + 1) \ 2
        If MeasureText(Left$(text, middle) & suffix, size, weight) <= width Then lo = middle Else hi = middle - 1
    Loop
    FitSingleLineText = Left$(text, lo) & suffix
End Function

' Pass the full caption and preferred size on every layout call so widening
' restores text and font. Native tips are limited to 255 characters. They are
' suitable for headers, not per-row values in a continuous bound table.
Public Sub FitHeaderText(ByVal heading As Control, ByVal fullCaption As String, Optional ByVal preferredSize As Long = 9, Optional ByVal minimumSize As Long = 8)
    Dim size As Long, width As Long, weight As Long, text As String
    If minimumSize < 1 Or preferredSize < minimumSize Then Err.Raise 5, , "Invalid header font range"
    width = heading.Width - heading.LeftMargin - heading.RightMargin
    weight = IIf(heading.FontBold, 700, 400): size = preferredSize
    text = Replace(Replace(Replace(fullCaption, vbCrLf, " "), vbCr, " "), vbLf, " ")
    Do While size > minimumSize
        If MeasureText(text, size, weight) <= width Then Exit Do
        size = size - 1
    Loop
    heading.FontSize = size
    heading.Caption = FitSingleLineText(text, width, size, weight)
    heading.ControlTipText = Left$(fullCaption, 255)
End Sub


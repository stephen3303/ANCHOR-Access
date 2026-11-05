Option Compare Database
Option Explicit

' Access UI kit - RFC 4180 CSV reading and writing (UTF-8).

' Whole file as text, read as UTF-8 with any BOM removed.
Public Function ReadTextFile(ByVal path As String) As String
    On Error GoTo Fail
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadTextFile = st.ReadText
    st.Close
    If Len(ReadTextFile) > 0 Then
        If AscW(Left$(ReadTextFile, 1)) = &HFEFF Then ReadTextFile = Mid$(ReadTextFile, 2)
    End If
    Exit Function
Fail:
    RaiseAgain "modCsv.ReadTextFile", "file " & path
End Function

' Write text as UTF-8 (ADODB.Stream adds a BOM).
Public Sub WriteTextFile(ByVal path As String, ByVal text As String)
    On Error GoTo Fail
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2 ' adSaveCreateOverWrite
    st.Close
    Exit Sub
Fail:
    RaiseAgain "modCsv.WriteTextFile", "file " & path
End Sub

' Parse CSV text into a Collection of rows; each row is a 0-based String array.
Public Function ParseCsv(ByVal text As String) As Collection
    Dim rows As New Collection, fields As Collection
    Dim i As Long, n As Long, ch As String, buf As String, inQ As Boolean
    Dim lineHasData As Boolean
    n = Len(text)
    Set fields = New Collection
    i = 1
    Do While i <= n
        ch = Mid$(text, i, 1)
        If inQ Then
            If ch = """" Then
                If i < n And Mid$(text, i + 1, 1) = """" Then
                    buf = buf & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                buf = buf & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQ = True
                    lineHasData = True
                Case ","
                    fields.Add buf
                    buf = ""
                    lineHasData = True
                Case vbCr
                    ' ignore; handled by LF
                Case vbLf
                    fields.Add buf
                    If lineHasData Or Len(buf) > 0 Then rows.Add CollToArray(fields)
                    Set fields = New Collection
                    buf = ""
                    lineHasData = False
                Case Else
                    buf = buf & ch
                    lineHasData = True
            End Select
        End If
        i = i + 1
    Loop
    If lineHasData Or Len(buf) > 0 Then
        fields.Add buf
        rows.Add CollToArray(fields)
    End If
    Set ParseCsv = rows
End Function

Private Function CollToArray(c As Collection) As String()
    Dim a() As String, i As Long
    ReDim a(0 To c.Count - 1)
    For i = 1 To c.Count
        a(i - 1) = c(i)
    Next
    CollToArray = a
End Function

' Read a CSV with a header row into a Collection of Scripting.Dictionary rows
' keyed by header name.
Public Function ReadCsvDicts(ByVal path As String) As Collection
    Dim raw As Collection, out As New Collection, hdr() As String
    Dim r As Long, c As Long, d As Object, row() As String
    Set raw = ParseCsv(ReadTextFile(path))
    If raw.Count = 0 Then
        Set ReadCsvDicts = out
        Exit Function
    End If
    hdr = raw(1)
    For r = 2 To raw.Count
        row = raw(r)
        Set d = CreateObject("Scripting.Dictionary")
        d.CompareMode = 1
        For c = 0 To UBound(hdr)
            If c <= UBound(row) Then d(Trim$(hdr(c))) = row(c) Else d(Trim$(hdr(c))) = ""
        Next
        out.Add d
    Next
    Set ReadCsvDicts = out
End Function

' One value formatted for CSV: booleans as 1/0, ISO dates, quoted when needed.
Public Function CsvField(ByVal v As Variant) As String
    Dim s As String
    If IsNull(v) Then
        CsvField = ""
        Exit Function
    End If
    Select Case VarType(v)
        Case vbBoolean
            s = IIf(v, "1", "0")
        Case vbDate
            If Int(v) = v Then s = Format$(v, "yyyy-mm-dd") Else s = Format$(v, "yyyy-mm-dd hh:nn:ss")
        Case vbDouble, vbSingle, vbCurrency, vbDecimal
            s = Trim$(Str$(v))
        Case Else
            s = CStr(v)
    End Select
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then
        s = """" & Replace(s, """", """""") & """"
    End If
    CsvField = s
End Function

' One CSV line (with CRLF) from any number of values.
Public Function CsvLine(ParamArray values() As Variant) As String
    Dim i As Long, parts() As String
    ReDim parts(0 To UBound(values))
    For i = 0 To UBound(values)
        parts(i) = CsvField(values(i))
    Next
    CsvLine = Join(parts, ",") & vbCrLf
End Function

' Write the result of a query (header + rows) to a CSV file. Returns row count.
Public Function ExportSqlToCsv(ByVal sql As String, ByVal path As String) As Long
    On Error GoTo Fail
    Dim rs As DAO.Recordset, i As Long, sb As String, parts() As String, n As Long
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    ReDim parts(0 To rs.Fields.Count - 1)
    For i = 0 To rs.Fields.Count - 1
        parts(i) = CsvField(rs.Fields(i).Name)
    Next
    sb = Join(parts, ",") & vbCrLf
    Do While Not rs.EOF
        For i = 0 To rs.Fields.Count - 1
            parts(i) = CsvField(rs.Fields(i).Value)
        Next
        sb = sb & Join(parts, ",") & vbCrLf
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    WriteTextFile path, sb
    ExportSqlToCsv = n
    Exit Function
Fail:
    RaiseAgain "modCsv.ExportSqlToCsv", "file " & path
End Function

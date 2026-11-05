Option Compare Database
Option Explicit

' Access UI kit - hashing through Windows CNG (bcrypt.dll).
' Used for password storage and export-package SHA-256 checksums.

Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
    ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, _
    ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptHash Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, _
    ByVal pbInput As LongPtr, ByVal cbInput As Long, ByVal pbOutput As LongPtr, ByVal cbOutput As Long) As Long
Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, ByVal pbBuffer As LongPtr, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long

Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = 2
Private Const HASH_ROUNDS As Long = 10000

Private m_hAlg As LongPtr

' SHA-256 provider handle, opened once and reused.
Private Function AlgHandle() As LongPtr
    Dim rc As Long, alg As String
    If m_hAlg = 0 Then
        alg = "SHA256"
        rc = BCryptOpenAlgorithmProvider(m_hAlg, StrPtr(alg), 0, 0)
        If rc <> 0 Then Err.Raise vbObjectError + 100, , "BCryptOpenAlgorithmProvider failed: " & Hex$(rc)
    End If
    AlgHandle = m_hAlg
End Function

' SHA-256 of the first dataLen bytes of data.
Private Function Sha256Bytes(data() As Byte, ByVal dataLen As Long) As Byte()
    Dim out(0 To 31) As Byte, rc As Long, dummy As Byte
    If dataLen = 0 Then
        rc = BCryptHash(AlgHandle(), 0, 0, VarPtr(dummy), 0, VarPtr(out(0)), 32)
    Else
        rc = BCryptHash(AlgHandle(), 0, 0, VarPtr(data(LBound(data))), dataLen, VarPtr(out(0)), 32)
    End If
    If rc <> 0 Then Err.Raise vbObjectError + 101, , "BCryptHash failed: " & Hex$(rc)
    Sha256Bytes = out
End Function

' Lower-case hex string of a byte array.
Private Function ToHex(b() As Byte) As String
    Dim i As Long, s As String, k As Long
    s = Space$(2 * (UBound(b) - LBound(b) + 1))
    k = 1
    For i = LBound(b) To UBound(b)
        Mid$(s, k, 2) = Mid$("0123456789abcdef", (b(i) \ 16) + 1, 1) & Mid$("0123456789abcdef", (b(i) And 15) + 1, 1)
        k = k + 2
    Next
    ToHex = s
End Function

' UTF-8 bytes of a string (without the BOM ADODB.Stream writes).
Private Function Utf8Bytes(ByVal s As String) As Byte()
    Dim st As Object, b() As Byte
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText s
    st.Position = 0: st.Type = 1: st.Position = 3   ' skip BOM
    b = st.Read
    st.Close
    Utf8Bytes = b
End Function

' SHA-256 of a string's UTF-8 bytes, as hex.
Public Function Sha256Hex(ByVal s As String) As String
    Dim b() As Byte
    If Len(s) = 0 Then
        ReDim b(0 To 0)
        Sha256Hex = ToHex(Sha256Bytes(b, 0))
    Else
        b = Utf8Bytes(s)
        Sha256Hex = ToHex(Sha256Bytes(b, UBound(b) - LBound(b) + 1))
    End If
End Function

' Hash of an ASCII-only string (hex digests, salts) without the stream round trip.
Private Function Sha256Ascii(ByVal s As String) As String
    Dim b() As Byte
    b = StrConv(s, vbFromUnicode)
    Sha256Ascii = ToHex(Sha256Bytes(b, UBound(b) + 1))
End Function

' SHA-256 of a whole file, as hex (export package checksums).
Public Function Sha256File(ByVal path As String) As String
    Dim f As Integer, b() As Byte, n As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n = 0 Then
        ReDim b(0 To 0)
    Else
        ReDim b(0 To n - 1)
        Get #f, , b
    End If
    Close #f
    Sha256File = ToHex(Sha256Bytes(b, n))
End Function

' nBytes of cryptographically random data, as hex (password salts).
Public Function RandomHex(ByVal nBytes As Long) As String
    Dim b() As Byte
    ReDim b(0 To nBytes - 1)
    BCryptGenRandom 0, VarPtr(b(0)), nBytes, BCRYPT_USE_SYSTEM_PREFERRED_RNG
    RandomHex = ToHex(b)
End Function

' Stored format: sha256$<rounds>$<salt>$<hash>
Public Function HashPassword(ByVal password As String, Optional ByVal salt As String = "", _
                             Optional ByVal rounds As Long = HASH_ROUNDS) As String
    Dim h As String, i As Long
    If Len(salt) = 0 Then salt = RandomHex(16)
    h = Sha256Hex(salt & ":" & password)
    For i = 2 To rounds
        h = Sha256Ascii(h & salt)
    Next
    HashPassword = "sha256$" & rounds & "$" & salt & "$" & h
End Function

' Re-hash with the stored salt and round count and compare.
Public Function VerifyPassword(ByVal password As String, ByVal stored As String) As Boolean
    Dim parts() As String
    If Len(stored) = 0 Then Exit Function
    parts = Split(stored, "$")
    If UBound(parts) <> 3 Then Exit Function
    VerifyPassword = (HashPassword(password, parts(2), CLng(parts(1))) = stored)
End Function

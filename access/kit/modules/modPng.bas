Option Compare Database
Option Explicit

' Access UI kit - rounded images. Access controls only draw square corners,
' so pills, badges, card corners, progress bars and toggle switches are image
' controls showing anti-aliased PNGs drawn here. Each image is written once
' per color and size to %LOCALAPPDATA%\<APP_NAME>\pills and reused. The
' image controls get them through ControlSource expressions such as
' =PillPng(<color>,w,h), so ThemeForm can repaint the colors written in them.

Private Const SCALE_UP As Long = 2          ' draw at 2x for crisp edges when zoomed
Private m_crc(255) As Long
Private m_crcReady As Boolean
Private m_cache As Collection
Private m_dir As String

' Path of a capsule image of the given Access color, sized for a control of
' wPx x hPx screen pixels. Null color (no matching condition) -> no image.
Public Function PillPng(ByVal clr As Variant, ByVal wPx As Long, ByVal hPx As Long) As Variant
    Dim key As String, path As String, draw As Boolean
    PillPng = Null
    If IsNull(clr) Or IsEmpty(clr) Then Exit Function
    If clr < 0 Or wPx < 2 Or hPx < 2 Then Exit Function
    On Error GoTo Fail
    key = "p_" & Hex6(clr) & "_" & wPx & "x" & hPx
    path = ImagePath(key, draw)
    If draw Then
        WriteCapsule path, CLng(clr), -1, 1, wPx * SCALE_UP, hPx * SCALE_UP
        Remember key, path
    End If
    PillPng = path
    Exit Function
Fail:
    PillPng = Null
End Function

' Progress capsule for a control of wPx x hPx: a track in color track, filled
' from the left in color fill to percent (0-100; Null or text counts as 0).
Public Function ProgressPng(ByVal percent As Variant, ByVal fill As Long, ByVal track As Long, _
                            ByVal wPx As Long, ByVal hPx As Long) As Variant
    Dim key As String, path As String, draw As Boolean, v As Long
    ProgressPng = Null
    If wPx < 2 Or hPx < 2 Then Exit Function
    On Error GoTo Fail
    If IsNumeric(percent) Then v = CLng(percent)
    If v < 0 Then v = 0
    If v > 100 Then v = 100
    key = "g_" & Hex6(fill) & "_" & Hex6(track) & "_" & v & "_" & wPx & "x" & hPx
    path = ImagePath(key, draw)
    If draw Then
        WriteCapsule path, fill, track, v / 100, wPx * SCALE_UP, hPx * SCALE_UP
        Remember key, path
    End If
    ProgressPng = path
    Exit Function
Fail:
    ProgressPng = Null
End Function

' On/off switch track for a control of wPx x hPx (modUIKit.Toggle): a capsule
' in onClr with a round knob at the right when isOn is true, in offClr with
' the knob at the left when it is false or Null. The knob is knobClr, 2px in
' from the track's edge.
Public Function TogglePng(ByVal isOn As Variant, ByVal onClr As Long, ByVal offClr As Long, ByVal knobClr As Long, _
                          ByVal wPx As Long, ByVal hPx As Long) As Variant
    Dim key As String, path As String, draw As Boolean, v As Boolean
    TogglePng = Null
    If wPx < 2 Or hPx < 2 Then Exit Function
    On Error GoTo Fail
    If Not IsNull(isOn) Then v = CBool(isOn)
    key = "t_" & IIf(v, "1", "0") & "_" & Hex6(IIf(v, onClr, offClr)) & "_" & Hex6(knobClr) & "_" & wPx & "x" & hPx
    path = ImagePath(key, draw)
    If draw Then
        WriteSwitch path, IIf(v, onClr, offClr), knobClr, v, wPx * SCALE_UP, hPx * SCALE_UP
        Remember key, path
    End If
    TogglePng = path
    Exit Function
Fail:
    TogglePng = Null
End Function

' Slider of wPx x hPx: a 6px track in color track across it, filled in color
' fill from the left to percent (0-100, Null or text is 0), and a round knob
' of diameter hPx in color knob with a 2px ring in color fill at that point.
Public Function SliderPng(ByVal percent As Variant, ByVal fill As Long, ByVal track As Long, ByVal knob As Long, _
                          ByVal wPx As Long, ByVal hPx As Long) As Variant
    Dim key As String, path As String, draw As Boolean, v As Double
    SliderPng = Null
    If wPx <= hPx Or hPx < 8 Then Exit Function
    On Error GoTo Fail
    If IsNumeric(percent) Then v = CDbl(percent)
    If v < 0 Then v = 0
    If v > 100 Then v = 100
    v = Int(v * 10 + 0.5) / 10
    key = "s_" & Hex6(fill) & "_" & Hex6(track) & "_" & Hex6(knob) & "_" & Replace(CStr(v), ",", ".") & "_" & wPx & "x" & hPx
    path = ImagePath(key, draw)
    If draw Then
        WriteSlider path, fill, track, knob, v / 100, wPx * SCALE_UP, hPx * SCALE_UP
        Remember key, path
    End If
    SliderPng = path
    Exit Function
Fail:
    SliderPng = Null
End Function

Private Sub WriteSlider(ByVal path As String, ByVal fill As Long, ByVal track As Long, ByVal knob As Long, _
                        ByVal fraction As Double, ByVal w As Long, ByVal h As Long)
    Dim raw() As Byte, x As Long, y As Long, i As Long, th As Double, kx As Double, r As Double, bw As Double
    Dim ta As Double, tc As Long, d As Double, outer As Double, inner As Double, kc As Long, a As Double, c As Long
    th = 6 * SCALE_UP
    r = h / 2
    bw = 2 * SCALE_UP
    kx = r + (w - h) * fraction
    ReDim raw(0 To h * (w * 4 + 1) - 1)
    i = 0
    For y = 0 To h - 1
        raw(i) = 0: i = i + 1
        For x = 0 To w - 1
            ' the track: a capsule th tall from the first knob center to the last
            ta = CapsuleCover(x + 0.5 - r + th / 2, y + 0.5 - (h - th) / 2, w - h + th, th, 0)
            tc = IIf(x + 0.5 <= kx, fill, track)
            d = Sqr((x + 0.5 - kx) ^ 2 + (y + 0.5 - r) ^ 2)
            outer = r - d + 0.5
            If outer < 0 Then outer = 0
            If outer > 1 Then outer = 1
            inner = r - bw - d + 0.5
            If inner < 0 Then inner = 0
            If inner > 1 Then inner = 1
            kc = Blend(fill, knob, inner)
            If outer > 0 Then
                c = Blend(tc, kc, IIf(ta > 0, outer, 1))
                a = outer + ta * (1 - outer)
            Else
                c = tc: a = ta
            End If
            PutPixel raw, i, c, a
            i = i + 4
        Next
    Next
    WritePng path, w, h, raw
End Sub

' Focus ring for a capsule control: the outline of a wPx x hPx capsule, 2px
' wide, in clr, transparent inside (modUIKit.Toggle draws it 4px larger than
' the track on every side, which leaves a 2px gap).
Public Function RingPng(ByVal clr As Long, ByVal wPx As Long, ByVal hPx As Long) As Variant
    Dim key As String, path As String, draw As Boolean
    RingPng = Null
    If wPx < 6 Or hPx < 6 Then Exit Function
    On Error GoTo Fail
    key = "r_" & Hex6(clr) & "_" & wPx & "x" & hPx
    path = ImagePath(key, draw)
    If draw Then
        WriteRing path, clr, wPx * SCALE_UP, hPx * SCALE_UP, 2 * SCALE_UP
        Remember key, path
    End If
    RingPng = path
    Exit Function
Fail:
    RingPng = Null
End Function

' One rounded corner of a card (modUIKit.Card): a square image of side rPx
' laid over the corner of a square box. Inside the arc it is fill, on the arc
' a 1px border, outside it back (what the card sits on). corner: 1 top-left,
' 2 top-right, 3 bottom-right, 4 bottom-left.
Public Function CardPng(ByVal fill As Long, ByVal border As Long, ByVal back As Long, ByVal rPx As Long, _
                        ByVal corner As Long) As Variant
    Dim key As String, path As String, draw As Boolean
    CardPng = Null
    If rPx < 2 Or corner < 1 Or corner > 4 Then Exit Function
    On Error GoTo Fail
    key = "c_" & Hex6(fill) & "_" & Hex6(border) & "_" & Hex6(back) & "_" & rPx & "_" & corner
    path = ImagePath(key, draw)
    If draw Then
        WriteCorner path, fill, border, back, rPx * SCALE_UP, SCALE_UP, corner
        Remember key, path
    End If
    CardPng = path
    Exit Function
Fail:
    CardPng = Null
End Function

' Six hex digits of a color, for file names.
Private Function Hex6(ByVal clr As Long) As String
    Hex6 = Right$("00000" & Hex$(clr), 6)
End Function

' Path for image key. draw is set when the file doesn't exist yet: the caller
' writes it, then calls Remember.
Private Function ImagePath(ByVal key As String, draw As Boolean) As String
    Dim path As String
    If m_cache Is Nothing Then Set m_cache = New Collection
    On Error Resume Next
    path = m_cache(key)
    On Error GoTo 0
    If Len(path) = 0 Then
        path = PillDir() & "\" & key & ".png"
        draw = (Len(Dir$(path)) = 0)
        If Not draw Then m_cache.Add path, key
    End If
    ImagePath = path
End Function

Private Sub Remember(ByVal key As String, ByVal path As String)
    m_cache.Add path, key
End Sub

' %LOCALAPPDATA%\<APP_NAME>\pills (TEMP if LOCALAPPDATA is missing), created on first use.
Private Function PillDir() As String
    Dim base As String
    If Len(m_dir) = 0 Then
        base = Environ$("LOCALAPPDATA")
        If Len(base) = 0 Then base = Environ$("TEMP")
        m_dir = base & "\" & APP_NAME
        If Len(Dir$(m_dir, vbDirectory)) = 0 Then MkDir m_dir
        m_dir = m_dir & "\pills"
        If Len(Dir$(m_dir, vbDirectory)) = 0 Then MkDir m_dir
    End If
    PillDir = m_dir
End Function

' Capsule (fully rounded ends) in color clr on a transparent background. With
' a track color (track >= 0) only the left fraction is clr and the rest is
' track, as in a progress bar. capsuleWidth (> 0) draws a narrower capsule
' centered on the w-wide canvas, so pills in one column share a width.
Private Sub WriteCapsule(ByVal path As String, ByVal clr As Long, ByVal track As Long, ByVal fraction As Double, _
                         ByVal w As Long, ByVal h As Long, Optional ByVal capsuleWidth As Double = 0)
    Dim raw() As Byte, x As Long, y As Long, i As Long, edge As Double, t As Double, inset As Double
    Dim r As Double, cx As Double, cy As Double, px As Double, py As Double, d As Double, a As Double
    If capsuleWidth <= 0 Or capsuleWidth > w Then capsuleWidth = w
    inset = (w - capsuleWidth) / 2
    r = h / 2
    If r > capsuleWidth / 2 Then r = capsuleWidth / 2
    cy = h / 2
    edge = capsuleWidth * fraction
    If track < 0 Then track = clr: edge = capsuleWidth
    ReDim raw(0 To h * (w * 4 + 1) - 1)
    i = 0
    For y = 0 To h - 1
        raw(i) = 0: i = i + 1                         ' filter: none
        py = y + 0.5
        For x = 0 To w - 1
            px = x + 0.5 - inset
            ' distance from the pixel center to the capsule's center segment
            If px < r Then
                cx = r
            ElseIf px > capsuleWidth - r Then
                cx = capsuleWidth - r
            Else
                cx = px
            End If
            d = Sqr((px - cx) ^ 2 + (py - cy) ^ 2)
            a = r - d + 0.5
            If a < 0 Then a = 0
            If a > 1 Then a = 1
            ' share of the pixel left of the fill's edge
            t = edge - px + 0.5
            If t < 0 Then t = 0
            If t > 1 Then t = 1
            PutPixel raw, i, Blend(track, clr, t), a
            i = i + 4
        Next
    Next
    WritePng path, w, h, raw
End Sub

' Card corner of side rr pixels with a border bw pixels wide (see CardPng).
Private Sub WriteCorner(ByVal path As String, ByVal fill As Long, ByVal border As Long, ByVal back As Long, _
                        ByVal rr As Long, ByVal bw As Long, ByVal corner As Long)
    Dim raw() As Byte, x As Long, y As Long, i As Long, ox As Double, oy As Double
    Dim d As Double, a As Double, t As Double
    ' center of the arc's circle, in image coordinates
    ox = IIf(corner = 1 Or corner = 4, rr, 0)
    oy = IIf(corner = 1 Or corner = 2, rr, 0)
    ReDim raw(0 To rr * (rr * 4 + 1) - 1)
    i = 0
    For y = 0 To rr - 1
        raw(i) = 0: i = i + 1
        For x = 0 To rr - 1
            ' distance outside the arc (negative inside)
            d = Sqr((x + 0.5 - ox) ^ 2 + (y + 0.5 - oy) ^ 2) - rr
            a = 0.5 - d                               ' share inside the card
            If a < 0 Then a = 0
            If a > 1 Then a = 1
            t = d + bw + 0.5                          ' share on the border
            If t < 0 Then t = 0
            If t > 1 Then t = 1
            PutPixel raw, i, Blend(back, Blend(fill, border, t), a), 1
            i = i + 4
        Next
    Next
    WritePng path, rr, rr, raw
End Sub

' Share (0-1) of the pixel centered at (ptX, ptY) that lies inside a capsule
' filling a w x h canvas less inset on every side.
Private Function CapsuleCover(ByVal ptX As Double, ByVal ptY As Double, ByVal w As Double, ByVal h As Double, _
                              ByVal inset As Double) As Double
    Dim r As Double, cx As Double, d As Double, a As Double
    r = (h - 2 * inset) / 2
    If r <= 0 Or w - 2 * inset <= 0 Then Exit Function
    If r > (w - 2 * inset) / 2 Then r = (w - 2 * inset) / 2
    ' nearest point on the capsule's center segment
    cx = ptX
    If cx < inset + r Then cx = inset + r
    If cx > w - inset - r Then cx = w - inset - r
    d = Sqr((ptX - cx) ^ 2 + (ptY - h / 2) ^ 2)
    a = r - d + 0.5
    If a < 0 Then a = 0
    If a > 1 Then a = 1
    CapsuleCover = a
End Function

' Switch track of w x h pixels in color track with a knob in color knob
' (see TogglePng). The knob's diameter is the track's height less a 2px
' margin (2 * SCALE_UP drawn pixels) top and bottom.
Private Sub WriteSwitch(ByVal path As String, ByVal track As Long, ByVal knob As Long, ByVal isOn As Boolean, _
                        ByVal w As Long, ByVal h As Long)
    Dim raw() As Byte, x As Long, y As Long, i As Long, kr As Double, kx As Double
    Dim d As Double, k As Double
    kr = h / 2 - 2 * SCALE_UP
    kx = IIf(isOn, w - h / 2, h / 2)
    ReDim raw(0 To h * (w * 4 + 1) - 1)
    i = 0
    For y = 0 To h - 1
        raw(i) = 0: i = i + 1                         ' filter: none
        For x = 0 To w - 1
            d = Sqr((x + 0.5 - kx) ^ 2 + (y + 0.5 - h / 2) ^ 2)
            k = kr - d + 0.5                          ' share of the pixel on the knob
            If k < 0 Then k = 0
            If k > 1 Then k = 1
            PutPixel raw, i, Blend(track, knob, k), CapsuleCover(x + 0.5, y + 0.5, w, h, 0)
            i = i + 4
        Next
    Next
    WritePng path, w, h, raw
End Sub

' Capsule outline of w x h pixels, bw pixels wide, in color clr (see RingPng).
Private Sub WriteRing(ByVal path As String, ByVal clr As Long, ByVal w As Long, ByVal h As Long, ByVal bw As Long)
    Dim raw() As Byte, x As Long, y As Long, i As Long, a As Double
    ReDim raw(0 To h * (w * 4 + 1) - 1)
    i = 0
    For y = 0 To h - 1
        raw(i) = 0: i = i + 1
        For x = 0 To w - 1
            a = CapsuleCover(x + 0.5, y + 0.5, w, h, 0) - CapsuleCover(x + 0.5, y + 0.5, w, h, bw)
            If a < 0 Then a = 0
            PutPixel raw, i, clr, a
            i = i + 4
        Next
    Next
    WritePng path, w, h, raw
End Sub

' Color between c0 (t = 0) and c1 (t = 1).
Private Function Blend(ByVal c0 As Long, ByVal c1 As Long, ByVal t As Double) As Long
    Dim k As Long, v As Long, s As Long
    For k = 0 To 2
        s = 256 ^ k
        v = v + CLng(((c0 \ s) And &HFF) * (1 - t) + ((c1 \ s) And &HFF) * t) * s
    Next
    Blend = v
End Function

' Write one RGBA pixel of color clr and opacity a (0-1) at raw(i).
Private Sub PutPixel(raw() As Byte, ByVal i As Long, ByVal clr As Long, ByVal a As Double)
    raw(i) = clr And &HFF
    raw(i + 1) = (clr \ &H100) And &HFF
    raw(i + 2) = (clr \ &H10000) And &HFF
    raw(i + 3) = CByte(a * 255)
End Sub

' ---------------------------------------------------------------------------
' Minimal PNG writer: RGBA8, one IDAT holding a zlib stream of stored blocks.
' ---------------------------------------------------------------------------

' Write raw (filtered RGBA scanlines) as a PNG file.
Private Sub WritePng(ByVal path As String, ByVal w As Long, ByVal h As Long, raw() As Byte)
    Dim out() As Byte, n As Long, hdr(12) As Byte, z() As Byte, fh As Integer
    ReDim out(0 To 7)
    out(0) = 137: out(1) = 80: out(2) = 78: out(3) = 71: out(4) = 13: out(5) = 10: out(6) = 26: out(7) = 10
    n = 8
    PutBE hdr, 0, w: PutBE hdr, 4, h
    hdr(8) = 8: hdr(9) = 6: hdr(10) = 0: hdr(11) = 0: hdr(12) = 0
    AddChunk out, n, "IHDR", hdr, 13
    z = Zlib(raw)
    AddChunk out, n, "IDAT", z, UBound(z) + 1
    AddChunk out, n, "IEND", hdr, 0
    ReDim Preserve out(0 To n - 1)
    If Len(Dir$(path)) > 0 Then Kill path
    fh = FreeFile
    Open path For Binary Access Write As #fh
    Put #fh, , out
    Close #fh
End Sub

' zlib stream of uncompressed (stored) blocks with an Adler-32 checksum. The
' images are tiny, so compression isn't worth the code.
Private Function Zlib(raw() As Byte) As Byte()
    Dim total As Long, blocks As Long, z() As Byte, i As Long, pos As Long, ln As Long, b As Long
    Dim s1 As Long, s2 As Long
    total = UBound(raw) + 1
    blocks = (total + 65534) \ 65535
    If blocks = 0 Then blocks = 1
    ReDim z(0 To 2 + blocks * 5 + total + 4 - 1)
    z(0) = &H78: z(1) = &H1
    pos = 2
    For b = 0 To blocks - 1
        ln = total - b * 65535
        If ln > 65535 Then ln = 65535
        z(pos) = IIf(b = blocks - 1, 1, 0)
        z(pos + 1) = ln And &HFF: z(pos + 2) = (ln \ &H100) And &HFF
        z(pos + 3) = (65535 - ln) And &HFF: z(pos + 4) = ((65535 - ln) \ &H100) And &HFF
        pos = pos + 5
        For i = 0 To ln - 1
            z(pos + i) = raw(b * 65535& + i)
        Next
        pos = pos + ln
    Next
    s1 = 1: s2 = 0
    For i = 0 To total - 1
        s1 = (s1 + raw(i)) Mod 65521
        s2 = (s2 + s1) Mod 65521
    Next
    z(pos) = (s2 \ &H100) And &HFF: z(pos + 1) = s2 And &HFF
    z(pos + 2) = (s1 \ &H100) And &HFF: z(pos + 3) = s1 And &HFF
    Zlib = z
End Function

' Append a PNG chunk (length, type, data, CRC) to out at position n.
Private Sub AddChunk(out() As Byte, n As Long, ByVal typ As String, data() As Byte, ByVal ln As Long)
    Dim i As Long, crc As Long, start As Long
    ReDim Preserve out(0 To n + 12 + ln - 1)
    PutBE out, n, ln
    start = n + 4
    For i = 1 To 4
        out(start + i - 1) = Asc(Mid$(typ, i, 1))
    Next
    For i = 0 To ln - 1
        out(start + 4 + i) = data(i)
    Next
    crc = Crc32(out, start, 4 + ln)
    PutBE out, start + 4 + ln, crc
    n = start + 4 + ln + 4
End Sub

' Write a 32-bit big-endian integer.
Private Sub PutBE(b() As Byte, ByVal at As Long, ByVal v As Long)
    b(at) = ((v And &HFF000000) \ &H1000000) And &HFF
    b(at + 1) = (v And &HFF0000) \ &H10000
    b(at + 2) = (v And &HFF00&) \ &H100
    b(at + 3) = v And &HFF
End Sub

' CRC-32 as PNG requires (table built on first use).
Private Function Crc32(b() As Byte, ByVal start As Long, ByVal ln As Long) As Long
    Dim c As Long, i As Long, k As Long
    If Not m_crcReady Then
        For i = 0 To 255
            c = i
            For k = 1 To 8
                If c And 1 Then
                    c = (((c And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor &HEDB88320
                Else
                    c = ((c And &HFFFFFFFE) \ 2) And &H7FFFFFFF
                End If
            Next
            m_crc(i) = c
        Next
        m_crcReady = True
    End If
    c = &HFFFFFFFF
    For i = start To start + ln - 1
        c = (((c And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor m_crc((c Xor b(i)) And &HFF)
    Next
    Crc32 = c Xor &HFFFFFFFF
End Function

' Measured Segoe UI width with padding; callers supply MeasureText via modWrap.
Public Function PillTextWidth(ByVal value As Variant, Optional ByVal fontSize As Long = 9, Optional ByVal padding As Long = 480) As Long
    PillTextWidth = MeasureText(Nz(value, ""), fontSize, 400) + padding
End Function
Public Function PillTextPng(ByVal clr As Variant, ByVal value As Variant, ByVal wPx As Long, ByVal hPx As Long, Optional ByVal uniformWidth As Long = 0) As Variant
    Dim fit As Long, path As String
    PillTextPng = Null
    If IsNull(clr) Or wPx < 2 Then Exit Function
    fit = PillTextWidth(value) \ 15
    If uniformWidth > 0 Then fit = uniformWidth \ 15
    If fit < hPx Then fit = hPx
    If fit > wPx Then fit = wPx
    path = PillDir() & "\text_" & Hex$(CLng(clr)) & "_" & fit & "_" & wPx & "_" & hPx & ".png"
    If Len(Dir(path)) = 0 Then WriteCapsule path, CLng(clr), -1, 1, wPx * 2, hPx * 2, fit * 2
    PillTextPng = path
End Function

' Whole rounded card background of w x h pixels, redrawn at the control's
' size (modUIKit.RoundedCard). CardPng draws the corners of a Card instead.
Public Function RoundedCardPng(ByVal w As Long, ByVal h As Long, ByVal back As Long, ByVal border As Long) As String
    Dim path As String, raw() As Byte, x As Long, y As Long, at As Long
    Dim dx As Double, dy As Double, distance As Double, radius As Double, alpha As Double, color As Long
    path = PillDir() & "\card_" & w & "_" & h & "_" & Hex$(back) & "_" & Hex$(border) & ".png"
    If Len(Dir(path)) > 0 Then RoundedCardPng = path: Exit Function
    radius = 9
    ReDim raw(0 To h * (w * 4 + 1) - 1)
    For y = 0 To h - 1
        at = at + 1
        For x = 0 To w - 1
            dx = Abs(x + 0.5 - w / 2) - (w / 2 - radius - 1)
            dy = Abs(y + 0.5 - h / 2) - (h / 2 - radius - 1)
            If dx < 0 Then dx = 0
            If dy < 0 Then dy = 0
            distance = Sqr(dx * dx + dy * dy)
            alpha = radius - distance + 0.5
            If alpha > 1 Then alpha = 1
            If alpha < 0 Then alpha = 0
            color = back
            If distance > radius - 1.3 Then color = border
            raw(at) = color And &HFF
            raw(at + 1) = (color \ 256) And &HFF
            raw(at + 2) = (color \ 65536) And &HFF
            raw(at + 3) = CByte(alpha * 255): at = at + 4
        Next
    Next
    WritePng path, w, h, raw: RoundedCardPng = path
End Function


' Fixed capsule width on a wider transparent cell canvas.
Public Function ColumnPillPng(ByVal color As Variant, ByVal width As Long, ByVal height As Long, ByVal uniformWidth As Long) As Variant
    ColumnPillPng = PillTextPng(color, Null, width, height, uniformWidth)
End Function

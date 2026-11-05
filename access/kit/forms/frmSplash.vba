Option Compare Database
Option Explicit

' Splash screen (BuildSplash builds it). It is the app's StartUpForm: it shows
' the app's name and version in the user's palette, then, from its timer so it
' has painted first, opens the form named in its Tag (frmMain), hidden when
' the Tag ends in SPLASH_HIDDEN (BuildSplash openHidden). That form's Load
' does the app's start-up work under the splash, calling SplashStatus to say
' what it is doing. Once the form is open the splash closes, after at least
' MIN_SHOW_MS so a fast start doesn't just flash it.

Private Const MIN_SHOW_MS As Long = 1200
Private Const SECONDS_PER_DAY As Long = 86400
Private m_openedAt As Single        ' Timer() when the splash opened
Private m_started As Boolean        ' the next form has been opened
Private m_name As String            ' this form's name, for after it may have closed
Private m_retried As Boolean        ' a failed Timer has rescheduled the close once

Private Sub Form_Load()
    On Error GoTo Err_Handler
    m_openedAt = Timer
    m_name = Me.Name
    Me.TimerInterval = 50           ' paint first, then start (set first, so the app starts whatever fails below)
    Me.lblSplashVersion.Caption = VersionText()
    CenterOnApp Me                  ' over the Access window, on its monitor
    Exit Sub
Err_Handler:
    HandleError "frmSplash.Form_Load"
End Sub

Private Sub Form_Timer()
    Dim waitMs As Long, nextForm As String, hidden As Boolean
    On Error GoTo Err_Handler
    Me.TimerInterval = 0
    If Not m_started Then
        m_started = True
        nextForm = Me.Tag
        hidden = (Right$(nextForm, Len(SPLASH_HIDDEN)) = SPLASH_HIDDEN)
        If hidden Then nextForm = Left$(nextForm, Len(nextForm) - Len(SPLASH_HIDDEN))
        If Len(nextForm) = 0 Then nextForm = "frmMain"
        DoCmd.OpenForm nextForm, WindowMode:=IIf(hidden, acHidden, acWindowNormal)
        ' A message during the next form's start-up closes the splash (KitMsgBox).
        If Not IsFormOpen(m_name) Then Exit Sub
    End If
    waitMs = MIN_SHOW_MS - ElapsedMs()
    If waitMs > 0 Then
        Me.TimerInterval = waitMs
    Else
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    Exit Sub
Err_Handler:
    ' HandleError's message closes the splash first (KitMsgBox, CloseSplash),
    ' but Access can refuse that (2585) while this Timer event is running. So
    ' run the timer once more: m_started is set, so it only closes the splash.
    ' Once only, so a close that keeps failing can't repeat its message.
    HandleError "frmSplash.Form_Timer"
    If Not m_retried And IsFormOpen(m_name) Then
        m_retried = True
        Me.TimerInterval = 100
    End If
End Sub

' Milliseconds since the splash opened (Timer() restarts at midnight).
Private Function ElapsedMs() As Long
    Dim s As Single
    s = Timer - m_openedAt
    If s < 0 Then s = s + SECONDS_PER_DAY
    ElapsedMs = CLng(s * 1000)
End Function

' "Version 1.0.0" from the database property AppVersion, else "Build <stamp>"
' from AppBuild (BuildAll sets it), else nothing. A missing property is
' expected, so its error is skipped.
Private Function VersionText() As String
    Dim db As DAO.Database, v As String
    Set db = CurrentDb
    On Error Resume Next
    v = db.Properties("AppVersion").Value
    If Len(v) > 0 Then
        VersionText = "Version " & v
    Else
        v = db.Properties("AppBuild").Value
        If Len(v) > 0 Then VersionText = "Build " & v
    End If
    On Error GoTo 0
End Function

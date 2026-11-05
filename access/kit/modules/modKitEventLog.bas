Option Compare Database
Option Explicit

' Access UI kit - writing to the Windows Event Log (the Application log), so a
' site's IT collects each app's errors on its own log server with Windows
' Event Forwarding, as it does for every other program on the network.
'
' LogError (modErrors) calls KitEventLogError for each entry it records. Each
' event's source is APP_NAME, so a forwarding subscription or an Event Viewer
' filter can pick out one app. Event ids:
'   EVENT_ID_ERROR (1001)  an error the app logged (LogError)
'   EVENT_ID_INFO  (1000)  anything else an app chooses to record (KitEventLog)
'
' Writing needs no admin rights. Until the source is registered on a PC, Event
' Viewer adds a line saying the event's description can't be found, and shows
' the text after it. To register it, IT runs once per PC, as an admin (or
' deploys it by Group Policy):
'     New-EventLog -LogName Application -Source "<APP_NAME>"
'
' What an error event holds: the error number and description, the procedure
' (with its line), the page, form and control, the app user and the front
' end's build and Access version. Not the context or call trail, which can name
' the record being worked on; those stay in the app's own error log. Writing
' never raises and never shows anything: a PC whose event log refuses the write
' simply has no event.
'
' The app provides APP_NAME.

Private Declare PtrSafe Function RegisterEventSourceW Lib "advapi32" (ByVal lpUNCServerName As LongPtr, _
    ByVal lpSourceName As LongPtr) As LongPtr
Private Declare PtrSafe Function ReportEventW Lib "advapi32" (ByVal hEventLog As LongPtr, ByVal wType As Integer, _
    ByVal wCategory As Integer, ByVal dwEventID As Long, ByVal lpUserSid As LongPtr, ByVal wNumStrings As Integer, _
    ByVal dwDataSize As Long, ByRef lpStrings As LongPtr, ByVal lpRawData As LongPtr) As Long
Private Declare PtrSafe Function DeregisterEventSource Lib "advapi32" (ByVal hEventLog As LongPtr) As Long

Public Const EVENT_ID_INFO As Long = 1000
Public Const EVENT_ID_ERROR As Long = 1001
Public Const EVENT_TYPE_ERROR As Integer = 1
Public Const EVENT_TYPE_WARNING As Integer = 2
Public Const EVENT_TYPE_INFO As Integer = 4

Private Const MAX_EVENT_TEXT As Long = 30000      ' an event's text tops out at 31,839 characters

Private m_written As Long                          ' events written this session

' Write one event to the Application log with APP_NAME as its source. True when
' Windows took it. Never raises.
Public Function KitEventLog(ByVal text As String, Optional ByVal eventType As Integer = EVENT_TYPE_INFO, _
                            Optional ByVal eventId As Long = EVENT_ID_INFO) As Boolean
    Dim h As LongPtr, src As String, p As LongPtr
    On Error GoTo Fail
    src = APP_NAME
    h = RegisterEventSourceW(0, StrPtr(src))
    If h = 0 Then Exit Function
    text = Left$(text, MAX_EVENT_TEXT)
    p = StrPtr(text)
    KitEventLog = (ReportEventW(h, eventType, 0, eventId, 0, 1, 0, p, 0) <> 0)
    DeregisterEventSource h
    If KitEventLog Then m_written = m_written + 1
    Exit Function
Fail:
    ' The event log is best effort: an error here must never reach the user or the error log.
End Function

' The error event for one LogError entry. id is the error log entry (0 when it
' went to the fallback file instead).
Public Function KitEventLogError(ByVal id As Long, ByVal num As Long, ByVal desc As String, ByVal source As String, _
                                 ByVal page As String, ByVal formName As String, ByVal controlName As String, _
                                 ByVal appUser As String, ByVal build As String) As Boolean
    Dim s As String
    On Error GoTo Fail
    s = APP_NAME & ": error " & num & ": " & desc & vbCrLf & _
        "Procedure: " & source & vbCrLf & _
        "Page: " & page & "   Form: " & formName & "   Control: " & controlName & vbCrLf & _
        "App user: " & IIf(Len(appUser) > 0, appUser, "(not signed in)") & vbCrLf & _
        "Build: " & build & "   Access " & Application.Version & "." & Application.Build & vbCrLf & _
        IIf(id > 0, "Error log entry " & id & ".", "Not in the error log (written to the fallback file).")
    KitEventLogError = KitEventLog(s, EVENT_TYPE_ERROR, EVENT_ID_ERROR)
    Exit Function
Fail:
End Function

' How many events this session wrote (for the tests and a diagnostics page).
Public Function KitEventsWritten() As Long
    KitEventsWritten = m_written
End Function

Attribute VB_Name = "modLogging"
Option Explicit

Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" ( _
                     ByVal lpUNCServerName As String, _
                     ByVal lpSourceName As String _
                     ) As Long

Declare Function DeregisterEventSource Lib "advapi32.dll" ( _
                     ByVal hEventLog As Long _
                     ) As Long

Declare Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" ( _
                     ByVal hEventLog As Long, ByVal wType As Integer, _
                     ByVal wCategory As Integer, ByVal dwEventID As Long, _
                     ByVal lpUserSid As Any, ByVal wNumStrings As Integer, _
                     ByVal dwDataSize As Long, plpStrings As Long, _
                     lpRawData As Any _
                     ) As Boolean
Declare Function GetLastError Lib "kernel32" () As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                     hpvDest As Any, _
                     hpvSource As Any, _
                     ByVal cbCopy As Long)
                     
Declare Function GlobalAlloc Lib "kernel32" ( _
                     ByVal wFlags As Long, _
                     ByVal dwBytes As Long _
                     ) As Long
Declare Function GlobalFree Lib "kernel32" ( _
                     ByVal hMem As Long _
                     ) As Long
   
Public Const EVENTLOG_SUCCESS = 0
Public Const EVENTLOG_ERROR_TYPE = 1
Public Const EVENTLOG_WARNING_TYPE = 2
Public Const EVENTLOG_INFORMATION_TYPE = 4
Public Const EVENTLOG_AUDIT_SUCCESS = 8
Public Const EVENTLOG_AUDIT_FAILURE = 10

Public Sub LogNTEvent(sString As String, iLogType As Integer, iEventID As Long)
Dim bRC As Boolean
Dim iNumStrings As Integer
Dim hEventLog As Long
Dim hMsgs As Long
Dim cbStringSize As Long

   hEventLog = RegisterEventSource("", App.Title)
   cbStringSize = Len(sString) + 1
   hMsgs = GlobalAlloc(&H40, cbStringSize)
   CopyMemory ByVal hMsgs, ByVal sString, cbStringSize
   iNumStrings = 1
   ReportEvent hEventLog, iLogType, 0, iEventID, 0&, iNumStrings, cbStringSize, hMsgs, hMsgs
   Call GlobalFree(hMsgs)
   DeregisterEventSource (hEventLog)
   
End Sub

Public Sub LogAction(sMsg As String, sFilename As String, sFilePath As String)
Dim lngFileNum As Long
Dim i As Integer
Dim sLine As String

    'Ignore Errors
    On Error Resume Next

    'Assign Free File Number
    lngFileNum = FreeFile

    'Open/Create File
    Open sFilePath & "\" & sFilename For Append As #lngFileNum
    
    'Construct Line
    sLine = Format(Now, "yyyy-mm-dd hh:MM:ss")
    sLine = sLine & "," & sMsg
    
    'Write Line
    Print #lngFileNum, sLine
    
    'Close File
    Close #lngFileNum   ' Close file.
End Sub





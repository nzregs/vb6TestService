VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Begin VB.Form frmMain 
   Caption         =   "Test VB6 COM Service"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmTimer 
      Left            =   3360
      Top             =   2160
   End
   Begin NTService.NTService NTService1 
      Left            =   1560
      Top             =   2160
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ServiceName     =   "Simple"
      StartMode       =   3
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Registry Settings
Public oSettings As New clsSettings

Private SERVICE_LOG_FILE As String
Private SERVICE_LOG_PATH As String

Private started            As Boolean

Private Sub Form_Load()
On Error GoTo Err_Load
    SERVICE_LOG_FILE = "TestVB6ComService.log"
    SERVICE_LOG_PATH = oSettings.ServiceLogDir
    
    LogAction "Loading Service", SERVICE_LOG_FILE, SERVICE_LOG_PATH
     
    Dim strDisplayName As String
    Dim bStarted As Boolean
    
    NTService1.ServiceName = App.EXEName
    NTService1.DisplayName = App.EXEName
    
    strDisplayName = NTService1.DisplayName
    
    If Command = "-install" Then
        ' enable interaction with desktop
        NTService1.Interactive = True
        LogAction App.EXEName & ": Service Install", SERVICE_LOG_FILE, SERVICE_LOG_PATH
        If NTService1.Install Then
            Call NTService1.SaveSetting("Parameters", "TimerInterval", "1000")
            MsgBox strDisplayName & " installed successfully"
            LogAction App.EXEName & ": Service installed", SERVICE_LOG_FILE, SERVICE_LOG_PATH
        Else
            MsgBox App.EXEName & ": " & strDisplayName & " failed to install"
            LogAction App.EXEName & ": Service failed to install", SERVICE_LOG_FILE, SERVICE_LOG_PATH
        End If
        bExitFlag = True
        End
    ElseIf Command = "-uninstall" Then
        LogAction "Service Uninstall Mode", SERVICE_LOG_FILE, SERVICE_LOG_PATH
        If NTService1.Uninstall Then
            MsgBox strDisplayName & " uninstalled successfully"
            LogAction "Service uninstalled", SERVICE_LOG_FILE, SERVICE_LOG_PATH
        Else
            MsgBox strDisplayName & " failed to uninstall"
            LogAction "Service failed to uninstall", SERVICE_LOG_FILE, SERVICE_LOG_PATH
        End If
        bExitFlag = True
        End
    ElseIf Command = "-debug" Then
        NTService1.Debug = True
        LogAction "Running debug mode", SERVICE_LOG_FILE, SERVICE_LOG_PATH
    ElseIf Command <> "" Then
        MsgBox "Invalid command option"
        LogAction "Invalid command option", SERVICE_LOG_FILE, SERVICE_LOG_PATH
        bExitFlag = True
        End
    End If
      ' enable Pause/Continue. Must be set before StartService
      ' is called or in design mode
      NTService1.ControlsAccepted = svcCtrlPauseContinue
      ' connect service to Windows NT services controller
      NTService1.StartService
    Exit Sub
Err_Load:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
    LogAction "[" & Err.Number & "] " & Err.Description, SERVICE_LOG_FILE, SERVICE_LOG_PATH
    bExitFlag = True
End Sub

Private Sub NTService1_Start(Success As Boolean)

   On Error GoTo Err_Start
       
   tmTimer.Interval = 1000
   tmTimer.Enabled = True
   
   LogAction "Service Started", SERVICE_LOG_FILE, SERVICE_LOG_PATH
Exit Sub
    
Err_Start:
   Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
   LogAction "[" & Err.Number & "] " & Err.Description, SERVICE_LOG_FILE, SERVICE_LOG_PATH
End Sub

Private Sub NTService1_Stop()

   On Error GoTo Err_Stop
   
   LogAction "Service Stopped", SERVICE_LOG_FILE, SERVICE_LOG_PATH
Exit Sub
    
Err_Stop:
   Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
   LogAction "[" & Err.Number & "] " & Err.Description, SERVICE_LOG_FILE, SERVICE_LOG_PATH
End Sub

Private Sub tmTimer_Timer()
Dim sDatetime As String
   
   sDatetime = Format(Now, "dd/mm/yy hh:mm:ss")
   
   'Let the OS do stuff for a moment
   DoEvents
   
   'on error GoTo tmTimer_Error
   LogAction "[" & sDatetime & "] Timer Does Something", SERVICE_LOG_FILE, SERVICE_LOG_PATH
   
Exit Sub
tmTimer_Error:
   Select Case Err.Number
      Case 5000:
         LogAction "Error during timer processing", SERVICE_LOG_FILE, SERVICE_LOG_PATH
         Resume Next
      Case Else:
         Resume Next
   End Select
End Sub

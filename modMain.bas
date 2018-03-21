Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()

Dim sCmdLine As String
Dim sCmdParms As String
Dim sCmdSwitch As String
On Error GoTo Err_Continue

    Load frmMain
    sCmdLine = Command
    If InStr(1, sCmdLine, "/") > 0 Then
        sCmdSwitch = Trim(Left(sCmdLine, InStr(1, sCmdLine, "/") - 1))
        sCmdParms = Trim(Right(sCmdLine, Len(sCmdLine) - InStr(1, sCmdLine, "/")))
    Else
        sCmdSwitch = Trim(sCmdLine)
    End If


    Select Case UCase(sCmdSwitch)
        Case "-INSTALL"
            If Len(Trim(sCmdParms)) > 0 Then
                ' disable interaction with desktop
                frmMain.NTService1.Interactive = False
                frmMain.NTService1.DisplayName = App.Title
                frmMain.NTService1.StartMode = svcStartAutomatic
                'get user and password from cmdparms
                'frmMain.NTService1.Account = Trim(Left(sCmdParms, InStr(1, sCmdParms, "/") - 1))
                'frmMain.NTService1.Password = Trim(Right(sCmdParms, Len(sCmdParms) - InStr(1, sCmdParms, "/")))
            Else
                ' enable interaction with desktop
                frmMain.NTService1.Interactive = True
                frmMain.NTService1.DisplayName = App.Title
                frmMain.NTService1.StartMode = svcStartAutomatic
            End If

            If frmMain.NTService1.Install Then
                MsgBox App.Title & " installed successfully"
            Else
                MsgBox App.Title & " failed to install"
            End If
            Unload frmMain
            
        Case "-UNINSTALL"
            If frmMain.NTService1.Uninstall Then
                MsgBox App.Title & " uninstalled successfully"
            Else
                MsgBox App.Title & " failed to uninstall"
            End If
            Unload frmMain
            
        Case ""
            'Call GetAppSettings
            ' enable Pause/Continue. Must be set before StartService
            ' is called or in design mode
            frmMain.NTService1.ControlsAccepted = svcCtrlPauseContinue
            frmMain.NTService1.StartService
            frmMain.tmTimer.Interval = 1000
            frmMain.tmTimer.Enabled = True
        Case Else
            MsgBox "Invalid command option"
        End Select
Exit Sub

Err_Continue:
        Call frmMain.NTService1.LogEvent(svcEventError, svcMessageError, "[" & Err.Number & "] " & Err.Description)
        Unload frmMain
End Sub


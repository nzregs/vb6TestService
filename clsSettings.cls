VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const RegistryBase = "SOFTWARE\MSFTCSE\"

Private pRegBase As String

Private Sub Class_Initialize()
    pRegBase = RegistryBase & App.EXEName & "\"
End Sub

Public Property Get SyncDelay()
Dim strSyncDelay As String
    loadFromRegistry "SyncDelay", strSyncDelay
    SyncDelay = strSyncDelay
End Property

Public Property Let SyncDelay(ByVal newSyncDelay)
    If newSyncDelay <> vbNullString Then
        loadToRegistry "SyncDelay", CStr(newSyncDelay)
    Else
        'Err.Raise
    End If
End Property



Public Property Get SQLSystem()
Dim strSQLSystem As String
    loadFromRegistry "SQLSystem", strSQLSystem
    SQLSystem = strSQLSystem
End Property

Public Property Let SQLSystem(ByVal newSQLSystem)
    If newSQLSystem <> vbNullString Then
        loadToRegistry "SQLSystem", CStr(newSQLSystem)
    Else
        'Err.Raise
    End If
End Property

Public Property Get SQLDBTable()
Dim strSQLDBTable As String
    loadFromRegistry "SQLDBTable", strSQLDBTable
    SQLDBTable = strSQLDBTable
End Property

Public Property Let SQLDBTable(ByVal newSQLDBTable)
    If newSQLDBTable <> vbNullString Then
        loadToRegistry "SQLDBTable", CStr(newSQLDBTable)
    Else
        'Err.Raise
    End If
End Property

Public Property Get ServiceLogDir()
Dim strServiceLogDir As String
    loadFromRegistry "ServiceLogDir", strServiceLogDir
    ServiceLogDir = strServiceLogDir
End Property

Public Property Let ServiceLogDir(ByVal newServiceLogDir)
    If newServiceLogDir <> vbNullString Then
        loadToRegistry "ServiceLogDir", CStr(newServiceLogDir)
    Else
        'Err.Raise
    End If
End Property

Private Sub loadFromRegistry(strRegKey As String, strRegValue As String)
    If CreateRegKey(HKEY_LOCAL_MACHINE, pRegBase) Then
        strRegValue = GetRegValue(HKEY_LOCAL_MACHINE, pRegBase, strRegKey, strRegValue)
    Else
        'err.raise
    End If
End Sub

Private Sub loadToRegistry(strRegKey As String, strRegValue As String)
    SetRegValue HKEY_LOCAL_MACHINE, pRegBase, strRegKey, strRegValue
End Sub


Public Property Get SMTPServer()
Dim strSMTPServer As String
    loadFromRegistry "SMTPServer", strSMTPServer
    SMTPServer = strSMTPServer
End Property

Public Property Let SMTPServer(ByVal newSMTPServer)
    If newSMTPServer <> vbNullString Then
        loadToRegistry "SMTPServer", CStr(newSMTPServer)
    Else
        'Err.Raise
    End If
End Property

Public Property Get SMTPFromAddress()
Dim strSMTPFromAddress As String
    loadFromRegistry "SMTPFromAddress", strSMTPFromAddress
    SMTPFromAddress = strSMTPFromAddress
End Property

Public Property Let SMTPFromAddress(ByVal newSMTPFromAddress)
    If newSMTPFromAddress <> vbNullString Then
        loadToRegistry "SMTPFromAddress", CStr(newSMTPFromAddress)
    Else
        'Err.Raise
    End If
End Property




    'Err.Raise
    End If
End Property

Public Property Get SMTPFromAddress()
Dim strSMTPFromAddress As String
    loadFromRegistry "SMTPFromAddress", strSMTPFromAddress
    SMTPFromAddress = strSMTPFromAddress
End Property

Public Property Let SMTPFromAddress(ByVal newSMTPFromAddress)
    If newSMTPFromAddress <> vbNullString Then
        loadToRegistry "SMTPFromAddress", CStr(newSMTPFromAddress)
    Else
        'Err.Raise
    End If
End Property





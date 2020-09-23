VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADOConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public oADODBConnection As ADODB.Connection
  
'  Dim OpenADO As New clsADOConnect
'  Dim MyDB As ADODB.Connection
'  Dim MySet As New ADODB.Recordset
'    OpenADO.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, DBname, DBname
'    MySet.open "SELECT Location.* FROM Location " & OrderBy, MyDB, adOpenStatic,adLockPessimistic

'Typical Connection String
'    Driver={Microsoft Access Driver (*.mdb)};Dbq=C:\PC_DATA\Security.pwd;DefaultDir=C:\PC_DATA\;PWD=MorganH;
' or
'    Adodc1.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & AppPath & dbFile & ";DefaultDir=" & AppPath & DB_PWD
'    Adodc1.CommandType = adCmdText
'    Adodc1.Mode = adModeShareDenyNone
'    Adodc1.RecordSource = "Select Location.* from Location Where Location.Void=False Order by Location.InLet;"
'    Adodc1.Refresh
' 0r
'    Adodc1.ConnectionString = OpenADO.adoConnectString(dbt_MicrosoftAccessFile, DBname, , , DB_PWD)
' Or
'    Call AdodcConnect(adodc1,SQLstmt)

Public Enum e_DatabaseTypes
    dbt_Undefined = 0
    dbt_OracleMSDA = 1
    dbt_OracleODBC = 2
    dbt_SQLserver = 3
    dbt_MicrosoftAccess97File = 4
    dbt_MicrosoftAccess2KFile = 5
    dbt_MicrosoftAccessFile = 6
    dbt_DSNFile = 7
End Enum

Public Function CompressAccessDatabase(mSourceDB As String, Optional Password As String = vbNullString, Optional DB_Type As e_DatabaseTypes = dbt_MicrosoftAccess2KFile) As Boolean
  Dim srcDB As String
  Dim destDB As String
  Dim JRO As JRO.JetEngine
  Dim EngineType As String

    On Error GoTo Err
    
    If DB_Type = dbt_MicrosoftAccess2KFile Then
        EngineType = "5"
    Else
        EngineType = "4"
    End If
    
    srcDB = mSourceDB
    destDB = App.PATH & "\MHbackup.mdb"
    If Dir(destDB) > vbNullString Then Kill destDB
    
    Set JRO = New JRO.JetEngine
    JRO.CompactDatabase _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & srcDB & ";Jet OLEDB:Database Password=" & Password & ";", _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & destDB & ";Jet OLEDB:Database Password=" & Password & ";Jet OLEDB:Engine Type=" & EngineType & ";"

    Kill srcDB

    DoEvents
    Name destDB As srcDB
    
    Set JRO = Nothing
    CompressAccessDatabase = True
Exit Function

Err:
    CompressAccessDatabase = False
    Set JRO = Nothing
End Function

Public Function BuildConnectString(ByVal DatabaseType As e_DatabaseTypes, ByVal ServerOrFilename As String, Optional ByVal databaseName As String, Optional ByVal UserName As String, Optional ByVal Password As String) As String
    Select Case DatabaseType
        Case dbt_OracleMSDA
            BuildConnectString = "Provider=MSDAORA;Data Source=" & ServerOrFilename & ";User ID=" & IIf(UserName <> vbNullString, UserName, vbNullString) & ";Password=" & IIf(Password <> vbNullString, Password, vbNullString) & ";" & IIf(databaseName <> vbNullString, "Initial Catalog=" & databaseName & ";", vbNullString)
        Case dbt_OracleODBC
            BuildConnectString = "DRIVER={Microsoft ODBC for Oracle};SERVER=" & ServerOrFilename & ";UID=" & UserName & ";PWD=" & Password & ";" & IIf(databaseName <> vbNullString, "Initial Catalog=" & databaseName & ";", vbNullString)
        Case dbt_SQLserver
            BuildConnectString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & ServerOrFilename & ";User ID=" & IIf(UserName <> vbNullString, UserName, vbNullString) & ";Password=" & IIf(Password <> vbNullString, Password, vbNullString) & ";" & IIf(databaseName <> vbNullString, "Initial Catalog=" & databaseName & ";", vbNullString)
        Case dbt_DSNFile
            BuildConnectString = "Provider=MSDASQL;DSN=" & ServerOrFilename & ";UID=" & IIf(UserName <> vbNullString, UserName, vbNullString) & ";PWD=" & IIf(Password <> vbNullString, Password & ";", vbNullString) & ";" & IIf(databaseName <> vbNullString, "Initial Catalog=" & databaseName & ";", vbNullString)
        Case dbt_MicrosoftAccess2KFile
            BuildConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                 "Data Source=" & ServerOrFilename & ";" & _
                                 "Jet OLEDB:Database Password=" & Password & ";" & _
                                 "Jet OLEDB:Engine Type=5;"
        Case dbt_MicrosoftAccess97File
            BuildConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                 "Data Source=" & ServerOrFilename & ";" & _
                                 "Jet OLEDB:Database Password=" & Password & ";" & _
                                 "Jet OLEDB:Engine Type=4;"
        Case dbt_MicrosoftAccessFile
            BuildConnectString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & ServerOrFilename & ";DefaultDir=" & DetermineDirectory(ServerOrFilename) & ";PWD=" & IIf(Password <> vbNullString, Password & ";", ";")
    End Select
End Function

Private Function DetermineDirectory(inputString As String) As String
  Dim i As Integer
    i = InStrRev(inputString, "\", , vbTextCompare)
    DetermineDirectory = Mid(inputString, 1, i)
End Function

Public Function adoConnectOpen(oADODBConnection As ADODB.Connection, _
                                DB_Type As e_DatabaseTypes, _
                                ByVal ServerOrFilename As String, _
                                Optional ByVal dbPathFilename As String = vbNullString, _
                                Optional CommandType As CommandTypeEnum = adCmdStoredProc, _
                                Optional CursorLocation As CursorLocationEnum = adUseClient, _
                                Optional ByVal UserName As String = vbNullString, _
                                Optional ByVal Password As String = vbNullString) As Boolean
    
    
    On Error GoTo adoConnectError
    
    Set oADODBConnection = New ADODB.Connection
    
    If dbPathFilename = vbNullString Then dbPathFilename = ServerOrFilename

'    oADODBConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\NorthWind.mdb;Jet OLEDB:Database Password=password;"

    
    With oADODBConnection
        .CursorLocation = CursorLocation '/* default=adUseClient(3)
        .Open BuildConnectString(DB_Type, ServerOrFilename, dbPathFilename, UserName, Password)
    End With

    adoConnectOpen = True
    
Exit Function
    
    
adoConnectError:
    MsgBox "Connect To database failed with: " & Err.Number & " " & Err.Description
    adoConnectOpen = False
    
End Function
Public Function adoConnectString(DB_Type As e_DatabaseTypes, _
                                ByVal ServerOrFilename As String, _
                                Optional ByVal dbPathFilename As String = vbNullString, _
                                Optional ByVal UserName As String = vbNullString, _
                                Optional ByVal Password As String = vbNullString) As String
    
    If dbPathFilename = vbNullString Then dbPathFilename = ServerOrFilename
    adoConnectString = BuildConnectString(DB_Type, ServerOrFilename, dbPathFilename, UserName, Password)
    
End Function


Public Function adoConnectClose(oADODBConnection As ADODB.Connection) As Boolean
    
    On Error GoTo adoConnectError
    oADODBConnection.Close
    adoConnectClose = True
    
Exit Function
    
    
adoConnectError:
    MsgBox "Connect To database failed with: " & Err.Number & " " & Err.Description
    adoConnectClose = False
    
End Function




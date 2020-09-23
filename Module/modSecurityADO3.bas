Attribute VB_Name = "modSecurity"
'/*************************************/
'/* Author: Morgan Haueisen
'/* Copyright (c) 1997-2002
'/*************************************/

Option Explicit

'/* For password protected database file (if required) */
Public Const DB_PWD As String = vbNullString
Public Const DB_Type As String = "4" '/* 4=Access97; 5=Access2000

Public Type goUserType
    UserName As String
    UserFullName As String
    Password As String
    MachineName As String
End Type
Public goUser As goUserType

Public Type goApplicationType
    SourceDatabasePath As String
    SecurityDatabasePath As String
    SystemID As String
End Type
Public goApplication As goApplicationType

Public LogOffT As Byte
Public Const LogOffM As Byte = 10
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function ADOFindFirst(MySet As ADODB.Recordset, ByVal Filter As String) As Boolean
  Dim mhRS As ADODB.Recordset
  Dim mhMatch As Boolean

    Set mhRS = New ADODB.Recordset
    Set mhRS = MySet.Clone
    mhRS.Filter = Filter
    
    If mhRS.RecordCount > 0 Then
        mhRS.MoveFirst
        MySet.Bookmark = mhRS.Bookmark
        mhMatch = True
    Else
        If MySet.RecordCount > 0 Then
            MySet.MoveLast
            MySet.MoveNext
        End If
        mhMatch = False
    End If
    
    mhRS.Close
    Set mhRS = Nothing
    
    ADOFindFirst = mhMatch

End Function

Public Function ADOFindNext(MySet As ADODB.Recordset, ByVal Filter As String) As Boolean
  Dim mhRS As ADODB.Recordset
  Dim mhMatch As Boolean

    Set mhRS = New ADODB.Recordset
    Set mhRS = MySet.Clone
    mhRS.Filter = Filter
    mhRS.Sort = MySet.Sort
    
    If mhRS.RecordCount > 0 Then
        mhRS.Bookmark = MySet.Bookmark
        mhRS.MoveNext
        If Not mhRS.EOF Then
            MySet.Bookmark = mhRS.Bookmark
            mhMatch = True
        Else
            mhMatch = False
        End If
    Else
        If MySet.RecordCount > 0 Then
            MySet.MoveLast: MySet.MoveNext
        End If
        mhMatch = False
    End If
    
    mhRS.Close
    Set mhRS = Nothing
    ADOFindNext = mhMatch
    
End Function

Public Sub ADOCreateQuery(ByVal sFilename As String, ByVal QueryName As String, ByVal QueryString As String, Optional AddDelay As Boolean = False)
  Dim OpenConnect As New clsADOConnect
  Dim CAT As New ADOX.Catalog
  Dim CMD As New ADODB.Command

    On Local Error Resume Next
    '/* Open the catalog
    CAT.ActiveConnection = OpenConnect.adoConnectString(dbt_MicrosoftAccess2KFile, sFilename, sFilename, , DB_PWD)
    
    '/* Create the query
    CMD.CommandText = QueryString
    CAT.Views.Append QueryName, CMD
    DoEvents
    
    Set CAT = Nothing
    'Set CMD = Nothing
    'Set OpenConnect = Nothing
    On Local Error GoTo 0
    DoEvents
    If AddDelay Then Sleep 5000

End Sub

Public Sub ADODeleteQuery(ByVal sFilename As String, ByVal QueryName As String)
  Dim OpenConnect As New clsADOConnect
  Dim CAT As New ADOX.Catalog
  Dim CMD As New ADODB.Command

    '/* Open the catalog
    CAT.ActiveConnection = OpenConnect.adoConnectString(dbt_MicrosoftAccess2KFile, sFilename, sFilename, , DB_PWD)
    '/* Delete the query
    CAT.Views.Delete QueryName
    
    Set CAT = Nothing
    Set CMD = Nothing
    Set OpenConnect = Nothing

End Sub


Public Sub ADOAttachTable(TableName As String, ByVal AttachToMDB As String, ByVal AttachFromMDB As String, Optional ByVal AsTableName As String = vbNullString)
  Dim OpenConnect As clsADOConnect
  Dim CAT As ADOX.Catalog
  Dim TBL As ADOX.Table

   Set OpenConnect = New clsADOConnect
   Set CAT = New ADOX.Catalog
   Set TBL = New ADOX.Table

   On Local Error Resume Next
   
   If AsTableName = vbNullString Then AsTableName = TableName
   
   '/* Open the catalog
   CAT.ActiveConnection = OpenConnect.BuildConnectString(dbt_MicrosoftAccess2KFile, AttachToMDB, , , DB_PWD)

   '/* Set the name and target catalog for the table
   TBL.Name = AsTableName
   Set TBL.ParentCatalog = CAT

   '/* Set the properties to create the link
   TBL.Properties("Jet OLEDB:Create Link") = True
   TBL.Properties("Jet OLEDB:Link Datasource") = AttachFromMDB
   TBL.Properties("Jet OLEDB:Link Provider String") = ";Pwd=" & DB_PWD
   TBL.Properties("Jet OLEDB:Remote Table Name") = TableName

   '/* Append the table to the collection
   CAT.Tables.Append TBL

   Set CAT = Nothing
   Set TBL = Nothing
   Set OpenConnect = Nothing
   On Local Error GoTo 0

End Sub

Private Sub SetMenuAccess_Example()
'  Dim oSecurity As clsSecurity
'  Dim vData() As Variant
'  Dim i As Integer
'
'    '/* Disable Secure Menu Items
'    mnuFile.Enabled = False
'    mnuReports.Enabled = False
'    mnuTables.Enabled = False
'    mnuInput.Enabled = False
'    mnuFinancial.Enabled = False
'    mnuPlan.Enabled = False
'    mnuShiftInfo.Enabled = False
'    mnuSecurity.Enabled = False
'    If goUser.UserName = vbNullString Then Exit Sub
'
'    Set oSecurity = New clsSecurity
'    '/* Enable Secure Menu Items based on User's rights
'    If oSecurity.GetMembership(goUser.UserName, vData, goApplication.SecurityDatabasePath) Then
'        For i = 0 To UBound(vData)
'            Select Case vData(i)
'            Case "Maintenance"
'                mnuFile.Enabled = True
'                mnuReports.Enabled = True
'            Case "Planning"
'                mnuInput.Enabled = True
'                mnuPlan.Enabled = True
'            Case "Production"
'                mnuFile.Enabled = True
'                mnuInput.Enabled = True
'                mnuShiftInfo.Enabled = True
'            Case "Reports"
'                mnuReports.Enabled = True
'            Case "Financial"
'                mnuFinancial.Enabled = True
'                mnuReports.Enabled = True
'            Case "Adminstrator"
'                mnuFile.Enabled = True
'                mnuReports.Enabled = True
'                mnuTables.Enabled = True
'                mnuInput.Enabled = True
'                mnuFinancial.Enabled = True
'                mnuPlan.Enabled = True
'                mnuShiftInfo.Enabled = True
'                mnuSecurity.Enabled = True
'            End Select
'        Next i
'    End If
'
'    Erase vData
'    Set oSecurity = Nothing

End Sub


Public Sub InitSecurity(ByVal MDBfile As String, ByVal MDWfile As String, Optional ByVal MDWSystemID As String = vbNullString, Optional IgnorFault As Boolean = False)
    
    MDWSystemID = "morganh" & Trim(MDWSystemID)
    
    goApplication.SourceDatabasePath = MDBfile
    goApplication.SecurityDatabasePath = MDWfile
    goApplication.SystemID = MDWSystemID
    
    If Dir$(MDBfile) = vbNullString Then
        If Not IgnorFault Then
            MsgBox "The Database file is missing.  Please contact your system adminstrator for assistance", vbCritical
            End
        End If
    ElseIf Dir$(MDWfile) = vbNullString Then
        MsgBox "The Security file is missing.  Please contact your system adminstrator for assistance", vbCritical
        End
    End If
    
End Sub

Public Sub OpenDB(MyDB As ADODB.Connection, Optional ByVal OpenMDB As Boolean = True, Optional ByVal DBPathName As String = vbNullString)
  Dim OpenConnect As clsADOConnect
  
    If DBPathName = vbNullString Then
        DBPathName = goApplication.SourceDatabasePath
    End If
    
    If OpenMDB Then
        '/* Password protected database file */
        Set OpenConnect = New clsADOConnect
        OpenConnect.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, DBPathName, DBPathName, , , , DB_PWD
        Set OpenConnect = Nothing
    Else
        MyDB.Close
        Set MyDB = Nothing
    End If
    DoEvents
End Sub

Public Sub OpenRS(oActiveRecordset As ADODB.Recordset, ByVal oSourceTable As String, oActiveConnection As ADODB.Connection, Optional oCursorType As CursorTypeEnum = adOpenStatic, Optional oLockType As LockTypeEnum = adLockOptimistic, Optional ByVal oOptions As Integer = -1)
    Set oActiveRecordset = New ADODB.Recordset
    oActiveRecordset.Open oSourceTable, oActiveConnection, oCursorType, oLockType, oOptions
End Sub
Public Function ADOFindPrevious(MySet As ADODB.Recordset, ByVal Filter As String) As Boolean
  Dim mhRS As ADODB.Recordset
  Dim mhMatch As Boolean

    Set mhRS = New ADODB.Recordset
    Set mhRS = MySet.Clone
    mhRS.Filter = Filter
    mhRS.Sort = MySet.Sort
    
    If mhRS.RecordCount > 0 Then
        mhRS.Bookmark = MySet.Bookmark
        mhRS.MovePrevious
        If (Not mhRS.BOF) Then
            MySet.Bookmark = mhRS.Bookmark
            mhMatch = True
        Else
            mhMatch = False
        End If
    Else
        If MySet.RecordCount > 0 Then
            MySet.MoveFirst: MySet.MovePrevious
        End If
        mhMatch = False
    End If
    
    mhRS.Close
    Set mhRS = Nothing

    ADOFindPrevious = mhMatch
    
End Function

Public Function ADOFindLast(MySet As ADODB.Recordset, ByVal Filter As String) As Boolean
  Dim mhRS As ADODB.Recordset
  Dim mhMatch As Boolean

    Set mhRS = New ADODB.Recordset
    Set mhRS = MySet.Clone
    mhRS.Filter = Filter
    
    If mhRS.RecordCount > 0 Then
        mhRS.MoveLast
        MySet.Bookmark = mhRS.Bookmark
        mhMatch = True
    Else
        If MySet.RecordCount > 0 Then
            MySet.MoveLast
            MySet.MoveNext
        End If
        mhMatch = False
    End If
    
    mhRS.Close
    Set mhRS = Nothing
    ADOFindLast = mhMatch

End Function

Public Sub ADODeleteTable(sFilename As String, sTableName As String)
  Dim OpenConnect As New clsADOConnect
  Dim CAT As ADOX.Catalog
    
    On Error GoTo ErrTrapD
    
    Set CAT = New ADOX.Catalog
    
    '/* Open Database
    CAT.ActiveConnection = OpenConnect.adoConnectString(dbt_MicrosoftAccess2KFile, sFilename, sFilename, , DB_PWD)
        
    '"Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & sFilename & ";" & _
               "Jet OLEDB:Database Password=" & DB_PWD & ";" & _
               "Jet OLEDB:Engine Type=5;"
    
    '/* Delete table
    'Dim TBL As ADOX.Table
    'Set TBL = New ADOX.Table
    'TBL.Name = sTableName
    'Set TBL.ParentCatalog = CAT
    CAT.Tables.Delete sTableName 'TBL
    
    Set OpenConnect = Nothing
    'Set TBL = Nothing
    Set CAT = Nothing
Exit Sub

ErrTrapD:
  'MsgBox Err.Number & " / " & Err.Description
  Exit Sub
  Resume

End Sub

Public Sub JROCompactDatabase(Optional DBPathName As String = vbNullString)
'  Dim JE As New JRO.JetEngine
'  Dim TempPath As String
'  Dim ConString As String
'  Dim cFile As clsFileOp
'
'    If DBPathName = vbNullString Then
'        DBPathName = goApplication.SourceDatabasePath
'    End If
'
'    ConString = ";Jet OLEDB:Database Password=" & DB_PWD & ";Jet OLEDB:Engine Type=" & DB_Type & ";"
'
'    Set cFile = New clsFileOp
'    TempPath = cFile.RetOnlyPath(DBPathName) & "CompactDatabase.mdb"
'
'    ' Make sure there isn't already a file with the name of the compacted database.
'    If Dir(TempPath) <> vbnullstring Then cFile.DeleteFile TempPath
'
'    ' Compact the database
'    JE.CompactDatabase "Data Source=" & DBPathName & ConString, "Data Source=" & TempPath & ConString
'    ' Delete the original database
'    cFile.DeleteFile DBPathName
'    ' Rename the file back to the original name
'    cFile.RenameFile TempPath, DBPathName
'
'    Set cFile = Nothing

End Sub



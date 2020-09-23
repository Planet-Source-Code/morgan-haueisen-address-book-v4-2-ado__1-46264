Attribute VB_Name = "modADOdc"
'/*************************************/
'/* Author: Morgan Haueisen
'/* Copyright (c) 1997-2002
'/*************************************/

Option Explicit
Public Sub ADOdcConnect(MyADOdc As Adodc, Optional ByVal SQLSource As String = vbNullString, Optional DBPathName As String = vbNullString)
  Dim OpenConnect As clsADOConnect
    
    On Local Error GoTo ErrMsg:
    Set OpenConnect = New clsADOConnect
    
    If DBPathName = vbNullString Then
        DBPathName = goApplication.SourceDatabasePath
    End If
    
    MyADOdc.CommandType = adCmdText
    MyADOdc.CursorType = adOpenStatic
    MyADOdc.LockType = adLockOptimistic 'adLockPessimistic
    MyADOdc.Mode = adModeShareDenyNone
    MyADOdc.CursorLocation = adUseClient
    
    MyADOdc.ConnectionString = OpenConnect.adoConnectString(dbt_MicrosoftAccess2KFile, DBPathName, DBPathName, , DB_PWD)
    Set OpenConnect = Nothing
    
    If SQLSource > vbNullString Then
        MyADOdc.RecordSource = SQLSource
        MyADOdc.Refresh
    End If
Exit Sub

ErrMsg:
    MsgBox Err.Number & vbCrLf & Err.Description
    Resume Next
    
End Sub

Public Function ADOdcFindFirst(MyADOdc As Adodc, ByVal FindString As String) As Boolean
  Dim MyDB As ADODB.Connection
  Dim MySet As ADODB.Recordset

    On Local Error Resume Next
    
    Set MyDB = New ADODB.Connection
    Set MySet = New ADODB.Recordset
    
    MyDB.CursorLocation = adUseClient
    MyDB.Open MyADOdc.ConnectionString
    
    MySet.Open MyADOdc.RecordSource, MyDB, adOpenStatic, adLockPessimistic

    If ADOFindFirst(MySet, FindString) Then
        MyADOdc.Recordset.BookMark = MySet.BookMark
        ADOdcFindFirst = True
    Else
        If Not (MyADOdc.Recordset.EOF And MyADOdc.Recordset.BOF) Then
            MyADOdc.Recordset.MoveLast
            MyADOdc.Recordset.MoveNext
        End If
        ADOdcFindFirst = False
    End If
        
    MySet.Close
    MyDB.Close
    
    Set MySet = Nothing
    Set MyDB = Nothing
    
    On Local Error GoTo 0
    
End Function
Public Function ADOdcFindNext(MyADOdc As Adodc, ByVal Filter As String) As Boolean
  Dim MyDB As ADODB.Connection
  Dim MySet As ADODB.Recordset
  Dim oNoMatch As Boolean

    On Local Error Resume Next
    
    Set MyDB = New ADODB.Connection
    Set MySet = New ADODB.Recordset
    
    MyDB.CursorLocation = adUseClient
    MyDB.Open MyADOdc.ConnectionString
    
    MySet.Open MyADOdc.RecordSource, MyDB, adOpenStatic, adLockPessimistic
    MySet.Filter = Filter
    MySet.Sort = MyADOdc.Recordset.Sort
    
    If Not (MySet.EOF And MySet.BOF) Then
        MySet.BookMark = MyADOdc.Recordset.BookMark
        MySet.MoveNext
        If (Not MySet.EOF) Then
            MyADOdc.Recordset.BookMark = MySet.BookMark
            oNoMatch = True
        Else
            oNoMatch = False
        End If
    Else
        If Not (MyADOdc.Recordset.EOF And MyADOdc.Recordset.BOF) Then
            MyADOdc.Recordset.MoveLast
            MyADOdc.Recordset.MoveNext
        End If
        oNoMatch = False
    End If
    
    MySet.Close
    MyDB.Close
    Set MyDB = Nothing
    Set MySet = Nothing
    
    ADOdcFindNext = oNoMatch
    
End Function

Public Function ADOdcFindLast(MyADOdc As Adodc, ByVal FindString As String) As Boolean
  Dim MyDB As ADODB.Connection
  Dim MySet As ADODB.Recordset

    On Local Error Resume Next
    
    Set MyDB = New ADODB.Connection
    Set MySet = New ADODB.Recordset
    
    MyDB.CursorLocation = adUseClient
    MyDB.Open MyADOdc.ConnectionString
    
    MySet.Open MyADOdc.RecordSource, MyDB, adOpenStatic, adLockPessimistic

    If ADOFindFirst(MySet, FindString) Then
        MyADOdc.Recordset.BookMark = MySet.BookMark
        ADOdcFindLast = True
    Else
        If Not (MyADOdc.Recordset.EOF And MyADOdc.Recordset.BOF) Then
            MyADOdc.Recordset.MoveLast
            MyADOdc.Recordset.MoveNext
        End If
        ADOdcFindLast = False
    End If
        
    MySet.Close
    MyDB.Close
    
    Set MySet = Nothing
    Set MyDB = Nothing
    
    On Local Error GoTo 0
    
End Function


Public Function ADOdcFindPrevious(MyADOdc As Adodc, ByVal FindString As String) As Boolean
  Dim MyDB As ADODB.Connection
  Dim MySet As ADODB.Recordset
  Dim mhMatch As Boolean

    On Local Error Resume Next
    
    Set MyDB = New ADODB.Connection
    Set MySet = New ADODB.Recordset
    
    MyDB.CursorLocation = adUseClient
    MyDB.Open MyADOdc.ConnectionString
    MySet.Open MyADOdc.RecordSource, MyDB, adOpenStatic, adLockPessimistic
    
    If ADOFindPrevious(MySet, FindString) Then
        MyADOdc.Recordset.BookMark = MySet.BookMark
        mhMatch = True
    Else
        If Not (MyADOdc.Recordset.EOF And MyADOdc.Recordset.BOF) Then
            MyADOdc.Recordset.MoveFirst
            MyADOdc.Recordset.MovePrevious
        End If
        mhMatch = False
    End If

    MySet.Close
    MyDB.Close
    Set MySet = Nothing
    Set MyDB = Nothing

    ADOdcFindPrevious = mhMatch
    
End Function


Attribute VB_Name = "modCreateDB"
Option Explicit

' ========================================================
' === Created         : 6/18/2003 10:51:00 AM
' === Access Database : ADDBOOK.MDB
' ========================================================

Private CAT As ADOX.Catalog

Public Sub CreateMDB(ByVal dbPathFilename As String)
On Error GoTo ErrTrap

  Set CAT = New ADOX.Catalog

  CAT.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & dbPathFilename & ";" & _
             "Jet OLEDB:Database Password=" & DB_PWD & ";" & _
             "Jet OLEDB:Engine Type=" & DB_Type & ";"

  Call CreateTables
  Call CreateViews
  Call CreateProcedures
  Call CreateIndexes
  Call CreateKeys

  Set CAT = Nothing

Exit Sub

ErrTrap:
  Exit Sub
  Resume
End Sub

Private Sub CreateTables()
On Error GoTo ErrTrap
Dim TBL As ADOX.Table

' ===[Create Table 'AddBook']===
  Set TBL = New ADOX.Table
  Set TBL.ParentCatalog = CAT
  With TBL
     .Name = "AddBook"
     .Columns.Append "ID", adInteger, 0
     .Columns("ID").Properties("AutoIncrement") = True
     .Columns("ID").Properties("NullAble") = True

     .Columns.Append "Prefix", adVarWChar, 5
     .Columns("Prefix").Properties("NullAble") = True
     .Columns("Prefix").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "FirstName", adVarWChar, 25
     .Columns("FirstName").Properties("NullAble") = True
     .Columns("FirstName").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "MiddleName", adVarWChar, 1
     .Columns("MiddleName").Properties("NullAble") = True
     .Columns("MiddleName").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "LastName", adVarWChar, 25
     .Columns("LastName").Properties("NullAble") = True
     .Columns("LastName").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "Title", adVarWChar, 50
     .Columns("Title").Properties("NullAble") = True
     .Columns("Title").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "Company", adVarWChar, 50
     .Columns("Company").Properties("NullAble") = True
     .Columns("Company").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "HomeAddress", adVarWChar, 50
     .Columns("HomeAddress").Properties("NullAble") = True
     .Columns("HomeAddress").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "HomeCity", adVarWChar, 50
     .Columns("HomeCity").Properties("NullAble") = True
     .Columns("HomeCity").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "HomeState", adVarWChar, 2
     .Columns("HomeState").Properties("NullAble") = True
     .Columns("HomeState").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "HomeZipCode", adVarWChar, 20
     .Columns("HomeZipCode").Properties("NullAble") = True
     .Columns("HomeZipCode").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "HomePhone", adVarWChar, 30
     .Columns("HomePhone").Properties("NullAble") = True
     .Columns("HomePhone").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "MobalPhone", adVarWChar, 30
     .Columns("MobalPhone").Properties("NullAble") = True
     .Columns("MobalPhone").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "WorkPhone", adVarWChar, 30
     .Columns("WorkPhone").Properties("NullAble") = True
     .Columns("WorkPhone").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "WorkAddress", adVarWChar, 50
     .Columns("WorkAddress").Properties("NullAble") = True
     .Columns("WorkAddress").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "WorkCity", adVarWChar, 50
     .Columns("WorkCity").Properties("NullAble") = True
     .Columns("WorkCity").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "WorkState", adVarWChar, 2
     .Columns("WorkState").Properties("NullAble") = True
     .Columns("WorkState").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "WorkZipCode", adVarWChar, 20
     .Columns("WorkZipCode").Properties("NullAble") = True
     .Columns("WorkZipCode").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "FaxNumber", adVarWChar, 30
     .Columns("FaxNumber").Properties("NullAble") = True
     .Columns("FaxNumber").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "User1Phone", adVarWChar, 30
     .Columns("User1Phone").Properties("NullAble") = True
     .Columns("User1Phone").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "User1Text", adVarWChar, 15
     .Columns("User1Text").Properties("NullAble") = True
     .Columns("User1Text").Properties("Jet OLEDB:Allow Zero Length") = True
     .Columns("User1Text").Properties("Default") = Chr(34) & "User1" & Chr(34)

     .Columns.Append "User2Phone", adVarWChar, 30
     .Columns("User2Phone").Properties("NullAble") = True
     .Columns("User2Phone").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "User2Text", adVarWChar, 15
     .Columns("User2Text").Properties("NullAble") = True
     .Columns("User2Text").Properties("Jet OLEDB:Allow Zero Length") = True
     .Columns("User2Text").Properties("Default") = Chr(34) & "User2" & Chr(34)

     .Columns.Append "User3Phone", adVarWChar, 30
     .Columns("User3Phone").Properties("NullAble") = True
     .Columns("User3Phone").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "User3Text", adVarWChar, 15
     .Columns("User3Text").Properties("NullAble") = True
     .Columns("User3Text").Properties("Jet OLEDB:Allow Zero Length") = True
     .Columns("User3Text").Properties("Default") = Chr(34) & "User3" & Chr(34)

     .Columns.Append "Birthdate", adVarWChar, 50
     .Columns("Birthdate").Properties("NullAble") = True
     .Columns("Birthdate").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "Anniversary", adVarWChar, 50
     .Columns("Anniversary").Properties("NullAble") = True
     .Columns("Anniversary").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "Note", adLongVarWChar, 0
     .Columns("Note").Properties("NullAble") = True

     .Columns.Append "Business", adBoolean, 2
     .Columns("Business").Properties("Default") = 0

     .Columns.Append "Relative", adBoolean, 2
     .Columns("Relative").Properties("Default") = 0

     .Columns.Append "Friend", adBoolean, 2
     .Columns("Friend").Properties("Default") = 0

     .Columns.Append "Other", adVarWChar, 10
     .Columns("Other").Properties("NullAble") = True
     .Columns("Other").Properties("Jet OLEDB:Allow Zero Length") = True

     .Columns.Append "EMail", adVarWChar, 50
     .Columns("EMail").Properties("NullAble") = True
     .Columns("EMail").Properties("Jet OLEDB:Allow Zero Length") = True
     .Columns("EMail").Properties("Default") = Chr(34) & "None" & Chr(34)

  End With

  CAT.Tables.Append TBL

' ===[Create Table 'Events']===
  Set TBL = New ADOX.Table
  Set TBL.ParentCatalog = CAT
  With TBL
     .Name = "Events"
     .Columns.Append "EventID", adInteger, 0
     .Columns("EventID").Properties("AutoIncrement") = True
     .Columns("EventID").Properties("NullAble") = True

     .Columns.Append "EventName", adVarWChar, 50
     .Columns("EventName").Properties("NullAble") = True

     .Columns.Append "StartDate", adDate, 0
     .Columns("StartDate").Properties("NullAble") = True
     .Columns("StartDate").Properties("Default") = 0

     .Columns.Append "StartTime", adDate, 0
     .Columns("StartTime").Properties("NullAble") = True
     .Columns("StartTime").Properties("Default") = 0

  End With

  CAT.Tables.Append TBL

' ===[Create Table 'Other']===
  Set TBL = New ADOX.Table
  Set TBL.ParentCatalog = CAT
  With TBL
     .Name = "Other"
     .Columns.Append "Other", adVarWChar, 10
     .Columns("Other").Properties("NullAble") = True
     .Columns("Other").Properties("Jet OLEDB:Allow Zero Length") = True

  End With

  CAT.Tables.Append TBL

  Set TBL = Nothing

Exit Sub

ErrTrap:
  'MsgBox Err.Number & " / " & Err.Description,,"Error In CreateTables"
  'Exit Sub
  'Resume
End Sub

Private Sub CreateViews()
On Error GoTo ErrTrap
Dim CMD As ADODB.Command

Exit Sub
ErrTrap:
  MsgBox Err.Number & " / " & Err.Description, , "Error In CreateViews"
  Exit Sub
  Resume
End Sub

Private Sub CreateProcedures()
On Error GoTo ErrTrap
Dim CMD As ADODB.Command

Exit Sub

ErrTrap:
  'MsgBox Err.Number & " / " & Err.Description,,"Error In CreateProcedures"
  'Exit Sub
  'Resume
End Sub

Private Sub CreateIndexes()
On Error GoTo ErrTrap
Dim IDX As ADOX.Index

' ===[Create Index 'PrimaryKey']===
  Set IDX = New ADOX.Index
  With IDX
     .Name = "PrimaryKey"
     .Columns.Append "ID"
     .PrimaryKey = True
     .Unique = True
     .Clustered = False
     .IndexNulls = adIndexNullsDisallow
  End With
  CAT.Tables("AddBook").Indexes.Append IDX
' ===[Create Index 'PrimaryKey']===
  Set IDX = New ADOX.Index
  With IDX
     .Name = "PrimaryKey"
     .Columns.Append "EventID"
     .PrimaryKey = True
     .Unique = True
     .Clustered = False
     .IndexNulls = adIndexNullsDisallow
  End With
  CAT.Tables("Events").Indexes.Append IDX
' ===[Create Index 'Other']===
  Set IDX = New ADOX.Index
  With IDX
     .Name = "Other"
     .Columns.Append "Other"
     .PrimaryKey = False
     .Unique = True
     .Clustered = False
     .IndexNulls = adIndexNullsAllow
  End With
  CAT.Tables("Other").Indexes.Append IDX
' ===[Create Index 'PrimaryKey']===
  Set IDX = New ADOX.Index
  With IDX
     .Name = "PrimaryKey"
     .Columns.Append "Other"
     .PrimaryKey = True
     .Unique = True
     .Clustered = False
     .IndexNulls = adIndexNullsDisallow
  End With
  CAT.Tables("Other").Indexes.Append IDX

  Set IDX = Nothing

  Exit Sub

ErrTrap:
  'MsgBox Err.Number & " / " & Err.Description,,"Error In CreateIndexes"
  'Exit Sub
  'Resume
End Sub

Private Sub CreateKeys()
On Error GoTo ErrTrap
Dim KEY As ADOX.KEY
Dim TBL As ADOX.Table

  Set KEY = New ADOX.KEY
  Set TBL = New ADOX.Table

  Set KEY = Nothing
  Set TBL = Nothing

  Exit Sub

ErrTrap:
  Select Case Err.Number
  Case -2147467259  ' Index already exists - Remove it...
    CAT.Tables(TBL.Name).Indexes.Delete KEY.Name
    Resume
  Case Else
    MsgBox Err.Number & " / " & Err.Description, , "Error In CreateKeys"
    Exit Sub
    Resume
  End Select
End Sub


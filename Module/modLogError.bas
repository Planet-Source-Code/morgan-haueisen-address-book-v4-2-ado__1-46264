Attribute VB_Name = "modLogError"
Option Explicit
Public Sub Err_Handler( _
    Optional ByVal bDisplayError As Boolean = True, _
    Optional ByVal ErrNumber As String = vbNullString, _
    Optional ByVal ErrDescription As String = vbNullString, _
    Optional ByVal ModuleName As String = vbNullString, _
    Optional ByVal ProcName As String = vbNullString)

  Dim tString As String

    '/* Purpose: Error handling - On Error

    '/* Show Error Message
    If bDisplayError Then
        tString = "Error occured: "
        If Len(ErrNumber) > 0 Then tString = tString & ErrNumber & vbNewLine Else tString = tString & vbNewLine
        If Len(ErrDescription) > 0 Then tString = tString & "Description: " & ErrDescription & vbNewLine
        If Len(ModuleName) > 0 Then tString = tString & "Module: " & ModuleName & vbNewLine
        If Len(ProcName) > 0 Then tString = tString & "Function: " & ProcName
        MsgBox tString, vbCritical, App.Title & " - ERROR"
    End If

    '/* Write error log
    Dim fnum As Integer
    fnum = FreeFile
    Open App.PATH & "\ErrorLog.txt" For Append As #fnum
    Write #fnum, Now, ErrNumber, ErrDescription, ModuleName, ProcName, Environ("username"), Environ("computername")
    Close #fnum

End Sub

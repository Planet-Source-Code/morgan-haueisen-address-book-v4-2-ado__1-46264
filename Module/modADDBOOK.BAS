Attribute VB_Name = "modADDBOOK"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public cValidate As New clsValidate
Public cScreen   As New clsScreenSize

Public Const PhoneProfile As String = "###-###-####"
Public Const ZipProfile   As String = "#####-####"
Public Const DateProfile  As String = "##/##/####"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, _
         ByVal lpFile As String, ByVal lpParameters As String, _
         ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public SortBy    As String
Public SortField As String
Public SortOther As String
Public SortQuery As String
Public PATH      As String

'/* Modem Port Setting */
Public ComPort As Integer

Public PrefixNumber  As String
Public PhoneNumber   As String
Public PnlPhoneFront As Boolean
Public QuitCommand   As Boolean


Public Sub Main()

    On Error GoTo Err_Proc

    If App.PrevInstance Then End
    
    With frmAbout
        .PreventClose = True
        .chameleonButton1.Visible = False
        .Show
    End With
    DoEvents
    
    PATH = App.PATH
    If right(PATH, 1) <> "\" Then PATH = PATH & "\"
    DoEvents

    If Dir(PATH & "ADDBOOK.MDB") = vbNullString Then
        Call CreateMDB(PATH & "ADDBOOK.MDB")
    End If
    
    Call InitSecurity(PATH & "ADDBOOK.MDB", PATH & "ADDBOOK.MDB")
    
    SortBy = GetSetting(App.Title, "Options", "SortBy", "COMPANY")
    SortOther = GetSetting(App.Title, "Options", "SortOther", "")
    SortField = GetSetting(App.Title, "Options", "SortField", "")
    
    MainMenu.Show

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "modADDBOOK", "Main"
    Err.Clear
    Resume Exit_Proc

End Sub

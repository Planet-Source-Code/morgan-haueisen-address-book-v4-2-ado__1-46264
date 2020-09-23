VERSION 5.00
Begin VB.Form frmSearchFor 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search For ?"
   ClientHeight    =   2535
   ClientLeft      =   3705
   ClientTop       =   4740
   ClientWidth     =   5280
   ForeColor       =   &H80000008&
   Icon            =   "FRMSEARC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2535
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5280
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   465
         TabIndex        =   7
         Top             =   660
         Width           =   4290
      End
      Begin VB.PictureBox Picture1 
         Height          =   645
         Index           =   0
         Left            =   510
         ScaleHeight     =   585
         ScaleWidth      =   1575
         TabIndex        =   4
         Top             =   1110
         Width           =   1635
         Begin VB.OptionButton opt_lookfor 
            Caption         =   "Company"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   6
            Top             =   60
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.OptionButton opt_lookfor 
            Caption         =   "Last Name"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   5
            Top             =   345
            Width           =   1515
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   645
         Index           =   1
         Left            =   2355
         ScaleHeight     =   585
         ScaleWidth      =   1575
         TabIndex        =   1
         Top             =   1110
         Width           =   1635
         Begin VB.OptionButton Option3D1 
            Caption         =   "Begining with"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   3
            Top             =   345
            Width           =   1515
         End
         Begin VB.OptionButton Option3D1 
            Caption         =   "Any"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   2
            Top             =   60
            Value           =   -1  'True
            Width           =   1515
         End
      End
      Begin AddressBook.chameleonButton cmd_ok 
         Height          =   555
         Left            =   3735
         TabIndex        =   9
         Top             =   1890
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Ok"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "FRMSEARC.frx":000C
         PICN            =   "FRMSEARC.frx":0028
         PICH            =   "FRMSEARC.frx":04EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_quit 
         Height          =   555
         Left            =   2700
         TabIndex        =   10
         Top             =   1890
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "FRMSEARC.frx":09B4
         PICN            =   "FRMSEARC.frx":09D0
         PICH            =   "FRMSEARC.frx":0E96
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   255
         TabIndex        =   8
         Top             =   165
         Width           =   4515
      End
   End
End
Attribute VB_Name = "frmSearchFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_SortByName As String

Private Const SHOWCOMPANY As String = "Please Enter Part or All of the COMPANY NAME you wish to find"
Private Const SHOWLAST    As String = "Please Enter Part or All of the LAST NAME you wish to find"

Private Sub cmd_ok_Click()

    On Error GoTo Err_Proc

 Dim tString As String
 Dim SQLstmt As String
 
    tString = Text1.Text
    If tString = vbNullString Then
        Me.Hide
        Exit Sub
    End If
    
    tString = UCase(tString)
    
    If Option3D1(1).Value Then
       SQLstmt = "[" & m_SortByName$ & "] LIKE '" & tString & "%'"
    Else
       SQLstmt = "[" & m_SortByName$ & "] LIKE '" & "%" & tString & "%'"
    End If
    
    If Not ADOdcFindNext(MainMenu!Data1, SQLstmt) Then
        ADOdcFindFirst MainMenu!Data1, SQLstmt
    End If
    Me.Hide


Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmSearchFor", "cmd_ok_Click"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub cmd_quit_Click()
 Me.Hide
End Sub

Private Sub Form_Activate()
 
 If Text1.Text > vbNullString Then
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SetFocus
 End If

End Sub

Private Sub Form_Load()

    On Error GoTo Err_Proc

    cScreen.CenterForm Me
    
    If UCase(SortBy) = "COMPANY" Then
       m_SortByName$ = "Company"
       Label1.Caption = SHOWCOMPANY$
       opt_lookfor(0).Value = True
    Else
       m_SortByName$ = "LastName"
       Label1.Caption = SHOWLAST$
       opt_lookfor(1).Value = True
    End If


Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmSearchFor", "Form_Load"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSearchFor = Nothing
End Sub


Private Sub opt_lookfor_Click(Index As Integer)
    On Local Error Resume Next
    Select Case Index
    Case 0
       m_SortByName$ = "Company"
       Label1.Caption = SHOWCOMPANY$
    Case 1
       m_SortByName$ = "Lastname"
       Label1.Caption = SHOWLAST$
    End Select
    Text1.SetFocus
End Sub

Private Sub Option3D1_Click(Index As Integer)
 Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then KeyAscii = False: cmd_ok_Click
 If KeyAscii = 27 Then KeyAscii = False: Me.Hide
End Sub


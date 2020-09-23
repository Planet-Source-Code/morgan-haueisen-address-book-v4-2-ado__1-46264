VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditOther 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "c"
   ClientHeight    =   2625
   ClientLeft      =   2910
   ClientTop       =   2250
   ClientWidth     =   5565
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "EDITOTH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5550
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Other"
         DataSource      =   "Data1"
         Height          =   300
         Left            =   585
         MaxLength       =   10
         TabIndex        =   1
         Top             =   855
         Width           =   2970
      End
      Begin AddressBook.chameleonButton cmd_new 
         Height          =   540
         Left            =   4350
         TabIndex        =   4
         Top             =   255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   953
         BTYPE           =   3
         TX              =   "Add New"
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
         MICON           =   "EDITOTH.frx":000C
         PICN            =   "EDITOTH.frx":0028
         PICH            =   "EDITOTH.frx":04EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmdEdit 
         Height          =   540
         Left            =   4350
         TabIndex        =   5
         Top             =   795
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   953
         BTYPE           =   3
         TX              =   "Edit"
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
         MICON           =   "EDITOTH.frx":09B4
         PICN            =   "EDITOTH.frx":09D0
         PICH            =   "EDITOTH.frx":0DA6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_ok 
         Height          =   540
         Left            =   4350
         TabIndex        =   6
         Top             =   1335
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   953
         BTYPE           =   3
         TX              =   "Save"
         ENAB            =   0   'False
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
         MICON           =   "EDITOTH.frx":117C
         PICN            =   "EDITOTH.frx":1198
         PICH            =   "EDITOTH.frx":169A
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
         Height          =   540
         Left            =   4350
         TabIndex        =   3
         Top             =   1875
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   953
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
         MICON           =   "EDITOTH.frx":1B9C
         PICN            =   "EDITOTH.frx":1BB8
         PICH            =   "EDITOTH.frx":207E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   390
         Left            =   870
         Top             =   1440
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   688
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Other"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   4140
         X2              =   4140
         Y1              =   105
         Y2              =   2565
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   4155
         X2              =   4155
         Y1              =   120
         Y2              =   2580
      End
      Begin VB.Label Label1 
         Caption         =   "Edit Other Listing"
         Height          =   300
         Left            =   555
         TabIndex        =   2
         Top             =   450
         Width           =   3105
      End
   End
End
Attribute VB_Name = "frmEditOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_new_Click()
    Data1.Recordset.AddNew
    Data1.Enabled = False
    cmd_new.Enabled = False
    cmdEdit.Enabled = False
    cmd_ok.Enabled = True
End Sub

Private Sub cmd_ok_Click()
    On Local Error GoTo SaveError
    cmdEdit.Enabled = True
    cmd_new.Enabled = True
    cmd_ok.Enabled = False
    Data1.Recordset.Update
    Data1.Enabled = True
Exit Sub

SaveError:
    MsgBox Error$ & vbCrLf & "This name already exists"
    Resume Next
   
End Sub

Private Sub cmd_quit_Click()

    On Error GoTo Err_Proc

    If cmdEdit.Enabled = False Then
        Data1.Recordset.CancelUpdate
    End If
    
    MainMenu!cboOther.Clear
    Data1.Recordset.MoveFirst
    Do
       MainMenu!cboOther.AddItem Data1.Recordset("Other")
       Data1.Recordset.MoveNext
    Loop Until Data1.Recordset.EOF
    MainMenu!cboOther.AddItem ""
    Me.Hide

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmEditOther", "cmd_quit_Click"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub cmdEdit_Click()
    'Data1.Recordset.Edit
    cmdEdit.Enabled = False
    cmd_new.Enabled = False
    cmd_ok.Enabled = True
End Sub

Private Sub Form_Load()
    
    cScreen.CenterForm Me
    On Local Error Resume Next
    ADOdcConnect Data1, "SELECT * FROM Other ORDER BY [Other]"
    Data1.Refresh
    On Local Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEditOther = Nothing
End Sub


Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    cValidate.FcaseTextBox Text1
End Sub



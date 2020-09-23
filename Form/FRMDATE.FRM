VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dates"
   ClientHeight    =   5040
   ClientLeft      =   2100
   ClientTop       =   1605
   ClientWidth     =   8430
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "FRMDATE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5040
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolBar 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   8370
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4230
      Width           =   8430
      Begin AddressBook.chameleonButton cmd_print 
         Height          =   660
         Left            =   510
         TabIndex        =   1
         Top             =   45
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1164
         BTYPE           =   3
         TX              =   "Print"
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
         MICON           =   "FRMDATE.frx":000C
         PICN            =   "FRMDATE.frx":0028
         PICH            =   "FRMDATE.frx":04EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_exit 
         Height          =   660
         Left            =   1575
         TabIndex        =   2
         Top             =   45
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1164
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
         MICON           =   "FRMDATE.frx":09B4
         PICN            =   "FRMDATE.frx":09D0
         PICH            =   "FRMDATE.frx":0E96
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSDataGridLib.DataGrid DBGrid1 
      Align           =   1  'Align Top
      Bindings        =   "FRMDATE.frx":135C
      Height          =   3555
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   6271
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Dates"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "FIRSTNAME"
         Caption         =   "FIRSTNAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "LASTNAME"
         Caption         =   "LASTNAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "BIRTHDATE"
         Caption         =   "BIRTHDATE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ANNIVERSARY"
         Caption         =   "ANNIVERSARY"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   390
      Left            =   180
      Top             =   3615
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
      CommandType     =   1
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
      Caption         =   "Events"
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
End
Attribute VB_Name = "frmDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_print_Click()

    On Error GoTo Err_Proc

 Dim X As Integer
 
 Screen.MousePointer = 11
 DoEvents
 Printer.FontName = "Times New Roman"
 Printer.FontSize = 11
 DBGrid1.Row = 0
 DBGrid1.Col = 0
 Printer.Print Tab(10); DBGrid1.Text;
 DBGrid1.Col = 1
 Printer.Print Tab(55); DBGrid1.Text;
 DBGrid1.Col = 2
 Printer.Print Tab(75); DBGrid1.Text
 Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
 Printer.Print
 
 For X = 1 To Adodc1.Recordset.RecordCount - 1
    DBGrid1.Row = X
    DBGrid1.Col = 0
    Printer.Print Tab(10); DBGrid1.Text;
    DBGrid1.Col = 1
    Printer.Print Tab(55); DBGrid1.Text;
    DBGrid1.Col = 2
    Printer.Print Tab(75); DBGrid1.Text
    Printer.Print
 Next X
 
 Printer.EndDoc
 Screen.MousePointer = 0
 

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmDates", "cmd_print_Click"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub cmd_exit_Click()
 Me.Hide
End Sub

Private Sub Form_Load()
    Me.Move MainMenu.left, MainMenu.top, MainMenu.Width, MainMenu.Height
    DoEvents
    
    ADOdcConnect Adodc1, "SELECT [FIRSTNAME],[LASTNAME],[BIRTHDATE],[ANNIVERSARY] FROM AddBook ORDER BY [LASTNAME],[FIRSTNAME]"
    Adodc1.Refresh
    DoEvents
 
End Sub

Private Sub Form_Resize()
    DBGrid1.Height = Me.Height - picToolBar.Height - 400
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmDates = Nothing
End Sub



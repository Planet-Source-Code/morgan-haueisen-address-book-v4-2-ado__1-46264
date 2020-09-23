VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAppointments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Appointments"
   ClientHeight    =   5070
   ClientLeft      =   1425
   ClientTop       =   1185
   ClientWidth     =   8595
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "FRMAPPOI.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5070
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolBar 
      Align           =   2  'Align Bottom
      Height          =   885
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   8535
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4185
      Width           =   8595
      Begin AddressBook.chameleonButton cmd_print 
         Height          =   660
         Left            =   90
         TabIndex        =   1
         Top             =   90
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
         MICON           =   "FRMAPPOI.frx":000C
         PICN            =   "FRMAPPOI.frx":0028
         PICH            =   "FRMAPPOI.frx":04EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_edit 
         Height          =   660
         Left            =   1125
         TabIndex        =   2
         Top             =   90
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1164
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
         MICON           =   "FRMAPPOI.frx":09B4
         PICN            =   "FRMAPPOI.frx":09D0
         PICH            =   "FRMAPPOI.frx":0DA6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_add 
         Height          =   660
         Left            =   2160
         TabIndex        =   3
         Top             =   90
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1164
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
         MICON           =   "FRMAPPOI.frx":117C
         PICN            =   "FRMAPPOI.frx":1198
         PICH            =   "FRMAPPOI.frx":165E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_delete 
         Height          =   660
         Left            =   3195
         TabIndex        =   4
         Top             =   90
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1164
         BTYPE           =   3
         TX              =   "Delete"
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
         MICON           =   "FRMAPPOI.frx":1B24
         PICN            =   "FRMAPPOI.frx":1B40
         PICH            =   "FRMAPPOI.frx":2006
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
         Left            =   4230
         TabIndex        =   5
         Top             =   90
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
         MICON           =   "FRMAPPOI.frx":24CC
         PICN            =   "FRMAPPOI.frx":24E8
         PICH            =   "FRMAPPOI.frx":29AE
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
   Begin MSDataGridLib.DataGrid Grid 
      Align           =   1  'Align Top
      Bindings        =   "FRMAPPOI.frx":2E74
      Height          =   3555
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
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
      Caption         =   "Appointments"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "EventID"
         Caption         =   "EventID"
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
         DataField       =   "EventName"
         Caption         =   "EventName"
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
         DataField       =   "StartDate"
         Caption         =   "StartDate"
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
         DataField       =   "StartTime"
         Caption         =   "StartTime"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
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
      Left            =   165
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
Attribute VB_Name = "frmAppointments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_add_Click()

    On Error GoTo Err_Proc

  Dim uID As Long
  
    Adodc1.Recordset.AddNew
    frmEditApp.Show vbModal, Me
    Unload frmEditApp
    uID = Adodc1.Recordset!Eventid
    Adodc1.Recordset.Requery
    ADOdcFindFirst Adodc1, "[EventID]=" & CStr(uID)


Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmAppointments", "cmd_add_Click"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub cmd_delete_Click()
    On Local Error Resume Next
    Adodc1.Recordset.Delete
End Sub

Private Sub cmd_edit_Click()
    On Local Error Resume Next
    frmEditApp!txt_desc.Text = Grid.Columns(1)
    frmEditApp!txt_date.Text = Grid.Columns(2)
    frmEditApp!txt_time.Text = Grid.Columns(3)
    
    ADOdcFindFirst Adodc1, "[EVENTID] = " & Grid.Columns(0)
    
    'Adodc1.Recordset.Edit
    frmEditApp.Show vbModal
    Unload frmEditApp

End Sub

Private Sub cmd_exit_Click()
 Me.Hide
End Sub

Private Sub cmd_print_Click()
' Screen.MousePointer = 11
' DoEvents
' Printer.FontName = "Times New Roman"
' Printer.FontSize = 11
' Grid.Row = 0
' Grid.Col = 0
' Printer.Print Tab(10); Grid.Text;
' Grid.Col = 1
' Printer.Print Tab(85); Grid.Text;
' Grid.Col = 2
' Printer.Print Tab(105); Grid.Text
' Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
' Printer.Print
'
' For x% = 1 To Grid.Rows - 1
'    Grid.Row = x%
'    Grid.Col = 0
'    Printer.Print Tab(10); Grid.Text;
'    Grid.Col = 1
'    Printer.Print Tab(85); Grid.Text;
'    Grid.Col = 2
'    Printer.Print Tab(105); Grid.Text
'    Printer.Print
' Next x%
' Printer.NewPage
' Printer.EndDoc
' Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
 
 Me.Move MainMenu.left, MainMenu.top, MainMenu.Width, MainMenu.Height
    
 DoEvents
 
 ADOdcConnect Adodc1, "SELECT * FROM EVENTS ORDER BY [STARTDATE],[STARTTIME]"
 Adodc1.Refresh
 DoEvents
 
End Sub

Private Sub Form_Resize()
    Grid.Height = Me.Height - picToolBar.Height - 400
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmAppointments = Nothing
End Sub



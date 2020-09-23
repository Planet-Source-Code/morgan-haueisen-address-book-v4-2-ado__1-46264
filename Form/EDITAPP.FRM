VERSION 5.00
Begin VB.Form frmEditApp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1785
   ClientLeft      =   1575
   ClientTop       =   1860
   ClientWidth     =   8535
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "EDITAPP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1785
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin AddressBook.chameleonButton cmdGetDate 
         Height          =   390
         Left            =   2850
         TabIndex        =   9
         ToolTipText     =   "Calendar"
         Top             =   1065
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "EDITAPP.frx":000C
         PICN            =   "EDITAPP.frx":0028
         PICH            =   "EDITAPP.frx":042B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txt_time 
         Height          =   285
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1305
         Width           =   1140
      End
      Begin VB.TextBox txt_date 
         Height          =   285
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   930
         Width           =   1140
      End
      Begin VB.TextBox txt_desc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   1
         Top             =   450
         Width           =   6555
      End
      Begin AddressBook.chameleonButton cmd_ok 
         Height          =   660
         Left            =   6900
         TabIndex        =   7
         Top             =   945
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1164
         BTYPE           =   3
         TX              =   "Save"
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
         MICON           =   "EDITAPP.frx":08F1
         PICN            =   "EDITAPP.frx":090D
         PICH            =   "EDITAPP.frx":0E0F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_quit 
         Height          =   660
         Left            =   5850
         TabIndex        =   8
         Top             =   945
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1164
         BTYPE           =   3
         TX              =   "Cancel"
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
         MICON           =   "EDITAPP.frx":1311
         PICN            =   "EDITAPP.frx":132D
         PICH            =   "EDITAPP.frx":17F3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Description: "
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   4
         Top             =   495
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Date: "
         Height          =   195
         Index           =   1
         Left            =   1110
         TabIndex        =   3
         Top             =   945
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Time: "
         Height          =   195
         Index           =   2
         Left            =   1110
         TabIndex        =   2
         Top             =   1350
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmEditApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_ok_Click()
 
 On Local Error GoTo ShowError

 frmAppointments!Adodc1.Recordset("EVENTNAME") = txt_desc.Text
 frmAppointments!Adodc1.Recordset("StartDate") = txt_date.Text
 frmAppointments!Adodc1.Recordset("StartTime") = txt_time.Text
 frmAppointments!Adodc1.Recordset.Update
 
' frmAppointments!Grid.Col = 0
' frmAppointments!Grid.Text = txt_desc.Text
' frmAppointments!Grid.Col = 1
' frmAppointments!Grid.Text = txt_date.Text
' frmAppointments!Grid.Col = 2
' frmAppointments!Grid.Text = txt_time.Text
 Me.Hide

ExitSub:
Exit Sub

ShowError:
 MsgBox "You must fill in both the date and time" & Chr(10) & "Date Format: MM/DD/YY   Time Format: HH:MM AM", 0, "ERROR"
 Resume ExitSub

End Sub

Private Sub cmd_quit_Click()
 Me.Hide
End Sub

Private Sub cmdGetDate_Click()
    On Error GoTo Err_Proc
  Dim tDate As Date
  
    With frmCalendar
        If IsDate(txt_date) Then
            tDate = CDate(txt_date & " " & txt_time)
            .PassDate = tDate
        End If
        .cfFirstWeekDay = vbSunday
        .ShowYear = True
        .ShowTime = True
        .ShowWeeks = False
        .Command1.Visible = False
        .Show vbModal
        txt_date = Format(.PassDate, "mm/dd/yy")
        txt_time = Format(.PassDate, "hh:mm AMPM")
    End With


Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmEditApp", "cmdGetDate_Click"
    Err.Clear
    Resume Exit_Proc
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    Me.Move MainMenu.left, MainMenu.top + 600, MainMenu.Width
End Sub

Private Sub Form_Resize()
    Frame1.Width = Me.Width - 120
End Sub


Private Sub txt_date_KeyUp(KeyCode As Integer, Shift As Integer)
    'cValidate.MaskedText txt_date, "##/##/##"
End Sub


Private Sub txt_desc_Change()
    cValidate.FcaseTextBox txt_desc
End Sub


Private Sub txt_time_KeyUp(KeyCode As Integer, Shift As Integer)
    'cValidate.MaskedText txt_time, "##:## @@"
End Sub



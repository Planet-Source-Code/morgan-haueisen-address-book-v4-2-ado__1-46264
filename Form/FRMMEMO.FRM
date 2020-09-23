VERSION 5.00
Begin VB.Form frmMemo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5265
   ClientLeft      =   2325
   ClientTop       =   3600
   ClientWidth     =   6000
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "FRMMEMO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5265
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   4110
         Left            =   75
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   135
         Width           =   5805
      End
   End
   Begin AddressBook.chameleonButton cmd_exit 
      Height          =   660
      Left            =   4605
      TabIndex        =   2
      Top             =   4470
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
      MICON           =   "FRMMEMO.frx":000C
      PICN            =   "FRMMEMO.frx":0028
      PICH            =   "FRMMEMO.frx":04EE
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
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_exit_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    cScreen.CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMemo = Nothing
End Sub


Private Sub Text1_Change()
    If MainMenu!chk_autocap.Value Then cValidate.FcaseTextBox Text1
End Sub


VERSION 5.00
Begin VB.Form frmSelectOther 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1395
   ClientLeft      =   3495
   ClientTop       =   3945
   ClientWidth     =   3900
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "FRMSELEC.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1395
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3885
      Begin VB.PictureBox picOpt 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   45
         ScaleHeight     =   570
         ScaleWidth      =   3780
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Width           =   3780
         Begin VB.OptionButton opt_last 
            Caption         =   "Order By Last Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   5
            Top             =   330
            Width           =   2475
         End
         Begin VB.OptionButton opt_company 
            Caption         =   "Order By Company"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   4
            Top             =   0
            Value           =   -1  'True
            Width           =   2475
         End
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   855
         Width           =   2085
      End
      Begin AddressBook.chameleonButton cmd_ok 
         Height          =   555
         Left            =   2730
         TabIndex        =   2
         Top             =   705
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
         MICON           =   "FRMSELEC.frx":000C
         PICN            =   "FRMSELEC.frx":0028
         PICH            =   "FRMSELEC.frx":04EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "frmSelectOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_ok_Click()
 Me.Hide
End Sub

Private Sub Form_Load()
  Dim i As Integer
  
    cScreen.CenterForm Me
    'cmd_ok.Picture = frmArt!art(0).Picture
    For i = 0 To MainMenu.cboOther.ListCount - 1
       Combo1.AddItem MainMenu!cboOther.List(i)
    Next i
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSelectOther = Nothing
End Sub



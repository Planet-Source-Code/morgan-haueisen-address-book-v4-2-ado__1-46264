VERSION 5.00
Begin VB.Form printing 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3645
   ClientLeft      =   2730
   ClientTop       =   2535
   ClientWidth     =   6690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "cfPrinting.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3645
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProgress 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3375
      Width           =   6690
      Begin VB.PictureBox picProgressSlide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   0
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.TextBox Label2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2280
      Left            =   1755
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "cfPrinting.frx":000C
      Top             =   660
      Width           =   4725
   End
   Begin VB.TextBox Label1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   1710
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "cfPrinting.frx":0025
      Top             =   15
      Width           =   4815
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   405
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   2715
      Width           =   900
   End
   Begin VB.Image pCursor 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   0
      Picture         =   "cfPrinting.frx":003D
      Top             =   3705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   3450
      Left            =   0
      Picture         =   "cfPrinting.frx":0347
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1650
   End
End
Attribute VB_Name = "printing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdQuit_Click()
    QuitCommand = True
End Sub
Private Sub Form_Activate()
    Label1.Refresh
    DoEvents
End Sub
Private Sub Form_Load()
    cScreen.CenterForm Me
    Screen.MouseIcon = pCursor.Picture
    Screen.MousePointer = vbCustom
    DoEvents
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub
Private Sub Label2_Change()
    Me.ZOrder
End Sub



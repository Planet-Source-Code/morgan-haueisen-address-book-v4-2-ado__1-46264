VERSION 5.00
Begin VB.Form frmCommSettings 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Communication Settings"
   ClientHeight    =   2865
   ClientLeft      =   4890
   ClientTop       =   3600
   ClientWidth     =   2550
   ForeColor       =   &H80000008&
   Icon            =   "COMMPORT.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame7 
      Caption         =   "Dial Prefix"
      Height          =   585
      Left            =   345
      TabIndex        =   1
      Top             =   240
      Width           =   1155
      Begin VB.TextBox txtPrefix 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Text            =   "1"
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Com Port"
      Height          =   885
      Left            =   330
      TabIndex        =   0
      Top             =   930
      Width           =   1860
      Begin VB.ComboBox cboComPort 
         Height          =   315
         Left            =   105
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   330
         Width           =   915
      End
      Begin AddressBook.chameleonButton Command1 
         Height          =   450
         Left            =   1155
         TabIndex        =   6
         Top             =   285
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   794
         BTYPE           =   14
         TX              =   "Auto Find"
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
         MICON           =   "COMMPORT.frx":000C
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
   Begin AddressBook.chameleonButton cmd_ok 
      Height          =   660
      Left            =   1245
      TabIndex        =   4
      Top             =   1980
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
      MICON           =   "COMMPORT.frx":0028
      PICN            =   "COMMPORT.frx":0044
      PICH            =   "COMMPORT.frx":0546
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
      Left            =   195
      TabIndex        =   5
      Top             =   1980
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
      MICON           =   "COMMPORT.frx":0A48
      PICN            =   "COMMPORT.frx":0A64
      PICH            =   "COMMPORT.frx":0F2A
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
Attribute VB_Name = "frmCommSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ok_Click()
 
    ComPort = cboComPort
    PrefixNumber = txtPrefix.Text
    
    SaveSetting App.Title, "Options", "ComPort", ComPort
    SaveSetting App.Title, "Options", "PrefixNumber", PrefixNumber
    
    Unload Me

End Sub

Private Sub cmd_quit_Click()
    Me.Hide
End Sub

Private Sub Command1_Click()
  Dim i As Byte
  
    On Error Resume Next

    frmDIALER.MSComm1.PortOpen = False
    
    For i = 1 To 16
        frmDIALER.MSComm1.CommPort = i
        If frmDIALER.MSComm1.PortOpen = False Then
            frmDIALER.MSComm1.PortOpen = True
            frmDIALER.MSComm1.Output = "AT" & vbCr '/* Sends "attention" command To the modem
            DoEvents
            If frmDIALER.MSComm1.InBufferCount >= 2 Then Exit For '/* once "OK" is read, exit Loop
            frmDIALER.MSComm1.PortOpen = False
        End If
    Next i

    If i = 17 Then
        MsgBox "Modem Not detected or In use.", vbInformation
    Else
        MsgBox "Your modem comm port is " & CStr(i)
        ComPort = CStr(i)
        SaveSetting App.Title, "Options", "ComPort", ComPort
    End If
    
    frmDIALER.MSComm1.PortOpen = False

End Sub

Private Sub Form_Load()
    On Error GoTo Err_Proc
  Dim i As Integer
    cScreen.CenterForm Me
    
    For i = 1 To 16
        cboComPort.AddItem CStr(i)
    Next i
    cboComPort = ComPort
        
    If PrefixNumber > vbNullString Then txtPrefix.Text = PrefixNumber


Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmCommSettings", "Form_Load"
    Err.Clear
    Resume Exit_Proc
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCommSettings = Nothing
End Sub


Private Sub txtPrefix_KeyPress(KeyAscii As Integer)
    cValidate.IsNumber txtPrefix, KeyAscii, False
End Sub


VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDIALER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Phone Dialer"
   ClientHeight    =   2955
   ClientLeft      =   4260
   ClientTop       =   2625
   ClientWidth     =   3165
   ForeColor       =   &H80000008&
   Icon            =   "DIALER.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2955
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   2070
      Width           =   3135
      Begin AddressBook.chameleonButton DialButton 
         Height          =   330
         Left            =   540
         TabIndex        =   10
         Top             =   120
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "Dial"
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "DIALER.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton CancelButton 
         Height          =   330
         Left            =   1230
         TabIndex        =   11
         Top             =   120
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   582
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "DIALER.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton QuitButton 
         Height          =   330
         Left            =   1920
         TabIndex        =   12
         Top             =   120
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   582
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "DIALER.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton Command1 
         Height          =   330
         Left            =   540
         TabIndex        =   13
         Top             =   495
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "Modem SetUp"
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "DIALER.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label DialNow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   60
         TabIndex        =   9
         Top             =   555
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2070
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.OptionButton Opt 
         Caption         =   "Dial Business Phone"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   5
         Top             =   255
         Value           =   -1  'True
         Width           =   2370
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Dial Home Phone"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   4
         Top             =   505
         Width           =   2370
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Dial Fax Phone"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   3
         Top             =   755
         Width           =   2370
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Dial Mobal Phone"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   2
         Top             =   1005
         Width           =   2370
      End
      Begin VB.TextBox txtPrefix 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   255
         TabIndex        =   1
         Text            =   "91"
         Top             =   1650
         Width           =   405
      End
      Begin VB.Label Status 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select a number to call"
         Height          =   285
         Left            =   75
         TabIndex        =   7
         Top             =   1275
         Width           =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Add prefix to phone number "
         Height          =   195
         Left            =   720
         TabIndex        =   6
         Top             =   1710
         Width           =   2040
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmDIALER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CancelFlag As Boolean

Private Sub CancelButton_Click()
    
    '-- CancelFlag tells the Dial routine to exit
    m_CancelFlag = True
    CancelButton.Enabled = False

End Sub

Private Sub Command1_Click()
 frmCommSettings.Show vbModal
 txtPrefix.Text = PrefixNumber
 DoEvents
End Sub

Private Sub Dial()
  Dim FromModem As String
  
    '-- Open the comm port
    On Local Error Resume Next
    MSComm1.PortOpen = False
    Err.Clear
    
    ComPort = Val(GetSetting(App.Title, "Options", "ComPort", "1"))
    MSComm1.CommPort = ComPort
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.PortOpen = True

    If Err.Number Then
       MsgBox Err.Description & "  COM" & CStr(ComPort) & ":" & vbCrLf & "Change the ComPort property to another port."
       Exit Sub
    End If
    
    '-- Flush the input/output buffer
    MSComm1.InBufferCount = 0
    MSComm1.OutBufferCount = 0
    DoEvents

    '--- AT is the Hayes compatible ATTENTION command an is required to send commands to the modem.
    '--- DT means "Dial Tone" - The Dial command, using touch tones as opposed to pulse (DP = Dial Pulse)
    '--- A semicolon tells the modem to return to command mode after dialing (important)
    '--- A Carriage return, vbcr, is required when sending commands to the modem.
    '-- Dial the number
    MSComm1.Output = "ATDT" & PhoneNumber & ";" & vbCr 'DialString$
    DoEvents
    Sleep 2000
    
    '-- Wait for "OK" to come back from the modem
    Do
       DoEvents
       '-- If there is data in the buffer, then read it.
       If MSComm1.InBufferCount >= 2 Then
          FromModem$ = FromModem$ & MSComm1.Input
          '-- Check for "OK"
          If InStr(FromModem$, "OK") Then
             '-- Notify the user to pick up the phone
             Beep
             MsgBox "Please pick up the phone and either press Enter, or click OK"
             Exit Do
          End If
       End If
       
       '-- Was Cancel pressed?
       If m_CancelFlag Then
          m_CancelFlag = False
          Exit Do
       End If
    Loop
    
    '-- Disconnect the modem
    MSComm1.Output = "ATH" & ";" & vbCr
    DoEvents

    '-- Close the port
    Do: DoEvents: Loop Until MSComm1.OutBufferCount = 0
    
    MSComm1.PortOpen = False
    Me.Hide
    
End Sub

Private Sub DialButton_click()
  Dim Temp$, m$

'    On Error Resume Next
    DialButton.Enabled = False
    QuitButton.Enabled = False
    Command1.Enabled = False
    Frame1.Visible = False
    Frame2.Visible = True
    CancelButton.Enabled = True
    CancelButton.SetFocus
    
    '-- Get the number to dial
    If txtPrefix.Text > vbNullString Then m$ = "-"
    PhoneNumber = txtPrefix.Text & m$ & PhoneNumber
    
    Temp$ = Status
    Status = "Dialing - " + PhoneNumber
    
    '-- Dial the selected phone number
    Call Dial

    Frame2.Visible = False
    DialButton.Enabled = True
    QuitButton.Enabled = True
    Command1.Enabled = True
    Frame1.Visible = True
    CancelButton.Enabled = False

    Status = Temp$

End Sub

Private Sub Form_Activate()

    On Error GoTo Err_Proc

 If PhoneNumber > vbNullString Then
    DialButton_click
    Exit Sub
 Else
 'If PhoneNumber = vbnullstring Then
     If Opt(0).Value Then PhoneNumber = MainMenu!txt_bphone.Text
     If Opt(1).Value Then PhoneNumber = MainMenu!txt_hphone.Text
     If Opt(2).Value Then PhoneNumber = MainMenu!txt_fphone.Text
     If Opt(3).Value Then PhoneNumber = MainMenu!txt_mphone.Text
 End If
 Status = "Selected " & txtPrefix.Text & "-" & PhoneNumber


Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmDIALER", "Form_Activate"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 And QuitButton.Enabled Then QuitButton_click
End Sub

Private Sub Form_Load()
    cScreen.CenterForm Me
    
    ComPort = GetSetting(App.Title, "Options", "ComPort", "1")
    PrefixNumber = GetSetting(App.Title, "Options", "PrefixNumber", "1")
    
    '-- Setting InputLen to 0 tells MSComm to read the entire
    '   contents of the input buffer when the Input property
    '   is used.
    MSComm1.InputLen = 0
    
    txtPrefix.Text = PrefixNumber
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDIALER = Nothing
End Sub

Private Sub Opt_Click(Index As Integer)
    '-- Get the number to dial
    If Opt(0).Value Then PhoneNumber = MainMenu!txt_bphone.Text
    If Opt(1).Value Then PhoneNumber = MainMenu!txt_hphone.Text
    If Opt(2).Value Then PhoneNumber = MainMenu!txt_fphone.Text
    If Opt(3).Value Then PhoneNumber = MainMenu!txt_mphone.Text
    Status = "Selected " & txtPrefix.Text & "-" & PhoneNumber
End Sub

Private Sub txtPrefix_Change()
    PrefixNumber = txtPrefix.Text
    Status = "Selected " & txtPrefix.Text & "-" & PhoneNumber
End Sub

Private Sub QuitButton_click()
    Me.Hide
End Sub

Private Sub txtPrefix_KeyPress(KeyAscii As Integer)
    cValidate.IsNumber txtPrefix, KeyAscii, False
End Sub



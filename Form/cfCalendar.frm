VERSION 5.00
Begin VB.Form frmCalendar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Date"
   ClientHeight    =   5475
   ClientLeft      =   3240
   ClientTop       =   1515
   ClientWidth     =   4890
   Icon            =   "cfCalendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   4890
   Begin VB.CommandButton cmdYRPrevNext 
      Height          =   435
      Index           =   0
      Left            =   2655
      Picture         =   "cfCalendar.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   30
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdYRPrevNext 
      Height          =   435
      Index           =   1
      Left            =   4410
      Picture         =   "cfCalendar.frx":03C4
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   30
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdMonPrevNext 
      Height          =   435
      Index           =   1
      Left            =   2130
      Picture         =   "cfCalendar.frx":047E
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   30
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdMonPrevNext 
      Height          =   435
      Index           =   0
      Left            =   135
      Picture         =   "cfCalendar.frx":0538
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   30
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   4785
      TabIndex        =   15
      Top             =   495
      Width           =   4845
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mon"
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   71
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tue"
         Height          =   255
         Index           =   2
         Left            =   1620
         TabIndex        =   70
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wed"
         Height          =   255
         Index           =   3
         Left            =   2220
         TabIndex        =   69
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thu"
         Height          =   255
         Index           =   4
         Left            =   2820
         TabIndex        =   68
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fri"
         Height          =   255
         Index           =   5
         Left            =   3420
         TabIndex        =   67
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sat"
         Height          =   255
         Index           =   6
         Left            =   4020
         TabIndex        =   66
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sun"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   65
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   0
         Left            =   420
         TabIndex        =   64
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   1
         Left            =   1020
         TabIndex        =   63
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   2
         Left            =   1620
         TabIndex        =   62
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   3
         Left            =   2220
         TabIndex        =   61
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   4
         Left            =   2820
         TabIndex        =   60
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   5
         Left            =   3420
         TabIndex        =   59
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   6
         Left            =   4020
         TabIndex        =   58
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   7
         Left            =   420
         TabIndex        =   57
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   8
         Left            =   1020
         TabIndex        =   56
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   9
         Left            =   1620
         TabIndex        =   55
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   10
         Left            =   2220
         TabIndex        =   54
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   11
         Left            =   2820
         TabIndex        =   53
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   12
         Left            =   3420
         TabIndex        =   52
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   13
         Left            =   4020
         TabIndex        =   51
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   14
         Left            =   420
         TabIndex        =   50
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   15
         Left            =   1020
         TabIndex        =   49
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   16
         Left            =   1620
         TabIndex        =   48
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   17
         Left            =   2220
         TabIndex        =   47
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   18
         Left            =   2820
         TabIndex        =   46
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   19
         Left            =   3420
         TabIndex        =   45
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   20
         Left            =   4020
         TabIndex        =   44
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   21
         Left            =   420
         TabIndex        =   43
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   22
         Left            =   1020
         TabIndex        =   42
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   23
         Left            =   1620
         TabIndex        =   41
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   24
         Left            =   2220
         TabIndex        =   40
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   25
         Left            =   2820
         TabIndex        =   39
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   26
         Left            =   3420
         TabIndex        =   38
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   27
         Left            =   4020
         TabIndex        =   37
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   28
         Left            =   420
         TabIndex        =   36
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   29
         Left            =   1020
         TabIndex        =   35
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   30
         Left            =   1620
         TabIndex        =   34
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   31
         Left            =   2220
         TabIndex        =   33
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   32
         Left            =   2820
         TabIndex        =   32
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   33
         Left            =   3420
         TabIndex        =   31
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   34
         Left            =   4020
         TabIndex        =   30
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   35
         Left            =   420
         TabIndex        =   29
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   36
         Left            =   1020
         TabIndex        =   28
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   37
         Left            =   1620
         TabIndex        =   27
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   38
         Left            =   2220
         TabIndex        =   26
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   39
         Left            =   2820
         TabIndex        =   25
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   40
         Left            =   3420
         TabIndex        =   24
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   495
         Index           =   41
         Left            =   4020
         TabIndex        =   23
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblWeek 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "55"
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   390
         Width           =   255
      End
      Begin VB.Label lblWeek 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "55"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   870
         Width           =   255
      End
      Begin VB.Label lblWeek 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "55"
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   20
         Top             =   1335
         Width           =   255
      End
      Begin VB.Label lblWeek 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "55"
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   19
         Top             =   1815
         Width           =   255
      End
      Begin VB.Label lblWeek 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "55"
         Height          =   225
         Index           =   4
         Left            =   135
         TabIndex        =   18
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblWeek 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "55"
         Height          =   225
         Index           =   5
         Left            =   135
         TabIndex        =   17
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label lblWeek 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wk"
         Height          =   225
         Index           =   6
         Left            =   75
         TabIndex        =   16
         Top             =   90
         Width           =   255
      End
   End
   Begin VB.Frame fraTime 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   3795
      Width           =   2970
      Begin VB.TextBox txtHours 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   540
         TabIndex        =   13
         Text            =   "23"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdUpHrs 
         Caption         =   "+"
         Height          =   255
         Left            =   1020
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdDownHrs 
         Caption         =   "-"
         Height          =   255
         Left            =   1020
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtMin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1380
         TabIndex        =   10
         Text            =   "23"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdUpMin 
         Caption         =   "+"
         Height          =   255
         Left            =   1860
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdDownMin 
         Caption         =   "-"
         Height          =   255
         Left            =   1860
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtSec 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2220
         TabIndex        =   7
         Text            =   "23"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdUpSec 
         Caption         =   "+"
         Height          =   255
         Left            =   2700
         TabIndex        =   6
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdDownSec 
         Caption         =   "-"
         Height          =   255
         Left            =   2700
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Time: "
         Height          =   195
         Left            =   15
         TabIndex        =   14
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.TextBox txtCurDate 
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Top             =   4935
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   3030
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   1335
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      Left            =   510
      TabIndex        =   0
      Text            =   "cmbMonth"
      Top             =   90
      Width           =   1575
   End
   Begin AddressBook.chameleonButton cmdSelect 
      Height          =   450
      Left            =   4095
      TabIndex        =   76
      Top             =   4530
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   794
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
      MICON           =   "cfCalendar.frx":05F2
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
      Height          =   450
      Left            =   2940
      TabIndex        =   77
      Top             =   4530
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   794
      BTYPE           =   3
      TX              =   "Appointments"
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
      MICON           =   "cfCalendar.frx":060E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   330
      X2              =   4680
      Y1              =   4425
      Y2              =   4425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   330
      X2              =   4680
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   585
      Index           =   42
      Left            =   120
      TabIndex        =   3
      Top             =   4575
      Width           =   2805
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/* Copyright (c) 1998-2003
'/*************************************/
'Option Explicit

Public PassDate As Variant
Public cfWeekNumber As Integer
Public cfFirstday As Date
Public cfFirstWeekDay As Integer
Public ShowYear As Boolean
Public ShowTime As Boolean
Public ShowWeeks As Boolean

Dim cfCurDay As Date
Dim cfDayIndex As Integer
Dim cfWeekDay As Integer
Dim cfYearIndex As Integer
Dim cfPassedFirstDay As Boolean

Private Sub cmbMonth_Click()
 Dim tDate As Variant
    tDate = cmbMonth.ListIndex + 1 & Format(txtCurDate, "/dd/yyyy")
    If IsDate(tDate) Then
        If CDate(txtCurDate) <> CDate(tDate) Then
            txtCurDate = CDate(tDate)
        End If
    Else
        txtCurDate = CDate(cmbMonth.ListIndex + 1 & "/01/" & Format(txtCurDate, "yyyy"))
    End If
End Sub



Private Sub cmbMonth_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 37 Then
        cfDayIndex = cfDayIndex - 1
        If cfDayIndex < 0 Then cfDayIndex = 0
    ElseIf KeyCode = 39 Then
        cfDayIndex = cfDayIndex + 1
        If cfDayIndex > 41 Then cfDayIndex = 41
    End If
    txtCurDate = lblDate(cfDayIndex).Tag
End Sub


Private Sub cmbYear_Click()
  Dim tDate As Date
            
    If Not cfPassedFirstDay Then
        cfFirstday = CDate("1/1/" & cmbYear)
        'cfWeekDay = vbSunday
        Call cfFixFirstDay
        tDate = Format(txtCurDate, "mm/dd/") & cmbYear
        txtCurDate = tDate
    End If
    
    tDate = Format(txtCurDate, "mm/dd/") & cmbYear
    If IsDate(tDate) Then
        If cmbYear.ListIndex <> cfYearIndex Then
            cfFirstday = DateAdd("yyyy", cmbYear.ListIndex - cfYearIndex, cfFirstday)
             cfYearIndex = cmbYear.ListIndex
        End If
        txtCurDate = tDate
    End If
End Sub

Private Sub cmdDownHrs_Click()
    If IsNumeric(txtHours) Then
        txtHours = Format(txtHours - 1, "00")
        If txtHours < 0 Then
            txtHours = 23
            txtCurDate = CDate(txtCurDate) - 1
        End If
    End If
End Sub

Private Sub cmdDownMin_Click()
        If IsNumeric(txtMin) Then
        txtMin = Format(txtMin - 1, "00")
        If txtMin < 0 Then
            txtMin = 59
            cmdDownHrs_Click
        End If
    End If
End Sub

Private Sub cmdDownSec_Click()
    If IsNumeric(txtSec) Then
        txtSec = Format(txtSec - 1, "00")
        If txtSec < 0 Then
            txtSec = 59
            cmdDownMin_Click
        End If
    End If
End Sub

Private Sub cmdMonPrevNext_Click(Index As Integer)
  Dim i As Integer
  
    i = cmbMonth.ListIndex
    Select Case Index
    Case 0
        i = i - 1
        If i < 0 Then
            If ShowYear Then
                i = cmbYear.ListIndex
                i = i - 1
                If i < 0 Then i = 0
                cmbYear.ListIndex = i
                cmbMonth.ListIndex = cmbMonth.ListCount - 1
            Else
                i = 0
            End If
        Else
            cmbMonth.ListIndex = i
        End If
    Case Else
        i = i + 1
        If i > cmbMonth.ListCount - 1 Then
            If ShowYear Then
                i = cmbYear.ListIndex
                i = i + 1
                If i > cmbYear.ListCount - 1 Then i = cmbYear.ListCount - 1
                cmbYear.ListIndex = i
                cmbMonth.ListIndex = 0
            Else
                i = cmbMonth.ListCount - 1
            End If
        Else
            cmbMonth.ListIndex = i
        End If
    End Select
    
End Sub

Private Sub cmdUpHrs_Click()
    If IsNumeric(txtHours) Then
        txtHours = Format(txtHours + 1, "00")
        If txtHours > 23 Then
            txtHours = "00"
            txtCurDate = CDate(txtCurDate) + 1
        End If
    End If
End Sub

Private Sub cmdUpMin_Click()
    If IsNumeric(txtMin) Then
        txtMin = Format(txtMin + 1, "00")
        If txtMin > 59 Then
            txtMin = "00"
            cmdUpHrs_Click
        End If
    End If
End Sub

Private Sub cmdUpSec_Click()
    If IsNumeric(txtSec) Then
        txtSec = Format(txtSec + 1, "00")
        If txtSec > 59 Then
            txtSec = 0
            cmdUpMin_Click
        End If
    End If
End Sub

Private Sub cmdSelect_Click()
    If ShowTime Then
        PassDate = CDate(Format(txtCurDate, "mm/dd/yyyy") & " " & txtHours & ":" & txtMin & ":" & txtSec)
    Else
        PassDate = CDate(txtCurDate)
    End If
    cfWeekNumber = cfRetWeekNumber(PassDate)
    
    ShowYear = False
    ShowTime = False
    ShowWeeks = False
    cfFirstday = 0
    
    iLogOff = cLogOff
    LogOffTime = LogOffMinutes
    Me.Hide

End Sub

Private Sub cmdYRPrevNext_Click(Index As Integer)
  Dim i As Integer
  
    i = cmbYear.ListIndex
    Select Case Index
    Case 0
        i = i - 1
        If i < 0 Then i = 0
        cmbYear.ListIndex = i
    Case Else
        i = i + 1
        If i > cmbYear.ListCount - 1 Then i = cmbYear.ListCount - 1
        cmbYear.ListIndex = i
    End Select
    
End Sub

Private Sub Command1_Click()
 Unload Me
 frmAppointments.Show vbModal
End Sub

Private Sub Form_Activate()
  Dim i As Integer
    
    For i = 0 To 6
        lblWeek(i).Visible = ShowWeeks
    Next i

    cmbYear.Enabled = ShowYear
    cmdYRPrevNext(0).Visible = ShowYear
    cmdYRPrevNext(1).Visible = ShowYear
    
    fraTime.Visible = ShowTime
    
    If Not ShowTime Then
        cmdSelect.top = fraTime.top
        Command1.top = fraTime.top
        lblDate(42).top = fraTime.top
        Me.Height = 4800
    End If
    iLogOff = cLogOff
    LogOffTime = LogOffMinutes

End Sub

Private Function RetFirstDayOfYear(Optional tDate As Date = 0) As Variant
 'Dim tDate As Date
    
    If tDate = 0 Then
        tDate = CDate("1/1/" & Year(PassDate))
    Else
        tDate = CDate("1/1/" & Year(tDate))
    End If
    
    If cfFirstWeekDay = 0 Then cfFirstWeekDay = 1
    
    If Weekday(tDate) <> cfFirstWeekDay Then
        Do
            tDate = tDate - 1
            If Weekday(tDate) = cfFirstWeekDay Then Exit Do
        Loop
    End If
    '/* Calculate first day of new year */
    tDate = DateAdd("d", 1, tDate) + 6
    If Day(tDate) > 5 Then tDate = tDate - 7
    RetFirstDayOfYear = tDate

End Function

Private Sub Form_Deactivate()
    FormShowing = False
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PassDate = txtCurDate
        Me.Hide
        'Unload Me
    End If
    If KeyAscii = vbKeyEscape Then
        Me.Hide
        'Unload Me
    End If
    
End Sub

Private Sub Form_Load()
 Dim i As Long
    
'    'Debug Only
'    cfFirstDay = CDate("1/4/99")
'    cfWeekDay = WeekDay(cfFirstDay)
'    ShowWeeks = True
'    ShowYear = True
'    ShowTime = True
    
    cScreen.CenterForm Me
    
    'Me.Icon = MainForm.Icon
    'If cTile.MaxColors(Me) > 300 Then Call cTile.TileBackground(Me, frmArt!Image2, 0)
    
    If IsEmpty(PassDate) Then PassDate = Now
    If PassDate = 0 Or Not IsDate(PassDate) Then PassDate = Now
    
    If cfFirstday > 0 And IsDate(cfFirstday) Then
        cfWeekDay = Weekday(cfFirstday)
        cfPassedFirstDay = True
    Else
        cfFirstday = RetFirstDayOfYear
        cfWeekDay = Weekday(cfFirstday)
        'cfFirstday = CDate("1/1/" & Year(PassDate))
        'cfWeekDay = vbSunday
        'Call cfFixFirstDay
    End If
    
    cmbYear.Clear
    cmbMonth.Clear
    For i = 1 To 12
        cmbMonth.AddItem Format$(i & "/1/1998", "mmmm"), i - 1
    Next
    
    For i = 1900 To 2100
        cmbYear.AddItem i
        cmbYear.ItemData(cmbYear.NewIndex) = i
        If i = Year(PassDate) Then cfYearIndex = cmbYear.NewIndex
    Next i
    
    txtHours = (Mid$(Format(PassDate, "hh:mm:ss"), 1, 2))
    txtMin = (Mid$(Format(PassDate, "hh:mm:ss"), 4, 2))
    txtSec = (Mid$(Format(PassDate, "hh:mm:ss"), 7, 2))
    txtCurDate = Format(PassDate, "mm/dd/yyyy")

End Sub
Private Sub cfFillDates(ByVal cfCurDate As Date)
  Dim cfSDate As Date, i As Integer, n As Integer
    
    Call cfFindYear(Year(cfCurDate))
'    cmbMonth.ListIndex = Month(cfCurDate) - 1
    
    cfSDate = CDate(Format$(cfCurDate, "mm/1/yyyy"))
    cfSDate = cfSDate - Weekday(cfSDate) + 1
    
    n = cfWeekDay
    For i = 0 To 6
        lblDays(i) = cfDayString(n)
        n = n + 1
        If n > 7 Then n = 1
    Next i
    
    cfSDate = CDate(Format$(cfCurDate, "mm/1/yyyy"))
    cfSDate = cfSDate - Weekday(cfSDate, cfWeekDay) + 1
    
    For i = 0 To 41
        lblDate(i).Tag = cfSDate + i
        lblDate(i).Caption = Format$(cfSDate + i, "d")
        If CDate(lblDate(i).Tag) = cfCurDate Then
            lblDate(i).BackColor = &H80000005
            cfDayIndex = i
        Else
            lblDate(i).BackColor = &H8000000F
        End If
        If Format(cfCurDate, "mm") <> Format(cfSDate + i, "mm") Then
            lblDate(i).ForeColor = &HFF&
        Else
            lblDate(i).ForeColor = &H0&
        End If
        
        cfFirstday = RetFirstDayOfYear(CDate(txtCurDate))
        
        '/* Display Week numbers */
        If i / 7 = Int(i / 7) Then
            lblWeek(i / 7).Caption = cfRetWeekNumber(cfSDate + i)
        End If
    Next
    cmbMonth.ListIndex = Month(cfCurDate) - 1
'    Call cfFindYear(Year(cfCurDate))
    
    LogOffTime = LogOffMinutes
    
End Sub
Private Function cfRetWeekNumber(ByVal InDate As Date)
  Dim cftFirstDayOfYear As Variant
  Dim a As Integer, B As Integer
    
    cftFirstDayOfYear = cfFirstday - Weekday(cfFirstday, cfWeekDay) + 1
    a = DateDiff("ww", cftFirstDayOfYear, InDate - Weekday(InDate, cfWeekDay) + 1) + 1
    'b = WeeksInAYear(cftFirstDayOfYear)
    If a = 0 Then a = 52
    cfRetWeekNumber = a

End Function

Private Function WeeksInAYear(Optional ByVal FirstDayOfYear As Variant, _
                            Optional ByVal CurrentYear As Variant, _
                            Optional WeekStartsOn As VbDayOfWeek = vbUseSystemDayOfWeek) As Integer
  
  Dim tDate As Date
    
    If IsMissing(FirstDayOfYear) Then
        If IsMissing(CurrentYear) Then CurrentYear = Year(Date)
        FirstDayOfYear = CDate("1/1/" & CurrentYear)
    Else
        If WeekStartsOn = vbUseSystemDayOfWeek Then
            WeekStartsOn = Weekday(FirstDayOfYear)
        End If
    End If
    
    tDate = DateAdd("yyyy", 1, FirstDayOfYear) + 6
    If Day(tDate) > 6 Then tDate = tDate - 7
    WeeksInAYear = DateDiff("w", FirstDayOfYear, tDate, WeekStartsOn)
    
End Function

Private Sub cfFindYear(InYear As Integer)
 Dim i As Integer
    For i = 0 To cmbYear.ListCount - 1
        If cmbYear.ItemData(i) = InYear Then
            cmbYear.ListIndex = i
            Exit For
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCalendar = Nothing
End Sub

Private Sub lblDate_Click(Index As Integer)
    If Index < 42 Then
        cfDayIndex = Index
        txtCurDate = lblDate(Index).Tag
    Else
        If ShowYear Or Year(txtCurDate) = Year(Date) Then
            txtCurDate = Date
            Call cfFillDates(txtCurDate)
        End If
    End If
End Sub

Private Sub lblDate_DblClick(Index As Integer)
    Call cmdSelect_Click
End Sub

Private Sub txtCurDate_Change()
    If IsDate(txtCurDate) Then
        If ShowYear Then
            Call cfFillDates(txtCurDate)
        Else
            txtCurDate = Format(txtCurDate, "mm/dd/" & Year(PassDate))
            Call cfFillDates(txtCurDate)
        End If
        lblDate(42) = "[Today]   " & Format(CDate(txtCurDate), "dddd - mmmm d, yyyy")
    End If
End Sub

Private Sub txtCurDate_KeyPress(KeyAscii As Integer)
    KeyAscii = False
End Sub



Private Function cfDayString(tDate As Variant) As String
    Select Case Weekday(tDate)
    Case 2
        cfDayString = "Mon"
    Case 3
        cfDayString = "Tue"
    Case 4
        cfDayString = "Wed"
    Case 5
        cfDayString = "Thu"
    Case 6
        cfDayString = "Fri"
    Case 7
        cfDayString = "Sat"
    Case 1
        cfDayString = "Sun"
    End Select
End Function

Private Sub cfFixFirstDay()
    If Weekday(cfFirstday) <> cfWeekDay Then
        Do
            cfFirstday = cfFirstday - 1
            If Weekday(cfFirstday) = cfWeekDay Then Exit Do
        Loop
    End If
End Sub

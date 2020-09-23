VERSION 5.00
Begin VB.Form MainMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Book"
   ClientHeight    =   6270
   ClientLeft      =   1245
   ClientTop       =   1755
   ClientWidth     =   8865
   ForeColor       =   &H80000008&
   Icon            =   "MAINMENI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   8865
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   870
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   8805
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   0
      Width           =   8865
      Begin VB.PictureBox picIndexButtons 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   720
         ScaleHeight     =   315
         ScaleWidth      =   6720
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   480
         Width           =   6720
         Begin VB.CommandButton cmd_letter 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   30
            TabIndex        =   77
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   285
            TabIndex        =   76
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   540
            TabIndex        =   75
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   795
            TabIndex        =   74
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "E"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1050
            TabIndex        =   73
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "F"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1305
            TabIndex        =   72
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "G"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1560
            TabIndex        =   71
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "H"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1815
            TabIndex        =   70
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   2070
            TabIndex        =   69
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "J"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   2325
            TabIndex        =   68
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "K"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   2580
            TabIndex        =   67
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "L"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   2835
            TabIndex        =   66
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "M"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   3090
            TabIndex        =   65
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   3345
            TabIndex        =   64
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   3600
            TabIndex        =   63
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   3855
            TabIndex        =   62
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "Q"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   4110
            TabIndex        =   61
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   4365
            TabIndex        =   60
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   4620
            TabIndex        =   59
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "T"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   4875
            TabIndex        =   58
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "U"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   5130
            TabIndex        =   57
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "V"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   5385
            TabIndex        =   56
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "W"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   5640
            TabIndex        =   55
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   5895
            TabIndex        =   54
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   6150
            TabIndex        =   53
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmd_letter 
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   6405
            TabIndex        =   52
            Top             =   30
            Width           =   255
         End
      End
      Begin AddressBook.chameleonButton cmd_Exit 
         Height          =   375
         Left            =   45
         TabIndex        =   96
         ToolTipText     =   "End"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "MAINMENI.frx":08CA
         PICN            =   "MAINMENI.frx":08E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_New 
         Height          =   375
         Left            =   1005
         TabIndex        =   98
         ToolTipText     =   "New"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":0DAC
         PICN            =   "MAINMENI.frx":0DC8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmdSave 
         Height          =   375
         Left            =   1380
         TabIndex        =   99
         ToolTipText     =   "Save"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":128E
         PICN            =   "MAINMENI.frx":12AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_Cut 
         Height          =   375
         Left            =   2190
         TabIndex        =   101
         ToolTipText     =   "Delete"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":1724
         PICN            =   "MAINMENI.frx":1740
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_Print 
         Height          =   375
         Left            =   2565
         TabIndex        =   102
         ToolTipText     =   "Print"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":1C06
         PICN            =   "MAINMENI.frx":1C22
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmdDates 
         Height          =   375
         Left            =   2940
         TabIndex        =   103
         ToolTipText     =   "Dates"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":20E8
         PICN            =   "MAINMENI.frx":2104
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_find 
         Height          =   375
         Left            =   3315
         TabIndex        =   104
         ToolTipText     =   "Find"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":25CA
         PICN            =   "MAINMENI.frx":25E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_Dial 
         Height          =   375
         Left            =   3690
         TabIndex        =   105
         ToolTipText     =   "Dial"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":2AAC
         PICN            =   "MAINMENI.frx":2AC8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmdSortBy 
         Height          =   375
         Left            =   4065
         TabIndex        =   106
         ToolTipText     =   "Change Filter"
         Top             =   30
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "Sort By Company"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":2F8E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_calander 
         Height          =   375
         Left            =   6285
         TabIndex        =   107
         ToolTipText     =   "Calendar"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":2FAA
         PICN            =   "MAINMENI.frx":2FC6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_app 
         Height          =   375
         Left            =   6660
         TabIndex        =   108
         ToolTipText     =   "Appoitments"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MAINMENI.frx":348C
         PICN            =   "MAINMENI.frx":34A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_About 
         Height          =   375
         Left            =   7035
         TabIndex        =   109
         ToolTipText     =   "About"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":396E
         PICN            =   "MAINMENI.frx":398A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmdImportExport 
         Height          =   375
         Left            =   525
         TabIndex        =   97
         ToolTipText     =   "Import/Export"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":3E50
         PICN            =   "MAINMENI.frx":3E6C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmdCancelNew 
         Height          =   375
         Left            =   1755
         TabIndex        =   100
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "MAINMENI.frx":4426
         PICN            =   "MAINMENI.frx":4442
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
         X1              =   0
         X2              =   8805
         Y1              =   435
         Y2              =   435
      End
   End
   Begin VB.OptionButton optAddress 
      Caption         =   "Home"
      Height          =   525
      Index           =   0
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4005
      Value           =   -1  'True
      Width           =   780
   End
   Begin VB.OptionButton optAddress 
      Caption         =   "Business"
      Height          =   525
      Index           =   1
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4530
      Width           =   780
   End
   Begin VB.CheckBox chk_autocap 
      Caption         =   "Auto Captilize"
      Height          =   480
      Left            =   7755
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5280
      Value           =   1  'Checked
      Width           =   915
   End
   Begin VB.PictureBox picNamePanel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2970
      Left            =   30
      ScaleHeight     =   2970
      ScaleWidth      =   8775
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   930
      Width           =   8775
      Begin VB.Frame Frame1 
         Caption         =   "Company"
         Height          =   570
         Index           =   0
         Left            =   90
         TabIndex        =   49
         Top             =   60
         Width           =   7050
         Begin VB.TextBox txt_Company 
            BackColor       =   &H0000FFFF&
            DataField       =   "Company"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   0
            Top             =   210
            Width           =   6810
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Title"
         Height          =   570
         Index           =   1
         Left            =   75
         TabIndex        =   48
         Top             =   630
         Width           =   4170
         Begin VB.TextBox txt_title 
            BackColor       =   &H0000FFFF&
            DataField       =   "Title"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   1
            Top             =   210
            Width           =   3930
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Birthday"
         Height          =   570
         Index           =   2
         Left            =   4245
         TabIndex        =   47
         Top             =   630
         Width           =   1290
         Begin VB.TextBox txt_bday 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            CausesValidation=   0   'False
            DataField       =   "Birthdate"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "00/00/00"
            ToolTipText     =   "DoubleClick for Calendar"
            Top             =   210
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Anniversary"
         Height          =   570
         Index           =   3
         Left            =   5535
         TabIndex        =   46
         Top             =   630
         Width           =   1290
         Begin VB.TextBox txt_anv 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            CausesValidation=   0   'False
            DataField       =   "Anniversary"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "00/00/00"
            ToolTipText     =   "DoubleClick for Calendar"
            Top             =   210
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Prefix"
         Height          =   570
         Index           =   4
         Left            =   75
         TabIndex        =   45
         Top             =   1200
         Width           =   1275
         Begin VB.TextBox txt_prefix 
            BackColor       =   &H0000FFFF&
            DataField       =   "Prefix"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            MaxLength       =   5
            TabIndex        =   4
            Top             =   210
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "First Name"
         Height          =   570
         Index           =   5
         Left            =   1350
         TabIndex        =   44
         Top             =   1200
         Width           =   2385
         Begin VB.TextBox txt_first 
            BackColor       =   &H0000FFFF&
            DataField       =   "FirstName"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   135
            MaxLength       =   25
            TabIndex        =   5
            Top             =   210
            Width           =   2010
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "MI"
         Height          =   570
         Index           =   6
         Left            =   3735
         TabIndex        =   43
         Top             =   1200
         Width           =   705
         Begin VB.TextBox txt_mi 
            BackColor       =   &H0000FFFF&
            DataField       =   "MiddleName"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   210
            MaxLength       =   20
            TabIndex        =   6
            Top             =   210
            Width           =   270
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Last Name"
         Height          =   570
         Index           =   7
         Left            =   4440
         TabIndex        =   42
         Top             =   1200
         Width           =   2385
         Begin VB.TextBox txt_last 
            BackColor       =   &H0000FFFF&
            DataField       =   "LastName"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   105
            MaxLength       =   25
            TabIndex        =   7
            Top             =   210
            Width           =   2130
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Business Phone"
         Height          =   570
         Index           =   8
         Left            =   75
         TabIndex        =   41
         Top             =   1770
         Width           =   1650
         Begin VB.TextBox txt_bphone 
            BackColor       =   &H0000FFFF&
            DataField       =   "WorkPhone"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   135
            MaxLength       =   30
            TabIndex        =   8
            Top             =   210
            Width           =   1350
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mobal Number"
         Height          =   570
         Index           =   9
         Left            =   1770
         TabIndex        =   40
         Top             =   1770
         Width           =   1650
         Begin VB.TextBox txt_mphone 
            BackColor       =   &H0000FFFF&
            DataField       =   "MobalPhone"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   165
            MaxLength       =   30
            TabIndex        =   9
            Top             =   210
            Width           =   1340
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fax Number"
         Height          =   570
         Index           =   10
         Left            =   3465
         TabIndex        =   39
         Top             =   1770
         Width           =   1650
         Begin VB.TextBox txt_fphone 
            BackColor       =   &H0000FFFF&
            DataField       =   "FaxNumber"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   165
            MaxLength       =   30
            TabIndex        =   10
            Top             =   210
            Width           =   1340
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Home Phone"
         Height          =   570
         Index           =   11
         Left            =   5160
         TabIndex        =   38
         Top             =   1770
         Width           =   1650
         Begin VB.TextBox txt_hphone 
            BackColor       =   &H0000FFFF&
            DataField       =   "HomePhone"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   180
            MaxLength       =   30
            TabIndex        =   11
            Top             =   210
            Width           =   1340
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Type"
         Height          =   1695
         Left            =   7245
         TabIndex        =   33
         Top             =   630
         Width           =   1395
         Begin VB.CheckBox chk_bus 
            Caption         =   "Business"
            DataField       =   "Business"
            DataSource      =   "Data1"
            Height          =   255
            Left            =   75
            TabIndex        =   37
            Top             =   240
            Width           =   960
         End
         Begin VB.CheckBox chk_rel 
            Caption         =   "Relative"
            DataField       =   "Relative"
            DataSource      =   "Data1"
            Height          =   255
            Left            =   75
            TabIndex        =   36
            Top             =   485
            Width           =   960
         End
         Begin VB.CheckBox chk_fri 
            Caption         =   "Friend"
            DataField       =   "Friend"
            DataSource      =   "Data1"
            Height          =   255
            Left            =   75
            TabIndex        =   35
            Top             =   730
            Width           =   960
         End
         Begin VB.ComboBox cboOther 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   75
            Sorted          =   -1  'True
            TabIndex        =   34
            Top             =   1230
            Width           =   1185
         End
         Begin VB.Label Label6 
            Caption         =   "Other"
            Height          =   165
            Left            =   345
            TabIndex        =   95
            Top             =   1005
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "E-Mail"
         Height          =   570
         Index           =   12
         Left            =   75
         TabIndex        =   32
         Top             =   2340
         Width           =   3345
         Begin VB.TextBox txt_EMail 
            BackColor       =   &H0000FFFF&
            DataField       =   "EMail"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   195
            Width           =   3150
         End
      End
      Begin AddressBook.chameleonButton cmdEmail 
         Height          =   420
         Left            =   3450
         TabIndex        =   110
         ToolTipText     =   "e-mail"
         Top             =   2460
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
         BTYPE           =   14
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":4889
         PICN            =   "MAINMENI.frx":48A5
         PICH            =   "MAINMENI.frx":4DA7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_next 
         Height          =   405
         Left            =   7740
         TabIndex        =   111
         ToolTipText     =   "Next"
         Top             =   2460
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   714
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MAINMENI.frx":52A9
         PICN            =   "MAINMENI.frx":52C5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_prev 
         Height          =   405
         Left            =   7020
         TabIndex        =   112
         ToolTipText     =   "Previous"
         Top             =   2460
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   714
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MAINMENI.frx":53F3
         PICN            =   "MAINMENI.frx":540F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmdGrid 
         Height          =   405
         Left            =   6255
         TabIndex        =   113
         ToolTipText     =   "Show Grid"
         Top             =   2460
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Grid"
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MAINMENI.frx":553D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AddressBook.chameleonButton cmd_phone 
         Height          =   405
         Left            =   4890
         TabIndex        =   114
         ToolTipText     =   "More Numbers"
         Top             =   2460
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "User Defined"
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "MAINMENI.frx":5559
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
   Begin VB.PictureBox picAddressPanel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1890
      Left            =   30
      ScaleHeight     =   1890
      ScaleWidth      =   7590
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4020
      Width           =   7590
      Begin VB.Frame Frame1 
         Caption         =   "Memo"
         Height          =   570
         Index           =   17
         Left            =   60
         TabIndex        =   26
         Top             =   1215
         Width           =   7050
         Begin VB.TextBox txt_memo 
            BackColor       =   &H0000FFFF&
            DataField       =   "Note"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   105
            TabIndex        =   21
            Top             =   210
            Width           =   6780
         End
      End
      Begin AddressBook.chameleonButton cmd_Expand 
         Height          =   420
         Left            =   7125
         TabIndex        =   115
         ToolTipText     =   "expand memo"
         Top             =   1335
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
         BTYPE           =   14
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
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "MAINMENI.frx":5575
         PICN            =   "MAINMENI.frx":5591
         PICH            =   "MAINMENI.frx":5A57
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame1 
         Caption         =   "Zip Code"
         Height          =   570
         Index           =   16
         Left            =   4830
         TabIndex        =   27
         Top             =   645
         Width           =   1830
         Begin VB.TextBox txt_bzip 
            BackColor       =   &H0000FFFF&
            DataField       =   "WorkZipCode"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            MaxLength       =   20
            TabIndex        =   20
            Top             =   210
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.TextBox txt_zip 
            BackColor       =   &H0000FFFF&
            DataField       =   "HomeZipCode"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            MaxLength       =   20
            TabIndex        =   16
            Top             =   210
            Width           =   1635
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "State"
         Height          =   570
         Index           =   15
         Left            =   4140
         TabIndex        =   28
         Top             =   645
         Width           =   690
         Begin VB.TextBox txt_bstate 
            BackColor       =   &H0000FFFF&
            DataField       =   "WorkState"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            MaxLength       =   2
            TabIndex        =   19
            Top             =   210
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txt_state 
            BackColor       =   &H0000FFFF&
            DataField       =   "HomeState"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            MaxLength       =   2
            TabIndex        =   15
            Top             =   210
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "City"
         Height          =   570
         Index           =   14
         Left            =   60
         TabIndex        =   29
         Top             =   645
         Width           =   4080
         Begin VB.TextBox txt_bcity 
            BackColor       =   &H0000FFFF&
            DataField       =   "WorkCity"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   105
            MaxLength       =   50
            TabIndex        =   18
            Top             =   210
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox txt_city 
            BackColor       =   &H0000FFFF&
            DataField       =   "HomeCity"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            MaxLength       =   50
            TabIndex        =   14
            Top             =   210
            Width           =   3870
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Address"
         Height          =   570
         Index           =   13
         Left            =   60
         TabIndex        =   30
         Top             =   75
         Width           =   7050
         Begin VB.TextBox txt_baddress 
            BackColor       =   &H0000FFFF&
            DataField       =   "WorkAddress"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   105
            MaxLength       =   80
            TabIndex        =   17
            Top             =   210
            Visible         =   0   'False
            Width           =   6795
         End
         Begin VB.TextBox txt_address 
            BackColor       =   &H0000FFFF&
            DataField       =   "HomeAddress"
            DataSource      =   "Data1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   105
            MaxLength       =   80
            TabIndex        =   13
            Top             =   210
            Width           =   6795
         End
      End
   End
   Begin VB.PictureBox pnl_phone 
      Height          =   3195
      Left            =   30
      ScaleHeight     =   3135
      ScaleWidth      =   8745
      TabIndex        =   78
      Top             =   2730
      Visible         =   0   'False
      Width           =   8805
      Begin AddressBook.chameleonButton cmdCloseUser 
         Height          =   465
         Left            =   6750
         TabIndex        =   116
         Top             =   2355
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   820
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
         MICON           =   "MAINMENI.frx":5F1D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txt_u3text 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "User3Text"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4110
         TabIndex        =   89
         Text            =   "International"
         Top             =   1260
         Width           =   1515
      End
      Begin VB.TextBox txt_u2text 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "User2Text"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4110
         TabIndex        =   88
         Text            =   "User2"
         Top             =   870
         Width           =   1515
      End
      Begin VB.TextBox txt_u3phone 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "User3Phone"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5685
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   1260
         Width           =   2500
      End
      Begin VB.TextBox txt_u1text 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "User1Text"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4095
         TabIndex        =   86
         Text            =   "User1"
         Top             =   480
         Width           =   1515
      End
      Begin VB.TextBox txt_u2phone 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "User2Phone"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5670
         TabIndex        =   85
         Text            =   "Text1"
         Top             =   870
         Width           =   2500
      End
      Begin VB.TextBox txt_u1phone 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "User1Phone"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5655
         TabIndex        =   84
         Text            =   "txt_u1phone"
         Top             =   480
         Width           =   2500
      End
      Begin VB.TextBox txt_hphone2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "HomePhone"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   1650
         Width           =   2500
      End
      Begin VB.TextBox txt_fphone2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "FaxNumber"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   1260
         Width           =   2500
      End
      Begin VB.TextBox txt_mphone2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "MobalPhone"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   870
         Width           =   2500
      End
      Begin VB.TextBox txt_bphone2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "WorkPhone"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   80
         Text            =   "Text1"
         Top             =   480
         Width           =   2500
      End
      Begin VB.TextBox txt_EMail2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         DataField       =   "EMail"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4770
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   1665
         Width           =   3405
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " E-Mail: "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4050
         TabIndex        =   94
         Top             =   1710
         Width           =   720
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Home:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   570
         TabIndex        =   93
         Top             =   1710
         Width           =   555
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   750
         TabIndex        =   92
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   525
         TabIndex        =   91
         Top             =   930
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Business:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   90
         Top             =   555
         Width           =   825
      End
   End
   Begin VB.Menu mnuMenu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu mnuAll 
         Caption         =   "Order By Company"
      End
      Begin VB.Menu mnuLast 
         Caption         =   "Order by Last Name"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBusiness 
         Caption         =   "Filter by Business"
      End
      Begin VB.Menu mnuFriend 
         Caption         =   "Filter by Friend"
      End
      Begin VB.Menu mnuRelative 
         Caption         =   "Filter by Relative"
      End
      Begin VB.Menu mnuOther 
         Caption         =   "Filter by Other"
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditOther 
         Caption         =   "Edit/Add 'Other' names"
      End
   End
   Begin VB.Menu mnuMenu2 
      Caption         =   "menu2"
      Visible         =   0   'False
      Begin VB.Menu mnuScreen 
         Caption         =   "Print Screen"
      End
      Begin VB.Menu mnuCurrent 
         Caption         =   "Print Current"
      End
      Begin VB.Menu mnuSortedBy 
         Caption         =   "Print Sorted By"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "mnuFile"
      Visible         =   0   'False
      Begin VB.Menu mnuExport 
         Caption         =   "Export to a VCard file"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import a VCard file"
      End
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private m_AddingNew As Boolean
Private m_BookMark  As Variant

Private m_DB  As ADODB.Connection
Public m_Set As ADODB.Recordset


Private Sub chk_bus_Click()
    If chk_bus.Value Then
       optAddress(1).Value = True
    Else
       optAddress(0).Value = True
    End If
End Sub




Private Sub cmd_about_Click()
    frmAbout.Show , Me ' vbModal
End Sub

Private Sub cmd_app_Click()
    frmAppointments.Show vbModal, Me
    Unload frmAppointments
End Sub

Private Sub cmd_calander_Click()

    With frmCalendar
        .ShowYear = True
        .ShowTime = False
        .ShowWeeks = False
        .Show vbModal, Me
    End With

End Sub

Private Sub cmd_cut_Click()

    On Error GoTo Err_Proc

    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Warning") = vbYes Then
        m_Set.Delete
        m_Set.MoveFirst
        Call UpdateData
    End If
    Call cmd_next_click

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "cmd_cut_Click"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub cmd_dial_Click()
    PhoneNumber = vbNullString
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_expand_Click()
    frmMemo!Text1.Text = txt_memo.Text
    frmMemo.Show vbModal, Me
    txt_memo.Text = frmMemo!Text1.Text
    Unload frmMemo
End Sub

Private Sub cmd_find_Click()
    
    On Local Error Resume Next

'    If m_AddingNew Then
'       m_AddingNew = False
'       cmdSave.Enabled = False
'    End If
    
    frmSearchFor.Show vbModal, Me
    Unload frmSearchFor
    
End Sub

Private Sub cmd_letter_Click(Index As Integer)
  
  Dim IDn        As Long
  Dim tString    As String
  Dim MyCriteria As String
  
    On Local Error Resume Next
    
    'If m_AddingNew Then
    '   m_AddingNew = False
    '   cmdSave.Enabled = False
    '   m_BookMark = False
    'End If
    
    IDn = m_Set!ID
    tString = Chr$(Index + 65)
    MyCriteria = "[" & SortField & "] LIKE '" & tString & "%'"
    
    If Not ADOFindFirst(m_Set, MyCriteria) Then
        Beep
        MyCriteria = "[ID] = " & CStr(IDn)
        ADOFindFirst m_Set, MyCriteria
        Call UpdateData
    End If
    Call UpdateData
    
    txt_Company.SetFocus
    
End Sub

Private Sub cmd_new_Click()
    On Local Error Resume Next
    m_BookMark = m_Set!ID
    m_Set.AddNew
    Call ShowButtons(False)
    
    txt_Company.Text = vbNullString
    txt_title.Text = vbNullString
    txt_bday.Text = vbNullString
    txt_anv.Text = vbNullString
    txt_prefix.Text = vbNullString
    txt_first.Text = vbNullString
    txt_mi.Text = vbNullString
    txt_last.Text = vbNullString
    txt_bphone.Text = vbNullString
    txt_mphone.Text = vbNullString
    txt_fphone.Text = vbNullString
    txt_hphone.Text = vbNullString
    txt_EMail.Text = vbNullString
    txt_baddress.Text = vbNullString
    txt_bcity.Text = vbNullString
    txt_bstate.Text = vbNullString
    txt_bzip.Text = vbNullString
    txt_memo.Text = vbNullString
    txt_bphone2.Text = vbNullString
    txt_u1text.Text = "User1"
    txt_u1phone.Text = vbNullString
    txt_mphone2.Text = vbNullString
    txt_u2text.Text = "User2"
    txt_u2phone.Text = vbNullString
    txt_fphone2.Text = vbNullString
    txt_u3text.Text = "User3"
    txt_u3phone.Text = vbNullString
    txt_hphone2.Text = vbNullString
    txt_EMail2.Text = vbNullString
    txt_address.Text = vbNullString
    txt_city.Text = vbNullString
    txt_state.Text = vbNullString
    txt_zip.Text = vbNullString
    chk_bus.Value = vbUnchecked
    chk_rel.Value = vbUnchecked
    chk_fri.Value = vbUnchecked
    cboOther.Text = vbNullString
    
    txt_Company.SetFocus
    
End Sub

Private Sub cmd_next_click()
    On Local Error Resume Next
    Call ShowButtons(True)
    
    'If Not (m_Set.EOF And m_Set.BOF) Then Call SaveData
    
    m_Set.MoveNext
    If m_Set.EOF Then
        m_Set.MoveFirst
        Beep
    End If
    Call UpdateData
    
End Sub

Private Sub cmd_phone_Click()
    PnlPhoneFront = True
    pnl_phone.Visible = True
    pnl_phone.ZOrder
End Sub

Private Sub cmd_prev_Click()
    On Local Error Resume Next
    Call ShowButtons(True)
    
    'If Not (m_Set.EOF And m_Set.BOF) Then Call SaveData
    
    m_Set.MovePrevious
    If m_Set.BOF Then
        m_Set.MoveLast
        Beep
    End If
    Call UpdateData
    
End Sub

Private Sub cmd_print_Click()
    PopupMenu mnuMenu2, , cmd_print.left, cmd_print.top + cmd_print.Height
End Sub

Private Sub cmdCancelNew_Click()

    On Error GoTo Err_Proc

    m_Set.CancelUpdate
    cmd_next_click
    cmd_prev_Click
    If m_Set.RecordCount = 0 Then
        cmd_new_Click
    Else
        Call UpdateData
    End If

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "cmdCancelNew_Click"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub cmdEmail_Click()
  Const SW_SHOW = 3
    ShellExecute Me.hWnd, "open", "mailto:" & txt_EMail, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub cmdGrid_Click()
    frmListing.Caption = "Listing - " & cmdSortBy.Caption
    frmListing.Show vbModal, Me
    Call UpdateData
End Sub

Private Sub cmdImportExport_Click()
  Dim xPos As Long, yPos As Long
  
    xPos = cmdImportExport.left
    yPos = cmdImportExport.top + cmdImportExport.Height
    PopupMenu mnuFile, , xPos, yPos

End Sub

Private Sub cmdSave_Click()
  
  Dim uID As Long
    
    On Error GoTo Err_Proc
    
    Call SaveData
    uID = m_Set!ID
    ADOFindFirst m_Set, "[ID]=" & CStr(uID)
    'm_AddingNew = False
    cmdSave.Enabled = False
    Call UpdateData
    Call ShowButtons(True)

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "cmdSave_Click"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub cmdSortBy_Click()
    
    PopupMenu mnuMenu1, , cmdSortBy.left, cmdSortBy.top + cmdSortBy.Height
    
End Sub

Private Sub cmdDates_Click()
    frmDates.Show vbModal, Me
    Unload frmDates
End Sub

Private Sub cmdCloseUser_Click()
    PnlPhoneFront = False
    pnl_phone.Visible = False
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       SendKeys "{TAB}"
       KeyAscii = False
    End If

End Sub

Private Sub Form_Load()

    On Error GoTo Err_Proc

  Dim X As Integer, FirstTime As Boolean
  Dim c3d As cls3Dcontrol
  
    cScreen.CenterForm Me
    
    Set c3d = New cls3Dcontrol
    picNamePanel.BackColor = &H8000000F
    picAddressPanel.BackColor = &H8000000F
    Show
    c3d.ThreeDControl Me, picToolBar
    c3d.ThreeDControl picToolBar, picIndexButtons
    c3d.ThreeDControl Me, picNamePanel
    c3d.ThreeDControl Me, picAddressPanel
    Set c3d = Nothing
    
    Call OpenDB(m_DB)
    Call OpenRS(m_Set, "SELECT * FROM Other ORDER BY [Other]", m_DB)
    If Not (m_Set.EOF And m_Set.BOF) Then
        Do
           cboOther.AddItem m_Set!Other
           m_Set.MoveNext
        Loop Until m_Set.EOF
    End If
    cboOther.AddItem vbNullString
    m_Set.Close
    
    '/* Check to see if there has been any data saved
    Call OpenRS(m_Set, "SELECT AddBook.* FROM AddBook", m_DB)
    If m_Set.EOF And m_Set.BOF Then FirstTime = True
    m_Set.Close
    
    Set m_Set = Nothing
    
    Me.Caption = "Address Book v" & App.Major & "." & Format(App.Revision, "00")
    
    Call SortData
    
    Call ShowButtons(True)
    
    Unload frmAbout
    
    If FirstTime Then Call cmd_new_Click
  

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "Form_Load"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If UnloadMode = vbFormControlMenu Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    m_Set.Close
    m_DB.Close

    Set cValidate = Nothing
    Set cScreen = Nothing
    Call EndProgram

End Sub



Private Sub optAddress_Click(Index As Integer)
    If Index Then
        Frame1(13).Caption = "Business Address"
        Frame1(14).Caption = "Business City"
        txt_baddress.Visible = True
        txt_address.Visible = False
        txt_bcity.Visible = True
        txt_city.Visible = False
        txt_bstate.Visible = True
        txt_state.Visible = False
        txt_bzip.Visible = True
        txt_zip.Visible = False
    Else
        Frame1(13).Caption = "Address"
        Frame1(14).Caption = "City"
        txt_baddress.Visible = False
        txt_address.Visible = True
        txt_bcity.Visible = False
        txt_city.Visible = True
        txt_bstate.Visible = False
        txt_state.Visible = True
        txt_bzip.Visible = False
        txt_zip.Visible = True
    End If
    
    On Error Resume Next
    txt_Company.SetFocus

End Sub

Private Sub txt_address_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_address
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)
    
End Sub

Private Sub txt_anv_DblClick()

    On Error GoTo Err_Proc

    With frmCalendar
        If IsDate(txt_anv) Then .PassDate = CDate(txt_anv)
        .Command1.Visible = False
        .ShowYear = True
        .ShowTime = False
        .ShowWeeks = False
        .Show vbModal, Me
        txt_anv = Format(.PassDate, "mm/dd/yyyy")
    End With
    Unload frmCalendar


Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "txt_anv_DblClick"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub txt_anv_KeyUp(KeyCode As Integer, Shift As Integer)
    
    cValidate.MaskedText txt_anv, DateProfile
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_anv_LostFocus()
    txt_anv = Format(txt_anv, "mm/dd/yyyy")
End Sub


Private Sub txt_baddress_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_baddress
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)
    
End Sub

Private Sub txt_bcity_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_bcity
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)
    
End Sub

Private Sub txt_bday_DblClick()

    On Error GoTo Err_Proc
    
    With frmCalendar
        If IsDate(txt_bday) Then .PassDate = CDate(txt_bday)
        .Command1.Visible = False
        .ShowYear = True
        .ShowTime = False
        .ShowWeeks = False
        .Show vbModal, Me
        txt_bday = Format(.PassDate, "mm/dd/yyyy")
    End With
    Unload frmCalendar
    

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "txt_bday_DblClick"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub txt_bday_KeyUp(KeyCode As Integer, Shift As Integer)
    
    cValidate.MaskedText txt_bday, DateProfile
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_bday_LostFocus()
    txt_bday = Format(txt_bday, "mm/dd/yyyy")
End Sub


Private Sub txt_bphone_DblClick()
    PhoneNumber = txt_bphone
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub txt_bphone_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode > 13 Then cValidate.MaskedText txt_bphone, PhoneProfile
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)
    
End Sub


Private Sub txt_bphone2_DblClick()
    PhoneNumber = txt_bphone2
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub txt_bphone2_KeyPress(KeyAscii As Integer)
    Call cValidate.MaskedText(txt_bphone2, PhoneProfile)
End Sub

Private Sub txt_bphone2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_bstate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_bstate_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)
    
End Sub


Private Sub txt_bzip_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Call cValidate.MaskedText(txt_bzip, ZipProfile)
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)
    
End Sub


Private Sub txt_city_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_city
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)
    
End Sub

Private Sub txt_Company_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_Company
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)
    
End Sub


Private Sub txt_EMail_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_EMail2_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_first_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_first
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_fphone_DblClick()
    PhoneNumber = txt_fphone
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub txt_fphone_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode > 36 Then cValidate.MaskedText txt_fphone, PhoneProfile
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_fphone2_DblClick()
    PhoneNumber = txt_fphone2
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub txt_fphone2_KeyPress(KeyAscii As Integer)
    Call cValidate.MaskedText(txt_fphone2, PhoneProfile)
End Sub

Private Sub txt_fphone2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_hphone_DblClick()
    PhoneNumber = txt_hphone
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub txt_hphone_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode > 36 Then cValidate.MaskedText txt_hphone, PhoneProfile
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_hphone2_DblClick()
    PhoneNumber = txt_hphone2
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub txt_hphone2_KeyPress(KeyAscii As Integer)
    
    Call cValidate.MaskedText(txt_hphone2, PhoneProfile)
    
End Sub

Private Sub txt_hphone2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_last_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_last
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_memo_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_memo
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_mi_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_mi_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_mphone_DblClick()
    PhoneNumber = txt_mphone
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub txt_mphone_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode > 36 Then cValidate.MaskedText txt_mphone, PhoneProfile
   If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_mphone2_DblClick()
    PhoneNumber = txt_mphone2
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub txt_mphone2_KeyPress(KeyAscii As Integer)
    Call cValidate.MaskedText(txt_mphone2, PhoneProfile)
End Sub

Private Sub txt_mphone2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_prefix_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_prefix
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_state_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_state_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_title_KeyUp(KeyCode As Integer, Shift As Integer)
 
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_title
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_u1phone_DblClick()
    PhoneNumber = txt_u1phone
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub


Private Sub txt_u1phone_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_u1text_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_u1text
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_u2phone_DblClick()
    PhoneNumber = txt_u2phone
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub txt_u2phone_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_u2text_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_u2text
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_u3phone_DblClick()
    PhoneNumber = txt_u3phone
    frmDIALER.Show vbModal, Me
    Unload frmDIALER
End Sub

Private Sub txt_u3phone_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub


Private Sub txt_u3text_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If chk_autocap.Value Then cValidate.FcaseTextBox txt_u3text
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Private Sub txt_zip_KeyUp(KeyCode As Integer, Shift As Integer)
    
    cValidate.MaskedText txt_zip, ZipProfile
    If KeyCode = 46 And Not cmdSave.Enabled Then Call ShowButtons(False)

End Sub

Public Sub SortData()

    On Error GoTo Err_Proc

  
    Select Case SortBy
    Case "COMPANY"
        SortQuery = "SELECT * FROM ADDBOOK ORDER BY [COMPANY],[LASTNAME],[FIRSTNAME]"
        cmdSortBy.Caption = "Sort By Company"
        SortField = "COMPANY"
        SortOther = vbNullString
    Case "BUSINESS"
        SortQuery = "SELECT * FROM ADDBOOK WHERE [BUSINESS] = TRUE ORDER BY [COMPANY],[LASTNAME],[FIRSTNAME]"
        cmdSortBy.Caption = "Filter By Business"
        SortField = "COMPANY"
        SortOther = vbNullString
    Case "FRIEND"
        SortQuery = "SELECT * FROM ADDBOOK WHERE [FRIEND] = TRUE ORDER BY [LASTNAME],[FIRSTNAME]"
        cmdSortBy.Caption = "Filter By Friend"
        SortField = "LASTNAME"
        SortOther = vbNullString
    Case "RELATIVE"
        SortQuery = "SELECT * FROM ADDBOOK WHERE [RELATIVE] = TRUE ORDER BY [LASTNAME],[FIRSTNAME]"
        cmdSortBy.Caption = "Filter By Relative"
        SortField = "LASTNAME"
        SortOther = vbNullString
    Case "OTHER"
        If SortOther = vbNullString Then
            frmSelectOther.Show vbModal, Me
            SortOther = frmSelectOther!Combo1.Text
            If frmSelectOther!opt_last.Value Then
               SortQuery = "SELECT * FROM ADDBOOK WHERE [OTHER] = '" & SortOther & "' ORDER BY [LASTNAME],[FIRSTNAME]"
               SortField = "LASTNAME"
            Else
               SortQuery = "SELECT * FROM ADDBOOK WHERE [OTHER] = '" & SortOther & "' ORDER BY [COMPANY],[LASTNAME],[FIRSTNAME]"
               SortField = "COMPANY"
            End If
            Unload frmSelectOther
        Else
            If SortField = "LASTNAME" Then
               SortQuery = "SELECT * FROM ADDBOOK WHERE [OTHER] = '" & SortOther & "' ORDER BY [LASTNAME],[FIRSTNAME]"
            Else
               SortQuery = "SELECT * FROM ADDBOOK WHERE [OTHER] = '" & SortOther & "' ORDER BY [COMPANY],[LASTNAME],[FIRSTNAME]"
               SortField = "COMPANY"
            End If
        End If
        cmdSortBy.Caption = "Filter By " & SortOther
    
    Case Else '"LASTNAME"
        SortQuery = "SELECT * FROM ADDBOOK ORDER BY [LASTNAME],[FIRSTNAME],[COMPANY]"
        cmdSortBy.Caption = "Sort By Last Name"
        SortField = "LASTNAME"
        SortOther = vbNullString
    
    End Select
    
    Call OpenRS(m_Set, SortQuery, m_DB)
    Call UpdateData
    
    SaveSetting App.Title, "Options", "SortBy", SortBy
    SaveSetting App.Title, "Options", "SortOther", SortOther
    SaveSetting App.Title, "Options", "SortField", SortField
    

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "SortData"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub ShowButtons(bEnabled As Boolean)

    On Error GoTo Err_Proc

    cmdImportExport.Enabled = bEnabled
    cmd_new.Enabled = bEnabled
    cmdSave.Enabled = Not bEnabled
    cmdCancelNew.Enabled = Not bEnabled
    cmd_Cut.Enabled = bEnabled
    cmd_print.Enabled = bEnabled
    cmdDates.Enabled = bEnabled
    cmd_find.Enabled = bEnabled
    cmd_Dial.Enabled = bEnabled
    cmdSortBy.Enabled = bEnabled
    cmd_calander.Enabled = bEnabled
    cmd_app.Enabled = bEnabled
    cmd_About.Enabled = bEnabled
    picIndexButtons.Enabled = bEnabled
    
    cmd_next.Enabled = bEnabled
    cmd_prev.Enabled = bEnabled
    cmdGrid.Enabled = bEnabled
    
    'm_AddingNew = Not bEnabled
    

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "ShowButtons"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Function FixString(ByVal tString As String, Optional AddCR As Boolean = False)

    On Error GoTo Err_Proc

  Dim i As Integer
    
TryAgain:
    i = InStr(tString, "=0D=0A")
    If i > 0 Then
        If AddCR Then
            tString = left(tString, i - 1) & vbCrLf & Mid(tString, i + 6)
        Else
            tString = left(tString, i - 1) & " " & Mid(tString, i + 6)
        End If
        GoTo TryAgain
    End If
    
    FixString = tString
        

Exit_Proc:
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "FixString"
    Err.Clear
    Resume Exit_Proc

End Function

Public Function FixNote(ByVal tString As String)

    On Error GoTo Err_Proc

  Dim i As Integer
    
TryAgain:
    i = InStr(tString, vbCrLf)
    If i > 0 Then
        tString = left(tString, i - 1) & "=0D=0A" & Mid(tString, i + 2)
        GoTo TryAgain
    End If
    
    FixNote = tString


Exit_Proc:
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "FixNote"
    Err.Clear
    Resume Exit_Proc

End Function

Private Sub UpdateData()

    On Error GoTo Err_Proc


    txt_Company.Text = m_Set!Company & vbNullString
    txt_title.Text = m_Set!Title & vbNullString
    txt_bday.Text = m_Set!Birthdate & vbNullString
    txt_anv.Text = m_Set!Anniversary & vbNullString
    txt_prefix.Text = m_Set!Prefix & vbNullString
    txt_first.Text = m_Set!FirstName & vbNullString
    txt_mi.Text = m_Set!MiddleName & vbNullString
    txt_last.Text = m_Set!LastName & vbNullString
    txt_bphone.Text = m_Set!WorkPhone & vbNullString
    txt_mphone.Text = m_Set!MobalPhone & vbNullString
    txt_fphone.Text = m_Set!FaxNumber & vbNullString
    txt_hphone.Text = m_Set!HomePhone & vbNullString
    txt_EMail.Text = m_Set!EMail & vbNullString
    txt_baddress.Text = m_Set!WorkAddress & vbNullString
    txt_bcity.Text = m_Set!WorkCity & vbNullString
    txt_bstate.Text = m_Set!WorkState & vbNullString
    txt_bzip.Text = m_Set!WorkZipCode & vbNullString
    txt_memo.Text = m_Set!Note & vbNullString
    txt_bphone2.Text = m_Set!WorkPhone & vbNullString
    txt_u1text.Text = m_Set!User1Text & vbNullString
    txt_u1phone.Text = m_Set!User1Phone & vbNullString
    txt_mphone2.Text = m_Set!MobalPhone & vbNullString
    txt_u2text.Text = m_Set!User2Text & vbNullString
    txt_u2phone.Text = m_Set!User2Phone & vbNullString
    txt_fphone2.Text = m_Set!FaxNumber & vbNullString
    txt_u3text.Text = m_Set!User3Text & vbNullString
    txt_u3phone.Text = m_Set!User3Phone & vbNullString
    txt_hphone2.Text = m_Set!HomePhone & vbNullString
    txt_EMail2.Text = m_Set!EMail & vbNullString
    txt_address.Text = m_Set!HomeAddress & vbNullString
    txt_city.Text = m_Set!HomeCity & vbNullString
    txt_state.Text = m_Set!HomeState & vbNullString
    txt_zip.Text = m_Set!HomeZipCode & vbNullString
    chk_bus.Value = Abs(m_Set!Business)
    chk_rel.Value = Abs(m_Set!Relative)
    chk_fri.Value = Abs(m_Set!Friend)
    cboOther.Text = m_Set!Other & vbNullString
    
    txt_bday.ForeColor = vbBlack
    txt_anv.ForeColor = vbBlack
    txt_bphone.ForeColor = vbBlack
    txt_bphone2.ForeColor = vbBlack
    txt_mphone.ForeColor = vbBlack
    txt_mphone2.ForeColor = vbBlack
    txt_fphone.ForeColor = vbBlack
    txt_fphone2.ForeColor = vbBlack
    txt_hphone.ForeColor = vbBlack
    txt_hphone2.ForeColor = vbBlack
    txt_u1phone.ForeColor = vbBlack
    txt_u2phone.ForeColor = vbBlack
    txt_u3phone.ForeColor = vbBlack

    

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "UpdateData"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub SaveData()

    On Error GoTo Err_Proc

    m_Set!Company = Trim$(txt_Company.Text)
    
    m_Set!Birthdate = Trim$(txt_bday.Text)
    m_Set!Anniversary = Trim$(txt_anv.Text)
    
    m_Set!Title = Trim$(txt_title.Text)
    m_Set!Prefix = Trim$(txt_prefix.Text)
    m_Set!FirstName = Trim$(txt_first.Text)
    m_Set!MiddleName = Trim$(txt_mi.Text)
    m_Set!LastName = Trim$(txt_last.Text)
    
    m_Set!WorkPhone = Trim$(txt_bphone.Text)
    m_Set!MobalPhone = Trim$(txt_mphone.Text)
    m_Set!FaxNumber = Trim$(txt_fphone.Text)
    m_Set!HomePhone = Trim$(txt_hphone.Text)
    
    m_Set!EMail = Trim$(txt_EMail.Text)
    
    m_Set!WorkAddress = Trim$(txt_baddress.Text)
    m_Set!WorkCity = Trim$(txt_bcity.Text)
    m_Set!WorkState = Trim$(txt_bstate.Text)
    m_Set!WorkZipCode = Trim$(txt_bzip.Text)
    
    m_Set!HomeAddress = Trim$(txt_address.Text)
    m_Set!HomeCity = Trim$(txt_city.Text)
    m_Set!HomeState = Trim$(txt_state.Text)
    m_Set!HomeZipCode = Trim$(txt_zip.Text)
    
    m_Set!User1Text = Trim$(txt_u1text.Text)
    m_Set!User1Phone = Trim$(txt_u1phone.Text)
    
    m_Set!User2Text = Trim$(txt_u2text.Text)
    m_Set!User2Phone = Trim$(txt_u2phone.Text)
    
    m_Set!User3Text = Trim$(txt_u3text.Text)
    m_Set!User3Phone = Trim$(txt_u3phone.Text)
    
    m_Set!Business = CBool(chk_bus.Value)
    m_Set!Relative = CBool(chk_rel.Value)
    m_Set!Friend = CBool(chk_fri.Value)
    m_Set!Other = Trim$(cboOther.Text)
    
    If LenB(txt_memo.Text) > 0 Then
        m_Set!Note = txt_memo.Text
    End If
        
    m_Set.Update
    
    DoEvents

Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "SaveData"
    Err.Clear
    Resume Exit_Proc

End Sub
Private Sub mnuAll_Click()
    SortBy = "COMPANY"
    Call SortData
End Sub
Private Sub mnuBusiness_Click()
    SortBy = "BUSINESS"
    Call SortData
End Sub
Private Sub mnuFriend_Click()
    SortBy = "FRIEND"
    Call SortData
End Sub
Private Sub mnuLast_Click()
    SortBy = "LASTNAME"
    Call SortData
End Sub
Private Sub mnuOther_Click()
    SortBy = "OTHER"
    SortOther = vbNullString
    Call SortData
End Sub
Private Sub mnuRelative_Click()
    SortBy = "RELATIVE"
    Call SortData
End Sub
Private Sub mnuCurrent_Click()

    On Error GoTo Err_Proc

 Dim msg As String
 
    Screen.MousePointer = 11
    DoEvents
    
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 11

    If txt_Company.Text > vbNullString Then Printer.Print Tab(5); "Company: "; Tab(25); txt_Company.Text
    If txt_title.Text > vbNullString Then Printer.Print Tab(25); txt_title.Text
    
    msg = Trim(Me.txt_prefix.Text & " " & Me.txt_first.Text)
    msg = Trim(msg & " " & Me.txt_mi.Text) & " " & Me.txt_last.Text
    If msg > vbNullString Then Printer.Print Tab(25); msg
    
    If txt_baddress.Text > vbNullString Then
       Printer.Print Tab(5); "Business Address:";
       Printer.Print Tab(25); txt_baddress.Text
       Printer.Print Tab(25); Me.txt_bcity.Text & ", " & Me.txt_bstate.Text & "  " & Me.txt_bzip.Text
    End If

    If txt_bphone.Text > vbNullString Then Printer.Print Tab(5); "Business Phone: "; Tab(25); txt_bphone.Text: Printer.Print
    
    If txt_address.Text > vbNullString Then
       Printer.Print Tab(5); "Home Address:";
       Printer.Print Tab(25); txt_address.Text
       Printer.Print Tab(25); Me.txt_city.Text & ", " & Me.txt_state.Text & "  " & Me.txt_zip.Text
    End If

    If txt_hphone.Text > vbNullString Then Printer.Print Tab(5); "Home Phone: "; Tab(25); txt_hphone.Text: Printer.Print
    
    If txt_fphone.Text > vbNullString Then Printer.Print Tab(5); "Fax Phone: "; Tab(25); txt_fphone.Text
    
    If txt_u1phone.Text > vbNullString Then Printer.Print Tab(5); Me.txt_u1text.Text; ": "; Tab(25); txt_u1phone.Text
    
    If txt_u2phone.Text > vbNullString Then Printer.Print Tab(5); Me.txt_u2text.Text; ": "; Tab(25); txt_u2phone.Text
    
    If txt_u3phone.Text > vbNullString Then Printer.Print Tab(5); Me.txt_u3text.Text; ": "; Tab(25); txt_u3phone.Text
    
    If txt_anv.Text > vbNullString Then Printer.Print Tab(5); "Anniversary: "; Tab(25); txt_anv.Text
    
    If txt_bday.Text > vbNullString Then Printer.Print Tab(5); "Birthday: "; Tab(25); txt_bday.Text
    
    Printer.Print
    
    If txt_memo.Text > vbNullString Then Printer.Print Tab(5); "MEMO:": Printer.Print txt_memo.Text
    
    Printer.EndDoc
    
    Screen.MousePointer = 0


Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "mnuCurrent_Click"
    Err.Clear
    Resume Exit_Proc

End Sub
Private Sub mnuScreen_Click()

    On Error GoTo Err_Proc

    
    Screen.MousePointer = 11
    If Not PnlPhoneFront Then pnl_phone.Visible = False
    picToolBar.Visible = False
    cmd_calander.Visible = False
    cmd_app.Visible = False
    cmd_next.Visible = False
    cmd_prev.Visible = False
    cmd_Expand.Visible = False
    cmd_About.Visible = False
    DoEvents
    
    Me.PrintForm
    
    pnl_phone.Visible = True
    picToolBar.Visible = True
    cmd_calander.Visible = True
    cmd_app.Visible = True
    cmd_next.Visible = True
    cmd_prev.Visible = True
    cmd_Expand.Visible = True
    cmd_About.Visible = True
    If Not PnlPhoneFront Then pnl_phone.ZOrder False
    Screen.MousePointer = False


Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "mnuScreen_Click"
    Err.Clear
    Resume Exit_Proc

End Sub
Private Sub mnuExport_Click()

    On Error GoTo Err_Proc

  Dim Filename      As String
  Dim FF            As Integer
  Dim getString     As String
  Dim tString       As String
  Dim i             As Integer
  Dim vData(0 To 4) As String
  Dim X             As Integer
  Dim cFile         As clsFileOpenSave
  
  
    Set cFile = New clsFileOpenSave
    If Trim(txt_last) > vbNullString Then
        Filename = Trim(txt_first & " " & txt_last)
    ElseIf Trim(txt_Company) > vbNullString Then
        Filename = Trim(txt_Company)
    End If
    cFile.VBGetSaveFileName Filename, "Export VCard file", , "vCard(*.vcf)|*.vcf", , , , "vcf"
    Set cFile = Nothing
    
    If Filename = vbNullString Then Exit Sub
    
    FF = FreeFile
    Open Filename For Output As #FF
    
    Print #FF, "BEGIN:VCARD"
    Print #FF, "VERSION:2.1"
    Print #FF, "N:" & txt_last & ";" & txt_first & ";" & txt_mi & ";" & txt_prefix & ";"
    If Trim(txt_mi) > vbNullString Then
        Print #FF, "FN:" & txt_first & " " & txt_mi & ". " & txt_last
    Else
        Print #FF, "FN:" & txt_first & " " & txt_last
    End If
    Print #FF, "ORG:" & txt_Company
    Print #FF, "TITLE:" & txt_title
    Print #FF, "NOTE;ENCODING=QUOTED-PRINTABLE:" & FixNote(txt_memo)
    Print #FF, "TEL;WORK;VOICE:" & txt_bphone
    Print #FF, "TEL;HOME;VOICE:" & txt_hphone
    Print #FF, "TEL;CELL;VOICE:" & txt_mphone
    Print #FF, "TEL;WORK;FAX:" & txt_fphone
    Print #FF, "ADR;WORK;ENCODING=QUOTED-PRINTABLE:;;" & txt_baddress & ";" & txt_bcity & ";" & txt_bstate & ";" & txt_bzip & ";"
    Print #FF, "ADR;HOME;ENCODING=QUOTED-PRINTABLE:;;" & txt_address & ";" & txt_city & ";" & txt_state & ";" & txt_zip & ";"
    

    Print #FF, "EMAIL;PREF;INTERNET:" & txt_EMail
    Print #FF, "REV:20030407T144607Z"
    Print #FF, "End:VCARD"
    Close #FF


Exit_Proc:
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "MainMenu", "mnuExport_Click"
    Err.Clear
    Resume Exit_Proc

End Sub

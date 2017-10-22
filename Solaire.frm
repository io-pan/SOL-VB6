VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fGbl 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8790
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   16695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   16695
   Begin VB.Frame frmGen 
      Caption         =   "Analemme"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   6
      Left            =   2880
      TabIndex        =   108
      Top             =   2760
      Width           =   2535
      Begin VB.CommandButton btAnalemme 
         Caption         =   "Command1"
         Height          =   195
         Left            =   1800
         TabIndex        =   121
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox CentreYAna 
         Height          =   285
         Left            =   2760
         TabIndex        =   120
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox CentreXAna 
         Height          =   285
         Left            =   2760
         TabIndex        =   119
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox MultAna 
         Height          =   285
         Left            =   2760
         TabIndex        =   118
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox picAnaleme 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   3975
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   3915
         ScaleWidth      =   2175
         TabIndex        =   117
         Top             =   840
         Width           =   2235
      End
      Begin VB.Frame frmTxtAna 
         BorderStyle     =   0  'None
         Height          =   550
         Left            =   120
         TabIndex        =   109
         Top             =   240
         Width           =   2295
         Begin VB.CommandButton btShowAnalemme 
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Papyrus"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2000
            TabIndex        =   112
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtDeclin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   111
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtDecalH 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   110
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label13 
            Caption         =   "°"
            Height          =   255
            Left            =   720
            TabIndex        =   116
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label12 
            Caption         =   "Equ. Temps"
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   115
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Déclinaison"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   114
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "min"
            Height          =   255
            Left            =   1440
            TabIndex        =   113
            Top             =   240
            Width           =   375
         End
      End
   End
   Begin VB.Frame frmGen 
      Caption         =   "Positions Maximum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   5
      Left            =   120
      TabIndex        =   89
      Top             =   6120
      Width           =   2535
      Begin VB.TextBox txtAziSolstEte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton btHautTheo 
         Caption         =   "maj"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtSolSolstEte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtAziEqui 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtAziSolstHiver 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtSolEqui 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtSolSolstHivers 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Azimuth"
         Height          =   255
         Left            =   960
         TabIndex        =   107
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "°"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   106
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "°"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   105
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "°"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   104
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Hauteur"
         Height          =   255
         Left            =   1800
         TabIndex        =   103
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "°"
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   102
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "°"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   101
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "°"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   100
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "Solst. d'hiv"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   99
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Equinoxes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   98
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Solst. d'été"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   97
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame frmGen 
      Caption         =   "Journée"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   4
      Left            =   120
      TabIndex        =   75
      Top             =   4080
      Width           =   2535
      Begin VB.TextBox txtAzimuthL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton btJourn 
         BackColor       =   &H80000013&
         Caption         =   "maj"
         Height          =   315
         Left            =   1440
         MaskColor       =   &H0000C000&
         TabIndex        =   81
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtSolMax 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtDuree 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   250
         Width           =   1455
      End
      Begin VB.TextBox txtLever 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtMidi 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtCoucher 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Durée"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Midi"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "Couch."
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Lever 
         Caption         =   "Lever"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "°Sud"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   84
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "°Hori"
         Height          =   255
         Index           =   9
         Left            =   2040
         TabIndex        =   83
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame frmGen 
      Caption         =   "Position du Soleil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   120
      TabIndex        =   67
      Top             =   2760
      Width           =   2535
      Begin VB.TextBox txtAzimuth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   250
         Width           =   735
      End
      Begin VB.TextBox txtHauteur 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtHeureSolaire 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   490
         Width           =   975
      End
      Begin VB.CommandButton btPosSol 
         Caption         =   "maj"
         Height          =   255
         Left            =   1440
         TabIndex        =   68
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "°Hori"
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   74
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "°Sud"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   73
         Top             =   250
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Heure Solaire"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   72
         Top             =   250
         Width           =   1095
      End
   End
   Begin VB.Frame frmGen 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   38
      Top             =   1320
      Width           =   3375
      Begin VB.TextBox txtDateHeure 
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CheckBox chkHeureLock 
         Caption         =   "Check1"
         Height          =   255
         Left            =   960
         TabIndex        =   54
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkHeureEte 
         Caption         =   "H. d'été"
         Height          =   195
         Left            =   2400
         TabIndex        =   53
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtJJ 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtHeure 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   51
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox chkDateLock 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1320
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton bdAddMinute 
         Caption         =   ">"
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton bdSubHeure 
         Caption         =   "<"
         Height          =   255
         Left            =   2040
         TabIndex        =   48
         Top             =   360
         Width           =   135
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   960
         TabIndex        =   47
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton btAddDay 
         Caption         =   ">"
         Height          =   255
         Left            =   1440
         TabIndex        =   46
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton btSubMonth 
         Caption         =   "<"
         Height          =   255
         Left            =   600
         TabIndex        =   45
         Top             =   360
         Width           =   135
      End
      Begin VB.CommandButton btAddMonth 
         Caption         =   ">"
         Height          =   255
         Left            =   1680
         TabIndex        =   44
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton btSubDay 
         Caption         =   "<"
         Height          =   255
         Left            =   720
         TabIndex        =   43
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton btDate 
         Caption         =   "maj"
         Height          =   195
         Left            =   1080
         TabIndex        =   42
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Timer timerTps 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1920
         Top             =   1080
      End
      Begin VB.CommandButton btNow 
         BackColor       =   &H00D6C5BE&
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton btSubMinute 
         Caption         =   "<"
         Height          =   255
         Left            =   2160
         TabIndex        =   40
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton bdAddHeure 
         Caption         =   ">"
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "jour n°"
         BeginProperty Font 
            Name            =   "@Malgun Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   240
         TabIndex        =   57
         Top             =   720
         Width           =   615
      End
      Begin VB.Image datelock 
         Height          =   225
         Left            =   840
         Picture         =   "Solaire.frx":0000
         Top             =   60
         Width           =   240
      End
      Begin VB.Image ImageLock 
         Height          =   225
         Left            =   600
         Picture         =   "Solaire.frx":0089
         Top             =   1200
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImageUnLock 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   2280
         Picture         =   "Solaire.frx":0111
         Top             =   1200
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label16 
         Caption         =   "Heure  "
         Height          =   255
         Left            =   1560
         TabIndex        =   56
         Top             =   0
         Width           =   495
      End
      Begin VB.Image heurelock 
         Height          =   225
         Left            =   2160
         Picture         =   "Solaire.frx":019A
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.Frame frmGen 
      Caption         =   "Lieu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtGMT 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Ville 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtLatitude 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "44.98"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtLongitude 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "4.58"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "°"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   37
         Top             =   600
         Width           =   135
      End
      Begin VB.Label lb_Latitude 
         Caption         =   "Longitude"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   36
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "°"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   35
         Top             =   600
         Width           =   110
      End
      Begin VB.Label lb_Latitude 
         Caption         =   "Latitude"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "GMT"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   33
         Top             =   240
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton initCour 
      Caption         =   "Option1"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame frmGen 
      Caption         =   "Fréquence d'affichage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   2
      Left            =   3720
      TabIndex        =   58
      Top             =   120
      Width           =   3735
      Begin VB.CheckBox chkSolstEte 
         Caption         =   "Solstice d'été"
         Height          =   255
         Left            =   240
         TabIndex        =   128
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton btClsTrajSol 
         Caption         =   "Effacer"
         Height          =   255
         Left            =   2520
         TabIndex        =   127
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton btTrajSol 
         Caption         =   "Tracer"
         Height          =   255
         Left            =   2520
         TabIndex        =   126
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkSolstHiver 
         Caption         =   "Solstices"
         Height          =   255
         Left            =   1200
         TabIndex        =   125
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox chkEqui 
         Caption         =   "Equinoxes"
         Height          =   255
         Left            =   1200
         TabIndex        =   124
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox chkTrajHeure 
         Caption         =   "Heures"
         Height          =   195
         Left            =   240
         TabIndex        =   123
         Top             =   1560
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTrajTxt 
         Caption         =   "Textes"
         Height          =   195
         Left            =   240
         TabIndex        =   122
         Top             =   1920
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Timer timerQTps 
         Left            =   3840
         Top             =   0
      End
      Begin VB.CommandButton btFreqReelle 
         Caption         =   "Réelle"
         Height          =   255
         Left            =   2520
         TabIndex        =   62
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtQtps 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   61
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtStp 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Text            =   "10"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton btFrequRapide 
         Caption         =   "10 min./msec"
         Height          =   255
         Left            =   2520
         TabIndex        =   59
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "/"
         Height          =   255
         Left            =   1200
         TabIndex        =   65
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label26 
         Caption         =   "msec"
         Height          =   255
         Left            =   2040
         TabIndex        =   64
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lbFrequence 
         Caption         =   "min."
         Height          =   255
         Left            =   840
         TabIndex        =   63
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame frmGen 
      BorderStyle     =   0  'None
      Caption         =   "Trajectoire du Soleil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Index           =   7
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.CheckBox chkTrace 
         Caption         =   "Trace"
         Height          =   195
         Left            =   7800
         TabIndex        =   28
         Top             =   9240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtHeureSolaire 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   9840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame frmSimAH 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Caption         =   "&H00404000&&H00404000&"
         Height          =   2175
         Left            =   4800
         TabIndex        =   22
         Top             =   120
         Width           =   2175
         Begin VB.Shape lineAHSol 
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   105
            Left            =   960
            Shape           =   3  'Circle
            Top             =   1080
            Width           =   105
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00404000&
            Index           =   5
            X1              =   2160
            X2              =   0
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00404000&
            Index           =   1
            X1              =   1080
            X2              =   1080
            Y1              =   0
            Y2              =   2160
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   1080
            TabIndex        =   25
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "E"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   24
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "W"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   23
            Top             =   840
            Width           =   255
         End
         Begin VB.Image ImaSphereNord 
            Height          =   2175
            Left            =   0
            Picture         =   "Solaire.frx":0223
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.TextBox txtHauteur 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   9840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame frmSimHaut 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2175
         Left            =   2280
         TabIndex        =   16
         Top             =   120
         Width           =   2175
         Begin VB.Shape lineHautSol 
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   105
            Left            =   840
            Shape           =   3  'Circle
            Top             =   1200
            Width           =   105
         End
         Begin VB.Line lineHaut 
            BorderColor     =   &H000000FF&
            X1              =   1080
            X2              =   1920
            Y1              =   1080
            Y2              =   360
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00404000&
            Index           =   4
            X1              =   1080
            X2              =   0
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00404000&
            Index           =   0
            X1              =   1080
            X2              =   1080
            Y1              =   0
            Y2              =   2160
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1080
            TabIndex        =   19
            Top             =   0
            Width           =   255
         End
         Begin VB.Line Line5 
            BorderColor     =   &H8000000D&
            X1              =   840
            X2              =   1320
            Y1              =   840
            Y2              =   1320
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   2640
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   18
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   17
            Top             =   840
            Width           =   255
         End
         Begin VB.Line Line8 
            Visible         =   0   'False
            X1              =   1080
            X2              =   1080
            Y1              =   0
            Y2              =   2640
         End
         Begin VB.Image ImaSphereEst 
            Height          =   2175
            Left            =   0
            Picture         =   "Solaire.frx":3733
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.Frame frmSimAzi 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   2175
         Left            =   7320
         TabIndex        =   9
         Top             =   120
         Width           =   2175
         Begin VB.Shape lineAziSol 
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   105
            Left            =   1080
            Shape           =   3  'Circle
            Top             =   1080
            Width           =   105
         End
         Begin VB.Line lineAzi 
            BorderColor     =   &H000000FF&
            X1              =   1200
            X2              =   1920
            Y1              =   1080
            Y2              =   240
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00404000&
            Index           =   2
            X1              =   1080
            X2              =   1080
            Y1              =   1080
            Y2              =   2160
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   1080
            TabIndex        =   14
            Top             =   840
            Width           =   255
         End
         Begin VB.Line Line4 
            BorderColor     =   &H8000000D&
            X1              =   720
            X2              =   1455
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "W"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1920
            TabIndex        =   13
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "E"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   12
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1080
            TabIndex        =   11
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1080
            TabIndex        =   10
            Top             =   1920
            Width           =   255
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00404000&
            Index           =   3
            X1              =   2160
            X2              =   0
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Image ImaSphereZenith 
            Height          =   2175
            Left            =   0
            Picture         =   "Solaire.frx":6C43
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.TextBox txtAzimuth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   8400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox picTrajSol 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000014&
         Height          =   4935
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   4875
         ScaleWidth      =   9345
         TabIndex        =   1
         Top             =   2760
         Width           =   9405
         Begin VB.Shape rondSolTraj 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00008080&
            Height          =   255
            Left            =   480
            Shape           =   3  'Circle
            Top             =   1800
            Width           =   255
         End
      End
      Begin VB.CommandButton btPleinEcran 
         BackColor       =   &H80000013&
         Caption         =   "^"
         Height          =   255
         Left            =   6840
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox MultTraj 
         Height          =   285
         Left            =   3600
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox centreYTraj 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox centreXTraj 
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "H. Solaire"
         Height          =   255
         Index           =   0
         Left            =   6840
         TabIndex        =   27
         Top             =   9840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Hauteur"
         Height          =   255
         Left            =   4320
         TabIndex        =   21
         Top             =   8160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Azimuth"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   15
         Top             =   8400
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Image ImaSphereBN 
      Height          =   2175
      Left            =   -240
      Picture         =   "Solaire.frx":A3BF
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image ImaSphereBJ 
      Height          =   2175
      Left            =   -240
      Picture         =   "Solaire.frx":DB25
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image ImaSphereN 
      Height          =   2175
      Left            =   -240
      Picture         =   "Solaire.frx":112A1
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image ImaSphereJ 
      Height          =   2175
      Left            =   -240
      Picture         =   "Solaire.frx":14841
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "¤"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "fGbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''
'''''''''''''''''''''''''''''
''- Private WithEvents vpb As LEGOVPBrickLib.VPBrick
Private Port$  'current port
Private crlf As String


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" _
  (ByVal hwnd As Long, ByVal _
  hWndInsertAfter As Long, ByVal X As Long, ByVal Y As _
  Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags _
  As Long) As Long
'-----------------------------------------------
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1

Private Const InclinaisonTerre = 23.45
Private colAnalemmes As Collection
Private frmCol As Collection 'Yo;Ho;Hm;Xo;Wo

Private Const vLieu = "0;255;1335;120;2535"
Private Const vDate = "480;615;1095;120;2535"
Private Const vPosS = "960;255;975;120;2535"
Private Const vJour = "840;255;1575;120;2535"
Private Const vTheo = "4320;255;1455;120;2535"
Private Const vAnal = "4320;255;4695;120;2535"
Private Const vFreq = "4320;255;975;120;2535"
Private Const vTraj = "0;255;5535;2760;6735"

Private iStpTrace As Integer
Private iTraceAzi
Private MaxTraceAzi
Private colOAzi As Collection
Private iTraceHaut
Private MaxTraceHaut
Private colOHaut As Collection
Private iTraceAH
Private MaxTraceAH
Private colOAH As Collection

Private chemin


Function InitfrmCol()
    Set frmCol = New Collection
    frmCol.Add vLieu   ' lieu
    frmCol.Add vDate   ' Date
    frmCol.Add vFreq   ' Fréquence d'affichage
    frmCol.Add vPosS   ' position soleil
    frmCol.Add vJour   ' journée
    frmCol.Add vTheo   ' hauteurs théoriques
    frmCol.Add vAnal   ' Analemme
    frmCol.Add vTraj   ' Traj Soleil
End Function

Private Sub incrHeureDate(periode, n As Integer)
    
    ' Incrémentation de l'heure
    txtDateHeure = txtDate & " " & txtHeure
    txtDateHeure = DateAdd(periode, n, txtDateHeure)
    
    If chkHeureLock = 0 Then
       txtHeure = TimeValue(txtDateHeure)
    End If
    If chkDateLock = 0 Then
       txtDate = DateValue(txtDateHeure)
    End If

End Sub
Private Sub bdAddHeure_Click()
    Call incrHeureDate("h", 1)
End Sub

Private Sub bdAddMinute_Click()
    Call incrHeureDate("n", 1)
End Sub

Private Sub btFrequRapide_Click()
    txtStp = 10
    iStpTrace = 10
    txtQtps = 1
    
    timerQTps.Enabled = False
        timerQTps.interval = txtQtps
    timerQTps.Enabled = True
End Sub

Private Sub btSubMinute_Click()
    Call incrHeureDate("n", -1)
End Sub

Private Sub bdSubHeure_Click()
    Call incrHeureDate("h", -1)
End Sub

Private Sub btAddDay_Click()
    Call incrHeureDate("d", 1)
End Sub

Private Sub btAddMonth_Click()
    Call incrHeureDate("m", 1)
End Sub

Private Sub btSubDay_Click()
    Call incrHeureDate("d", -1)
End Sub

Private Sub btSubMonth_Click()
    Call incrHeureDate("m", -1)
End Sub



Private Sub btHautTheo_Click()
    txtSolEqui = 90 - txtLatitude
    txtSolSolstEte = txtSolEqui + InclinaisonTerre
    txtSolSolstHivers = txtSolEqui - InclinaisonTerre
    txtAziEqui = VBA.Format(Round(AzimuthLever(JourJ(DateValue("20.03")), txtLatitude), 2), "00.00")
    txtAziSolstEte = VBA.Format(Round(AzimuthLever(JourJ(DateValue("21.06")), txtLatitude), 2), "00.00")
    txtAziSolstHiver = VBA.Format(Round(AzimuthLever(JourJ(DateValue("21.12")), txtLatitude), 2), "00.00")
End Sub

Private Sub btJourn_Click()
    txtMidi.Text = VBA.Format(Midi(JourJ(DateValue(txtDate.Text)), txtLongitude.Text, txtGMT, Hete()), "hh:nn")
    txtCoucher.Text = VBA.Format(Coucher(JourJ(DateValue(txtDate.Text)), txtLatitude.Text, txtLongitude.Text, txtGMT.Text, Hete()), "hh:nn")
    txtLever.Text = VBA.Format(LeverS(JourJ(DateValue(txtDate.Text)), txtLatitude.Text, txtLongitude.Text, txtGMT.Text, Hete()), "hh:nn")
    txtDuree.Text = VBA.Format(Duree(JourJ(DateValue(txtDate.Text)), txtLatitude.Text, txtLongitude.Text), "hh:nn")
    txtAzimuthL.Text = VBA.Round(AzimuthLever(JourJ(DateValue(txtDate.Text)), txtLatitude.Text), 2)
    txtSolMax.Text = VBA.Format(deg(hauteurSolMax(txtJJ, txtLatitude.Text)), "00.00")
End Sub

Private Sub initPositionsShapesCadrans()
    lineAzi.X1 = frmSimAzi.Width / 2 - 1
    lineAzi.Y1 = frmSimAzi.Height / 2 - 1
    lineAzi.X2 = frmSimAzi.Width / 2 - 1
    lineAzi.Y2 = frmSimAzi.Height / 2 - 1
    lineAziSol.Top = frmSimAzi.Width / 2 - lineAziSol.Width / 2
    lineAziSol.Left = frmSimAzi.Width / 2 - lineAziSol.Width / 2
    
    lineHaut.X1 = frmSimHaut.Width / 2 - 1
    lineHaut.Y1 = frmSimHaut.Width / 2 - 1
    lineHaut.X2 = frmSimHaut.Width / 2 - 1
    lineHaut.Y2 = 0
    lineHautSol.Left = (frmSimHaut.Width / 2 - 1) - lineHautSol.Width / 2
    lineHautSol.Top = 0
End Sub

Private Sub btPosSol_Click()
  Dim Haut As Double
  Dim Azi As Double

    If txtHeure <> "" Then
        txtHeureSolaire(0) = VBA.Format(heureSolaire(txtJJ.Text, TimeValue(txtHeure.Text), txtLongitude.Text, txtGMT.Text, Hete()), "hh:nn")
        txtHeureSolaire(1).Text = txtHeureSolaire(0).Text
        
        Haut = deg(hauteurSol(txtJJ.Text, TimeValue(txtHeure.Text), txtLatitude.Text, txtLongitude.Text, txtGMT.Text, Hete()))
        txtHauteur(0).Text = VBA.Format(Round(Haut, 2), "00.00")
        txtHauteur(1).Text = txtHauteur(0).Text
        
        Azi = deg(Azimuth(txtJJ.Text, TimeValue(txtHeure.Text), txtLatitude.Text, txtLongitude.Text, txtGMT.Text, Hete()))
        txtAzimuth(0).Text = VBA.Format(Round(Azi, 2), "00.00")
        txtAzimuth(1).Text = txtAzimuth(0).Text
        
        Call updPosSol(txtAzimuth(0).Text, txtHauteur(0).Text)
        Call updPanneaux(txtAzimuth(0).Text, txtHauteur(0).Text)
    End If
End Sub


Private Sub Form_Resize()
   ' frmGenerales.Height = fGbl.Height
    'VScrollGen.Height = fGbl.Height
    'Call initScrollGen
End Sub
Private Sub initScrollGen()
'    Dim htot As Integer
'    htot = 0
'
'    Dim i As Integer
'    For i = 0 To 6
'        htot = htot + frmGen(i).Height
'    Next
'    If htot > fGbl.Height Then
'        VScrollGen.Enabled = True
'        VScrollGen.Max = htot - fGbl.Height
'        VScrollGen.Min = 0
'    Else
'        frmGenerales.Top = 5
'        VScrollGen.Enabled = False
'    End If
'    frmGenerales.Height = htot
End Sub

Private Sub timerQTps_Timer()
    Call simule
End Sub

Private Sub simule()
 Dim heure As Date
 Dim heureP As Date
 Dim interval As String
    
    If lbFrequence.Caption = "j." Then
        interval = "d"
    Else    'Minutes
        interval = "n"
    End If
    
    Call incrHeureDate(interval, iStpTrace)
    heure = txtDateHeure
    heureP = DateAdd(interval, -iStpTrace, heure)
    Call SimulTraj(heure, heureP)

errh:
End Sub

Private Sub SimulTraj(heure As Date, heureP As Date)
 Dim Ja As Double
 Dim Jap As Double
 Dim Xp As Double
 Dim Yp As Double
 Dim Xc As Double
 Dim Yc As Double
 Dim Xpg As Double
 Dim Ypg As Double
 Dim Xcg As Double
 Dim Ycg As Double
 Dim Latitude As Double
 Dim Longitude As Double
 Dim GMT As Integer
 Dim Decl As Double
 

    'init  à l'arrache
    Latitude = txtLatitude
    Longitude = txtLongitude
    GMT = txtGMT
   
    Ja = JourJ(heure)
    Jap = JourJ(heureP)

    Xp = deg(Azimuth(Jap, heureP, Latitude, Longitude, GMT, Hete()))
    Xpg = Xp * MultTraj + centreXTraj
    Yp = deg(hauteurSol(Jap, heureP, Latitude, Longitude, GMT, Hete()))
    Ypg = Yp * -MultTraj + centreYTraj
    
    Xc = deg(Azimuth(Ja, heure, Latitude, Longitude, GMT, Hete()))
    Xcg = Xc * MultTraj + centreXTraj
    Yc = deg(hauteurSol(Ja, heure, Latitude, Longitude, GMT, Hete()))
    Ycg = Yc * -MultTraj + centreYTraj
    
    Dim fon As Integer
    If chkTrajHeure = 1 Then
        If Hour(heureP) < Hour(heure) Then
            picTrajSol.Line (Xpg, 0)-(Xpg, picTrajSol.Height), vbGreen
            If chkTrajTxt = 1 Then
              fon = picTrajSol.FontSize
              picTrajSol.FontSize = picTrajSol.FontSize - 2
              picTrajSol.Font = vbBlack
              picTrajSol.CurrentX = (Xpg) - 100
              picTrajSol.CurrentY = centreYTraj
              picTrajSol.Print VBA.Format(Hour(heure), "00") & ":" & VBA.Format(Minute(heure), "00")
              picTrajSol.FontSize = fon
            End If
        End If
    End If
    If chkTrajTxt = 1 Then
        'Lever
        If Yp < 0 And Yc >= 0 Then
            picTrajSol.CurrentX = Xcg - 400
            picTrajSol.CurrentY = centreYTraj - 200
            picTrajSol.Font = vbBlack
            picTrajSol.Print VBA.Format(Hour(heure), "00") & ":" & VBA.Format(Minute(heure), "00")
        End If
        'midi
        If Yp > 0 Then
            If Yp < Yc And Yc >= deg(hauteurSol(Ja, DateAdd("n", 1, heure), Latitude, Longitude, GMT, Hete)) Then
                picTrajSol.CurrentX = Xcg
                picTrajSol.CurrentY = Ycg - 200
                picTrajSol.Font = vbBlack
                picTrajSol.Print VBA.Format(Hour(heure), "00") & ":" & VBA.Format(Minute(heure) - 1, "00")
            End If
        End If
        'Coucher
        If Yp > 0 And Yc <= 0 Then
            picTrajSol.CurrentX = Xcg
            picTrajSol.CurrentY = centreYTraj - 200
            picTrajSol.Font = vbBlack
            picTrajSol.Print VBA.Format(Hour(heure), "00") & ":" & VBA.Format(Minute(heure), "00")
        
         End If

    End If ' txt

    picTrajSol.Line (Xpg, Ypg)-(Xcg, Ycg), vbRed

    Xp = Xc
    Yp = Yc
    Xpg = Xcg
    Ypg = Ycg
        
End Sub

Private Sub updPanneaux(Azimu As Double, Haut As Double)
    
    'couleurs
    If Haut < 0 Then
        ImaSphereZenith.Picture = ImaSphereBN.Picture
        ImaSphereEst.Picture = ImaSphereN.Picture
        ImaSphereNord.Picture = ImaSphereN.Picture
        
        lineAziSol.FillColor = &HC0C0&
    Else
        ImaSphereZenith.Picture = ImaSphereBJ.Picture
        ImaSphereEst.Picture = ImaSphereJ.Picture
        ImaSphereNord.Picture = ImaSphereJ.Picture
        
        lineAziSol.FillColor = vbYellow
    End If
    If Azimu >= 0 Then
        lineHautSol.FillColor = vbYellow
    Else
        lineHautSol.FillColor = &HC0C0&
    End If
    If Azimu >= -90 And Azimu <= 90 Then
        lineAHSol.FillColor = &HC0C0&
    Else
        lineAHSol.FillColor = vbYellow
    End If
    
    ' Positions
    lineAzi.X2 = (Cos(rad(Azimu - 90)) * frmSimAzi.Width / 2) + frmSimAzi.Width / 2
    lineAzi.Y2 = (Sin(rad(Azimu - 90)) * frmSimAzi.Height / 2) + frmSimAzi.Height / 2
                lineAziSol.Width = Cos(rad(Haut - 90)) * 80 + 150
                lineAziSol.Height = lineAziSol.Width
    lineAziSol.Left = Cos(rad(Azimu - 90)) * Cos(rad(Haut)) * frmSimAzi.Width / 2 + frmSimAzi.Width / 2 - lineAziSol.Width / 2
    lineAziSol.Top = Sin(rad(Azimu - 90)) * Cos(rad(Haut)) * frmSimAzi.Height / 2 + frmSimAzi.Height / 2 - lineAziSol.Height / 2

    
    
    lineHaut.X2 = frmSimHaut.Width / 2 + Cos(rad(Haut)) * (frmSimHaut.Width / 2)
    lineHaut.Y2 = (frmSimHaut.Height - lineHaut.X1) - (Sin(rad(Haut)) * (frmSimHaut.Height - lineHaut.X1))
    lineHautSol.Width = Cos(rad(Azimu - 90)) * 80 + 150
    lineHautSol.Height = lineHautSol.Width
    
    Dim X2 As Double
    Dim Y2 As Double
    
    X2 = Cos(rad(Haut)) * Cos(rad(Azimu))
    Y2 = -Sin(rad(Haut))
    lineHautSol.Left = frmSimHaut.Width / 2 + X2 * (frmSimHaut.Width / 2) - lineHautSol.Height / 2
    lineHautSol.Top = (frmSimHaut.Width / 2) - (Sin(rad(Haut)) * (frmSimHaut.Width / 2)) - lineAHSol.Height / 2
    
    lineAHSol.Width = Cos(rad(Azimu + 180)) * 80 + 150
    lineAHSol.Height = lineAHSol.Width
    lineAHSol.Left = lineAziSol.Left    ' lineAHSol.Left = -Sin(rad(Azimu - 180)) * Cos(rad(Haut)) * frmSimAH.Width / 2 + frmSimAH.Width / 2 - lineAHSol.Width / 2
    lineAHSol.Top = (frmSimHaut.Width / 2) - (Sin(rad(Haut)) * (frmSimHaut.Width / 2)) - lineAHSol.Height / 2
    
    If chkTrace = 1 Then
        Call Trace(colOAzi, lineAziSol.Left + lineAziSol.Height / 2, lineAziSol.Top + lineAziSol.Height / 2, lineAziSol.FillColor)
        Call Trace(colOHaut, lineHautSol.Left + lineHautSol.Height / 2, lineHautSol.Top + lineHautSol.Height / 2, lineHautSol.FillColor)
        Call Trace(colOAH, lineAHSol.Left + lineAHSol.Height / 2, lineAHSol.Top + lineAHSol.Height / 2, lineAHSol.FillColor)
        If iTraceAzi < MaxTraceAzi Then
            iTraceAzi = iTraceAzi + 1
        Else
            iTraceAzi = 1
        End If
    End If
End Sub














Private Sub datelock_Click()
    If chkDateLock = 0 Then
        datelock.Picture = ImageLock.Picture
        chkDateLock = 1
    Else
        datelock.Picture = ImageUnLock.Picture
        chkDateLock = 0
    End If
    
    If chkDateLock = 0 And chkHeureLock = 1 Then
        lbFrequence.Caption = "j."
    Else
        lbFrequence.Caption = "min."
    End If
End Sub

Private Sub heurelock_Click()
    If chkHeureLock = 0 Then
        heurelock.Picture = ImageLock.Picture
        chkHeureLock = 1
        lbFrequence.Caption = "j."
    Else
        heurelock.Picture = ImageUnLock.Picture
        chkHeureLock = 0
        lbFrequence.Caption = "min."
    End If
End Sub



Private Sub MaskTraj_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picTrajSol_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub MaskTraj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picTrajSol_MouseMove(Button, Shift, X, Y)
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call sauveLieu
End Sub




Private Sub lbFrequence_Change()
    Call changeNbTraces
End Sub




Private Sub chkEqui_Click()
    Call btTrajSol_Click
End Sub

Private Sub chkSolstEte_Click()
    Call btTrajSol_Click
End Sub

Private Sub chkSolstHiver_Click()
    chkSolstEte = -chkSolstEte + 1
End Sub
    
    
Private Sub frmGen_DblClick(Index As Integer)
'    Dim Ho As Integer
'    Dim Hm As Integer
'    'Yo;Ho;Hm;Xo;Wo
'
'    Dim z() As String
'    z = Split(frmCol(Index + 1), ";")
'    Ho = z(1)
'    Hm = z(2)
'
'    Dim i As Integer
'    If frmGen(Index).Height = Hm Then
'        For i = Index + 1 To frmCol.Count - 1
'            If frmGen(i).Left = frmGen(i - 1).Left Then
'                frmGen(i).Top = frmGen(i).Top - (Hm - Ho)
'            End If
'        Next
'        frmGen(Index).Height = Ho
'    Else
'        For i = Index + 1 To frmCol.Count - 1
'            If frmGen(i).Left = frmGen(i - 1).Left Then
'                frmGen(i).Top = frmGen(i).Top + (Hm - Ho)
'            End If
'        Next
'        frmGen(Index).Height = Hm
'    End If
'    Call initScrollGen
End Sub
Private Sub frmGen_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim mx As Integer
    Dim my As Integer
    
    For i = 0 To frmGen.Count - 1
        frmGen(i).ForeColor = vbBlack
    Next
    frmGen(Index).ForeColor = vbBlue
    
    If txtAzimuth(1) <> txtAzimuth(0) Then
        txtAzimuth(1) = txtAzimuth(0)
    End If
    If txtHauteur(1) <> txtHauteur(0) Then
        txtHauteur(1) = txtHauteur(0)
    End If
    
    ' jour courant
    mx = min2deg(eqt(txtJJ)) * MultAna + CentreXAna
    my = Declinaison1(txtJJ) * -MultAna + CentreYAna
    picAnaleme.Line (mx, 0)-(mx, picAnaleme.Height), vbGreen
    picAnaleme.Line (0, my)-(picAnaleme.Width, my), vbGreen
    txtDeclin = VBA.Format(-(my - CentreYAna) / MultAna, "00.00")
    txtDecalH = VBA.Round((mx - CentreXAna) / min2deg(MultAna), 2)
End Sub














'''''''''''''''
''''''''''''''''
'''''''''''''''
''''''''''''''''


Private Sub btNow_Click()

    Call changeNbTraces
    txtDate = VBA.Date
    txtHeure = VBA.Time
   ' txtHeure = Format(txtHeure, "HH:mm")
    txtJJ.Text = JourJ(DateValue(txtDate.Text))
    
    btTrajSol_Click
    
End Sub




Private Sub cmdClose_Click()
  'vpb.Close
  'txtResult = "Ok"
End Sub

Private Sub cmdDownload_Click()
  'Dim nPos As Long
  
  'On Error GoTo except
  
  'Source$ = txtSource
  'vpb.Download Source$, , nPos
  'Exit Sub

'except:
'  txtSource.SelStart = nPos
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  txtSource.SetFocus
'  ErrHelp
End Sub

Private Sub cmdExecute_Click()
'  Dim nResult As Long
'  Dim nPos As Long
'  Dim result As Variant

'  On Error GoTo except

'  Source$ = txtSource
'  nResult = vpb.Execute(Source$, , nPos, result)
'  txtResult = "Result: " + Str(nResult)
'  If VarType(result) = vbByte + vbArray Then
'    If UBound(result) > LBound(result) Then
'      txtResult = txtResult & ", variant result: "
'      For i = LBound(result) To UBound(result)
'        txtResult = txtResult + Str(result(i)) + " "
'      Next i
'    End If
'  End If
'  Exit Sub
'
'except:
'  txtSource.SelStart = nPos
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  txtSource.SetFocus
'  ErrHelp
End Sub

Private Sub FindPort()
'  On Error GoTo except
'  vpb.FindPort Port$
'  txtResult = "Port: " + Port$
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub cmdFirmware_Click()
'  On Error GoTo except
'  If vpb.BrickType = "RCX2" Then
'    vpb.DownloadFirmware "firm0328.lgo"
'  Else
'    txtResult = "Invalid PBrick type"
'  End If
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub ErrHelp()
'  If chkHelp.Value = Checked Then
'    If err.HelpFile = "" Then
'      MsgBox "No Err.HelpFile"
'      Exit Sub
'    End If
'
'    With CommonDialog1
'      .HelpFile = err.HelpFile
'      .HelpContext = Val(err.HelpContext)
'      .HelpCommand = cdlHelpContext
'      .ShowHelp
'    End With
'  End If
End Sub

Private Sub cmdHelp_Click()
'  With CommonDialog1
'    .HelpFile = vpb.HelpFile
'    .HelpCommand = cdlHelpContents
'    .ShowHelp
'  End With
End Sub

Private Sub cmdMemMap_Click()
'  Dim subs, tasks, sounds, displays As Variant
'  Dim nMemTop As Long
'  Dim nMemLast As Long
'  Dim nDataStart As Long
'  Dim nDataLast As Long
'  Dim MM As String
'
'  On Error GoTo except
'
'  vpb.MemMap nMemLast, nMemTop, nDataStart, nDataLast, tasks, subs, sounds, displays
'
'  MM = MM & "Free RAM: " & Str(nMemTop - nMemLast)
'  MM = MM & " (MemLast: " & Str(nMemLast) & ", MemTop: " & Str(nMemTop) & ")" & crlf
'
'  If vpb.BrickType = Spybot Then
'    MM = MM & "Data: " & Str(nDataLast - nDataStart)
'    MM = MM & ", start: " & Str(nDataStart)
'    MM = MM & ", end: " & Str(nDataLast) & crlf
'  End If
'
'  If vpb.BrickType = RCX2 Then
'    MM = MM & "Free Data: " & Str(nMemLast - nDataLast) & " (" & Str(nMemLast - nDataLast) / 3 & " items)"
'    MM = MM & "  Logged Data: " & Str(nDataLast - nDataStart)
'    MM = MM & " (" & Str(nDataLast - nDataStart) / 3 & " items)" & crlf
'    MM = MM & "DataStart: " & Str(nDataStart) & ", DataLast: " & Str(nDataLast) & crlf
'  End If
'
'  If VarType(tasks) = vbLong + vbArray Then
'    For nSlot = LBound(tasks, 1) To UBound(tasks, 1)
'      MM = MM & crlf & "Tasks (slot " & Str(nSlot + 1) & "): "
'      For nTask = LBound(tasks, 2) To UBound(tasks, 2)
'        MM = MM + Str(tasks(nSlot, nTask)) + ","
'      Next nTask
'      MM = MM
'    Next nSlot
'  End If
'
'  If VarType(subs) = vbLong + vbArray Then
'    For nSlot = LBound(subs, 1) To UBound(subs, 1)
'      MM = MM & crlf & "Subs (slot " & Str(nSlot + 1) & "): "
'      For nSub = LBound(subs, 2) To UBound(subs, 2)
'        MM = MM + Str(subs(nSlot, nSub)) + ","
'      Next nSub
'      MM = MM
'    Next nSlot
'  End If
  
'  If VarType(sounds) = vbLong + vbArray Then
'    MM = MM & "Sounds: "
'    For nSlot = LBound(sounds, 1) To UBound(sounds, 1)
'      For nSound = LBound(sounds, 2) To UBound(sounds, 2)
'        MM = MM + Str(sounds(nSlot, nSound)) + ","
'      Next nSound
'      MM = MM & crlf
'    Next nSlot
'  End If
  
'  If VarType(displays) = vbLong + vbArray Then
'    MM = MM & "Displays: "
'    For nSlot = LBound(displays, 1) To UBound(displays, 1)
'      For nDisplay = LBound(displays, 2) To UBound(displays, 2)
'        MM = MM + Str(displays(nSlot, nDisplay)) + ","
'      Next nDisplay
'      MM = MM & crlf
'    Next nSlot
'  End If
  
'  txtSource = MM

'  Exit Sub

'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub OpenPort()
'  On Error GoTo except
'  vpb.Open Port$
'  txtResult = "Opened " + Port$
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub cmdOpen_Click()
'  On Error GoTo except
'  FindPort
'  OpenPort
'  GetStatus
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub cmdPath_Click()
'  Path$ = InputBox("Path?", "New header path", vpb.Path)
'  If Len(Path$) > 0 Then
'    vpb.Path = Path$
'  End If
End Sub

Private Sub cmdRetries_Click()
'  Dim nRetries As Integer
'  On Error GoTo except
'  nRetries = vpb.Retries(ExecuteRetries)
'  nRetries = Val(InputBox("Execute retries?", , nRetries))
'  vpb.Retries(ExecuteRetries) = nRetries
  
'  nRetries = vpb.Retries(DownloadRetries)
'  nRetries = Val(InputBox("Download retries?", , nRetries))
'  vpb.Retries(DownloadRetries) = nRetries
'  Exit Sub

'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub cmdStats_Click()
'  Dim nRequests As Long
'  Dim nFails As Long
'  Dim nAborts As Long
'  Dim nTxRequests As Long
'  Dim nTxFails As Long
'  Dim nRxRequests As Long
'  Dim nRxErrors As Long
'  On Error GoTo except
'
'  vpb.GetStatistics nRequests, nFails, nAborts, nTxRequests, nTxFails, nRxRequests, nRxErrors
'  txtResult = "Requests:" + Str(nRequests) + ", fails:" + Str(nFails) + ", aborts:" + Str(nAborts) + ", TxRequests:" + Str(nTxRequests) + ", TxFails:" + Str(nTxFails) + ", RxRequests:" + Str(nRxRequests) + ", RxErrors:" + Str(nRxErrors)
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub cmdStatus_Click()
'  GetStatus
End Sub

Private Sub GetStatus()
'  Dim nStatus As StatusResult
'  Dim nBrickType As BrickTypes
'
'  On Error GoTo except
'
'  txtResult = "Working..."
'  Refresh
'
'  nBrickType = vpb.Status(CheckBrickType)
'  SetBrickType (nBrickType)
'  nStatus = vpb.Status(BrickStatus)
'  Select Case nStatus
'    Case StatusReady
'      txtResult = GetBrickType(vpb.Status(CheckBrickType)) + " ready"
'      If nBrickType = RCX2 Then
'        comboProgramSlot = vpb.ProgramSlot
'        comboIR.ListIndex = vpb.BrickTxRange - ShortRange
'        txtSleep = Str(vpb.PowerDownTime)
'      End If
'    Case StatusBusy
'      txtResult = "Busy"
'    Case Downloading
'      txtResult = "Downloading"
'    Case NotOpened
'      txtResult = "Not opened"
'    Case NoTower
'      txtResult = "No tower"
'    Case BadTower
'      txtResult = "Bad tower"
'    Case NoBrick
'      txtResult = "No brick"
'    Case NoFirmware
'      txtResult = "No firmware"
'    Case BadBrickBattery
'      txtResult = GetBrickType(nBrickType) + " bad battery"
'    Case BrickMismatch
'      txtResult = GetBrickType(nBrickType) + " brick type mismatch, expecting " + GetBrickType(vpb.BrickType)
'      vpb.BrickType = nBrickType
'    Case BadComms
'      txtResult = "Bad comms"
    'Case Else
'      txtResult = "Status: " + Str(nStatus)
'  End Select

'  If vpb.PortType = USBTowerLink Then
'    comboUSB.ListIndex = vpb.PortTxRange - ShortRange
'  End If
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

''- Private Function SetBrickType(nBrickType As BrickTypes)
''-   Select Case nBrickType
''-     Case Scout
''-       comboBrickType = "Scout"
''-       vpb.BrickType = Scout
''-     Case Spybot
''-       comboBrickType = "Spybot"
''-       vpb.BrickType = Spybot
''-     Case RCX2
''-       comboBrickType = "RCX2"
''-       vpb.BrickType = RCX2
''-     Case MicroScout
''-       comboBrickType = "MicroScout"
''-       vpb.BrickType = MicroScout
''-     Case Else
''-       comboBrickType.Clear
''-   End Select
''- End Function

''- Private Function GetBrickType(nBrickType As BrickTypes) As String
''-   Select Case nBrickType
''-     Case RCXnoFirmware
''-       GetBrickType = "RCX (no firmware)"
''-     Case RCX
''-       GetBrickType = "RCX"
''-     Case Scout
''-       GetBrickType = "Scout"
''-     Case Spybot
''-       GetBrickType = "Spybot"
''-     Case RCX2
''-       GetBrickType = "RCX2"
''-     Case MicroScout
''-       GetBrickType = "MicroScout"
''-     Case Else
''-       GetBrickType = "unknown PBrick"
''-   End Select
''-End Function

Private Sub cmdTest_Click()
'  On Error GoTo except
'  nBrickType = vpb.Status(CheckBrickType)
'  SetBrickType (nBrickType)
'  nStatus = vpb.Status(BrickStatus)  'unlocks Scout
'  If nStatus = StatusReady Then
'    If nBrickType = Scout Then
'      vpb.Execute "sound 25"
'    Else
'      vpb.Execute "sound 3"
'    End If
'    txtResult = "Ok"
'  Else
'    txtResult = "Not ready"
'  End If
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub cmdTrace_Click()
'  Dim nTrace As Integer
'  On Error GoTo except
'
'  nTrace = vpb.Trace()
'  nTrace = Val(InputBox("Trace options (0-15)?", , nTrace))
'  vpb.Trace = nTrace
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub cmdUpload_Click()
'  Dim nStartAddress As Long
'  Dim nSize As Long
'  Dim memData As Variant
'  On Error GoTo except
'
'  nStartAddress = Val(InputBox("Start Address?"))
'  nSize = Val(InputBox("Bytes?"))
'  memData = vpb.Upload(nStartAddress, nSize)
'  If VarType(memData) = vbByte + vbArray Then
'    For i = LBound(memData) To UBound(memData)
'      txtSource = txtSource + Str(memData(i)) + ","
'    Next i
'  End If
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub comboBrickType_Click()
'  On Error GoTo except
'
'  If comboBrickType = "MicroScout" Then
'    If vpb.BrickType = Spybot Then
'      vpb.Close
'      Port$ = ""
'      vpb.BrickType = MicroScout
'    End If
'  ElseIf comboBrickType = "Scout" Then
'    If vpb.BrickType = Spybot Then
'      vpb.Close
'      Port$ = ""
'      vpb.BrickType = Scout
'    End If
'  ElseIf comboBrickType = "Spybot" Then
'    If vpb.BrickType <> Spybot Then
'      vpb.Close
'      Port$ = ""
'      vpb.BrickType = Spybot
'    End If
'  ElseIf comboBrickType = "RCX2" Then
'    If vpb.BrickType = Spybot Then
'      vpb.Close
'      Port$ = ""
'      vpb.BrickType = RCX2
'    End If
'  End If
'  Exit Sub

'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub comboIR_Click()
'  On Error GoTo except
'  If vpb.BrickType = RCX2 Then
'    vpb.BrickTxRange = comboIR.ListIndex + ShortRange
'  End If
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub comboProgramSlot_Click()
'  On Error GoTo except
'  If vpb.BrickType = RCX2 Then
'    vpb.ProgramSlot = comboProgramSlot
'  End If
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub comboUSB_Click()
'  On Error GoTo except
'  If vpb.PortType = USBTowerLink Then
'    vpb.PortTxRange = comboUSB.ListIndex + ShortRange
'  End If
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub





Private Sub frmDate_DblClick()
'    If frmDate.Height = 1095 Then
'        frmDate.Height = 495
'        frmJourn.Top = frmJourn.Top + (lieuHm - lieuHo)
'    Else
'        frmDate.Height = 1095
'
'    End If
End Sub


Private Sub frmLieu_DblClick()
'    If frmLieu.Height = lieuHm Then
'        frmLieu.Height = lieuHo
'        frmDate.Top = frmDate.Top - (lieuHm - lieuHo)
'        frmJourn.Top = frmJourn.Top - (lieuHm - lieuHo)
'    Else
'        frmLieu.Height = lieuHm
'        frmDate.Top = frmDate.Top + (lieuHm - lieuHo)
'        frmJourn.Top = frmJourn.Top + (lieuHm - lieuHo)
'    End If
End Sub



Private Sub frmTheorique_DblClick()
'    If frmTheorique.Height = 1215 Then
'        frmTheorique.Height = 255
'    Else
'        frmTheorique.Height = 1215
'    End If
End Sub


Private Sub frmTraj_DblClick()
'    If frmTraj.Height = 255 Then
'        frmTraj.Height = 3255
'    Else
'        frmTraj.Height = 255
'    End If
End Sub

Private Sub frmTraj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 1 To frmGen.Count - 1
        frmGen(i).ForeColor = vbBlack
    Next
    'frmTraj.ForecolOAzir = vbBlue
End Sub



Private Sub btFreqReelle_Click()
    txtStp = 1
    iStpTrace = 1
    txtQtps = 60000
    
    timerQTps.Enabled = False
        timerQTps.interval = txtQtps
    timerQTps.Enabled = True
End Sub
Private Sub timerTps_Timer()
   ' btNow.Caption = DateValue(Date) & "       " & Format(TimeValue(Time), "hh:nn")
End Sub
Private Sub chkTrace_Click()
    Call changeNbTraces
End Sub
Private Sub initTraces()
    ' trace
    
    'Dim tS As Shape
    Dim i As Integer
    Dim nom As String
    
    MaxTraceAzi = 1440
    For i = 1 To MaxTraceAzi
        nom = "tA" & i
        colOAzi.Add Controls.Add("VB.shape", nom, frmSimAzi)
        nom = "tH" & i
        colOHaut.Add Controls.Add("VB.shape", nom, frmSimHaut)
        nom = "tAH" & i
        colOAH.Add Controls.Add("VB.shape", nom, frmSimAH)
        
        colOAzi(i).BorderColor = &HFFFF&
        colOAzi(i).Shape = 3
        colOAzi(i).Top = 0
        colOAzi(i).Left = 0
        colOAzi(i).Width = 1
        colOAzi(i).Height = 1
        colOAzi(i).ZOrder (0)
        
        colOHaut(i).BorderColor = &HFFFF&
        colOHaut(i).Shape = 3
        colOHaut(i).Top = 0
        colOHaut(i).Left = 0
        colOHaut(i).Width = 1
        colOHaut(i).Height = 1
        colOHaut(i).ZOrder (0)
        
        colOAH(i).BorderColor = &HFFFF&
        colOAH(i).Shape = 3
        colOAH(i).Top = 0
        colOAH(i).Left = 0
        colOAH(i).Width = 1
        colOAH(i).Height = 1
        colOAH(i).ZOrder (0)
    Next
    iTraceAzi = 1
End Sub

Private Sub changeNbTraces()
    Dim i As Integer
    
    For i = 1 To 1440
        colOAzi(i).Visible = False
        colOHaut(i).Visible = False
        colOAH(i).Visible = False
    Next
    If lbFrequence = "min." Then
        MaxTraceAzi = Int(1440 / txtStp)
    Else 'jours
        MaxTraceAzi = Int(365 / txtStp)
    End If
    
    If iTraceAzi > MaxTraceAzi Then
        iTraceAzi = MaxTraceAzi
    End If
    
End Sub

Private Sub Trace(coll As Collection, X As Double, Y As Double, col As Double)
    coll(iTraceAzi).BorderColor = col
    coll(iTraceAzi).Left = X
    coll(iTraceAzi).Top = Y
    coll(iTraceAzi).Visible = True
End Sub
Private Sub Form_Load()

''-     Set vpb = New LEGOVPBrickLib.VPBrick
    Port$ = ""
''-     comboBrickType = "RCX2"
''-     crlf = Chr$(13) & Chr$(10)

    'Initialisations
    chemin = VBA.CurDir$ + "\"
    
    initCour = True
     iStpTrace = 1
    Call initPositionsShapesCadrans     ' cadrans
    Set colOAzi = New Collection
    Set colOHaut = New Collection
    Set colOAH = New Collection
    Call initTraces
    Call initTraj(0)
    Call InitfrmCol                     ' collection de frames Génerales
    Call initLieux                      ' dropdown lieux + sélection lieu
    
    txtDate = VBA.Date
    txtHeure = VBA.Time
        
    Call btClsTrajSol_Click
    Call btTrajSol_Click
    
    
    txtQtps = 60000
    timerQTps.interval = 60000
    timerQTps.Enabled = True
    txtStp = 1
    initCour = False
End Sub

Private Sub initLieux()
  Dim lieu
  lieu = ""
    
    Open chemin + "lieux.txt" For Input As #1
    Do While Not EOF(1)
        Input #1, lieu    'Print #1, htmlConc
        Ville.AddItem (lieu)
    Loop
    Ville.ListIndex = 1
    Close #1
    
    ''- Open chemin + "lieu.txt" For Input As #1
    ''-     Input #1, lieuxpardefaut 'Print #1, htmlConc
    ''- Close #1
    Ville.ListIndex = 1
End Sub
Private Sub sauveLieu()
    Open chemin + "lieu.txt" For Output As #1
        Print #1, Ville.ListIndex
    Close #1
End Sub
Private Sub Form_Terminate()
    
    'vpb.Close
End Sub



Private Sub txtHeure_Change()
    If Not txtHeure.ForeColor = vbBlue And Not txtHeure.ForeColor = vbRed Then
        Call btPosSol_Click
    End If
End Sub

Private Sub txtHeure_GotFocus()
    txtHeure.ForeColor = vbBlue
    txtHeure = VBA.Format(txtHeure, "hh:nn")
End Sub

Private Sub txtHeure_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         txtHeure_Validate (False)
    End If
End Sub

Private Sub txtHeure_LostFocus()
        txtHeure.ForeColor = vbBlack
End Sub

Private Sub txtHeure_Validate(Cancel As Boolean)
    If (txtHeure <> "") Then
        If InStr(txtHeure, ":") = 0 Then
             If Len(txtHeure) = 4 Then
                txtHeure = Left(txtHeure, 2) & ":" & VBA.Right(txtHeure, 2)
            ElseIf Len(txtHeure) = 2 Then
                txtHeure = Left(txtHeure, 2) & ":00"
            End If
        End If
        
        If IsDate(txtHeure) Then
           txtHeure = VBA.Format(txtHeure, "HH:nn")
        Else
            Cancel = True
        End If
    Else
       txtHeure = Hour(VBA.Time) & ":" & Minute(VBA.Time)
    End If
    
    If Cancel = False Then

        
        Call btPosSol_Click
        'txtdate.SetFocus
    Else
        txtHeure.ForeColor = vbRed
    End If
    
End Sub

Private Sub txtDate_Change()
    If Not txtDate.ForeColor = vbBlue And Not txtDate.ForeColor = vbRed Then
         If (chkHeureLock = 0) Then
            chkHeureEte = heureEte(DateValue(txtDate))
         End If
        txtJJ = JourJ(DateValue(txtDate))
        Call btJourn_Click
        Call btAnalemme_Click
        Call btPosSol_Click
    End If
End Sub

Private Sub txtDate_GotFocus()
    txtDate.ForeColor = vbBlue
    txtDate = VBA.Format(txtDate, "dd.mm")
End Sub

Private Sub txtDate_LostFocus()
    txtDate.ForeColor = vbBlack
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         txtDate_Validate (False)
    End If
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    If (txtDate <> "") Then
        If InStr(txtDate, ".") = 0 Then
            If Len(txtDate) = 4 Then
                txtDate = VBA.Left(txtDate, 2) & "." & VBA.Right(txtDate, 2)
            ElseIf Len(txtDate) = 2 Then
                txtDate = VBA.Left(txtDate, 2) & "." & Month(VBA.Date)
            End If
        End If
        
        If IsDate(txtDate) Then
           
        Else
            Cancel = True
        End If
    Else
       txtDate = VBA.Date
    End If
    
    If Cancel = False Then
         
         If (chkHeureLock = 0) Then
            chkHeureEte = heureEte(DateValue(txtDate))
         End If
        txtJJ = JourJ(DateValue(txtDate))
        Call btJourn_Click
        Call btAnalemme_Click
        'txtHeure.SetFocus
    Else
        txtDate.ForeColor = vbRed
    End If
End Sub


Private Sub txtQtps_GotFocus()
    txtQtps.ForeColor = vbBlue
End Sub

Private Sub txtQtps_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         txtQtps_Validate (False)
    End If
End Sub

Private Sub txtQtps_Validate(Cancel As Boolean)
    On Error GoTo err
        timerQTps.Enabled = False
            timerQTps.interval = txtQtps
        timerQTps.Enabled = True
        
        txtQtps.ForeColor = vbBlack
    Exit Sub
err:
    txtQtps.ForeColor = vbRed
    Cancel = True
End Sub

Private Sub txtStp_GotFocus()
    txtStp.ForeColor = vbBlue
End Sub

Private Sub txtStp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         txtStp_Validate (False)
    End If
End Sub

Private Sub txtStp_Validate(Cancel As Boolean)
    txtStp.ForeColor = vbBlack
    iStpTrace = txtStp
    Call changeNbTraces
End Sub

Private Sub Ville_Click()
    Dim txt As String
    
    txt = VBA.Right(Ville.Text, 19)
    txtGMT = VBA.Right(txt, 3)
    txtLatitude = VBA.Left(txt, 7)
    txtLongitude = VBA.Mid(txt, 9, 7)
    btHautTheo_Click
    
    If Not initCour Then
        btJourn_Click
        btPosSol_Click
        btTrajSol_Click
    End If
    
End Sub

Private Sub txtLatitude_Validate(Cancel As Boolean)
    
    On Error GoTo err
    If (txtLatitude > 90 Or txtLatitude < 0) Then
        GoTo err
    Else
        txtSolEqui = 90 - txtLatitude
        txtSolSolstEte = txtSolEqui + InclinaisonTerre
        txtSolSolstHivers = txtSolEqui - InclinaisonTerre
    End If
Exit Sub
err:
     MsgBox "Veuillez entrer une latitude correcte"
     Cancel = True
End Sub

Private Sub txtSleep_Change()
'  On Error GoTo except
'  vpb.PowerDownTime = Val(txtSleep)
'  Exit Sub
'
'except:
'  txtResult = "Error " + Hex$(err.Number) + " " + err.Description
'  ErrHelp
End Sub

Private Sub txtSource_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyF1 Then
'    help$ = txtSource.SelText
'    If Len(help$) = 0 Then
'      nStart = txtSource.SelStart
'      nEnd = txtSource.SelStart
'      While nStart > 0 And Mid$(txtSource, nStart + 1, 1) <> " "
'        nStart = nStart - 1
'      Wend
'      While nEnd < Len(txtSource) And Mid$(txtSource, nEnd + 1, 1) <> " "
'        nEnd = nEnd + 1
'      Wend
'      help$ = Mid$(txtSource, nStart + 1, nEnd - nStart)
'    End If
'    With CommonDialog1
'      .HelpFile = "VPB.hlp"
'      If Len(help$) = 0 Then
'        .HelpCommand = cdlHelpContents
'      Else
'        .HelpKey = help$
'        .HelpCommand = cdlHelpKey
'      End If
'      .ShowHelp
'    End With
'  End If
End Sub

Private Sub vpb_DownloadDone(ByVal nErrorCode As Long)
'  If nErrorCode = 0 Then
'    txtResult = "Downloaded ok"
'  Else
'    txtResult = "Download error: " + Hex$(nErrorCode)
'  End If
End Sub

Private Sub vpb_DownloadProgress(ByVal nPercent As Long)
'  txtResult = "Download progress: " + Str(nPercent) + "%"
End Sub















''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''

' A N A L E M M E

''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub btAnalemme_Click()
    

    If MultAna = "" Then
            Call initAnalemme(2)
        Else
            Dim i As Integer
            Dim mul As Double
            
            mul = MultAna
            MultAna = picAnaleme.Height / 90
            i = 0
            Do While mul > MultAna
                MultAna = MultAna * 2
                i = i + 1
            Loop
            Call initAnalemme(i)
        End If
        Call Analemme

End Sub

Private Sub initAnalemme(nbZoom As Integer)

    Dim mx As Double
    Dim my As Double
    
    CentreXAna = picAnaleme.Width / 2
    CentreYAna = picAnaleme.Height / 2
    MultAna = picAnaleme.Height / 90
    
    Dim i As Integer
    For i = 1 To nbZoom
        mx = min2deg(eqt(txtJJ)) * MultAna + CentreXAna
        my = Declinaison1(txtJJ) * -MultAna + CentreYAna
        Call zoomAnalemme(Round(mx), Round(my), 1)
    Next
End Sub

Private Sub zoomAnalemme(X As Single, Y As Single, zoom As Integer)

    If zoom = 1 Then
        If MultAna < 1600 Then
            MultAna = MultAna * 2
            CentreXAna = (2 * CentreXAna - X)
            CentreYAna = (2 * CentreYAna - Y)
        End If
    Else
        initAnalemme (2)
    End If

End Sub

Private Sub picAnaleme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = vbHourglass
    Call zoomAnalemme(X, Y, Button)
    Call Analemme
    Screen.MousePointer = vbDefault
End Sub

Private Sub picAnaleme_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mx As Double
    Dim my As Double
    Dim decal As Double
    mx = X
    my = Y
    txtDeclin = VBA.Format(-Round((my - CentreYAna) / MultAna, 2), "00.00")
    
    decal = ((mx - CentreXAna) / min2deg(MultAna))
    
'    -(Ho - Int(Ho))
   txtDecalH = Round((mx - CentreXAna) / min2deg(MultAna), 2)
End Sub

Private Sub Analemme()
    Dim j As Date
    Dim dj As Date
    Dim iJ As Double
    Dim decal As Double
        Dim X As Double
        Dim Y As Double
    
    picAnaleme.Cls
    ' Axes
    picAnaleme.Line (CentreXAna, 0)-(CentreXAna, picAnaleme.Height), vbBlue
    picAnaleme.Line (0, CentreYAna)-(picAnaleme.Width, CentreYAna), vbBlue
    
    'Analemme
    iJ = 1
    j = DateValue("01.01." & Year(VBA.Date))
    dj = DateValue("31.12." & Year(VBA.Date))
    Do While DateValue(j) <> DateValue(dj)
        'eqt
        X = min2deg(eqt(iJ)) 'affichage en degré
        X = X * MultAna + CentreXAna
        Y = Declinaison1(iJ) * -MultAna + CentreYAna
        picAnaleme.Line (X, Y)-(X + 20, Y), vbRed
        
        If Day(j) = 1 Then
            If X >= CentreXAna Then
                picAnaleme.Line (X, Y)-(X + 400, Y), vbBlack
                picAnaleme.CurrentX = X + 400
            Else
                picAnaleme.Line (X, Y)-(X - 400, Y), vbBlack
                picAnaleme.CurrentX = X - 800
            End If
            picAnaleme.CurrentY = Y + 10
            picAnaleme.Font = vbBlue
            picAnaleme.Print VBA.Format(j, "mmmm")
        End If
        
       'eqt0
        'X = min2deg(eqt0(iJ)) ' affichage en degré
        'X = X * MultAna + CentreXAna
        'Y = Declinaison(iJ) * -MultAna + CentreYAna
        'picAnaleme.Line (X, Y)-(X + 20, Y), vbBlue
                
        j = DateAdd("d", 1, j)
        iJ = iJ + 1
    Loop
   
   ' jour courant
    Dim mx As Double
    Dim my As Double
   
    iJ = JourJ(DateValue(fGbl.txtDate))
    mx = min2deg(eqt(iJ)) * MultAna + CentreXAna
    my = Declinaison1(iJ) * -MultAna + CentreYAna
    picAnaleme.Line (mx, 0)-(mx, picAnaleme.Height), vbGreen
    picAnaleme.Line (0, my)-(picAnaleme.Width, my), vbGreen
    txtDeclin = VBA.Format(-(my - CentreYAna) / MultAna, "00.00")
    
   ' decal = ((mx - CentreXAna) / (MultAna))
    txtDecalH = Round((mx - CentreXAna) / min2deg(MultAna), 2)
   'If decal >= 0 Then
   '     txtDecalH = Int(decal) & "' "
   '     decal = decal - Int(decal)
   ' Else
   '     decal = -decal
   '     txtDecalH = "-" & Int(decal) & "' "
   '     decal = decal - Int(decal)
   ' End If
   ' decal = Round(decal * 60)
   ' txtDecalH = txtDecalH & Format(decal, "00") & "''"
End Sub

Private Sub btShowAnalemme_click()
        Screen.MousePointer = vbHourglass
        Call initAnalemme(0)
        Call Analemme
        Screen.MousePointer = vbDefault
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'
'    T R A J E T   S O L E I L
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''

Private Sub btTrajSol_Click()
        Dim Ja As Double
        '
    If chkEqui Then
        Ja = JourJ(DateValue(20.03))
        Call TrajSol(Ja, txtLatitude, txtLongitude, txtGMT, Hete(), vbYellow, False)
        Ja = JourJ(DateValue(22.09))
        Call TrajSol(Ja, txtLatitude, txtLongitude, txtGMT, Hete(), vbYellow, False)
    End If
        
    If chkSolstEte Then
        Ja = JourJ(DateValue(21.06))
        Call TrajSol(Ja, txtLatitude, txtLongitude, txtGMT, Hete(), vbYellow, False)
    End If
    
    If chkSolstHiver Then
        Ja = JourJ(DateValue(21.12))
        Call TrajSol(Ja, txtLatitude, txtLongitude, txtGMT, Hete(), vbYellow, False)
    End If
    
    Call TrajSol(txtJJ, txtLatitude, txtLongitude, txtGMT, Hete(), vbRed, True)

End Sub
Private Sub initTraj(nbZoom As Integer)
    
    picTrajSol.Cls
    MultTraj = picTrajSol.Width / 360
    centreXTraj = picTrajSol.Width / 2
    centreYTraj = picTrajSol.Height / 1.25
    
    Dim i As Integer
    Dim Yc As Integer
    Dim Xc As Integer
    
    For i = 1 To nbZoom
        ' Position courante
        Yc = deg(hauteurSol(txtJJ, TimeValue(txtHeure), txtLatitude, txtLongitude, txtGMT, Hete()))
        Yc = Yc * MultTraj + centreYTraj
        Xc = deg(Azimuth(txtJJ, TimeValue(txtHeure), txtLatitude, txtLongitude, txtGMT, Hete()))
        Xc = Xc * MultTraj + centreXTraj
        Call zoomTraj(Round(Yc), Round(Yc), 1)
    Next
End Sub
Private Sub zoomTraj(X As Single, Y As Single, boutton As Integer)
    
    If (boutton = 1) Then 'click droit
        If MultTraj < 1600 Then
            MultTraj = MultTraj * 2
            centreXTraj = (2 * centreXTraj - X)
            centreYTraj = (2 * centreYTraj - Y)
        End If
    Else    'click gauche
        Call initTraj(0)
    End If
End Sub

Private Sub btClsTrajSol_Click()

    picTrajSol.Cls
    Dim Xc As Integer
    
    Xc = -90
    picTrajSol.Line (centreXTraj + MultTraj * Xc, 0)-(centreXTraj + MultTraj * Xc, picTrajSol.Height), vbBlue
    picTrajSol.CurrentX = centreXTraj + MultTraj * Xc + 50
    picTrajSol.CurrentY = 50
    picTrajSol.Font = vbBlue
    picTrajSol.Print "E"
     
     Xc = 0
    picTrajSol.Line (centreXTraj, 0)-(centreXTraj, picTrajSol.Height), vbBlue
    picTrajSol.CurrentX = centreXTraj + MultTraj * Xc + 50
    picTrajSol.CurrentY = 50
    picTrajSol.Font = vbBlue
    picTrajSol.Print "S"
    
    Xc = 90
    picTrajSol.Line (centreXTraj + (MultTraj) * Xc, 0)-(centreXTraj + MultTraj * Xc, picTrajSol.Height), vbBlue
    picTrajSol.CurrentX = centreXTraj + MultTraj * Xc + 50
    picTrajSol.CurrentY = 50
    picTrajSol.Font = vbBlue
    picTrajSol.Print "W"
    
    Xc = -180
    picTrajSol.Line (centreXTraj + MultTraj * Xc, 0)-(centreXTraj + MultTraj * Xc, picTrajSol.Height), vbBlue
    
    Xc = 180
    picTrajSol.Line (centreXTraj + MultTraj * Xc, 0)-(centreXTraj + MultTraj * Xc, picTrajSol.Height), vbBlue
    
    picTrajSol.Line (0, centreYTraj)-(picTrajSol.Width, centreYTraj), vbBlue
    picTrajSol.Line (0, centreYTraj - 90 * MultTraj)-(picTrajSol.Width, centreYTraj - 90 * MultTraj), vbBlue

    Dim i As Integer
    For i = 1 To 1440
        colOAzi(i).Visible = False
        colOHaut(i).Visible = False
        colOAH(i).Visible = False
    Next
End Sub


Sub TrajSol(Ja As Double, Latitude As Double, Longitude As Double, GMT As Integer, Hete As Integer, col As Double, txt As Boolean)
 Dim Xp As Double
 Dim Yp As Double
 Dim Xc As Double
 Dim Yc As Double
 Dim Xpg As Double
 Dim Ypg As Double
 Dim Xcg As Double
 Dim Ycg As Double
 Dim heure As Date
 Dim Ymidi As Double
 
    Screen.MousePointer = vbHourglass

    heure = TimeValue("0:0")
    Xp = deg(Azimuth(Ja, heure, Latitude, Longitude, GMT, Hete)) '+ 180
    Xpg = Xp * MultTraj + centreXTraj
    
    Yp = deg(hauteurSol(Ja, heure, Latitude, Longitude, GMT, Hete))
    Ypg = Yp * -MultTraj + centreYTraj
    Ymidi = 0
    
    Dim h As Integer
    Dim M As Integer
    For h = 0 To 23
        If txt And chkTrajHeure = 1 Then
            picTrajSol.Line (Xpg, 0)-(Xpg, picTrajSol.Height), vbGreen
        
        If txt And chkTrajTxt = 1 Then
            picTrajSol.CurrentX = (Xpg) - 100
            picTrajSol.CurrentY = centreYTraj
            picTrajSol.Font = vbBlack
            picTrajSol.Print h
        End If
        End If
       For M = 0 To 59 Step 1
        
            heure = TimeValue(h & ":" & M)
            
            Yc = deg(hauteurSol(Ja, heure, Latitude, Longitude, GMT, Hete))
            Ycg = Yc * -MultTraj + centreYTraj
            
            Xc = deg(Azimuth(Ja, heure, Latitude, Longitude, GMT, Hete))
            Xcg = Xc * MultTraj + centreXTraj
            picTrajSol.Line (Xpg, Ypg)-(Xcg, Ycg), col
            
            If txt And chkTrajTxt = 1 Then
                'Lever
                If Yp < 0 And Yc >= 0 Then
                    picTrajSol.CurrentX = Xcg - 400
                    picTrajSol.CurrentY = centreYTraj - 200
                    picTrajSol.Font = vbBlack
                    picTrajSol.Print VBA.Format(h, "00") & ":" & VBA.Format(M, "00")
                End If
                'midi (juste passé)
                If Yc < Yp And Yc > 0 Then
                    If Ymidi = 0 Then
                        Ymidi = Yc
                        picTrajSol.CurrentX = Xcg
                        picTrajSol.CurrentY = Ycg - 200
                        picTrajSol.Font = vbBlack
                        picTrajSol.Print VBA.Format(h, "00") & ":" & VBA.Format(M - 1, "00")
                    End If
                End If
                'Coucher
                If Yp > 0 And Yc <= 0 Then
                    picTrajSol.CurrentX = Xcg
                    picTrajSol.CurrentY = centreYTraj - 200
                    picTrajSol.Font = vbBlack
                    picTrajSol.Print VBA.Format(h, "00") & ":" & VBA.Format(M, "00")
                End If
            End If
            Xp = Xc
            Yp = Yc
            Xpg = Xcg
            Ypg = Ycg
        Next
    Next
        
    Screen.MousePointer = Default
End Sub

Private Sub updPosSol(Azi As Double, Haut As Double)
        ' Position courante sur trajet
        Dim Ycg
        Dim Xcg
        
        Ycg = Haut * -MultTraj.Text + centreYTraj.Text
        Xcg = Azi * MultTraj.Text + centreXTraj.Text
        rondSolTraj.Left = Xcg - rondSolTraj.Width / 2
        rondSolTraj.Top = Ycg - rondSolTraj.Height / 2
        
        txtAzimuth(0).Text = VBA.Format(Round(Azi, 2), "00.00")
        txtHauteur(0).Text = VBA.Format(Round(Haut, 2), "00.00")
        txtAzimuth(1).Text = VBA.Format(Round(Azi, 2), "00.00")
        txtHauteur(1).Text = VBA.Format(Round(Haut, 2), "00.00")
End Sub
Private Sub paysage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picTrajSol_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub paysage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picTrajSol_MouseUp(Button, Shift, X, Y)
End Sub
Private Sub picTrajSol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mx As Double
    Dim my As Double
    
    If MultTraj <> "" Then
        mx = X
        my = Y
        txtHauteur(1) = VBA.Format(Round((my - centreYTraj) / -MultTraj, 2), "00.00")
        txtAzimuth(1) = VBA.Format(Round((mx - centreXTraj) / MultTraj, 2), "00.00")
    End If
End Sub

Private Sub picTrajSol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ja As Double
    Screen.MousePointer = vbHourglass
    Call zoomTraj(X, Y, Button)
    picTrajSol.Cls

    Call btClsTrajSol_Click
    If chkEqui Then
        Ja = JourJ(DateValue(20.03))
        Call TrajSol(Ja, txtLatitude, txtLongitude, txtGMT, Hete(), vbYellow, False)
        Ja = JourJ(DateValue(22.09))
        Call TrajSol(Ja, txtLatitude, txtLongitude, txtGMT, Hete(), vbYellow, False)
    End If
        
    If chkSolstEte Then
        Ja = JourJ(DateValue(21.06))
        Call TrajSol(Ja, txtLatitude, txtLongitude, txtGMT, Hete(), vbYellow, False)
    End If
    
    If chkSolstHiver Then
        Ja = JourJ(DateValue(21.12))
        Call TrajSol(Ja, txtLatitude, txtLongitude, txtGMT, Hete(), vbYellow, False)
    End If
    
    Call TrajSol(txtJJ, txtLatitude, txtLongitude, txtGMT, Hete(), vbRed, True)
    Screen.MousePointer = vbDefault
    btPosSol_Click
End Sub

Private Sub btPleinEcran_Click()
'    z = Split(frmCol(fTraj + 1), ";")
'     'Yo;Ho;Hm;Xo;Wo
'    If z(3) <> frmGen(fTraj).Left Then
'         frmGen(fTraj).Left = z(3)
'         frmGen(fTraj).Top = z(0)
'         frmGen(fTraj).Width = z(4)
'         frmGen(fTraj).Height = z(2)
'         btPleinEcran.Caption = "^"
'         picTrajSol.SetFocus
'    Else
'         frmGen(fTraj).Left = 0
'         frmGen(fTraj).Top = 0
'         frmGen(fTraj).Width = fGbl.Width
'         frmGen(fTraj).Height = fGbl.Height
'         btPleinEcran.Caption = "_"
'         picTrajSol.SetFocus
'    End If
'    picTrajSol.Width = frmGen(fTraj).Width - 300
'    picTrajSol.Height = frmGen(fTraj).Height - 840
'
'    btPleinEcran.Left = frmGen(fTraj).Width - 300
'    Call picTrajSol_MouseUp(2, 0, 0, 0)
End Sub

Function Hete() As Integer
    '
    Hete = chkHeureEte * heureEte(txtJJ)
End Function

Private Sub VScrollGen_Change()
'    Dim htot As Integer
'    Dim i As Integer
'
'    htot = 0
'    For i = 0 To 6
'        htot = htot + frmGen(i).Height
'    Next
'    frmGenerales.Top = -VScrollGen
'    frmGenerales.Height = frmGenerales.Height + VScrollGen
'    picTrajSol.SetFocus
End Sub

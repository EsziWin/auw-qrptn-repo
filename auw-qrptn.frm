VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form form1 
   Caption         =   "Fõprogram"
   ClientHeight    =   8304
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   12072
   ControlBox      =   0   'False
   HelpContextID   =   1
   Icon            =   "auw-qrptn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8304
   ScaleWidth      =   12072
   Begin VB.ListBox List1 
      Height          =   2544
      Left            =   10800
      TabIndex        =   115
      Top             =   840
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   252
      Left            =   8160
      TabIndex        =   113
      Top             =   5640
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4200
      TabIndex        =   109
      Top             =   1200
      Width           =   732
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   5880
      TabIndex        =   107
      Text            =   "Text8"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   2292
      Left            =   6240
      TabIndex        =   95
      Top             =   5640
      Visible         =   0   'False
      Width           =   1572
      Begin VB.CommandButton Command18 
         Caption         =   "Command18"
         Height          =   372
         Left            =   720
         TabIndex        =   100
         Top             =   1560
         Width           =   492
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Command21"
         Height          =   372
         Left            =   120
         TabIndex        =   99
         Top             =   1560
         Width           =   372
      End
      Begin VB.ListBox List4 
         Height          =   240
         Left            =   120
         TabIndex        =   98
         Top             =   1200
         Width           =   852
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   372
         Left            =   120
         TabIndex        =   97
         Top             =   720
         Width           =   1332
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   372
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "frame9"
      Height          =   732
      Left            =   5880
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   1572
      Begin VB.CommandButton Command23 
         Caption         =   "Command23"
         Height          =   192
         Left            =   600
         TabIndex        =   102
         Top             =   360
         Width           =   492
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Command22"
         Height          =   252
         Left            =   360
         TabIndex        =   101
         Top             =   240
         Width           =   372
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Command17"
         Height          =   372
         Left            =   3720
         TabIndex        =   94
         Top             =   1200
         Width           =   852
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   492
         Left            =   3120
         TabIndex        =   93
         Top             =   360
         Width           =   1212
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Command16"
         Height          =   492
         Left            =   2400
         TabIndex        =   92
         Top             =   2040
         Width           =   1212
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Command16"
         Height          =   492
         Left            =   2280
         TabIndex        =   91
         Top             =   1440
         Width           =   1092
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Option11"
         Height          =   372
         Left            =   2880
         TabIndex        =   89
         Top             =   1800
         Width           =   972
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Option10"
         Height          =   492
         Left            =   1800
         TabIndex        =   88
         Top             =   1440
         Width           =   852
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Option9"
         Height          =   612
         Left            =   840
         TabIndex        =   87
         Top             =   1320
         Width           =   612
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   252
         Left            =   1320
         TabIndex        =   86
         Top             =   960
         Width           =   732
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   252
         Left            =   2520
         TabIndex        =   85
         Top             =   2880
         Width           =   612
         _ExtentX        =   1080
         _ExtentY        =   445
         _Version        =   327680
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Command15"
         Height          =   192
         Left            =   1560
         TabIndex        =   84
         Top             =   3480
         Width           =   732
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Command14"
         Height          =   252
         Left            =   720
         TabIndex        =   83
         Top             =   2520
         Width           =   732
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Command13"
         Height          =   192
         Left            =   720
         TabIndex        =   82
         Top             =   240
         Width           =   612
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         Height          =   252
         Left            =   120
         TabIndex        =   76
         Top             =   3120
         Width           =   1212
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Command11"
         Height          =   252
         Left            =   120
         TabIndex        =   75
         Top             =   2040
         Width           =   1332
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   252
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   1332
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   252
         Left            =   120
         TabIndex        =   70
         Top             =   720
         Width           =   1332
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   252
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   1332
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   252
         Left            =   120
         TabIndex        =   68
         Top             =   1200
         Width           =   1332
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   252
         Left            =   120
         TabIndex        =   67
         Top             =   1440
         Width           =   1332
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Command10"
         Height          =   252
         Left            =   120
         TabIndex        =   66
         Top             =   1680
         Width           =   1332
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   252
         Left            =   1680
         TabIndex        =   65
         Top             =   360
         Width           =   1212
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   252
         Left            =   1680
         TabIndex        =   64
         Top             =   600
         Width           =   1092
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   252
         Left            =   1680
         TabIndex        =   63
         Top             =   840
         Width           =   852
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   252
         Left            =   1680
         TabIndex        =   62
         Top             =   1080
         Width           =   972
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   252
         Left            =   1680
         TabIndex        =   61
         Top             =   1320
         Width           =   972
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check6"
         Height          =   252
         Left            =   1680
         TabIndex        =   60
         Top             =   1560
         Width           =   852
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check7"
         Height          =   252
         Left            =   1680
         TabIndex        =   59
         Top             =   1800
         Width           =   852
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check8"
         Height          =   252
         Left            =   1680
         TabIndex        =   58
         Top             =   2040
         Width           =   972
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Check9"
         Height          =   252
         Left            =   1680
         TabIndex        =   57
         Top             =   2280
         Width           =   972
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Check10"
         Height          =   252
         Left            =   1680
         TabIndex        =   56
         Top             =   2520
         Width           =   1092
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Check11"
         Height          =   252
         Left            =   2760
         TabIndex        =   55
         Top             =   360
         Width           =   1212
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Check12"
         Height          =   252
         Left            =   2760
         TabIndex        =   54
         Top             =   600
         Width           =   1092
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Check13"
         Height          =   252
         Left            =   2760
         TabIndex        =   53
         Top             =   840
         Width           =   1092
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Check14"
         Height          =   192
         Left            =   2760
         TabIndex        =   52
         Top             =   1080
         Width           =   1092
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Check15"
         Height          =   252
         Left            =   2760
         TabIndex        =   51
         Top             =   1320
         Width           =   1092
      End
      Begin VB.CheckBox Check16 
         Caption         =   "Check16"
         Height          =   252
         Left            =   2760
         TabIndex        =   50
         Top             =   1560
         Width           =   972
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Check17"
         Height          =   252
         Left            =   2760
         TabIndex        =   49
         Top             =   1800
         Width           =   1092
      End
      Begin VB.CheckBox Check18 
         Caption         =   "Check18"
         Height          =   252
         Left            =   2760
         TabIndex        =   48
         Top             =   2040
         Width           =   972
      End
      Begin VB.CheckBox Check19 
         Caption         =   "Check19"
         Height          =   252
         Left            =   2760
         TabIndex        =   47
         Top             =   2280
         Width           =   1092
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Check20"
         Height          =   252
         Left            =   2760
         TabIndex        =   46
         Top             =   2520
         Width           =   1092
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Check21"
         Height          =   252
         Left            =   3960
         TabIndex        =   45
         Top             =   360
         Width           =   1092
      End
      Begin VB.CheckBox Check22 
         Caption         =   "Check22"
         Height          =   252
         Left            =   3960
         TabIndex        =   44
         Top             =   600
         Width           =   972
      End
      Begin VB.CheckBox Check23 
         Caption         =   "Check23"
         Height          =   252
         Left            =   3960
         TabIndex        =   43
         Top             =   840
         Width           =   1092
      End
      Begin VB.CheckBox Check24 
         Caption         =   "Check24"
         Height          =   252
         Left            =   3960
         TabIndex        =   42
         Top             =   1080
         Width           =   972
      End
      Begin VB.CheckBox Check42 
         Caption         =   "Check42"
         Height          =   252
         Left            =   3960
         TabIndex        =   41
         Top             =   1320
         Width           =   972
      End
      Begin VB.CheckBox Check43 
         Caption         =   "Check43"
         Height          =   252
         Left            =   3960
         TabIndex        =   40
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   372
         Left            =   3960
         TabIndex        =   39
         Top             =   1920
         Width           =   972
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   372
         Left            =   3960
         TabIndex        =   38
         Top             =   2400
         Width           =   972
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   372
         Left            =   3960
         TabIndex        =   37
         Top             =   2880
         Width           =   972
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   252
         Left            =   5040
         TabIndex        =   34
         Top             =   360
         Width           =   852
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   252
         Left            =   5040
         TabIndex        =   33
         Top             =   600
         Width           =   972
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   372
         Left            =   5040
         TabIndex        =   32
         Top             =   840
         Width           =   852
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   252
         Left            =   5040
         TabIndex        =   31
         Top             =   1200
         Width           =   852
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option5"
         Height          =   252
         Left            =   5040
         TabIndex        =   30
         Top             =   1440
         Width           =   852
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   252
         Left            =   5040
         TabIndex        =   29
         Top             =   1680
         Width           =   972
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Option7"
         Height          =   252
         Left            =   5040
         TabIndex        =   28
         Top             =   1920
         Width           =   852
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option8"
         Height          =   252
         Left            =   5040
         TabIndex        =   27
         Top             =   2160
         Width           =   852
      End
      Begin VB.TextBox Text9 
         Height          =   288
         Left            =   4080
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   3360
         Width           =   1452
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
         Height          =   252
         Left            =   2640
         TabIndex        =   35
         Top             =   3360
         Width           =   972
         _ExtentX        =   1715
         _ExtentY        =   445
         _Version        =   327680
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   252
         Left            =   1560
         TabIndex        =   36
         Top             =   3000
         Width           =   852
         _ExtentX        =   1503
         _ExtentY        =   445
         _Version        =   327680
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   252
         Left            =   600
         TabIndex        =   106
         Top             =   360
         Width           =   252
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   252
         Left            =   960
         TabIndex        =   105
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         Height          =   132
         Left            =   600
         TabIndex        =   104
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   252
         Left            =   240
         TabIndex        =   103
         Top             =   240
         Width           =   252
      End
      Begin VB.Label Label20 
         Caption         =   "Label15"
         Height          =   252
         Left            =   1920
         TabIndex        =   90
         Top             =   960
         Width           =   612
      End
      Begin VB.Label Label14 
         Caption         =   "Label14"
         Height          =   372
         Left            =   4440
         TabIndex        =   81
         Top             =   120
         Width           =   732
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   252
         Left            =   3240
         TabIndex        =   80
         Top             =   240
         Width           =   732
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   252
         Left            =   5160
         TabIndex        =   79
         Top             =   2640
         Width           =   612
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   252
         Left            =   5160
         TabIndex        =   78
         Top             =   3000
         Width           =   732
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   252
         Left            =   5160
         TabIndex        =   77
         Top             =   2520
         Width           =   732
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   252
         Left            =   120
         TabIndex        =   74
         Top             =   2400
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   252
         Left            =   120
         TabIndex        =   73
         Top             =   2640
         Width           =   732
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   252
         Left            =   120
         TabIndex        =   72
         Top             =   2880
         Width           =   732
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Listák"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5640
      TabIndex        =   24
      Top             =   5040
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ellenõrzés"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5160
      TabIndex        =   23
      Top             =   5160
      Width           =   1452
   End
   Begin VB.TextBox Text7 
      Height          =   852
      Left            =   6720
      TabIndex        =   22
      Text            =   "Text7"
      Top             =   2400
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox Text6 
      Height          =   1092
      Left            =   7680
      TabIndex        =   21
      Text            =   "Text6"
      Top             =   2760
      Visible         =   0   'False
      Width           =   492
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Height          =   132
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   233
      _Version        =   327680
   End
   Begin ComctlLib.ProgressBar ProgressBar3 
      Height          =   252
      Left            =   8040
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   445
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar ProgressBar2 
      Height          =   252
      Left            =   5640
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   445
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   3480
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   445
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ListBox List3 
      Height          =   240
      Left            =   3480
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.ListBox List2 
      Height          =   240
      Left            =   3480
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      HelpContextID   =   1
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.TextBox Text4 
      Height          =   492
      Left            =   3480
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.TextBox Text3 
      Height          =   492
      Left            =   3480
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   8040
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   4092
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3852
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   6795
      _Version        =   327680
      ScrollBars      =   0
      Appearance      =   0
   End
   Begin VB.ListBox Info 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   432
      Left            =   8040
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   4092
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Karbantartása"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      HelpContextID   =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "A munka folytatása"
      Top             =   720
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kilépés"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      HelpContextID   =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Kilépés a programból"
      Top             =   720
      Width           =   1452
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Command24"
      Height          =   372
      Left            =   5280
      TabIndex        =   110
      Top             =   1200
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label21 
      Caption         =   "Szállítólevelek"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   10800
      TabIndex        =   116
      Top             =   600
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label29 
      Height          =   372
      Left            =   8280
      TabIndex        =   114
      Top             =   6120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label Label22 
      Height          =   252
      Left            =   8280
      TabIndex        =   112
      Top             =   5040
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label25 
      Height          =   132
      Left            =   8640
      TabIndex        =   111
      Top             =   3600
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Mai dátum:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2760
      TabIndex        =   108
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   264
      Left            =   4920
      TabIndex        =   19
      Top             =   750
      Visible         =   0   'False
      Width           =   84
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Raktár:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3360
      TabIndex        =   18
      Top             =   750
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      ForeColor       =   &H0003178B&
      Height          =   360
      Left            =   720
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   8160
      Width           =   10212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   7920
      Width           =   10212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   6612
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' A mukalap számot a pszbrec$ 111 pozícióján tárolta és ha azt felülírták
' számlázás közben, akkor a számlaszám nem írodott át a munkalapba.
' Ezért munkalapról számlázáskor ne csak a Text8-ba írja be munkalapszámot,
' hanem egy másik komponensbe is, és ezt a komponenst mentse a pszbrec 171. pozíciójába.
' Sztornónál innen tudja a program, hogy munkalapról készült
' Több szállítóról készülhet egy számla
' Nyugel1 8 oszlopába került a raktár kód, munkalapról számla, munkalapról szállító, majd számla

Dim dbxneve$, objneve$, param$(10)
Dim kcikkek$(2000), kcikkcim&(2000), kegysar@(2000), cikkekdb%
Dim szalrend@(2000), vevorend@(2000), kkeszlet@(2000)
Dim orendel@(2000, 7), cikknev$(2000)
Dim fr4hed$(50), fr4fej$(50), fr4sor$(50), fr4lab$(50), fr4tra$(50), fmezok$(200), afatomb@(6, 2), rafatomb@(6, 2)
Dim afakodok$(5), afaalapok@(5), afaosszegek@(5), afaszamlak$(5)
Dim ellszla$(50), ellosszegek@(50), devellosszegek@(50), elto$(50), krak$, kertmoz$
Dim regimt$(1001), gysz$(200), hiv$(200), ntafa$(10)
     
Dim nevcimt$(), ncazonosito$()

Public ncdb&
Dim wrtt$(200)


Public sztornoszamla%, sztornomasolat%, masolat%, sztornoszamlaX%
Public ttkrak$, kmegrec$, stax%, exparvolt%, torlnyugta%, rnyugtavolt%, nyugtavolt%, munkalap$, megrendelesbol, megrend$, szallitobol, szallito$, szarmaztatasbol
Public kerafakod, kerbev, kerraf, elolegprefix$, szamlaszam$, rpartner$

Const KPPARTN As String = " KP"
Const BANKPARTN As String = "BANK"
Private Sub KPAR_Tolt()


If Trim$(Nyugel1.Text4(1)) = "" Or Trim$(Nyugel1.Text4(1)) = "BANK" Or Trim$(Nyugel1.Text4(1)) = "BANK" Then

             van = False
             ncv$ = Mid$(Nyugel1.Text2 + Space(60), 1, 60) + Mid$(Nyugel1.Text3 + Space(60), 1, 60)
             For j1% = 1 To ncdb&
               If nevcimt$(j1%) = ncv$ Then
                 van = True
                 Exit For
               End If
             Next
             If Not van Then
               
               kpar$ = Space(200)
               Mid$(kpar$, 1, 60) = Nyugel1.Text2
               Mid$(kpar$, 61, 60) = Nyugel1.Text3
               Call dbxki("KPAR", kpar$, ";", "U", "G", hiba%)
               Call nevcimtolt(ncv$, Mid$(kpar$, 190, 7))
             End If
End If
End Sub
             

Private Function nincs_tetelsor()
  tetdar% = 0
  For i1% = 1 To 200
    tkod$ = Left(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) + Space(15), 15)
    If Trim(tkod$) <> "" Then
       tetdar% = tetdar% + 1
    End If
  Next
  If tetdar% = 0 Then
    nincs_tetelsor = True
  Else
    nincs_tetelsor = False
  End If
End Function
Private Sub szallsztor(szallito$, erbiz)
For i% = 1 To List1.ListCount
szallito$ = List1.List(i% - 1)

kszbszlrec$ = dbxkey("KSZB", szallito$)
If Not kszbszlrec$ = "" Then
  Call dbxtrkezd("KSZB")
  Mid$(kszbszlrec$, 11, 10) = "Számlázva"
  Mid$(kszbszlrec$, 35, 1) = "S"
  Mid$(kszbszlrec$, 36, 6) = maidatum$
  Mid$(kszbszlrec$, 42, 8) = terminal$ + ugyintezo$
  Mid$(kszbszlrec$, 282, 15) = erbiz
  Call dbxki("KSZB", kszbszlrec$, ";", "", "", hiba%)
  Call dbxtrvege
  
  trdarab% = Val(Mid$(kszbszlrec$, 50, 3))
  szamla2$ = Mid$(kszbszlrec$, 1, 10)
  trind$ = szamla2$ + "001"
  psztrec$ = dbxkey("KSZT", trind$)
  If psztrec$ <> "" Then
     w1% = obsorszama("KSZT")
     kezdoix& = OBJTAB(w1%).obind
     dbfi = FreeFile
     Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #dbfi
     ndfi = FreeFile
     Open auditorutvonal$ + "auw-kszt.ndx" For Binary Shared As #ndfi
     rc& = Int(LOF(ndfi) / 18)
     If trdarab% > 0 Then
        For i1% = 1 To trdarab%
           i1d& = kezdoix& + i1% - 1
           Get #ndfi, (i1d& - 1) * 18& + 1, rcim&
           Seek #dbfi, rcim& + 9
           psztrec$ = Space(170): Get #dbfi, , psztrec$
           Mid$(psztrec$, 14, 1) = "S"
           kfrec$ = Mid$(psztrec$, 15, 120)
           Call dbxvir("KSZT", dbfi, psztrec$, rcim& + 9, 150)
           aktucim& = 0
                     
           '--- készlet sztornózása
           '
           ktetikt$ = Mid$(psztrec$, 158, 7)
           kfttrec$ = dbxkey("KKFT", ktetikt$)
           If kfttrec$ <> "" Then
              Mid(kfttrec$, 95, 1) = "S"
              Call dbxki("KKFT", kfttrec$, ";", "", "", hiba%)
              ymenny@ = -xval(Mid$(kfttrec$, 71, 12))
              Mid$(kfttrec$, 71, 12) = ertszam(Str(ymenny@), 12, 3)
              ymenny@ = -xval(Mid$(kfttrec$, 83, 12))
              Mid$(kfttrec$, 83, 12) = ertszam(Str(ymenny@), 12, 3)
              Call keszletvalt("B", kfttrec$, "K")
              If i1% = 1 Then
                 '--- forgalmi bizonylat sztornózása
                 bizikt$ = Mid(kfttrec$, 8, 7)
                 kbizrec$ = dbxkey("KKBZ", bizikt$)
                 Mid$(kbizrec$, 92, 1) = "S"
                 Mid$(kbizrec$, 93, 6) = maidatum$
                 Mid$(kbizrec$, 99, 8) = ugyintezo$
                 Call dbxki("KKBZ", kbizrec$, ";", "", "", hiba%)
              
              End If
           End If
        Next
     End If
     Close dbfi
     Close ndfi
  End If
   
End If
Next

End Sub
Private Sub Futtat(param$)
  On Error GoTo hibakez
    Kill auditorutvonal$ + terminal$ + task$ + "vege.par"
hibakez:
    Shell param$, vbNormalFocus
    fnev$ = terminal$ + task$ + "vege.par"
    Do
     DoEvents
     s$ = UCase(Dir(auditorutvonal$ + fnev$))
    Loop While s$ <> UCase(fnev$)
    sm1& = FileLen(auditorutvonal$ + fnev$)
    Do
      DoEvents
      Call waitsec(1)
      sm& = FileLen(auditorutvonal$ + fnev$)
      If sm1& = sm& Then Exit Do
      sm1& = sm&
    Loop While sm& > 0

  On Error GoTo vege
  
    Kill auditorutvonal$ + terminal$ + task$ + "vege.par"

vege:
End Sub
Private Sub GyariszamRogzit(ktikt$, szamlaszam$, sor%, kbikt$)

              gyariszamok$ = Nyugel1.GyariszamAtad(sor%, gysz$(), hiv$())
              pzx% = InStr(gyariszamok$, ":")
              db% = Val(Mid$(gyariszamok$, 1, pzx% - 1))
              gyariszamok$ = Mid$(gyariszamok$, pzx% + 1)
              For i2% = 1 To db%
               ' pzx% = InStr(gyariszamok$, ";")
               ' Gyariszamakt$ = Mid$(gyariszamok$, 1, pzx% - 1)
               Gyariszamakt$ = gysz$(i2%)
               Hivszamakt$ = hiv$(i2%)
                Call dbxtrkezd("KSZX")
              ' Gyári szám felvitel
                kszxrec$ = Space$(350)
                Mid$(kszxrec$, 320, 10) = ktikt$ + Right("000" + Trim(Str(i2%)), 3)    ' Iktató
                ' 2015.7.29
                Mid$(kszxrec$, 8, 4) = Nyugel1.MSFlexGrid1.TextMatrix(sor%, 8)  '  ttkrak$ ' A raktár kódja
                
                Mid$(kszxrec$, 12, 15) = Nyugel1.MSFlexGrid1.TextMatrix(sor%, 1)     ' Termék azonosító kódja (cikkszám)
                Mid$(kszxrec$, 27, 10) = szamlaszam$  ' Nyugta száma vagy számlaszám
                Mid$(kszxrec$, 37, 6) = Nyugel1.Text6.Text    ' Teljesítés kelte
                Mid$(kszxrec$, 43, 15) = Nyugel1.Text4(1).Text ' A vevõ (címzett) partnerkódja
                Mid$(kszxrec$, 200, 40) = Gyariszamakt$ 'Gyariszám
                Mid$(kszxrec$, 310, 7) = kbikt$
                Mid$(kszxrec$, 330, 10) = Hivszamakt$
                If i2% = 1 Then
                   Mid$(kszxrec$, 300, 10) = "         0"   'pointer
                Else
                   Mid$(kszxrec$, 300, 10) = Right(Space(10) + Str(aktucim&), 10)
                End If
                Call dbxki("KSZX", kszxrec$, ";", "U", "", hiba%)
                Call dbxtrvege
                
                w1% = obsorszama("KSZX")
                aktucim& = OBJTAB(w1%).obcim
              

                Call dbxtrkezd("KKFX")
                ' munkalapnál üres
                elozorec$ = dbxkey("KKFX", Hivszamakt$)
                If Not elozorec$ = "" Then
                 Mid$(elozorec$, 57, 10) = Mid$(kszxrec$, 310, 10)
                 Mid$(elozorec$, 67, 1) = "S"
                 Call dbxki("KKFX", elozorec$, ";", "", "", hiba%)
                End If
                Call dbxtrvege


                
              

              
              Next
              
End Sub

Private Sub Command1_Click()
  '--- kilépés program vége
  On Error GoTo hibakez
  If exparvolt% = 1 Then
    Kill "c:\auwset\auw" + terminal$ + task$ + ".auw"
  End If
  Close
  End
hibakez:
  Call mess(langmodul(156) + "XXX", 1, 0, langmodul(157), valasz%)
  Close: End
End Sub

Private Sub Command24_Click()
  Dim kkftikt$(100000)
  fil1 = FreeFile
  Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #fil1
  fil2 = FreeFile
  Open auditorutvonal$ + "auw-kkft.ndx" For Binary Shared As #fil2
  rc& = Int(LOF(fil2) / 12)
  kdb = 0
  For i& = 1 To rc&
            DoEvents
            ProgressBar3.Value = pscale(i&, rc&)
            Get #fil2, (i& - 1) * 12 + 1, cim&
            frec$ = Space(130)
            Get #fil1, cim& + 9, frec$
            Ikt$ = Mid$(frec$, 1, 7)
            mozg$ = Mid$(frec$, 21, 3)
            If mozg$ = "003" Or mozg$ = "009" Then
            kdb = kdb + 1
            kkftikt$(kdb) = Ikt$
            End If
   Next
   Close fil1, fil2
   dbfi = FreeFile
   Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #dbfi
   ndfi = FreeFile
   Open auditorutvonal$ + "auw-kszt.ndx" For Binary Shared As #ndfi
   rc& = Int(LOF(ndfi) / 18)
   For i1& = 1 To rc&
                  
         Get #ndfi, (i1& - 1) * 18& + 1, rcim&
         Seek #dbfi, rcim& + 9
         psztrec$ = Space(170): Get #dbfi, , psztrec$
         iktato$ = Mid$(psztrec$, 158, 7)
         For j1% = 1 To kdb
           If iktato$ = kkftikt$(j1%) Then
             kkftikt$(j1%) = 0
           End If
         Next
   Next
        
   Close dbfi, ndfi
   
   For j1% = 1 To kdb
       If Not kkftikt$(j1%) = 0 Then
             a$ = "0"
        End If
   Next

End Sub

Private Sub Command3_Click()
     '--- mehet a munka
On Error GoTo vege
maidatum = Text8.Text
'If maidatum < "100101" Then
'  Call mess("Állítsa be a mai dátumot!", 2, 0, langprg(1), valasz%)
'  Exit Sub
'End If
krakrec$ = ""
krak$ = Left(Trim(Text5.Text) + Space(4), 4)
If krak$ <> "" Then krakrec$ = dbxkey("KRAK", krak$)
If krakrec$ <> "" Then
  ttkrak$ = krak$
  Label6.Caption = " " + Mid$(krakrec$, 5, 40)
  Label6.Visible = True
  'Nyugel1.Caption = ttkrak + Label6.Caption
  Command1.Visible = False
  Command3.Visible = False
  Do
    ' 2015.03.17
     ttkrak$ = krak$
    
    megrendelesbol = False
    rpartner$ = ""
    partrec$ = ""
    
    Call nyelad
    munkalapbol = False
    munkalap$ = ""

    szallitobol = False
    szallito$ = ""

    szarmaztatasbol = False

    kmegrec$ = ""
    szoveg18$ = ""
    melyik% = 0
    List1.Clear
szamlaz:

    rnyugtavolt = nyugtavolt
    ' Eszi
    szlamod$ = " "
    ssz$ = " "
    szlacim$ = " "
    szlafaj$ = " "
    
    Select Case nyugtavolt
      Case 1

 
        '--- nyugta eladás
        '--- kiadási bizonylat rögzítése
       If nincs_tetelsor() Then
          Exit Do
       End If
       If Not Nyugel1.vaneloleg Then
        Call dbxtrkezd("INST")
        irec$ = dbxkey("INST", "INST")
        nyugtaszam$ = novel(irec$, 815, 6)
        Mid$(irec$, 815, 6) = nyugtaszam$
        Call dbxki("INST", irec$, ";", "", "", hiba%)
        Call dbxtrvege
        ' Eszi - 2009.12.02
        fejr$ = Space$(300)
    
        If Not Nyugel1.Text9 = "" Then
           If Mid$(Nyugel1.Text9, 5, 1) = "." Then
              Mid$(fejr$, 24, 6) = Mid$(Nyugel1.Text9, 3, 2) + Mid$(Nyugel1.Text9, 6, 2) + Mid$(Nyugel1.Text9, 9, 2)
           Else
             Mid$(fejr$, 24, 6) = Nyugel1.Text9
           End If
        Else
           Mid$(fejr$, 24, 6) = maidatum$  ' Teljesítés kelte
        End If
    
    
        Call dbxtrkezd("KSZB")
        pszbrec$ = Space$(300)
        Mid$(pszbrec$, 1, 10) = ttkrak$ + nyugtaszam$
        Mid$(pszbrec$, 21, 8) = ugyintezo$
        Mid$(pszbrec$, 29, 6) = maidatum$
        If Trim(Nyugel1.Text4(1).Text) = "" Then
          Mid$(pszbrec$, 61, 15) = Mid$(irec$, 800, 15)
        Else
          Mid$(pszbrec$, 61, 15) = Left(Trim(Nyugel1.Text4(1).Text) + Space(15), 15)
        End If
        Mid$(pszbrec$, 76, 2) = "BS"
        Mid$(pszbrec$, 78, 6) = maidatum$
    
        ' Eszi - 2009.12.02
        'Mid$(pszbrec$, 84, 6) = maidatum$
        Mid$(fejr$, 24, 6) = Mid$(fejr$, 24, 6)
    
        Mid$(pszbrec$, 90, 6) = maidatum$
        Mid$(pszbrec$, 231, 6) = maidatum$
        ' Eszi
        Mid$(pszbrec$, 250, 10) = terminal$ + " Nyugta"
        Mid$(pszbrec$, 111, 60) = Nyugel1.Text8.Text
        Mid$(pszbrec$, 171, 60) = Nyugel1.Text10.Text
    
        Call dbxki("KSZB", pszbrec$, ";", "U", "", hiba%)
        Call dbxtrvege
    
        fejr$ = Space$(300)
    
        If Not Nyugel1.Text9 = "" Then
           If Mid$(Nyugel1.Text9, 5, 1) = "." Then
              Mid$(fejr$, 24, 6) = Mid$(Nyugel1.Text9, 3, 2) + Mid$(Nyugel1.Text9, 6, 2) + Mid$(Nyugel1.Text9, 9, 2)
           Else
              Mid$(fejr$, 24, 6) = Nyugel1.Text9
           End If
        Else
           Mid$(fejr$, 24, 6) = maidatum$  ' Teljesítés kelte
        End If
        ' Eszi - 2009.12.04
        Mid$(pszbrec$, 84, 6) = Mid$(fejr$, 24, 6)

        Mid$(fejr$, 18, 6) = maidatum$  ' Számla kelte
        Mid$(fejr$, 30, 6) = maidatum$  ' Lejárat kelte
        Mid$(fejr$, 36, 2) = "01"       ' Fizetési mód
        Mid$(fejr$, 171, 8) = ""        ' Üzletkötõ
        Mid$(fejr$, 111, 60) = Nyugel1.Text8.Text
        Mid$(fejr$, 171, 58) = Nyugel1.Text11.Text

        Call dbxtrkezd("KKBZ")
        kbrec$ = Space$(140)
        Mid$(kbrec$, 8, 6) = maidatum$
         Mid$(kbrec$, 14, 40) = "Pénztárgépes eladás"
         Mid$(kbrec$, 54, 1) = "E"
         If Trim(Nyugel1.Text4(1).Text) = "" Then
           Mid$(kbrec$, 55, 15) = Mid$(irec$, 800, 15)
         Else
           Mid$(kbrec$, 55, 15) = Left(Trim(Nyugel1.Text4(1).Text) + Space(15), 15)
         End If
         Mid$(kbrec$, 107, 15) = ttkrak$ + nyugtaszam$
         szamlaszam$ = "Ny" + Mid$(ttkrak$, 1, 3) + nyugtaszam$
         Mid$(kbrec$, 78, 6) = maidatum$
         Mid$(kbrec$, 84, 8) = ugyintezo$
         Call dbxki("KKBZ", kbrec$, ";", "U", "G", hiba%)
         Call dbxtrvege
    
         tetdar% = 0
         For i1% = 1 To 200
           tkod$ = Left(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) + Space(15), 15)
           If Trim(tkod$) <> "" Then
             ' Hiba nullával osztás
             'a% = 5 / tetdar%
        
             tetdar% = tetdar% + 1
             termrec$ = dbxkey("KTRM", tkod$)
             ktrec$ = Space$(130)
             psztrec$ = Space$(170)
             Mid$(psztrec$, 130, 6) = Mid$(pszbrec$, 84, 6)
             Mid$(psztrec$, 1, 10) = ttkrak$ + nyugtaszam$
             Mid$(psztrec$, 122, 8) = terminal$ + " Ny"
             Mid$(psztrec$, 11, 3) = Right$("000" + Trim$(Str$(tetdar%)), 3)
             Mid$(psztrec$, 107, 15) = tkod$
             menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
             Mid$(psztrec$, 21, 12) = ertszam(Str(menny@), 12, 3)
             Mid$(psztrec$, 33, 6) = Mid$(termrec$, 484, 6)
             blar@ = xval(Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 4)))
             Mid$(psztrec$, 57, 12) = ertszam(Str(blar@), 12, 2)
             enge@ = -xval(Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 5)))
             blar@ = blar@ + blar@ * enge@ / 100
             blar@ = xval(Format(blar@, "###########0.00"))
             afakod$ = Mid$(termrec$, 706, 2)
             afrec$ = dbxkey("PAFA", afakod$)
             afakulcs@ = xval(Mid$(afrec$, 33, 6))
             elar@ = blar@ / ((100 + afakulcs@) / 100)
             Mid$(psztrec$, 51, 6) = ertszam(Str(enge@), 6, 2)
             Mid$(psztrec$, 39, 12) = ertszam(Str(elar@), 12, 2)
             Mid$(psztrec$, 69, 12) = ertszam(Str(elar@), 12, 2)
             Mid$(psztrec$, 81, 2) = Mid$(termrec$, 706, 2)
             Mid$(ktrec$, 8, 7) = Mid$(kbrec$, 1, 7)
             Mid$(ktrec$, 15, 6) = Mid$(kbrec$, 8, 6)
             ' Mozgás kód - visszaru
             menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
             Mid$(ktrec$, 71, 12) = ertszam(Str(-menny@), 12, 3)
             If menny@ >= 0 Then
               Mid$(ktrec$, 21, 3) = Mid$(irec$, 631, 3)
             Else
               Mid$(ktrec$, 21, 3) = Mid$(irec$, 655, 3)
             End If
             If Nyugel1.MSFlexGrid1.TextMatrix(i1%, 6) = "Elõleg besz." Then
                 Mid$(psztrec$, 150, 7) = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 7)
                 Mid$(psztrec$, 157, 1) = "B"
             End If
             
             ' ide kell betenni a raktárat a munkalapból készült számlába
             ' 2015.7.29
             Mid$(ktrec$, 24, 4) = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 8) ' ttkrak
             If munkalapbol Then
                kkftiktato$ = Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 6))
                
                If Len(kkftiktato$) = 7 Then
                  kkftml$ = dbxkey("KKFT", kkftiktato$)
                  If Not kkftml$ = "" Then
                     Mid$(ktrec$, 24, 4) = Mid$(kkftml$, 24, 4)
                  End If
                End If
             End If

             Mid$(ktrec$, 36, 15) = tkod$
             Mid$(ktrec$, 59, 12) = ertszam(Mid$(termrec$, 554, 14), 12, 2)
             Call dbxtrkezd("KKFT")
             Call dbxki("KKFT", ktrec$, ";", "U", "G", hiba%)
             Call dbxtrvege
        
             Mid(psztrec$, 158, 7) = Mid$(ktrec$, 1, 7)
             Call dbxtrkezd("KSZT")
             Call dbxki("KSZT", psztrec$, ";", "U", "", hiba%)
             Call dbxtrvege

             Call GyariszamRogzit(Mid$(ktrec$, 1, 7), nyugtaszam$, i1%, Mid$(kbrec$, 1, 7))
        
             If i1% = 1 Then
               Call dbxtrkezd("KKBZ")
               Mid$(kbrec$, 122, 7) = Mid$(ktrec$, 1, 7)
               Call dbxki("KKBZ", kbrec$, ";", "", "", hiba%)
               Call dbxtrvege
             End If
             ' 2011.10.14
             Call dbxtrkezd("KSZT")
             w1% = obsorszama("KSZT")
             aktucim& = OBJTAB(w1%).obcim
             Call lancra("AUWKER", "KSZTERM", termrec$, aktucim&, psztrec$)
             w1% = obsorszama("KKFT")
             aktucim& = OBJTAB(w1%).obcim
             Call lancra("AUWKER", "KKFKFRG", termrec$, aktucim&, ktrec$)
             Call keszletvalt("K", ktrec$, "N")
             Call dbxtrvege
             ' Eszi - megrendelõ kezelés
             If kmegrec$ <> "" Then
               kimenny@ = Abs(xval(Mid$(ktrec$, 71, 12)))
               tkod$ = Mid$(ktrec$, 36, 15)
               rakod$ = Mid(kmegrec$, 480, 4)
               For i111% = 1 To 1000
                 If i111% <= 200 Then
                   ele2$ = Mid$(kmegrec$, (i111% - 1) * 59 + 500, 59)
                   If Mid$(ele2$, 1, 15) = Mid$(ktrec$, 36, 15) Then
                     Call hozzad(ele2$, 48, 12, kimenny@, 3)
                     If i111% <= 200 Then Mid$(kmegrec$, (i111% - 1) * 59 + 500, 59) = ele2$
                     Call dbxtrkezd("KMEG")
                     Call dbxki("KMEG", kmegrec$, ";", "", "", hiba%)
                     Call dbxtrvege
                     If Mid$(kmegrec$, 192, 1) = "D" Then
                         minta$ = Mid$(ele2$, 16, 8)
                        Call foglal("N", rakod$, tkod$, minta$, kimenny@)
                     End If
                     Exit For
                   End If
                 End If
               Next
             End If
           End If
         Next
         ' Eszi - megrendelõ kezelés - lezárás
         If kmegrec$ <> "" Then
            Mid$(kmegrec$, 406, 1) = "S"
            Mid$(kmegrec$, 407, 8) = nyugtaszam$
            Mid$(kmegrec$, 484, 6) = maidatum$
            Call dbxtrkezd("KMEG")
            Call dbxki("KMEG", kmegrec$, ";", "", "", hiba%)
            Call dbxtrvege
         End If
         Call dbxtrkezd("KSZB")
         Mid$(pszbrec$, 50, 3) = Right$("   " + Str$(tetdar%), 3)
         Call dbxki("KSZB", pszbrec$, ";", "", "", hiba%)
         Call dbxtrvege
         If szallitobol Then
           ' Eredeti szállító sztonója
           Call szallsztor(szallito$, nyugtaszam$)
         End If
         
         ' Partner név,cím letárolása
         Call partnevcimtarol(kbrec$, pszbrec$, nyugtavolt)
         'Call dbxtrkezd("KSYB")
         'psybrec$ = Space$(200)
         'Mid$(psybrec$, 1, 7) = Mid$(kbrec$, 1, 7)
         'Mid$(psybrec$, 8, 15) = Mid$(pszbrec$, 61, 15)
         'If RTrim$(Mid$(pszbrec$, 61, 15)) = "" Or Mid$(pszbrec$, 61, 4) = "BANK" Then
         '   Mid$(psybrec$, 8, 15) = Mid$(pszbrec$, 61, 15)
         '   Mid$(psybrec$, 16, 7) = Nyugel1.Text2
         'Else
         '   Mid$(psybrec$, 8, 15) = Mid$(pszbrec$, 61, 15)
         'End If
         'Mid$(psybrec$, 23, 60) = Nyugel1.Text2          ' Partner neve
         'Mid$(psybrec$, 83, 60) = Nyugel1.Text3          ' Partner címe
         'Mid$(psybrec$, 143, 58) = Nyugel1.Text11        ' Megjegyzés 2. sor

         'Call dbxki("KSYB", psybrec$, ";", "U", "", hiba%)
    
         'Call dbxtrvege
    ' Eszi
    ' Bizonylat nyontatás errõl is
'          fi3 = FreeFile
'          Open listautvonal$ + terminal$ + task$ + "szle.lst" For Output As #fi3
         kpsszamla = 1
         GoSub szamlair
'          Close fi3
'          Shell programutvonal$ + "dbx4-sho.exe " + terminal$ + task$ + "szle/" + listautvonal$, vbNormalFocus
         Else
           Call mess("Elõleget számlába számítson be!", 2, 0, langprg(1), valasz%)
         End If
    
       Case 2
         '--- ÁFÁS készpénzes számla
         If nincs_tetelsor() Then
            Exit Do
         End If
    
         Call dbxtrkezd("INST")
         irec$ = dbxkey("INST", "INST")
         nyugtaszam$ = novel(irec$, 815, 6)
         Mid$(irec$, 815, 6) = nyugtaszam$
         Call dbxki("INST", irec$, ";", "", "", hiba%)
         Call dbxtrvege
    
         'Call dbxtrkezd("SINS")
         'sirec$ = dbxkey("SINS", "INST")
         'sp% = 500
         'szamlaszam$ = Mid$(sirec$, sp%, 4) + novel(sirec$, sp% + 4, 6)
         'Mid$(sirec$, sp%, 10) = szamlaszam$
         'Call dbxki("SINS", sirec$, ";", "", "", hiba%)
         'Call dbxtrvege
    
         fejr$ = Space$(300)
    
         Mid$(fejr$, 16, 2) = "BS"
         If Trim$(Nyugel1.Text4(1).Text) = "" Then
            Mid$(fejr$, 1, 15) = KPPARTN
         Else
            Mid$(fejr$, 1, 15) = Nyugel1.Text4(1).Text    ' Partner kód
         End If
         If Not Nyugel1.Text9 = "" Then
            If Mid$(Nyugel1.Text9, 5, 1) = "." Then
               Mid$(fejr$, 24, 6) = Mid$(Nyugel1.Text9, 3, 2) + Mid$(Nyugel1.Text9, 6, 2) + Mid$(Nyugel1.Text9, 9, 2)
            Else
               Mid$(fejr$, 24, 6) = Nyugel1.Text9
            End If
         Else
            Mid$(fejr$, 24, 6) = maidatum$  ' Teljesítés kelte
         End If
         Mid$(fejr$, 18, 6) = maidatum$  ' Számla kelte
         Mid$(fejr$, 30, 6) = maidatum$  ' Lejárat kelte
         Mid$(fejr$, 36, 2) = "01"       ' Fizetési mód
         Mid$(fejr$, 171, 8) = ""        ' Üzletkötõ
         Mid$(fejr$, 111, 60) = Nyugel1.Text8.Text
         Mid$(fejr$, 171, 58) = Nyugel1.Text11.Text

         szlamod$ = "U"
         szlafaj$ = "K"
         szlacim$ = "KÉSZPÉNZFIZETÉSI   S Z Á M L A"
         kpsszamla = 1
         GoSub szamlair

         
         
         Call dbxtrkezd("KSZB")
         pszbrec$ = Space$(300)
         Mid$(pszbrec$, 1, 10) = szamlaszam$ ' ttkrak$ + nyugtaszam$
         Mid$(pszbrec$, 21, 8) = ugyintezo$
         Mid$(pszbrec$, 29, 6) = maidatum$
         If Trim(Nyugel1.Text4(1).Text) = "" Then
           Mid$(pszbrec$, 61, 15) = Mid$(irec$, 800, 15) ' alap paraméter kp-és partner
         Else
           Mid$(pszbrec$, 61, 15) = Left(Trim(Nyugel1.Text4(1).Text) + Space(15), 15)
         End If
         Mid$(pszbrec$, 76, 2) = "BS"
         Mid$(pszbrec$, 78, 6) = maidatum$
    ' Eszi - 2009.12.02
    'Mid$(pszbrec$, 84, 6) = maidatum$
         Mid$(pszbrec$, 84, 6) = Mid$(fejr$, 24, 6)
    
         Mid$(pszbrec$, 90, 6) = maidatum$
         Mid$(pszbrec$, 231, 6) = maidatum$
         Mid$(pszbrec$, 96, 2) = Nyugel1.Text7
    
    
         Mid$(pszbrec$, 250, 10) = terminal$ + " Kp szl"
         Mid$(pszbrec$, 111, 60) = Nyugel1.Text8.Text
         Mid$(pszbrec$, 171, 60) = Nyugel1.Text10.Text
         
    
         Call dbxki("KSZB", pszbrec$, ";", "U", "", hiba%)
         Call dbxtrvege
         Call dbxtrkezd("KKBZ")
         kbrec$ = Space$(140)
         Mid$(kbrec$, 8, 6) = maidatum$
         Mid$(kbrec$, 14, 40) = "Pénztárgépes eladás"
         Mid$(kbrec$, 54, 1) = "E"
         If Trim(Nyugel1.Text4(1).Text) = "" Then
           Mid$(kbrec$, 55, 15) = Mid$(irec$, 800, 15)
         Else
           Mid$(kbrec$, 55, 15) = Left(Trim(Nyugel1.Text4(1).Text) + Space(15), 15)
         End If
         Mid$(kbrec$, 107, 15) = szamlaszam$
         Mid$(kbrec$, 78, 6) = maidatum$
         Mid$(kbrec$, 84, 8) = ugyintezo$
         Call dbxki("KKBZ", kbrec$, ";", "U", "G", hiba%)
         Call dbxtrvege
    
         tetdar% = 0
         For i1% = 1 To 200
           tkod$ = Left(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) + Space(15), 15)
           If Trim(tkod$) <> "" Then
             tetdar% = tetdar% + 1
             termrec$ = dbxkey("KTRM", tkod$)
             ktrec$ = Space$(130)
             psztrec$ = Space$(170)
             Mid$(psztrec$, 130, 6) = Mid$(pszbrec$, 84, 6)
             Mid$(psztrec$, 1, 10) = szamlaszam$
             Mid$(psztrec$, 122, 8) = terminal$ + " Kp"
             Mid$(psztrec$, 11, 3) = Right$("000" + Trim$(Str$(tetdar%)), 3)
             Mid$(psztrec$, 107, 15) = tkod$
             menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
             Mid$(psztrec$, 21, 12) = ertszam(Str(menny@), 12, 3)
             Mid$(psztrec$, 33, 6) = Mid$(termrec$, 484, 6)
             blar@ = xval(Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 4)))
             Mid$(psztrec$, 57, 12) = ertszam(Str(blar@), 12, 2)
             enge@ = -xval(Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 5)))
             blar@ = blar@ + blar@ * enge@ / 100
             blar@ = xval(Format(blar@, "###########0.00"))
             afakod$ = Mid$(termrec$, 706, 2)
             afrec$ = dbxkey("PAFA", afakod$)
             afakulcs@ = xval(Mid$(afrec$, 33, 6))
             elar@ = blar@ / ((100 + afakulcs@) / 100)
             Mid$(psztrec$, 51, 6) = ertszam(Str(enge@), 6, 2)
             Mid$(psztrec$, 39, 12) = ertszam(Str(elar@), 12, 2)
             Mid$(psztrec$, 69, 12) = ertszam(Str(elar@), 12, 2)
             Mid$(psztrec$, 81, 2) = Mid$(termrec$, 706, 2)
             Mid$(ktrec$, 8, 7) = Mid$(kbrec$, 1, 7)
             Mid$(ktrec$, 15, 6) = Mid$(kbrec$, 8, 6)
             Mid$(ktrec$, 21, 3) = Mid$(irec$, 631, 3)
             ' Mozgás kód - visszaáru
             menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
             Mid$(ktrec$, 71, 12) = ertszam(Str(-menny@), 12, 3)
             If menny@ >= 0 Then
               Mid$(ktrec$, 21, 3) = Mid$(irec$, 631, 3)
             Else
               Mid$(ktrec$, 21, 3) = Mid$(irec$, 655, 3)
             End If
             If Nyugel1.MSFlexGrid1.TextMatrix(i1%, 6) = "Elõleg besz." Then
                 Mid$(psztrec$, 150, 7) = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 7)
                 Mid$(psztrec$, 157, 1) = "B"
             End If

             '  ide kell betenni a raktárat a munkalapból készült számlába
             ' 2015.7.29
             Mid$(ktrec$, 24, 4) = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 8) ' ttkrak
             If munkalapbol Then
                kkftiktato$ = Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 6))
                
                If Len(kkftiktato$) = 7 Then
                  kkftml$ = dbxkey("KKFT", kkftiktato$)
                  If Not kkftml$ = "" Then
                     Mid$(ktrec$, 24, 4) = Mid$(kkftml$, 24, 4)
                  End If
                End If
             End If
             Mid$(ktrec$, 36, 15) = tkod$
             Mid$(ktrec$, 59, 12) = ertszam(Mid$(termrec$, 554, 14), 12, 2)
             Call dbxtrkezd("KKFT")
             Call dbxki("KKFT", ktrec$, ";", "U", "G", hiba%)
             Call dbxtrvege
        
             Call dbxtrkezd("KSZT")
             Mid(psztrec$, 158, 7) = Mid$(ktrec$, 1, 7)
             Call dbxki("KSZT", psztrec$, ";", "U", "", hiba%)
             Call dbxtrvege
        
             Call GyariszamRogzit(Mid$(ktrec$, 1, 7), szamlaszam$, i1%, Mid$(kbrec$, 1, 7))
        
             If i1% = 1 Then
               Call dbxtrkezd("KKBZ")
               Mid$(kbrec$, 122, 7) = Mid$(ktrec$, 1, 7)
               Call dbxki("KKBZ", kbrec$, ";", "", "", hiba%)
               Call dbxtrvege
             End If
             ' 2011.10.14
             Call dbxtrkezd("KSZT")
             w1% = obsorszama("KSZT")
             aktucim& = OBJTAB(w1%).obcim
             Call lancra("AUWKER", "KSZTERM", termrec$, aktucim&, psztrec$)
             w1% = obsorszama("KKFT")
             aktucim& = OBJTAB(w1%).obcim
             Call lancra("AUWKER", "KKFKFRG", termrec$, aktucim&, ktrec$)
             Call keszletvalt("K", ktrec$, "N")
             Call dbxtrvege
             ' Eszi - megrendelõ kezelés
             If kmegrec$ <> "" Then
               kimenny@ = Abs(xval(Mid$(ktrec$, 71, 12)))
               tkod$ = Mid$(ktrec$, 36, 15)
               rakod$ = Mid(kmegrec$, 480, 4)

               For i111% = 1 To 1000
                 If i111% <= 200 Then
                   ele2$ = Mid$(kmegrec$, (i111% - 1) * 59 + 500, 59)
                   If Mid$(ele2$, 1, 15) = Mid$(ktrec$, 36, 15) Then
                     Call hozzad(ele2$, 48, 12, kimenny@, 3)
                     If i111% <= 200 Then Mid$(kmegrec$, (i111% - 1) * 59 + 500, 59) = ele2$
                     Call dbxtrkezd("KMEG")
                     Call dbxki("KMEG", kmegrec$, ";", "", "", hiba%)
                     Call dbxtrvege
                     If Mid$(kmegrec$, 192, 1) = "D" Then
                        minta$ = Mid$(ele2$, 16, 8)
                        Call foglal("N", rakod$, tkod$, minta$, kimenny@)
                     End If
                
                     Exit For
                   End If
                 End If
               Next
             End If

           End If
         Next
         ' Eszi - megrendelõ kezelés - lezárás
         If kmegrec$ <> "" Then
            Mid$(kmegrec$, 406, 1) = "S"
            Mid$(kmegrec$, 407, 8) = Mid$(szamlaszam$, 3, 8)
            Mid$(kmegrec$, 484, 6) = maidatum$
            Call dbxtrkezd("KMEG")
            Call dbxki("KMEG", kmegrec$, ";", "", "", hiba%)
            Call dbxtrvege
         End If

         Call dbxtrkezd("KSZB")
         Mid$(pszbrec$, 50, 3) = Right$("   " + Str$(tetdar%), 3)
         Call dbxki("KSZB", pszbrec$, ";", "", "", hiba%)
         Call dbxtrvege
         If szallitobol Then
           ' Eredeti szállító sztonója
           Call szallsztor(szallito$, szamlaszam$)
         End If
    
         ' Partner név,cím letárolása
         Call partnevcimtarol(kbrec$, pszbrec$, nyugtavolt)
         
         'Call dbxtrkezd("KSYB")
         'psybrec$ = Space$(200)
         'Mid$(psybrec$, 1, 7) = Mid$(kbrec$, 1, 7)
         'Mid$(psybrec$, 8, 15) = Mid$(pszbrec$, 61, 15)
         'Mid$(psybrec$, 23, 60) = Nyugel1.Text2          ' Partner neve
         'Mid$(psybrec$, 83, 60) = Nyugel1.Text3          ' Partner címe
         'Mid$(psybrec$, 143, 58) = Nyugel1.Text11        ' Megjegyzés 2. sor
         

         'Call KPAR_Tolt
   
         'For i99% = 1 To 4
         '   elem$ = Nyugel1.NtAtad(i99%)
         '   If Trim$(elem$) <> "" Then
         '      Mid$(psybrec$, (i99% - 1) + 143, 36) = Mid$(elem$, 1, 36)
         '   End If
         'Next
    
         'Call dbxki("KSYB", psybrec$, ";", "U", "", hiba%)
    
         'Call dbxtrvege
    
         ' KP-és számla elõleg beszámítás
         'For i99% = 1 To 5
         '   elem$ = Nyugel1.NtAtad(i99%)
         '   If Trim$(elem$) <> "" Then
         '     konyveldat$ = Mid$(fejr$, 24, 6)
         '     Call elobekonyvel(elem$, "V", konyveldat$)
         '   End If
         'Next
         
         Call Eloleg_beszamitas(" ")

    
'          fi3 = FreeFile
'          Open listautvonal$ + terminal$ + task$ + "szle.lst" For Output As #fi3
''         szlamod$ = "U"
''         kpsszamla = 1
''         GoSub szamlair
'          Close fi3
'          Shell programutvonal$ + "dbx4-sho.exe " + terminal$ + task$ + "szle/" + listautvonal$, vbNormalFocus
       Case 3, 12
        '--- ÁFÁS átutalásos számla
         '--- ÁFÁS számla készítése
         '--- PVSZ-be is rögzítei kell
         If nincs_tetelsor() Then
            Exit Do
         End If
    
         'Call dbxtrkezd("INST")
    'irec$ = dbxkey("INST", "INST")
    'sp% = 287
    'szamlaszam$ = Mid$(irec$, sp%, 4) + novel(irec$, sp% + 4, 6)
    
    'Mid$(irec$, sp%, 10) = szamlaszam$
    'Call dbxki("INST", irec$, ";", "", "", hiba%)
    'Call dbxtrvege
    
    fejr$ = Space$(300)
    If Trim$(Nyugel1.Text4(1).Text) = "" Then
       Mid$(fejr$, 1, 15) = BANKPARTN
    Else
       Mid$(fejr$, 1, 15) = Nyugel1.Text4(1).Text    ' Partner kód
    End If

    Mid$(fejr$, 16, 2) = "BS"
    
    If Not Nyugel1.Text9 = "" Then
       If Mid$(Nyugel1.Text9, 5, 1) = "." Then
          Mid$(fejr$, 24, 6) = Mid$(Nyugel1.Text9, 3, 2) + Mid$(Nyugel1.Text9, 6, 2) + Mid$(Nyugel1.Text9, 9, 2)
       Else
          Mid$(fejr$, 24, 6) = Nyugel1.Text9
       End If
    Else
       Mid$(fejr$, 24, 6) = maidatum$  ' Teljesítés kelte
    End If
    Mid$(fejr$, 18, 6) = maidatum$  ' Számla kelte
    Mid$(fejr$, 30, 6) = Nyugel1.Text6.Text  ' Lejárat kelte
    Mid$(fejr$, 36, 2) = Nyugel1.Text7.Text  ' Fizetési mód
    Mid$(fejr$, 171, 8) = ""        ' Üzletkötõ
    Mid$(fejr$, 96, 2) = Nyugel1.Text7 ' Fizetési mód
    Mid$(fejr$, 111, 60) = Nyugel1.Text8.Text
    Mid$(fejr$, 171, 58) = Nyugel1.Text11.Text

    szlacim$ = "S Z Á M L A"
    szlamod$ = "U"
    szlafaj$ = "U"
    GoSub szamlair
  
    
    Call dbxtrkezd("KSZB")
    pszbrec$ = Space$(300)
    Mid$(pszbrec$, 1, 10) = szamlaszam$
    Mid$(pszbrec$, 21, 8) = ugyintezo$
    Mid$(pszbrec$, 29, 6) = maidatum$
    If Trim(Nyugel1.Text4(1).Text) = "" Then
      Mid$(pszbrec$, 61, 15) = Mid$(irec$, 800, 15)
    Else
      Mid$(pszbrec$, 61, 15) = Left(Trim(Nyugel1.Text4(1).Text) + Space(15), 15)
    End If
    Mid$(pszbrec$, 76, 2) = "BS"
    Mid$(pszbrec$, 78, 6) = maidatum$
    
    'Mid$(pszbrec$, 84, 6) = maidatum$
    Mid$(pszbrec$, 84, 6) = Mid$(fejr$, 24, 6)
    
    Mid$(pszbrec$, 90, 6) = Mid$(fejr$, 30, 6)
    Mid$(pszbrec$, 231, 6) = maidatum$
    Mid$(pszbrec$, 96, 2) = Nyugel1.Text7
    
    Mid$(pszbrec$, 250, 10) = terminal$ + " Ht szl"
    Mid$(pszbrec$, 111, 60) = Nyugel1.Text8.Text
    Mid$(pszbrec$, 171, 60) = Nyugel1.Text10.Text
    Mid$(fejr$, 171, 58) = Nyugel1.Text11.Text
    
    Call dbxki("KSZB", pszbrec$, ";", "U", "", hiba%)
    Call dbxtrvege
    
    Call dbxtrkezd("KKBZ")
    kbrec$ = Space$(140)
    Mid$(kbrec$, 8, 6) = maidatum$
    Mid$(kbrec$, 14, 40) = "Hiteles eladás"
    Mid$(kbrec$, 54, 1) = "E"
    If Trim(Nyugel1.Text4(1).Text) = "" Then
      Mid$(kbrec$, 55, 15) = Mid$(irec$, 800, 15)
    Else
      Mid$(kbrec$, 55, 15) = Left(Trim(Nyugel1.Text4(1).Text) + Space(15), 15)
    End If
    Mid$(kbrec$, 107, 15) = szamlaszam$
    Mid$(kbrec$, 78, 6) = maidatum$
    Mid$(kbrec$, 84, 8) = ugyintezo$
    Call dbxki("KKBZ", kbrec$, ";", "U", "G", hiba%)
    Call dbxtrvege
    
    tetdar% = 0
    For i1% = 1 To 200
      tkod$ = Left(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) + Space(15), 15)
      If Trim(tkod$) <> "" Then
        tetdar% = tetdar% + 1
        termrec$ = dbxkey("KTRM", tkod$)
        ktrec$ = Space$(130)
        psztrec$ = Space$(170)
        Mid$(psztrec$, 130, 6) = Mid$(pszbrec$, 84, 6)
        Mid$(psztrec$, 1, 10) = szamlaszam$
        Mid$(psztrec$, 122, 8) = terminal$ + " Ht"
        Mid$(psztrec$, 11, 3) = Right$("000" + Trim$(Str$(tetdar%)), 3)
        Mid$(psztrec$, 107, 15) = tkod$
        menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
        Mid$(psztrec$, 21, 12) = ertszam(Str(menny@), 12, 3)
        Mid$(psztrec$, 33, 6) = Mid$(termrec$, 484, 6)
        blar@ = xval(Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 4)))
        Mid$(psztrec$, 57, 12) = ertszam(Str(blar@), 12, 2)
        enge@ = -xval(Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 5)))
        blar@ = blar@ + blar@ * enge@ / 100
        blar@ = xval(Format(blar@, "###########0.00"))
        afakod$ = Mid$(termrec$, 706, 2)
        afrec$ = dbxkey("PAFA", afakod$)
        afakulcs@ = xval(Mid$(afrec$, 33, 6))
        elar@ = blar@ / ((100 + afakulcs@) / 100)
        Mid$(psztrec$, 51, 6) = ertszam(Str(enge@), 6, 2)
        Mid$(psztrec$, 39, 12) = ertszam(Str(elar@), 12, 2)
        Mid$(psztrec$, 69, 12) = ertszam(Str(elar@), 12, 2)
        Mid$(psztrec$, 81, 2) = Mid$(termrec$, 706, 2)
        Mid$(ktrec$, 8, 7) = Mid$(kbrec$, 1, 7)
        Mid$(ktrec$, 15, 6) = Mid$(kbrec$, 8, 6)

        ' Mozgás kód - visszaáru
        menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
        Mid$(ktrec$, 71, 12) = ertszam(Str(-menny@), 12, 3)
        If menny@ >= 0 Then
          Mid$(ktrec$, 21, 3) = Mid$(irec$, 631, 3)
        Else
          Mid$(ktrec$, 21, 3) = Mid$(irec$, 655, 3)
        End If

        ' ide kell betenni a raktárat a munkalapból készült számlába
        ' 2015.7.29
        Mid$(ktrec$, 24, 4) = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 8) ' ttkrak
        If munkalapbol Then
           kkftiktato$ = Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 6))
           If Len(kkftiktato$) = 7 Then
             kkftml$ = dbxkey("KKFT", kkftiktato$)
             If Not kkftml$ = "" Then
                Mid$(ktrec$, 24, 4) = Mid$(kkftml$, 24, 4)
             End If
           End If
        End If
        Mid$(ktrec$, 36, 15) = tkod$
        Mid$(ktrec$, 59, 12) = ertszam(Mid$(termrec$, 554, 14), 12, 2)
        
        Call dbxtrkezd("KKFT")
        Call dbxki("KKFT", ktrec$, ";", "U", "G", hiba%)
        Call dbxtrvege
        
        Call dbxtrkezd("KSZT")
        Mid(psztrec$, 158, 7) = Mid$(ktrec$, 1, 7)
        Call dbxki("KSZT", psztrec$, ";", "U", "", hiba%)
        Call dbxtrvege
        
        Call GyariszamRogzit(Mid$(ktrec$, 1, 7), szamlaszam$, i1%, Mid$(kbrec$, 1, 7))
        
        If i1% = 1 Then
          Mid$(kbrec$, 122, 7) = Mid$(ktrec$, 1, 7)
          Call dbxtrkezd("KKBZ")
          Call dbxki("KKBZ", kbrec$, ";", "", "", hiba%)
          Call dbxtrvege
        End If
        ' 2011.10.14
        Call dbxtrkezd("KSZT")
        w1% = obsorszama("KSZT")
        aktucim& = OBJTAB(w1%).obcim
        Call lancra("AUWKER", "KSZTERM", termrec$, aktucim&, psztrec$)
        w1% = obsorszama("KKFT")
        aktucim& = OBJTAB(w1%).obcim
        Call lancra("AUWKER", "KKFKFRG", termrec$, aktucim&, ktrec$)
        Call keszletvalt("K", ktrec$, "N")
        Call dbxtrvege
        ' Eszi - megrendelõ kezelés
        If kmegrec$ <> "" Then
          kimenny@ = Abs(xval(Mid$(ktrec$, 71, 12)))
          tkod$ = Mid$(ktrec$, 36, 15)
          rakod$ = Mid(kmegrec$, 480, 4)
          
          For i111% = 1 To 1000
            If i111% <= 200 Then
              ele2$ = Mid$(kmegrec$, (i111% - 1) * 59 + 500, 59)
              If Mid$(ele2$, 1, 15) = Mid$(ktrec$, 36, 15) Then
                Call hozzad(ele2$, 48, 12, kimenny@, 3)
                If i111% <= 200 Then Mid$(kmegrec$, (i111% - 1) * 59 + 500, 59) = ele2$
                Call dbxtrkezd("KMEG")
                Call dbxki("KMEG", kmegrec$, ";", "", "", hiba%)
                Call dbxtrvege
                If Mid$(kmegrec$, 192, 1) = "D" Then
                   minta$ = Mid$(ele2$, 16, 8)
                   Call foglal("N", rakod$, tkod$, minta$, kimenny@)
                End If
                
                Exit For
              End If
            End If
          Next
        End If

      End If
    Next
    ' Eszi - megrendelõ kezelés - lezárás
    If kmegrec$ <> "" Then
       Mid$(kmegrec$, 406, 1) = "S"
       Mid$(kmegrec$, 407, 8) = Mid$(szamlaszam$, 3, 8)
       Mid$(kmegrec$, 484, 6) = maidatum$
       Call dbxtrkezd("KMEG")
       Call dbxki("KMEG", kmegrec$, ";", "", "", hiba%)
       Call dbxtrvege
    End If

    Mid$(pszbrec$, 50, 3) = Right$("   " + Str$(tetdar%), 3)
    Call dbxtrkezd("KSZB")
    Call dbxki("KSZB", pszbrec$, ";", "", "", hiba%)
    Call dbxtrvege
    
    ' Partner név,cím letárolása
    Call partnevcimtarol(kbrec$, pszbrec$, nyugtavolt)
    
    'Call dbxtrkezd("KSYB")
    'psybrec$ = Space$(200)
    'Mid$(psybrec$, 1, 7) = Mid$(kbrec$, 1, 7)
    'Mid$(psybrec$, 8, 15) = Mid$(pszbrec$, 61, 15)
    'Mid$(psybrec$, 23, 60) = Nyugel1.Text2          ' Partner neve
    'Mid$(psybrec$, 83, 60) = Nyugel1.Text3          ' Partner címe
    'Mid$(psybrec$, 143, 58) = Nyugel1.Text11        ' Megjegyzés 2. sor
    
    
    'Call KPAR_Tolt
    
    'For i99% = 1 To 4
    '   elem$ = Nyugel1.NtAtad(i99%)
    '   If Trim$(elem$) <> "" Then
    '   Mid$(psybrec$, (i99% - 1) + 143, 36) = Mid$(elem$, 1, 36)
    '   End If
    'Next


    'Call dbxki("KSYB", psybrec$, ";", "U", "", hiba%)
    
    'Call dbxtrvege
    ' 2015.07.02 -
    
    Call Eloleg_beszamitas(" ")
    
    kpsszamla = 0
    Call dbxtrkezd("PVSZ")
    GoSub folyokonyvel
    Call dbxtrvege
    
    Call dbxtrkezd("KSZB")
    Mid$(pszbrec$, 53, 7) = vikt$
    Call dbxki("KSZB", pszbrec$, ";", "", "", hiba%)
    Call dbxtrvege
    
    If szallitobol Then
      ' Eredeti szállító sztonója
      Call szallsztor(szallito$, szamlaszam$)
    End If


'          fi3 = FreeFile
    
'          Open listautvonal$ + terminal$ + task$ + "szla.lst" For Output As #fi3
'    kpsszamla = 0
''    szlamod$ = "U"
''    GoSub szamlair
'          Close fi3
'          Shell programutvonal$ + "dbx4-sho.exe " + terminal$ + task$ + "szla/" + listautvonal$, vbNormalFocus
  
  Case 7
    '--- gyarancia jegy - eladás szállító levélre
    '--- kiadási bizonylat rögzítése

   If nincs_tetelsor() Then
      Exit Do
   End If
   If Not Nyugel1.vaneloleg Then
    If Nyugel1.Check1.Value = 1 Then
     Call dbxtrkezd("INST")
     irec$ = dbxkey("INST", "INST")
     nyugtaszam$ = novel(irec$, 638, 10)
     Mid$(irec$, 638, 10) = nyugtaszam$
     nyugtaszam$ = nyugtaszam$
     Call dbxki("INST", irec$, ";", "", "", hiba%)
     Call dbxtrvege
    
    Else
     Call dbxtrkezd("SINS")
     sirec$ = dbxkey("SINS", "INST")
     nyugtaszam$ = novel(sirec$, 534, 6)
     Mid$(sirec$, 534, 6) = nyugtaszam$
     nyugtaszam$ = Mid$(sirec$, 530, 4) + nyugtaszam$
     Call dbxki("SINS", sirec$, ";", "", "", hiba%)
     Call dbxtrvege
    End If
    Call dbxtrkezd("KSZB")
    pszbrec$ = Space$(300)
    Mid$(pszbrec$, 1, 10) = nyugtaszam$
    Mid$(pszbrec$, 21, 8) = ugyintezo$
    Mid$(pszbrec$, 29, 6) = maidatum$
    If Trim(Nyugel1.Text4(1).Text) = "" Then
      Mid$(pszbrec$, 61, 15) = Mid$(irec$, 800, 15)
    Else
      Mid$(pszbrec$, 61, 15) = Left(Trim(Nyugel1.Text4(1).Text) + Space(15), 15)
    End If
    Mid$(pszbrec$, 76, 2) = "SL"
    Mid$(pszbrec$, 78, 6) = maidatum$
    Mid$(pszbrec$, 84, 6) = maidatum$
    Mid$(pszbrec$, 90, 6) = maidatum$
    Mid$(pszbrec$, 231, 6) = maidatum$
    ' Eszi
    If Not Nyugel1.Text9 = "" Then
       If Mid$(Nyugel1.Text9, 5, 1) = "." Then
          Mid$(pszbrec$, 84, 6) = Mid$(Nyugel1.Text9, 3, 2) + Mid$(Nyugel1.Text9, 6, 2) + Mid$(Nyugel1.Text9, 9, 2)
       Else
          Mid$(pszbrec$, 84, 6) = Nyugel1.Text9
       End If
    Else
       Mid$(pszbrec$, 84, 6) = maidatum$ ' Teljesítés kelte
    End If
    
    If Nyugel1.Check1.Value = 1 Then
      Mid$(pszbrec$, 250, 10) = terminal$ + " Száll."
    Else
      Mid$(pszbrec$, 250, 10) = terminal$ + " Gar."
    End If
    Mid$(pszbrec$, 111, 60) = Nyugel1.Text8.Text
    Mid$(pszbrec$, 171, 60) = Nyugel1.Text10.Text
        
    Call dbxki("KSZB", pszbrec$, ";", "U", "", hiba%)
    Call dbxtrvege
    
    szamlaszam$ = nyugtaszam$
    
    Call dbxtrkezd("KKBZ")
    kbrec$ = Space$(140)
    Mid$(kbrec$, 8, 6) = maidatum$
    Mid$(kbrec$, 14, 40) = "Szállító levél"
    Mid$(kbrec$, 54, 1) = "E"
    If Trim(Nyugel1.Text4(1).Text) = "" Then
      Mid$(kbrec$, 55, 15) = Mid$(irec$, 800, 15)
    Else
      Mid$(kbrec$, 55, 15) = Left(Trim(Nyugel1.Text4(1).Text) + Space(15), 15)
    End If
    Mid$(kbrec$, 107, 15) = nyugtaszam$
    Mid$(kbrec$, 78, 6) = maidatum$
    Mid$(kbrec$, 84, 8) = ugyintezo$
    Call dbxki("KKBZ", kbrec$, ";", "U", "G", hiba%)
    Call dbxtrvege
    

              
    fejr$ = Space$(300)
    Mid$(fejr$, 24, 6) = Mid$(pszbrec$, 84, 6)  ' Teljesítés kelte
    Mid$(fejr$, 18, 6) = maidatum$  ' Számla kelte
    Mid$(fejr$, 30, 6) = maidatum$  ' Lejárat kelte
    Mid$(fejr$, 36, 2) = "01"       ' Fizetési mód
    Mid$(fejr$, 171, 8) = ""        ' Üzletkötõ
    Mid$(fejr$, 111, 60) = Nyugel1.Text8.Text
    Mid$(fejr$, 171, 58) = Nyugel1.Text11.Text
    
    tetdar% = 0
    For i1% = 1 To 200
      tkod$ = Left(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) + Space(15), 15)
      If Trim(tkod$) <> "" Then
        tetdar% = tetdar% + 1
        termrec$ = dbxkey("KTRM", tkod$)
        ktrec$ = Space$(130)
        psztrec$ = Space$(170)
        Mid$(psztrec$, 130, 6) = Mid$(pszbrec$, 84, 6)
        Mid$(psztrec$, 1, 10) = nyugtaszam$
        Mid$(psztrec$, 122, 8) = terminal$ + " Gr"
        Mid$(psztrec$, 11, 3) = Right$("000" + Trim$(Str$(tetdar%)), 3)
        Mid$(psztrec$, 107, 15) = tkod$
        menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
        Mid$(psztrec$, 21, 12) = ertszam(Str(menny@), 12, 3)
        Mid$(psztrec$, 33, 6) = Mid$(termrec$, 484, 6)
        blar@ = xval(Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 4)))
        Mid$(psztrec$, 57, 12) = ertszam(Str(blar@), 12, 2)
        enge@ = -xval(Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 5)))
        blar@ = blar@ + blar@ * enge@ / 100
        blar@ = xval(Format(blar@, "###########0.00"))
        afakod$ = Mid$(termrec$, 706, 2)
        afrec$ = dbxkey("PAFA", afakod$)
        afakulcs@ = xval(Mid$(afrec$, 33, 6))
        elar@ = blar@ / ((100 + afakulcs@) / 100)
        Mid$(psztrec$, 51, 6) = ertszam(Str(enge@), 6, 2)
        Mid$(psztrec$, 39, 12) = ertszam(Str(elar@), 12, 2)
        Mid$(psztrec$, 69, 12) = ertszam(Str(elar@), 12, 2)
        Mid$(psztrec$, 81, 2) = Mid$(termrec$, 706, 2)
        Mid$(ktrec$, 8, 7) = Mid$(kbrec$, 1, 7)
        Mid$(ktrec$, 15, 6) = Mid$(kbrec$, 8, 6)
        Mid$(ktrec$, 24, 4) = ttkrak
        ' Mozgás kód - visszaáru
        menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
        Mid$(ktrec$, 71, 12) = ertszam(Str(-menny@), 12, 3)
        
        If menny@ >= 0 Then
          Mid$(ktrec$, 21, 3) = Mid$(irec$, 631, 3)
        Else
          Mid$(ktrec$, 21, 3) = Mid$(irec$, 655, 3)
        End If
        ' ide kell betenni a raktárat a munkalapból készült számlába
        ' 2015.7.29
        Mid$(ktrec$, 24, 4) = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 8) ' ttkrak
        If munkalapbol Then
           kkftiktato$ = Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 6))
           
           If Len(kkftiktato$) = 7 Then
             kkftml$ = dbxkey("KKFT", kkftiktato$)
             If Not kkftml$ = "" Then
                Mid$(ktrec$, 24, 4) = Mid$(kkftml$, 24, 4)
             End If
           End If
        End If
        
        Mid$(ktrec$, 36, 15) = tkod$
        Mid$(ktrec$, 59, 12) = ertszam(Mid$(termrec$, 554, 14), 12, 2)
        Call dbxtrkezd("KKFT")
        Call dbxki("KKFT", ktrec$, ";", "U", "G", hiba%)
        Call dbxtrvege
        
        Mid(psztrec$, 158, 7) = Mid$(ktrec$, 1, 7)
        Call dbxtrkezd("KSZT")
        Call dbxki("KSZT", psztrec$, ";", "U", "", hiba%)
        Call dbxtrvege
        
        Call GyariszamRogzit(Mid$(ktrec$, 1, 7), nyugtaszam$, i1%, Mid$(kbrec$, 1, 7))
        
        If i1% = 1 Then
          Call dbxtrkezd("KKBZ")
          Mid$(kbrec$, 122, 7) = Mid$(ktrec$, 1, 7)
          Call dbxki("KKBZ", kbrec$, ";", "", "", hiba%)
          Call dbxtrvege
        End If
        ' 2011.10.14
        Call dbxtrkezd("KSZT")
        w1% = obsorszama("KSZT")
        aktucim& = OBJTAB(w1%).obcim
        Call lancra("AUWKER", "KSZTERM", termrec$, aktucim&, psztrec$)
        w1% = obsorszama("KKFT")
        aktucim& = OBJTAB(w1%).obcim
        Call lancra("AUWKER", "KKFKFRG", termrec$, aktucim&, ktrec$)
        Call keszletvalt("K", ktrec$, "N")
        Call dbxtrvege
       
       ' Eszi - megrendelõ kezelés
        If kmegrec$ <> "" Then
          kimenny@ = Abs(xval(Mid$(ktrec$, 71, 12)))
          tkod$ = Mid$(ktrec$, 36, 15)
          rakod$ = Mid(kmegrec$, 480, 4)
          
          For i111% = 1 To 1000
            If i111% <= 200 Then
              ele2$ = Mid$(kmegrec$, (i111% - 1) * 59 + 500, 59)
              If Mid$(ele2$, 1, 15) = Mid$(ktrec$, 36, 15) Then
                Call hozzad(ele2$, 48, 12, kimenny@, 3)
                If i111% <= 200 Then Mid$(kmegrec$, (i111% - 1) * 59 + 500, 59) = ele2$
                Call dbxtrkezd("KMEG")
                Call dbxki("KMEG", kmegrec$, ";", "", "", hiba%)
                Call dbxtrvege
                If Mid$(kmegrec$, 192, 1) = "D" Then
                   minta$ = Mid$(ele2$, 16, 8)
                   Call foglal("N", rakod$, tkod$, minta$, kimenny@)
                End If
                
                Exit For
              End If
            End If
          Next
        End If

      End If
    Next
    ' Eszi - megrendelõ kezelés - lezárás
    If kmegrec$ <> "" Then
       Mid$(kmegrec$, 406, 1) = "S"
       Mid$(kmegrec$, 407, 8) = Mid$(szamlaszam$, 3, 8)
       Mid$(kmegrec$, 484, 6) = maidatum$
       Call dbxtrkezd("KMEG")
       Call dbxki("KMEG", kmegrec$, ";", "", "", hiba%)
       Call dbxtrvege
    End If

    Call dbxtrkezd("KSZB")
    Mid$(pszbrec$, 50, 3) = Right$("   " + Str$(tetdar%), 3)
    Call dbxki("KSZB", pszbrec$, ";", "", "", hiba%)
    Call dbxtrvege
    If szallitobol Then
      ' Eredeti szállító sztonója
      Call szallsztor(szallito$, szamlaszam$)
    End If
        
    ' Partner név,cím letárolása
    Call partnevcimtarol(kbrec$, pszbrec$, nyugtavolt)
    
    'Call dbxtrkezd("KSYB")
    'psybrec$ = Space$(200)
    'Mid$(psybrec$, 1, 7) = Mid$(kbrec$, 1, 7)
    'Mid$(psybrec$, 8, 15) = Mid$(pszbrec$, 61, 15)
    'Mid$(psybrec$, 23, 60) = Nyugel1.Text2          ' Partner neve
    'Mid$(psybrec$, 83, 60) = Nyugel1.Text3          ' Partner neve
    'Mid$(psybrec$, 143, 58) = Nyugel1.Text11        ' Megjegyzés 2. sor
    

    'Call KPAR_Tolt


    'Call dbxki("KSYB", psybrec$, ";", "U", "", hiba%)
    
    'Call dbxtrvege
    ' Eszi
    ' Bizonylat nyontatás errõl is
'          fi3 = FreeFile
'          Open listautvonal$ + terminal$ + task$ + "szle.lst" For Output As #fi3
    kpsszamla = 0
    GoSub szamlair
'          Close fi3
'          Shell programutvonal$ + "dbx4-sho.exe " + terminal$ + task$ + "szle/" + listautvonal$, vbNormalFocus
    Else
      Call mess("Elõleget számlába számítson be!", 2, 0, langprg(1), valasz%)
    End If
  
  
  Case 4
    ' Egyébb készlet változás
'          Shell programutvonal$ + "auw-qregy AUWKER/KTRM/" + terminal$ + task$ + "/" + ugyintezo$ + "/" + auditorutvonal$ + "/1"
'          Call gomb(" Tovább &", gg%, 8320, 100, "V")
     Futtat (programutvonal$ + "auw-qregy AUWKER/KTRM/" + terminal$ + task$ + "/" + ugyintezo$ + "/" + auditorutvonal$ + "/1")
     
  Case 5
'           Kill auditorutvonal$ + "\" + terminal$ + task$ + "szla.txt"
    ' MUNKALAP
'           Futtat ("C:\Program Files\Win Investor RT\Win Munkalap és Garancia Nyilvántartás\munkalap.exe " + ugyintezo$ + "  " + auditorutvonal$ + "\" + terminal$ + task$)
     'Futtat (programutvonal$ + "auw-qgsml AUWKER/GSML/" + terminal$ + task$ + "/" + ugyintezo$ + "/" + auditorutvonal$ + "/1")
munkaujra:
     If UCase(Dir$(auditorutvonal$ + "\" + terminal$ + task$ + "szla.txt")) = UCase(terminal$ + task$ + "szla.txt") Then
       ' feltölt
       sf = FreeFile
       Open auditorutvonal$ + terminal$ + task$ + "szla.txt" For Input As sf
       Line Input #sf, tdb$
       Line Input #sf, partkod$
       Line Input #sf, nev$
       Line Input #sf, cim$
       Line Input #sf, munkalap$
       Line Input #sf, teljkelt$
       
       
       Nyugel1.Text4(1) = partkod$
       Nyugel1.Text2 = nev$
       Nyugel1.Text3 = cim$
       
       Nyugel1.Text8 = munkalap$
       Nyugel1.Text10 = munkalap$
       
       Nyugel1.Text9 = teljkelt$
       
       
       tho = Val(Mid$(teljkelt$, 3, 2) + Mid$(teljkelt$, 6, 2))
       aho = Val(Mid$(Text8.Text, 1, 4))
       anap = Val(Mid$(Text8.Text, 5, 2))
       ' Ha a teljesítés két hónappal korábbi vagy ha a számla kelte napja 20 vagy annál nagyobb
       If (tho < aho - 1) Or ((tho < aho) And (anap >= 20)) Then
         teljkelt$ = "20" + Mid$(Text8.Text, 1, 2) + "." + Mid$(Text8.Text, 3, 2) + "." + Mid$(Text8.Text, 5, 2)
         Nyugel1.Text9 = teljkelt$
       End If
       prec$ = dbxkey("PART", partkod$)
       If prec$ = "" Then
          prec$ = Space$(340)
       End If
       fmkod$ = Mid$(prec$, 328, 2)
       If Not Trim$(fmkod$) = "" Then
          fmrec$ = dbxkey("PFIZ", fmkod$)
          Nyugel1.Text5.Text = Mid$(fmrec$, 3, 30)
          Nyugel1.Text7.Text = fmkod$
       End If

       fizhatido% = Val(Mid$(prec$, 330, 3))
       
       fidat$ = maidatum$
       For i13% = 1 To fizhatido%
          xxxx$ = novdat(fidat$)
          fidat$ = xxxx$
       Next
       Nyugel1.Text6.Text = fidat$

       
       
       pos = InStr(munkalap$, " ")
       munkalap$ = Mid$(munkalap$, 1, pos - 1)
       Call Nyugel1.GyariszamTorol
       
       'Meg kell keresni az elsõ üreset - több munkalap
       ' Ugyanazt a munkalapot kétszer ne lehessen felvinni.
       ' Most nem engedi a munkalapot felvinni
       
       For j1% = 1 To 200
         If Trim(Nyugel1.MSFlexGrid1.TextMatrix(j1%, 1)) = "" Then
            Exit For
         End If
       Next
       teteldb% = Val(tdb$)
       If j1% - 1 + teteldb% > 200 Then
         Call mess("Ez a munkalap már nem fér a számlára!", 2, 0, langprg(1), valasz%)
         Close sf
       Else
       For i1% = j1% To teteldb%
         Line Input #sf, tetel$
         mennyiseg@ = Val(Mid$(tetel$, 90, 14))
'               If mennyiseg@ > 0 Then
           termkod$ = Mid$(tetel$, 1, 15)
           Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) = termkod$
           megnev$ = Mid$(tetel$, 16, 60)
           Nyugel1.MSFlexGrid1.TextMatrix(i1%, 2) = megnev$
           egysegar$ = Mid$(tetel$, 77, 14)
           Nyugel1.MSFlexGrid1.TextMatrix(i1%, 4) = egysegar$
           mennyisegs$ = Mid$(tetel$, 90, 14)
           Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3) = mennyisegs$
           Nyugel1.MSFlexGrid1.TextMatrix(i1%, 7) = mennyisegs$
           kkftiktato$ = Mid$(tetel$, 104, 7)
           Nyugel1.MSFlexGrid1.TextMatrix(i1%, 6) = kkftiktato$
           ' munkalap raktárkód
           ' 2015.7.29
           kkftrec$ = dbxkey("KKFT", kkftiktato$)
           mrakt$ = Mid$(kkftrec$, 24, 4)
           Nyugel1.MSFlexGrid1.TextMatrix(i1%, 8) = mrakt$
           

           Line Input #sf, gyarisz$
           Do While Not (gyarisz$ = "*")
             poz = InStr(gyarisz$, "|")
             hivatkoz$ = Mid$(gyarisz$, poz + 1, Len(gyarisz$) - poz)
             gyarisz$ = Mid$(gyarisz$, 1, poz - 1)
             ' hivatkozást is be kellene tölteni
             Call Nyugel1.GyariszamFeltolt(i1%, gyarisz$, hivatkoz$)
             Line Input #sf, gyarisz$
           Loop
'               Else
           ' A sztornózott tételek kezelése
'               End If
       Next
       For i1% = 1 To teteldb%
         mennyiseg@ = Val(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
         If mennyiseg@ < 0 Then
           termkod$ = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1)
           kkftiktato$ = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 6)
           For j1% = 1 To teteldb%
              mennyiseg2@ = Val(Nyugel1.MSFlexGrid1.TextMatrix(j1%, 3))
              termkod2$ = Nyugel1.MSFlexGrid1.TextMatrix(j1%, 1)
              kkftiktato2$ = Nyugel1.MSFlexGrid1.TextMatrix(j1%, 6)
              If termkod$ = termkod2$ And kkftiktato$ = kkftiktato2$ And Not j1% = i1% Then
                 Nyugel1.MSFlexGrid1.TextMatrix(j1%, 3) = Right$(Space$(6) + Str$(mennyiseg2@ + mennyiseg@), 6)
                 Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3) = "     0"
                 Exit For
              End If
           Next
         End If
       Next
       
       For i1% = teteldb% To 1 Step -1
         mennyiseg@ = Val(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
         If mennyiseg@ = 0 Then
            Nyugel1.MSFlexGrid1.Row = i1%
            Call Nyugel1.sorttorol
                            
          End If
       Next
       Nyugel1.Text1 = Nyugel1.MSFlexGrid1.TextMatrix(1, 1)
       
       Call Nyugel1.ujraszamol
       If Not xval(Trim(Nyugel1.Label1.Caption)) = 0 Then
        Nyugel1.Command11.Enabled = False
        Nyugel1.Command10.Enabled = False
        Nyugel1.Command9.Enabled = False
        Nyugel1.Command14.Enabled = False
        Nyugel1.Command17.Enabled = False
        Nyugel1.Command15.Enabled = False
        Nyugel1.Command6.Enabled = False
        Nyugel1.Command12.Enabled = False
        'Nyugel1.Command2.Enabled = False
        Nyugel1.Text9.Enabled = False

       End If
       Close sf
       
       For j1% = 1 To 200
         If Trim(Nyugel1.MSFlexGrid1.TextMatrix(j1%, 1)) = "" Then
            Exit For
         End If
       Next
       
       Nyugel1.Text1 = Nyugel1.MSFlexGrid1.TextMatrix(j1%, 1)
       Nyugel1.MSFlexGrid1.Row = j1%
       
       End If
       
       Nyugel1.Show vbModal
       ' Eszi most
       If UCase(Dir$(auditorutvonal$ + "\" + terminal$ + task$ + "szla.txt")) = UCase(terminal$ + task$ + "szla.txt") Then
            Kill auditorutvonal$ + "\" + terminal$ + task$ + "szla.txt"
       End If
       munkalapbol = True
       ' Kpsszamla=
       'GoTo szamlaz
       GoSub szamlaz
       
       munkalap$ = ""
       
     Else
       Futtat ("C:\Program Files\Win Investor RT\Win Munkalap és Garancia Nyilvántartás\munkalap.exe " + ugyintezo$ + "  " + auditorutvonal$ + "\" + terminal$ + task$)
       If UCase(Dir$(auditorutvonal$ + "\" + terminal$ + task$ + "szla.txt")) = UCase(terminal$ + task$ + "szla.txt") Then
         GoTo munkaujra
       End If
     End If
             
  Case 6
    ' Megrendelés
    
megrujra:
    ' If UCase(Dir$(auditorutvonal$ + "\" + terminal$ + task$ + "szlm.txt")) = UCase(terminal$ + task$ + "szlm.txt") Then
    '   ' feltölt
    '   sf = FreeFile
    '   Open auditorutvonal$ + terminal$ + task$ + "szlm.txt" For Input As sf
    '   Line Input #sf, tdb$
    '   Line Input #sf, partkod$
    '   Line Input #sf, nev$
    '   Line Input #sf, cim$
    '   Line Input #sf, megrend$
    '   Line Input #sf, teljkelt$
    '   teljkelt$ = ""
       
    '   Nyugel1.Text4(1) = partkod$
    '   Nyugel1.Text2 = nev$
    '   Nyugel1.Text3 = cim$
    '   Nyugel1.Text8 = megrend$ + " számú megrendelés"
    '   Nyugel1.Text9 = teljkelt$
    '   kmegrec$ = dbxkey("KMEG", megrend$)
       
    '   fmkod$ = Mid$(kmegrec$, 197, 2)
    '   fmrec$ = dbxkey("PFIZ", fmkod$)
    '   Nyugel1.Text5.Text = Mid$(fmrec$, 3, 30)
    '   Nyugel1.Text7.Text = fmkod$
    '   fizhatido% = Val(Mid$(kmegrec$, 199, 3))
    '   fidat$ = maidatum$
    '   For i13% = 1 To fizhatido%
    '     xxxx$ = novdat(fidat$)
    '     fidat$ = xxxx$
    '   Next
    '   Nyugel1.Text6.Text = fidat$

       
       
       
       
     '  teteldb% = Val(tdb$)
     '  Call Nyugel1.GyariszamTorol
     '  For i1% = 1 To teteldb%
     '    Line Input #sf, tetel$
     '    mennyiseg@ = Val(Mid$(tetel$, 90, 14))
     '    If mennyiseg@ > 0 Then
     '      termkod$ = Mid$(tetel$, 1, 15)
     '      Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) = termkod$
     '      megnev$ = Mid$(tetel$, 16, 60)
     '      Nyugel1.MSFlexGrid1.TextMatrix(i1%, 2) = megnev$
     '      egysegar$ = Mid$(tetel$, 77, 14)
     '      Nyugel1.MSFlexGrid1.TextMatrix(i1%, 4) = egysegar$
     '      mennyisegs$ = Mid$(tetel$, 90, 14)
     '      Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3) = mennyisegs$
     '     End If
     
     '  Next
     '  Close sf
     '  Call Nyugel1.ujraszamol
     '  If Not xval(Trim(Nyugel1.Label1.Caption)) = 0 Then
     '   Nyugel1.Command11.Enabled = False
     '   Nyugel1.Command10.Enabled = False
     '   Nyugel1.Command9.Enabled = False
     '   Nyugel1.Command14.Enabled = False
     '   Nyugel1.Command17.Enabled = False
     '   Nyugel1.Command15.Enabled = False
     '   Nyugel1.Command6.Enabled = False
     '   Nyugel1.Command12.Enabled = False
        
        
     '  End If
       
       
     '  Nyugel1.Text1 = Nyugel1.MSFlexGrid1.TextMatrix(1, 1)
       
     '  megrendelesbol = True
       
     '  Nyugel1.Show vbModal
       ' Eszi Most
     '  If UCase(Dir$(auditorutvonal$ + "\" + terminal$ + task$ + "szlm.txt")) = UCase(terminal$ + task$ + "szlm.txt") Then
     '      Kill auditorutvonal$ + "\" + terminal$ + task$ + "szlm.txt"
     '  End If
       
       
     '  GoTo szamlaz
       
    
    'Else
    '   Futtat (programutvonal$ + "auw-qrmeg AUWKER/KMEG/" + terminal$ + task$ + "/" + ugyintezo$ + "/" + auditorutvonal$ + "/1")
    '   If UCase(Dir$(auditorutvonal$ + "\" + terminal$ + task$ + "szlm.txt")) = UCase(terminal$ + task$ + "szlm.txt") Then
    '     GoTo megrujra
    '   End If
    'End If
  Case 8
    ' Pénztár
    
    Futtat (programutvonal$ + "auw-qptrg AUWSZAMV/PKTE/" + terminal$ + task$ + "/" + ugyintezo$ + "/" + auditorutvonal$ + "/1")
    
  Case 10
  ' Sztornó, másolat
    szalind% = 0
    voltszall = False
ujravallaszt:
   ' Ellenõrzés: partner ugyanaz lehet, kétszer nem lehet ugyanazt kiválasztani
    Call dbxker("Sztornó&Másolat&Másolat képernyõre&KP szla név&Számla szállítóról&Származtatás&Szállító lista törlése", "KSZB", 0, talalat%, rec$)
    melyik% = 0
    If gombsorszam% = 3 Then
       melyik% = gombsorszam%
       gombsorszam% = 2
    End If
    szamlaeset% = gombsorszam%
    gombs1% = gombsorszam%
    
mnyomtat:
    ' Számla szállítóról esetén ellenõrizni, hogy szállító-e a tétel
    ' Származtatást munkalapos és megrendelõs tételre nem szabad engedni.
    If szamlaeset% = 1 Or szamlaeset% = 2 Or szamlaeset% = 5 Or szamlaeset% = 6 Or (szamlaeset% = 0 And voltszall) Then
      '--- sztornó bizonylat, vagy számla másolat nyomtatása
      If (szamlaeset% = 0 And voltszall) Then
        Nyugel1.Text8.Text = Nyugel1.Text8.Text + ". szállító levél"
        gombs1% = 5
        GoTo folytat
      End If
      masolat% = 0: sztornoszamla% = 0: sztornomasolat% = 0
      szoveg18$ = ""
      irec$ = dbxkey("INST", "INST")
      sirec$ = dbxkey("SINS", "INST")
      If rec$ <> "" Then
      For i1% = 1 To 1000
        mt$(i1%) = Space$(120)
        regimt$(i1%) = Space$(120)
      Next
      xrec$ = Space$(700)
      vsorszam& = 0
      w1% = obsorszama("KSZF")
      For i1% = 1 To 5: nt$(i1%) = Space$(43): Next
      '--- számla beolvasása
      szamla1$ = Mid$(rec$, 1, 10)
      erxszamla$ = szamla1$
      pszbrec$ = dbxkey("KSZB", szamla1$)
      szallevszam$ = Mid$(pszbrec$, 250, 10)
      If Mid$(pszbrec$, 76, 2) = "SL" Then
         szallitasicim$ = Mid$(pszbrec$, 260, 15)
      End If
      munkalap$ = Mid$(pszbrec$, 171, 60)
      ppos = InStr(munkalap$, "számú munkalap")
      If ppos > 0 Then
        munkalap$ = Trim(Mid$(munkalap$, 1, ppos - 1))
      Else
        munkalap$ = ""
      End If
      
      megrendelesiktato$ = Mid$(pszbrec$, 275, 7)
      teljesdat$ = Mid$(pszbrec$, 84, 6)
      konyveldat$ = Mid$(pszbrec$, 231, 6)
      fejr$ = Mid$(pszbrec$, 61, 170) + Mid$(pszbrec$, 237, 8) + Space$(122)
      Mid$(fejr$, 111, 60) = Mid$(pszbrec$, 111, 60)
      ' export kiegészítés
      'psxbrec$ = dbxkey("KSXB", szamla1$)
      'If psxbrec$ <> "" Then xrec$ = Mid$(psxbrec$, 12, 689)
      pkod$ = Mid$(fejr$, 1, 15)
      partrec$ = dbxkey("PART", pkod$)
      ' --- hibaellenõrzés
      If Mid$(pszbrec$, 35, 1) = "S" Then
        If gombs1% = 1 Or gombs1% = 5 Then
          Call mess("Sztornózott", 2, 0, langprg(1), valasz%)
        
          pszbrec$ = ""
        Else
          If gombs1% = 2 Then
             sztornomasolat% = 1
          End If
        End If
      Else
        If gombs1% = 1 Then
        '--- sztornózhatóság logikai vizsgálata
          sztornoszamla% = 1
          If dtm(konyveldat$) <= dtm(lezardat$) Then
             pszbrec$ = ""
             Call mess("Lezárt idõszak nem sztornóztahó", 2, 0, langprg(1), valasz%)
          End If
          '--- vevo rekord beolvasása
          vikt$ = Mid$(pszbrec$, 53, 7)
          If Trim$(vikt$) <> "" Then
            vrec$ = dbxkey("PVSZ", vikt$)
            kio@ = 0
            For i15% = 1 To 10
              kio@ = kio@ + xval(Mid$(vrec$, (i15% - 1) * 35 + 930, 14))
            Next
            If kio@ <> 0 Then
              pszbrec$ = ""
              Call mess("Már volt kiegyenlítés erre a számlára", 2, 0, langprg(1), valasz%)

            End If
          End If
          herec$ = dbxkey("PSHL", "V" + vikt$)
          If herec$ <> "" Then
            For i15% = 1 To 3
              elem3$ = Mid$(herec$, (i15% - 1) * 790 + 1, 790)
              If Mid$(elem3$, 86, 1) <> "S" Then
                If xval(Mid$(elem3$, 36, 14)) <> 0 Then
                  If Mid$(elem3$, 9, 10) <> szamla1$ Then
                    pszbrec$ = ""
                    Call mess("Már volt helyesbítés erre a számlára", 2, 0, langprg(1), valasz%)
                    
                    Exit For
                  End If
                End If
              End If
            Next
            herec$ = "": elem3$ = ""
          End If
        Else
          masolat% = 1
          If gombs1% = 5 Or gombs1% = 6 Then
             masolat% = 0
          End If
        End If
        ' Eszi - elõleg beolvasása
        pvszikt$ = Mid$(pszbrec$, 53, 7)
        ' Ha számla, nem szállító levél
        If Not Trim$(pvszikt$) = "" Then
          pvszrec$ = dbxkey("PVSZ", pvszikt$)
          For i9% = 1 To 5
            nt$(i9%) = Mid$(pvszrec$, (i9% - 1) * 43 + 1280, 43)
                              
            If Not nt$(i9%) = "" Then
                            ' partner kód, számlaszám
              ntafa$(i9%) = elolegafa(Mid$(pvszrec$, 38, 15), Mid$(nt$(i9%), 8, 10))
            End If
            ' elolegafa()
          Next
        Else
          trind$ = szamla1$ + "001"
          psztrec$ = dbxkey("KSZT", trind$)
          ktetikt$ = Mid$(psztrec$, 158, 7)
          kfttrec$ = dbxkey("KKFT", ktetikt$)
          If kfttrec$ <> "" Then
            ' 2015.7.29
            mrakt$ = Mid$(kfttrec$, 24, 4)
            
            bizikt$ = Mid(kfttrec$, 8, 7)
            ksybrec$ = dbxkey("KSYB", bizikt$)
          End If
          For i9% = 1 To 4
             nt$(i9%) = Mid$(ksybrec$, (i9% - 1) * 36 + 143, 36)
             If Not Trim$(nt$(i9%)) = "" Then
               ' partner kód, számlaszám
                 ntafa$(i9%) = elolegafa(Mid$(pszbrec$, 61, 15), Mid$(nt$(i9%), 8, 10))
             End If
          Next

               

        End If
      End If
      If pszbrec$ <> "" Then
        If gombs1% = 1 Then
          Call elsolap("KSZF", fejr$)
          Call sztornokelt(targyev$, Mid$(pszbrec$, 231, 6), Mid$(pszbrec$, 78, 6), irec$, "S", sztordat$, sztornosz$)
          If Trim(sztordat$) = "" Then
            pszbrec$ = ""
          Else
            ' Sztornó dátum
            Mid$(fejr$, 30, 6) = sztordat$
            Mid$(fejr$, 24, 6) = Mid$(pszbrec$, 84, 6)
            Mid$(fejr$, 18, 6) = sztordat$
          End If
        End If
      End If
      If gombs1% = 6 And (Not Trim$(munkalap$) = "" Or Not Trim$(megrendelesiktato$) = "") Then
        pszbrec$ = ""
        Call mess("Munkalapról ill megrendelõrõl készült számla nem származtatható!", 2, 0, langprg(1), valasz%)
      End If
      If gombs1% = 5 And Not Mid$(pszbrec$, 254, 6) = "Száll." Then
        pszbrec$ = ""
        
        Call mess("Ez nem szállítólevél!", 2, 0, langprg(1), valasz%)

      End If
      If gombs1% = 5 Then
        'Mid$(pszbrec$, 1, 10) = ""
        For i10% = 0 To form1.List1.ListCount - 1
            If form1.List1.List(i10%) = Mid$(pszbrec$, 1, 10) Then
               Call mess("Ezt a szállítólevelet már kiválasztotta!", 2, 0, langprg(1), valasz%)
               pszbrec$ = ""
            End If
        Next
        If form1.List1.ListCount = 0 Then
           szallpartner = Mid$(pszbrec$, 61, 15)
        Else
           If Not szallpartner = Mid$(pszbrec$, 61, 15) And Not pszbrec$ = "" Then
              Call mess("Partner nem egyezik!", 2, 0, langprg(1), valasz%)
              pszbrec$ = ""
           End If
        End If
        
        
        

      End If

      If pszbrec$ <> "" Then
        trdarab% = xval(Mid$(pszbrec$, 50, 3))
        
        If (szalind% + trdarab%) < 200 Then
        
        trind$ = szamla1$ + "001"
        psztrec$ = dbxkey("KSZT", trind$)
        If psztrec$ <> "" Then
          w1% = obsorszama("KSZT")
          kezdoix& = OBJTAB(w1%).obind
          dbfi = FreeFile
          Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #dbfi
          ndfi = FreeFile
          Open auditorutvonal$ + "auw-kszt.ndx" For Binary Shared As #ndfi
          rc& = Int(LOF(ndfi) / 12)
          For i1% = 1 To trdarab%
            i1d& = kezdoix& + i1% - 1
            Get #ndfi, (i1d& - 1) * 18& + 1, rcim&
            Seek #dbfi, rcim& + 9
            psztrec$ = Space(170): Get #dbfi, , psztrec$
            mt$(i1%) = Mid$(psztrec$, 15, 120)
            ktikt$ = Mid$(psztrec$, 158, 7)
            'kszxrec$ = dbxkey("KSZX", ktikt$)
            ' ÁFA kód itt
            afakod$ = Mid$(psztrec$, 81, 2)
            Nyugel1.MSFlexGrid1.TextMatrix(szalind% + i1%, 7) = afakod$
            
            Nyugel1.MSFlexGrid1.TextMatrix(szalind% + i1%, 1) = Mid$(psztrec$, 107, 15)
            Nyugel1.MSFlexGrid1.TextMatrix(szalind% + i1%, 3) = Mid$(psztrec$, 21, 12)
            Nyugel1.MSFlexGrid1.TextMatrix(szalind% + i1%, 4) = Mid$(psztrec$, 57, 12)
            If Mid$(psztrec$, 157, 1) = "B" Then
               ' Elõleg
               Nyugel1.MSFlexGrid1.TextMatrix(i1%, 6) = "Elõleg besz."
               Nyugel1.MSFlexGrid1.TextMatrix(i1%, 7) = Mid$(psztrec$, 150, 7)
               wbeszeloikt(i1%) = Mid$(psztrec$, 150, 7)
            End If
            If gombs1% = 5 Or gombs1% = 6 Then
               termkod$ = Mid$(psztrec$, 107, 15)
               termrec$ = dbxkey("KTRM", termkod$)
               Nyugel1.MSFlexGrid1.TextMatrix(szalind% + i1%, 2) = Mid$(termrec$, 16, 60)
               Nyugel1.MSFlexGrid1.TextMatrix(szalind% + i1%, 7) = Mid$(termrec$, 484, 6)
               
               krak$ = Left(Trim(form1.Text5.Text) + "    ", 4)
               rkszkod$ = krak$ + termkod$
               rkszrec$ = dbxkey("RKSZ", rkszkod$)
               keszle@ = xval(Mid$(rkszrec$, 20, 12))
               ' 2015.7.29
               Nyugel1.MSFlexGrid1.TextMatrix(szalind% + i1%, 8) = mrakt$
               
               If gombs1% = 5 Then
                  keszle@ = keszle@ + Val(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
               End If
               If jelleg$ = "S" Then
                 Nyugel1.MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = "Szolgáltatás"
               Else
                 If gombs1% = 5 Then
                    Nyugel1.MSFlexGrid1.TextMatrix(szalind% + i1%, 6) = Mid$(psztrec$, 1, 10)
                 Else
                    Nyugel1.MSFlexGrid1.TextMatrix(szalind% + i1%, 6) = Str(keszle@)
                 End If
               End If


               
               
            End If
            If gombs1% = 5 And i1% = 1 Then
               Nyugel1.Text1.Text = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1)
               'Nyugel1.Text8.Text = Mid$(pszbrec$, 1, 10) + ". szállító levél"
               If Trim$(Nyugel1.Text8.Text) = "" Then
                  vesszo = ""
               Else
                  vesszo = ","
               End If
               'Nyugel1.Text8.Text = Nyugel1.Text8.Text + vesszo + Mid$(pszbrec$, 1, 10)
               ' vezetõ nullák elhagyása
               szlaszam$ = Mid$(pszbrec$, 1, 10)
               For i11% = 1 To 10
                 If Mid$(szlaszam$, i11%, 1) = "0" Then
                   Mid$(szlaszam$, i11%, 1) = " "
                 Else
                   If Mid$(szlaszam$, i11%, 1) > "0" Then
                      Exit For
                   End If
                 End If
               Next
               szlaszam$ = LTrim$(szlaszam$)
               
               Nyugel1.Text8.Text = Nyugel1.Text8.Text + vesszo + szlaszam$
               Nyugel1.Text9.Text = Mid$(pszbrec$, 84, 6)

               
               szallito$ = Mid$(pszbrec$, 1, 10)
            End If
'                  If masolat% = 1 Then
              Nyugel1.MSFlexGrid1.TextMatrix(i1%, 5) = Str$(-Val(Mid$(psztrec$, 51, 6)))
'                  Else
'                     Nyugel1.MSFlexGrid1.TextMatrix(i1%, 5) = Mid$(psztrec$, 51, 6)
'                  End If
                              ' IDe
            ktetikt$ = Mid$(psztrec$, 158, 7)
            kfttrec$ = dbxkey("KKFT", ktetikt$)
            ' 2015.03.17
            If gombs1% = 5 Then
               ttkrak$ = Mid$(kfttrec$, 24, 4)
            End If
            If kfttrec$ <> "" Then
               bizikt$ = Mid(kfttrec$, 8, 7)
               ksybrec$ = dbxkey("KSYB", bizikt$)
               Call feltolt2(ksybrec$, pkod$)
               If Mid$(pszbrec$, 254, 1) = "H" Then
                   If Mid$(pkod$, 1, 1) = "T" Then
                     kod1$ = Trim$(Mid$(partrec$, 363, 60))
                     kod2$ = Trim$(Mid$(partrec$, 423, 60))
                  
                    Nyugel1.Text8.Text = kod1$
                    Nyugel1.Text11.Text = kod2$
                  End If
               End If
               ' Másolat Támop számláról nincs név, cam ,adószám
               
               Mid$(fejr$, 171, 58) = Nyugel1.Text11.Text
               ' KSZB-ben benne van a fizetési mód
               'If Trim$(Mid$(KSYBREC$, 8, 15)) = "" Then
               '   Nyugel1.Text5 = "Készpénz"
               'ElseIf Trim$(Mid$(KSYBREC$, 8, 15)) = "BANK" Then
               '   Nyugel1.Text5 = "bankkártya"
               'Else
               '   Nyugel1.Text5 = "Átutalás"
               'End If
               fizmod$ = Mid$(pszbrec$, 96, 2)
               ' 2010.08.10
               If fizmod$ = "  " Then
                  pvszikt$ = Mid$(pszbrec$, 53, 7)
                  pvszrec$ = dbxkey("PVSZ", pvszikt$)
                  If Not pvszrec$ = "" Then
                     fizmod$ = Mid$(pvszrec$, 76, 2)
                  End If
               End If
               fmrec$ = dbxkey("PFIZ", fizmod$)
               Nyugel1.Text5.Text = Mid$(fmrec$, 3, 30)
               Nyugel1.Text7.Text = fizmod$

               If Not gombs1% = 1 Then
                If Mid$(pszbrec$, 53, 7) = Space$(7) Then
               
                  For i9% = 1 To 4
                      nt$(i9%) = Mid$(ksybrec$, (i9% - 1) * 36 + 143, 36)
                      If Not Trim$(nt$(i9%)) = "" Then
                            ' partner kód, számlaszám
                          ntafa$(i9%) = elolegafa(Mid$(pszbrec$, 61, 15), Mid$(nt$(i9%), 8, 10))
                       End If
                  Next
                Else
                  pvszikt$ = Mid$(pszbrec$, 53, 7)
                  pvszrec$ = dbxkey("PVSZ", pvszikt$)
                  For i9% = 1 To 5
                     nt$(i9%) = Mid$(pvszrec$, (i9% - 1) * 43 + 1280, 43)
                              
                      
                      If Not Trim$(nt$(i9%)) = "" Then
                            ' partner kód, számlaszám
                         ntafa$(i9%) = elolegafa(Mid$(pvszrec$, 38, 15), Mid$(nt$(i9%), 8, 10))
                       End If
            
                  Next
                
                End If
               End If
            End If
            Call Nyugel1.GyariszamBetolt(i1%, ktikt$)
                              
          Next
          'Nyugel1.Show vbModal
          szalind% = szalind% + trdarab%
          Close dbfi
          Close ndfi
        End If
        pkod$ = Mid$(fejr$, 1, 15)
        partrec$ = dbxkey("PART", pkod$)
        If Trim$(Nyugel1.Text2) = "" Then
           Call feltolt
           Mid$(fejr$, 171, 58) = Trim$(Nyugel1.Text11.Text) + Space$(58)
        End If
        
        dat$ = Mid$(fejr$, 18, 6)
        If Mid$(fejr$, 17, 1) = "P" Then proforma% = 1 Else proforma% = 0
        If Mid$(fejr$, 16, 1) <> "B" And Mid$(fejr$, 16, 1) <> "S" Then megnbeal.Show vbModal
        'List1.Clear
        proforma% = 0
        rnyugtavolt = nyugtavolt
        Nyugel1.Check1.Value = 0
        Select Case Mid$(pszbrec$, 254, 1)
          Case "N"
             nyugtavolt = 1
        Case "K"
             nyugtavolt = 2
        Case "H", " ", "0"
          nyugtavolt = 3
          If Mid$(pkod$, 1, 1) = "T" Then
            nyugtavolt = 12
            kod1$ = Trim$(Mid$(partrec$, 363, 60))
            kod2$ = Trim$(Mid$(partrec$, 423, 60))
            Nyugel1.Text8.Text = kod1$
            Nyugel1.Text11.Text = kod2$
            ' ide
          End If
        Case "G"
          nyugtavolt = 7
        Case "S"
          nyugtavolt = 7
          Nyugel1.Check1.Value = 1
        Case Else
      End Select
      ' Ez mi?
'            If Mid$(fejr$, 38, 3) <> "   " And Mid$(fejr$, 38, 3) <> "HUF" Then
'              List1.AddItem langprg(39) + " " + Mid$(fejr$, 38, 3) + "=" + Trim$(Mid$(fejr$, 41, 10)) + " " + langprg(41) + " " + langprg(42)
'            Else
'              List1.AddItem langprg(40)
'            End If
'            List1.AddItem Trim$(Mid$(partrec$, 16, 60)) + "  " + Trim$(Mid$(partrec$, 106, 8)) + " " + Trim$(Mid$(partrec$, 114, 30)) + ", " + Trim$(Mid$(partrec$, 144, 30)) + " " + Trim$(Mid$(partrec$, 174, 10))
'            List1.AddItem langprg(43) + Trim$(Mid$(partrec$, 645, 14)) + " " + langprg(44) + Trim$(Mid$(partrec$, 631, 14)) + " " + langprg(45) + Trim$(Mid$(partrec$, 659, 14)) + " " + langprg(46) + Trim$(Mid$(partrec$, 617, 14)) + " " + langprg(47) + Trim$(Mid$(partrec$, 603, 14))
'            List1.Visible = True
      
'            Call elsolap("KSZF", fejr$)
'            Call elsotab("KSZL", mt$(), trdarab%)
' Ide számlaszám tömb feltöltés
      szamlaszam$ = Mid$(pszbrec$, 1, 10)
      '--- számla nyomtatása
      If Mid$(pszbrec$, 11, 10) <> Space$(10) Then
        szoveg18$ = langprg(60) + " " + LCase(langprg(51)) + ":" + Mid$(pszbrec$, 11, 10)
        helyes% = 0
'              jovairas% = 1
      End If
      If helyes% = 1 And Mid$(pszbrec$, 11, 10) <> Space$(10) Then
      Else
        '--- számla
        
        If sztornoszamla% = 1 And sztornomasolat% = 0 Then
          If nyugtavolt <> 0 Then
            
            If nyugtavolt = 3 Or nyugtavolt = 12 Then
              szlacim$ = "S Z Á M L A"
              szlamod$ = "S"
              szlafaj$ = "U"
              ssz$ = szamlaszam$
              'Call dbxtrkezd("INST")
              'irec$ = dbxkey("INST", "INST")
              'sp% = 704
              'sztornoszamlaszam$ = Mid$(irec$, sp%, 4) + novel(irec$, sp% + 4, 6)
              'Mid$(irec$, sp%, 10) = sztornoszamlaszam$
              'Call dbxki("INST", irec$, ";", "", "", hiba%)
              'Call dbxtrvege
            ElseIf nyugtavolt = 2 Then
              szlacim$ = "KÉSZPÉNZFIZETÉSI   S Z Á M L A"
              szlamod$ = "S"
              szlafaj$ = "K"
              ssz$ = szamlaszam$
              'Call dbxtrkezd("SINS")
              'sirec$ = dbxkey("SINS", "INST")
              ' Kézpénzes számla sztornója ugyan azon az intervallumon
              'sp% = 510
              'sp% = 500
              'sztornoszamlaszam$ = Mid$(sirec$, sp%, 4) + novel(sirec$, sp% + 4, 6)
              'Mid$(sirec$, sp%, 10) = sztornoszamlaszam$
              'Call dbxki("SINS", sirec$, ";", "", "", hiba%)
              'Call dbxtrvege
            ElseIf nyugtavolt = 1 Then
              szlamod$ = " "
              Call dbxtrkezd("SINS")
              sirec$ = dbxkey("SINS", "INST")
              sp% = 520
              sztornoszamlaszam$ = Mid$(sirec$, sp%, 4) + novel(sirec$, sp% + 4, 6)
              Mid$(sirec$, sp%, 10) = sztornoszamlaszam$
              Call dbxki("SINS", sirec$, ";", "", "", hiba%)
              Call dbxtrvege
            ElseIf nyugtavolt = 7 Then
              szlamod$ = " "
              sztornoszamlaszam$ = szamlaszam$
            Else
              szlamod$ = " "
              sztornoszamlaszam$ = ""
            End If
'                  Call dbxtrvege
            'GoSub sztornokezel
          End If
        End If
        If sztornomasolat% = 1 Then
           Call gomb(" Eredeti & Sztornó &", gg3%, 8320, 100, "V")
           If gg3% = 1 Then
             sztornoszamlaszam$ = Mid$(rec$, 1, 10)
           Else
             sztornoszamlaszam$ = Mid$(rec$, 282, 10)
           End If

          szlamod$ = "M"
          ssz$ = sztornoszamlaszam$
        End If
        
        If masolat% = 1 Or gombs1% = 5 Then
           Nyugel1.Text4(1).Text = Mid$(pszbrec$, 61, 15)
           szlamod$ = "M"
           ssz$ = szamlaszam$
        End If
        ' Ide : eredetit sztornóról
        If melyik% = 3 Then
        ' ITT MOD
          If ugyintezo$ = "ESZES" Then
'            masolat% = 0
'            sztornomasolat% = 0
''                sztornoszamla% = 1
          End If
        ElseIf melyik% = 4 Then
          masolat% = 0
          sztornomasolat% = 0
          sztornomasolatX% = 1
        End If
        If gombs1% = 5 Then
           ' Újra választ, listába beletesz
           List1.Visible = True
           Label21.Visible = True
           a = List1.ListCount
           voltszall = True
           List1.AddItem (szallito$)
           GoTo ujravallaszt
        End If
folytat:
     
        If gombs1% = 5 Or gombs1% = 6 Then
           If gombs1% = 5 Then
              szallitobol = True
           Else
              szarmaztatasbol = True
           End If
           For j1% = 1 To 200
              If Trim(Nyugel1.MSFlexGrid1.TextMatrix(j1%, 1)) = "" Then
                 Exit For
              End If
           Next
       
           Nyugel1.Text1 = Nyugel1.MSFlexGrid1.TextMatrix(j1%, 1)
           Nyugel1.MSFlexGrid1.Row = j1%
           
           Nyugel1.Text1.Top = Nyugel1.MSFlexGrid1.CellTop + Nyugel1.MSFlexGrid1.Top
           Nyugel1.Text1.SelStart = Len(Trim(Nyugel1.Text1.Text))

           Call Nyugel1.ujraszamol
           If Not xval(Trim(Nyugel1.Label1.Caption)) = 0 Then
              Call Beallit
'              Nyugel1.Command11.Enabled = False
'              Nyugel1.Command10.Enabled = False
'              Nyugel1.Command9.Enabled = False
'              Nyugel1.Command14.Enabled = False
'              Nyugel1.Command17.Enabled = False
'              Nyugel1.Command15.Enabled = False
'              Nyugel1.Command6.Enabled = False
'              Nyugel1.Command12.Enabled = False
              If gombs1% = 5 Then
'               Nyugel1.Text9.Enabled = False
                Nyugel1.Check1.Value = 0
                Nyugel1.Check1.Enabled = False
                Nyugel1.Command4.Enabled = False
              End If

           End If
 
          If szallitobol Then
            Call Partnertolt
          End If
          For i1% = 1 To 200
            tetnevmt(i1%) = Space$(150)
            wbeszeloikt(i1%) = ""
          Next
          
          Nyugel1.Show vbModal
          GoSub szamlaz

        Else
        
          GoSub szamlair
        End If
      End If
      '--- sztornó jel visszaírása
      Call dbxtrkezd("KSZB")
      If sztornoszamla% = 1 And sztornomasolat% = 0 Then
        ' sZTORNÓ KEZEL
        If nyugtavolt = 3 Then
           Call pszsztrn(vikt$, "V", sztordat$, szamlaszam$)
        End If
        pszbrec$ = dbxkey("KSZB", erxszamla$)
        If pszbrec$ <> "" Then
          Mid$(pszbrec$, 35, 1) = "S"
          Mid$(pszbrec$, 36, 6) = sztordat$
          Mid$(pszbrec$, 42, 8) = terminal$ + ugyintezo$
          Mid$(pszbrec$, 282, 15) = szamlaszam$                         'sztornoszamlaszam$
          Call dbxki("KSZB", pszbrec$, ";", "", "", hiba%)
          trdarab% = Val(Mid$(pszbrec$, 50, 3))
          szamla2$ = Mid$(pszbrec$, 1, 10)
          trind$ = szamla2$ + "001"
          psztrec$ = dbxkey("KSZT", trind$)
          If psztrec$ <> "" Then
            w1% = obsorszama("KSZT")
            kezdoix& = OBJTAB(w1%).obind
            dbfi = FreeFile
            Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #dbfi
            ndfi = FreeFile
            Open auditorutvonal$ + "auw-kszt.ndx" For Binary Shared As #ndfi
            rc& = Int(LOF(ndfi) / 18)
            If trdarab% > 0 Then
              For i1% = 1 To trdarab%
                i1d& = kezdoix& + i1% - 1
                Get #ndfi, (i1d& - 1) * 18& + 1, rcim&
                Seek #dbfi, rcim& + 9
                psztrec$ = Space(170): Get #dbfi, , psztrec$
                Mid$(psztrec$, 14, 1) = "S"
                kfrec$ = Mid$(psztrec$, 15, 120)
                Call dbxvir("KSZT", dbfi, psztrec$, rcim& + 9, 150)
                aktucim& = 0
               
                '--- készlet sztornózása
                '
                ktetikt$ = Mid$(psztrec$, 158, 7)
                kfttrec$ = dbxkey("KKFT", ktetikt$)
                If kfttrec$ <> "" Then
                  Mid(kfttrec$, 95, 1) = "S"
                  Call dbxki("KKFT", kfttrec$, ";", "", "", hiba%)
                  ymenny@ = -xval(Mid$(kfttrec$, 71, 12))
                  Mid$(kfttrec$, 71, 12) = ertszam(Str(ymenny@), 12, 3)
                  ymenny@ = -xval(Mid$(kfttrec$, 83, 12))
                  Mid$(kfttrec$, 83, 12) = ertszam(Str(ymenny@), 12, 3)
                  Call keszletvalt("B", kfttrec$, "K")
                  If i1% = 1 Then
                    '--- forgalmi bizonylat sztornózása
                    bizikt$ = Mid(kfttrec$, 8, 7)
                    kbizrec$ = dbxkey("KKBZ", bizikt$)
                    Mid$(kbizrec$, 92, 1) = "S"
                    Mid$(kbizrec$, 93, 6) = sztordat$
                    Mid$(kbizrec$, 99, 8) = ugyintezo$
                    Call dbxki("KKBZ", kbizrec$, ";", "", "", hiba%)
                    ' Eszi - megrendelésre visszatenni
                  End If
                End If
              Next
            End If
            Close dbfi
            Close ndfi
          End If
        End If
        Call Eloleg_beszamitas("S")
      End If
      Call dbxtrvege
      masolat% = 0
      sztornoszamla% = 0
      sztornomasolat% = 0
      Call Command3_Click
  '    Call kilepp
      Exit Sub
      Else
       ' Hiba - nem fér a számlára
       Call mess("Túl sok tétel! Nem fér a számlára.", 2, 0, langprg(1), valasz%)
      End If
    Else
      If gombs1% = 5 Then
         GoTo ujravallaszt
      End If
      masolat% = 0
      sztornoszamla% = 0
      sztornomasolat% = 0
      Call Command3_Click
  '    Call kilepp
      Exit Sub
      
    End If
  End If
  Else
    If gombsorszam% = 3 Then
      Call altx("PART", azonosito$)
      Clipboard.Clear
      Clipboard.SetText (azonosito$)
      
      GoTo ujravallaszt
    End If
    If gombsorszam% = 4 Then
      Call dbxker("Választ", "KSYB", 0, talalat%, rec$)
      If gombsorszam% = 1 Then
        szamlaeset% = 1
        masolat% = 1
        iktato$ = Mid$(rec$, 1, 7)
        kkbzrec$ = dbxkey("KKBZ", iktato$)
        szlaszam$ = Mid$(kkbzrec$, 107, 15)
        rec$ = dbxkey("KSZB", szlaszam$)
        GoTo mnyomtat
      End If
     End If
     If gombsorszam% = 7 Then
       ListEdit1.Show vbModal
       GoTo ujravallaszt
     Else
       List1.Visible = False
       Label21.Visible = False
     End If
   End If
  Case Else
End Select

    Loop While nyugtavolt <> 0
    Command1.Visible = True
    Command3.Visible = True
    Command3.SetFocus
  Else
    Call mess("Raktár kötelezõ!", 2, 0, "Hiba", valasz%)
  End If
  
Exit Sub
folyokonyvel:
  '--- számla tartalmának könyvelése pvsz-be
  pvszrec$ = Space$(1500)
  Call feltolt_folyokonyvel(pvszrec$, szamlaszam$, fejr$)
  'Mid$(pvszrec$, 8, 10) = szamlaszam$
  'Mid$(pvszrec$, 38, 15) = Mid$(fejr$, 1, 15)
  'Mid$(pvszrec$, 211, 6) = Mid$(fejr$, 24, 6)
  'Mid$(pvszrec$, 58, 6) = Mid$(fejr$, 24, 6)
  'Mid$(pvszrec$, 64, 6) = Mid$(fejr$, 18, 6)
  'Mid$(pvszrec$, 70, 6) = Mid$(fejr$, 30, 6)
  'Mid$(pvszrec$, 76, 2) = Mid$(fejr$, 36, 2)
  '' Mid$(pvszrec$, 201, 8) = Mid$(fejr$, 171, 8)
  Mid$(pvszrec$, 1495, 4) = krak$
  arf@ = xval(Mid$(fejr$, 41, 10))
  GoSub szamlaszamol
  If arf@ <> 0 Then
    Mid$(pvszrec$, 92, 3) = Mid$(fejr$, 38, 3)
    Mid$(pvszrec$, 95, 14) = ertszam(Str$(devizaertek@), 14, 2)
  End If
  Mid$(pvszrec$, 78, 14) = ertszam(Str$(forintertek@), 14, 2)
  Mid$(pvszrec$, 109, 1) = "N"
  Mid$(pvszrec$, 110, 3) = Mid$(fejr$, 51, 30)
  Mid$(pvszrec$, 160, 6) = maidatum$
  Mid$(pvszrec$, 173, 8) = ugyintezo$
  '--- áfa bontás
  For i9% = 1 To 5
    If afaalapok@(i9%) <> 0 Then
elem9$ = Space$(30)
Mid$(elem9$, 1, 2) = afakodok$(i9%)
Mid$(elem9$, 3, 14) = ertszam(Str$(afaalapok(i9%)), 14, 2)
Mid$(elem9$, 17, 14) = ertszam(Str$(afaosszegek@(i9%)), 14, 2)
Mid$(pvszrec$, (i9% - 1) * 30 + 250, 30) = elem9$
    End If
  Next
  '--- kontírozás
  kiegrec$ = ""
  For i9% = 1 To 50
    If ellosszegek@(i9%) <> 0 Then
If i9% > 10 Then
  If kiegrec$ = "" Then
    kiegrec$ = Space(3000)
    Mid$(kiegrec$, 1, 7) = Mid$(pvszrec$, 1, 7)
    Mid$(kiegrec$, 8, 2) = "UJ"
  End If
  elem9$ = Space$(53)
  Mid$(elem9$, 1, 14) = ertszam(Str$(ellosszegek@(i9%)), 14, 2)
  Mid$(elem9$, 15, 32) = ellszla$(i9%)
  Mid$(kiegrec$, (i9% - 11) * 53 + 10, 53) = elem9$
Else
  elem9$ = Space$(53)
  Mid$(elem9$, 1, 14) = ertszam(Str$(ellosszegek@(i9%)), 14, 2)
  Mid$(elem9$, 15, 32) = ellszla$(i9%)
  Mid$(pvszrec$, (i9% - 1) * 53 + 400, 53) = elem9$
End If
    End If
  Next
  '--- elõleg beszámítások beírása
  For i9% = 1 To 5
    If Mid$(Nyugel1.NtAtad(i9%), 8, 15) <> Space$(15) Then
Mid$(pvszrec$, (i9% - 1) * 43 + 1280, 43) = Nyugel1.NtAtad(i9%)
    End If
  Next
  '--- pvsz kiírás
  Call pszkonyvel(pvszrec$, kiegrec$, "V")
  vikt$ = Mid$(pvszrec$, 1, 7)
Return
sztornokezel:
  '--- számla sztornó könyvelése
  'Call pszsztrn(vikt$, "V", sztordat$, sztornoszamlaszam$)
  Call pszsztrn(vikt$, "V", sztordat$, szamlaszam$)
Return

szamlaszamol:
  '--- számla kiszámítása áfa és kontírozás tombok feltoltese
  '--- devizakonverzió
  '--- elõleg iktatók beírása nt$-ba
  '--- feltölteni afakodok$(5), afaalapok@(5), afaosszegek@(5),afaszamlak$(5)
  '---            ellszla$(10), ellosszegek@(10), devellosszegek@(10)
  '--- 080229 kerekítés
  'kpsszamla = 0
  'If seset% = 1 Then
  '  '--- belföldi számla
  '  If Mid$(Fejr$, 38, 3) = "   " Then
  '    fizm$ = Mid$(Fejr$, 36, 2)
  '    fizmrec$ = dbxkey("PFIZ", fizm$)
  '    If Mid$(fizmrec$, 33, 1) = "K" Then kpsszamla = 1
  '  End If
  'End If
  For i99% = 1 To 50
    If i99% < 6 Then
afakodok$(i99%) = "": afaalapok@(i99%) = 0: afaosszegek@(i99%) = 0: afaszamlak$(i99%) = ""
    End If
    ellszla$(i99%) = "": ellosszegek@(i99%) = 0: devellosszegek@(i99%) = 0
  Next
  devizaertek@ = 0
  forintertek@ = 0
  mennyker% = Val(Mid$(irec$, 343, 1))
  ertker% = Val(Mid$(irec$, 344, 1))
  afaker% = Val(Mid$(irec$, 345, 1))
  If mennyker% = 0 Then fstm$ = "############0" Else fstm$ = "#############0." + String(mennyker%, "0")
  If ertker% = 0 Then fste$ = "############0" Else fste$ = "#############0." + String(ertker%, "0")
  If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
  If arf@ <> 0 Then fste$ = "#############0.00": fst$ = "#############0.00"
  onert@ = 0: obert@ = 0
  For i13% = 1 To 200
    elem$ = Nyugel1.MSFlexGrid1.TextMatrix(i13%, 1)
    If Trim$(elem$) <> "" Then
tkod$ = Nyugel1.MSFlexGrid1.TextMatrix(i13%, 1)
termrec$ = dbxkey("KTRM", tkod$)
' ÁFA itt
afakod$ = Mid$(termrec$, 706, 2)
afrec$ = dbxkey("PAFA", afakod$)
afakulcs@ = xval(Mid$(afrec$, 33, 6))
afajel$ = Mid$(afrec$, 39, 1)
afszla$ = Mid$(afrec$, 49, 8)
menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 3))
liar@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 4))
pensz@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 5))
'      penft@ = xval(Mid$(elem$, 43, 12))
elar@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 4))
' kedvezmény kezelése
If pensz@ <> 0 Then
    elar@ = elar@ / 100 * (100 - pensz@)
End If


If menny@ <> 0 Then
  bert@ = elar * menny@
  If arf@ <> 0 Then bert@ = bert@ * arf@
  bert@ = xval(Format(bert@, fste$))
Else
  bert@ = elar@
  If arf@ <> 0 Then bert@ = bert@ * arf@
End If

nert@ = bert@ / (1 + afakulcs@ / 100)
nert@ = xval(Format(nert@, fste$))
afaosz@ = bert@ - nert@
afaosz@ = xval(Format(afaosz@, fst$))
    
If menny@ <> 0 Then
   elar@ = nert@ / menny@
   elar@ = xval(Format(elar@, fste$))
Else
   elar@ = nert@
End If

'      Ez volt a rossz
'      afaosz@ = (elar@ * afakulcs@) / 100
'      afaosz@ = xval(Format(afaosz@, fst$))
'      elar@ = elar@ - afaosz@


'      afaosz@ = (nert@ * afakulcs@) / 100
'      afaosz@ = xval(Format(afaosz@, fst$))
'      bert@ = nert@
'      nert@ = bert@ - afaosz@

For j99% = 1 To 5
  If afakodok$(j99%) = afakod$ Or afakodok$(j99%) = "" Then
    If afakodok$(j99%) = "" Then
      afakodok$(j99%) = afakod$
      afaalapok@(j99%) = nert@
      afaosszegek@(j99%) = afaosz@
      afaszamlak$(j99%) = afszla$
    Else
      afaalapok@(j99%) = afaalapok@(j99%) + nert@
      afaosszegek@(j99%) = afaosszegek@(j99%) + afaosz@
    End If
    Exit For
  End If
Next
krakrec$ = dbxkey("KRAK", krak$)
ellszi$ = Mid$(elem$, 69, 8) + Space$(8) + Mid$(elem$, 77, 16)
If Trim(ellszi$) = "" Then
  ellszi$ = Mid$(termrec$, 716, 8) + Space$(8) + Mid$(termrec$, 724, 16)
End If
If Trim(ellszi$) = "" Then
  ellszi$ = Mid$(krakrec$, 151, 8) + Space$(8) + Mid$(krakrec$, 159, 16)
End If
For j99% = 1 To 50
  If ellszla$(j99%) = ellszi$ Or ellszla$(j99%) = "" Then
    If ellszla$(j99%) = "" Then
      ellszla$(j99%) = ellszi$
      ellosszegek@(j99%) = nert@
      If arf@ <> 0 Then devellosszegek@(j99%) = nert@ / arf@
    Else
      ellosszegek@(j99%) = ellosszegek@(j99%) + nert@
      If arf@ <> 0 Then devellosszegek@(j99%) = devellosszegek@(j99%) + nert@ / arf@
    End If
    forintertek@ = forintertek@ + nert@
    If arf@ <> 0 Then devizaertek@ = devizaertek@ + nert@ / arf@
    Exit For
  End If
Next
    End If
  Next
  '--- afa kontirozas
  For i99 = 1 To 5
    If afakodok$(i99%) <> "" Then
ellszi$ = afaszamlak$(i99%) + Space$(24)
nert@ = afaosszegek@(i99%)
For j99% = 1 To 50
  If ellszla$(j99%) = ellszi$ Or ellszla$(j99%) = "" Then
    If ellszla$(j99%) = "" Then
      ellszla$(j99%) = ellszi$
      ellosszegek@(j99%) = nert@
      If arf@ <> 0 Then devellosszegek@(j99%) = nert@ / arf@
    Else
      ellosszegek@(j99%) = ellosszegek@(j99%) + nert@
      If arf@ <> 0 Then devellosszegek@(j99%) = devellosszegek@(j99%) + nert@ / arf@
    End If
    forintertek@ = forintertek@ + nert@
    If arf@ <> 0 Then devizaertek@ = devizaertek@ + nert@ / arf@
    Exit For
  End If
Next
    End If
  Next
  '--- 080229 kerekítés
  If kpsszamla = 1 Then
    Call kerekit510(forintertek@, kerekossz@, kerek@, "K")
    If kerek@ <> 0 Then
'--- bruttó összeg
forintertek@ = kerekossz@
'--- áfa nem adóalap
afakod$ = kerafakod
afrec$ = dbxkey("PAFA", afakod$)
afakulcs@ = xval(Mid$(afrec$, 33, 6))
afajel$ = Mid$(afrec$, 39, 1)
afszla$ = Mid$(afrec$, 49, 8)
For j99% = 1 To 5
  If afakodok$(j99%) = afakod$ Or afakodok$(j99%) = "" Then
    If afakodok$(j99%) = "" Then
      afakodok$(j99%) = afakod$
      afaalapok@(j99%) = kerek@
      afaosszegek@(j99%) = 0
      afaszamlak$(j99%) = afszla$
    Else
      afaalapok@(j99%) = afaalapok@(j99%) + kerek@
    End If
    Exit For
  End If
Next
'--- kerekítés kontírozása
If kerek@ > 0 Then
  ellszi$ = kerbev
Else
  ellszi$ = kerraf
End If
For j99% = 1 To 50
  If ellszla$(j99%) = ellszi$ Or ellszla$(j99%) = "" Then
    If ellszla$(j99%) = "" Then
      ellszla$(j99%) = ellszi$
      ellosszegek@(j99%) = kerek@
    Else
      ellosszegek@(j99%) = ellosszegek@(j99%) + kerek@
    End If
    Exit For
  End If
Next
    End If
  End If
Return

szamlair:
  '--- formátumvezérelt számla elkészítése
  '--- formatumfile beolvasása
  '--- 080229 kerekítés
  
  If nyugtavolt = 1 Then
    ' kp-s bizonylat
    fr4nev$ = "auw-pszu"
    peldany$ = "1"
  ElseIf nyugtavolt = 2 Then
    ' kp-s száma
    ' külön sorszámon fusson
    If Mid$(fejr$, 24, 6) > "151231" Then
       fr4nev$ = "auw-uszk"
    Else
       fr4nev$ = "auw-pszk"
    End If
    peldany$ = "2"
  ElseIf nyugtavolt = 3 Or nyugtavolt = 12 Then
    ' hiteles számla
    peldany$ = "2"
    If nyugtavolt = 12 Then
      fr4nev$ = "auw-psza"
      peldany$ = "2"
    Else
      If Mid$(fejr$, 24, 6) > "151231" Then
        fr4nev$ = "auw-uszb"
      Else
        fr4nev$ = "auw-pszb"
      End If
    End If
    nyugtavolt = 3
    
  ElseIf nyugtavolt = 7 Then
    ' garania jegy
    If Nyugel1.Check1.Value = 1 Then
fr4nev$ = "auw-slev"
    Else
fr4nev$ = "auw-pszg"
    End If
    peldany$ = "2"
  Else
    Return
  End If
  
  GoSub fr4beolv
  arf@ = xval(Mid$(fejr$, 41, 10))
  
  '--- belföldi számla
  For i9% = 1 To 50: fmezok$(i9%) = "": Next
'--- fejlec mezok feltoltése
  fmezok$(1) = Mid$(irec$, 5, 60)
  fmezok$(2) = Trim$(Mid$(irec$, 95, 8)) + " " + Trim$(Mid$(irec$, 103, 30)) + " " + Trim$(Mid$(irec$, 133, 30)) + " " + Trim$(Mid$(irec$, 163, 10))
  fmezok$(3) = Mid$(irec$, 173, 15)
  If bankvalasztas% = 0 Then
fmezok$(4) = Trim$(Mid$(irec$, 203, 30)) + " " + banktagol(Mid$(irec$, 233, 24))
fmezok$(5) = Mid$(irec$, 257, 30)
  Else
fkkod$ = Left(Bankval.MSFlexGrid1.TextMatrix(bankvalasztas% - 1, 0) + Space(8), 8)
fkkrec$ = dbxkey("FKSZ", fkkod$)
fmezok$(4) = Trim$(Mid$(fkkrec$, 389, 30)) + " " + banktagol(Mid$(fkkrec$, 419, 24))
fmezok$(5) = Mid$(fkkrec$, 443, 28)
  End If
 
    If Not Trim$(Nyugel1.Text4(1).Text) = "" Then

'      fmezok$(7) = Mid$(pec$, 16, 60)
'      fmezok$(8) = Trim$(Mid$(prec$, 106, 8)) + " " + Trim$(Mid$(prec$, 114, 30)) + " " + Trim$(Mid$(prec$, 144, 30)) + " " + Trim$(Mid$(prec$, 174, 10))
' fmezok$(9) = Mid$(prec$, 184, 15)
    End If
    fmezok$(9) = Nyugel1.Text12.Text
    fmezok$(6) = Mid$(fejr$, 1, 15)
    fmezok$(7) = Nyugel1.Text2.Text
    fmezok$(8) = Nyugel1.Text3.Text
    fmezok$(41) = ""
    fmezok$(42) = ""
    'If fr4nev$ = "auw-psza" Then
    Call szetvalaszt(fmezok$(7), fmezok$(42), 43, 1)
    fmezok$(8) = Trim$(fmezok$(8))
    If Len(fmezok$(8)) > 43 Then
       p% = InStr(fmezok$(8), ",")
       If p% > 0 Then
          hossz% = Len(fmezok$(8)) - p% - 1
          fmezok$(41) = Right$(fmezok$(8), hossz%)
          fmezok$(8) = Mid$(fmezok$(8), 1, p)
       End If
     End If
    'End If
    fmezok$(10) = szamlaszam$
    

    
    fmezok$(11) = datki(Mid$(fejr$, 24, 6))
    fmezok$(12) = datki(Mid$(fejr$, 18, 6))
    fmezok$(13) = datki(Mid$(fejr$, 30, 6))
    fizm$ = Nyugel1.Text7.Text
    fizmrec$ = dbxkey("PFIZ", fizm$)
    '--- 080229 kerekítés
    If fr4nev$ = "auw-pszk" Or fr4nev$ = "auw-pszu" Then
If Mid$(fizmrec$, 33, 1) = "K" Then kpsszamla = 1
    End If
    fmezok$(14) = Nyugel1.Text5.Text
    fmezok$(15) = Trim$(Mid$(fejr$, 111, 60))
    'fmezok$(15) = Nyugel1.Text8.Text
    devnem$ = Mid$(fejr$, 38, 3)
    If devnem$ <> "   " Then
fmezok$(16) = devnem$
fmezok$(17) = ertszam(Mid$(fejr$, 41, 10), 10, 4)
    End If
    If szoveg18$ = "" Then
       szoveg18$ = Mid$(fejr$, 171, 58)
    End If
    fmezok$(18) = szoveg18$
    'If sztornoszamla% = 1 And (seset% = 1 Or seset% = 3 Or seset% = 5) Then
    If sztornoszamla% = 1 Then
fmezok$(10) = sztornoszamlaszam$
If Mid$(fejr$, 24, 6) > "151231" Then
 fmezok$(26) = " Eredeti számla: <ereszam>"               ' szamlaszam$
Else
 fmezok$(26) = " Eredeti számla: " + szamlaszam$
End If
    End If
    If sztornomasolat% = 1 Then
fmezok$(10) = szamlaszam$
   If Mid$(fejr$, 24, 6) > "151231" Then
fmezok$(26) = " Szornó számla: " + "<szamlaszam>"             'sztornoszamlaszam$
   Else
fmezok$(26) = " Szornó számla: " + sztornoszamlaszam$
   End If
    End If
    
    lfi = FreeFile
    ' garancia jegy
    If fr4nev$ = "auw-slev" And form852% = 1 Then
Open listautvonal$ + terminal$ + task$ + "SLEV.lst" For Output As #lfi
    Else
Open listautvonal$ + terminal$ + task$ + "SZLA.lst" For Output As #lfi
    End If
    fr4kod$ = "F"
    GoSub fr4ir
    '--- sorok összeállítása
    mennyker% = xval(Mid$(irec$, 343, 1))
    ertker% = xval(Mid$(irec$, 344, 1))
    afaker% = xval(Mid$(irec$, 345, 1))
    If ertker% = 0 Then fste$ = "############0" Else fste$ = "#############0." + String(ertker%, "0")
    If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
    If arf@ <> 0 Then fste$ = "#############0.00": fst$ = "#############0.00"
    onert@ = 0: obert@ = 0: For i7% = 1 To 6: afatomb@(i7%, 1) = 0: afatomb@(i7%, 2) = 0: Next
    osuly@ = 0
    tetdb% = 0
    For i13% = 1 To 200
For i9% = 1 To 50: fmezok$(i9%) = "": Next
elem$ = Nyugel1.MSFlexGrid1.TextMatrix(i13%, 1)

If Trim$(elem$) <> "" Then
  tetdb% = tetdb% + 1
  tkod$ = Nyugel1.MSFlexGrid1.TextMatrix(i13%, 1)
  termrec$ = dbxkey("KTRM", tkod$)
  cikkszam$ = tkod$
  ' ÁFA kezelés itt - újjat ill. másolatot , sztornót megkülönböztetni
  'If rnyugtavolt = 10 Then
  '   afakod$ = Nyugel1.MSFlexGrid1.TextMatrix(i13%, 7)
  'Else
  '   afakod$ = Mid$(termrec$, 706, 2)
  'End If
  afakod$ = Mid$(termrec$, 706, 2)
  afrec$ = dbxkey("PAFA", afakod$)
  afakulcs@ = xval(Mid$(afrec$, 33, 6))
  afajel$ = Mid$(afrec$, 39, 1)
  
  menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 3))
  megys$ = Mid$(termrec$, 484, 6)
  liar@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 4))
  pensz@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 5))
  elar@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 4))
  ' kedvezmény kezelése
  If pensz@ <> 0 Then
    elar@ = elar@ / 100 * (100 - pensz@)
  End If
  
  If menny@ <> 0 Then
    bert@ = elar * menny@
    bert@ = xval(Format(bert@, fste$))
  Else
' Eszi - 2011.09.22
'    bert@ = elar@
    bert@ = 0
  End If
  ' Eszi a 0 bruttó értékû tételt nem írta ki.
  'If bert@ <> 0 Then
    ' oda - vissza jó legyen nettóár, nettó érték, bruttó érték
    nert@ = bert@ / (1 + afakulcs@ / 100)
    nert@ = xval(Format(nert@, fste$))
    afaosz@ = bert@ - nert@
    afaosz@ = xval(Format(afaosz@, fst$))
    
    If menny@ <> 0 Then
      elar@ = nert@ / menny@
      elar@ = xval(Format(elar@, fste$))
    Else
      elar@ = nert@
    End If
    
    obert@ = obert@ + bert@
    onert@ = onert@ + nert@
    If afajel$ = "N" Then
      afatomb@(1, 1) = afatomb@(1, 1) + nert@
    Else
      If afajel$ = "M" Then
        afatomb@(2, 1) = afatomb@(2, 1) + nert@
      Else
        For afai% = 1 To afakulcsokdb
          If afakulcs@ = afakulcsok(afai%) Then
            afatomb(afai% + 2, 1) = afatomb(afai% + 2, 1) + nert@
            afatomb(afai% + 2, 2) = afatomb(afai% + 2, 2) + afaosz@
            Exit For
          End If
        Next
        'If afakulcs@ = 5 Then afatomb@(3, 1) = afatomb@(3, 1) + nert@: afatomb@(3, 2) = afatomb@(3, 2) + afaosz@
        'If afakulcs@ = 15 Then afatomb@(4, 1) = afatomb@(4, 1) + nert@: afatomb@(4, 2) = afatomb@(4, 2) + afaosz@
        'If afakulcs@ = 25 Then afatomb@(5, 1) = afatomb@(5, 1) + nert@: afatomb@(5, 2) = afatomb@(5, 2) + afaosz@
      End If
    End If



    fmezok$(1) = cikkszam$
    fmezok$(2) = Mid$(termrec$, 444, 12)
    fmezok$(3) = Mid$(termrec$, 16, 55)
    
    If fr4nev$ = "auw-psza" Then
       fmezok$(2) = Mid$(termrec$, 16, 55)
       fmezok$(3) = Nyugel1.Text8.Text
    
    End If
    If Nyugel1.MSFlexGrid1.TextMatrix(i13%, 6) = "Elõleg besz." Then
       fmezok$(3) = Mid$(tetnevmt(i13%), 61, 60)
    Else
       fmezok$(4) = Mid$(termrec$, 196, 55)
    End If

    'fmezok$(4) = Mid$(termrec$, 196, 55)
    fmezok$(5) = ertszam(Str$(stornoelojel@(menny@)), 10, mennyker%)
    fmezok$(6) = megys$
    fmezok$(7) = ertszam(Str$(liar@), 12, ertker%)
    fmezok$(8) = ertszam(Str$(pensz@), 12, ertker%)
'          fmezok$(9) = ertszam(Str$(penft@), 12, ertker%)
    fmezok$(10) = ertszam(Str$(elar@), 12, ertker%)
    fmezok$(11) = ertszam(Str$(afakulcs@), 5, ertker%)
    fmezok$(12) = ertszam(Str$(stornoelojel@(afaosz@)), 12, ertker%)
    fmezok$(13) = ertszam(Str$(stornoelojel@(nert@)), 12, ertker%)
    fmezok$(14) = ertszam(Str$(stornoelojel@(bert@)), 12, ertker%)
    'fmezok$(15) = Mid$(termrec$, 76, 60)
    fmezok$(16) = Mid$(termrec$, 136, 60)
    fmezok$(17) = Mid$(termrec$, 256, 60)
    fmezok$(18) = Mid$(termrec$, 316, 60)
    fmezok$(19) = Mid$(termrec$, 376, 60)
    fmezok$(20) = Mid$(termrec$, 522, 20)
    egysuly@ = xval(Mid$(termrec$, 836, 12))
    netsuly@ = egysuly@ * menny@
    osuly@ = osuly@ + netsuly@
    fmezok$(21) = " " + Trim(ertszam(Str(netsuly@), 14, 2)) + " kg"
    fr4kod$ = "S"
    
    gyariszamok$ = Nyugel1.GyariszamAtad(i13%, gysz$(), hiv$())
    pzx% = InStr(gyariszamok$, ":")
    db% = Val(Mid$(gyariszamok$, 1, pzx% - 1))
    fmezok$(15) = ""
    For i14% = 1 To db%
      fmezok$(15) = fmezok$(15) + Trim(gysz$(i14%)) + ";"
    Next

    kozvszolg = " "
    If Mid$(termrec$, 716, 4) = "9133" Then
       kozvszolg = "K"
    End If
    
    wrttrec$ = Space(300)
    Mid$(wrttrec$, 1, 10) = szamlaszam$                      ' Csak ha készpénzes
    Mid$(wrttrec$, 11, 3) = Right(Space(3) + Str(tetdb%), 3)
    Mid$(wrttrec$, 14, 1) = "E"
    Mid$(wrttrec$, 15, 6) = cikkszam$
    Mid$(wrttrec$, 21, 60) = fmezok$(3)                  ' Megnevezés
    Mid$(wrttrec$, 199, 12) = Mid$(termrec$, 444, 12)    ' Stat. szám
    Mid$(wrttrec$, 81, 12) = "  " + fmezok$(5)           ' Mennyiség
    Mid$(wrttrec$, 93, 6) = megys$                       ' Mennyiségi egység
    Mid$(wrttrec$, 299, 1) = kozvszolg                    ' Közvetített szolgáltatás - termékcsop=98
    Mid$(wrttrec$, 220, 14) = "  " + fmezok$(13) ' nettó érték
    Mid$(wrttrec$, 129, 12) = fmezok$(10) ' nettóár
    Mid$(wrttrec$, 141, 2) = afakod$
    Mid$(wrttrec$, 234, 10) = "    " + Mid$(afrec$, 33, 6)
    Mid$(wrttrec$, 244, 14) = ertszam(Str$(stornoelojel@(afaosz@)), 14, 2) ' ÁFA összeg
    Mid$(wrttrec$, 268, 14) = ertszam(Str$(stornoelojel@(bert@)), 14, 2)
    
    wrtt$(tetdb%) = wrttrec$

  
    GoSub fr4ir
  'End If
End If
    Next
    '--- lablec összeállítása
    For i9% = 1 To 50: fmezok$(i9%) = "": Next
    If masolat% = 1 Or (masolat% = 0 And rnyugtavolt = 10) Then
oeloleg@ = 0
For i11% = 1 To 4
  elolegkulcs@ = 0: elolegkod$ = "": elolalap@ = 0: elolafa@ = 0
  erelszikt$ = Mid$(nt$(i11%), 1, 7)
  elsz$ = Mid$(nt$(i11%), 8, 15)
  eloo@ = xval(Mid$(nt$(i11%), 23, 14))
  If eloo@ <> 0 And Trim(erelszikt$) <> "" Then
    elolegkod$ = ntafa$(i11%)
    If elolegkod$ <> "" Then
      elpafrec$ = dbxkey("PAFA", elolegkod$)
      If elpafrec$ <> "" Then elolegkulcs@ = xval(Mid$(elpafrec$, 33, 6))
    End If
    If elolegkulcs@ <> 0 Then elolafa@ = (eloo@ * elolegkulcs@) / (100 + elolegkulcs@) Else elolafa@ = 0
    elolafa@ = xval(Format(elolafa@, fst$))
    elolalap@ = eloo@ - elolafa@
    If fr4nev$ = "auw-psza" Then
      fmezok$(41 + i11%) = elsz$ + "számú bizonylat alapján"
      fmezok$(15 + i11%) = "Alap:" + ertszam(Str$(elolalap@), 12, 2) + " " + ertszam(Str(elolegkulcs@), 6, 2) + " % ÁFA" + ertszam(Str$(elolafa@), 12, 2)
    Else
       fmezok$(15 + i11%) = elsz$ + "Alap:" + ertszam(Str$(elolalap@), 12, 2) + " " + ertszam(Str(elolegkulcs@), 6, 2) + " % ÁFA" + ertszam(Str$(elolafa@), 12, 2)
    End If
    oeloleg@ = oeloleg@ + eloo@
  End If
Next
oeloleg@ = oeloleg@
    Else
    
oeloleg@ = 0
For i11% = 1 To 4
  elem$ = Nyugel1.NtAtad(i11%)
  elolegkulcs@ = 0: elolegkod$ = "": elolalap@ = 0: elolafa@ = 0
  erelszikt$ = Mid$(elem$, 1, 7)
  elsz$ = Mid$(elem$, 8, 15)
  eloo@ = xval(Mid$(elem$, 23, 14))
  If eloo@ <> 0 And Trim(erelszikt$) <> "" Then
    elolegkod$ = elolegafa(Mid$(fejr$, 1, 15), Mid$(elem$, 8, 10))
    If elolegkod$ <> "" Then
      elpafrec$ = dbxkey("PAFA", elolegkod$)
      If elpafrec$ <> "" Then elolegkulcs@ = xval(Mid$(elpafrec$, 33, 6))
    End If
    If elolegkulcs@ <> 0 Then elolafa@ = (eloo@ * elolegkulcs@) / (100 + elolegkulcs@) Else elolafa@ = 0
    elolafa@ = xval(Format(elolafa@, fst$))
    elolalap@ = eloo@ - elolafa@
    If fr4nev$ = "auw-psza" Then
      fmezok$(41 + i11%) = elsz$ + "számú bizonylat alapján"
      fmezok$(15 + i11%) = "Alap:" + ertszam(Str$(elolalap@), 12, 2) + " " + ertszam(Str(elolegkulcs@), 6, 2) + " % ÁFA" + ertszam(Str$(elolafa@), 12, 2)
    Else
      fmezok$(15 + i11%) = elsz$ + "Alap:" + ertszam(Str$(elolalap@), 12, 2) + " " + ertszam(Str(elolegkulcs@), 6, 2) + " % ÁFA" + ertszam(Str$(elolafa@), 12, 2)
    End If
    oeloleg@ = oeloleg@ + eloo@
    
  End If
Next
    oeloleg@ = stornoelojel@(oeloleg@)
    ' oeloleg@ = 0
    ' For i11% = 1 To 4
    '   elsz$ = Mid$(Nyugel1.NtAtad(i11%), 8, 15)
    '   eloo@ = xval(Mid$(Nyugel1.NtAtad(i11%), 23, 14))
    '   If eloo@ <> 0 Then
    '     fmezok$(15 + i11%) = elsz$ + " " + ertszam(Str$(eloo@), 14, 2)
    '     oeloleg@ = oeloleg@ + eloo@
    '   End If
    ' Next
    End If
    If devnem$ = "   " Then devnem$ = ""
    fmezok$(1) = ertszam(Str$(stornoelojel@(onert@)), 14, afaker%) + " " + devnem$
    fmezok$(2) = ertszam(Str$(stornoelojel@(afatomb@(1, 1))), 14, afaker%) + " " + devnem$
    fmezok$(3) = ertszam(Str$(stornoelojel@(afatomb@(2, 1))), 14, afaker%) + " " + devnem$
    fmezok$(4) = ertszam(Str$(stornoelojel@(afatomb@(3, 1))), 14, afaker%) + " " + devnem$
    fmezok$(5) = ertszam(Str$(stornoelojel@(afatomb@(3, 2))), 14, afaker%) + " " + devnem$
    fmezok$(6) = ertszam(Str$(stornoelojel@(afatomb@(4, 1))), 14, afaker%) + " " + devnem$
    fmezok$(7) = ertszam(Str$(stornoelojel@(afatomb@(4, 2))), 14, afaker%) + " " + devnem$
    fmezok$(8) = ertszam(Str$(stornoelojel@(afatomb@(5, 1))), 14, afaker%) + " " + devnem$
    fmezok$(9) = ertszam(Str$(stornoelojel@(afatomb@(5, 2))), 14, afaker%) + " " + devnem$
    fmezok$(34) = ertszam(Str$(stornoelojel@(afatomb@(6, 1))), 14, afaker%) + " " + devnem$
    fmezok$(35) = ertszam(Str$(stornoelojel@(afatomb@(6, 2))), 14, afaker%) + " " + devnem$
    '--- 080229 kerekítés
    fizetnem@ = obert@ - oeloleg@
    If kpsszamla = 1 Then
'--- készpénzes
Call kerekit510(fizetnem@, fizetni@, kerek@, "K")
fmezok$(36) = ertszam(Str$(kerek@), 14, 0) + " " + devnem$
    Else
fizetni@ = fizetnem@: kerek@ = 0
    End If
    fmezok$(10) = ertszam(Str$(stornoelojel@(obert@)), 14, ertker%) + " " + devnem$
    fmezok$(11) = ertszam(Str$(stornoelojel@(oeloleg@)), 14, ertker%) + " " + devnem$
    If fr4nev$ = "auw-psza" Then
       fmezok$(11) = ertszam(Str$(stornoelojel@(-oeloleg@)), 14, ertker%) + " " + devnem$
    End If
    '--- 080229 kerekítés
    fmezok$(39) = ertszam(Str(stornoelojel@(kerek@)), 14, ertker%)
    fmezok$(12) = ertszam(Str$(stornoelojel@(fizetni@)), 14, ertker%) + " " + devnem$
    fmezok$(13) = Trim$(Mid$(partrec$, 363, 60)) + " " + Trim$(Mid$(partrec$, 423, 60))
    fmezok$(40) = betuvel$(stornoelojel@(fizetni@), devnem$)
    fmezok$(41) = Nyugel1.Text8.Text
    fmezok$(46) = " számú bizonylat alapján"
    If devnem$ <> "" Then
arf@ = xval(Mid$(fejr$, 41, 10))
arfi# = arf@
fmezok$(20) = ARfkonv(fmezok$(2), arfi#, 14)
fmezok$(21) = ARfkonv(fmezok$(3), arfi#, 14)
fmezok$(22) = ARfkonv(fmezok$(4), arfi#, 14)
fmezok$(23) = ARfkonv(fmezok$(5), arfi#, 14)
fmezok$(24) = ARfkonv(fmezok$(6), arfi#, 14)
fmezok$(25) = ARfkonv(fmezok$(7), arfi#, 14)
fmezok$(26) = ARfkonv(fmezok$(8), arfi#, 14)
fmezok$(27) = ARfkonv(fmezok$(9), arfi#, 14)
fmezok$(36) = ARfkonv(fmezok$(34), arfi#, 14)
fmezok$(37) = ARfkonv(fmezok$(35), arfi#, 14)
    End If
    If fr4nev$ = "auw-sleb" Then
fmezok$(28) = langprg(78) + Trim(bovslev.Text1.Text)
fmezok$(29) = langprg(79) + ertszam(Str(osuly@), 10, 2) + " kg"
fmezok$(30) = Trim(bovslev.Text2.Text)
fmezok$(31) = Trim(bovslev.Text3.Text)
fmezok$(32) = Trim(bovslev.Text4.Text)
fmezok$(33) = Trim(bovslev.Text5.Text)
    End If
    fr4kod$ = "L"
    GoSub fr4ir
    '--- megjelenítés
    Close lfi
 ' Call mess("Nyugtavolt:" + Str(nyugtavolt), 1, 0, langmodul(157), valasz%)
  
  'If (masolat% = 1 Or sztornomasolat% = 1) And Not melyik = 3 Then
  If (masolat% = 1 Or sztornomasolat% = 1) Then
    
    If Mid$(fejr$, 24, 6) > "151231" And (nyugtavolt = 2 Or nyugtavolt = 3) Then
'       sztornoszamlaszam$ = Mid$(pszbrec$, 282, 15)
'       If Not RTrim$(sztornoszamlaszam$) = "" Then
'          Call mess("Az eredeti számlát kéri?", 5, 0, "Választás", valasz%)
'          If valasz% = 1 Then
'            ssz$ = szamlaszam$
'          Else
'            ssz$ = sztornoszamlaszam$
'          End If
'       End If
       Call szlaparair(szlamod$, ssz$, szlacim$, szlafaj$)
       If melyik% = 3 Then
         Shell programutvonal$ + "AUW-QSZLAM " + terminal$ + task$ + "/" + listautvonal$, vbNormalFocus
       Else
         Shell programutvonal$ + "AUW-SZLA " + terminal$ + task$ + "/" + listautvonal$, vbNormalFocus
       End If
    Else
       Shell programutvonal$ + "dbx4-sho " + terminal$ + task$ + "SZLA/" + listautvonal$, vbNormalFocus
    End If
  Else
     If nyugtavolt = 2 Or nyugtavolt = 3 Or nyugtavolt = 12 Then
        xszamlaszam$ = szamlaszam$
        'If nyugtavolt = 2 And sztornoszamla% = 1 Then
        '   xszamlaszam$ = sztornoszamlaszam$
        'End If
        Call szamlakeszit("dbx4-qsho", fejr$, irec$, partrec$, fizmrec$, onert@, obert@, szlacim$, szlamod$, ssz$, xszamlaszam$, tetdb%, Nyugel1.Text13.Text, fr4nev$, nyugtavolt, szlafaj$, rpartner$)
     Else
       Shell programutvonal$ + "dbx4-qsho " + peldany$ + terminal$ + task$ + "SZLA/" + listautvonal$, vbNormalFocus
     End If
  End If
  If munkalapbol Then
   munkalapbol = False
   Shell "C:\Program Files\Win Investor RT\Win Munkalap és Garancia Nyilvántartás\szlavisz.exe " + munkalap$ + " " + szamlaszam$
  End If
  If sztornoszamla% = 1 And Not Trim$(munkalap$) = "" Then
    Shell "C:\Program Files\Win Investor RT\Win Munkalap és Garancia Nyilvántartás\szlavisz.exe " + munkalap$ + "           "
  End If
  munkalap$ = ""
  Call gomb(" Tovább &", gg%, 8320, 100, "V")
Return

fr4beolv:
  If langutvonal$ = "" Then
    ffi = FreeFile
    Open programutvonal$ + fr4nev$ + ".fx4" For Binary Shared As #ffi
    formfm& = LOF(ffi)
    Close ffi
    ffi = FreeFile
    If formfm& > 0 Then
Open programutvonal$ + fr4nev$ + ".fx4" For Input Shared As #ffi
    Else
Open programutvonal$ + fr4nev$ + ".fr4" For Input Shared As #ffi
    End If
  Else
    ffi = FreeFile
    Open langutvonal$ + fr4nev$ + ".fx4" For Binary Shared As #ffi
    formfm& = LOF(ffi)
    Close ffi
    ffi = FreeFile
    If formfm& > 0 Then
Open langutvonal$ + fr4nev$ + ".fx4" For Input Shared As #ffi
    Else
Open langutvonal$ + fr4nev$ + ".fr4" For Input Shared As #ffi
    End If
  End If
  heddb% = 0: fejdb% = 0: sordb% = 0: labdb% = 0: tradb% = 0
  Do
    Line Input #ffi, fs$
    ko$ = Left$(fs$, 2)
    If ko$ <> "* " Then
Select Case ko$
  Case "F="
    fejdb% = fejdb% + 1
    fr4fej$(fejdb%) = fs$
  Case "S="
    sordb% = sordb% + 1
    fr4sor$(sordb%) = fs$
  Case "L="
    labdb% = labdb% + 1
    fr4lab$(labdb%) = fs$
  Case Else
    If fejdb% = 0 Then
      heddb% = heddb% + 1
      fr4hed$(heddb%) = fs$
    Else
      If labdb% > 0 Then
        tradb% = tradb% + 1
        fr4tra$(tradb%) = fs$
      End If
    End If
End Select
    End If
  Loop While Not EOF(ffi)
  Close ffi
Return

fr4ir:
  '--- listafile írása
  If fr4kod$ = "F" Then
    For i14% = 1 To heddb%
sr$ = fr4hed$(i14%)
Print #lfi, sr$
    Next
    If masolat% = 1 Then
sr$ = "F=Ariel/10/B/K &" + "Másolat" + "            &"
Print #lfi, sr$
    End If
    If sztornoszamla% = 1 Or sztornomasolatX% = 1 Then
sr$ = "F=Ariel/10/B/K &" + "Érvénytelenítõ" + "            &"
Print #lfi, sr$
    End If
    If sztornomasolat% = 1 Then
sr$ = "F=Ariel/10/B/K &" + "Sztornó" + " " + "másolat" + "            &"
Print #lfi, sr$
    End If
    ' Mindíg ugyanannyi sor legyen
    If masolat% = 0 And sztornoszamla% = 0 And sztornomasolat% = 0 And sztornomasolatX% = 0 Then
sr$ = "F=Ariel/10/B/K &                 &"
Print #lfi, sr$
    End If

    For i15% = 1 To fejdb%
sr$ = fr4fej$(i15%)
pzx% = InStr(sr$, "#")
Do While pzx% > 0
  sosz% = Val(Mid$(sr$, pzx% + 1, 2))
  mzx$ = fmezok$(sosz%)
  If Len(mzx$) < 3 Then mzx$ = mzx$ + "   "
  pzz% = InStr(pzx% + 1, sr$, "#")
  pzzz% = InStr(pzx% + 1, sr$, "[")
  If pzz% <> 0 And pzzz% <> 0 Then
    If pzzz% < pzz% Then pzz% = pzzz%
  Else
    If pzz% = 0 And pzzz% <> 0 Then pzz% = pzzz%
  End If
  If pzz% > 0 Then
    ureshely% = pzz% - pzx% - 1
    If Len(mzx$) > ureshely% Then mzx$ = Left(mzx$, ureshely%)
  End If
  Mid$(sr$, pzx%) = mzx$
  pzx% = InStr(sr$, "#")
Loop
Print #lfi, sr$
    Next
    Return
  End If
  If fr4kod$ = "S" Then
    For i15% = 1 To sordb%
sr$ = fr4sor$(i15%)
pzx% = InStr(sr$, "#")
Do While pzx% > 0
  sosz% = Val(Mid$(sr$, pzx% + 1, 2))
  mzx$ = fmezok$(sosz%)
  If Len(mzx$) < 3 Then mzx$ = mzx$ + "   "
  pzz% = InStr(pzx% + 1, sr$, "#")
  pzzz% = InStr(pzx% + 1, sr$, "[")
  If pzz% <> 0 And pzzz% <> 0 Then
    If pzzz% < pzz% Then pzz% = pzzz%
  Else
    If pzz% = 0 And pzzz% <> 0 Then pzz% = pzzz%
  End If
  If pzz% > 0 Then
    ureshely% = pzz% - pzx% - 1
    If Len(mzx$) > ureshely% Then mzx$ = Left(mzx$, ureshely%)
  End If
  Mid$(sr$, pzx%) = mzx$
  pzx% = InStr(sr$, "#")
Loop
Print #lfi, sr$
    Next
    Return
  End If
  If fr4kod$ = "L" Then
    For i15% = 1 To labdb%
sr$ = fr4lab$(i15%)
pzx% = InStr(sr$, "#")
Do While pzx% > 0
  sosz% = Val(Mid$(sr$, pzx% + 1, 2))
  mzx$ = fmezok$(sosz%)
  If Len(mzx$) < 3 Then mzx$ = mzx$ + "   "
  pzz% = InStr(pzx% + 1, sr$, "#")
  pzzz% = InStr(pzx% + 1, sr$, "[")
  If pzz% <> 0 And pzzz% <> 0 Then
    If pzzz% < pzz% Then pzz% = pzzz%
  Else
    If pzz% = 0 And pzzz% <> 0 Then pzz% = pzzz%
  End If
  If pzz% > 0 Then
    ureshely% = pzz% - pzx% - 1
    If Len(mzx$) > ureshely% Then mzx$ = Left(mzx$, ureshely%)
  End If
  Mid$(sr$, pzx%) = mzx$
  pzx% = InStr(sr$, "#")
Loop
Print #lfi, sr$
    Next
    For i14% = 1 To tradb%
sr$ = fr4tra(i14%)
Print #lfi, sr$
    Next
  End If
Return


' Eszi - nem formátum vezérelt
Exit Sub
vege:
  hbuz$ = "Hiba: AUW-QRPTN (Command3.Click)" + " Sor:" + Str$(Erl) + " " + Str$(Err) + Err.Description
  Call mess(hbuz$, 1, 0, "Hiba", valasz%)
  Resume Next
 
 Call hibakiir(1, hbuz$)
 Resume Next
End Sub
Private Sub hibakiir(hiv%, hbuz$)
  hbdatum$ = "* " + Right(Date$, 2) + "." + Left(Date$, 2) + "." + Mid$(Date$, 4, 2) + " " + Str(Time) + " " + terminal$ + task$ + " " + ugyintezo$
  
  
  fhb = FreeFile
  Open auditorutvonal$ + "hiba.txt" For Append As #fhb
  Print #fhb, hbdatum$
  Print #fhb, hbuz$
  
  If hiv% = 1 Then
    hbegyb$ = "Nyugta volt:" + Str(nyugtavolt)
    Print #fhb, hbegyb$
    For i1% = 1 To 200
      hbegyb$ = Left(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) + Space(15), 15) + Nyugel1.MSFlexGrid1.TextMatrix(i1%, 2)
      If Trim(hbegyb$) <> "" Then
         Print #fhb, hbegyb$
      End If
    Next
  End If
  Close fhb
End Sub
Private Sub feltolt2(ksybrec$, pkod$)
        If Not ksybrec$ = "" Then
           wpkod = Mid$(ksybrec$, 8, 15)
           If Mid$(wpkod, 1, 8) = Space(8) Then
              wpkod2 = Mid$(wpkod, 9, 7)
              If Len(RTrim$(wpkod2)) = 7 Then
                 Nyugel1.Text13 = wpkod2
              End If
              
           Else
             If Mid$(wpkod, 1, 4) = "BANK" Then
                Nyugel1.Text4(1) = "BANK"
                wpkod2 = Mid$(wpkod, 9, 7)
                If Len(RTrim$(wpkod2)) = 7 Then
                   Nyugel1.Text13 = wpkod2
                End If

             Else
                Nyugel1.Text4(1) = Mid$(ksybrec$, 8, 15)
             End If
           End If
           
           Nyugel1.Text2 = Mid$(ksybrec$, 23, 60)
           Nyugel1.Text3 = Mid$(ksybrec$, 83, 60)
           Nyugel1.Text11 = Mid$(ksybrec$, 143, 58)
        End If
        If Trim(Nyugel1.Text4(1)) = "" Then
              Nyugel1.Text4(1) = pkod$
        End If
     
End Sub

Private Sub feltolt()


        Nyugel1.Text2 = Mid$(partrec$, 16, 60)
        Nyugel1.Text3 = postacim(partrec$, 106)
        Nyugel1.Text12.Text = Trim(Mid$(partrec$, 184, 15))
        If Not Mid$(Nyugel1.Text12.Text, 9, 1) = "-" And Len(Nyugel1.Text12.Text) = 11 Then
             Nyugel1.Text12.Text = Mid$(Nyugel1.Text12.Text, 1, 8) + "-" + Mid$(Nyugel1.Text12.Text, 9, 1) + "-" + Mid$(Nyugel1.Text12.Text, 10, 2)
        End If
        poz = InStr(Nyugel1.Text11.Text, "Bev.")
        If poz > 0 Then
           Nyugel1.Text11.Text = Mid$(Nyugel1.Text11.Text, 1, poz - 1)
        End If
End Sub
Private Function stornoelojel@(osszeg@)
  
  If sztornoszamla% = 1 Or sztornomasolat% = 1 Then
    stornoelojel@ = -osszeg@
  Else
    stornoelojel@ = osszeg@
  End If
End Function
Private Function elolegafa$(pkod$, elolegszla$)
                    partrec$ = dbxkey("PART", pkod$)
                    nxptr& = Val(Mid$(partrec$, 702, 10))
                    elolegdb% = 0
                    dxfi = FreeFile
                    Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #dxfi
                    fim& = LOF(dxfi)
                    Do While nxptr& > 0
                      Seek #dxfi, nxptr& + 9
                      elrec$ = Space(650): Get #dxfi, , elrec$
                      nxptr& = Val(Mid$(elrec$, 204, 10))
                      elokod$ = "BV"
                      If Mid$(elrec$, 90, 1) <> "S" And Mid$(elrec$, 224, 2) = elokod$ Then
                        If Mid$(elrec$, 8, 10) = elolegszla$ Then
                          elem1$ = Mid$(elrec$, 8, 15) + " " + Mid$(elrec$, 38, 6) + " " + Right$(Space$(14) + Format(szabossz@, "##########0.00"), 14) + " " + Mid(elrec$, 1, 7)
                          For jij% = 1 To 5:
                            jelem$ = Mid$(elrec$, (jij% - 1) * 30 + 230, 30)
                            If Trim(jelem$) <> "" Then
                              elolegafa$ = Mid$(jelem$, 1, 2)
                              Exit Function
                            End If
                          Next
                        End If
                      End If
                    Loop
                    Close dxfi

End Function


Private Sub Form_Activate()
 If stax% = 0 Then
    stax% = 1
    'Call Command3_Click
    ReDim nevcimt$(30000)
    ReDim ncazonosito$(30000)
    irec$ = dbxkey("INST", "INST")
    Call afakulcstolt(irec$)
    Text5.Text = Mid$(irec$, 634, 4)
    Command1.Visible = True
    Command2.Visible = False
    Command3.Caption = "Új vevõ"
    Command3.Visible = True
    Command4.Visible = False
    Call hetinap(maidatum$, maisorsz%, napnev$)
    Label4.Top = 100
    Label4.Left = 8000
    Label4.Visible = True
    Label4.AutoSize = True
    Label4.Font.Name = "Microsoft Sans Serif"
    Label4.Font.Size = 11: Label4.Font.Bold = True
    Label4.Caption = " A mai nap:" + datki(maidatum) + " " + napnev$ + " "
    Label5.Visible = True: Text5.Visible = True
    ncdb = 0
    dbfi = FreeFile
    Open auditorutvonal$ + "auwker2.dbx" For Binary Shared As #dbfi
    ndfi = FreeFile
    Open auditorutvonal$ + "AUW-KPAR.ndx" For Binary Shared As #ndfi
    rc& = Int(LOF(ndfi) / 12)
    For i1% = 1 To rc&
         Get #ndfi, (i1% - 1) * 12& + 1, rcim&
         Seek #dbfi, rcim& + 9
         kparrec$ = Space(200): Get #dbfi, , kparrec$
         Call nevcimtolt(Mid$(kparrec$, 1, 120), Mid$(kparrec$, 190, 7))
    Next
    Close dbfi, ndfi

  End If
  
  
  
End Sub

Private Function betuvel$(fizetni@, devnem$)
  filler$ = Mid$(ertszam(Str$(fizetni@), 14, 2), 13, 2)
  If devnem$ = "" Or devnem$ = "HUF" Then
       devmas$ = "Ft"
  Else
       devmas$ = devnem$
  End If
  betuvel$ = szamszoveg(fizetni@, 0, "") + " " + filler$ + "/100 " + devmas$

End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    KeyCode = 0
    Call Command3_Click
  End If
  If KeyCode = vbKeyEscape Then
    KeyCode = 0:
    Call Command1_Click
  End If
End Sub

Private Sub Form_Load()
  '--- fõprogram kód
  '--- közös rész
  parancssor% = 0
  stax% = 0
  exparvolt% = 1
  expar$ = Trim$(UCase$(Command$))
'  If expar$ = "" Then expar$ = "AUWKER/RKSZ/A011/C:\AUWIN\": parancssor% = 0: exparvolt% = 0: jogok$ = "IIIIISSSSSSIIII"
  If expar$ = "" Then
    parancssor% = 1
    exparvolt% = 0
    dbxneve$ = "auwker"
    objneve$ = "RKRE"
    terminal$ = "A01"
    task$ = "1"
    auditorutvonal$ = "C:\auwin\proba\"
'    auditorutvonal$ = "Q:\AUWIN\"
    ugyintezo$ = "ESZES"
  Else
    paramdb% = 10
    Call linpar(expar$, param$(), "/", paramdb%)
    dbxneve$ = param$(1)
    objneve$ = param$(2)
    terminal$ = Mid$(param$(3), 1, 3)
    task$ = Mid$(param$(3), 4, 1)
    auditorutvonal$ = param$(4)
  End If
  Call prglang("dbx4-bws")
  programnev$ = "AUW-RPTN"
  If parancssor% = 0 Then Call auwini(inihiba%) Else programutvonal$ = "C:\AUWIN\": listautvonal$ = "C:\AUWIN\"
  If inihiba% <> 0 Then Close: End
  form1.Picture = LoadPicture(programutvonal$ + "auwht01.jpg")
  form1.Caption = task$ + "-Bolti értékesítés pénztárgépes kapcsolat nélkül" + " (" + regszam$ + "/" + terminal$ + "," + Trim(ugyintezo$) + "," + Trim(auditorutvonal$) + ")-" + Trim(cegneve$)
  igen123% = ini123%()
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWSZAMV", "PART", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER", "KTRM", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  
  Call dbxopen("AUWKER", "KKBZ", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER", "KKFT", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER", "KSZB", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER", "KSZT", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER2", "KSZX", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER2", "KCIM", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
 
  Call dbxopen("AUWKER2", "KCIM", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  
  
  Call dbxopen("AUWKER", "RKSZ", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER", "ARMG", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER", "KRAK", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER", "REAN", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER", "KMEG", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  
  Call dbxopen("AUWKER2", "KKFX", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER2", "KSZX", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWKER2", "KSYB", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  ' Eszi - kivenni
  Call dbxopen("AUWKER2", "KPAR", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  
  Call dbxopen("AUWKER2", "KXCB", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  
  Call dbxopen("AUWKER2", "KCIM", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
    
  Call dbxopen("AUWKER2", "SINS", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  
  Call dbxopen("AUWSZAMV", "PFIZ", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWSZAMV", "PVSZ", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWSZAMV", "PVSK", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWSZAMV", "PSHL", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWSZAMV", "PELO", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  Call dbxopen("AUWSZAMV", "PSEL", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  
  Call dbxopen("AUWSZAMV", "FKTE", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  
  Call dbxopen("AUWSZAMV", "PKTE", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  
  Call dbxopen("AUWSZAMV", "PELV", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  
  Call dbxopen("AUWSZAMV", "INST", 0, runtimhiba%)
  If runtimhiba% = 1 Then Call Command1_Click
  obs% = obsorszama(objneve$)
 
  irec$ = dbxkey("INST", "INST")
  kerafakod = Mid$(irec$, 468, 2)
  kerbev = Mid$(irec$, 570, 8)
  kerraf = Mid$(irec$, 578, 8)
 
  elolegprefix$ = Mid$(irec$, 450, 4)
 
  'Label1.Caption = Trim$(OBJTAB(obs%).obnev) + " " + langprg(1)
  Label1.Caption = "Bolti értékesítés pénztárgépes kapcsolat nélkül"
  maidatum$ = Right(Date$, 2) + Left(Date$, 2) + Mid$(Date$, 4, 2)
  Text8.Text = maidatum$
  teljesossz@ = 80000

  
  'Label1.Caption = expar$
End Sub
Private Sub elsolap(objazon, rek$)
  MSFlexGrid1.Clear
  MSFlexGrid1.Cols = 2
  MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  MSFlexGrid1.Font.Size = 8
  w1& = obsorszama(objazon)
  w2& = OBJTAB(w1&).oba(1)
  odarab& = ABLTAB(w2&).adatsorsz(0)
  MSFlexGrid1.Rows = odarab& + 1
  MSFlexGrid1.TextMatrix(0, 0) = langprg(44)
  MSFlexGrid1.TextMatrix(0, 1) = langprg(44) + " " + langprg(45)
  MSFlexGrid1.ColAlignment(1) = 1
  mmax% = 12: hxmax% = 0
  mhmax% = 10
  For i1& = 1 To odarab&
    ne$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatnev)
    ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).attr)
    mh% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatho
    kp% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatkp
    MSFlexGrid1.TextMatrix(i1&, 0) = ne$
    w3& = Len(ne$)
    hxh% = Vektor.TextWidth(ne$) + 100
    If hxmax% < hxh% Then hxmax% = hxh%
    MSFlexGrid1.ColWidth(0) = hxmax%
    'MSFlexGrid1.ColWidth(0) = h%
    If mh% > mhmax% Then h1% = mh% * 120: mhmax% = mh% Else h1% = mhmax% * 120
    MSFlexGrid1.ColWidth(1) = h1%
    MSFlexGrid1.Width = hxmax% + mhmax% * 120 + 70
    amezo$ = Mid$(rek$, kp%, mh%)
    MSFlexGrid1.TextMatrix(i1&, 1) = amezo$
  Next
  MSFlexGrid1.Height = (odarab&) * 240
  'Text2.Text = Trim$(OBJTAB(w2&).obnev)
  Text2.Top = 500
  MSFlexGrid1.Top = 512
  Text2.Text = ABLTAB(w2&).fejlec
  Text2.Width = MSFlexGrid1.Width
  Text2.Height = 312
  'Text2.Visible = True
  MSFlexGrid1.Visible = True
End Sub
Private Sub elsotab(objazon$, mt$(), sordarab%)
  MSFlexGrid2.Clear
  If sordarab% > 20 Then
    darab% = 20
    MSFlexGrid2.Rows = 21
  Else
    darab% = sordarab%
    MSFlexGrid2.Rows = sordarab% + 1
  End If
  MSFlexGrid2.Rows = 21
  w1& = obsorszama(objazon)
  w2& = OBJTAB(w1&).oba(1)
  odarab& = ABLTAB(w2&).adatsorsz(0)
  MSFlexGrid2.Cols = odarab& + 1
  For i1& = 1 To darab%
    MSFlexGrid2.TextMatrix(i1&, 0) = Str$(i1&) + "." + langprg(97)
  Next
  gw& = 0
  '--- adatok feltöltése
  For i1& = 1 To darab%
    For i2& = 1 To odarab&
      mh% = ADATAB(ABLTAB(w2&).adatsorsz(i2&)).adatho
      kp% = ADATAB(ABLTAB(w2&).adatsorsz(i2&)).adatkp
      amezo$ = Mid$(mt$(i1&), kp%, mh%)
      MSFlexGrid2.TextMatrix(i1&, i2&) = Trim$(amezo$)
    Next
  Next
  '--- fejlec
  For i1& = 1 To odarab&
    ne$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatnev)
    ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).attr)
    mh% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatho
    kp% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatkp
    MSFlexGrid2.TextMatrix(0&, i1&) = ne$
    w3& = Len(ne$)
    If w3& > mh% Then h% = w3& * 110 Else h% = mh% * 110
    MSFlexGrid2.ColWidth(i1&) = h%
    gw& = gw& + h%
    Tabla.Caption = ABLTAB(w2&).fejlec
    If InStr(ar$, "J") > 0 Then
      MSFlexGrid2.ColAlignment(i1&) = 6
    Else
      MSFlexGrid2.ColAlignment(i1&) = 1
    End If
    mtb$(i1&) = ar$: mho%(i1&) = mh%
    mesor%(i1&) = ABLTAB(w2&).adatsorsz(i1&)
  Next
  '--- méretek beállítása
  MSFlexGrid2.Top = 2600
  MSFlexGrid2.Left = 120
  MSFlexGrid2.Width = 12080
  MSFlexGrid2.Height = 5150
  MSFlexGrid2.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call Command1_Click
End Sub
Private Sub nyelad()
  Unload Nyugel1
  Nyugel1.betoltve = 0
  Nyugel1.Caption = ttkrak + Label6.Caption
  Nyugel1.MSFlexGrid1.Clear
  Nyugel1.MSFlexGrid1.Rows = 201
  ' 2015.7.29
  'Nyugel1.MSFlexGrid1.Cols = 8
  Nyugel1.MSFlexGrid1.Cols = 9
  
  Nyugel1.MSFlexGrid1.FixedCols = 1
  Nyugel1.MSFlexGrid1.ColWidth(0) = 400
  Nyugel1.MSFlexGrid1.ColAlignment(0) = 1
  Nyugel1.MSFlexGrid1.TextMatrix(0, 0) = "Ssz"
  Nyugel1.MSFlexGrid1.ColWidth(1) = 1400
  Nyugel1.MSFlexGrid1.ColAlignment(1) = 1
  Nyugel1.MSFlexGrid1.TextMatrix(0, 1) = "Termék kód"
  Nyugel1.MSFlexGrid1.ColWidth(2) = 5100
  Nyugel1.MSFlexGrid1.ColAlignment(2) = 1
  Nyugel1.MSFlexGrid1.TextMatrix(0, 2) = "Megnevezés"
  ' 2015.7.29
  Nyugel1.MSFlexGrid1.ColWidth(3) = 1000  ' 1200
  
  Nyugel1.MSFlexGrid1.ColAlignment(3) = 7
  Nyugel1.MSFlexGrid1.TextMatrix(0, 3) = "Mennyiség"
  Nyugel1.MSFlexGrid1.ColWidth(4) = 1200
  Nyugel1.MSFlexGrid1.ColAlignment(4) = 7
  Nyugel1.MSFlexGrid1.TextMatrix(0, 4) = "Egységár"
  Nyugel1.MSFlexGrid1.ColWidth(5) = 800
  Nyugel1.MSFlexGrid1.ColAlignment(5) = 7
  Nyugel1.MSFlexGrid1.TextMatrix(0, 5) = "Eng.%"
  Nyugel1.MSFlexGrid1.ColWidth(6) = 1000
  Nyugel1.MSFlexGrid1.ColAlignment(6) = 7
  Nyugel1.MSFlexGrid1.TextMatrix(0, 6) = "Készlet"
  Nyugel1.MSFlexGrid1.ColWidth(7) = 700
  Nyugel1.MSFlexGrid1.ColAlignment(7) = 1
  Nyugel1.MSFlexGrid1.TextMatrix(0, 7) = "M.egys"
  
  ' 2015.7.29
  Nyugel1.MSFlexGrid1.ColWidth(8) = 700
  Nyugel1.MSFlexGrid1.ColAlignment(8) = 1
  Nyugel1.MSFlexGrid1.TextMatrix(0, 8) = "Raktár"

  Nyugel1.MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  Nyugel1.MSFlexGrid1.Font.Size = 9
  Nyugel1.Text1.Text = ""
  Nyugel1.Text2.Text = ""
  Nyugel1.Text3.Text = ""
  Nyugel1.Text4(1).Text = ""
  Nyugel1.Text5.Text = "Készpénz"
  Nyugel1.Text6.Text = ""
  Nyugel1.Text7.Text = "01"
  Nyugel1.Text9.Text = ""
  Nyugel1.Label1.Caption = "0,00"
  Nyugel1.MSFlexGrid1.Row = 1: Nyugel1.MSFlexGrid1.Col = 1
  Nyugel1.Command11.Enabled = True
  Nyugel1.Command10.Enabled = True
  Nyugel1.Command9.Enabled = True
  Nyugel1.Command14.Enabled = True
  Nyugel1.Command17.Enabled = True
  Nyugel1.Command15.Enabled = True
  Nyugel1.Command6.Enabled = True
  Nyugel1.Command8.Enabled = True
  Nyugel1.Command12.Enabled = True
  Nyugel1.Command2.Enabled = True
  Nyugel1.Text9.Enabled = True
  Nyugel1.Text3.Enabled = True
  Nyugel1.Check1.Enabled = True
  Nyugel1.Check1.Value = 0

  
  

  
  Nyugel1.Show vbModal
End Sub

Private Sub keretkarb(rec$)
  '--- járat karbantartása
  sznev$ = " " + Trim(Mid$(rec$, 1, 15)) + " " + Trim(Mid$(rec$, 31, 60)) + " "
  telep$ = "  " + Trim(Mid$(rec$, 121, 8)) + " " + Trim(Mid$(rec$, 129, 30)) + ", " + Trim(Mid$(rec$, 159, 30)) + " " + Trim(Mid$(rec$, 189, 10)) + " "
  Rkeret.Label2.Caption = sznev$ + Chr$(13) + telep$
  Rkeret.MSFlexGrid1.Clear
  Rkeret.Top = 1000
  Rkeret.MSFlexGrid1.Rows = 201
  Rkeret.MSFlexGrid1.Cols = 17
  Rkeret.MSFlexGrid1.FixedCols = 1
  Rkeret.MSFlexGrid1.RowHeight(0) = 400
  Rkeret.MSFlexGrid1.ColWidth(0) = 300
  Rkeret.MSFlexGrid1.ColAlignment(0) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 0) = "Sz"
  Rkeret.MSFlexGrid1.ColWidth(1) = 1200
  Rkeret.MSFlexGrid1.ColAlignment(1) = 1
  Rkeret.MSFlexGrid1.TextMatrix(0, 1) = "Termék kód"
  Rkeret.MSFlexGrid1.ColWidth(2) = 850
  Rkeret.MSFlexGrid1.ColAlignment(2) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 2) = "Hétfõ keret"
  Rkeret.MSFlexGrid1.ColWidth(3) = 850
  Rkeret.MSFlexGrid1.ColAlignment(3) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 3) = "Hétfõ módosított"
  Rkeret.MSFlexGrid1.ColWidth(4) = 850
  Rkeret.MSFlexGrid1.ColAlignment(4) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 4) = "Kedd keret"
  Rkeret.MSFlexGrid1.ColWidth(5) = 850
  Rkeret.MSFlexGrid1.ColAlignment(5) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 5) = "Kedd módosított"
  Rkeret.MSFlexGrid1.ColWidth(6) = 850
  Rkeret.MSFlexGrid1.ColAlignment(6) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 6) = "Szerda keret"
  Rkeret.MSFlexGrid1.ColWidth(7) = 850
  Rkeret.MSFlexGrid1.ColAlignment(7) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 7) = "Szerda módosított"
  Rkeret.MSFlexGrid1.ColWidth(8) = 850
  Rkeret.MSFlexGrid1.ColAlignment(8) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 8) = "Csütörtök keret"
  Rkeret.MSFlexGrid1.ColWidth(9) = 850
  Rkeret.MSFlexGrid1.ColAlignment(9) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 9) = "Csütörtök módosított"
  Rkeret.MSFlexGrid1.ColWidth(10) = 850
  Rkeret.MSFlexGrid1.ColAlignment(10) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 10) = "Péntek keret"
  Rkeret.MSFlexGrid1.ColWidth(11) = 850
  Rkeret.MSFlexGrid1.ColAlignment(11) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 11) = "Péntek módosított"
  Rkeret.MSFlexGrid1.ColWidth(12) = 850
  Rkeret.MSFlexGrid1.ColAlignment(12) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 12) = "Szombat keret"
  Rkeret.MSFlexGrid1.ColWidth(13) = 850
  Rkeret.MSFlexGrid1.ColAlignment(13) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 13) = "Szombat módosított"
  Rkeret.MSFlexGrid1.ColWidth(14) = 850
  Rkeret.MSFlexGrid1.ColAlignment(14) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 14) = "Vasárnap keret"
  Rkeret.MSFlexGrid1.ColWidth(15) = 850
  Rkeret.MSFlexGrid1.ColAlignment(15) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 15) = "Vasárnap módosított"
  Rkeret.MSFlexGrid1.ColWidth(16) = 850
  Rkeret.MSFlexGrid1.ColAlignment(16) = 7
  Rkeret.MSFlexGrid1.TextMatrix(0, 16) = "Nettó egységár"
  For i77% = 1 To 200
    Rkeret.MSFlexGrid1.TextMatrix(i77%, 0) = Trim(Str(i77%))
  Next
  opo% = 0
  For i77% = 1 To cikkekdb%
    azo$ = Mid$(rec$, 1, 15) + kcikkek$(i77%)
    keretrec$ = dbxkey("RKRE", azo$)
    If keretrec$ <> "" Then
      opo% = opo% + 1
      Rkeret.MSFlexGrid1.TextMatrix(opo%, 1) = kcikkek$(i77%)
      Rkeret.MSFlexGrid1.TextMatrix(opo%, 16) = kegysar@(i77%)
      For j3% = 1 To 14
        Rkeret.MSFlexGrid1.Row = i77%
        Rkeret.MSFlexGrid1.Col = j3% + 1
        If Rkeret.MSFlexGrid1.Col <> 1 Then
          If Rkeret.MSFlexGrid1.Col Mod 2 = 0 Then
            Rkeret.MSFlexGrid1.CellForeColor = RGB(0, 60, 40)
          Else
            Rkeret.MSFlexGrid1.CellForeColor = RGB(110, 0, 0)
          End If
        Else
          Rkeret.MSFlexGrid1.CellForeColor = 0
        End If
        o@ = xval(Mid$(keretrec$, (j3% - 1) * 10 + 44, 10))
        If o@ = 0 Then
          Rkeret.MSFlexGrid1.TextMatrix(opo%, j3% + 1) = ""
        Else
          Rkeret.MSFlexGrid1.TextMatrix(opo%, j3% + 1) = Trim(ertszam(Str(o@), 10, 2))
        End If
      Next
    End If
  Next
  Rkeret.MSFlexGrid1.Row = 1
  Rkeret.MSFlexGrid1.Col = 1
  Rkeret.Text1.Text = ""
  Rkeret.betoltve = 0
  Rkeret.ipartkod = Mid$(rec$, 16, 15)
  Rkeret.Show vbModal
  opo% = 0
  If rogzites% <> 0 Then
    Call dbxtrkezd("RKRE")
    If cikkekdb% > 0 Then
      '--- sztornózás
      For j& = 1 To cikkekdb%
        azo$ = Mid$(rec$, 1, 15) + kcikkek$(j&)
        keretrec$ = dbxkey("RKRE", azo$)
        If keretrec$ <> "" Then
          Mid$(keretrec$, 43, 1) = "S"
          Call dbxki("RKRE", keretrec$, ";", "", "", hiba%)
        End If
      Next
    End If
    '--- újak rögzítése
    For j& = 1 To 200
      ktrmkod$ = Left(Trim(Rkeret.MSFlexGrid1.TextMatrix(j&, 1)) + Space(15), 15)
      If Trim(ktrmkod$) <> "" Then
        azo$ = Mid$(rec$, 1, 15) + ktrmkod$
        keretrec$ = dbxkey("RKRE", azo$)
        If keretrec$ = "" Then
          '--- újonan rendelt
          keretrec$ = Space(200)
          Mid$(keretrec$, 1, 30) = azo$
          o@ = xval(Trim(Rkeret.MSFlexGrid1.TextMatrix(j&, 16)))
          Mid$(keretrec$, 184, 12) = ertszam(Str(o@), 12, 2)
          For j3% = 1 To 14
            o@ = xval(Trim(Rkeret.MSFlexGrid1.TextMatrix(j&, j3% + 1)))
            Mid$(keretrec$, (j3% - 1) * 10 + 44, 10) = ertszam(Str(o@), 10, 2)
          Next
          Call dbxki("RKRE", keretrec$, ";", "U", "", hiba%)
        Else
          '--- korábban is rendelt
          Mid$(keretrec$, 43, 1) = " "
          o@ = xval(Trim(Rkeret.MSFlexGrid1.TextMatrix(j&, 16)))
          Mid$(keretrec$, 184, 12) = ertszam(Str(o@), 12, 2)
          For j3% = 1 To 14
            o@ = xval(Trim(Rkeret.MSFlexGrid1.TextMatrix(j&, j3% + 1)))
            Mid$(keretrec$, (j3% - 1) * 10 + 44, 10) = ertszam(Str(o@), 10, 2)
          Next
          Call dbxki("RKRE", keretrec$, ";", "", "", hiba%)
        End If
      End If
    Next
    Call dbxtrvege
  End If
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyH Then
    KeyCode = 0
    If Shift And vbAltMask Then
      Call alth("KRAK", ujazonosito$)
      If ujazonosito$ <> "" Then
        Text5.Text = ujazonosito$
      End If
    End If
  End If
End Sub
Private Sub kilepp()
  '--- számla mégsem
  List1.Visible = False
  Command1.Visible = True
  Command3.Visible = True
  MSFlexGrid1.Visible = False
  Info.Visible = False
  Text1.Visible = False
  Text2.Visible = False
  Command4.Visible = True
  Command5.Visible = True
  Command6.Visible = True
  Command7.Visible = True
  Command8.Visible = True
  Text6.Visible = True
  Label7.Visible = True
  Info.Visible = False
  MSFlexGrid1.Visible = False
  MSFlexGrid2.Visible = False
  Label4.Visible = True
  Label5.Visible = True
  Text3.Visible = True
  Text4.Visible = True
  Command1.SetFocus
End Sub

Private Sub Text8_LostFocus()
   
    maidatum = Text8.Text
    Call hetinap(maidatum$, maisorsz%, napnev$)
    Label4.Caption = " A mai nap:" + datki(maidatum) + " " + napnev$ + " "
End Sub
Private Sub szetvalaszt(sor1$, sor2$, hossz, tip)
Dim valaszt As Integer

If Len(Trim$(sor1)) > hossz Then
  valaszt = 0
  For i = hossz To 20 Step -1
    a$ = Mid$(sor1$, i, 1)
    If tip = 1 Then
      If a$ = " " Or a$ = "," Or a$ = "." Or a$ = ";" Or a$ = ")" Then
        valaszt = i
        Exit For
      End If
    Else
      If a$ = ")" Then
        valaszt = i + 1
        Exit For
      End If
   
    End If
  Next
  sor2$ = Mid$(sor1, valaszt + 1, 60 - valaszt)
  sor1$ = Mid$(sor1, 1, valaszt)
End If
  
End Sub

Private Sub Beallit()

        Nyugel1.Command11.Enabled = False
        Nyugel1.Command10.Enabled = False
        Nyugel1.Command9.Enabled = False
        Nyugel1.Command14.Enabled = False
        Nyugel1.Command17.Enabled = False
        Nyugel1.Command15.Enabled = False
        Nyugel1.Command6.Enabled = False
        Nyugel1.Command12.Enabled = False
End Sub
Private Sub Partnertolt()
      pkod$ = UCase(Left(Trim(Nyugel1.Text4(1).Text + Space(15)), 15))
      Nyugel1.Text4(1).Text = pkod$
      If Trim(pkod$) <> "" Then
        prec$ = dbxkey("PART", pkod$)
        If prec$ = "" Then
          Call mess("Hibás vevõ kód!", 3, 0, "Hiba", valasz%)
          KeyCode = 0
        Else
          If Trim$(pkod$) = "BANK" Then
            If Trim$(Nyugel1.Text2.Text) = "" Then
              Nyugel1.Text2.Text = Trim(Mid$(prec$, 16, 60))
              Nyugel1.Text3.Text = postacim(prec$, 106)
            End If
          Else
            Nyugel1.Text2.Text = Trim(Mid$(prec$, 16, 60))
            Nyugel1.Text3.Text = postacim(prec$, 106)
            Nyugel1.Text12.Text = Trim(Mid$(prec$, 184, 15))
            If Not Len(Trim$(Nyugel1.Text12.Text)) = 0 Then
            If Not Mid$(Nyugel1.Text12.Text, 9, 1) = "-" Then
              Nyugel1.Text12.Text = Mid$(Nyugel1.Text12.Text, 1, 8) + "-" + Mid$(Nyugel1.Text12.Text, 9, 1) + "-" + Mid$(Nyugel1.Text12.Text, 10, 2)
            End If
            End If
          End If
          fmkod$ = Mid$(prec$, 328, 2)
          If Not Trim$(fmkod$) = "" Then
            fmrec$ = dbxkey("PFIZ", fmkod$)
            Nyugel1.Text5.Text = Mid$(fmrec$, 3, 30)
            Nyugel1.Text7.Text = fmkod$
          End If
          fizhatido% = Val(Mid$(prec$, 330, 3))
          fidat$ = maidatum$
          For i13% = 1 To fizhatido%
            xxxx$ = novdat(fidat$)
            fidat$ = xxxx$
          Next
          Nyugel1.Text6.Text = fidat$
          Nyugel1.Text2.SelStart = Len(Trim(Nyugel1.Text2.Text)) + 1
          
        End If
    End If
End Sub

Sub feltolt_folyokonyvel(pvszrec$, szamlaszam$, fejr$)

  Mid$(pvszrec$, 8, 10) = szamlaszam$
  Mid$(pvszrec$, 38, 15) = Mid$(fejr$, 1, 15)
  Mid$(pvszrec$, 211, 6) = Mid$(fejr$, 24, 6)
  Mid$(pvszrec$, 58, 6) = Mid$(fejr$, 24, 6)
  Mid$(pvszrec$, 64, 6) = Mid$(fejr$, 18, 6)
  Mid$(pvszrec$, 70, 6) = Mid$(fejr$, 30, 6)
  Mid$(pvszrec$, 76, 2) = Mid$(fejr$, 36, 2)

End Sub

Sub szamlakeszit(qshow$, fejr$, irec$, partrec$, fizmrec$, onert@, obert@, szlacim$, szlamod$, ssz$, pszamlaszam$, tdb, kpkod$, fr4nev$, nyugtavolt, szlafaj$, rpartner$)
    ssz$ = pszamlaszam$
    Call szlaparair(szlamod$, ssz$, szlacim$, szlafaj$)
    
    If partrec$ = "" Then
      pkod$ = Mid$(fejr$, 1, 15)
      If Trim(pkod) = "BANK" Then
         pkod$ = rpartner$
      End If
      If Not RTrim$(pkod$) = "" Then
        partrec$ = dbxkey("PART", pkod$)
      End If
    End If
    
    filwli = FreeFile
    Open listautvonal$ + terminal$ + task$ + "AUW.WLI" For Output Shared As #filwli
    
    Wrtbrec$ = Space(3000)
    Mid$(Wrtbrec$, 11, 3) = pszamlaszam$
    Mid$(Wrtbrec$, 11, 3) = Right("   " + Str(tdb), 3)
    Mid$(Wrtbrec$, 24, 8) = "auw-qrptn"
    Mid$(Wrtbrec$, 32, 12) = terminal$ + task$ + "SZLA.LST"
    Mid$(Wrtbrec$, 44, 12) = fr4nev$ + ".fx4"          'Formátum
    Mid$(Wrtbrec$, 56, 10) = qshow$    '"dbx4-qshz"
    Mid$(Wrtbrec$, 66, 2) = "BS"
    Mid$(Wrtbrec$, 68, 1) = "1"
    Mid$(Wrtbrec$, 69, 1) = "P"
    Mid$(Wrtbrec$, 70, 6) = Mid$(fejr$, 18, 6)             'Számla kelte
    Mid$(Wrtbrec$, 76, 6) = Mid$(fejr$, 24, 6)             'Teljesítés kelte
    Mid$(Wrtbrec$, 100, 12) = Mid$(irec$, 173, 15)        'szállító adószáma
    'Mid$(Wrtbrec$, 112, 12) = Mid$(irec$, 173, 15)       'csoportos adószám
    Mid$(Wrtbrec$, 124, 20) = Mid$(irec$, 188, 10)        'szállító EU adószáma
    Mid$(Wrtbrec$, 226, 60) = Mid$(irec$, 5, 60)          'Név
    Mid$(Wrtbrec$, 286, 10) = Mid$(irec$, 95, 8)          'Irsz
    Mid$(Wrtbrec$, 296, 30) = Mid$(irec$, 103, 30)        'Település
    Mid$(Wrtbrec$, 326, 10) = Mid$(irec$, 901, 10)        'Kerület
    Mid$(Wrtbrec$, 336, 30) = Mid$(irec$, 133, 30)        'Közterület
    Mid$(Wrtbrec$, 366, 10) = Mid$(irec$, 911, 30)        'Közterület jell
    Mid$(Wrtbrec$, 376, 10) = Mid$(irec$, 163, 10)        'Házszám
    Mid$(Wrtbrec$, 386, 10) = Mid$(irec$, 921, 10)        'Épület
    Mid$(Wrtbrec$, 396, 10) = Mid$(irec$, 931, 10)        'Lépcsõház
    Mid$(Wrtbrec$, 406, 10) = Mid$(irec$, 941, 10)        'Szint
    Mid$(Wrtbrec$, 416, 10) = Mid$(irec$, 951, 10)        'ajtó
    
    Mid$(Wrtbrec$, 426, 10) = Mid$(partrec$, 1, 15)        'Part kód
    Mid$(Wrtbrec$, 441, 12) = Mid$(partrec$, 184, 15)      'Part adószáma
    Mid$(Wrtbrec$, 453, 12) = Mid$(partrec$, 855, 15)      'Part csop adószáma
    Mid$(Wrtbrec$, 465, 12) = Mid$(partrec$, 199, 15)      'Part adószáma
    ' KP-ésszámla esetén - auwker - KCIM -ból
    If kpkod$ = "" Then
      If partrec$ = "" Then
         Call mess("Elõzõ évi számla sztonója! Hívja a fejlesztõt!", 2, 0, langprg(1), valasz%)
      End If
      Mid$(Wrtbrec$, 505, 60) = Mid$(partrec$, 16, 60)         'Név
      Mid$(Wrtbrec$, 565, 10) = Mid$(partrec$, 106, 8)         'Irsz
      Mid$(Wrtbrec$, 575, 30) = Mid$(partrec$, 114, 30)        'Település
      Mid$(Wrtbrec$, 605, 10) = Mid$(partrec$, 543, 10)        'Kerület
      Mid$(Wrtbrec$, 615, 30) = Mid$(partrec$, 144, 30)        'Közterület
      Mid$(Wrtbrec$, 645, 10) = Mid$(partrec$, 553, 10)        'Közterület jell
      Mid$(Wrtbrec$, 655, 10) = Mid$(partrec$, 174, 10)        'Házszám
      Mid$(Wrtbrec$, 665, 10) = Mid$(partrec$, 563, 10)        'Épület
      Mid$(Wrtbrec$, 675, 10) = Mid$(partrec$, 573, 10)        'Lépcsõház
      Mid$(Wrtbrec$, 685, 10) = Mid$(partrec$, 583, 10)        'Szint
      Mid$(Wrtbrec$, 695, 10) = Mid$(partrec$, 593, 10)        'ajtó
    Else
        cimbrec$ = dbxkey("KCIM", kpkod$)
        If Not cimbrec$ = "" Then
          kparrec$ = dbxkey("KPAR", kpkod$)
          Mid$(Wrtbrec$, 505, 60) = Mid$(kparrec$, 1, 60)          'Név
          Mid$(Wrtbrec$, 565, 10) = Mid$(cimbrec$, 1, 8)         'Irsz
          Mid$(Wrtbrec$, 575, 30) = Mid$(cimbrec$, 10, 30)        'Település
          Mid$(Wrtbrec$, 605, 10) = Mid$(cimbrec$, 40, 10)        'Kerület
          Mid$(Wrtbrec$, 615, 30) = Mid$(cimbrec$, 50, 30)        'Közterület
          Mid$(Wrtbrec$, 645, 10) = Mid$(cimbrec$, 80, 10)        'Közterület jell
          Mid$(Wrtbrec$, 655, 10) = Mid$(cimbrec$, 90, 10)        'Házszám
          Mid$(Wrtbrec$, 665, 10) = Mid$(cimbrec$, 100, 10)        'Épület
          Mid$(Wrtbrec$, 675, 10) = Mid$(cimbrec$, 110, 10)        'Lépcsõház
          Mid$(Wrtbrec$, 685, 10) = Mid$(cimbrec$, 120, 10)        'Szint
          Mid$(Wrtbrec$, 695, 10) = Mid$(cimbrec$, 130, 10)        'ajtó
        Else
          Call mess("Nincs cím felbontva", 2, 0, langprg(1), valasz%)
          kparrec$ = dbxkey("KPAR", kpkod$)
          Mid$(Wrtbrec$, 505, 60) = Mid$(kparrec$, 1, 60)          'Név
        End If
      
    End If
    Mid$(Wrtbrec$, 1013, 6) = Mid$(fejr$, 30, 6)             'fizetési hatid
    Mid$(Wrtbrec$, 1019, 2) = Mid$(fejr$, 36, 2)             'fizetési mód
    Mid$(Wrtbrec$, 1021, 20) = Mid$(fizmrec$, 3, 30)         'fizetési mód
    Mid$(Wrtbrec$, 1041, 20) = "M"
    Mid$(Wrtbrec$, 1072, 24) = Mid$(irec$, 233, 24)           ' Bankszámla
    
    'Áfa összesítés
    jafa = 1500
    If afatomb(1, 1) + afatomb(2, 1) <> 0 Then
       Mid$(Wrtbrec$, jafa, 6) = 0
       Mid$(Wrtbrec$, jafa + 6, 14) = afatomb(1, 1) + afatomb(2, 1)
       Mid$(Wrtbrec$, jafa + 20, 14) = afatomb(1, 2) + afatomb(2, 2)
       Mid$(Wrtbrec$, jafa + 34, 14) = afatomb(1, 1) + afatomb(2, 1) + afatomb(1, 2) + afatomb(2, 2)
       jafa = jafa + 48
    End If
    oafaert@ = 0
    For ij = 3 To 6
       If afatomb(ij, 1) <> 0 Then
          Mid$(Wrtbrec$, jafa, 6) = Right$(Space$(6) + Str(afakulcsok(ij - 2)), 6)
          Mid$(Wrtbrec$, jafa + 6, 14) = ertszam(Str(afatomb(ij, 1)), 14, 2)
          Mid$(Wrtbrec$, jafa + 20, 14) = ertszam(Str(afatomb(ij, 2)), 14, 2)
          Mid$(Wrtbrec$, jafa + 34, 14) = ertszam(Str(afatomb(ij, 1) + afatomb(ij, 2)), 14, 2)
          oafaert@ = oafaert@ + afatomb(ij, 2)
          jafa = jafa + 48
       End If
    Next
    Mid$(Wrtbrec$, 1840, 14) = ertszam(Str$(onert@), 14, 2)
    Mid$(Wrtbrec$, 1854, 14) = ertszam(Str$(oafaert@), 14, 2)
    Mid$(Wrtbrec$, 1868, 14) = ertszam(Str$(obert@), 14, 2)
    Mid$(Wrtbrec$, 1888, 14) = ertszam(0, 14, 0)
    Mid$(Wrtbrec$, 1902, 14) = ertszam(Str$(obert@), 14, 2)
       
    Print #filwli, Wrtbrec$
    
    For i = 1 To tdb
      Print #filwli, wrtt$(i)
    Next
    
    Close filwli
         
    fnev$ = terminal$ + task$ + "AUW.WLE"
    
    If UCase(Dir(listautvonal$ + fnev$)) = UCase(fnev$) Then
       Kill listautvonal$ + fnev$
    End If
    prog$ = programutvonal$ + "auw-szla " + terminal$ + task$ + "/" + listautvonal$
    Shell prog$, vbNormalFocus
    
    ' Várni a fájlra
    Text1.Text = "Számlázásra várok"
    Text1.Visible = True
        
    Do
      DoEvents
      s$ = UCase(Dir(listautvonal$ + fnev$))
    Loop While s$ <> UCase(fnev$)

    Text1.Text = "Majdnem megvan   "
    sm1& = FileLen(listautvonal$ + fnev$)
    Do
      DoEvents
      Call waitsec(1)
      sm& = FileLen(listautvonal$ + fnev$)
      If sm1& = sm& Then Exit Do
      sm1& = sm&
    Loop While sm& > 0
    
    Text1.Text = "Számlaszámra olv."
    filwle = FreeFile
    Open listautvonal$ + terminal$ + task$ + "AUW.WLE" For Input Shared As #filwle
    Line Input #filwle, szamlaszam$
    Close filwle
    
    Text1.Visible = False
    
End Sub

Sub szlaparair(szlamod$, ssz$, szlacim$, szlafaj$)
    filwlh = FreeFile
    Open listautvonal$ + terminal$ + task$ + "AUW.WLH" For Output Shared As #filwlh
    Print #filwlh, "CIM=" + szlacim$
    Print #filwlh, "MOD=" + szlamod$
    Print #filwlh, "FAJ=" + szlafaj$
    Print #filwlh, "SSZ=" + ssz$
    
    Close filwlh
End Sub

Sub partnevcimtarol(kbrec$, pszbrec$, nyugtavolt)
   Call dbxtrkezd("KSYB")
   psybrec$ = Space$(200)
   Mid$(psybrec$, 1, 7) = Mid$(kbrec$, 1, 7)
   Mid$(psybrec$, 8, 15) = Mid$(pszbrec$, 61, 15)
   If RTrim$(Mid$(pszbrec$, 61, 15)) = "" Or Mid$(pszbrec$, 61, 4) = "BANK" Then
      Mid$(psybrec$, 8, 15) = Mid$(pszbrec$, 61, 15)
      Mid$(psybrec$, 16, 7) = Nyugel1.Text13
   Else
      Mid$(psybrec$, 8, 15) = Mid$(pszbrec$, 61, 15)
   End If
   Mid$(psybrec$, 23, 60) = Nyugel1.Text2          ' Partner neve
   Mid$(psybrec$, 83, 60) = Nyugel1.Text3          ' Partner címe
   Mid$(psybrec$, 143, 58) = Nyugel1.Text11        ' Megjegyzés 2. sor

   Call dbxki("KSYB", psybrec$, ";", "U", "", hiba%)
   If nyugtavolt = 2 Or nyugtavolt = 3 Or nyugtavolt = 7 Then
         Call KPAR_Tolt
   
         For i99% = 1 To 4
            elem$ = Nyugel1.NtAtad(i99%)
            If Trim$(elem$) <> "" Then
               Mid$(psybrec$, (i99% - 1) + 143, 36) = Mid$(elem$, 1, 36)
            End If
         Next
   
   End If
   Call dbxtrvege

End Sub
Sub Eloleg_beszamitas(stjel$)
 For i1% = 1 To 200
                 '--- elõlegbeszámítás könyvelése
                If wbeszeloikt(i1%) <> "" Then
                  pelvrec$ = dbxkey("PELV", wbeszeloikt(i1%))
                  If pelvrec$ <> "" Then
                    If xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3)) <> 0 Then
                      wmenny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
                      wbruttoo@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3)) * xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 4))
                    Else
                      wmenny@ = 0
                      wbruttoo@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 4))
                    End If
                    
                    tkod$ = Left(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) + Space(15), 15)
                    termrec$ = dbxkey("KTRM", tkod$)
                    afakod$ = Mid$(termrec$, 706, 2)
                    afrec$ = dbxkey("PAFA", afakod$)
                    afakulcs@ = xval(Mid$(afrec$, 33, 6))
                    wneert@ = wbruttoo@ / ((100 + afakulcs@) / 100)
                    wneert@ = ertszam(Str(wneert@), 12, 0)
                    
                    wafaosz@ = wbruttoo@ - wneert@
                    If stjel$ = "S" Then
                       wbruttoo@ = xval(Mid$(pelvrec$, 449, 14)) + wbruttoo@
                       wmenny@ = xval(Mid$(pelvrec$, 435, 14)) + wmenny@
                       wneert@ = xval(Mid$(pelvrec$, 463, 14)) + wneert@
                       wafaosz@ = xval(Mid$(pelvrec$, 477, 14)) + wafaosz@
                    Else
                       wbruttoo@ = xval(Mid$(pelvrec$, 449, 14)) - wbruttoo@
                       wmenny@ = xval(Mid$(pelvrec$, 435, 14)) - wmenny@
                       wneert@ = xval(Mid$(pelvrec$, 463, 14)) - wneert@
                       wafaosz@ = xval(Mid$(pelvrec$, 477, 14)) - wafaosz@
                    End If
                    Mid$(pelvrec$, 435, 14) = ertszam(Str(wmenny@), 14, 2)
                    Mid$(pelvrec$, 449, 14) = ertszam(Str(wbruttoo@), 14, 2)
                    Mid$(pelvrec$, 463, 14) = ertszam(Str(wneert), 14, 2)
                    If Trim(Mid(pszbrec$, 98, 3)) <> "" Then
                      Mid$(pelvrec$, 477, 14) = ertszam(Str(wafaosz@), 14, 2)
                    Else
                      Mid$(pelvrec$, 477, 14) = ertszam(Str(wafaosz@), 14, 0)
                    End If
                    Call dbxki("PELV", pelvrec$, ";", "", "", hiba%)
                  End If
                End If
 Next
End Sub

Public Sub nevcimtolt(ncv$, azo$)
    i& = UBound(nevcimt$)
    ncdb& = ncdb& + 1
    If ncdb > i& Then
      ReDim nevcimt$(1 To i& + 300)
      ReDim ncazonosito$(1 To i& + 300)
    End If
    nevcimt$(ncdb) = ncv$
    ncazonosito$(ncdb) = azo$
    
End Sub
Public Function nevcimkeres$(ncv$)
 van = False
 For j1% = 1 To ncdb&
   If nevcimt$(j1%) = ncv$ Then
    kparrec$ = dbxkey("KCIM", ncazonosito$(j1%))
    If Not Trim$(kparrec$) = "" Then
      van = True
      Exit For
    End If
   End If
Next
If van Then
    nevcimkeres$ = kparrec$
    Nyugel1.Text13.Text = ncazonosito$(j1%)
Else
    nevcimkeres$ = ""
End If

End Function

Public Function nevcimmod$(ncv$, azo$)
 van = False
 For j1% = 1 To ncdb&
   If ncazonosito$(j1%) = azo$ Then
      van = True
      Exit For
   End If
Next
If van Then
    nevcimt$(j1%) = ncv$
    nevcimmod$ = azo$

Else
    nevcimmod$ = ""
End If

End Function


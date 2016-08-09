VERSION 5.00
Begin VB.Form Mesbox 
   BackColor       =   &H0091E9FB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   1248
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   5124
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1248
   ScaleWidth      =   5124
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rendben"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mégsem"
      Default         =   -1  'True
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Igen"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   720
      TabIndex        =   3
      Top             =   180
      Width           =   4296
   End
   Begin VB.Image Image7 
      Height          =   456
      Left            =   5280
      Picture         =   "Mesbox.frx":0000
      Top             =   3120
      Width           =   552
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   3360
      Picture         =   "Mesbox.frx":150A
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   384
      Left            =   2280
      Picture         =   "Mesbox.frx":2954
      Top             =   3120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Image4 
      Height          =   384
      Left            =   960
      Picture         =   "Mesbox.frx":2D96
      Top             =   3120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Image3 
      Height          =   384
      Left            =   1680
      Picture         =   "Mesbox.frx":31D8
      Top             =   3120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Image2 
      Height          =   384
      Left            =   120
      Picture         =   "Mesbox.frx":361A
      Top             =   3120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   120
      Picture         =   "Mesbox.frx":3A5C
      Top             =   240
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "Mesbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  msgvalasz = 1:
  Call langclos
  Unload Me
End Sub

Private Sub Command2_Click()
  msgvalasz = 0
  Call langclos
  Unload Me
End Sub

Private Sub Command3_Click()
  Call langclos
  Unload Me
End Sub

Private Sub Form_Activate()
  If Command2.Visible = True Then Command2.SetFocus
End Sub

Private Sub Form_Load()
  Call langinit("Mesbox", 2)
  Call szkriptel("Mesbox")
End Sub

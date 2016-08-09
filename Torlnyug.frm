VERSION 5.00
Begin VB.Form Torlnyug 
   BackColor       =   &H009BF8FD&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Törlés nyugtázása"
   ClientHeight    =   2472
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2472
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command4"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mégsem"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Törlés"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1212
   End
   Begin VB.Image Image2 
      Height          =   384
      Left            =   120
      Picture         =   "Torlnyug.frx":0000
      Stretch         =   -1  'True
      Top             =   130
      Width           =   384
   End
   Begin VB.Image Image1 
      Height          =   372
      Left            =   120
      Picture         =   "Torlnyug.frx":0442
      Stretch         =   -1  'True
      Top             =   150
      Width           =   372
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bizosan törölni akarja?"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.2
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   252
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ha mégis a törlést választja, és van olyan tranzakció, amelyet nem sztornózott, a késõbbiekben ez hibát okozhat!"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.2
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   492
      Left            =   600
      TabIndex        =   3
      Top             =   1440
      Width           =   6252
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   6252
   End
End
Attribute VB_Name = "Torlnyug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Call langclos
  form1.torlnyugta = 1
  Torlnyug.Hide
End Sub

Private Sub Command2_Click()
  form1.torlnyugta = 0
  Call langclos
  Torlnyug.Hide
End Sub

Private Sub Form_Activate()
  If Label1.Caption = "" Then
    Torlnyug.Height = 1400
    Torlnyug.Width = 3200
    Label2.Left = Label3.Left
    Label2.Top = Label3.Top
    Command1.Left = Command3.Left
    Command2.Left = Command4.Left
    Command1.Top = Command3.Top
    Command2.Top = Command4.Top
    Image2.Visible = True: Image1.Visible = False
  Else
    Torlnyug.Height = 3000
    Torlnyug.Width = 7000
    Label2.Left = 300
    Label2.Top = 1440
    Label1.Left = 600
    Command1.Left = 4140
    Command2.Left = 5460
    Command1.Top = 2040
    Command2.Top = 2040
    Image1.Visible = True: Image2.Visible = False
  End If
End Sub

Private Sub Form_Load()
  Call langinit("torlnyug", 2)
  Call szkriptel("torlnyug")
End Sub

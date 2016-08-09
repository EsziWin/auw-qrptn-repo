VERSION 5.00
Begin VB.Form Szenged 
   BackColor       =   &H0091E9FB&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Számla végösszegébõl adott engedmény"
   ClientHeight    =   1332
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   4932
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1332
   ScaleWidth      =   4932
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mehet"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1212
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   3840
      TabIndex        =   4
      Text            =   "0"
      Top             =   480
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   3840
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label4 
      BackColor       =   &H0091E9FB&
      Caption         =   "%"
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
      TabIndex        =   6
      Top             =   490
      Width           =   252
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "xEngedmény a számla bruttó értékébõl:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   490
      Width           =   3612
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4680
      TabIndex        =   2
      Top             =   150
      Width           =   372
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "xEngedmény a számla nettó értékébõl:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   156
      Width           =   3612
   End
End
Attribute VB_Name = "Szenged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Call langclos
  Szenged.Hide
End Sub

Private Sub Form_Activate()
  If utfrmnev$ <> "szenged" Then Call frmlang("szenged")
End Sub

Private Sub Form_Load()
  Call langinit("szenged", 2)
  Call szkriptel("szenged")
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    KeyCode = 0
    Text2.SetFocus
  End If
  If KeyCode = vbKeyDelete Then
    KeyCode = 0
    Text1.Text = ""
  End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    KeyCode = 0
    Text1.SetFocus
  End If
  If KeyCode = vbKeyReturn Then
    KeyCode = 0
    Command1.SetFocus
  End If
  If KeyCode = vbKeyDelete Then
    KeyCode = 0
    Text2.Text = ""
  End If
End Sub


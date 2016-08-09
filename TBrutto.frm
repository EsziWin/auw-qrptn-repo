VERSION 5.00
Begin VB.Form TBrutto 
   Caption         =   "TÁMOP számla"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5460
   LinkTopic       =   "Form2"
   ScaleHeight     =   2496
   ScaleWidth      =   5460
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Csökkent"
      Height          =   252
      Left            =   4080
      TabIndex        =   4
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Növel"
      Height          =   252
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tovább"
      Height          =   612
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   2532
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      TabIndex        =   0
      Text            =   " 80000"
      Top             =   240
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Számla bruttó összege:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2772
   End
End
Attribute VB_Name = "TBrutto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Command2_Click()
Text1.Text = Str(Val(Text1.Text) + 10000)
End Sub

Private Sub Command3_Click()
Text1.Text = Str(Val(Text1.Text) - 10000)
End Sub

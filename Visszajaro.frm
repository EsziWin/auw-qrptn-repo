VERSION 5.00
Begin VB.Form Visszajaro 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Visszajáró számítás"
   ClientHeight    =   4320
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7392
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7392
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Bezár"
      Height          =   492
      Left            =   2880
      TabIndex        =   5
      Top             =   3840
      Width           =   1812
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
      Left            =   3360
      TabIndex        =   1
      Top             =   1440
      Width           =   1932
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   252
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   7452
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   372
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   7332
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   252
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   7332
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Index           =   1
      Left            =   3600
      TabIndex        =   8
      Top             =   3360
      Width           =   2292
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   492
      Index           =   1
      Left            =   3600
      TabIndex        =   7
      Top             =   2760
      Width           =   2292
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   3600
      TabIndex        =   6
      Top             =   2160
      Width           =   2292
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Visszajáró:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   3360
      Width           =   2292
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Kapott:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   492
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   2760
      Width           =   2292
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fizetendõ:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   2292
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Készpénz:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1800
      TabIndex        =   0
      Top             =   1560
      Width           =   1332
   End
End
Attribute VB_Name = "Visszajaro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kilep
Private Sub Command1_Click()
Me.Hide
End Sub


Private Sub Form_Activate()
Text1.Text = ""
kilep = 0

Label3(1) = ""
Label4(1) = ""

Text1.SetFocus
End Sub

Private Sub Form_Load()
kilep = 0
Text1.Text = ""

Label3(1) = ""
Label4(1) = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    '--- mezõ ellenõrzés
    kilep = 0
    Command1.SetFocus
  End If
  If KeyCode = vbKeyEscape Then
    kilep = 1
    Me.Hide
  End If
End Sub

Private Sub Text1_LostFocus()
If kilep = 0 Then
  If Not Text1.Text = "" Then
   kapot@ = Val(Text1.Text)
   fizet@ = Val(Label2(1))
   If kapot@ < fizet@ Then
    Call mess("Kevés a készpénz!", 3, 0, "Figyelmeztetés", valasz%)
    ' Text1.SetFocus
   End If
   visszajar@ = kapot@ - fizet@

   Label3(1) = Right$(Space$(12) + Format(kapot@, "# ### ### ##0"), 12)
   Label4(1) = Right$(Space$(12) + Format(visszajar@, "# ### ### ##0"), 12)
  Else
    Me.Hide
  End If
End If
End Sub

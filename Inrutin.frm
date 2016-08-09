VERSION 5.00
Begin VB.Form Inrutin 
   BackColor       =   &H0091E9FB&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Azonosító megadása"
   ClientHeight    =   1548
   ClientLeft      =   156
   ClientTop       =   300
   ClientWidth     =   4848
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   10.8
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1548
   ScaleWidth      =   4848
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Esc"
      Top             =   1080
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ok"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Enter"
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   4572
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.2
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Adatmezõ:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   96
      TabIndex        =   0
      Top             =   120
      Width           =   3852
   End
End
Attribute VB_Name = "Inrutin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Call langclos
  Inrutin.Hide
End Sub

Private Sub Command2_Click()
  Text1.Text = ""
  Call langclos
  Inrutin.Hide
End Sub

Private Sub Form_Activate()
  Text1.SetFocus
End Sub

Private Sub Form_Load()
  Call langinit("inrutin", 2)
  Call szkriptel("inrutin")
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyH Then
    '--- Alt+H kapcsolat
    If Shift And vbAltMask Then
      '--- Alt+H
      If objektum$ = "JPAR" Then
        KeyCode = 0
        azonosito$ = ""
        Call alth("PART", azonosito$)
        If azonosito$ <> "" Then
          Text1.Text = Trim$(azonosito$)
        End If
      End If
    End If
  End If
  If KeyCode = vbKeyEscape Then
    KeyCode = 0
    Text1.Text = ""
    Call langclos
    Inrutin.Hide
    'Unload Me
  Else
    If KeyCode = vbKeyReturn Then
      atex$ = Trim(Text1.Text)
      atexh% = Len(atex$)
      If atexh% > 0 Then
        For iaa% = 1 To atexh%
          If Mid$(atex$, iaa%, 1) = "ö" Then Mid$(atex$, iaa%, 1) = "0"
        Next
        Text1.Text = atex$
      End If
      KeyCode = 0
      Call langclos
      'Unload Me
      Inrutin.Hide
    Else
      If KeyCode = vbKeyDelete Then
        Text1.Text = ""
      End If
    End If
  End If
End Sub

VERSION 5.00
Begin VB.Form ListEdit1 
   BackColor       =   &H0091E9FB&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Számlázandó szállítólevelek"
   ClientHeight    =   3816
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
   ScaleHeight     =   3816
   ScaleWidth      =   4848
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mind töröl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1212
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Töröl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   972
   End
   Begin VB.ListBox List1 
      Height          =   3216
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2172
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Esc"
      Top             =   3360
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Enter"
      Top             =   3360
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
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3360
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label Label2 
      Height          =   252
      Left            =   3480
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Szállító száma:"
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
Attribute VB_Name = "ListEdit1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  form1.List1.Clear
  For i = 1 To List1.ListCount
    form1.List1.AddItem (List1.List(i - 1))
  Next

  Call langclos
  ListEdit1.Hide
End Sub

Private Sub Command2_Click()
  
  Call langclos
  ListEdit1.Hide
End Sub

Private Sub Command3_Click()
If List1.ListIndex >= 0 Then
szallevikt = List1.List(List1.ListIndex)
List1.RemoveItem (List1.ListIndex)
' Nyugel-t módosítani
poz% = InStr(Nyugel1.Text8, Trim$(Str$(Val(szallevikt))))
tSzallevikt = Trim$(Str$(Val(szallevikt)))
If poz% = 1 And Len(tSzallevikt) = Len(Nyugel1.Text8) Then
  Nyugel1.Text8 = ""
Else
  hossz% = Len(Nyugel1.Text8) + Len(tSzallevikt)
  Nyugel1.Text8 = Left$(Nyugel1.Text8, poz% - 1) + Mid$(Nyugel1.Text8, poz% + Len(tSzallevikt) + 1, hossz%)
End If
For i% = 1 To 200
  If Not Trim(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1)) = "" Then
     If Nyugel1.MSFlexGrid1.TextMatrix(i%, 6) = szallevikt Then
        For j% = 1 To 7
           Nyugel1.MSFlexGrid1.TextMatrix(i%, j%) = " "
        Next
     End If
  End If
Next
End If
End Sub

Private Sub Command4_Click()
 Nyugel1.Text8 = ""
 List1.Clear
End Sub

Private Sub Form_Activate()
  List1.Clear
  For i = 1 To form1.List1.ListCount
    List1.AddItem (form1.List1.List(i - 1))
  Next

End Sub

Private Sub Form_Load()
  Call langinit("inrutin", 2)
  Call szkriptel("inrutin")
End Sub


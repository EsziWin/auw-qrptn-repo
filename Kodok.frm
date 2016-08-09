VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Kodok 
   BackColor       =   &H0055D7F7&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Kódok értelmezése"
   ClientHeight    =   3636
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   5508
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3636
   ScaleWidth      =   5508
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
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
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   4092
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2892
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5292
      _ExtentX        =   9335
      _ExtentY        =   5101
      _Version        =   327680
      BackColorFixed  =   9562619
      BackColorSel    =   16777152
      BackColorBkg    =   9562619
      ScrollBars      =   0
      SelectionMode   =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1092
   End
End
Attribute VB_Name = "Kodok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  vakod$ = ""
  Call langclos
  Kodok.Hide
End Sub

Private Sub Form_Activate()
  vakod$ = ""
  MSFlexGrid1.Clear
  kdfil = FreeFile
  Open langutvonal$ + "auwin.kod" For Input As #kdfil
  Do
    Line Input #kdfil, scrp$
    If Left(scrp$, 1) <> "*" Then
      Call scpar(scrp$, "\")
      If UCase(spar(1)) = UCase(betukod$) Then
        Text1.Text = spar(2)
        kdara% = xval(spar(3))
        Call flexinit(kdara%)
        For ji31 = 1 To kdara%
          Line Input #kdfil, scrp$
          Call scpar(scrp$, "\")
          MSFlexGrid1.TextMatrix(ji31 - 1, 0) = spar(1)
          MSFlexGrid1.TextMatrix(ji31 - 1, 1) = spar(2)
        Next
      End If
    End If
  Loop While Not EOF(kdfil)
  Close kdfil
End Sub

Private Sub flexinit(sor%)
  MSFlexGrid1.Cols = 2
  MSFlexGrid1.Rows = sor%
  MSFlexGrid1.ColAlignment(0) = 0
  MSFlexGrid1.ColWidth(0) = 700
  MSFlexGrid1.ColWidth(1) = 4540
  MSFlexGrid1.ColAlignment(1) = 0
  MSFlexGrid1.Width = 5292
  MSFlexGrid1.Height = sor% * 235 + 40
  MSFlexGrid1.FixedRows = 0
  Kodok.Height = sor% * 240 + 1300
End Sub

Private Sub Form_Load()
  MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  MSFlexGrid1.Font.Size = 8
  Call langinit("kodok", 2)
  Call szkriptel("kodok")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call Command1_Click
End Sub

Private Sub MSFlexGrid1_DblClick()
  vakod$ = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0))
  Call langclos
  Kodok.Hide
End Sub

Private Sub MSFlexGrid1_entercell()
  'MSFlexGrid1.CellFontName = "Ariel"
  'MSFlexGrid1.CellFontSize = 11
  'MSFlexGrid1.CellFontItalic = True
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  Call MSFlexGrid1_DblClick
End Sub

Private Sub MSFlexGrid1_LeaveCell()
  'MSFlexGrid1.CellFontName = "MS Sans Serif"
  'MSFlexGrid1.CellFontSize = 8
  'MSFlexGrid1.CellFontItalic = False
End Sub

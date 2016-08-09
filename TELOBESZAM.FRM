VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Telobeszam 
   BackColor       =   &H00BFDBD5&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tételes elõlegbeszámítás"
   ClientHeight    =   4656
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   11676
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4656
   ScaleWidth      =   11676
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Beszámítás"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   10.2
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1452
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   10.2
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6360
      TabIndex        =   5
      Top             =   4080
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   10.2
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   3
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mégsem"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   10.2
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1332
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3852
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11412
      _ExtentX        =   20130
      _ExtentY        =   6795
      _Version        =   327680
      BackColorBkg    =   12573653
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Beszámítandó bruttó összeg:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   10.2
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Top             =   4080
      Width           =   2652
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Beszámítandó mennyiség:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   10.2
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   2172
   End
End
Attribute VB_Name = "Telobeszam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  welobeikt = ""
  welobemenny = 0
  welobebrutto = 0
  Telobeszam.Hide
End Sub

Private Sub Command2_Click()
  welobeikt = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0))
  welobemenny = xval(Trim(Text1.Text))
  welobebrutto = xval(Trim(Text2.Text))
  Telobeszam.Hide
End Sub

Private Sub Form_Activate()
  If wujeloleg = 1 Then
    MSFlexGrid1.Enabled = True
    wujeloleg = 0
    Command2.Visible = False
    Label1(0).Visible = False
    Label1(1).Visible = False
    Text1.Text = "": Text1.Visible = False
    Text2.Text = "": Text2.Visible = False
    MSFlexGrid1.Clear
    MSFlexGrid1.Cols = 7
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.FixedCols = 0
    MSFlexGrid1.FixedRows = 1
    MSFlexGrid1.ColAlignment(0) = 1
    MSFlexGrid1.ColWidth(0) = 800
    MSFlexGrid1.TextMatrix(0, 0) = "Iktató"
    MSFlexGrid1.ColAlignment(1) = 1
    MSFlexGrid1.ColWidth(1) = 1000
    MSFlexGrid1.TextMatrix(0, 1) = "Számlaszám"
    MSFlexGrid1.ColAlignment(2) = 1
    MSFlexGrid1.ColWidth(2) = 800
    MSFlexGrid1.TextMatrix(0, 2) = "Telj.kelt"
    MSFlexGrid1.ColAlignment(3) = 1
    MSFlexGrid1.ColWidth(3) = 1000
    MSFlexGrid1.TextMatrix(0, 3) = "Termék kód"
    MSFlexGrid1.ColAlignment(4) = 1
    MSFlexGrid1.ColWidth(4) = 4600
    MSFlexGrid1.TextMatrix(0, 4) = "Megnevezés"
    MSFlexGrid1.ColAlignment(5) = 1
    MSFlexGrid1.ColWidth(5) = 1200
    MSFlexGrid1.TextMatrix(0, 5) = "Besz.menny"
    MSFlexGrid1.ColAlignment(6) = 1
    MSFlexGrid1.ColWidth(6) = 1200
    MSFlexGrid1.TextMatrix(0, 6) = "Besz.összeg"
    fil1 = FreeFile
    Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #fil1
    fil2 = FreeFile
    Open auditorutvonal$ + "auw-pelv.ndx" For Binary Shared As #fil2
    rce& = Int(LOF(fil2) / 12)
    If rce& <> 0 Then
      edb% = 0
      For i41% = 1 To rce&
        Get #fil2, (i41% - 1) * 12 + 1, rcim&
        ere$ = Space(650)
        Get #fil1, rcim& + 9, ere$
        If Mid$(ere$, 90.1) <> "S" Then
          If Mid$(ere$, 23, 15) = Mid$(welobepartrec, 1, 15) And Mid$(ere$, 90, 1) <> "S" And Mid$(ere$, 224, 2) = "EV" Then
            beszhato@ = xval(Mid$(ere$, 44, 14)) - xval(Mid$(ere$, 449, 14))
            beszhme@ = xval(Mid$(ere$, 295, 14)) - xval(Mid$(ere$, 435, 14))
            If beszhato@ > 0 Or beszhme@ > 0 Then
              edb% = edb% + 1
              MSFlexGrid1.Rows = edb% + 1
              MSFlexGrid1.Row = ebd%
              MSFlexGrid1.TextMatrix(edb%, 0) = Mid$(ere$, 1, 7)
              MSFlexGrid1.TextMatrix(edb%, 1) = Mid$(ere$, 8, 15)
              MSFlexGrid1.TextMatrix(edb%, 2) = Mid$(ere$, 38, 6)
              MSFlexGrid1.TextMatrix(edb%, 3) = Mid$(ere$, 266, 15)
              MSFlexGrid1.TextMatrix(edb%, 4) = Trim(Mid$(ere$, 315, 60))
              MSFlexGrid1.TextMatrix(edb%, 5) = Trim(ertszam(Str(beszhme@), 14, 2))
              MSFlexGrid1.TextMatrix(edb%, 6) = Trim(ertszam(Str(beszhato@), 14, 2))
            End If
          End If
        End If
      Next
      Close fil1: Close fi2
      If edb% = 0 Then
        Call mess("A partnernek nincs beszámítható elõlege!", 3, 0, "Figyelmeztetés", valasz%)
        Call Command1_Click
      End If
    Else
      Close fil1: Close fi2
      Call mess("A partnernek nincs beszámítható elõlege!", 3, 0, "Figyelmeztetés", valasz%)
      Call Command1_Click
    End If
  End If
End Sub

Private Sub MSFlexGrid1_Click()
  If xval(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)) > 0 Then
    MSFlexGrid1.Enabled = False
    Command2.Visible = True
    Label1(0).Visible = True
    Text2.Text = ""
    Text1.Visible = True
    Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
    Text1.SelStart = Len(Trim(Text1.Text)) + 1
    Text1.SetFocus
  Else
    If xval(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)) > 0 Then
      MSFlexGrid1.Enabled = False
      Command2.Visible = True
      Text1.Text = ""
      Label1(1).Visible = True
      Text2.Visible = True
      Text2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
      Text2.SelStart = Len(Trim(Text2.Text)) + 1
      Text2.SetFocus
    End If
  End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Call kodvizsg(Trim(Text1.Text), "NJT", khiba%, 14)
    If khiba% <> 0 Then
      Call mess("Hibás mennyiség!", 3, 0, "Hiba", valasz%)
    Else
      If xval(Trim(Text1.Text)) = 0 Or xval(Trim(Text1.Text)) > xval(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))) Then
        Call mess("Hibás mennyiség!", 3, 0, "Hiba", valasz%)
      Else
        Command2.SetFocus
      End If
    End If
  End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Call kodvizsg(Trim(Text1.Text), "NJT", khiba%, 14)
    If khiba% <> 0 Then
      Call mess("Hibás érték!", 3, 0, "Hiba", valasz%)
    Else
      If xval(Trim(Text1.Text)) = 0 Or xval(Trim(Text1.Text)) > xval(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))) Then
        Call mess("Hibás érték!", 3, 0, "Hiba", valasz%)
      Else
        Command2.SetFocus
      End If
    End If
  End If
End Sub


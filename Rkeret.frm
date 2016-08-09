VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Rkeret 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0055D7F7&
   ClientHeight    =   7116
   ClientLeft      =   48
   ClientTop       =   48
   ClientWidth     =   12120
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7116
   ScaleWidth      =   12120
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sor törlés"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rögzítés"
      Height          =   492
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
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
      Left            =   4320
      TabIndex        =   5
      Top             =   6720
      Width           =   1212
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4812
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   11892
      _ExtentX        =   20976
      _ExtentY        =   8488
      _Version        =   327680
      BackColorSel    =   12648384
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      Redraw          =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      FillStyle       =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mégsem"
      Height          =   492
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   972
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Termék"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   732
      Left            =   120
      TabIndex        =   3
      Top             =   910
      Width           =   11892
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Szállító"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   480
      Left            =   120
      TabIndex        =   2
      Top             =   400
      Width           =   11892
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keretrendelés és napi módosítás"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   3504
   End
End
Attribute VB_Name = "Rkeret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim utermkod$, billscr%
Public betoltve%, ipartkod$

Private Sub Command1_Click()
  rogzites% = 0
  Rkeret.Hide
End Sub

Private Sub Command2_Click()
  rogzites% = 1
  Rkeret.Hide
End Sub

Private Sub Command3_Click()
  If MSFlexGrid1.Row < 200 Then
    For i77% = MSFlexGrid1.Row To 199
      For i78% = 1 To 15
        MSFlexGrid1.TextMatrix(i77%, i78%) = MSFlexGrid1.TextMatrix(i77% + 1, i78%)
      Next
    Next
  End If
  For i78% = 1 To 15
    MSFlexGrid1.TextMatrix(200, i78%) = ""
  Next
  Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
  Text1.SelStart = Len(Trim(Text1.Text))
  Text1.SetFocus
End Sub

Private Sub Form_Activate()
  If betoltve% = 0 Then
    billscr% = 0
    betoltve% = 1
    MSFlexGrid1.SetFocus
  End If
End Sub

Private Sub MSFlexGrid1_gotfocus()
  Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
  Text1.Height = MSFlexGrid1.CellHeight
  Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
  Text1.Width = MSFlexGrid1.CellWidth
  Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
  Text1.SelStart = Len(Trim(Text1.Text))
  termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
  If termkod$ <> utermkod$ Then
    utermkod$ = termkod$
    If Trim(termkod$) = "" Then
      Label3.Caption = ""
    Else
      ktrmrec$ = dbxkey("KTRM", termkod$)
      If ktrmrec$ = "" Then
        Label3.Caption = ""
      Else
        egysar@ = xval(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 16))
        krakod$ = Trim(form1.Text5.Text)
        If krakod$ = "" Then
          keszme@ = 0
        Else
          Call rkeszlet(termkod$, "", krakod$, keszme@, foglme@)
        End If
        Label3.Caption = " " + Trim(Mid$(ktrmrec$, 16, 60)) + Chr$(13) + "  " + Trim(Mid$(ktrmrec$, 196, 60)) + Chr$(13) + "  " + "Mértékegység: " + Mid$(ktrmrec$, 484, 6) + "   Rendelhetõ: " + Trim(Mid$(ktrmrec$, 542, 10)) + "   Nettó eladási ár:" + ertszam(Str(egysar@), 12, 2) + "   Készlet:" + ertszam(Str(keszme@), 12, 2)
      End If
    End If
  End If
  Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_Scroll()
  If billscr% = 0 Then
    MSFlexGrid1.Row = MSFlexGrid1.toprow
    MSFlexGrid1.Col = MSFlexGrid1.LeftCol
  End If
  Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
  Text1.Height = MSFlexGrid1.CellHeight
  Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
  Text1.Width = MSFlexGrid1.CellWidth
  Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
  Text1.SelStart = Len(Trim(Text1.Text))
  billscr% = 0
End Sub

Private Sub Text1_Click()
  Text1.SelStart = Len(Trim(Text1.Text))
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  billscr% = 0
  If KeyCode = vbKeyReturn Then
    '--- mezõ ellenõrzés
    khiba% = 0
    If MSFlexGrid1.Col = 1 Then
      '--- termék kód
      ktrmkod$ = Left(Trim(Text1.Text) + Space(15), 15)
      ktrmrec$ = dbxkey("KTRM", ktrmkod$)
      If ktrmrec$ = "" Then
        Call mess("Hibás termék kód!", 2, 0, "Hiba", valasz%)
        khiba% = 1
      Else
        For i79% = 1 To 200
          If i79% <> MSFlexGrid1.Row And ktrmkod$ = Left(Trim(MSFlexGrid1.TextMatrix(i79%, 1)) + Space(15), 15) Then
            khiba% = 1
            Call mess("A megadott termék egy másik sorban már szerepel!", 2, 0, "Hiba", valasz%)
            Exit For
          End If
        Next
        If khiba% = 0 Then
          egysar@ = arazo(ipartkod$, ktrmkod$, "", maidatum$, "")
          krakod$ = Trim(form1.Text5.Text)
          If krakod$ = "" Then
            keszme@ = 0
          Else
            Call rkeszlet(ktrmkod$, "", krakod$, keszme@, foglme@)
          End If
          Label3.Caption = " " + Trim(Mid$(ktrmrec$, 16, 60)) + Chr$(13) + " " + Trim(Mid$(ktrmrec$, 196, 60)) + Chr$(13) + "  " + "Mértékegység: " + Mid$(ktrmrec$, 484, 6) + "   Rendelhetõ: " + Trim(Mid$(ktrmrec$, 542, 10)) + "   Nettó eladási ár:" + ertszam(Str(egysar@), 12, 2) + "   Készlet:" + ertszam(Str(keszme@), 12, 2)
          MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 16) = egysar@
        End If
      End If
    Else
      '--- egyéb mezõk
      mezo$ = Trim(Text1.Text)
      If MSFlexGrid1.Col Mod 2 = 1 Then
        Call kodvizsg(mezo$, "NJT-", khiba%, 10)
      Else
        Call kodvizsg(mezo$, "NJT", khiba%, 10)
      End If
      If khiba% <> 0 Then
        Call mess("Hibás mennyiség!", 2, 0, "Hiba", valasz%)
      End If
      If khiba% = 0 And MSFlexGrid1.Col Mod 2 = 1 Then
        ere@ = xval(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col - 1)))
        csok@ = xval(Text1.Text)
        If ere@ + csok@ < 0 Then
          Call mess("Negatív mennyiség nem szállítható ki!", 2, 0, "Hiba", valasz%)
          khiba% = 1
        End If
      End If
      If khiba% = 0 And xval(Text1.Text) <> 0 Then
        If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) = "" Then
          Call mess("Termék kód kötelezõ!", 2, 0, "Hiba", valasz%)
          khiba% = 1
        End If
      End If
    End If
    If khiba% = 0 Then
      If MSFlexGrid1.Col <> 1 Then
        If MSFlexGrid1.Col Mod 2 = 0 Then
          MSFlexGrid1.CellForeColor = RGB(0, 60, 40)
        Else
          MSFlexGrid1.CellForeColor = RGB(110, 0, 0)
        End If
      Else
        MSFlexGrid1.CellForeColor = 0
      End If
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col) = Text1.Text
      If MSFlexGrid1.Col < 16 Then
        billscr% = 1
        MSFlexGrid1.Col = MSFlexGrid1.Col + 1
      Else
        If MSFlexGrid1.Row < 200 Then
          billscr% = 1
          MSFlexGrid1.Row = MSFlexGrid1.Row + 1
          MSFlexGrid1.Col = 1
        End If
      End If
      KeyCode = 0
      MSFlexGrid1.SetFocus
    End If
  Else
    Select Case KeyCode
      Case vbKeyDelete
        Text1.Text = ""
      Case vbKeyX, vbKeyH
        If Shift And vbAltMask Then
          oszlop% = MSFlexGrid1.Col
          If oszlop% = 1 Then
            If KeyCode = vbKeyX Then
              Call altx("KTRM", azonosito$)
            Else
              Call alth("KTRM", azonosito$)
            End If
            If azonosito$ <> "" Then
              Text1.Text = azonosito$
            End If
          End If
        End If
      Case vbKeyHome
        billscr% = 1
        MSFlexGrid1.Row = 1: MSFlexGrid1.Col = 1: KeyCode = 0: MSFlexGrid1.SetFocus
      Case vbKeyEnd
        billscr% = 1
        MSFlexGrid1.Row = 200: MSFlexGrid1.Col = 1: KeyCode = 0: MSFlexGrid1.SetFocus
      Case vbKeyPageDown
        billscr% = 1
        If MSFlexGrid1.Row + 17 <= 200 Then
          MSFlexGrid1.Row = MSFlexGrid1.Row + 17: KeyCode = 0: MSFlexGrid1.SetFocus
        Else
          MSFlexGrid1.Row = 200: KeyCode = 0: MSFlexGrid1.SetFocus
        End If
      Case vbKeyPageUp
        billscr% = 1
        If MSFlexGrid1.Row - 17 >= 1 Then
          MSFlexGrid1.Row = MSFlexGrid1.Row - 17: KeyCode = 0: MSFlexGrid1.SetFocus
        Else
          MSFlexGrid1.Row = 1: KeyCode = 0: MSFlexGrid1.SetFocus
        End If
      Case vbKeyUp
        billscr% = 1
        If MSFlexGrid1.Row > 1 Then MSFlexGrid1.Row = MSFlexGrid1.Row - 1: KeyCode = 0: MSFlexGrid1.SetFocus
      Case vbKeyDown
        billscr% = 1
        If MSFlexGrid1.Row < 200 Then MSFlexGrid1.Row = MSFlexGrid1.Row + 1: KeyCode = 0: MSFlexGrid1.SetFocus
      Case vbKeyLeft
        billscr% = 1
        If MSFlexGrid1.Col > 1 Then MSFlexGrid1.Col = MSFlexGrid1.Col - 1: KeyCode = 0: MSFlexGrid1.SetFocus
      Case vbKeyRight
        billscr% = 1
        If MSFlexGrid1.Col < 16 Then MSFlexGrid1.Col = MSFlexGrid1.Col + 1: KeyCode = 0: MSFlexGrid1.SetFocus
      Case Else
    End Select
  End If
End Sub

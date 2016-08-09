VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Vektor 
   BackColor       =   &H0096E2FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sima ablak"
   ClientHeight    =   5148
   ClientLeft      =   1848
   ClientTop       =   4152
   ClientWidth     =   5136
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   7.8
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
   ScaleHeight     =   5148
   ScaleWidth      =   5136
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H0096E2FF&
      Caption         =   "Ciril betû"
      Height          =   192
      Left            =   3720
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0096E2FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   450
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0091E9FB&
      Caption         =   "Alt+X"
      Height          =   372
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   850
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   6
      Text            =   "Text2"
      ToolTipText     =   "Az aktuális adatmezõ"
      Top             =   3840
      Width           =   3492
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0091E9FB&
      BorderStyle     =   0  'None
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
      IMEMode         =   3  'DISABLE
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   2772
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mégsem"
      Height          =   372
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Kilépés a tartalom rögzítése nélkül (Ctrl+T)"
      Top             =   4680
      Width           =   1692
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Kilépés a tartalom rögzítésével (Esc)"
      Top             =   4680
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nincs kapcsolat"
      Height          =   372
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "A kapcsolódó objektum táblázata (Alt+H)"
      Top             =   4200
      Width           =   3492
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3612
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   6371
      _Version        =   327680
      BackColor       =   16777215
      BackColorFixed  =   14745086
      BackColorSel    =   -2147483639
      BackColorBkg    =   -2147483643
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12
      Left            =   1080
      TabIndex        =   5
      Top             =   360
      Width           =   12
   End
End
Attribute VB_Name = "Vektor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- VEKTOR form vektoros adatbevitel
Public texelo$, backtext$, ujcar%, com1wid%
Dim kikapcsol%

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Text1.Font.Name = "ER Kurier 1251"
Else
   Text1.Font.Name = "Microsoft Sans Serif"
End If
End Sub

Private Sub Command2_Click()
  '--- Kilépés az adatok rögzítésével
  For i1& = 1 To MSFlexGrid1.Rows - 2
    ar$ = RTrim$(ADATAB(mesor%(i1&)).attr)
    amez$ = MSFlexGrid1.TextMatrix(i1&, 1)
    If InStr(ar$, "*") > 0 Then
      If Trim$(amez$) = "" Then
        nx$ = RTrim$(ADATAB(mesor%(i1&)).adatnev)
        Call mess(nx$ + " " + langform(9), 2, 0, langform(10), valasz%)
        'MsgBox nx$ + " " + langform(9), 48, langform(10)
        Exit Sub
      End If
    End If
  Next
  rogzites% = 1
  form1.Label2.Caption = ""
  form1.Label3.Caption = ""
  Call langclos
  Vektor.Hide
End Sub
Private Sub Command1_Click()
  '--- Alt+H kapcsolat
  sor% = MSFlexGrid1.Row
  w5& = ADATAB(mesor%(sor%)).kapcsob
  If w5& > 0 Then
    objazon1$ = OBJTAB(w5&).obaz
    azonosito$ = ""
    Call alth(objazon1$, azonosito$)
    If azonosito$ <> "" Then
      MSFlexGrid1.Text = azonosito$
      Text1.Text = Trim$(azonosito$)
    End If
  End If
  Text1.SetFocus
End Sub

Private Sub Command3_Click()
  '--- Kilépés rögzítés nélkül
  rogzites% = 0
  form1.Label2.Caption = ""
  form1.Label3.Caption = ""
  Call langclos
  Vektor.Hide
End Sub

Private Sub Command4_Click()
  sor% = MSFlexGrid1.Row
  w5& = ADATAB(mesor%(sor%)).kapcsob
  If w5& > 0 Then
    objazon1$ = OBJTAB(w5&).obaz
    azonosito$ = ""
    Call altx(objazon1$, azonosito$)
    If azonosito$ <> "" Then
      MSFlexGrid1.Text = azonosito$
      Text1.Text = Trim$(azonosito$)
    End If
  End If
  Text1.SetFocus
End Sub

Private Sub Form_Activate()
  Command1.Width = com1wid
  'com1wid = Command1.Width
  If Check1.Value = 1 Then
   Text1.Font.Name = "ER Kurier 1251"
  Else
   Text1.Font.Name = "Microsoft Sans Serif"
  End If

End Sub


Private Sub Form_Load()
  'com1wid = Command1.Width
  Call langinit("vektor", 2)
  Call szkriptel("vektor")
  'For i17% = 0 To 1
  '  MSFlexGrid1.Row = 0: MSFlexGrid1.Col = i17%
  '  MSFlexGrid1.CellBackColor = RGB(150, 20, 0)
  '  MSFlexGrid1.CellForeColor = RGB(255, 255, 255)
  'Next
End Sub

Private Sub MSFlexGrid1_DblClick()
  sor% = MSFlexGrid1.Row
  adatazonosito$ = Trim$(ADATAB(mesor%(sor%)).adkod)
  vakod$ = ""
  Call kodtablak(adatazonosito$, vakod$)
  If vakod$ <> "" Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = vakod$
  End If
End Sub

Private Sub MSFlexGrid1_entercell()
  '--- Belépés egy cellába
  If kikapcsol% = 1 Then Exit Sub
  sor% = MSFlexGrid1.Row
  ujcar% = 0
  w5& = ADATAB(mesor%(sor%)).kapcsob
  If w5& <> 0 Then
    Command1.Font.Size = 9
    Command1.BackColor = RGB(240, 220, 85)
    Command1.Caption = Trim(OBJTAB(w5&).obnev) + " (Alt+H)"
    If OBJTAB(w5&).hashcod <> 0 Then
      Command4.Top = Command1.Top
      Command4.Left = Command1.Left + com1wid% - 850
      Command4.Visible = True
      Command1.Width = com1wid% - 860
    Else
      Command1.Width = com1wid%
      Command4.Visible = False
    End If
  Else
    Command1.Font.Size = 8
    Command1.BackColor = RGB(240, 240, 240)
    Command1.Caption = langform(4)
    Command1.Width = com1wid%
    Command4.Visible = False
  End If
  If InStr(mtb(sor%), "*") > 0 Then
    form1.Label3.ForeColor = &HC0&
    capelo$ = langform(12) + " "
    Text1.Locked = False
  Else
    form1.Label3.ForeColor = &H808080
    capelo$ = ""
    Text1.Locked = False
  End If
  If InStr(mtb(sor%), "W") > 0 Then
    form1.Label3.ForeColor = &HC0&
    capelo$ = langform(11) + " "
    Text1.Locked = True
  End If
  form1.Label2.Caption = MAGYARAZAT$(mesor%(sor%))
  If InStr(mtb(sor%), "Q") > 0 Then
    form1.Label3.ForeColor = RGB(0, 0, 120)
    form1.Label3.Caption = capelo$ + langform(18)
  Else
    form1.Label3.Caption = capelo$ + Mid$(ELLENORZO$(mesor%(sor%)), 3)
  End If
  Text1.MaxLength = mho%(sor%)
  Text2.Text = MSFlexGrid1.TextMatrix(sor%, 0) + ": " + Text1.Text
  If InStr(mtb(sor%), "D") > 0 And langhun% > 1 Then
    Text1.Text = datfor(Trim$(MSFlexGrid1.Text))
  Else
    Text1.Text = Trim$(MSFlexGrid1.Text)
  End If
  Text1.SelStart = Len(Text1.Text) + 1
  Vektor.texelo = Text1.Text
  MSFlexGrid1.CellBackColor = QBColor(14)
  If InStr(mtb$(sor%), "J") > 0 And InStr(mtb$(sor%), "NZJ") = 0 Then
    MSFlexGrid1.CellAlignment = 1
  Else
    MSFlexGrid1.CellAlignment = 1
  End If
  Text1.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
  Text1.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft + 25
  Text1.Height = MSFlexGrid1.CellHeight - 50
  Text1.Width = MSFlexGrid1.CellWidth - 25
  If InStr(mtb(sor%), "NJ") Then
    If Trim(Text1.Text) <> "" Then
      munimo@ = xval(Text1.Text)
      Text3.Left = 1
      Text3.Width = Vektor.Width
      Text3.Top = Command1.Top - 70
      Text3.Visible = True
      Text3.Text = szamszoveg(munimo@, 0, "")
    Else
      Text3.Visible = False
    End If
  Else
    Text3.Visible = False
  End If
  If InStr(mtb$(sor%), "X") > 0 Then
    w5& = ADATAB(mesor%(sor%)).kapcsob
    If w5& <> 0 Then
      objazon1$ = OBJTAB(w5&).obaz
      azonosito$ = Text1.Text
      If InStr(mtb$(MSFlexGrid1.Row), "J") > 0 And InStr(mtb$(MSFlexGrid1.Row), "NZJ") = 0 Then
        azonosito$ = Right$(Space$(14) + azonosito$, mho%(MSFlexGrid1.Row))
      Else
        azonosito$ = Left$(azonosito$ + Space$(mho%(MSFlexGrid1.Row)), mho%(MSFlexGrid1.Row))
      End If
      If azonosito$ <> Space$(mho%(MSFlexGrid1.Row)) Then
        Call rekinfo(objazon1, azonosito$)
      Else
        '--- 071108
        'form1.Info.Visible = False
        'form1.Text1.Visible = False
      End If
    End If
  Else
    '--- 071108
    'form1.Info.Visible = False
    'form1.Text1.Visible = False
  End If
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_gotfocus()
  '--- Táblázatra fókuszálás
  If kikapcsol% = 1 Then Exit Sub
  Memo& = MSFlexGrid1.Row
  For i6% = 1 To MSFlexGrid1.Rows - 2
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = i6%
  Next
  MSFlexGrid1.Row = Memo&
  sor% = MSFlexGrid1.Row
  If mesor%(1) = foindex% Then
    If MSFlexGrid1.toprow = 1 And MSFlexGrid1.Row = 1 Then
      MSFlexGrid1.Row = 2: sor% = 2
    End If
  End If
  If programnev$ = "AUW-RLEL" Then
     If MSFlexGrid1.toprow = 1 And MSFlexGrid1.Row = 1 Then
        MSFlexGrid1.Row = 6: sor% = 6
     End If
  End If

  
  ujcar% = 0
  w5& = ADATAB(mesor%(sor%)).kapcsob
  If w5& <> 0 Then
    Command1.Font.Size = 9
    Command1.BackColor = RGB(240, 220, 85)
    Command1.Caption = Trim(OBJTAB(w5&).obnev) + " (Alt+H)"
    If OBJTAB(w5&).hashcod <> 0 Then
      Command4.Top = Command1.Top
      Command4.Left = Command1.Left + com1wid% - 850
      Command4.Visible = True
      Command1.Width = com1wid% - 860
    Else
      Command1.Width = com1wid%
      Command4.Visible = False
    End If
  Else
    Command1.Font.Size = 8
    Command1.BackColor = RGB(240, 240, 240)
    Command1.Caption = langform(4)
    Command1.Width = com1wid%
    Command4.Visible = False
  End If
  Text1.MaxLength = mho%(sor%)
  If InStr(mtb(sor%), "D") > 0 And langhun% > 1 Then
    Text1.Text = datfor(Trim(MSFlexGrid1.Text))
  Else
    Text1.Text = Trim(MSFlexGrid1.Text)
  End If
  'If InStr(mtb(sor%), "W") > 0 Then
  '  Text1.SelStart = 0
  'Else
    Text1.SelStart = Len(Text1.Text) + 1
  'End If
  Vektor.texelo = Text1.Text
  MSFlexGrid1.CellBackColor = QBColor(14)
  If MSFlexGrid1.RowIsVisible(MSFlexGrid1.Rows - 2) = True Then MSFlexGrid1.ScrollBars = flexScrollBarNone
  Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_LeaveCell()
  '--- Cella elhagyása
  Text1.Locked = False
  If InStr(mtb$(MSFlexGrid1.Row), "NJ") > 0 Then
    txx$ = Trim(MSFlexGrid1.Text)
    tw% = TextWidth(String(mho%(MSFlexGrid1.Row), "0"))
    tw2% = TextWidth(" ")
    tw3% = TextWidth(txx$)
    twsp% = (tw% - tw3%) / tw2%
    kikapcsol% = 1
    MSFlexGrid1.Text = Space$(twsp%) + txx$
    kikapcsol% = 0
    MSFlexGrid1.CellBackColor = &HF4F4F4
  Else
    MSFlexGrid1.CellBackColor = QBColor(15)
  End If
End Sub

Private Sub Text1_Change()
  '--- Adat változás
  Dim param$(30)
  sor% = MSFlexGrid1.Row
  'If InStr(mtb$(sor%), "W") > 0 Then Exit Sub
  Text2.Text = MSFlexGrid1.TextMatrix(sor%, 0) + ": " + Text1.Text
  If InStr(mtb(sor%), "NJ") Then
    If Trim(Text1.Text) <> "" Then
      munimo@ = xval(Text1.Text)
      Text3.Left = 1
      Text3.Top = Command1.Top - 70
      Text3.Width = Vektor.Width
      Text3.Visible = True
      Text3.Text = szamszoveg(munimo@, 0, "")
    Else
      Text3.Visible = False
    End If
  Else
    Text3.Visible = False
  End If
  If ujcar% <> 0 And Text1.SelStart > 0 Then
    pont% = 0
    minusz% = 0
    h% = Len(Text1.Text)
    For j% = 1 To h%
      If Mid$(Text1.Text, j%, 1) = "." Or Mid$(Text1.Text, j%, 1) = "," Then pont% = pont% + 1
      If Mid$(Text1.Text, j%, 1) = "-" Then minusz% = minusz% + 1
    Next
    texcar$ = Mid$(Text1.Text, Text1.SelStart, 1)
    texhiba% = 0
    If InStr(mtb$(sor%), "N") > 0 Then
      '--- numerikus adatok vizsgálata
      If texcar$ = "-" Then
        If InStr(mtb(sor%), "-") > 0 And (Text1.SelStart <> 1 Or minusz% > 1) Then texhiba% = 1
      Else
        If texcar$ = "." Or texcar$ = "," Then
          If InStr(mtb(sor%), "T") > 0 And pont% > 1 Then texhiba% = 1
          If InStr(mtb(sor%), "T") And texcar$ = "," Then
            texcar$ = "."
          End If
        Else
          If texcar$ < "0" Or texcar$ > "9" Then texhiba% = 1
        End If
      End If
    End If
    If InStr(mtb$(sor%), "D") > 0 Then
      '--- dátum adat vizsgálata
      w1$ = "000000"
      Mid$(w1$, 1) = Text1.Text
      hon% = xval(Mid$(w1$, 3, 2))
      If hon% = 4 Or hon% = 6 Or hon% = 9 Or hon% = 11 Then nap% = 30 Else nap% = 31
      If hon% = 2 Then nap% = 29
      If hon% > 12 Then texhiba% = 1
      If langhun% > 1 Then
        If xval(Mid$(w1$, 1, 2)) > nap% Then texhiba% = 1
      Else
        If xval(Mid$(w1$, 5, 2)) > nap% Then texhiba% = 1
      End If
    End If
    '--- ellenõrzõ string vizsgálata
    ellst$ = Trim$(ELLENORZO$(mesor%(sor%)))
    If ellst$ <> "" Then
      elh% = Len(Text1.Text)
      ellko$ = Left$(ellst$, 2)
      ellst$ = Mid$(ellst$, 3)
      elldb% = 30
      Call linpar(ellst$, param$(), "\", elldb%)
      elhiba% = 1
      For eli1% = 1 To elldb%
        If Mid$(param$(eli1%), 1, elh%) = UCase$(Text1.Text) Then
          elhiba% = 0: Exit For
        End If
      Next
      If elhiba% = 1 Then texhiba% = 1
    End If
    'If InStr(mtb$(sor%), "W") > 0 Then texhiba% = 1
    If texhiba% > 0 Then
      '--- hibás karakter törlése
      If Len(Text1.Text) > 1 Then
        px% = Text1.SelStart
        ujcar% = 0
        Text1.Text = backtext
        Text1.SelStart = px% - 1
      Else
        Text1.Text = ""
      End If
    End If
  End If
  If InStr(mtb$(sor%), "U") > 0 Then
    MSFlexGrid1.Text = UCase$(Text1.Text)
  Else
    If InStr(mtb$(sor%), "D") > 0 And langhun% > 1 Then
      MSFlexGrid1.Text = datfor(Text1.Text)
    Else
      MSFlexGrid1.Text = Text1.Text
    End If
  End If
  If InStr(mtb$(sor%), "D") > 0 And langhun% > 1 Then
    backtext$ = datfor(MSFlexGrid1.Text)
  Else
    backtext$ = MSFlexGrid1.Text
  End If
End Sub

Private Sub Text1_GotFocus()
  '--- Adatbeviteli mezõre fókuszálás
  If InStr(mtb$(MSFlexGrid1.Row), "D") > 0 And langhun% > 1 Then
    Text1.Text = datfor(Trim$(MSFlexGrid1.Text))
  Else
    Text1.Text = Trim$(MSFlexGrid1.Text)
  End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- Billentyûlenyomás adatbeviteli mezõben
  ujcar% = 0
  sor = MSFlexGrid1.Row
  If InStr(mtb(sor), "W") Then
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    Else
      If InStr(mtb(sor), "D") And langhun% > 1 Then
        Text1.Text = datfor(MSFlexGrid1.TextMatrix(sor, 1))
      Else
        Text1.Text = MSFlexGrid1.TextMatrix(sor, 1)
      End If
      KeyCode = 0
      Exit Sub
    End If
  End If
  sorok = MSFlexGrid1.Rows - 2
  If KeyCode = vbKeyLeft And Text1.SelStart = 0 Or KeyCode = vbKeyRight And Text1.SelStart = Len(Text1.Text) Then
    If InStr(mtb(sor), "D") And langhun% > 1 Then
      MSFlexGrid1.Text = datfor(Vektor.texelo)
    Else
      MSFlexGrid1.Text = Vektor.texelo
    End If
    Text1.Text = Vektor.texelo
  End If
  If KeyCode = vbKeyEscape Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    If InStr(mtb(sor), "D") And langhun% > 1 Then
      MSFlexGrid1.Text = datfor(Vektor.texelo)
    Else
      MSFlexGrid1.Text = Vektor.texelo
    End If
    Text1.Text = Vektor.texelo
  End If
  Select Case KeyCode
    Case vbKeyBack
    Case vbKeyDelete
      '--- adatmezõ törlése (Del)
      Text1.Text = ""
    Case vbKeyF1
      If Shift And vbAltMask Then
        rre = Shell(programutvonal$ + "auw-gyi " + terminal$ + task$, vbNormalFocus)
      End If
    Case vbKeyInsert
      If Shift And vbAltMask Then
        sors% = MSFlexGrid1.Row
        w5& = ADATAB(mesor%(sors%)).kapcsob
        If w5& > 0 And InStr(mtb$(MSFlexGrid1.Row), "Q") > 0 Then
          objazon1$ = OBJTAB(w5&).obaz
          '--- ide jön a beléptetõ program hívása
          rreobj$ = OBJTAB(ADATAB(mesor%(sors%)).kapcsob).obaz
          rrepar$ = dbxneve(rreobj$) + "/" + rreobj$ + "/" + terminal$ + task$ + "/" + auditorutvonal$
          rre = Shell(programutvonal$ + "DBX4-NEW.EXE " + rrepar$, vbNormalFocus)
          Call mess(langmodul(162), 4, 0, langmodul(163), valasz%)
          pxfx = FreeFile
          Open programutvonal$ + terminal$ + task$ + "new.txt" For Binary Shared As pxfx
          csx$ = Space(3)
          Get #pxfx, 1, csx$
          If csx$ <> "NIX" Then
            csx$ = Space(ADATAB(mesor%(sors%)).adatho)
            Get #pxfx, 1, csx$
            Text1.Text = csx$
          End If
          Close pxfx
        End If
        Text1.SetFocus
      End If
    Case vbKeyX
      If Shift And vbAltMask Then
        sors% = MSFlexGrid1.Row
        w5& = ADATAB(mesor%(sors%)).kapcsob
        If w5& > 0 Then
          objazon1$ = OBJTAB(w5&).obaz
          azonosito$ = ""
          Call altx(objazon1$, azonosito$)
          If azonosito$ <> "" Then
            MSFlexGrid1.Text = azonosito$
            Text1.Text = Trim$(azonosito$)
          End If
        End If
        Text1.SetFocus
      Else
        ujcar% = 1
      End If
    Case vbKeyH
      '--- Alt+H kapcsolat
      If Shift And vbAltMask Then
        Call Command1_Click
      Else
        ujcar% = 1
      End If
    Case vbKeyT
      '--- Ctrl+T kilépés rögzítés nélkül
      If Shift And vbCtrlMask Then Call Command3_Click: Exit Sub Else ujcar% = 1
    Case vbKeyI
      '--- Alt+I kapcsolt rekord tartalmának megjelenítés listboxban
      If Shift And vbAltMask Then
        sorox% = MSFlexGrid1.Row
        w5& = ADATAB(mesor%(sorox%)).kapcsob
        If w5& > 0 Then
          objazon1$ = OBJTAB(w5&).obaz
          azonosito$ = Text1.Text
          If InStr(mtb$(MSFlexGrid1.Row), "J") > 0 And InStr(mtb$(MSFlexGrid1.Row), "NZJ") = 0 Then
            azonosito$ = Right$(Space$(14) + azonosito$, mho%(MSFlexGrid1.Row))
          Else
            azonosito$ = Left$(azonosito$ + Space$(mho%(MSFlexGrid1.Row)), mho%(MSFlexGrid1.Row))
          End If
          Call rekinfo(objazon1$, azonosito$)
        End If
      Else
        ujcar% = 1
      End If
    Case vbKeyDown
      '--- következõ mezõ
      If sor < sorok Then MSFlexGrid1.Row = MSFlexGrid1.Row + 1 Else MSFlexGrid1.Row = 1: MSFlexGrid1.toprow = 1
      KeyCode = 0
    Case vbKeyUp
      '--- elõzõ mezõ
      If sor > 1 Then MSFlexGrid1.Row = MSFlexGrid1.Row - 1
      Text1.SelStart = Text1.SelStart + 1
      KeyCode = 0
    Case vbKeyEscape
      '--- kilépés rögzítéssel
      Call Command2_Click
      Exit Sub
    Case vbKeyF12
      If MSFlexGrid1.Row > 1 Then
        If InStr(mtb$(MSFlexGrid1.Row), "[") > 0 Then xax1$ = Mid$(mtb(MSFlexGrid1.Row), 1, InStr(mtb$(MSFlexGrid1.Row), "[") - 1) Else xax1$ = mtb$(MSFlexGrid1.Row)
        If InStr(mtb$(MSFlexGrid1.Row - 1), "[") > 0 Then xax2$ = Mid$(mtb(MSFlexGrid1.Row), 1, InStr(mtb$(MSFlexGrid1.Row - 1), "[") - 1) Else xax2$ = mtb$(MSFlexGrid1.Row - 1)
        If xax1$ = xax2$ Then
          If InStr(mtb$(MSFlexGrid1.Row), "D") > 0 And langhun% > 1 Then
            Text1.Text = datfor(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 1))
          Else
            Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 1)
          End If
        End If
      End If
    Case vbKeyReturn
      '--- teljes mezo ellenorzése
      If InStr(mtb$(MSFlexGrid1.Row), "NZJ") > 0 Then
        txx$ = Right$("000000000000000" + LTrim$(Text1.Text), mho%(MSFlexGrid1.Row))
        Text1.Text = txx$
        MSFlexGrid1.Text = txx$
      End If
      '--- kapcsolódó objektum ellenõrzése, ha van
      soro% = MSFlexGrid1.Row
      w5& = ADATAB(mesor%(soro%)).kapcsob
      If w5& <> 0 Then
        objazon1$ = OBJTAB(w5&).obaz
        objazon1$ = OBJTAB(w5&).obaz
        If objazon1$ = "KTRM" Then
          atex$ = Trim(Text1.Text)
          atexh% = Len(atex$)
          If atexh% > 0 Then
            For iaa% = 1 To atexh%
              If Mid$(atex$, iaa%, 1) = "ö" Then Mid$(atex$, iaa%, 1) = "0"
            Next
            btex$ = Left(Trim(atex$) + Space(15), 13)
            reanrec$ = dbxkey("REAN", btex$)
            If reanrec$ <> "" Then
              Text1.Text = Mid$(reanrec$, 14, 15)
            Else
              Text1.Text = atex$
            End If
          End If
        End If
        azonosito$ = Text1.Text
        If InStr(mtb$(MSFlexGrid1.Row), "J") > 0 And InStr(mtb$(MSFlexGrid1.Row), "NZJ") = 0 Then
          azonosito$ = Right$(Space$(14) + azonosito$, mho%(MSFlexGrid1.Row))
        Else
          azonosito$ = Left$(azonosito$ + Space$(mho%(MSFlexGrid1.Row)), mho%(MSFlexGrid1.Row))
        End If
        If azonosito$ <> Space$(mho%(MSFlexGrid1.Row)) Then
          w6$ = dbxkey$(objazon1$, azonosito$)
          If w6$ = "" And InStr(mtb$(MSFlexGrid1.Row), "0") = 0 Then
            w7$ = OBJTAB(w5&).obnev
            '--- nincs ilyen
            Call mess(langform(13) + " " + Trim(w7$) + "!", 2, 0, langform(14), valasz%)
            MSFlexGrid1.Text = Vektor.texelo
            Text1.Text = Vektor.texelo
            KeyCode = 0
          Else
            w6$ = ""
            If InStr(mtb$(MSFlexGrid1.Row), "X") > 0 Then
              Call rekinfo(objazon1, azonosito$)
            End If
          End If
        End If
      End If
      '--- ide betenni
      If KeyCode <> 0 Then
        If InStr(mtb$(MSFlexGrid1.Row), "K") > 0 And Len(Text1.Text) <> mho%(MSFlexGrid1.Row) And Trim(Text1.Text) <> "" Then
          KeyCode = 0
        Else
          Call mezovege(MSFlexGrid1.Row, 0, mezohiba%, 1)
          If mezohiba% = 0 Then
            Call informa
            If sor < sorok Then
              MSFlexGrid1.Row = MSFlexGrid1.Row + 1
              If InStr(mtb(MSFlexGrid1.Row), "D") > 0 And langhun% > 1 Then
                Text1.Text = datfor(Trim(MSFlexGrid1.Text))
              Else
                Text1.Text = Trim(MSFlexGrid1.Text)
              End If
            Else
              MSFlexGrid1.Row = 1: MSFlexGrid1.toprow = 1
            End If
          Else
            KeyCode = 0
          End If
        End If
      End If
      KeyCode = 0
    Case Else
      '--- rögzített normál karakter
      ujcar% = 1
  End Select
  If mesor%(1) = foindex% And programnev$ <> "DBX4-SET" Then
    If MSFlexGrid1.toprow = 1 And MSFlexGrid1.Row = 1 Then MSFlexGrid1.Row = 2
  End If
  If MSFlexGrid1.Row < MSFlexGrid1.toprow Then
    MSFlexGrid1.toprow = MSFlexGrid1.toprow - 1
  Else
    a& = MSFlexGrid1.Row: lep% = 0
    If a& < sorok Then a& = a& + 1
    If MSFlexGrid1.RowIsVisible(a&) = False Then
      MSFlexGrid1.toprow = MSFlexGrid1.toprow + 1
    End If
  End If
End Sub


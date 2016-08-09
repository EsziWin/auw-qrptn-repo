VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Tabla 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adatbeviteli táblázat"
   ClientHeight    =   3612
   ClientLeft      =   4308
   ClientTop       =   4908
   ClientWidth     =   7968
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3612
   ScaleWidth      =   7968
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      BackColor       =   &H0091E9FB&
      Caption         =   "Alt+X"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C5FCC8&
      Caption         =   "Sor törlés"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "A kijelölt sor törlése (Ctrl+S)"
      Top             =   2760
      Width           =   1212
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   5520
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   2760
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   5
      Text            =   "Text2"
      ToolTipText     =   "Az aktuális adatmezõ"
      Top             =   2760
      Width           =   3372
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nincs kapcsolat"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "A kapcsolódó objektum táblázata (Alt+H)"
      Top             =   3120
      Width           =   3372
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mégsem"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Kilépés a tartalom rögzítése nélkül (Ctrl+T)"
      Top             =   3120
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0091E9FB&
      BorderStyle     =   0  'None
      Height          =   288
      Left            =   240
      MaxLength       =   5
      TabIndex        =   2
      Top             =   2160
      Width           =   4092
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Kilépés a tartalom rögzítésével (Esc)"
      Top             =   3120
      Width           =   1212
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2532
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7332
      _ExtentX        =   12933
      _ExtentY        =   4466
      _Version        =   327680
      ForeColorSel    =   16777152
      BackColorBkg    =   14737632
      AllowBigSelection=   -1  'True
      FocusRect       =   2
      GridLinesFixed  =   3
      SelectionMode   =   1
   End
End
Attribute VB_Name = "Tabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- TABLA form (táblázatos adatbevitel) kódja
Public texelo$, backtext$, ujcar%, com3wid%, etext$
Dim mparam$(20), mparamdb%
Public Sub Fcopy(inp$, outp$)
ffi = FreeFile
Open inp$ For Input Shared As #ffi
ffo = FreeFile
Open outp$ For Output Shared As #ffo
Do
   Line Input #ffi, s$
   Print #ffo, s$
Loop While Not EOF(ffi)
Close ffi, ffo
End Sub

Private Sub Command1_Click()
  '--- Kilépés az adatok rögzítésével
  form1.Label2.Caption = ""
  form1.Label3.Caption = ""
  sortorol% = 0
  rogzites% = 1
  Call langclos
  Tabla.Hide
End Sub

Private Sub Command2_Click()
  '--- Kilépés rögzítés nélkül
  form1.Label2.Caption = ""
  form1.Label3.Caption = ""
  sortorol% = 0
  rogzites% = 0
  Call langclos
  Tabla.Hide
End Sub

Private Sub Command3_Click()
  '--- Alt+H kapcsolat
  oszlop% = MSFlexGrid1.Col
  w5& = ADATAB(mesor%(oszlop%)).kapcsob
  If w5& > 0 Then
    objazon1$ = OBJTAB(w5&).obaz
    azonosito$ = ""
    Call alth(objazon1$, azonosito$)
    If azonosito$ <> "" Then
      MSFlexGrid1.Text = azonosito$
      Text1.Text = azonosito$
      Text1.SetFocus
    End If
  End If
  Text1.SetFocus
End Sub

Private Sub Command4_Click()
  '--- sor törlése
  torstart& = MSFlexGrid1.Row
  sortorol% = 1
  rogzites% = 1
  Call langclos
  Tabla.Hide
End Sub

Private Sub Command5_Click()
  '--- Alt+X kapcsolat
  oszlop% = MSFlexGrid1.Col
  w5& = ADATAB(mesor%(oszlop%)).kapcsob
  If w5& > 0 Then
    objazon1$ = OBJTAB(w5&).obaz
    azonosito$ = ""
    Call altx(objazon1$, azonosito$)
    If azonosito$ <> "" Then
      MSFlexGrid1.Text = azonosito$
      Text1.Text = azonosito$
      Text1.SetFocus
    End If
  End If
  Text1.SetFocus

End Sub

Private Sub Form_Activate()
  Command3.Width = com3wid
  'com1wid% = Command3.Width
  'If programnev$ = "AUW-QLIK" And keresobj$ = "PSSZ" And (MSFlexGrid1.Col = 1 Or MSFlexGrid1.Col = 2 Or MSFlexGrid1.Col = 5 Or MSFlexGrid1.Col = 6) Then
  '  MSFlexGrid1.Col = 3
  'End If
End Sub

Private Sub Form_Load()
  '--- TABLA for betöltése
  '--- 0-adik oszlop feltöltése sorszámokkal
  Call langinit("tabla", 2)
  Call szkriptel("tabla")
  'com1wid% = Command3.Width
  MSFlexGrid1.ColWidth(0) = 600
  MSFlexGrid1.ColAlignment(0) = 6
  For s& = 1 To MSFlexGrid1.Rows - 2
    MSFlexGrid1.TextMatrix(s&, 0&) = Str$(s&) + "." + langform(11)
  Next
  MSFlexGrid1.ToolTipText = langform(12) + " " + Str$(MSFlexGrid1.Rows - 2) + " " + langform(11)
End Sub

Private Sub MSFlexGrid1_DblClick()
  vakod$ = ""
  oszlop% = MSFlexGrid1.Col
  adatazonosito$ = Trim$(ADATAB(mesor%(oszlop%)).adkod)
  Call kodtablak(adatazonosito$, vakod$)
End Sub

Private Sub MSFlexGrid1_entercell()
  '--- belépés a cellába
  oszlop% = MSFlexGrid1.Col
  ujcar% = 0
  w5& = ADATAB(mesor%(oszlop%)).kapcsob
  If w5& <> 0 Then
    Command3.Font.Size = 9
    Command3.BackColor = RGB(240, 220, 85)
    Command3.Caption = Trim(OBJTAB(w5&).obnev)
    If OBJTAB(w5&).hashcod <> 0 Then
      Command5.Top = Command3.Top
      Command5.Left = Command3.Left + com3wid% - 1200
      Command5.Visible = True
      Command3.Width = com3wid% - 1220
    Else
      Command3.Width = com3wid%
      Command5.Visible = False
    End If
  Else
    Command3.Font.Size = 8
    Command3.BackColor = RGB(240, 240, 240)
    Command3.Caption = langform(3)
    Command3.Width = com3wid%
    Command5.Visible = False
  End If
  If InStr(mtb$(oszlop%), "*") > 0 Then
    form1.Label3.ForeColor = &HC0&
    capelo$ = langform(13) + " "
  Else
    form1.Label3.ForeColor = &H808080
    capelo$ = ""
  End If
  If InStr(mtb(oszlop%), "Q") > 0 Then
    form1.Label3.ForeColor = RGB(0, 0, 120)
    form1.Label3.Caption = capelo$ + langform(16)
  Else
    form1.Label3.Caption = capelo$ + Mid$(ELLENORZO$(mesor%(sor%)), 3)
  End If
  etext$ = Text1.Text
  form1.Label2.Caption = MAGYARAZAT$(mesor%(oszlop%))
  Text1.MaxLength = mho%(oszlop%)
  Text1.Text = Trim$(MSFlexGrid1.Text)
  etext$ = Text1.Text
  Text1.SelStart = Len(Text1.Text) '+ 1
  Tabla.texelo = Text1.Text
  MSFlexGrid1.CellBackColor = QBColor(14)
  Text1.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
  Text1.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft + 25
  Text1.Height = MSFlexGrid1.CellHeight - 50
  Text1.Width = MSFlexGrid1.CellWidth - 25
  Text2.Text = MSFlexGrid1.TextMatrix(0, oszlop%) + ": " + Text1.Text
  Call mezoeleje(MSFlexGrid1.Row, MSFlexGrid1.Col)
  If InStr(mtb$(oszlop%), "X") > 0 Then
    w5& = ADATAB(mesor%(oszlop%)).kapcsob
    If w5& <> 0 Then
      objazon1$ = OBJTAB(w5&).obaz
      azonosito$ = Text1.Text
      If InStr(mtb$(oszlop%), "J") > 0 And InStr(mtb$(oszlop%), "NZJ") = 0 Then
        azonosito$ = Right$(Space$(14) + azonosito$, mho%(oszlop%))
      Else
        azonosito$ = Left$(azonosito$ + Space$(mho%(oszlop%)), mho%(oszlop%))
      End If
      If azonosito$ <> Space$(mho%(oszlop%)) Then
        Call rekinfo(objazon1, azonosito$)
      Else
        '--- 071108
        form1.Info.Visible = False
        form1.Text1.Visible = False
      End If
    End If
  Else
    '--- 071108
    form1.Info.Visible = False
    form1.Text1.Visible = False
  End If
  Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_gotfocus()
  '--- táblázatra click
  oszlop% = MSFlexGrid1.Col
  ujcar% = 0
  w5& = ADATAB(mesor%(oszlop%)).kapcsob
  If w5& <> 0 Then
    Command3.Caption = Trim(OBJTAB(w5&).obnev)
  Else
    Command3.Caption = langform(3)
  End If
  Text1.MaxLength = mho%(oszlop%)
  Text1.Text = MSFlexGrid1.Text
  Text1.SelStart = Len(Text1.Text) + 1
  Tabla.texelo = Text1.Text
  MSFlexGrid1.CellBackColor = QBColor(14)
  Text1.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
  Text1.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft + 25
  Text1.Height = MSFlexGrid1.CellHeight
  Text1.Width = MSFlexGrid1.CellWidth - 25
  Call MSFlexGrid1_entercell
  Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_LeaveCell()
  '--- cella elhagyása
  MSFlexGrid1.CellBackColor = QBColor(15)
  If programnev$ = "AUW-QLIK" And keresobj$ = "PSSZ" And (MSFlexGrid1.Col = 1 Or MSFlexGrid1.Col = 2 Or MSFlexGrid1.Col = 5 Or MSFlexGrid1.Col = 6) Then
    Text1.Text = etext$
  End If
End Sub

Private Sub MSFlexGrid1_Scroll()
  '--- táblázat scroll esetén biztosítja az aktuális sor láthatóságát
  If MSFlexGrid1.RowIsVisible(MSFlexGrid1.Row) = False Then Text1.Visible = False Else Text1.Visible = True
  Text1.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
End Sub

Private Sub Text1_Change()
  '--- adatmezõ on-line kódviszgálta karakterenként
  Dim param$(30)
  oszlop% = MSFlexGrid1.Col
  If ujcar% <> 0 And Text1.SelStart > 0 Then
    pont% = 0
    minusz% = 0
    h% = Len(Text1.Text)
    For j% = 1 To h%
      If Mid$(Text1.Text, j%, 1) = "." Or Mid$(Text1.Text, j%, 1) = "," Then pont% = pont% + 1
      'If Mid$(Text1.Text, j%, 1) = "." Then pont% = pont% + 1
      If Mid$(Text1.Text, j%, 1) = "-" Then minusz% = minusz% + 1
    Next
    texcar$ = Mid$(Text1.Text, Text1.SelStart, 1)
    texhiba% = 0
    If InStr(mtb$(oszlop%), "N") > 0 Then
      '--- numerikus adatok vizsgálata
      If texcar$ = "-" Then
        If InStr(mtb(oszlop%), "-") > 0 And (Text1.SelStart <> 1 Or minusz% > 1) Then texhiba% = 1
      Else
        If texcar$ = "." Or texcar$ = "," Then
          If InStr(mtb(oszlop%), "T") > 0 And pont% > 1 Then texhiba% = 1
          If InStr(mtb(oszlop%), "T") And texcar$ = "," Then
            texcar$ = "."
          End If
        Else
          If texcar$ < "0" Or texcar$ > "9" Then texhiba% = 1
        End If
      End If
    End If
    If InStr(mtb$(oszlop%), "D") > 0 Then
      '--- dátum adat vizsgálata
      w1$ = "000000"
      Mid$(w1$, 1) = Text1.Text
      hon% = xval(Mid$(w1$, 3, 2))
      If hon% = 4 Or hon% = 6 Or hon% = 9 Or hon% = 11 Then nap% = 30 Else nap% = 31
      If hon% = 2 Then nap% = 29
      If hon% > 12 Then texhiba% = 1
      If xval(Mid$(w1$, 5, 2)) > nap% Then texhiba% = 1
    End If
    '--- ellenõrzõ string vizsgálata
    ellst$ = Trim$(ELLENORZO$(mesor%(oszlop%)))
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
  If InStr(mtb$(oszlop%), "U") > 0 Then
    MSFlexGrid1.Text = UCase$(Text1.Text)
  Else
    MSFlexGrid1.Text = Text1.Text
  End If
  backtext$ = MSFlexGrid1.Text
  Text2.Text = MSFlexGrid1.TextMatrix(0, oszlop%) + ": " + Text1.Text
End Sub

Private Sub Text1_GotFocus()
  '--- text1-re focus esetén
  If MSFlexGrid1.RowIsVisible(MSFlexGrid1.Row) = False Then Text1.Visible = False Else Text1.Visible = True
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- billenytû lenyomása az adatbeviteli mezõben
  '--- enter esetén tartalom ellenõrzése
  ujcar% = 0
  sor = MSFlexGrid1.Row
  osl = MSFlexGrid1.Col
  sorok = MSFlexGrid1.Rows - 2
  oslok = MSFlexGrid1.Cols - 1
  If KeyCode = vbKeyLeft And Text1.SelStart = 0 Or KeyCode = vbKeyRight And Text1.SelStart = Len(Text1.Text) Then
    MSFlexGrid1.Text = Tabla.texelo
    Text1.Text = Tabla.texelo
  End If
  If KeyCode = vbKeyEscape Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    MSFlexGrid1.Text = Tabla.texelo
    Text1.Text = Tabla.texelo
  End If
  Select Case KeyCode
    Case vbKeyBack
    Case vbKeyDelete
      '--- Mezõ törlése
      Text1.Text = ""
    Case vbKeyF1
      If Shift And vbAltMask Then
        rre = Shell(programutvonal$ + "auw-gyi " + terminal$ + task$, vbNormalFocus)
      End If
    Case vbKeyInsert
      If Shift And vbAltMask Then
        oszlop% = MSFlexGrid1.Col
        w5& = ADATAB(mesor%(oszlop%)).kapcsob
        If w5& > 0 And InStr(mtb$(MSFlexGrid1.Col), "Q") > 0 And Not programnev$ = "AUW-QSOPTRG" Then
          objazon1$ = OBJTAB(w5&).obaz
          '--- ide jön a beléptetõ program hívása
          rreobj$ = OBJTAB(ADATAB(mesor%(oszlop%)).kapcsob).obaz
          rrepar$ = dbxneve(rreobj$) + "/" + rreobj$ + "/" + terminal$ + task$ + "/" + auditorutvonal$
          rre = Shell(programutvonal$ + "DBX4-NEW.EXE " + rrepar$, vbNormalFocus)
          Call mess(langmodul(162), 4, 0, langmodul(163), valasz%)
          pxfx = FreeFile
          Open programutvonal$ + terminal$ + task$ + "new.txt" For Binary Shared As pxfx
          csx$ = Space(3)
          Get #pxfx, 1, csx$
          If csx$ <> "NIX" Then
            csx$ = Space(ADATAB(mesor%(oszlop%)).adatho)
            Get #pxfx, 1, csx$
            Text1.Text = csx$
          End If
          Close pxfx
        Else
        ' Eszi - pénztár
          If programnev$ = "AUW-QSOPTRG" And (oszlop% = 6 Or oszlop% = 7) Then
             prgpara$ = "auwszamv/PSSZ"
             If oszlop% = 6 Then
                prgpara$ = "auwszamv/PVSZ"
             End If
             
             Call Fcopy("c:\auwset\auw" + terminal$ + task$ + ".auw", "c:\auwset\auw" + terminal$ + task$ + ".aux")
             
             prgpara$ = prgpara$ + "/" + terminal$ + task$ + "/" + auditorutvonal$
             rre = Shell(programutvonal$ + "AUW-PVRG.EXE " + prgpara$, vbNormalFocus)
             
             Call mess(langmodul(162), 4, 0, langmodul(163), valasz%)
             Call Fcopy("c:\auwset\auw" + terminal$ + task$ + ".aux", "c:\auwset\auw" + terminal$ + task$ + ".auw")
             
          
          End If
        End If
        Text1.SetFocus
      End If
    Case vbKeyHome
      MSFlexGrid1.Row = 1
      KeyCode = 0
    Case vbKeyEnd
      MSFlexGrid1.Row = sorok
      KeyCode = 0
    Case vbKeyPageUp
      If sor > 10 Then
        MSFlexGrid1.Row = MSFlexGrid1.Row - 10
      Else
        MSFlexGrid1.Row = 1
      End If
      KeyCode = 0
    Case vbKeyPageDown
      If sor < sorok - 10 Then
        MSFlexGrid1.Row = MSFlexGrid1.Row + 10
      Else
        MSFlexGrid1.Row = sorok
      End If
      KeyCode = 0
    Case vbKeyDown
      '--- következõ sor
      If sor < sorok Then MSFlexGrid1.Row = MSFlexGrid1.Row + 1
      KeyCode = 0
    Case vbKeyUp
      '--- elõzõ sor
      If sor > 1 Then MSFlexGrid1.Row = MSFlexGrid1.Row - 1
      Text1.SelStart = Text1.SelStart + 1
      KeyCode = 0
    Case vbKeyRight
      '--- következõ oszlop
      If Text1.SelStart = Len(Text1.Text) Then
        If osl < oslok Then MSFlexGrid1.Col = MSFlexGrid1.Col + 1
        Do
          If InStr(tablamaszk, "\" + Trim(Str(MSFlexGrid1.Col)) + "\") > 0 Then
            MSFlexGrid1.Col = MSFlexGrid1.Col + 1
          End If
        Loop While InStr(tablamaszk, "\" + Trim(Str(MSFlexGrid1.Col)) + "\") > 0
        KeyCode = 0
      End If
    Case vbKeyLeft
      '--- elõzõ oszlop
      If Text1.SelStart = 0 Then
        If osl > 1 Then MSFlexGrid1.Col = MSFlexGrid1.Col - 1
        Do
          If InStr(tablamaszk, "\" + Trim(Str(MSFlexGrid1.Col)) + "\") > 0 Then
            MSFlexGrid1.Col = MSFlexGrid1.Col - 1
            If MSFlexGrid1.Col < 1 Then MSFlexGrid1.Col = oslok
          End If
        Loop While InStr(tablamaszk, "\" + Trim(Str(MSFlexGrid1.Col)) + "\") > 0
        KeyCode = 0
      End If
    Case vbKeyM
      If Shift And vbAltMask Then
        If programnev$ = "AUW-PSZL" Then
          If pdarab > 0 Then
            Pszmeg.Show vbModal
            If pszmegikt$ <> "" Then
              mrec$ = dbxkey("PMEG", pszmegikt)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = Mid$(mrec$, 55, 6)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Mid$(mrec$, 138, 14)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = Mid$(mrec$, 226, 6)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = Mid$(mrec$, 121, 14)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = Mid$(mrec$, 121, 14)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12) = pszmegikt
              Text1.Text = Mid$(mrec$, 55, 6)
            End If
            Text1.SetFocus
          End If
        End If
        If programnev$ = "AUW-JOV" Then
          If pdarab > 0 Then
            Jszmeg.Show vbModal
            If pszmegikt$ <> "" Then
              mrec$ = dbxkey("JMEG", pszmegikt)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = Mid$(mrec$, 235, 15)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Mid$(mrec$, 55, 4)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = Mid$(mrec$, 138, 14)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = Mid$(mrec$, 226, 6)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = Mid$(mrec$, 121, 14)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11) = pszmegikt
              Text1.Text = Mid$(mrec$, 235, 15)
            End If
            Text1.SetFocus
          End If
        End If
        If programnev$ = "AUW-QEGY" Or programnev$ = "AUW-QSLK" Or programnev$ = "AUW-QPTRG" Then
            
            If RTrim(Me.Caption) = "Forgalmi tételek" Then
             MSFlexGrid1.Row = MSFlexGrid1.Row + 1
             If MSFlexGrid1.Cols = 7 Then
              MSFlexGrid1.Col = 5
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 1)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 2)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 3)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 4)
             
             Else
              MSFlexGrid1.Col = 8
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 3)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 4)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 5)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 6)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 7)
              
              End If
            Else
             If RTrim(Me.Caption) = "Pénztári tételek" Then
                MSFlexGrid1.Col = 9
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 1)
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 2)
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 3)
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 4)
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 12)
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 13) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, 13)
             '  Eszi
             '  HivKiiras (MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7))
             End If
            End If
            
            Text1.SetFocus
          
        End If
        
      Else
        ujcar% = 1
      End If
    Case vbKeyZ
      If Shift And vbAltMask Then
        If programnev$ = "AUW-RMEG" Or programnev$ = "AUW-RSZL" Then
          If programnev$ = "AUW-RMEG" Then
            armeg.ciko = Left$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) + Space$(15), 15)
            armeg.Show vbModal
            Text1.SetFocus
          Else
            armeg.ciko = Left$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) + Space$(15), 15)
            armeg.Show vbModal
            Text1.SetFocus
          End If
        End If
      Else
        ujcar% = 1
      End If
    Case vbKeyY
      If Shift And vbAltMask Then
        If programnev$ = "AUW-RSZL" Or programnev$ = "AUW-RMEG" Or programnev$ = "AUW-REGY" Then
          '--- készlet lekérdezés
          If programnev$ = "AUW-RMEG" Then
            kerkesz.ciko = Left$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) + Space$(15), 15)
            kerkesz.Show vbModal
            Text1.SetFocus
          Else
            If programnev$ = "AUW-RSZL" Then
              kerkesz.ciko = Left$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) + Space$(15), 15)
              kerkesz.Show vbModal
              If MSFlexGrid1.Col = 2 Then
                Text1.Text = kerminta
              End If
              Text1.SetFocus
            Else
              kerkesz.ciko = Left$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) + Space$(15), 15)
              rkkk$ = Left$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) + Space$(4), 4)
              If Trim(rkkk$) <> "" Then kerkesz.rako = rkkk$
              kerkesz.Show vbModal
              If MSFlexGrid1.Col = 4 Then
                Text1.Text = kerminta
              End If
              Text1.SetFocus
            End If
          End If
        End If
      Else
        ujcar% = 1
      End If
    Case vbKeyX
      If Shift And vbAltMask Then
        oszlop% = MSFlexGrid1.Col
        w5& = ADATAB(mesor%(oszlop%)).kapcsob
        If w5& > 0 Then
          objazon1$ = OBJTAB(w5&).obaz
          azonosito$ = ""
          Call altx(objazon1$, azonosito$)
          If azonosito$ <> "" Then
            MSFlexGrid1.Text = azonosito$
            Text1.Text = azonosito$
            Text1.SetFocus
          End If
        End If
        Text1.SetFocus
      Else
        ujcar% = 1
      End If
    Case vbKeyH
      '--- Alt+H kapcsolat
      If Shift And vbAltMask Then
        Call Command3_Click
      Else
        ujcar% = 1
      End If
    Case vbKeyI
      '--- Alt+I kapcsolt rekord megjelenítése
      If Shift And vbAltMask Then
        oszlop% = MSFlexGrid1.Col
        w5& = ADATAB(mesor%(oszlop%)).kapcsob
        If w5& > 0 Then
          objazon1$ = OBJTAB(w5&).obaz
          azonosito$ = Text1.Text
          If InStr(mtb$(MSFlexGrid1.Col), "J") > 0 Then
            azonosito$ = Right$(Space$(14) + azonosito$, mho%(MSFlexGrid1.Col))
          Else
            azonosito$ = Left$(azonosito$ + Space$(mho%(MSFlexGrid1.Col)), mho%(MSFlexGrid1.Col))
          End If
          Call rekinfo(objazon1$, azonosito$)
        End If
      Else
        ujcar% = 1
      End If
    Case vbKeyS
      '--- Ctrl+S sor törlése
      If Shift And vbCtrlMask Then
        Call Command4_Click
      Else
        ujcar% = 1
      End If
    Case vbKeyT
      '--- Ctrl+T kilépés rögzítés nélkül
      If Shift And vbCtrlMask Then
        Call Command2_Click
      Else
        ujcar% = 1
      End If
    Case vbKeyEscape
      '--- Esc kilépés rögzítéssel
      Call Command1_Click
      'Call langclos
      'Tabla.Hide
    Case vbKeyF12
      If MSFlexGrid1.Row > 1 Then
        Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row - 1, MSFlexGrid1.Col)
      End If
    Case vbKeyReturn
      '--- teljes mezo ellenorzése
      If programnev$ = "AUW-QLIK" And keresobj$ = "PSSZ" And (MSFlexGrid1.Col = 1 Or MSFlexGrid1.Col = 2) Then
        MSFlexGrid1.Col = 4
        If Len(Trim$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))) > 0 Then
           Text1.Text = "L"
           MSFlexGrid1.Text = "L"
        End If
        
        ertek@ = 0
        For i1% = 1 To 200
           If MSFlexGrid1.TextMatrix(i1%, 4) = "L" Then
               ertek@ = ertek@ + xval(MSFlexGrid1.TextMatrix(i1%, 2)) * xval(MSFlexGrid1.TextMatrix(i1%, 3))
           End If
        Next
        form1.Label25.Caption = ertszamy(Str(ertek@), 14, 0)
        netto1$ = Trim(form1.Label22.Caption)
        netto2$ = ""
        For i12% = 1 To Len(netto1$)
           If Not Asc(Mid$(netto1$, i12%, 1)) = 160 Then
             netto2$ = netto2$ + Mid$(netto1$, i12%, 1)
           End If
        Next
        form1.Label29.Caption = ertszamy(Str(Val(netto2$) - ertek@), 14, 0)
        
      End If
      If programnev$ = "AUW-QLIK" And keresobj$ = "PSSZ" And (MSFlexGrid1.Col = 5 Or MSFlexGrid1.Col = 6) Then
           MSFlexGrid1.Col = 1
           MSFlexGrid1.Row = MSFlexGrid1.Row + 1
      End If
     
      If InStr(mtb$(MSFlexGrid1.Col), "NZJ") > 0 Then
        txx$ = Right$("000000000000000" + LTrim$(Text1.Text), mho%(MSFlexGrid1.Col))
        Text1.Text = txx$
        MSFlexGrid1.Text = txx$
      End If
      '--- kapcsolódó objektum ellenõrzése, ha van
      oszlop% = MSFlexGrid1.Col
      w5& = ADATAB(mesor%(oszlop%)).kapcsob
      If w5& <> 0 Then
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
        If azonosito$ = "0000000" And InStr(mtb$(MSFlexGrid1.Col), "NZJ") > "0" Then Text1.Text = "       ": azonosito$ = "       "
        If InStr(mtb$(MSFlexGrid1.Col), "J") > 0 Then
          azonosito$ = Right$(Space$(14) + azonosito$, mho%(MSFlexGrid1.Col))
        Else
          azonosito$ = Left$(azonosito$ + Space$(mho%(MSFlexGrid1.Col)), mho%(MSFlexGrid1.Col))
        End If
        If azonosito$ <> Space$(mho%(MSFlexGrid1.Col)) Then
          w6$ = dbxkey$(objazon1$, azonosito$)
          
          If w6$ = "" Then
            w7$ = OBJTAB(w5&).obnev
            Call mess(langform(14) + " " + w7$, 2, 0, langform(15), valasz%)
            MSFlexGrid1.Text = texelo$
            Text1.Text = texelo$
            KeyCode = 0
          Else
            w6$ = ""
            If InStr(mtb$(MSFlexGrid1.Col), "X") > 0 Then
              Call rekinfo(objazon1, azonosito$)
            End If
          End If
        End If
      End If
      '--- ide betenni
      If InStr(mtb$(MSFlexGrid1.Col), "K") > 0 And Len(Text1.Text) <> mho%(MSFlexGrid1.Col) And Trim(Text1.Text) <> "" Then
        KeyCode = 0
      End If
      
      If KeyCode <> 0 Then
        Call mezovege(MSFlexGrid1.Row, MSFlexGrid1.Col, mezohiba%, 0)
        If mezohiba% = 0 Then
          Call informa
          If osl < oslok Then
            MSFlexGrid1.Col = MSFlexGrid1.Col + 1
          Else
            MSFlexGrid1.Col = 1
            MSFlexGrid1.LeftCol = 1
            If sor < sorok Then MSFlexGrid1.Row = MSFlexGrid1.Row + 1
          End If
        End If
      End If
    Case Else
      '--- rögzített karakter
      ujcar% = 1
      If InStr(mtb$(MSFlexGrid1.Col), "M") > 0 Then
        ujcar% = 0
        KeyCode = 0
      End If

  End Select
  If MSFlexGrid1.Col < MSFlexGrid1.LeftCol Then
    MSFlexGrid1.LeftCol = MSFlexGrid1.LeftCol - 1
  Else
    a& = MSFlexGrid1.Col: lep% = 0
    If a& < oslok Then a& = a& + 1 Else lep% = 1
    If lep% Or MSFlexGrid1.ColIsVisible(a&) = False Then
      MSFlexGrid1.LeftCol = MSFlexGrid1.LeftCol + 1
    End If
  End If
  If MSFlexGrid1.Row < MSFlexGrid1.toprow Then
    MSFlexGrid1.toprow = MSFlexGrid1.toprow - 1
  Else
    a& = MSFlexGrid1.Row
    If a& < sorok Then a& = a& + 1
    If MSFlexGrid1.RowIsVisible(a&) = False Then
      MSFlexGrid1.toprow = MSFlexGrid1.toprow + 1
    End If
  End If
End Sub

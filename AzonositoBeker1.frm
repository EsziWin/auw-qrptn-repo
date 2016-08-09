VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form AzonositoBeker 
   Caption         =   "Csoportos elõleg számla másolat"
   ClientHeight    =   8712
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12816
   LinkTopic       =   "Form2"
   ScaleHeight     =   8712
   ScaleWidth      =   12816
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Mind"
      Height          =   372
      Left            =   9240
      TabIndex        =   11
      Top             =   7680
      Width           =   1812
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Egysem"
      Height          =   372
      Left            =   7080
      TabIndex        =   10
      Top             =   7680
      Width           =   1812
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   8280
      Width           =   6732
      _ExtentX        =   11875
      _ExtentY        =   656
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command 
      Caption         =   "Keresés"
      Height          =   492
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   1452
   End
   Begin VB.TextBox Text2 
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
      Left            =   5400
      TabIndex        =   5
      Text            =   "K"
      Top             =   120
      Width           =   1212
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
      Left            =   3720
      TabIndex        =   3
      Text            =   "P"
      Top             =   120
      Width           =   1332
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nyomtat"
      Height          =   492
      Left            =   7080
      TabIndex        =   1
      Top             =   8160
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mégsem"
      Height          =   492
      Left            =   9240
      TabIndex        =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   1812
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6252
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   12612
      _ExtentX        =   22246
      _ExtentY        =   11028
      _Version        =   327680
      Rows            =   50
      Cols            =   7
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Db:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   8
      Top             =   7680
      Width           =   2052
   End
   Begin VB.Label Label 
      Caption         =   "-"
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
      Index           =   0
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   132
   End
   Begin VB.Label Label1 
      Caption         =   "Kérem a tanfolyam azonosítóját:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3492
   End
End
Attribute VB_Name = "AzonositoBeker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fr4hed$(50), fr4fej$(50), fr4sor$(50), fr4lab$(50), fr4tra$(50), fmezok$(50), afatomb@(6, 2), rafatomb@(6, 2), fr4nev$
Dim regimt$(1001), gysz$(200), hiv$(200), ntafa$(10)
Dim nt$(100)
Public sor, pelo$, rec$, nyugtavolt, fejr$, szamlaszam$
Private Sub Command_Click()
   MSFlexGrid1.Visible = False
   Command1.Visible = False
   Command2.Visible = False
   
   
   dbfi = FreeFile
   Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #dbfi
   ndfi = FreeFile
   
   Open auditorutvonal$ + "auw-part.ndx" For Binary Shared As #ndfi
   rc& = Int(LOF(ndfi) / 20)
   sor = 0
   keres1$ = Trim(Text1.Text)
   keres2$ = Trim(Text2.Text)
   
   For i1d& = 1 To rc&
       Get #ndfi, (i1d& - 1) * 20& + 1, rcim&
       Seek #dbfi, rcim& + 9
       prec$ = Space(900): Get #dbfi, , prec$
       
       azon$ = Mid(prec$, 363, 60)
       If InStr(azon$, keres1$) > 0 And InStr(azon$, keres2$) > 0 Then
         sor = sor + 1
         MSFlexGrid1.TextMatrix(sor, 0) = Str(sor)
         MSFlexGrid1.TextMatrix(sor, 1) = Mid(prec$, 16, 60)
         MSFlexGrid1.TextMatrix(sor, 2) = Trim(Mid(prec$, 106, 8)) + " " + Trim(Mid(prec$, 114, 30)) + " " + Trim(Mid(prec$, 144, 30)) + " " + Trim(Mid(prec$, 174, 10))
         MSFlexGrid1.TextMatrix(sor, 3) = Trim(Mid(prec$, 423, 60))
         MSFlexGrid1.TextMatrix(sor, 4) = Mid(prec$, 184, 15)
         MSFlexGrid1.TextMatrix(sor, 5) = Mid(prec$, 1, 15)
         MSFlexGrid1.TextMatrix(sor, 6) = "I"
       End If
       
   Next
   MSFlexGrid1.Visible = True
   Command1.Visible = True
   Command2.Visible = True
   Label2.Caption = "Db: " + Str(sor)
   Close dbfi, ndfi
   
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub
Private Sub szamlair()
If programnev$ = "AUW-QPTRG" Then
 '--- elõleg számla nyomtatása
  fr4nev$ = "auw-pelo"
  GoSub fr4beolv
  For i9% = 1 To 50: fmezok$(i9%) = "": Next
  irec$ = dbxkey("INST", "INST")
  eloji% = 1
  If masolat% = 0 And sztornoszamla% = 0 Then
   ' bzszam$ = novel(irec$, 454, 6)
   ' Mid$(irec$, 454, 6) = bzszam$
   ' Call  ("INST", irec$, ";", "", "", hiba%)
   ' Mid$(pelo$, 8, 4) = Mid$(irec$, 450, 4)
   ' Mid$(pelo$, 12, 6) = bzszam$
    'Mid$(rec$, 123, 10) = Mid$(pelo$, 8, 10)
    fmezok$(12) = langprg(32)
  Else
    If masolat% = 1 And sztornoszamla% = 1 Then
      fmezok$(12) = langprg(32) + " " + langprg(44) + " " + langprg(33)
      eloji% = -1
    Else
      If masolat% = 1 Then
        ' Eszi
        'fmezok$(12) = langprg(32) + " " + langprg(33)
        fmezok$(12) = langprg(32)
      Else
        fmezok$(12) = langprg(32) + " " + langprg(44)
        eloji% = -1
      End If
    End If
    bzszam$ = Mid$(pelo$, 8, 10)
  End If
  bzx$ = Mid$(pelo, 8, 10)
  '--- fejlec mezok feltoltése
  fmezok$(1) = Mid$(irec$, 5, 60)
  fmezok$(2) = Trim$(Mid$(irec$, 95, 8)) + " " + Trim$(Mid$(irec$, 103, 30)) + ", " + Trim$(Mid$(irec$, 133, 30)) + Trim$(Mid$(irec$, 163, 10))
  fmezok$(3) = Mid$(irec$, 173, 15)
  fmezok$(4) = Mid$(pelo$, 23, 15)
  fmezok$(5) = Mid$(partrec$, 16, 60)
  fmezok$(6) = Trim(Mid$(partrec$, 106, 8)) + " " + Trim(Mid$(partrec$, 114, 30))
  fmezok$(13) = Trim(Mid$(partrec$, 144, 30)) + " " + Trim(Mid$(partrec$, 174, 10))
  fmezok$(7) = Mid$(partrec$, 184, 15)
  fmezok$(8) = datki(Mid$(pelo$, 38, 6))
  ' Eszi
  'fmezok$(9) = datki(maidatum$)
  fmezok$(9) = fmezok$(8)
  
  fmezok$(10) = bzx$
  fmezok$(11) = Mid$(fejr$, 38, 60)
  lfi = FreeFile
  If form852% = 1 Then
    Open listautvonal$ + terminal$ + task$ + "PELO.lst" For Output As #lfi
  Else
    Open listautvonal$ + terminal$ + task$ + "SZLA.lst" For Output As #lfi
  End If
  fr4kod$ = "F"
  GoSub fr4ir
  onert@ = 0: obert@ = 0: For i7% = 1 To 5: afatomb@(i7%, 1) = 0: afatomb@(i7%, 2) = 0: Next
  For i2% = 1 To 5
    elma$ = Mid$(pelo$, (i2% - 1) * 30 + 230, 30)
    If Trim$(elma$) <> "" Then
      For i9% = 1 To 50: fmezok$(i9%) = "": Next
      fmezok$(1) = Mid$(rec$, 30, 25)
      nert@ = xval(Mid$(elma$, 3, 14))
      afaosz@ = xval(Mid$(elma$, 17, 14))
      btto@ = nert@ + afaosz@
      onert@ = onert@ + eloji% * nert@
      obert@ = obert@ + eloji% * btto@
      fmezok$(2) = ertszam(Str(eloji% * nert@), 14, 2)
      afik$ = Mid$(elma$, 1, 2)
      afrec$ = dbxkey("PAFA", afik$)
      fmezok$(3) = Mid$(afrec$, 33, 6)
      fmezok$(4) = ertszam(Str(eloji% * afaosz@), 14, 2)
      afakulcs@ = xval(Mid$(afrec$, 33, 6))
      afajel$ = Mid$(afrec$, 39, 1)
      fmezok$(5) = ertszam(Str$(eloji% * btto@), 14, 2)
      fmezok$(11) = Mid$(partrec$, 363, 60)
      fmezok$(12) = Mid$(partrec$, 423, 60)
 
      If afajel$ = "N" Then
        afatomb@(1, 1) = afatomb@(1, 1) + eloji% * nert@
      Else
        If afajel$ = "M" Then
          afatomb@(2, 1) = afatomb@(2, 1) + eloji% * nert@
        Else
          For afai% = 1 To afakulcsokdb
            If afakulcs@ = afakulcsok(afai%) Then
              afatomb(afai% + 2, 1) = afatomb(afai% + 2, 1) + eloji% * nert@
              afatomb(afai% + 2, 2) = afatomb(afai% + 2, 2) + eloji% * afaosz@
              Exit For
            End If
          Next
        End If
      End If
      fr4kod$ = "S"
      GoSub fr4ir
    End If
  Next
  For i9% = 1 To 50: fmezok$(i9%) = "": Next
  fmezok$(1) = ertszam(Str$(onert@), 14, 2)
  fmezok$(2) = ertszam(Str$(afatomb@(1, 1)), 14, 2)
  fmezok$(3) = ertszam(Str$(afatomb@(2, 1)), 14, 2)
  fmezok$(4) = ertszam(Str$(afatomb@(3, 1)), 14, 2)
  fmezok$(5) = ertszam(Str$(afatomb@(3, 2)), 14, 2)
  fmezok$(6) = ertszam(Str$(afatomb@(4, 1)), 14, 2)
  fmezok$(7) = ertszam(Str$(afatomb@(4, 2)), 14, 2)
  fmezok$(8) = ertszam(Str$(afatomb@(5, 1)), 14, 2)
  fmezok$(9) = ertszam(Str$(afatomb@(5, 2)), 14, 2)
  fmezok$(10) = ertszam(Str$(obert@), 14, 2)
  fr4kod$ = "L"
  GoSub fr4ir
  '--- megjelenítés
  Close lfi
  If form852% = 1 Then
    Shell programutvonal$ + "dbx4-sho " + terminal$ + task$ + "PELO/" + listautvonal$, vbNormalFocus
  Else
    Shell programutvonal$ + "dbx4-qsho 2" + terminal$ + task$ + "SZLA/" + listautvonal$, vbNormalFocus
  End If
Else
' végszámla
  
  If nyugtavolt = 1 Then
    ' kp-s bizonylat
    fr4nev$ = "auw-pszu"
    peldany$ = "1"
  ElseIf nyugtavolt = 2 Then
    ' kp-s száma
    ' külön sorszámon fusson
    fr4nev$ = "auw-pszk"
    peldany$ = "2"
  ElseIf nyugtavolt = 3 Or nyugtavolt = 12 Then
    ' hiteles számla
    peldany$ = "2"
    If nyugtavolt = 12 Then
      fr4nev$ = "auw-psza"
      peldany$ = "2"
    Else
      fr4nev$ = "auw-pszb"
    End If
    nyugtavolt = 3
    
  ElseIf nyugtavolt = 7 Then
    ' garania jegy
    If Nyugel1.Check1.Value = 1 Then
fr4nev$ = "auw-slev"
    Else
fr4nev$ = "auw-pszg"
    End If
    peldany$ = "2"
  Else
    Return
  End If
  
  GoSub fr4beolv
  arf@ = xval(Mid$(fejr$, 41, 10))
  
  '--- belföldi számla
  For i9% = 1 To 50: fmezok$(i9%) = "": Next
'--- fejlec mezok feltoltése
  fmezok$(1) = Mid$(irec$, 5, 60)
  fmezok$(2) = Trim$(Mid$(irec$, 95, 8)) + " " + Trim$(Mid$(irec$, 103, 30)) + " " + Trim$(Mid$(irec$, 133, 30)) + " " + Trim$(Mid$(irec$, 163, 10))
  fmezok$(3) = Mid$(irec$, 173, 15)
  If bankvalasztas% = 0 Then
fmezok$(4) = Trim$(Mid$(irec$, 203, 30)) + " " + banktagol(Mid$(irec$, 233, 24))
fmezok$(5) = Mid$(irec$, 257, 30)
  Else
fkkod$ = Left(Bankval.MSFlexGrid1.TextMatrix(bankvalasztas% - 1, 0) + Space(8), 8)
fkkrec$ = dbxkey("FKSZ", fkkod$)
fmezok$(4) = Trim$(Mid$(fkkrec$, 389, 30)) + " " + banktagol(Mid$(fkkrec$, 419, 24))
fmezok$(5) = Mid$(fkkrec$, 443, 28)
  End If
 
    If Not Trim$(Nyugel1.Text4(1).Text) = "" Then

'      fmezok$(7) = Mid$(pec$, 16, 60)
'      fmezok$(8) = Trim$(Mid$(prec$, 106, 8)) + " " + Trim$(Mid$(prec$, 114, 30)) + " " + Trim$(Mid$(prec$, 144, 30)) + " " + Trim$(Mid$(prec$, 174, 10))
' fmezok$(9) = Mid$(prec$, 184, 15)
    End If
    fmezok$(9) = Nyugel1.Text12.Text
    fmezok$(6) = Mid$(fejr$, 1, 15)
    fmezok$(7) = Nyugel1.Text2.Text
    fmezok$(8) = Nyugel1.Text3.Text

    fmezok$(10) = szamlaszam$
    

    
    fmezok$(11) = datki(Mid$(fejr$, 24, 6))
    fmezok$(12) = datki(Mid$(fejr$, 18, 6))
    fmezok$(13) = datki(Mid$(fejr$, 30, 6))
    fizm$ = Nyugel1.Text7.Text
    fizmrec$ = dbxkey("PFIZ", fizm$)
    '--- 080229 kerekítés
    If fr4nev$ = "auw-pszk" Or fr4nev$ = "auw-pszu" Then
If Mid$(fizmrec$, 33, 1) = "K" Then kpsszamla = 1
    End If
    fmezok$(14) = Nyugel1.Text5.Text
    fmezok$(15) = Trim$(Mid$(fejr$, 111, 60))
    'fmezok$(15) = Nyugel1.Text8.Text
    devnem$ = Mid$(fejr$, 38, 3)
    If devnem$ <> "   " Then
fmezok$(16) = devnem$
fmezok$(17) = ertszam(Mid$(fejr$, 41, 10), 10, 4)
    End If
    If szoveg18$ = "" Then
       szoveg18$ = Mid$(fejr$, 171, 58)
    End If
    fmezok$(18) = szoveg18$
    'If sztornoszamla% = 1 And (seset% = 1 Or seset% = 3 Or seset% = 5) Then
    If sztornoszamla% = 1 Then
fmezok$(10) = sztornoszamlaszam$
fmezok$(26) = " Eredeti számla: " + szamlaszam$
    End If
    If sztornomasolat% = 1 Then
fmezok$(10) = szamlaszam$
fmezok$(26) = " Szornó számla: " + sztornoszamlaszam$
    End If
    
    lfi = FreeFile
    ' garancia jegy
    If fr4nev$ = "auw-slev" And form852% = 1 Then
Open listautvonal$ + terminal$ + task$ + "SLEV.lst" For Output As #lfi
    Else
Open listautvonal$ + terminal$ + task$ + "SZLA.lst" For Output As #lfi
    End If
    fr4kod$ = "F"
    GoSub fr4ir
    '--- sorok összeállítása
    mennyker% = xval(Mid$(irec$, 343, 1))
    ertker% = xval(Mid$(irec$, 344, 1))
    afaker% = xval(Mid$(irec$, 345, 1))
    If ertker% = 0 Then fste$ = "############0" Else fste$ = "#############0." + String(ertker%, "0")
    If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
    If arf@ <> 0 Then fste$ = "#############0.00": fst$ = "#############0.00"
    onert@ = 0: obert@ = 0: For i7% = 1 To 6: afatomb@(i7%, 1) = 0: afatomb@(i7%, 2) = 0: Next
    osuly@ = 0
    For i13% = 1 To 200
For i9% = 1 To 50: fmezok$(i9%) = "": Next
elem$ = Nyugel1.MSFlexGrid1.TextMatrix(i13%, 1)

If Trim$(elem$) <> "" Then
  tkod$ = Nyugel1.MSFlexGrid1.TextMatrix(i13%, 1)
  termrec$ = dbxkey("KTRM", tkod$)
  cikkszam$ = tkod$
  ' ÁFA kezelés itt - újjat ill. másolatot , sztornót megkülönböztetni
  If rnyugtavolt = 10 Then
     afakod$ = Nyugel1.MSFlexGrid1.TextMatrix(i13%, 7)
  Else
     afakod$ = Mid$(termrec$, 706, 2)
  End If
  afrec$ = dbxkey("PAFA", afakod$)
  afakulcs@ = xval(Mid$(afrec$, 33, 6))
  afajel$ = Mid$(afrec$, 39, 1)
  
  menny@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 3))
  megys$ = Mid$(termrec$, 484, 6)
  liar@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 4))
  pensz@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 5))
  elar@ = xval(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 4))
  ' kedvezmény kezelése
  If pensz@ <> 0 Then
    elar@ = elar@ / 100 * (100 - pensz@)
  End If
  
  If menny@ <> 0 Then
    bert@ = elar * menny@
    bert@ = xval(Format(bert@, fste$))
  Else
' Eszi - 2011.09.22
'    bert@ = elar@
    bert@ = 0
  End If
  ' Eszi a 0 bruttó értékû tételt nem írta ki.
  'If bert@ <> 0 Then
    ' oda - vissza jó legyen nettóár, nettó érték, bruttó érték
    nert@ = bert@ / (1 + afakulcs@ / 100)
    nert@ = xval(Format(nert@, fste$))
    afaosz@ = bert@ - nert@
    afaosz@ = xval(Format(afaosz@, fst$))
    
    If menny@ <> 0 Then
      elar@ = nert@ / menny@
      elar@ = xval(Format(elar@, fste$))
    Else
      elar@ = nert@
    End If
    
    obert@ = obert@ + bert@
    onert@ = onert@ + nert@
    If afajel$ = "N" Then
      afatomb@(1, 1) = afatomb@(1, 1) + nert@
    Else
      If afajel$ = "M" Then
        afatomb@(2, 1) = afatomb@(2, 1) + nert@
      Else
        For afai% = 1 To afakulcsokdb
          If afakulcs@ = afakulcsok(afai%) Then
            afatomb(afai% + 2, 1) = afatomb(afai% + 2, 1) + nert@
            afatomb(afai% + 2, 2) = afatomb(afai% + 2, 2) + afaosz@
            Exit For
          End If
        Next
        'If afakulcs@ = 5 Then afatomb@(3, 1) = afatomb@(3, 1) + nert@: afatomb@(3, 2) = afatomb@(3, 2) + afaosz@
        'If afakulcs@ = 15 Then afatomb@(4, 1) = afatomb@(4, 1) + nert@: afatomb@(4, 2) = afatomb@(4, 2) + afaosz@
        'If afakulcs@ = 25 Then afatomb@(5, 1) = afatomb@(5, 1) + nert@: afatomb@(5, 2) = afatomb@(5, 2) + afaosz@
      End If
    End If



    fmezok$(1) = cikkszam$
    fmezok$(2) = Mid$(termrec$, 444, 12)
    fmezok$(3) = Mid$(termrec$, 16, 55)
    
    If fr4nev$ = "auw-psza" Then
       fmezok$(2) = Mid$(termrec$, 16, 55)
       fmezok$(3) = Nyugel1.Text8.Text
    
    End If
    
    fmezok$(4) = Mid$(termrec$, 196, 55)
    fmezok$(5) = ertszam(Str$(stornoelojel@(menny@)), 10, mennyker%)
    fmezok$(6) = megys$
    fmezok$(7) = ertszam(Str$(liar@), 12, ertker%)
    fmezok$(8) = ertszam(Str$(pensz@), 12, ertker%)
'          fmezok$(9) = ertszam(Str$(penft@), 12, ertker%)
    fmezok$(10) = ertszam(Str$(elar@), 12, ertker%)
    fmezok$(11) = ertszam(Str$(afakulcs@), 5, ertker%)
    fmezok$(12) = ertszam(Str$(stornoelojel@(afaosz@)), 12, ertker%)
    fmezok$(13) = ertszam(Str$(stornoelojel@(nert@)), 12, ertker%)
    fmezok$(14) = ertszam(Str$(stornoelojel@(bert@)), 12, ertker%)
    'fmezok$(15) = Mid$(termrec$, 76, 60)
    fmezok$(16) = Mid$(termrec$, 136, 60)
    fmezok$(17) = Mid$(termrec$, 256, 60)
    fmezok$(18) = Mid$(termrec$, 316, 60)
    fmezok$(19) = Mid$(termrec$, 376, 60)
    fmezok$(20) = Mid$(termrec$, 522, 20)
    egysuly@ = xval(Mid$(termrec$, 836, 12))
    netsuly@ = egysuly@ * menny@
    osuly@ = osuly@ + netsuly@
    fmezok$(21) = " " + Trim(ertszam(Str(netsuly@), 14, 2)) + " kg"
    fr4kod$ = "S"
    
    gyariszamok$ = Nyugel1.GyariszamAtad(i13%, gysz$(), hiv$())
    pzx% = InStr(gyariszamok$, ":")
    db% = Val(Mid$(gyariszamok$, 1, pzx% - 1))
    fmezok$(15) = ""
    For i14% = 1 To db%
      fmezok$(15) = fmezok$(15) + Trim(gysz$(i14%)) + ";"
    Next
  
    GoSub fr4ir
  'End If
End If
    Next
    '--- lablec összeállítása
    For i9% = 1 To 50: fmezok$(i9%) = "": Next
  oeloleg@ = 0
  For i11% = 1 To 4
  elolegkulcs@ = 0: elolegkod$ = "": elolalap@ = 0: elolafa@ = 0
  erelszikt$ = Mid$(nt$(i11%), 1, 7)
  elsz$ = Mid$(nt$(i11%), 8, 15)
  eloo@ = xval(Mid$(nt$(i11%), 23, 14))
  If eloo@ <> 0 And Trim(erelszikt$) <> "" Then
    elolegkod$ = ntafa$(i11%)
    If elolegkod$ <> "" Then
      elpafrec$ = dbxkey("PAFA", elolegkod$)
      If elpafrec$ <> "" Then elolegkulcs@ = xval(Mid$(elpafrec$, 33, 6))
    End If
    If elolegkulcs@ <> 0 Then elolafa@ = (eloo@ * elolegkulcs@) / (100 + elolegkulcs@) Else elolafa@ = 0
    elolafa@ = xval(Format(elolafa@, fst$))
    elolalap@ = eloo@ - elolafa@
    If fr4nev$ = "auw-psza" Then
      fmezok$(41 + i11%) = elsz$ + "számú bizonylat alapján"
      fmezok$(15 + i11%) = "Alap:" + ertszam(Str$(elolalap@), 12, 2) + " " + ertszam(Str(elolegkulcs@), 6, 2) + " % ÁFA" + ertszam(Str$(elolafa@), 12, 2)
    Else
       fmezok$(15 + i11%) = elsz$ + "Alap:" + ertszam(Str$(elolalap@), 12, 2) + " " + ertszam(Str(elolegkulcs@), 6, 2) + " % ÁFA" + ertszam(Str$(elolafa@), 12, 2)
    End If
    oeloleg@ = oeloleg@ + eloo@
  End If
Next
oeloleg@ = oeloleg@
    If devnem$ = "   " Then devnem$ = ""
    fmezok$(1) = ertszam(Str$(stornoelojel@(onert@)), 14, afaker%) + " " + devnem$
    fmezok$(2) = ertszam(Str$(stornoelojel@(afatomb@(1, 1))), 14, afaker%) + " " + devnem$
    fmezok$(3) = ertszam(Str$(stornoelojel@(afatomb@(2, 1))), 14, afaker%) + " " + devnem$
    fmezok$(4) = ertszam(Str$(stornoelojel@(afatomb@(3, 1))), 14, afaker%) + " " + devnem$
    fmezok$(5) = ertszam(Str$(stornoelojel@(afatomb@(3, 2))), 14, afaker%) + " " + devnem$
    fmezok$(6) = ertszam(Str$(stornoelojel@(afatomb@(4, 1))), 14, afaker%) + " " + devnem$
    fmezok$(7) = ertszam(Str$(stornoelojel@(afatomb@(4, 2))), 14, afaker%) + " " + devnem$
    fmezok$(8) = ertszam(Str$(stornoelojel@(afatomb@(5, 1))), 14, afaker%) + " " + devnem$
    fmezok$(9) = ertszam(Str$(stornoelojel@(afatomb@(5, 2))), 14, afaker%) + " " + devnem$
    fmezok$(34) = ertszam(Str$(stornoelojel@(afatomb@(6, 1))), 14, afaker%) + " " + devnem$
    fmezok$(35) = ertszam(Str$(stornoelojel@(afatomb@(6, 2))), 14, afaker%) + " " + devnem$
    '--- 080229 kerekítés
    fizetnem@ = obert@ - oeloleg@
    If kpsszamla% = 1 Then
'--- készpénzes
Call kerekit510(fizetnem@, fizetni@, kerek@, "K")
fmezok$(36) = ertszam(Str$(kerek@), 14, 0) + " " + devnem$
    Else
fizetni@ = fizetnem@: kerek@ = 0
    End If
    fmezok$(10) = ertszam(Str$(stornoelojel@(obert@)), 14, ertker%) + " " + devnem$
    fmezok$(11) = ertszam(Str$(stornoelojel@(oeloleg@)), 14, ertker%) + " " + devnem$
    If fr4nev$ = "auw-psza" Then
       fmezok$(11) = ertszam(Str$(stornoelojel@(-oeloleg@)), 14, ertker%) + " " + devnem$
    End If
    '--- 080229 kerekítés
    fmezok$(39) = ertszam(Str(stornoelojel@(kerek@)), 14, ertker%)
    fmezok$(12) = ertszam(Str$(stornoelojel@(fizetni@)), 14, ertker%) + " " + devnem$
    fmezok$(13) = Trim$(Mid$(partrec$, 363, 60)) + " " + Trim$(Mid$(partrec$, 423, 60))
    fmezok$(40) = betuvel$(stornoelojel@(fizetni@), devnem$)
    fmezok$(41) = Nyugel1.Text8.Text
    fmezok$(46) = " számú bizonylat alapján"
    If devnem$ <> "" Then
arf@ = xval(Mid$(fejr$, 41, 10))
arfi# = arf@
fmezok$(20) = ARfkonv(fmezok$(2), arfi#, 14)
fmezok$(21) = ARfkonv(fmezok$(3), arfi#, 14)
fmezok$(22) = ARfkonv(fmezok$(4), arfi#, 14)
fmezok$(23) = ARfkonv(fmezok$(5), arfi#, 14)
fmezok$(24) = ARfkonv(fmezok$(6), arfi#, 14)
fmezok$(25) = ARfkonv(fmezok$(7), arfi#, 14)
fmezok$(26) = ARfkonv(fmezok$(8), arfi#, 14)
fmezok$(27) = ARfkonv(fmezok$(9), arfi#, 14)
fmezok$(36) = ARfkonv(fmezok$(34), arfi#, 14)
fmezok$(37) = ARfkonv(fmezok$(35), arfi#, 14)
    End If
    If fr4nev$ = "auw-sleb" Then
fmezok$(28) = langprg(78) + Trim(bovslev.Text1.Text)
fmezok$(29) = langprg(79) + ertszam(Str(osuly@), 10, 2) + " kg"
fmezok$(30) = Trim(bovslev.Text2.Text)
fmezok$(31) = Trim(bovslev.Text3.Text)
fmezok$(32) = Trim(bovslev.Text4.Text)
fmezok$(33) = Trim(bovslev.Text5.Text)
    End If
    fr4kod$ = "L"
    GoSub fr4ir
    '--- megjelenítés
    Close lfi
 ' Call mess("Nyugtavolt:" + Str(nyugtavolt), 1, 0, langmodul(157), valasz%)
  
  If (masolat% = 1 Or sztornomasolat% = 1) And Not melyik = 3 Then
   If form852% = 1 Then
     'Shell programutvonal$ + "dbx4-sho " + terminal$ + task$ + "SLEV/" + listautvonal$, vbNormalFocus
     Shell programutvonal$ + "dbx4-qsho " + peldany$ + terminal$ + task$ + "SLEV/" + listautvonal$, vbNormalFocus
   Else
     Shell programutvonal$ + "dbx4-sho " + terminal$ + task$ + "SZLA/" + listautvonal$, vbNormalFocus
   End If
  
  Else
   If form852% = 1 Then
     Shell programutvonal$ + "dbx4-qsho " + peldany$ + terminal$ + task$ + "SLEV/" + listautvonal$, vbNormalFocus
   Else
     Shell programutvonal$ + "dbx4-qsho " + peldany$ + terminal$ + task$ + "SZLA/" + listautvonal$, vbNormalFocus
   End If
  End If



End If
  Call gomb("Tovább " + " &", gg%, 8320, 100, "V")
  GoTo vege
fr4beolv:
  If langutvonal$ = "" Then
    ffi = FreeFile
    Open programutvonal$ + fr4nev$ + ".fx4" For Binary Shared As #ffi
    formfm& = LOF(ffi)
    Close ffi
    ffi = FreeFile
    If formfm& > 0 Then
      Open programutvonal$ + fr4nev$ + ".fx4" For Input Shared As #ffi
    Else
      Open programutvonal$ + fr4nev$ + ".fr4" For Input Shared As #ffi
    End If
  Else
    ffi = FreeFile
    Open langutvonal$ + fr4nev$ + ".fx4" For Binary Shared As #ffi
    formfm& = LOF(ffi)
    Close ffi
    ffi = FreeFile
    If formfm& > 0 Then
      Open langutvonal$ + fr4nev$ + ".fx4" For Input Shared As #ffi
    Else
      Open langutvonal$ + fr4nev$ + ".fr4" For Input Shared As #ffi
    End If
  End If
  heddb% = 0: fejdb% = 0: sordb% = 0: labdb% = 0: tradb% = 0
  olvkez% = 0
  Do
    Line Input #ffi, fs$
    ko$ = Left$(fs$, 2)
    If ko$ <> "* " Then
      If olvkez% = 0 Then
        If UCase$(fs$) = "PRINT=852" Then form852% = 1 Else form852% = 0
        olvkez% = 1
      End If
      Select Case ko$
        Case "F="
          fejdb% = fejdb% + 1
          fr4fej$(fejdb%) = fs$
        Case "S="
          sordb% = sordb% + 1
          fr4sor$(sordb%) = fs$
        Case "L="
          labdb% = labdb% + 1
          fr4lab$(labdb%) = fs$
        Case Else
          If fejdb% = 0 Then
            heddb% = heddb% + 1
            fr4hed$(heddb%) = fs$
          Else
            If labdb% > 0 Then
              tradb% = tradb% + 1
              fr4tra$(tradb%) = fs$
            End If
          End If
      End Select
    End If
  Loop While Not EOF(ffi)
  Close ffi
Return
fr4ir:
  '--- listafile írása
  If fr4kod$ = "F" Then
    If form852% = 0 Then
      For i14% = 1 To heddb%
        sr$ = fr4hed$(i14%)
        Print #lfi, sr$
      Next
    End If
    If masolat% = 0 And sztornoszamla% = 0 And sztornomasolat% = 0 And sztornomasolatX% = 0 Then
        sr$ = "F=Ariel/10/B/K &                 &"
        If programnev$ = "AUW-QPTRG" Then
        Else
         Print #lfi, sr$
        End If
    End If
    
    For i15% = 1 To fejdb%
      sr$ = fr4fej$(i15%)
      pzx% = InStr(sr$, "#")
      Do While pzx% > 0
        sosz% = Val(Mid$(sr$, pzx% + 1, 2))
        mzx$ = fmezok$(sosz%)
        If Len(mzx$) < 3 Then mzx$ = mzx$ + "   "
        pzz% = InStr(pzx% + 1, sr$, "#")
        pzzz% = InStr(pzx% + 1, sr$, "[")
        If pzz% <> 0 And pzzz% <> 0 Then
          If pzzz% < pzz% Then pzz% = pzzz%
        Else
          If pzz% = 0 And pzzz% <> 0 Then pzz% = pzzz%
        End If
        If pzz% > 0 Then
          ureshely% = pzz% - pzx% - 1
          If Len(mzx$) > ureshely% Then mzx$ = Left(mzx$, ureshely%)
        End If
        Mid$(sr$, pzx%) = mzx$
        pzx% = InStr(sr$, "#")
      Loop
      If form852% = 1 Then
        sis$ = Mid$(sr$, 3): psis% = InStr(sis$, "&"): If psis% > 0 Then sis$ = Left(sis$, psis% - 1)
        If i15% = 1 Then sr$ = "CM" + sis$ Else sr$ = "FL" + sis$
      End If
      Print #lfi, sr$
    Next
    Return
  End If
  If fr4kod$ = "S" Then
    For i15% = 1 To sordb%
      sr$ = fr4sor$(i15%)
      pzx% = InStr(sr$, "#")
      Do While pzx% > 0
        sosz% = Val(Mid$(sr$, pzx% + 1, 2))
        mzx$ = fmezok$(sosz%)
        If Len(mzx$) < 3 Then mzx$ = mzx$ + "   "
        pzz% = InStr(pzx% + 1, sr$, "#")
        pzzz% = InStr(pzx% + 1, sr$, "[")
        If pzz% <> 0 And pzzz% <> 0 Then
          If pzzz% < pzz% Then pzz% = pzzz%
        Else
          If pzz% = 0 And pzzz% <> 0 Then pzz% = pzzz%
        End If
        If pzz% > 0 Then
          ureshely% = pzz% - pzx% - 1
          If Len(mzx$) > ureshely% Then mzx$ = Left(mzx$, ureshely%)
        End If
        Mid$(sr$, pzx%) = mzx$
        pzx% = InStr(sr$, "#")
      Loop
      If form852% = 1 Then
        sis$ = Mid$(sr$, 3): psis% = InStr(sis$, "&"): If psis% > 0 Then sis$ = Left(sis$, psis% - 1)
        sr$ = "TS" + sis$
      End If
      Print #lfi, sr$
    Next
    Return
  End If
  If fr4kod$ = "L" Then
    For i15% = 1 To labdb%
      sr$ = fr4lab$(i15%)
      pzx% = InStr(sr$, "#")
      Do While pzx% > 0
        sosz% = Val(Mid$(sr$, pzx% + 1, 2))
        mzx$ = fmezok$(sosz%)
        If Len(mzx$) < 3 Then mzx$ = mzx$ + "   "
        pzz% = InStr(pzx% + 1, sr$, "#")
        pzzz% = InStr(pzx% + 1, sr$, "[")
        If pzz% <> 0 And pzzz% <> 0 Then
          If pzzz% < pzz% Then pzz% = pzzz%
        Else
          If pzz% = 0 And pzzz% <> 0 Then pzz% = pzzz%
        End If
        If pzz% > 0 Then
          ureshely% = pzz% - pzx% - 1
          If Len(mzx$) > ureshely% Then mzx$ = Left(mzx$, ureshely%)
        End If
        Mid$(sr$, pzx%) = mzx$
        pzx% = InStr(sr$, "#")
      Loop
      If form852% = 1 Then
        sis$ = Mid$(sr$, 3): psis% = InStr(sis$, "&"): If psis% > 0 Then sis$ = Left(sis$, psis% - 1)
        sr$ = "TS" + sis$
      End If
      Print #lfi, sr$
    Next
    If form852% = 0 Then
      For i14% = 1 To tradb%
        sr$ = fr4tra(i14%)
        Print #lfi, sr$
      Next
    End If
  End If
Return

vege:
End Sub

Private Sub Command2_Click()
irec$ = dbxkey("INST", "INST")
        
If programnev$ = "AUW-QPTRG" Then
        
  dbfi = FreeFile
  Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + "auw-pelo.ndx" For Binary Shared As #ndfi
  rc& = Int(LOF(ndfi) / 12)

  szamladb = 0
  ProgressBar1.Min = 0
  ProgressBar1.Max = 100
   
  For i1d& = 1 To rc&
       ProgressBar1.Value = pscale(i1d&, rc&)
       
       Get #ndfi, (i1d& - 1) * 12& + 1, rcim&
       Seek #dbfi, rcim& + 9
       pelo$ = Space(550): Get #dbfi, , pelo$
       
       epkod$ = Mid(pelo$, 23, 15)
       van = False
       For i = 1 To sor
         kell$ = Mid(MSFlexGrid1.TextMatrix(i, 6) + Space$(1), 1, 1)
         pkod$ = Mid(MSFlexGrid1.TextMatrix(i, 5) + Space$(15), 1, 15)
         If epkod$ = pkod$ Then
            If kell$ = "I" Then
               van = True
            End If
            Exit For
         End If
       Next
       
       If van Then
       
          masolat% = 1
          ptikt$ = Mid$(pelo$, 105, 7)
          rec$ = dbxkey("PKTE", ptikt$)
          pkod$ = Mid$(pelo$, 23, 15)
          partrec$ = dbxkey("PART", pkod$)
          osszeg@ = Val(Mid$(pelo$, 44, 14))
          If Mid$(pelo$, 90, 1) = "S" Then sztornoszamla% = 1 Else sztornoszamla% = 0
          If sztornoszamla% = 0 And osszeg > 0 Then
            If partrec$ <> "" Then
              Call szamlair
              szamladb = szamladb + 1
              
            End If
          End If
          masolat% = 0
       
       
       End If
        
  Next
Else

  sirec$ = dbxkey("SINS", "INST")


  dbxfi = FreeFile
  Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #dbxfi
  ndxfi = FreeFile
  Open auditorutvonal$ + "auw-kszb.ndx" For Binary Shared As #ndxfi
  rc& = Int(LOF(ndxfi) / 15)
    
  ProgressBar1.Min = 0
  ProgressBar1.Max = 100

  szamladb = 0
  gombs1% = 2
  For i2d& = 1 To rc&
       ProgressBar1.Value = pscale(i2d&, rc&)
       
       Get #ndxfi, (i2d& - 1) * 15& + 1, rcim&
       Seek #dbxfi, rcim& + 9
       rec$ = Space(300): Get #dbxfi, , rec$
       
       epkod$ = Mid(rec$, 61, 15)
       stornojel$ = Mid(rec$, 35, 1)
       van = False
       For i = 1 To sor
         kell$ = Mid(MSFlexGrid1.TextMatrix(i, 6) + Space$(1), 1, 1)
         pkod$ = Mid(MSFlexGrid1.TextMatrix(i, 5) + Space$(15), 1, 15)
         If epkod$ = pkod$ And stornojel$ = " " Then
            If kell$ = "I" Then
               van = True
            End If
            Exit For
         End If
       Next
       
       If van Then

          masolat% = 0: sztornoszamla% = 0: sztornomasolat% = 0
          szoveg18$ = ""
          
          If rec$ <> "" Then
             For i1% = 1 To 1000
                mt$(i1%) = Space$(120)
                regimt$(i1%) = Space$(120)
             Next
             xrec$ = Space$(700)
             vsorszam& = 0
             w1% = obsorszama("KSZF")
             For i1% = 1 To 5: nt$(i1%) = Space$(43): Next
      '--- számla beolvasása
             szamla1$ = Mid$(rec$, 1, 10)
             erxszamla$ = szamla1$
             pszbrec$ = dbxkey("KSZB", szamla1$)
             teljesdat$ = Mid$(pszbrec$, 84, 6)
             konyveldat$ = Mid$(pszbrec$, 231, 6)
             fejr$ = Mid$(pszbrec$, 61, 170) + Mid$(pszbrec$, 237, 8) + Space$(122)
             Mid$(fejr$, 111, 60) = Mid$(pszbrec$, 111, 60)
             pkod$ = Mid$(fejr$, 1, 15)
             partrec$ = dbxkey("PART", pkod$)
         
             ' Eszi - elõleg beolvasása
             pvszikt$ = Mid$(pszbrec$, 53, 7)
             ' Ha számla, nem szállító levél
             If Not Trim$(pvszikt$) = "" Then
                pvszrec$ = dbxkey("PVSZ", pvszikt$)
                For i9% = 1 To 5
                   nt$(i9%) = Mid$(pvszrec$, (i9% - 1) * 43 + 1280, 43)
                              
                   If Not nt$(i9%) = "" Then
                     ' partner kód, számlaszám
                     ntafa$(i9%) = elolegafa(Mid$(pvszrec$, 38, 15), Mid$(nt$(i9%), 8, 10))
                   End If
                 ' elolegafa()
                Next
             End If
             

            If pszbrec$ <> "" Then
              trdarab% = xval(Mid$(pszbrec$, 50, 3))
              trind$ = szamla1$ + "001"
              psztrec$ = dbxkey("KSZT", trind$)
              If psztrec$ <> "" Then
                w1% = obsorszama("KSZT")
                kezdoix& = OBJTAB(w1%).obind
                dbfi = FreeFile
                Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #dbfi
                ndfi = FreeFile
                Open auditorutvonal$ + "auw-kszt.ndx" For Binary Shared As #ndfi
                rc& = Int(LOF(ndfi) / 12)
                For i1% = 1 To trdarab%
                  i1d& = kezdoix& + i1% - 1
                  Get #ndfi, (i1d& - 1) * 18& + 1, rcim&
                  Seek #dbfi, rcim& + 9
                  psztrec$ = Space(170): Get #dbfi, , psztrec$
                  mt$(i1%) = Mid$(psztrec$, 15, 120)
                  ktikt$ = Mid$(psztrec$, 158, 7)
                  'kszxrec$ = dbxkey("KSZX", ktikt$)
                  ' ÁFA kód itt
                  afakod$ = Mid$(psztrec$, 81, 2)
                  Nyugel1.MSFlexGrid1.TextMatrix(i1%, 7) = afakod$
            
                  Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) = Mid$(psztrec$, 107, 15)
                  Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3) = Mid$(psztrec$, 21, 12)
                  Nyugel1.MSFlexGrid1.TextMatrix(i1%, 4) = Mid$(psztrec$, 57, 12)
               
                  Nyugel1.MSFlexGrid1.TextMatrix(i1%, 5) = Str$(-Val(Mid$(psztrec$, 51, 6)))
                  ktetikt$ = Mid$(psztrec$, 158, 7)
                  kfttrec$ = dbxkey("KKFT", ktetikt$)
                  If kfttrec$ <> "" Then
                     bizikt$ = Mid(kfttrec$, 8, 7)
                     ksybrec$ = dbxkey("KSYB", bizikt$)
                     Call feltolt2(ksybrec$, pkod$)
                     If Mid$(pszbrec$, 254, 1) = "H" Then
                        If Mid$(pkod$, 1, 1) = "T" Then
                          kod1$ = Trim$(Mid$(partrec$, 363, 60))
                          kod2$ = Trim$(Mid$(partrec$, 423, 60))
                  
                          Nyugel1.Text8.Text = kod1$
                          Nyugel1.Text11.Text = kod2$
                        End If
                     End If
               
                     Mid$(fejr$, 171, 58) = Nyugel1.Text11.Text
                     fizmod$ = Mid$(pszbrec$, 96, 2)
                     If fizmod$ = "  " Then
                        pvszikt$ = Mid$(pszbrec$, 53, 7)
                        pvszrec$ = dbxkey("PVSZ", pvszikt$)
                        If Not pvszrec$ = "" Then
                           fizmod$ = Mid$(pvszrec$, 76, 2)
                        End If
                     End If
                     fmrec$ = dbxkey("PFIZ", fizmod$)
                     Nyugel1.Text5.Text = Mid$(fmrec$, 3, 30)
                     Nyugel1.Text7.Text = fizmod$

                  End If
                             
                Next
                
                Close dbfi
                Close ndfi
              
                pkod$ = Mid$(fejr$, 1, 15)
                partrec$ = dbxkey("PART", pkod$)
                Call feltolt
                Mid$(fejr$, 171, 58) = Trim$(Nyugel1.Text11.Text) + Space$(58)
        
                dat$ = Mid$(fejr$, 18, 6)
                If Mid$(fejr$, 17, 1) = "P" Then proforma% = 1 Else proforma% = 0
                If Mid$(fejr$, 16, 1) <> "B" And Mid$(fejr$, 16, 1) <> "S" Then megnbeal.Show vbModal
                
                proforma% = 0
                rnyugtavolt = nyugtavolt
                Nyugel1.Check1.Value = 0
                Select Case Mid$(pszbrec$, 254, 1)
                Case "N"
                  nyugtavolt = 1
                Case "K"
                  nyugtavolt = 2
                Case "H", " "
                  nyugtavolt = 3
                  If Mid$(pkod$, 1, 1) = "T" Then
                     nyugtavolt = 12
                     kod1$ = Trim$(Mid$(partrec$, 363, 60))
                     kod2$ = Trim$(Mid$(partrec$, 423, 60))
                     Nyugel1.Text8.Text = kod1$
                     Nyugel1.Text11.Text = kod2$
                  End If
                Case "G"
                  nyugtavolt = 7
                Case "S"
                  nyugtavolt = 7
                   Nyugel1.Check1.Value = 1
                Case Else
                End Select
                szamlaszam$ = Mid$(pszbrec$, 1, 10)
      '--- számla nyomtatása
                If Mid$(pszbrec$, 11, 10) <> Space$(10) Then
                  szoveg18$ = langprg(60) + " " + LCase(langprg(51)) + ":" + Mid$(pszbrec$, 11, 10)
                  helyes% = 0
                End If
        
        
                Call szamlair
                
                
              End If
            End If
       End If
     End If
  Next
End If
Call mess("Elkészült a nyomtatás!", 2, 0, langprg(1), valasz%)
        
End Sub

Private Sub Command3_Click()
 For i = 1 To sor
   MSFlexGrid1.TextMatrix(i, 6) = " "
 Next
End Sub

Private Sub Command4_Click()
 For i = 1 To sor
   MSFlexGrid1.TextMatrix(i, 6) = "I"
 Next
End Sub

Private Sub Form_Load()
  MSFlexGrid1.TextMatrix(0, 0) = "Halgatók"
  MSFlexGrid1.TextMatrix(0, 1) = "Név"
  MSFlexGrid1.TextMatrix(0, 2) = "Cím"
  MSFlexGrid1.TextMatrix(0, 3) = "HIR azonosító"
  MSFlexGrid1.TextMatrix(0, 4) = "Adószám"
  MSFlexGrid1.TextMatrix(0, 5) = "Partnerkód"
  MSFlexGrid1.TextMatrix(0, 6) = "I/N"

  MSFlexGrid1.ColAlignment(2) = 1
  MSFlexGrid1.ColAlignment(3) = 1
  MSFlexGrid1.ColAlignment(4) = 1
  MSFlexGrid1.ColWidth(0) = 500
  MSFlexGrid1.ColWidth(1) = 3500
  MSFlexGrid1.ColWidth(2) = 3500
  MSFlexGrid1.ColWidth(3) = 2000
  MSFlexGrid1.ColWidth(4) = 1500

  
End Sub
Private Function elolegafa$(pkod$, elolegszla$)
                    partrec$ = dbxkey("PART", pkod$)
                    nxptr& = Val(Mid$(partrec$, 702, 10))
                    elolegdb% = 0
                    dxfi = FreeFile
                    Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #dxfi
                    fim& = LOF(dxfi)
                    Do While nxptr& > 0
                      Seek #dxfi, nxptr& + 9
                      elrec$ = Space(650): Get #dxfi, , elrec$
                      nxptr& = Val(Mid$(elrec$, 204, 10))
                      elokod$ = "BV"
                      If Mid$(elrec$, 90, 1) <> "S" And Mid$(elrec$, 224, 2) = elokod$ Then
                        If Mid$(elrec$, 8, 10) = elolegszla$ Then
                          elem1$ = Mid$(elrec$, 8, 15) + " " + Mid$(elrec$, 38, 6) + " " + Right$(Space$(14) + Format(szabossz@, "##########0.00"), 14) + " " + Mid(elrec$, 1, 7)
              For jij% = 1 To 5:
                            jelem$ = Mid$(elrec$, (jij% - 1) * 30 + 230, 30)
                            If Trim(jelem$) <> "" Then
                              elolegafa$ = Mid$(jelem$, 1, 2)
                              Exit Function
                            End If
                          Next
                        End If
                      End If
                    Loop
                    Close dxfi

End Function


Private Sub feltolt2(ksybrec$, pkod$)
        If Not ksybrec$ = "" Then
        Nyugel1.Text4(1) = Mid$(ksybrec$, 8, 15)
        Nyugel1.Text2 = Mid$(ksybrec$, 23, 60)
        Nyugel1.Text3 = Mid$(ksybrec$, 83, 60)
        Nyugel1.Text11 = Mid$(ksybrec$, 143, 58)
        End If
        If Trim(Nyugel1.Text4(1)) = "" Then
              Nyugel1.Text4(1) = pkod$
        End If
    
End Sub

Private Sub feltolt()


        Nyugel1.Text2 = Mid$(partrec$, 16, 60)
        Nyugel1.Text3 = postacim(partrec$, 106)
        Nyugel1.Text12.Text = Trim(Mid$(partrec$, 184, 15))
        If Not Mid$(Nyugel1.Text12.Text, 9, 1) = "-" Then
             adoazjel = Trim(Mid$(partrec$, 184, 15))
             If Len(adoazjel) = 10 Then
             Nyugel1.Text12.Text = Trim(Mid$(partrec$, 184, 15))
             Else
             Nyugel1.Text12.Text = Mid$(Nyugel1.Text12.Text, 1, 8) + "-" + Mid$(Nyugel1.Text12.Text, 9, 1) + "-" + Mid$(Nyugel1.Text12.Text, 10, 2)
             End If
        End If
        poz = InStr(Nyugel1.Text11.Text, "Bev.")
        If poz > 0 Then
           Nyugel1.Text11.Text = Mid$(Nyugel1.Text11.Text, 1, poz - 1)
        End If
End Sub
Private Function stornoelojel@(osszeg@)
  
  If sztornoszamla% = 1 Or sztornomasolat% = 1 Then
    stornoelojel@ = -osszeg@
  Else
    stornoelojel@ = osszeg@
  End If
End Function

Private Function betuvel$(fizetni@, devnem$)
  filler$ = Mid$(ertszam(Str$(fizetni@), 14, 2), 13, 2)
  If devnem$ = "" Or devnem$ = "HUF" Then
       devmas$ = "Ft"
  Else
       devmas$ = devnem$
  End If
  betuvel$ = szamszoveg(fizetni@, 0, "") + " " + filler$ + "/100 " + devmas$

End Function


Private Sub MSFlexGrid1_DblClick()
   s = MSFlexGrid1.RowSel
   
   If MSFlexGrid1.TextMatrix(s, 6) = "I" Then
         MSFlexGrid1.TextMatrix(s, 6) = "N"
   Else
         MSFlexGrid1.TextMatrix(s, 6) = "I"
   End If
End Sub

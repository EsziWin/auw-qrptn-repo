Attribute VB_Name = "auwrutin"
Type MTYP
  cs As Integer
  CO As Integer
  rp As Integer
  mh As Integer
End Type
Type kontrol               '-kontroll mezo tipus
  kopoz As Integer         '-kontroll mezo poz. a rekordban
  kohossz As Integer       '-kontroll mezo hossza
  car As String * 1        '-aláhúzás elõtte
  cax As String * 1        '-aláhúzás utánna
End Type
Dim kontomb(10) As kontrol '-kontroll mezo tomb pointer=kontp%
Dim konert$(10)            '-kontroll ertekek
Dim hasert$(10)            '-hasonlito ertekek
Dim kosor$(10)             '-kontroll osszegsor
Dim gyujto@(10, 20)        '-gyujto
Type telepkom              '-telepleíró
  Ltvan As Integer          '-1-van auwkom.par 0-nincs auwkom.par
  Lipcim As String         '-központi szerve IP címe
  Luser As String          '-felhasználó a belépéshez
  Ljelszo As String        '-jelszó a belépéshez
  Ltelepdb As Integer      '-telepek száma
  Lsajtip As String * 2    '-saját telep típusa (KP vagy LR)
  Lraktar As String * 4    '-saját telep raktár kódja
  Ldirnev As String        '-saját telep cég könyvtár a központi szerver auwin alatt
  Lsajkom As String        '-kommunikációs könyvtár a saját szerveren
  Ltavkom As String        '-kommunikációs könyvtár a központi szerveren
  Ltelindex As Integer     '-telep sorszáma
  Ltelepek(99) As String * 12  '-raktárkod(4)+alkönyvtár(8)
End Type
Public telkom As telepkom
Type sorleir               '-sor mezoi tipus
  mspoz As Integer         '-pozicio a sorban
  msorszam As Integer      '-mezosorszam az ux-ben
  fmt As String * 60       '-formatumstring
  ksorsz As Integer        '-kifejezes sorszama
  halm As Integer
  flem As Integer
End Type
Dim sor(30) As sorleir     '-sor mezoi tomb pointer=spp%
Dim fejt$(20)              '-fejlecek pointer=fejp%
Dim kif$(10)               '-kifejezesek pointer=kifp%
                           '-srr$ a listasor
Dim elo$(30), abert$(30)
Dim mp(1001) As MTYP
Dim parr$(1001)
Dim narr$(1001)
Dim eljp%(1001), elje$(1001)
Dim ux%(1001), param$(20)
Public komirakt$(50), komiraktdb%, Arszorzo1%, wrogzites%, leltardat$
Public partnermeg(10000), pdarab, pszlpart$, pszmegikt$
Public sseged%(), oseged%(), sortores%, szelesseg%
Public mennyisegpozicio%, mennyiseghossz%, mennyisegelojel$
Public afakulcsokdb%, afakulcsok@(5), msgvalasz%, vakod$
Public keresztarfolyam#(200), kxarf#, arpartkod$, arbizdatum$
Public cimvekt&(), cimvektdb&, kerminta$, afakodja$(1000), arkateg%
Public adatfiltomb$(), adatcimtomb&(), adatelemszam&, filterek$(10), filterdarab%
Public tapgonkod$(1001), tapgonelar@(1001), tapgonbear@(1001)
Public cikszam$(), vokod$(), vdb&, kosarfajl$, kiadjegy%
Public partkarb%, kulsobolt%, aktualistermek$, aktualiscsoport$, aktualisafakod$, aktualisarvalt$
'--- 080229 kerekítés
Public kpsszamla%, kerafakod$, kerbev$, kerraf$

Public Sub scrinit(kepernyomod$)
  If kepernyomod$ = "T" Then
    form1.Label1.FontName = "Arial Narrow"
    form1.Label1.FontSize = 14
    form1.Label1.FontBold = True
    form1.Label1.ForeColor = RGB(130, 20, 0)
    form1.Label1.Top = 100: form1.Label1.Left = 2600
    form1.Label2.FontName = "Arial Narrow"
    form1.Label2.FontSize = 10
    form1.Label2.FontBold = False
    form1.Label2.Top = 7884: form1.Label2.Left = 30
    form1.Label3.FontName = "Arial Narrow"
    form1.Label3.FontSize = 9
    form1.Label3.FontBold = True
    form1.Label3.Top = 8100: form1.Label3.Left = 30
  Else
  End If
End Sub

Public Sub elsolap(objazon, rek$, txleft%, txtop%, langisor%)
  form1.MSFlexGrid1.Clear
  form1.MSFlexGrid1.Cols = 2
  form1.MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  form1.MSFlexGrid1.Font.Size = 8
  form1.MSFlexGrid1.ForeColorFixed = RGB(0, 0, 0)
  form1.MSFlexGrid1.BackColorFixed = RGB(254, 253, 224)
  form1.MSFlexGrid1.BackColorBkg = RGB(255, 240, 180)
  form1.Text2.BackColor = RGB(96, 125, 64)
  'form1.Text2.BackColor = RGB(170, 184, 131)
  form1.Text2.ForeColor = RGB(255, 255, 255)
  form1.Text2.Font.Name = "Microsoft Sans Serif"
  form1.Text2.Font.Size = 8
  form1.Text2.Font.Bold = True
  form1.Font.Name = "Microsoft Sans Serif"
  form1.Font.Size = 8
  w1& = obsorszama(objazon)
  w2& = OBJTAB(w1&).oba(1)
  odarab& = ABLTAB(w2&).adatsorsz(0)
  form1.MSFlexGrid1.Rows = 1
  form1.MSFlexGrid1.TextMatrix(0, 0) = langprg(langisor)
  form1.MSFlexGrid1.TextMatrix(0, 1) = langprg(langisor) + " " + langprg(langisor + 1)
  form1.MSFlexGrid1.ColAlignment(1) = 1
  mmax% = 12: hxmax% = 0
  mhmax% = 10
  For i1& = 1 To odarab&
    ne$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatnev)
    ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).attr)
    If InStr(ar$, "R") = 0 Then
      form1.MSFlexGrid1.Rows = form1.MSFlexGrid1.Rows + 1
      mh% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatho
      kp% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatkp
      form1.MSFlexGrid1.TextMatrix(form1.MSFlexGrid1.Rows - 1, 0) = ne$
      w3& = Len(ne$)
      hxh% = form1.TextWidth(ne$) + 100
      If hxmax% < hxh% Then hxmax% = hxh%
      form1.MSFlexGrid1.ColWidth(0) = hxmax%
      If mh% > mhmax% Then h1% = mh% * 120: mhmax% = mh% Else h1% = mhmax% * 120
      form1.MSFlexGrid1.ColWidth(1) = h1%
      form1.MSFlexGrid1.Width = mmax% * 120 + mhmax% * 120 + 70
      amezo$ = Mid$(rek$, kp%, mh%)
      form1.MSFlexGrid1.TextMatrix(form1.MSFlexGrid1.Rows - 1, 1) = amezo$
    End If
  Next
  form1.MSFlexGrid1.Height = (form1.MSFlexGrid1.Rows) * 230
  form1.Text2.Text = " " + ABLTAB(w2&).fejlec
  form1.Text2.Height = 234
  form1.Text2.Top = txtop
  form1.Text2.Left = txleft
  form1.MSFlexGrid1.Top = form1.Text2.Top + form1.Text2.Height
  form1.MSFlexGrid1.Left = form1.Text2.Left
  form1.MSFlexGrid1.Width = form1.MSFlexGrid1.ColWidth(0) + form1.MSFlexGrid1.ColWidth(1)
  form1.Text2.Width = form1.MSFlexGrid1.Width
  form1.Text2.Visible = True
  form1.MSFlexGrid1.Visible = True
End Sub

Public Sub elsotab(objazon$, mt$(), sordarab%, txleft, txtop)
  form1.MSFlexGrid2.Clear
  'form1.MSFlexGrid2.Cols = 2
  form1.MSFlexGrid1.Appearance = 0
  form1.MSFlexGrid2.Font.Name = "Microsoft Sans Serif"
  form1.MSFlexGrid2.Font.Size = 8
  form1.MSFlexGrid2.ForeColorFixed = RGB(0, 0, 0)
  form1.MSFlexGrid2.BackColorFixed = RGB(254, 253, 224)
  form1.MSFlexGrid2.BackColorBkg = RGB(255, 240, 180)
  form1.Text12.BackColor = RGB(96, 125, 64)
  form1.Text12.ForeColor = RGB(255, 255, 255)
  form1.Text12.Font.Name = "Microsoft Sans Serif"
  form1.Text12.Font.Size = 8
  form1.Text12.Font.Bold = True
  form1.Font.Name = "Microsoft Sans Serif"
  form1.Font.Size = 8
  If sordarab% > 20 Then
    darab% = 20
    form1.MSFlexGrid2.Rows = 21
  Else
    darab% = sordarab%
    form1.MSFlexGrid2.Rows = sordarab% + 1
  End If
  w1& = obsorszama(objazon)
  w2& = OBJTAB(w1&).oba(1)
  odarab& = ABLTAB(w2&).adatsorsz(0)
  form1.MSFlexGrid2.Cols = odarab& + 1
  For i1& = 1 To darab%
    form1.MSFlexGrid2.TextMatrix(i1&, 0) = Str$(i1&) + "." '+ langprg(67)
  Next
  gw& = 0
  '--- adatok feltöltése
  For i1& = 1 To darab%
    If Trim(mt$(i1&)) <> "" Then
      For i2& = 1 To odarab&
        mh% = ADATAB(ABLTAB(w2&).adatsorsz(i2&)).adatho
        kp% = ADATAB(ABLTAB(w2&).adatsorsz(i2&)).adatkp
        amezo$ = Mid$(mt$(i1&), kp%, mh%)
        form1.MSFlexGrid2.TextMatrix(i1&, i2&) = Trim$(amezo$)
      Next
    End If
  Next
  '--- fejlec
  form1.MSFlexGrid2.ColWidth(0) = 400
  For i1& = 1 To odarab&
    ne$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatnev)
    ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).attr)
    mh% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatho
    kp% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatkp
    form1.MSFlexGrid2.TextMatrix(0&, i1&) = ne$
    w3& = Len(ne$)
    If w3& > mh% Then h% = w3& * 110 Else h% = mh% * 110
    form1.MSFlexGrid2.ColWidth(i1&) = h%
    gw& = gw& + h%
    If InStr(ar$, "J") > 0 Then
      form1.MSFlexGrid2.ColAlignment(i1&) = 6
    Else
      form1.MSFlexGrid2.ColAlignment(i1&) = 1
    End If
    mtb$(i1&) = ar$: mho%(i1&) = mh%
    mesor%(i1&) = ABLTAB(w2&).adatsorsz(i1&)
  Next
  '--- méretek beállítása
  If objazon <> "PSHL" Then
    If txtop = 0 Then
      form1.Text12.Top = form1.Text2.Top
    Else
      If txtop = -1 Then
        form1.Text12.Top = form1.Text2.Top + form1.Text2.Height + form1.MSFlexGrid1.Height
      Else
        form1.Text12.Top = txtop
      End If
    End If
    If txleft = 0 Then
      form1.Text12.Left = form1.Text2.Left + form1.Text2.Width
    Else
      If txleft = -1 Then
        form1.Text12.Left = form1.Text2.Left
      Else
        form1.Text12.Left = txleft
      End If
    End If
    form1.MSFlexGrid2.Height = (form1.MSFlexGrid2.Rows) * 230
    form1.Text12.Text = " " + ABLTAB(w2&).fejlec
    form1.Text12.Height = 234
    form1.MSFlexGrid2.Top = form1.Text12.Top + form1.Text12.Height
    form1.MSFlexGrid2.Left = form1.Text12.Left
    oz% = 0: For i1& = 0 To odarab&: oz% = oz% + form1.MSFlexGrid2.ColWidth(i1&): Next
    form1.MSFlexGrid2.Width = oz%
    form1.Text12.Width = form1.MSFlexGrid2.Width
    form1.Text12.Visible = True
    form1.MSFlexGrid2.Visible = True
    form1.MSFlexGrid2.Visible = True
  Else
    form1.MSFlexGrid2.Top = 5950
    form1.MSFlexGrid2.Left = 300
    form1.MSFlexGrid2.Width = 11000
    form1.MSFlexGrid2.Height = 1400
    form1.MSFlexGrid2.Visible = True
    Gombok.BackColor = &H808080
  End If
End Sub


Public Sub kerekit510(osszeg@, kerekitettosszeg@, kerekites@, yfizmod$)
  '--- osszeg kerekitese 5-10 forintra
  kerekites@ = 0@
  osw1$ = ertszam(Str(osszeg@), 14, 2)
  If yfizmod$ = "K" Then
    '--- 5,10 forintra
    vegzodes@ = Abs(xval(Right$(osw1$, 4)))
    h1@ = xval(Mid$(irec$, 460, 4))
    h2@ = xval(Mid$(irec$, 464, 4))
    If h1@ = 0 Then h1@ = 2.5
    If h2@ = 0 Then h2@ = 7.5
    If vegzodes@ < h1@ Then
      kerekites@ = -vegzodes@
    Else
      If vegzodes@ < h2@ Then
        kerekites@ = 5 - vegzodes@
      Else
        kerekites@ = 10 - vegzodes@
      End If
    End If
    If osszeg@ < 0 Then kerekites@ = -kerekites@
  Else
    '--- forintra
    vegzodes@ = Abs(xval(Right$(osw1$, 2)))
    If vegzodes@ < 50 Then
      kerekites@ = -vegzodes@
    Else
      kerekites@ = 100 - vegzodes@
    End If
    kerekites@ = kerekites@ / 100: If osszeg@ < 0 Then kerekites@ = -kerekites@
  End If
  kerekitettosszeg@ = osszeg@ + kerekites@
End Sub
Public Function ltkeszlet@(rec$)
  ltkeszlet = xval(Mid$(rec$, 748, 14))
End Function
Public Function karcsere$(rec$, mit$, mire$)
  '--- karakter csere
  iw1% = Len(rec$)
  wb$ = rec$
  For iw91% = 1 To iw1%
    If Mid$(rec$, iw91%, 1) = mit$ Then Mid$(wb$, iw91%, 1) = mire$
  Next
  karcsere = wb$
End Function

Public Sub telkombe()
  '--- auwkom.par beolvasása, telepkom feltöltése
  Dim lparam$(20), lparamdb%
  tefi = FreeFile
  Open Left(auditorutvonal$, 1) + ":\auwin\auwkom.par" For Binary As #tefi
  tefim& = LOF(tefi)
  Close tefi
  If tefim& < 5 Then
    telkom.Ltvan = 0
  Else
    telkom.Ltvan = 1
    tefi = FreeFile
    Open Left(auditorutvonal$, 1) + ":\auwin\auwkom.par" For Input As #tefi
    Line Input #tefi, wt1$
    Call linpar(wt1$, lparam(), ";", lparamdb)
    telkom.Lipcim = Trim(lparam(1))
    telkom.Luser = Trim(lparam(2))
    telkom.Ljelszo = Trim(lparam(3))
    Line Input #tefi, wt1$
    Call linpar(wt1$, lparam(), ";", lparamdb)
    telkom.Ltelepdb = xval(lparam(1))
    telkom.Lsajtip = Trim(lparam(2))
    telkom.Lraktar = Left(Trim(lparam(3)) + "   ", 4)
    megvani% = 0
    For iw1% = 1 To telkom.Ltelepdb
      Line Input #tefi, wt1$
      Call linpar(wt1$, lparam(), ";", lparamdb)
      raki$ = Left(Trim(lparam(1)) + "   ", 4)
      If megvani% = 0 And raki$ = telkom.Lraktar Then
        telkom.Ltelindex = iw1%
        telkom.Ldirnev = Trim(lparam(2))
        If lparamdb > 2 Then telkom.Lsajkom = Trim(lparam(3))
        If lparamdb > 3 Then telkom.Ltavkom = Trim(lparam(4))
        megvani% = 1
      End If
      telkom.Ltelepek(iw1%) = raki$ + Left(Trim(lparam(2)) + Space(8), 8)
    Next
    Close tefi
  End If
End Sub

Public Function szokozkihagy$(s$)
  If s$ = "" Then szokozkihagy = "": Exit Function
  z$ = ""
  For i443% = 1 To Len(s$)
    If Mid$(s$, i443%, 1) <> " " Then z$ = z$ + Mid$(s$, i443%, 1)
  Next
  szokozkihagy = z$
End Function

Public Function minuszkihagy$(s$)
  If s$ = "" Then minuszkihagy = "": Exit Function
  z$ = ""
  For i443% = 1 To Len(s$)
    If Mid$(s$, i443%, 1) <> " " And Mid$(s$, i443%, 1) <> "-" Then z$ = z$ + Mid$(s$, i443%, 1)
  Next
  minuszkihagy = z$
End Function

Public Sub kulsoakciosar(ktrmrec$, dat$, valt%)
  '--- ha van hatályba lépõ vagy lejáró akció a dat$ napon
  valt% = 0
  If kulsobolt = 1 Then
    akckod$ = Mid$(ktrmrec$, 1, 15)
    akcrec$ = dbxkey("AKCT", akckod$)
    If akcrec$ <> "" Then
      For jui% = 1 To 5
        valt% = 0
        erar@ = xval(Mid$(ktrmrec$, 678, 14))
        akcelem$ = Mid$(akcrec$, (jui% - 1) * 140 + 100, 140)
        akcd1$ = Mid$(akcelem$, 23, 6)
        akcd2$ = Mid$(akcelem$, 29, 6)
        akcar$ = Mid$(akcelem$, 125, 10)
        If akcd1$ = dat$ Then
          '--- az akció a megadott napon lép hatályba
          valt% = 1
          Mid$(ktrmrec$, 678, 14) = Right(Space(14) + akcar$, 14)
          Exit Sub
        End If
        If novdat(akcd2$) = dat$ Then
          '--- az akció tegnap járt le
          valt% = -1
          Exit Sub
        End If
      Next
    End If
  End If
End Sub

Public Sub kulsoakcio(ktrmrec$)
  '--- ha van érvényes AKCT az abban szereplõ árat beírja a ktrmrec-be
  If kulsobolt = 1 Then
    akckod$ = Mid$(ktrmrec$, 1, 15)
    akcrec$ = dbxkey("AKCT", akckod$)
    If akcrec$ <> "" Then
      For jui% = 1 To 5
        erar@ = xval(Mid$(ktrmrec$, 678, 14))
        akcelem$ = Mid$(akcrec$, (jui% - 1) * 140 + 100, 140)
        akcd1$ = Mid$(akcelem$, 23, 6)
        akcd2$ = Mid$(akcelem$, 29, 6)
        akcar$ = Mid$(akcelem$, 125, 10)
        If xval(akcar$) <> 0 And xval(akcar$) < erar@ Then
          If dtm(akcd1$) <= dtm(maidatum$) And dtm(akcd2$) >= dtm(maidatum$) Then
            Mid$(ktrmrec$, 678, 14) = Right(Space(14) + akcar$, 14)
          End If
        End If
      Next
    End If
  End If
End Sub

Public Function adocdv%(adosz$)
  cdvo& = 0: cdvs$ = "9731973"
  For i888% = 1 To 7
    cdvo& = cdvo& + xval(Mid$(adosz$, i888%, 1)) * xval(Mid$(cdvs$, i888%, 1))
  Next
  aa$ = Right(Trim(Str(cdvo&)), 1)
  cdvk& = 10 - xval(aa$): If cdvk& = 10 Then cdvk& = 0
  cdv$ = Trim(Str(cdvk&))
  If cdv$ = Mid$(adosz$, 8, 1) Then adocdv = 1 Else adocdv = 0
End Function
Public Function szamszoveg$(szam@, wrapp%, tkjel$)
  '--- szám kiírása szöveggel
  Dim eg$(9), ti$(9), tizen$(9), uto$(9)
  eg$(0) = "": eg$(1) = "egy": eg$(2) = "kettõ": eg$(3) = "három": eg$(4) = "négy": eg$(5) = "öt": eg$(6) = "hat": eg$(7) = "hét": eg$(8) = "nyolc": eg$(9) = "kilenc"
  ti$(0) = "": ti$(1) = "tíz": ti$(2) = "húsz": ti$(3) = "harminc": ti$(4) = "negyven": ti$(5) = "ötven": ti$(6) = "hatvan": ti$(7) = "hetven": ti$(8) = "nyolcvan": ti$(9) = "kilencven"
  uto$(1) = "": uto$(2) = "ezer": uto(3) = "millió": uto$(4) = "miliárd": uto$(5) = "billió"
  For j313% = 1 To 9: tizen$(j313%) = ti$(j313%): Next
  tizen$(1) = "tizen": tizen$(2) = "huszon"
  sztring$ = Trim(Format(Abs(szam@), "###############"))
  sztl% = Len(sztring$)
  If sztl% = 0 Then szamszoveg = "nulla": Exit Function
  ox$ = ""
  For j314% = 1 To 5
    If sztring$ = "" Then Exit For
    If Len(sztring$) > 3 Then
      triad$ = Right(sztring$, 3)
      sztring$ = Left(sztring$, Len(sztring$) - 3)
    Else
      triad$ = sztring$
      sztring$ = ""
    End If
    '--- triad meghatározása
    tlen% = Len(triad$)
    qj1% = xval(Right(triad$, 1))
    Select Case tlen%
      Case 1: triszov$ = eg$(qj1%)
      Case 2
        qj2% = xval(Left(triad$, 1))
        If qj1% = 0 Then triszov$ = ti$(qj2%) Else triszov$ = tizen(qj2%) + eg$(qj1%)
      Case 3
        qj3% = xval(Left(triad$, 1))
        If qj3% = 0 Then triszov$ = "" Else triszov$ = eg$(qj3%) + "száz"
        qj2% = xval(Mid$(triad$, 2, 1))
        If qj1% = 0 Then triszov$ = triszov$ + ti$(qj2%) Else triszov$ = triszov$ + tizen(qj2%) + eg$(qj1%)
      Case Else
    End Select
    '--- szöveg hozzáadása
    If wrapp% = 1 Then koto$ = "- " Else koto$ = "-"
    If triszov$ <> "" Then ox$ = triszov$ + uto$(j314%) + koto$ + ox$
  Next
  If Right$(Trim(ox$), 1) = "-" Then ox$ = Left(Trim(ox$), Len(Trim(ox$)) - 1)
  'If szam@ < 0 Then szamszoveg$ = "mínusz " + ox$ Else szamszoveg$ = ox$
  If tkjel = "KT" Then
    If szam@ < 0 Then szamszoveg$ = "követel " + ox$ Else szamszoveg$ = "tartozik " + ox$
  Else
    If tkjel = "TK" Then
      If szam@ < 0 Then szamszoveg$ = "tartozik " + ox$ Else szamszoveg$ = "követel " + ox$
    Else
      If szam@ < 0 Then szamszoveg$ = "mínusz " + ox$ Else szamszoveg$ = ox$
    End If
  End If
End Function

Public Function tpontoz$(szam$)
  '--- tizedes vesszo lecserélése pontra
  oszam$ = szam$
  ppz% = InStr(oszam$, ",")
  If ppz% > 0 Then Mid$(oszam$, ppz%, 1) = "."
  tpontoz = oszam$
End Function

Public Function ekodol$(mit$, umod$)
  '--- mit$=inputstring, umod=K-kódol umod=D-dekódol
  ekodt1$ = "0123456789aábcdeéfgiíjklmnoóöõpqrstuúüûvwxyz\.-AÁBCDEÉFGHIÍJKLMNOÓÖÕPQRSTUÚÜÛVWXYZ "
  ekodt2$ = "gwHó\.Ín-ÉxöÁyFPéA0vBqWNu7C3áGpjÓYMmXOÛi1acz4ZÜfbíoLÕ dõ2U6VRDtlürTEeSKQÖIs8ûJkúÚ59"
  mire$ = mit$
  If umod$ = "K" Then
    '--- kódolás
    h% = Len(mit)
    For i81% = 1 To h%
      car$ = Mid$(mit$, i81%, 1)
      pz1% = InStr(ekodt1$, car$)
      If pz1% > 0 Then Mid$(mire$, i81%, 1) = Mid$(ekodt2$, pz1%, 1)
    Next
    zire$ = mire$
    For i81% = 1 To h%: Mid$(zire$, i81%, 1) = Mid$(mire$, h% - i81% + 1, 1): Next
    mire$ = zire$
  Else
    '--- dekódolás
    h% = Len(mit)
    zire$ = mit$
    For i81% = 1 To h%: Mid$(zire$, i81%, 1) = Mid$(mit$, h% - i81% + 1, 1): Next
    For i81% = 1 To h%
      car$ = Mid$(zire$, i81%, 1)
      pz1% = InStr(ekodt2$, car$)
      If pz1% > 0 Then Mid$(zire$, i81%, 1) = Mid$(ekodt1$, pz1%, 1)
    Next
    mire$ = zire$
  End If
  ekodol = mire$
End Function
Public Function sraktar$(szamlakezdet$)
  Select Case szamlakezdet$
    Case "MERK": sraktar = "0001"
    Case "MOHA": sraktar = "1001"
    Case "KOML": sraktar = "1002"
    Case "CSER": sraktar = "1003"
    Case "ADRI": sraktar = "1004"
    Case "ZEGE": sraktar = "1005"
    Case "SZVR": sraktar = "1006"
    Case "NAGY": sraktar = "1007"
    Case Else
  End Select
End Function
Public Function nevjavit$(xx$)
  zz$ = xx$
  poz% = InStr(zz$, "&")
  Do While poz% > 0
    Mid$(zz$, poz%, 1) = " "
    poz% = InStr(zz$, "&")
  Loop
  nevjavit = zz$
End Function

Public Sub mlog(azonosito$, szoveg$)
  '--- terminál LOG írása
  filog = FreeFile
  Open "c:\auwin\terminal.log" For Append As #filog
  srf$ = azonosito + "-" + ugyintezo$ + "-" + terminal$ + task$ + "-" + datki(maidatum) + "-" + Time$ + "\" + szoveg$
  Print #filog, srf$
  Close filog
End Sub
Public Function fajlista$(fejlec$, utvonal$, kiterj$)
  '--- konkret fájl kiválasztása
  Fajlist.Caption = fejlec$
  Fajlist.File1.Path = utvonal$
  Fajlist.File1.Pattern = kiterj$
  Fajlist.Show vbModal
  fajlista$ = kosarfajl$
End Function
Public Sub tkosar(cikszam$(), vokod$(), vdb&)
  '--- termékek kijelölése kosárral
  vdb& = 0: Kosar.Text1.Text = ""
  Kosar.List1.Clear
  Kosar.Show vbModal
  If rogzites% <> 0 Then
    vdb& = Kosar.List1.ListCount
    If vdb& > 0 Then
      For i93% = 1 To vdb&
        ReDim Preserve cikszam(1 To vdb&)
        ReDim Preserve vokod(1 To vdb&)
        elem$ = Kosar.List1.List(i93% - 1)
        vokod(i93%) = Mid$(elem$, 1, 13)
        cikszam(i93%) = Mid$(elem$, 75, 15)
      Next
    End If
  Else
    vdb& = 0
  End If
End Sub

Public Function parfilbe$(filn$)
  '--- paraméterfájl beolvasása
  fio = FreeFile
  Open auditorutvonal$ + filn$ For Binary Shared As #fio
  fiom& = LOF(fio)
  Close fio
  If fiom& > 2 Then
    fio = FreeFile
    Open auditorutvonal$ + filn$ For Input Shared As #fio
    Line Input #fio, xaxy$
    Close fio
    parfilbe$ = xaxy$
  Else
    parfilbe$ = ""
  End If
End Function

Public Sub waitsec(mperc%)
  '--- várakozás mperc másodpercig
  ssec = Timer
  Do While Timer < ssec + mperc%
    DoEvents
  Loop
End Sub

Public Function holnap$(dat$)
  '--- a következõ munkanap meghatározása
  xd$ = novdat(dat$)
  dx$ = "20" + Left(xd$, 2) + "." + Mid$(xd$, 3, 2) + "." + Mid$(xd$, 5, 2) + "."
  naps% = WeekDay(dx$, vbSunday)
  If naps% = 1 Then
    holnap = novdat(xd$)
  Else
    holnap = xd$
  End If
End Function
Public Function fsbford$(ertbrec$, kszcrec$)
  '--- ERTB objektum árfordítása KSZB szerkezetre
  '--- kimenet KSZB rekord
  sta$ = Mid$(ertbrec$, 26, 1)
  Select Case sta$
    Case "L", "S", "J": o$ = Space(300)
    Case Else: fsbford$ = "": Exit Function
  End Select
  If sta$ = "S" Then
    Mid$(o$, 1, 10) = Mid$(ertbrec$, 135, 10)        '--- számlaszám
    Mid$(o$, 250, 10) = Mid$(ertbrec$, 155, 10)      '--- szállítólevél szám
    Mid$(o$, 76, 2) = "BS"
  Else
    Mid$(o$, 1, 10) = Mid$(ertbrec$, 155, 10)        '--- száll.levél szám
    Mid$(o$, 250, 10) = Mid$(ertbrec$, 155, 10)      '--- száll.levél szám
    Mid$(o$, 76, 2) = "SL"
  End If
  Mid$(o$, 21, 8) = Mid$(ertbrec$, 193, 8)           '--- számlázó
  Mid$(o$, 29, 6) = Mid$(ertbrec$, 20, 6)            '--- rögzítés kelte
  Mid$(o$, 35, 1) = Mid$(ertbrec$, 201, 1)           '--- sztornó jel
  Mid$(o$, 36, 6) = Mid$(ertbrec$, 202, 6)           '--- sztornó kelte
  Mid$(o$, 42, 8) = Mid$(ertbrec$, 208, 8)           '--- sztornózta
  Mid$(o$, 282, 15) = Mid$(ertbrec$, 225, 10)        '--- sztornószámla száma
  Mid$(o$, 53, 7) = Mid$(ertbrec$, 175, 7)           '--- vevõ iktató folyószámlában
  Mid$(o$, 61, 15) = Mid$(kszcrec$, 16, 15)          '--- partner kód
  Mid$(o$, 260, 15) = Mid$(ertbrec$, 105, 15)        '--- száll.cím kód
  Mid$(o$, 275, 7) = Mid$(ertbrec$, 1, 7)            '--- megrendelés iktató
  Mid$(o$, 237, 8) = Mid$(ertbrec$, 216, 8)          '--- üzletkötõ
  Mid$(o$, 78, 6) = Mid$(ertbrec$, 39, 6)            '--- számla kelte
  Mid$(o$, 84, 6) = Mid$(ertbrec$, 39, 6)            '--- teljesítés kelte
  Mid$(o$, 231, 6) = Mid$(ertbrec$, 39, 6)           '--- könyvelés kelte
  Mid$(o$, 90, 6) = Mid$(ertbrec$, 39, 6)            '--- fizetési határidõ
  fsbford$ = o$
End Function
Public Function fstford$(ertbrec$, erttrec$, ktrmrec$, gyujtogongyoleg$, refarkell$)
  '--- ERTT objektum átfordítása KSZT szerkezetre
  '--- gyujtogongyoleg="I" gyûjtõgöngyöleg rekord kell
  '--- refarkell=I esetén a nyilv.ár helyett a ref.ár kell
  '--- kimenet KSZT rekord
  sta$ = Mid$(ertbrec$, 26, 1)
  Select Case sta$
    Case "L", "S", "J": o$ = Space(170)
    Case Else: fstford$ = "": Exit Function
  End Select
  If sta$ = "S" Then
    Mid$(o$, 1, 10) = Mid$(ertbrec$, 135, 10)    '--- számlaszám
  Else
    Mid$(o$, 1, 10) = Mid$(ertbrec$, 155, 10)    '--- száll.levél száma
  End If
  Mid$(o$, 130, 6) = Mid$(ertbrec$, 39, 6)       '--- teljesítés kelte
  Mid$(o$, 14, 1) = Mid$(erttrec$, 23, 1)        '--- sztorno jel
  Mid$(o$, 107, 15) = Mid$(erttrec$, 28, 15)     '--- termék kód
  If gyujtogongyoleg$ = "I" Then
    menny@ = xval(Mid$(erttrec$, 91, 4))
    Mid$(o$, 21, 12) = ertszam(Str(menny@), 12, 0) '--- mennyiség
    If refarkell$ = "I" Then
      If Len(ktrmrec$) = 99 Then
        Mid$(o$, 39, 12) = Mid$(ktrmrec$, 70, 12)      '--- referencia ár
      Else
        Mid$(o$, 39, 12) = Mid$(ktrmrec$, 1276, 12)    '--- referencia ár
      End If
    Else
      If Len(ktrmrec$) = 99 Then
        Mid$(o$, 39, 12) = Mid$(ktrmrec$, 56, 12)     '--- nyilv.ár
      Else
        Mid$(o$, 39, 12) = Mid$(ktrmrec$, 554, 12)     '--- nyilv.ár
      End If
    End If
    Mid$(o$, 69, 12) = Mid$(erttrec$, 95, 12)      '--- eladási ár
    If Len(ktrmrec$) = 99 Then
      Mid$(o$, 81, 2) = Mid$(ktrmrec$, 82, 2)       '--- ÁFA kód
    Else
      Mid$(o$, 81, 2) = Mid$(ktrmrec$, 706, 2)       '--- ÁFA kód
    End If
  Else
    Mid$(o$, 21, 12) = Mid$(erttrec$, 67, 12)      '--- mennyiség
    If refarkell$ = "I" Then
      If Len(ktrmrec$) = 99 Then
        Mid$(o$, 39, 12) = Mid$(ktrmrec$, 70, 12)      '--- referencia ár
      Else
        Mid$(o$, 39, 12) = Mid$(ktrmrec$, 1276, 12)    '--- referencia ár
      End If
    Else
      Mid$(o$, 39, 12) = Mid$(erttrec$, 79, 12)      '--- nyilv.ár
    End If
    Mid$(o$, 69, 12) = Mid$(erttrec$, 43, 12)      '--- eladási ár
    If Len(ktrmrec$) = 99 Then
      Mid$(o$, 81, 2) = Mid$(ktrmrec$, 82, 2)       '--- ÁFA kód
    Else
      Mid$(o$, 81, 2) = Mid$(ktrmrec$, 706, 2)       '--- ÁFA kód
    End If
  End If
  'Mid$(o$, 83, 8) = ellszla$
  'Mid$(o$, 91, 16) = szervegys + munkaszam
  Mid$(o$, 150, 8) = Mid$(ertbrec$, 216, 8)      '--- üzletkötõ
  If ikonfrec$ <> "" Then
    If Len(ktrmrec$) = 99 Then
      rako$ = Mid$(ktrmrec$, 84, 1)                '--- raktározási kód
    Else
      rako$ = Mid$(ktrmrec$, 907, 1)               '--- raktározási kód
    End If
    Select Case rako$
      Case "I": rak$ = Mid$(ikonfrec$, 288, 4)
      Case "E": rak$ = Mid$(ikonfrec$, 280, 4)
      Case "D": rak$ = Mid$(ikonfrec$, 280, 4)
      Case "V": rak$ = Mid$(ikonfrec$, 284, 4)
      Case "G": rak$ = Mid$(ikonfrec$, 280, 4)
      Case Else: rak$ = Mid$(ikonfrec$, 280, 4)
    End Select
  Else
    rak$ = Mid$(irec$, 634, 4)
  End If
  Mid$(o$, 146, 4) = rak$                        '--- raktár kód
  If Len(ktrmrec$) = 99 Then
    gdkod$ = Trim(Mid$(ktrmrec$, 85, 15))
  Else
    gdkod$ = Trim(Mid$(ktrmrec$, 1067, 15))
  End If
  If gdkod$ <> "" Then
    gelar@ = xval(Mid$(erttrec$, 55, 12))
    gbear@ = xval(Mid$(erttrec$, 113, 8))
    Mid$(o$, 51, 1) = "G"
    Mid$(o$, 52, 9) = ertszam(Str(gelar@), 9, 1) '--- göngyöleg elad.ár
    Mid$(o$, 61, 8) = ertszam(Str(gbear@), 8, 1) '--- göngyöleg besz.ár
  End If
  fstford$ = o$
End Function

Public Function fbford$(ertbrec$, kszcrec$)
  '--- ERTB objektum átfordítása KKBZ szerkezetre
  '--- kimenet KKBZ rekord
  o$ = Space(140)
  Mid$(o$, 1, 7) = Mid$(ertbrec$, 1, 7)         '--- biz.iktató
  Mid$(o$, 8, 6) = Mid$(ertbrec$, 39, 6)        '--- dátum
  Mid$(o$, 14, 40) = "Értékesítés"              '--- szöveg
  sta$ = Mid$(ertbrec$, 26, 1)
  Select Case sta$
    Case "L", "S": Mid$(o$, 54, 1) = "E"        '--- biz.típus
    Case "J", "R": Mid$(o$, 54, 1) = "R"        '--- biz.típus
    Case Else: Mid$(o$, 54, 1) = "X"            '--- biz.típus
  End Select
  Mid$(o$, 55, 15) = Mid$(kszcrec$, 16, 15)     '--- partner kód
  Mid$(o$, 70, 8) = Mid$(ertbrec$, 216, 8)      '--- üzletkötõ
  If sta$ = "S" Then
    Mid$(o$, 107, 15) = Mid$(ertbrec$, 135, 15) '--- számlaszám
  End If
  If sta$ = "L" Or sta$ = "R" Then
    Mid$(o$, 107, 15) = Mid$(ertbrec$, 155, 15) '--- száll.levél szám
  End If
  Mid$(o$, 78, 6) = Mid$(ertbrec$, 20, 6)       '--- feldolgozás kelte
  Mid$(o$, 84, 8) = Mid$(ertbrec$, 193, 8)      '--- ügyintézõ
  Mid$(o$, 92, 1) = Mid$(ertbrec$, 201, 1)      '--- sztornó jel
  Mid$(o$, 93, 6) = Mid$(ertbrec$, 202, 6)      '--- sztornó kelte
  Mid$(o$, 99, 8) = Mid$(ertbrec$, 208, 8)      '--- sztornózta
  fbford$ = o$
End Function

Public Function frford$(ertbrec$, erttrec$, ktrmrec$, gyujtogongyoleg$, refarkell$)
  '--- ERTT objektum átfordítása KKFT szerkezetre
  '--- gyujtogongyoleg="I" gyûjtõgöngyöleg rekord kell, ktrmrec-ben a gy.göngyöleg rekordja
  '--- refarkell=I esetén a besz.ár helyett ref.ár kell
  '--- kimenet KKFT rekord
  o$ = Space(130)
  Mid$(o$, 8, 7) = Mid$(ertbrec$, 1, 7)             '--- biz.iktató
  Mid$(o$, 15, 6) = Mid$(ertbrec$, 39, 6)           '--- dátum
  sta$ = Mid$(ertbrec$, 26, 1)
  Select Case sta$
    Case "L", "S": mozg$ = "003"
    Case "J": mozg$ = "888"
    Case "R": mozg$ = "006"
    Case Else: mozg$ = "   "
  End Select
  Mid$(o$, 21, 3) = mozg$                           '--- mozgás
  If ikonfrec$ <> "" Then
    If Len(ktrmrec$) = 99 Then
      rako$ = Mid$(ktrmrec$, 84, 1)
    Else
      rako$ = Mid$(ktrmrec$, 907, 1)
    End If
    Select Case rako$
      Case "I": rak$ = Mid$(ikonfrec$, 288, 4)
      Case "E": rak$ = Mid$(ikonfrec$, 280, 4)
      Case "D": rak$ = Mid$(ikonfrec$, 280, 4)
      Case "V": rak$ = Mid$(ikonfrec$, 284, 4)
      Case "G": rak$ = Mid$(ikonfrec$, 280, 4)
      Case Else: rak$ = Mid$(ikonfrec$, 280, 4)
    End Select
  Else
    rak$ = Mid$(irec$, 634, 4)
  End If
  Mid$(o$, 24, 4) = rak$                              '--- raktár kód
  If gyujtogongyoleg$ = "I" Then
    Mid$(o$, 36, 15) = Mid$(ktrmrec$, 1, 15)          '--- termék kód
    meny@ = -xval(Mid$(erttrec$, 91, 4))
    menn$ = ertszam(Str(meny@), 12, 2)
    Mid$(o$, 71, 12) = menn$                          '--- mennyiség
    If Len(ktrmrec$) = 99 Then
      nyar@ = xval(Mid$(ktrmrec$, 54, 14))
    Else
      nyar@ = xval(Mid$(ktrmrec$, 552, 14))
    End If
    If refarkell$ = "I" Then
      If Len(ktrmrec$) = 99 Then
        nyar@ = xval(Mid$(ktrmrec$, 68, 14))
      Else
        nyar@ = xval(Mid$(ktrmrec$, 1274, 14))
      End If
    Else
      If Len(ktrmrec$) = 99 Then
        nyar@ = xval(Mid$(ktrmrec$, 54, 14))
      Else
        nyar@ = xval(Mid$(ktrmrec$, 552, 14))
      End If
    End If
    bszr$ = ertszam(Str(nyar@), 12, 2)
    Mid$(o$, 59, 12) = bszr$                          '--- nyilv.ár
  Else
    Mid$(o$, 36, 15) = Mid$(erttrec$, 28, 15)         '--- termék kód
    If refarkell = "I" Then
      If Len(ktrmrec$) = 99 Then
        nyar@ = xval(Mid$(ktrmrec$, 68, 14))
      Else
        nyar@ = xval(Mid$(ktrmrec$, 1274, 14))
      End If
    Else
      nyar@ = xval(Mid$(erttrec$, 79, 12))
    End If
    bszr$ = ertszam(Str(nyar@), 12, 2)
    Mid$(o$, 59, 12) = bszr$                          '--- nyilv.ár
    meny@ = -xval(Mid$(erttrec$, 67, 12))
    menn$ = ertszam(Str(meny@), 12, 2)
    Mid$(o$, 71, 12) = menn$                          '--- mennyiség
    If Len(ktrmrec$) = 99 Then
      gdkod$ = Trim(Mid$(ktrmrec$, 85, 15))
    Else
      gdkod$ = Trim(Mid$(ktrmrec$, 1067, 15))
    End If
    If gdkod$ <> "" Then
      gbear@ = xval(Mid$(erttrec$, 113, 8))
      Mid$(o$, 116, 1) = "G"                          '--- tapadó göngyöleg jele
      Mid$(o$, 117, 10) = ertszam(Str(gbear@), 8, 1)  '--- göngyöleg nyilv ára
    End If
  End If
  frford$ = o$
End Function

Public Function aruafa@(afakod$, egysar@)
  '--- ÁFA meghatározása
  afrec$ = dbxkey("PAFA", afakod$)
  afkul@ = xval(Mid$(afrec$, 33, 6))
  aruafa@ = egysar@ * (afkul@ / 100)
End Function

Public Sub adatfilbe(objazon$, akp%, aho%, sorted%, skp%, sho%)
  '--- adatfájl részletének beolvasása tömbbe
  '--- adatfiltomb adatcimtomb adatelemszam&
  '--- akp% kezdet  aho% hossz
  '--- sorted=1 rendezett, egyébként nem
  '--- skp%,sho% a rendezõ mezõ kezdete és hossza
  '--- filterek alapján szûr
  Dim rel$(20), xtol%(20), xig%(20), xert1$(20), xert2$(20), param$(20), paramdb%
  If filterdarab% > 0 Then
    For i17% = 1 To filterdarab
      Call linpar(filterek(i17%), param(), "\", paramdb%)
      rel$(i17%) = param(1)
      xtol%(i17%) = Val(param(2))
      xig%(i17%) = Val(param(3))
      xert1$(i17%) = param(4)
      If paramdb% > 4 Then
        xert2$(i17%) = param(5)
      End If
    Next
  End If
  adatelemszam& = 0
  ob% = obsorszama(objazon)
  dbxn$ = dbxneve(objazon)
  w1% = OBJTAB(ob%).obi(1)
  indn$ = RTrim$(INDTAB(w1%).indnev)
  ih& = ADATAB(INDTAB(w1%).adatsorsz).adatho + 5
  rh& = OBJTAB(ob%).rekhossz
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  rc& = Int(LOF(ndfi) / ih&)
  If rc& = 0 Then Exit Sub
  For iti& = 1 To rc&
    DoEvents
    Get #ndfi, (iti& - 1) * ih& + 1, cci&
    r$ = Space(rh&)
    Get #dbfi, cci& + 9, r$
    If filterdarab% > 0 Then
      kell% = 1
      For i17% = 1 To filterdarab%
        relac$ = rel(i17%)
        Select Case relac$
          Case "="
            If Mid$(r$, xtol(i17%), xig(i17%)) <> xert1$(i17%) Then kell% = 0: Exit For
          Case "<>"
            If Mid$(r$, xtol(i17%), xig(i17%)) = xert1$(i17%) Then kell% = 0: Exit For
          Case "<"
            If Mid$(r$, xtol(i17%), xig(i17%)) >= xert1$(i17%) Then kell% = 0: Exit For
          Case ">"
            If Mid$(r$, xtol(i17%), xig(i17%)) <= xert1$(i17%) Then kell% = 0: Exit For
          Case "INT"
            If Mid$(r$, xtol(i17%), xig(i17%)) < xert1$(i17%) Or Mid$(r$, xtol(i17%), xig(i17%)) > xert2$(i17%) Then kell% = 0: Exit For
          Case "INSTR"
            If InStr(konver(Mid$(r$, xtol(i17%), xig(i17%))), konver(xert1$(i17%))) = 0 Then kell% = 0: Exit For
          Case Else
        End Select
      Next
    Else
      kell% = 1
    End If
    If kell% = 1 Then
      adatelemszam& = adatelemszam& + 1
      ReDim Preserve adatfiltomb$(1 To adatelemszam&)
      ReDim Preserve adatcimtomb&(1 To adatelemszam&)
      If akp% = 0 Then
        adatfiltomb(adatelemszam&) = r$
      Else
        If aho% = 0 Then
          adatfiltomb(adatelemszam&) = Mid$(r$, akp%)
        Else
          adatfiltomb(adatelemszam&) = Mid$(r$, akp%, aho%)
        End If
      End If
      adatcimtomb(adatelemszam&) = cci&
    End If
  Next
  Close dbfi: Close ndfi
  If sorted% = 1 Then
    If skp% = 0 Or sho% = 0 Then
      Call qsort(adatfiltomb(), adatcimtomb(), adatelemszam&, "N")
    Else
      Call qsortr(adatfiltomb(), adatcimtomb(), adatelemszam&, "N", skp%, sho%)
    End If
  End If
End Sub

Public Function netto@(bruttoar@, termrec$)
  '--- netto ár meghatározása bruttóból
  If Len(termrec$) = 2 Then
    afakod$ = termrec$: If afakod$ = "99" Then afakod$ = "00"
  Else
    afakod$ = Mid$(termrec$, 706, 2)
  End If
  afrec$ = dbxkey("PAFA", afakod$)
  afkulcs@ = xval(Mid$(afrec$, 33, 6))
  netar@ = ertszam(Str(bruttoar@ / (1 + afkulcs@ / 100)), 10, 2)
  netto@ = netar@
End Function

Public Function afbr@(nettoar@, afakod$)
  '--- áfás bruttó ár meghatásozása 2 tizedes
  If afakod$ = "99" Then afakod$ = "00"
  afkulcs@ = xval(afakod$)
  fogyar@ = ertszam(Str(nettoar@ * (1 + afkulcs@ / 100)), 10, 2)
  afbr@ = fogyar@
End Function

Public Function afasbrutto@(nettoar@, termrec$)
  '--- áfás bruttó ár meghatásozása
  If Len(termrec$) = 6 Then
    '--- afa kulcs
    afkulcs@ = xval(termrec$)
    fogyar@ = ertszam(Str(nettoar@ * (1 + afkulcs@ / 100)), 10, 0)
    afasbrutto@ = fogyar@
  Else
    If Len(termrec$) = 2 Then
      afakod$ = termrec$: If afakod$ = "99" Then afakod$ = "00"
    Else
      afakod$ = Mid$(termrec$, 706, 2)
    End If
    afrec$ = dbxkey("PAFA", afakod$)
    afkulcs@ = xval(Mid$(afrec$, 33, 6))
    fogyar@ = ertszam(Str(nettoar@ * (1 + afkulcs@ / 100)), 10, 0)
    afasbrutto@ = fogyar@
  End If
End Function
Public Function torolvas$(objazon$, kulcs$, xqto%, xqhossz%)
  If Trim(kulcs$) = "" Then torolvas = "": Exit Function
  torox$ = dbxkey(objazon, kulcs)
  If torox$ <> "" Then
    If xqto% = 0 Then
      torolvas = torox$
    Else
      torolvas = Mid$(torox$, xqto%, xqhossz%)
    End If
  Else
    torolvas = ""
  End If
End Function

Public Function banosz@(oz$)
  '--- bankból összeg konvertálása
  blh% = Len(oz$): eloj$ = " "
  If blh% = 0 Then banosz@ = 0: Exit Function
  boz$ = ""
  For i88% = 1 To blh%
    c$ = Mid$(oz$, i88%, 1)
    If c$ = "." Then
    Else
      If c$ = "," Then c$ = "."
      If c$ = "-" Then
        eloj$ = "-"
      Else
        boz$ = boz$ + c$
      End If
    End If
  Next
  boo@ = xval(boz$)
  If eloj$ = "-" Then banosz = -boo@ Else banosz@ = boo@
End Function
Public Function komirak%(raktarkod$)
  '--- 1-komissiózó raktár, 0 nem
  For ww1% = 1 To komiraktdb%
    If komirakt$(ww1%) = raktarkod$ Then komirak = 1: Exit Function
  Next
  komirak = 0
End Function

Public Sub kominit()
  Dim pa$(30)
  '--- komissiózáshoz használható raktárak beolvasása
  If ikonfrec$ = "" Then komiraktdb = 0: Exit Sub
  komiraktdb = 0
  pas$ = Trim(Mid$(ikonfrec$, 220, 60))
  Call linpar(pas$, pa$(), ",", padb%)
  If padb% = 0 Then
    komiraktdb = 0: Exit Sub
  Else
    komiraktdb = padb%
    For kid% = 1 To komiraktdb%
      komirakt(kid%) = Left(pa(kid%) + "    ", 4)
    Next
  End If
End Sub

Public Sub felvalt(menny@, db&, valtoszam@, csomag&, tored@)
  '--- egyszerû felváltá
  If valtoszam@ <> 0 Then
    If menny@ <> 0 Then
      csomag& = menny@ \ valtoszam@
      tored@ = menny@ Mod valtoszam@
    Else
      If db& <> 0 Then
        csomag& = db& \ valtoszam@
        tored@ = db& Mod valtoszam@
      Else
        csomag& = 0
        tored@ = 0
      End If
    End If
  Else
    csomag& = 0: tored@ = 0
  End If
End Sub

Public Sub kiszvalt(menny@, egyscsom&, qkarton&, qsor&, raklap&, toredek@, torede@, toredk@, toreds@, arurec$)
  '--- kiszerelések átváltása
  If menny@ > 0 Then
    valtoe@ = xval(Mid$(arurec$, 1219, 7))
    toredek@ = 0
    If valtoe@ <> 0 Then
      Call felvalt(menny@, 0&, valtoe@, egyscsom&, toredek@)
    Else
      valtoe@ = 1: egyscsom& = menny@: toredek@ = 0
    End If
    valtok@ = xval(Mid$(arurec$, 1226, 7))
    If valtok@ <> 0 Then
      Call felvalt(0@, egyscsom&, valtok@, qkarton&, torede@)
    Else
      valtok@ = 1: qkarton& = egyscsom&: torede@ = 0
    End If
    valtos@ = xval(Mid$(arurec$, 1233, 7))
    If valtos@ <> 0 Then
      Call felvalt(0@, qkarton&, valtos@, qsor&, toredk@)
    Else
      valtos@ = 1: qsor& = qkarton: toredk@ = 0
    End If
    valtor@ = xval(Mid$(arurec$, 1240, 7))
    If valtor@ <> 0 Then
      Call felvalt(0@, qsor&, valtor@, raklap&, toreds@)
    Else
      valtor@ = 0: raklap& = 0: toreds@ = 0
    End If
  Else
    valtoe@ = xval(Mid$(arurec$, 1219, 7)): If valtoe@ = 0 Then valtoe@ = 1
    valtok@ = xval(Mid$(arurec$, 1226, 7)): If valtok@ = 0 Then valtok@ = 1
    valtos@ = xval(Mid$(arurec$, 1233, 7)): If valtos@ = 0 Then valtos@ = 1
    valtor@ = xval(Mid$(arurec$, 1240, 7)): If valtor@ = 0 Then valtor@ = 1
    If egyscsom& <> 0 Then
      menny@ = egyscsom * valtoe@
    Else
      If qkarton& <> 0 Then
        menny@ = qkarton& * valtok@ * valtoe@
      Else
        If qsor& <> 0 Then
          menny@ = qsor& * valtos@ * valtok@ * valtoe@
        Else
          If raklap& <> 0 Then
            menny@ = raklap& * valtor@ * valtos@ * valtok@ + valtoe@
          Else
          End If
        End If
      End If
    End If
  End If
End Sub

Public Sub arkalk(rec$, volt%)
  '--- azonnali automatikus árkalkuláció
  volt% = 0
  tcs$ = Mid$(rec$, 438, 4)
  If Trim(tcs$) <> "" Then
    tcsrec$ = dbxkey("KCSP", tcs$)
    If tcsrec$ <> "" Then
      If Mid$(tcsrec$, 68, 1) = "A" Then
        hatm$ = Mid$(tcsrec$, 68, 1)
        hatd$ = Mid$(tcsrec$, 69, 6)
        kalkb$ = Mid$(tcsrec$, 67, 1)
        tizd$ = Mid$(tcsrec$, 155, 1)
        Select Case kalkb$
          Case "R": bazis@ = xval(Mid$(rec$, 1274, 14))
          Case "N": bazis@ = xval(Mid$(rec$, 552, 14))
          Case "U": bazis@ = xval(Mid$(rec$, 566, 14))
          Case Else: bazis@ = 0
        End Select
        If bazis@ <> 0 Then
          For j% = 1 To 8
            If Mid$(tcsrec$, (j% - 1) * 10 + 75, 10) <> Space(10) Then
              regar@ = xval(Mid$(rec$, (j% - 1) * 14 + 580, 14))
              hk@ = xval(Mid$(tcsrec$, (j% - 1) * 10 + 75, 10))
              ujar@ = bazis@ * (100 + hk@) / 100
              Select Case tizd$
                Case "0": uja$ = ertszam(Str(ujar@), 14, 0)
                Case "1": uja$ = ertszam(Str(ujar@), 14, 1)
                Case "2": uja$ = ertszam(Str(ujar@), 14, 2)
                Case Else: uja$ = ertszam(Str(ujar@), 14, 0)
              End Select
              ujar@ = xval(uja$)
              If ujar@ <> 0 And regar@ <> ujar@ Then Mid$(rec$, (j% - 1) * 14 + 580, 14) = uja$: volt% = 1
            End If
          Next
        End If
      End If
    End If
  End If
End Sub

Public Function levelkelte$()
  levelkelte = Trim(Mid$(irec$, 103, 30)) + ", " + datki(maidatum)
End Function

Public Function houtnap$(dat$)
  '--- hó ulsó napja
  ho% = xval(Mid$(dat$, 3, 2))
  Select Case ho%
    Case 1, 3, 5, 7, 8, 10, 12
      utn% = 31
    Case 2
      If ev% Mod 4 = 0 Then utn% = 29 Else utn% = 28
    Case 4, 6, 9, 11
      utn% = 30
    Case Else
  End Select
  houtnap$ = Mid$(dat$, 1, 4) + Right("00" + Trim(Str(utn%)), 2)
End Function
Public Sub qsortr(sortkulcs$(), sortcimek&(), elemszam&, irany$, vkp%, vmh%)
  '--- címvektor rendezése
  '--- irany=N-növekvõ C-csökkenõ
  Dim qpoz%(4000), qosort$(400000), qscim&(400000)
  yqm$ = "K"
  If Len(irany$) = 2 Then
    If Mid$(irany$, 2, 1) = "N" Then yqm$ = "N"
  End If
  If Left(irany$, 1) = "N" Then qhossz& = 500 Else ghossz% = 100
  qhossz& = 100
  If elemszam& < 2 Then Exit Sub
  elemhossz% = vmh% 'Len(sortkulcs(1))
  If yqm$ = "K" Then
    minima$ = String(elemhossz%, Chr$(255))
    maxima$ = String(elemhossz%, Chr$(0))
  Else
    minima$ = String(elemhossz%, "9")
    maxima$ = -String(elemhossz% - 1, "9")
  End If
  qsorok& = Int(elemszam& / qhossz&): qmar% = elemszam& Mod qhossz&
  If qmar% > 0 Then qsorok& = qsorok& + 1: utel% = qmar% Else utel% = qhossz&
  '--- rendezések
  For qi& = 1 To qsorok&
    If qi& = qsorok& Then qeli% = utel% Else qeli% = qhossz&
    GoSub q1sort
  Next
  '--- merge
  If qsorok& > 1 Then
    For qi& = 1 To qsorok&: qpoz(qi&) = 1: Next
    Do
      If Left(irany$, 1) = "N" Then
        '--- növekvõ
        qmin$ = minima$
        For qii& = 1 To qsorok&
          If qii& = qsorok& Then qeli% = utel% Else qeli% = qhossz&
          If qpoz(qii&) <= qeli% Then
            qhas$ = Mid$(sortkulcs((qii& - 1) * qhossz& + qpoz(qii&)), vkp%, vmh%)
            If yqm$ = "K" Then
              If qhas$ < qmin$ Then
                qminsor& = qii&: qmin$ = qhas$
              End If
            Else
              If xval(qhas$) < xval(qmin$) Then
                qminsor& = qii&: qmin$ = qhas$
              End If
            End If
          End If
        Next
      Else
        '--- csökkenõ
        qmax$ = maxima$
        For qii& = 1 To qsorok&
          If qii& = qsorok& Then qeli% = utel% Else qeli% = qhossz&
          If qpoz(qii&) <= qeli% Then
            qhas$ = Mid$(sortkulcs((qii& - 1) * qhossz& + qpoz(qii&)), vkp, vmh)
            If yqm$ = "K" Then
              If qhas$ > qmax$ Then
                qminsor& = qii&: qmax$ = qhas$
              End If
            Else
              If xval(qhas$) > xval(qmax$) Then
                qminsor& = qii&: qmax$ = qhas$
              End If
            End If
          End If
        Next
      End If
      qdb& = qdb& + 1
      qosort(qdb&) = sortkulcs((qminsor& - 1) * qhossz& + qpoz(qminsor&))
      qscim&(qdb&) = sortcimek&((qminsor& - 1) * qhossz& + qpoz(qminsor&))
      qpoz(qminsor&) = qpoz(qminsor&) + 1
    Loop While qdb& < elemszam&
    For qqi& = 1 To elemszam&
      sortkulcs(qqi&) = qosort(qqi&)
      sortcimek(qqi&) = qscim(qqi&)
    Next
  End If
Exit Sub
q1sort:
  '--- qeli rendezése
  ok% = 0
  qeltol& = (qi& - 1) * qhossz&
  Do While ok% = 0
    ok% = 1
    For qii& = 1 To qeli% - 1
      csere% = 0
      If Left(irany$, 1) = "N" Then
        If yqm$ = "K" Then
          If Mid$(sortkulcs(qeltol& + qii&), vkp, vmh) > Mid(sortkulcs(qeltol& + qii& + 1), vkp, vmh) Then csere% = 1
        Else
          If xval(Mid$(sortkulcs(qeltol& + qii&), vkp, vmh)) > xval(Mid(sortkulcs(qeltol& + qii& + 1), vkp, vmh)) Then csere% = 1
        End If
      Else
        If yqm$ = "K" Then
          If Mid(sortkulcs(qeltol& + qii&), vkp, vmh) < Mid(sortkulcs(qeltol& + qii& + 1), vkp, vmh) Then csere% = 1
        Else
          If xval(Mid(sortkulcs(qeltol& + qii&), vkp, vmh)) < xval(Mid(sortkulcs(qeltol& + qii& + 1), vkp, vmh)) Then csere% = 1
        End If
      End If
      If csere% = 1 Then
        ok% = 0
        mumu$ = sortkulcs(qeltol& + qii&)
        sortkulcs(qeltol& + qii&) = sortkulcs(qeltol& + qii& + 1)
        sortkulcs(qeltol& + qii& + 1) = mumu$
        mumu1& = sortcimek&(qeltol& + qii&)
        sortcimek(qeltol& + qii&) = sortcimek(qeltol& + qii& + 1)
        sortcimek(qeltol& + qii& + 1) = mumu1&
      End If
    Next
    If ok% = 0 Then
      For qii& = qeli% - 1 To 1 Step -1
        csere% = 0
        If Left(irany$, 1) = "N" Then
          If yqm$ = "K" Then
            If Mid$(sortkulcs(qeltol& + qii&), vkp, vmh) > Mid(sortkulcs(qeltol& + qii& + 1), vkp, vmh) Then csere% = 1
          Else
            If xval(Mid$(sortkulcs(qeltol& + qii&), vkp, vmh)) > xval(Mid(sortkulcs(qeltol& + qii& + 1), vkp, vmh)) Then csere% = 1
          End If
        Else
          If yqm$ = "K" Then
            If Mid(sortkulcs(qeltol& + qii&), vkp, vmh) < Mid(sortkulcs(qeltol& + qii& + 1), vkp, vmh) Then csere% = 1
          Else
            If xval(Mid(sortkulcs(qeltol& + qii&), vkp, vmh)) < xval(Mid(sortkulcs(qeltol& + qii& + 1), vkp, vmh)) Then csere% = 1
          End If
        End If
        If csere% = 1 Then
          ok% = 0
          mumu$ = sortkulcs(qeltol& + qii&)
          sortkulcs(qeltol& + qii&) = sortkulcs(qeltol& + qii& + 1)
          sortkulcs(qeltol& + qii& + 1) = mumu$
          mumu1& = sortcimek&(qeltol& + qii&)
          sortcimek(qeltol& + qii&) = sortcimek(qeltol& + qii& + 1)
          sortcimek(qeltol& + qii& + 1) = mumu1&
        End If
      Next
    End If
  Loop
Return
End Sub

Public Sub qsort(sortkulcs$(), sortcimek&(), elemszam&, irany$)
  '--- címvektor rendezése
  '--- irany=N-növekvõ C-csökkenõ
  Dim qpoz%(1500), qosort$(150000), qscim&(150000)
  If irany$ = "N" Then qhossz& = 500 Else ghossz% = 100
  qhossz& = 100
  If elemszam& < 2 Then Exit Sub
  elemhossz% = Len(sortkulcs(1))
  minima$ = String(elemhossz%, Chr$(255))
  maxima$ = String(elemhossz%, Chr$(0))
  qsorok& = Int(elemszam& / qhossz&): qmar% = elemszam& Mod qhossz&
  If qmar% > 0 Then qsorok& = qsorok& + 1: utel% = qmar% Else utel% = qhossz&
  '--- rendezések
  For qi& = 1 To qsorok&
    If qi& = qsorok& Then qeli% = utel% Else qeli% = qhossz&
    GoSub q1sort
  Next
  '--- merge
  If qsorok& > 1 Then
    For qi& = 1 To qsorok&: qpoz(qi&) = 1: Next
    Do
      If irany$ = "N" Then
        '--- növekvõ
        qmin$ = minima$
        For qii& = 1 To qsorok&
          If qii& = qsorok& Then qeli% = utel% Else qeli% = qhossz&
          If qpoz(qii&) <= qeli% Then
            qhas$ = sortkulcs((qii& - 1) * qhossz& + qpoz(qii&))
            If qhas$ < qmin$ Then
              qminsor& = qii&: qmin$ = qhas$
            End If
          End If
        Next
      Else
        '--- csökkenõ
        qmax$ = maxima$
        For qii& = 1 To qsorok&
          If qii& = qsorok& Then qeli% = utel% Else qeli% = qhossz&
          If qpoz(qii&) <= qeli% Then
            qhas$ = sortkulcs((qii& - 1) * qhossz& + qpoz(qii&))
            If qhas$ > qmax$ Then
              qminsor& = qii&: qmax$ = qhas$
            End If
          End If
        Next
      End If
      qdb& = qdb& + 1
      qosort(qdb&) = sortkulcs((qminsor& - 1) * qhossz& + qpoz(qminsor&))
      qscim&(qdb&) = sortcimek&((qminsor& - 1) * qhossz& + qpoz(qminsor&))
      qpoz(qminsor&) = qpoz(qminsor&) + 1
    Loop While qdb& < elemszam&
    For qqi& = 1 To elemszam&
      sortkulcs(qqi&) = qosort(qqi&)
      sortcimek(qqi&) = qscim(qqi&)
    Next
  End If
Exit Sub
q1sort:
  '--- qeli rendezése
  ok% = 0
  qeltol& = (qi& - 1) * qhossz&
  Do While ok% = 0
    ok% = 1
    For qii& = 1 To qeli% - 1
      csere% = 0
      If irany$ = "N" Then
        If sortkulcs(qeltol& + qii&) > sortkulcs(qeltol& + qii& + 1) Then csere% = 1
      Else
        If sortkulcs(qeltol& + qii&) < sortkulcs(qeltol& + qii& + 1) Then csere% = 1
      End If
      If csere% = 1 Then
        ok% = 0
        mumu$ = sortkulcs(qeltol& + qii&)
        sortkulcs(qeltol& + qii&) = sortkulcs(qeltol& + qii& + 1)
        sortkulcs(qeltol& + qii& + 1) = mumu$
        mumu1& = sortcimek&(qeltol& + qii&)
        sortcimek(qeltol& + qii&) = sortcimek(qeltol& + qii& + 1)
        sortcimek(qeltol& + qii& + 1) = mumu1&
      End If
    Next
    If ok% = 0 Then
      For qii& = qeli% - 1 To 1 Step -1
        csere% = 0
        If irany$ = "N" Then
          If sortkulcs(qeltol& + qii&) > sortkulcs(qeltol& + qii& + 1) Then csere% = 1
        Else
          If sortkulcs(qeltol& + qii&) < sortkulcs(qeltol& + qii& + 1) Then csere% = 1
        End If
        If csere% = 1 Then
          ok% = 0
          mumu$ = sortkulcs(qeltol& + qii&)
          sortkulcs(qeltol& + qii&) = sortkulcs(qeltol& + qii& + 1)
          sortkulcs(qeltol& + qii& + 1) = mumu$
          mumu1& = sortcimek&(qeltol& + qii&)
          sortcimek(qeltol& + qii&) = sortcimek(qeltol& + qii& + 1)
          sortcimek(qeltol& + qii& + 1) = mumu1&
        End If
      Next
    End If
  Loop
Return
End Sub

Public Function postacim$(rec$, poz%)
  '--- postacím egyesítése
  'pcii$ = Trim(Mid$(rec$, poz%, 30))
  'pcii$ = pcii$ + ", " + Trim(Mid$(rec$, poz% + 8, 30))
  pcii$ = Trim(Mid$(rec$, poz%, 8))
  If Not Trim(Mid$(rec$, poz% + 8, 30)) = "" Then
    pcii$ = pcii$ + " " + Trim(Mid$(rec$, poz% + 8, 30))
  End If
  If Not Trim(Mid$(rec$, poz% + 437, 10)) = "" Then
     szep = ","
     If Right(RTrim$(Mid$(rec$, poz% + 437, 10)), 1) = "," Then
          szep = ""
     End If

    pcii$ = pcii$ + " " + Trim(Mid$(rec$, poz% + 437, 10)) + szep
  Else
     szep = ","
     If Right(RTrim$(pcii$), 1) = "," Then
          szep = ""
     End If
    pcii$ = pcii$ + szep
  End If
  If Not Trim(Mid$(rec$, poz% + 38, 30)) = "" Then
    pcii$ = pcii$ + " " + Trim(Mid$(rec$, poz% + 38, 30))
  End If
  If Not Trim(Mid$(rec$, poz% + 447, 10)) = "" Then
     pcii$ = pcii$ + " " + Trim(Mid$(rec$, poz% + 447, 10))
  End If
  If Not Trim(Mid$(rec$, poz% + 68, 10)) = "" Then
     szep = "."
     If Right(RTrim$(Mid$(rec$, poz% + 68, 10)), 1) = "." Then
          szep = ""
     End If
     pcii$ = pcii$ + " " + Trim(Mid$(rec$, poz% + 68, 10)) + szep
  End If
  If Not Trim(Mid$(rec$, poz% + 457, 10)) = "" Then
     szep = ".ép"
     If Right(RTrim$(Mid$(rec$, poz% + 457, 10)), 3) = ".ép" Then
          szep = ""
     End If
     pcii$ = pcii$ + " " + Trim(Mid$(rec$, poz% + 457, 10)) + szep
  End If
  If Not Trim(Mid$(rec$, poz% + 467, 10)) = "" Then
     szep = ".lh"
     If Right(RTrim$(Mid$(rec$, poz% + 467, 10)), 3) = ".lh" Then
          szep = ""
     End If
     pcii$ = pcii$ + " " + Trim(Mid$(rec$, poz% + 467, 10)) + szep
  End If
  If Not Trim(Mid$(rec$, poz% + 477, 10)) = "" Then
     szep = ".em"
     If Right(RTrim$(Mid$(rec$, poz% + 477, 10)), 3) = ".em" Or UCase(RTrim$(Mid$(rec$, poz% + 477, 10))) = "FSZ" Or UCase(RTrim$(Mid$(rec$, poz% + 477, 10))) = "FSZ." Or UCase(RTrim$(Mid$(rec$, poz% + 477, 10))) = "FÖLDSZINT" Then
          szep = ""
     End If
     pcii$ = pcii$ + " " + Trim(Mid$(rec$, poz% + 477, 10)) + szep
  End If
  
  If Not Trim(Mid$(rec$, poz% + 487, 10)) = "" Then
     szep = ".ajto"
     If Right(RTrim$(Mid$(rec$, poz% + 487, 10)), 5) = ".ajtó" Then
          szep = ""
     End If
     If Right(RTrim$(Mid$(rec$, poz% + 487, 10)), 1) = "." Then
          szep = "ajto"
     End If
     
     pcii$ = pcii$ + " " + Trim(Mid$(rec$, poz% + 487, 10)) + szep
  End If
  postacim = pcii$
End Function
Public Function arfolyambe#(bank$, dat$, devnem$)
  '--- árfolyam megállapítása
  If bank$ = "" Then bank$ = Mid$(irec$, 470, 8)
  dkod$ = bank$ + devnem$
  devrec$ = dbxkey("PDEV", dkod$)
  If devrec$ = "" Then arfolyambe = 0: Exit Function
  dkod$ = bank$ + devnem$ + dat$
  arfrec$ = dbxkey("PDRF", dkod$)
  If arfrec$ = "" Then arfolyambe = 0: Exit Function
  egyseg@ = xval(Mid$(arfrec$, 18, 6)): If egyseg@ = 0 Then egyseg@ = 1
  kod$ = Mid$(irec$, 479, 1)
  Select Case kod$
    Case "V"
      arf@ = xval(Mid$(arfrec$, 24, 10))
    Case "K"
      arf@ = xval(Mid$(arfrec$, 34, 10))
    Case "E"
      arf@ = xval(Mid$(arfrec$, 44, 10))
    Case "S"
      devk@ = xval(Mid$(devrec$, 12, 14))
      fint@ = xval(Mid$(devrec$, 26, 14))
      If devk@ <> 0 Then
        arf@ = fint@ / devk@
      Else
        arfolyambe = 0: Exit Function
      End If
    Case Else
  End Select
  arfolyambe = arf@
End Function
Public Sub cimvektor(objazon, indexdarabszam&)
  ob% = obsorszama(objazon)
  dbxn$ = dbxneve(objazon)
  w1% = OBJTAB(ob%).obi(1)
  indn$ = RTrim$(INDTAB(w1%).indnev)
  w2% = INDTAB(w1%).adatsorsz
  kp% = ADATAB(w2%).adatkp
  ho% = ADATAB(w2%).adatho
  ih& = ho% + 5
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  indexmeret& = LOF(ndfi)
  indexdarabszam& = Int(indexmeret& / ih&)
  Close ndfi
  If indexdarabszam& > 0 Then
    cimvektdb& = indexdarabszam&
    ReDim cimvekt(indexdarabszam&)
    cmvfi = FreeFile
    Open auditorutvonal$ + objazon + ".cmv" For Input Shared As #cmvfi
    i1i& = 0
    Do
      i1i& = i1i& + 1
      Line Input #cmvfi, cicike$
      cimvekt(i1i&) = Val(cicike$)
      If i1i& = indexdarabszam& Then Exit Do
    Loop While Not EOF(1)
    If i1i& < indexdarabszam& Then indexdarabszam& = i1i&
    Close cmvfi
  End If
End Sub


Public Function xarfolyam#(sdev$, bdev$)
  '--- keresztárfolyam bekérése
  '--- sdev$ számlázási devizanem
  '--- bdev$ banki devizanem (kiegyenlítés)
  Xarfolym.Text1 = ""
  Xarfolym.Label4.Caption = bdev$
  Xarfolym.Label6.Caption = sdev$
  Xarfolym.Show vbModal
  xarfolyam = xval(Xarfolym.Text1.Text)
End Function

Public Sub partmeg(termekkod$)
  pdarab = 0
  fi1 = FreeFile
  Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #fi1
  fi2 = FreeFile
  Open auditorutvonal$ + "auw-pmeg.ndx" For Binary Shared As #fi2
  rcm& = Int(LOF(fi2) / 12)
  If rcm& > 0 Then
    For iim& = 1 To rcm&
      Get #fi2, (iim& - 1) * 12 + 1, rcim&
      mrec$ = Space(250)
      Get #fi1, rcim& + 9, mrec$
      If Mid$(mrec$, 226, 1) <> "S" Then
        If Mid$(mrec$, 8, 15) = pszlpart$ And (termekkod$ = "" Or termekkod$ = Mid$(mrec$, 55, 6)) Then
          pdarab = pdarab + 1
          partnermeg(pdarab) = mrec$
        End If
      End If
    Next
  End If
  Close fi1: Close fi2
End Sub


Public Function xval@(mez$)
  '--- a val függvény javítása
  nx$ = mez$
  poz% = InStr(nx$, ",")
  If poz% > 0 Then Mid$(nx$, poz%, 1) = "."
  xval = Val(nx$)
End Function
Public Sub tombrend(thtomb$(), tombertek@(), thdb%, tertekdb&, ir$, kulcskezd%, kulcshossz%, kulcstip$)
  '--- tomb rendezése
  '--- ir$="N" növekvõ "C" csökkenõ
  '--- kulcstip$="K" karakteres "N" numerikus
  If thdb% < 2 Then Exit Sub
  Do
    ok% = 1
    For i88% = 1 To thdb% - 1
      mez$ = Mid$(thtomb$(i88%), kulcskezd%, kulcshossz%)
      mez1$ = Mid$(thtomb$(i88% + 1), kulcskezd%, kulcshossz%)
      cs% = 1
      If kulcstip$ = "N" Then
        If ir$ = "N" Then
          If xval(mez$) > xval(mez1$) Then cs% = 0
        Else
          If xval(mez$) < xval(mez1$) Then cs% = 0
        End If
      Else
        If ir$ = "N" Then
          If mez$ > mez1$ Then cs% = 0
        Else
          If mez$ < mez1$ Then cs% = 0
        End If
      End If
      If cs% = 0 Then
        ok% = 0
        mun$ = thtomb$(i88%)
        thtomb(i88%) = thtomb$(i88% + 1)
        thtomb(i88% + 1) = mun$
        If tertekdb& > 0 Then
          mune@ = tombertek@(i88%)
          tombertek@(i88%) = tombertek@(i88% + 1)
          tombertek@(i88% + 1) = mune@
        End If
      End If
    Next
    For i88% = thdb% - 1 To 1 Step -1
      mez$ = Mid$(thtomb$(i88%), kulcskezd%, kulcshossz%)
      mez1$ = Mid$(thtomb$(i88% + 1), kulcskezd%, kulcshossz%)
      cs% = 1
      If kulcstip$ = "N" Then
        If ir$ = "N" Then
          If xval(mez$) > xval(mez1$) Then cs% = 0
        Else
          If xval(mez$) < xval(mez1$) Then cs% = 0
        End If
      Else
        If ir$ = "N" Then
          If mez$ > mez1$ Then cs% = 0
        Else
          If mez$ < mez1$ Then cs% = 0
        End If
      End If
      If cs% = 0 Then
        ok% = 0
        mun$ = thtomb$(i88%)
        thtomb(i88%) = thtomb$(i88% + 1)
        thtomb(i88% + 1) = mun$
        If tertekdb& > 0 Then
          mune@ = tombertek@(i88%)
          tombertek@(i88%) = tombertek@(i88% + 1)
          tombertek@(i88% + 1) = mune@
        End If
      End If
    Next
  Loop While ok% = 0
End Sub

Public Function nagykezdo$(x$)
  '--- kezdõbetû naggyá alakítása
  xx$ = x$: Mid$(xx$, 1, 1) = UCase(Mid$(x$, 1, 1)): nagykezdo = xx$
End Function

Public Sub objmasol(dbxne$, objn$, indexnev$, rekhossz%, inhossz%, honnan$, hova$)
  '--- egy objektum átírás másik adatbázis végére
  fi1 = FreeFile
  Open honnan$ + dbxne$ + ".dbx" For Binary Shared As #fi1
  fi2 = FreeFile
  Open honnan$ + indexnev$ + ".ndx" For Binary Shared As #fi2
  irc& = Int(LOF(fi2) / (inhossz% + 5))
  fi3 = FreeFile
  Open hova$ + dbxne$ + ".dbx" For Binary As #fi3
  ofm& = LOF(fi3)
  opoz& = ofm& + 1
  If irc& > 0 Then
    For i91& = 1 To irc&
      Get #fi2, (i91& - 1) * (inhossz% + 5) + 1, rcim&
      r$ = Space$(rekhossz%)
      Get #fil, rcim& + 9, r$
      Put #fi3, opoz&, objn$ + ";"
      opoz& = opoz& + 5
      Put #fi3, opoz&, 0&
      opoz& = opoz& + 4
      Put #fi3, opoz&, r$
      opoz& = opoz& + rekhossz%
    Next
  End If
  Close fi1: Close fi2: Close fi3
End Sub

Public Sub filmasol(mit$, honnan$, hova$)
  '--- fájlok másolása
  finnev$ = Dir(honnan$ + mit$)
  If finnev$ <> "" Then
    If Left$(finnev$, 1) <> "." Then
      Do
        FileCopy honnan$ + finnev$, hova$ + finnev$
        finnev$ = Dir
      Loop While finnev$ <> ""
    End If
  End If
End Sub
  
Public Function numszam@(szov$)
  '--- formázott szám
  a$ = Trim(szov$)
  p1% = InStr(a$, ",")
  If p1% > 0 Then Mid$(a$, p1%, 1) = "."
  p2% = InStr(a$, Chr$(160))
  Do While p2% > 0
    a$ = Left$(a$, p2% - 1) + Mid$(a$, p2% + 1)
    p2% = InStr(a$, Chr$(160))
  Loop
  numszam = xval(a$)
End Function

Public Sub tordel(szoveg$, hossz%, hatarolo$, param$(), paramdb%)
  For i31% = 1 To paramdb%: param$(i31%) = "": Next
  paramdb% = 0
  a$ = Trim(szoveg$)
  Do While Len(a$) > hossz%
    dsh% = hossz%
    wx1$ = Left$(a$, hossz%)
    For i42% = hossz% To 1 Step -1
      If Mid$(wx1$, i42%, 1) = hatarolo$ Then dsh% = i42%: Exit For
    Next
    wx1$ = Left(a$, dsh%): paramdb% = paramdb% + 1: param$(paramdb%) = wx1$
    a$ = Trim(Mid$(a$, dsh% + 1))
  Loop
  paramdb% = paramdb% + 1
  param$(paramdb%) = a$
End Sub

Public Function scal&(idb&, inszam&)
  ixc& = (idb& / 32000) + 1
  ixscal& = inszam& / ixc&
  If ixscal& = 0 Then scal& = 1 Else scal& = ixscal&
End Function

Public Sub utanir(inpfil$, outfil$, rekhossz%)
  '--- inpfil utánírás outfil végére
  form1.ProgressBar3.Visible = True
  form1.Label6.Visible = True
  ifil = FreeFile
  Open listautvonal$ + terminal$ + task$ + inpfil$ For Binary As #ifil
  ofil = FreeFile
  Open listautvonal$ + terminal$ + task$ + outfil$ For Append As #ofil
  rrc& = Int(LOF(ifil) / rekhossz%)
  If rrc& > 0 Then
    form1.ProgressBar3.Max = 100
    For i32& = 1 To rrc&
      form1.ProgressBar3.Value = pscale(i32&, rrc&)
      Seek #ifil, (i32& - 1) * rekhossz% + 1
      ir$ = Space(rekhossz%): Get #ifil, , ir$
      Print #ofil, ir$;
    Next
  End If
  Close ifil: Close ofil
End Sub

Public Function napkul(dat1$, dat2$)
  '--- intervallum naptári napjainak a száma
  '--- dat1-tõl dat2-ig eltelt napok
  ye1% = Val(Mid$(dat1$, 1, 2)): mo1% = Val(Mid$(dat1$, 3, 2)): da1% = Val(Mid$(dat1$, 5, 2))
  ye2% = Val(Mid$(dat2$, 1, 2)): mo2% = Val(Mid$(dat2$, 3, 2)): da2% = Val(Mid$(dat2$, 5, 2))
  If ye1% > 39 Then ye1% = 1900 + ye1% Else ye1% = 2000 + ye1%
  If ye2% > 39 Then ye2% = 1900 + ye2% Else ye2% = 2000 + ye2%
  napkul = DateSerial(ye2%, mo2%, da2%) - DateSerial(ye1%, mo1%, da1%)
End Function

Function muvjel%(po%, kif$)
  Dim a%(5)
  mje% = 0
  a%(1) = InStr(po%, kif$, "+")
  a%(2) = InStr(po%, kif$, "-"): If a%(2) = po% Then a%(2) = InStr(po% + 1, kif$, "-")
  a%(3) = InStr(po%, kif$, "*")
  a%(4) = InStr(po%, kif$, "/")
  a%(5) = InStr(po%, kif$, "^")
  For i% = 1 To 5
    If a%(i%) <> 0 Then
      If a%(i%) < mje% Or mje% = 0 Then mje% = a%(i%)
    End If
  Next
  muvjel% = mje%
End Function

Public Sub kiegeszit(inpfil$, outfil$, rekhossz%, objne$, kulcskezd%, kulcshossz%)
  '--- forgalmi rekord kiegészítése törzsadatokkal
  '--- epoztomb%(i,0)=kezdet a rekordban
  '--- epoztomb%(i,1)=kezdet a törzsben
  '--- epoztomb%(i,2)=hossz
  form1.ProgressBar3.Visible = True
  form1.Label6.Visible = True
  epozdb% = 0
  For i33% = 1 To 14
    If epoztomb%(i33%, 2) <> 0 Then epozdb% = epozdb% + 1
  Next
  ifil = FreeFile
  Open listautvonal$ + terminal$ + task$ + inpfil$ For Binary As #ifil
  rrc& = Int(LOF(ifil) / rekhossz%)
  If rrc& > 0 Then
    ofil = FreeFile
    Open listautvonal$ + terminal$ + task$ + outfil$ For Output As #ofil
    If rrc& > 0 Then
      form1.ProgressBar3.Max = 100
      For i32& = 1 To rrc&
        DoEvents
        form1.ProgressBar3.Value = pscale(i32&, rrc&)
        Seek #ifil, (i32& - 1) * rekhossz% + 1
        hrec$ = Space(rekhossz%): Get #ifil, , hrec$
        kulcs$ = Mid$(hrec$, kulcskezd%, kulcshossz%)
        torzsrec$ = dbxkey(objne$, kulcs$)
        For i33% = 1 To epozdb%
          Mid$(hrec$, epoztomb%(i33%, 0), epoztomb%(i33%, 2)) = Mid$(torzsrec$, epoztomb%(i33%, 1), epoztomb%(i33%, 2))
        Next
        If mennyisegelojel$ = "-" Or mennyisegelojel$ = "+" Then
          mexxi@ = xval(Mid$(hrec$, mennyisegpozicio, mennyiseghossz))
          If mennyisegelojel$ = "-" And mexxi@ < 0 Then Print #ofil, hrec$;
          If mennyisegelojel$ = "+" And mexxi@ > 0 Then Print #ofil, hrec$;
        Else
          Print #ofil, hrec$;
        End If
      Next
    End If
    'If mennyisegelojel$ = "-" Or mennyisegelojel$ = "+" Then
    '  mexxi@ = xval(Mid$(hrec$, mennyisegpozicio, mennyiseghossz))
    '  If mennyisegelojel$ = "-" And mexxi@ < 0 Then Print #ofil, hrec$;
    '  If mennyisegelojel$ = "+" And mexxi@ > 0 Then Print #ofil, hrec$;
    'Else
    '  Print #ofil, hrec$;
    'End If
    Close ofil
  End If
  Close ifil
End Sub


Public Sub osszevon(inpfil$, outfil$, rekhossz%, rendezokod%, renhiba%, inullpoz%, inullhossz%, onullpoz%, onullhossz%, mimaxhonnan%, mimaxhossz%, minhova%, maxhova%)
  '--- binaris fájl összevonása adott kódmezõkre
  '--- összetett kulcs megállapítása
  '--- rendezokod=0 nem kell rendezni, =1 rendezni kell
  renhiba% = 0
  form1.ProgressBar3.Visible = True
  form1.Label6.Visible = True
  kulcsdb% = 0
  For i33% = 1 To 14
    If kulcstomb%(i33%, 1) <> 0 And kulcstomb%(i33%, 2) <> 0 Then kulcsdb% = kulcsdb% + 1
  Next
  epozdb% = 0
  For i33% = 1 To 14
    If epoztomb%(i33%, 1) <> 0 And epoztomb%(i33%, 2) <> 0 Then epozdb% = epozdb% + 1
  Next
  If rendezokod% = 1 Then
    '--- input file rendezése
    rsor$ = "FIE="
    For i33% = 1 To kulcsdb%
      rsor$ = rsor$ + Trim(kulcstomb%(i33%, 1)) + "," + Trim(kulcstomb%(i33%, 2))
      If i33% <> kulcsdb% Then rsor$ = rsor$ + ","
    Next
    rsor$ = rsor$ + "/INP=" + listautvonal$ + terminal$ + task$ + inpfil$
    rsor$ = rsor$ + "/OUT=" + listautvonal$ + terminal$ + task$ + "SWF.TMP/RLE=" + Trim(Str(rekhossz%)) + "/MOD=A"
    Call rendez(rsor$)
    If rendezohiba% = 1 Then renhiba% = 1: Exit Sub
    form1.Refresh
    ifil = FreeFile
    Open listautvonal$ + terminal$ + task$ + "SWF.TMP" For Binary As #ifil
  Else
    ifil = FreeFile
    Open listautvonal$ + terminal$ + task$ + inpfil$ For Binary As #ifil
  End If
  rrc& = Int(LOF(ifil) / rekhossz%)
  If rrc& > 0 Then
    ofil = FreeFile
    Open listautvonal$ + terminal$ + task$ + outfil$ For Output As #ofil
    Seek #ifil, 1
    krec$ = Space(rekhossz%): Get #ifil, , krec$
    If rrc& > 1 Then
      form1.ProgressBar3.Max = 100
      For i32& = 2 To rrc&
        DoEvents
        form1.ProgressBar3.Value = pscale(i32&, rrc&)
        Seek #ifil, (i32& - 1) * rekhossz% + 1
        hrec$ = Space(rekhossz%): Get #ifil, , hrec$
        If inullpoz% <> 0 Then
          If xval(Mid$(hrec$, inullpoz%, inullhossz%)) <> 0 Then ikell% = 1 Else ikell% = 0
        Else
          ikell% = 1
        End If
        If ikell% = 1 Then
          ill% = 1
          For i33% = 1 To kulcsdb%
            If Mid$(krec$, kulcstomb%(i33%, 1), kulcstomb%(i33%, 2)) <> Mid$(hrec$, kulcstomb%(i33%, 1), kulcstomb%(i33%, 2)) Then ill% = 0: Exit For
          Next
          If ill% = 0 Then
            If onullpoz% <> 0 Then
              If xval(Mid$(krec$, onullpoz%, onullhossz%)) <> 0 Then
                Print #ofil, krec$;
              End If
            Else
              Print #ofil, krec$;
            End If
            krec$ = hrec$
          Else
            For i33% = 1 To epozdb%
              szam@ = xval(Mid$(hrec$, epoztomb%(i33%, 1), epoztomb%(i33%, 2)))
              Call hozzad(krec$, epoztomb%(i33%, 1), epoztomb%(i33%, 2), szam@, 2)
            Next
            If mimaxhonnan% <> 0 Then
              minimx@ = xval(Mid$(krec$, minhova%, mimaxhossz%))
              maximx@ = xval(Mid$(krec$, maxhova%, mimaxhossz%))
              mimaxert@ = xval(Mid$(hrec$, mimaxhonnan%, mimaxhossz%))
              If mimaxert@ < minimx@ Then
                Mid$(krec$, minhova%, mimaxhossz%) = Mid$(hrec$, mimaxhonnan%, mimaxhossz%)
              End If
              If mimaxert@ > maximx@ Then
                Mid$(krec$, maxhova%, mimaxhossz%) = Mid$(hrec$, mimaxhonnan%, mimaxhossz%)
              End If
            End If
          End If
        End If
      Next
    End If
    If onullpoz% <> 0 Then
      If xval(Mid$(krec$, onullpoz%, onullhossz%)) <> 0 Then Print #ofil, krec$;
    Else
      Print #ofil, krec$;
    End If
    Close ofil
  End If
  Close ifil
End Sub

Public Function szamit@(rec$, kif$)
 Dim d$(3, 2)
 d$(1, 1) = "^": d$(1, 2) = "^"
 d$(2, 1) = "*": d$(2, 2) = "/"
 d$(3, 1) = "+": d$(3, 2) = "-"
 a$ = kif$
 If muvjel%(1, a$) = 0 Then
    If Left$(a$, 1) = "%" Then
      mm% = Val(Mid$(a$, 2))
      midi% = ux%(mm%)
      sza@ = xval(Mid$(rec$, mp(midi%).rp, mp(midi%).mh))
      If Mid$(rec$, mp(midi%).rp, mp(midi%).mh) = "-" Then sza@ = -1 Else If mp(midi%).mh = 1 Then sza@ = 1
      szamit@ = sza@
    Else
      szamit@ = xval(a$)
    End If
    Exit Function
 End If
 For iq% = 1 To 3
   Do
     mtw% = 1
     If iq% = 3 Then
       p1% = InStr(a$, "+"): p0% = InStr(a$, "-"): If p0% > 1 And (p0% < p1% Or p1% = 0) Then p1% = p0%: mtw% = 2
       If p1% = 0 And p0% = 1 Then p0% = InStr(2, a$, "-"): p1% = p0%: mtw% = 2
     Else
       p1% = InStr(a$, d$(iq%, 1)): p0% = InStr(a$, d$(iq%, 2)): If p0% <> 0 And p0% < p1% Or p1% = 0 Then p1% = p0%: mtw% = 2
     End If
     If p1% > 0 Then
       '--- van hatvanyozas
       p2% = muvjel%(p1% + 1, a$)
       If p2% > 0 Then
         x$ = Mid$(a$, p1% + 1, p2% - p1% - 1)
         If Left$(x$, 1) = "%" Then
           mm% = Val(Mid$(x$, 2))
           midi% = ux%(mm%)
           j@ = xval(Mid$(rec$, mp(midi%).rp, mp(midi%).mh))
           If Mid$(rec$, mp(midi%).rp, mp(midi%).mh) = "-" Then j@ = -1 Else If mp(midi%).mh = 1 Then j@ = 1
         Else
           j@ = xval(x$)
         End If
         jobb$ = Mid$(a$, p2%)
       Else
         x$ = Mid$(a$, p1% + 1)
         If Left$(x$, 1) = "%" Then
           mm% = Val(Mid$(x$, 2))
           midi% = ux%(mm%)
           j@ = xval(Mid$(rec$, mp(midi%).rp, mp(midi%).mh))
           If Mid$(rec$, mp(midi%).rp, mp(midi%).mh) = "-" Then j@ = -1 Else If mp(midi%).mh = 1 Then j@ = 1
         Else
           j@ = xval(x$)
         End If
         jobb$ = ""
       End If
       p0% = 0
       Do
         p% = muvjel%(p0% + 1, a$)
         If p% < p1% Then p0% = p%
       Loop While p% < p1%
       If p0% > 0 Then
         x$ = Mid$(a$, p0% + 1, p1% - p0% - 1)
         If Left$(x$, 1) = "%" Then
           mm% = Val(Mid$(x$, 2))
           midi% = ux%(mm%)
           b@ = xval(Mid$(rec$, mp(midi%).rp, mp(midi%).mh))
           If Mid$(rec$, mp(midi%).rp, mp(midi%).mh) = "-" Then b@ = -1 Else If mp(midi%).mh = 1 Then b@ = 1
         Else
           b@ = xval(x$)
         End If
         bal$ = Left$(a$, p0%)
       Else
         x$ = Left$(a$, p1% - 1)
         If Left$(x$, 1) = "%" Then
           mm% = Val(Mid$(x$, 2))
           midi% = ux%(mm%)
           b@ = xval(Mid$(rec$, mp(midi%).rp, mp(midi%).mh))
           If Mid$(rec$, mp(midi%).rp, mp(midi%).mh) = "-" Then b@ = -1 Else If mp(midi%).mh = 1 Then b@ = 1
         Else
           b@ = xval(x$)
         End If
         bal$ = ""
       End If
       Select Case iq%
         Case 1: er@ = b@ ^ j@
         Case 2
           If mtw% = 1 Then
             er@ = b@ * j@
           Else
             If j@ = 0 Then er@ = 0@ Else er@ = b@ / j@
           End If
         Case 3
           If mtw% = 1 Then er@ = b@ + j@ Else er@ = b@ - j@
         Case Else
       End Select
       a$ = bal$ + LTrim$(Str$(er@)) + jobb$
     End If
   Loop While p1% > 0
 Next
 szamit@ = er@
End Function

Public Function kszamit@(rec$, kif$)
  a$ = kif$
  Do
    p1% = InStr(a$, "(")
    If p1% > 0 Then
      par% = 1
      p2% = p1%
      Do
        p3% = InStr(p2% + 1, a$, ")")
        p4% = InStr(p2% + 1, a$, "(")
        If p3% <= p4% Or p4% = 0 Then par% = par% - 1: p2% = p3% Else par% = par% + 1: p2% = p4%
      Loop While par% <> 0
      bal$ = Left$(a$, p1% - 1)
      jobb$ = Mid$(a$, p2% + 1)
      a$ = Mid$(a$, p1% + 1, p2% - p1% - 1)
      aq@ = kszamit@(rec$, a$)
      a$ = bal$ + LTrim$(Str$(aq@)) + jobb$
    End If
  Loop While p1% <> 0
  kszamit@ = szamit@(rec$, a$)
End Function

Public Function jelpoz%(po%, kif$)
  Dim a%(7)
  mje% = 0
  a%(1) = InStr(po%, kif$, "+")
  a%(2) = InStr(po%, kif$, "-")
  a%(3) = InStr(po%, kif$, "*")
  a%(4) = InStr(po%, kif$, "/")
  a%(5) = InStr(po%, kif$, "^")
  a%(6) = InStr(po%, kif$, "(")
  a%(7) = InStr(po%, kif$, ")")
  For i% = 1 To 7
    If a%(i%) <> 0 Then
      If a%(i%) < mje% Or mje% = 0 Then mje% = a%(i%)
    End If
  Next
  jelpoz% = mje%
End Function
Public Function fordit$(kif$)
  po% = 1
  a$ = "(" + kif$ + ")"
  Do
    Do
      p1% = jelpoz%(po% + 1, a$)
      If p1% = po% + 1 Then
        po% = po% + 1
      End If
    Loop While p1% = po%
    If p1% <> 0 Then
      nev$ = UCase$(Mid$(a$, po% + 1, p1% - po% - 1))
      bal$ = Left$(a$, po%)
      jobb$ = Mid$(a$, p1%)
      po1% = p1%
      For i% = 1 To ux%(0)
        n1$ = UCase$(narr$(ux%(i%)))
        If InStr(n1$, nev$) = 1 Then
          a$ = bal$ + "%" + LTrim$(Str$(i%)) + jobb$
          po1% = jelpoz%(po% + 1, a$)
          Exit For
        End If
      Next
      po% = po1%
    End If
  Loop While p1% <> 0
  fordit$ = Mid$(a$, 2, Len(a$) - 2)
End Function

Public Sub listazo(tdef$, listanev$, komment$, listhiba%)
  '--- TBT kiterjesztésû fájl interpretálása
  On Error GoTo hibakez
  listhiba% = 0
  soso% = 100
  szelesseg% = 0
  sortores% = 0
  dat$ = Right$(Date$, 2) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
  fi19 = FreeFile
  If Trim(langutvonal$) = "" Then
    Open programutvonal$ + tdef$ + ".tbx" For Binary Shared As #fi19
    lm& = LOF(fi19)
    Close fi19
    If lm& > 0 Then tdef$ = tdef$ + ".tbx" Else tdef$ = tdef$ + ".tbt"
    '--- tbt fájl olvasása
    Open programutvonal$ + tdef$ For Input Shared As #fi19
  Else
    Open langutvonal$ + tdef$ + ".tbx" For Binary Shared As #fi19
    lm& = LOF(fi19)
    Close fi19
    If lm& > 0 Then tdef$ = tdef$ + ".tbx" Else tdef$ = tdef$ + ".tbt"
    '--- tbt fájl olvasása
    Open langutvonal$ + tdef$ For Input Shared As #fi19
  End If
  '--- címsor olvasása
  Line Input #fi19, cime$
  Line Input #fi19, x$
  '--- FILE param‚ter
  If UCase$(Left$(x$, 5)) <> "FILE=" Then
    Call mess(langmodul(69), 1, 0, langmodul(70), valasz%)
    'MsgBox langmodul(69), 48, langmodul(70)
    Close fi19
    listhiba% = 1
    Exit Sub
  End If
  pz1% = InStr(x$, "/")
  filnev$ = listautvonal$ + terminal$ + task$ + Mid$(x$, 6, pz1% - 6)
  rh& = xval(Mid$(x$, pz1% + 1))
  pz2% = InStr(filnev$, ".")
  Line Input #fi19, x$
  '--- REC paraméterek (input fájl rekordleírása)
  If UCase$(Left$(x$, 4)) <> "REC=" Then
    Call mess(langmodul(71), 1, 0, langmodul(70), valasz%)
    'MsgBox "REC " + langmodul(71), 48, langmodul(70)
    Close fi19
    listhiba% = 1
    Exit Sub
  End If
  ms% = 0
  Do While UCase$(Left$(x$, 4)) = "REC="
    a$ = Mid$(x$, 5)
    ux%(0) = ux%(0) + 1
    ux%(ux%(0)) = ux%(0)
    ms% = ms% + 1
    paramdb% = 3
    Call linpar(a$, param$(), "/", paramdb%)
    If paramdb% < 3 Then
      Call mess(langmodul(72), 1, 0, langmodul(70), valasz%)
      'MsgBox "REC " + langmodul(72), 48, langmodul(70)
      Close fi19
      listhiba% = 1
      Exit Sub
    End If
    mp(ms%).rp = xval(param$(2))
    mp(ms%).mh = xval(param$(3))
    narr$(ms%) = param$(1)
    Line Input #fi19, x$
  Loop
  rho& = rh&
  fil$ = filnev$
  bon$ = ""
  lapszam% = 0
  '--- LIST parameter
  If Left$(x$, 4) = "LIST" Then
    lf$ = listautvonal$ + terminal$ + task$ + Mid$(x$, 6)
    Line Input #fi19, x$
  Else
    Call mess(langmodul(71), 1, 0, langmodul(70), valasz%)
    'MsgBox "LIST " + langmodul(71), 48, langmodul(70)
    Close fi19
    listhiba% = 1
    Exit Sub
  End If
  '--- LINE parameter
  If Left$(x$, 4) = "LINE" Then
    linepar% = Val(Mid$(x$, 6, 1))
    Line Input #fi19, x$
  Else
    linepar% = 1
  End If
  '--- PAGE parameter
  If Left$(x$, 4) = "PAGE" Then
    pagpar% = Val(Mid$(x$, 6, 1))
    Line Input #fi19, x$
  Else
    pagpar% = 0
  End If
  '--- MODE parameter
  modvolt% = 0
  If Left$(x$, 4) = "MODE" Then
    Line Input #fi19, x$
  End If
  '--- NULL parameter
  If Left$(UCase$(x$), 2) = "NU" Then
    nullpar% = 1
    Line Input #fi19, x$
  Else
    nullpar% = 0
  End If
  '--- LF fejlec parameterek
  If Left$(x$, 5) = "TORES" Then
    sortores% = Val(Mid$(x$, 7))
    Line Input #fi19, x$
  End If
  If Left$(x$, 6) = "SZELES" Then
    szelesseg% = Val(Mid$(x$, 8))
    Line Input #fi19, x$
  End If
  fejp% = 0
  Do While UCase$(Left$(x$, 2)) = "LF"
    fejp% = fejp% + 1
    fejt$(fejp%) = Mid$(x$, 5)
    Line Input #fi19, x$
  Loop
  fejt$(1) = Trim$(fejt$(1))
  dinfo$ = datki(maidatum$)
  fejt$(3) = komment$
  '--- LSO sor leiro
  If Mid$(x$, 1, 4) <> "LSO=" Then
    Call mess(langmodul(72), 1, 0, langmodul(70), valasz%)
    'MsgBox "LSO " + langmodul(72), 48, langmodul(70)
    Close fi19
    listhiba% = 1
    Exit Sub
  Else
    mssor$ = "MS" + Mid$(x$, 5)
    If szelesseg% > 0 Then
       mssor$ = Mid$(mssor$, 1, szelesseg%)
    End If
    x$ = UCase$(Mid$(x$, 5))
    srr$ = x$ + "  "
    po3% = 1
    spp% = 0
    Do
      po1% = InStr(po3%, x$, "X")
      po11% = InStr(po3%, x$, "F")
      po2% = InStr(po3%, x$, "#")
      po21% = InStr(po3%, x$, "H")
      If po2% = po21% + 1 Then po2% = po21%: hal% = 1 Else hal% = 0
      If po1% = po11% + 1 And po11% <> 0 Then po1% = po11%: fle% = 1 Else fle% = 0
      po% = po1%: If po% = 0 Or po2% < po% And po2% <> 0 Then po% = po2%
      If po% > 0 Then
        po3% = InStr(po%, x$, " ")
        If po3% > 0 Then maszk$ = Mid$(x$, po%, po3% - po%) Else maszk$ = Mid$(x$, po%)
        spp% = spp% + 1
        sor(spp%).mspoz = po%
        If hal% = 1 Then Mid$(maszk$, 1, 1) = "#": sor(spp%).halm = 1
        If fle% = 1 Then Mid$(maszk$, 1, 1) = "X": sor(spp%).flem = 1
        sor(spp%).fmt = maszk$
      End If
    Loop While po% > 0 And po3% > 0
    Line Input #fi19, x$
  End If
  '--- LK kontrollsorok
  kontp% = 0
  Do While UCase$(Left$(x$, 2)) = "LK"
    kontp% = kontp% + 1
    kosor$(kontp%) = Mid$(x$, 5) + "  "
    Line Input #fi19, x$
  Loop
  '--- DK kontrollmezok
  kontp% = 0
  Do While UCase$(Left$(x$, 2)) = "DK"
    kontp% = kontp% + 1
    mnev$ = RTrim$(Mid$(x$, 5))
    pl11% = InStr(mnev$, ":")
    If pl11% > 0 Then
      mxnev$ = Mid$(mnev$, pl11% + 1)
      mnev$ = Left$(mnev$, pl11% - 1)
    Else
      mxnev$ = mnev$
    End If
    pl1% = InStr(mnev$, ";")
    If pl1% > 0 Then
      kontomb(kontp%).cax = Mid$(mnev$, pl1% + 1, 1)
      mnev$ = Left$(mnev$, pl1% - 1)
    Else
      kontomb(kontp%).cax = "0"
    End If
    pl% = InStr(mnev$, "/")
    If pl% > 0 Then
      kontomb(kontp%).car = Mid$(mnev$, pl% + 1, 1)
      mnev$ = Left$(mnev$, pl% - 1)
    Else
      kontomb(kontp%).car = "0"
    End If
    If InStr(mnev$, "(") > 0 Then
      poq1% = InStr(mnev$, "("): poq2% = InStr(mnev$, ")")
      poq$ = Mid$(mnev$, poq1% + 1, poq2% - poq1% - 1)
      poq3% = InStr(poq$, ",")
      kontomb(kontp%).kopoz = xval(Left$(poq$, poq3% - 1))
      kontomb(kontp%).kohossz = xval(Mid$(poq$, poq3% + 1))
    Else
      If mnev$ = "*" Then kontomb(kontp%).kopoz = 1: kontomb(kontp%).kohossz = 0
      For j% = 1 To ux%(0)
        If InStr(narr$(ux%(j%)), mnev$) > 0 Then
          kontomb(kontp%).kopoz = mp(ux%(j%)).rp
          kontomb(kontp%).kohossz = mp(ux%(j%)).mh
          Exit For
        End If
      Next
    End If
    Line Input #fi19, x$
    If modvolt% = 1 Then
      modne$ = modne$ + mxnev$ + "&"
      modn% = modn% + 1
    End If
  Loop
  If Left$(x$, 1) <> "M" Then
    Call mess(langmodul(71), 1, 0, langmodul(70), valasz%)
    'MsgBox "Mnn " + langmodul(71), 48, langmodul(70)
    Close fi19
    listhiba% = 1
    Exit Sub
  Else
    '--- M rovatok
    Do While UCase$(Left$(x$, 1)) = "M"
      sorszam% = Val(Mid$(x$, 2, 2))
      If sorszam% = 0 Then
        Call mess(langmodul(72), 1, 0, langmodul(70), valasz%)
        'MsgBox "Mnn " + langmodul(72), 48, langmodul(70)
        Close fi19
        listhiba% = 1
        Exit Sub
      Else
        mnev$ = Trim$(Mid$(x$, 5))
        If InStr(mnev$, "/+") > 0 Or InStr(mnev$, "/-") > 0 Or InStr(mnev$, "/*") > 0 Then
          pxz% = InStr(mnev$, "/")
          If Right$(mnev$, 3) = "ABS" Then abe$ = "1" Else abe$ = " "
          elojele$ = Mid$(mnev$, pxz% + 1, 1): mnev$ = Left$(mnev$, pxz% - 1)
        Else
          abe$ = " "
          elojele$ = "*"
        End If
        If InStr(mnev$, ",") > 0 Then
          poq1% = InStr(mnev$, "("): poq2% = InStr(mnev$, ")")
          poq$ = Mid$(mnev$, poq1% + 1, poq2% - poq1% - 1)
          poq3% = InStr(poq$, ",")
          kopoz1% = Val(Left$(poq$, poq3% - 1))
          kohossz1% = Val(Mid$(poq$, poq3% + 1))
          ux%(0) = ux%(0) + 1: ux%(ux%(0)) = ux%(0): mp(ux%(ux%(0))).rp = kopoz1%: mp(ux%(ux%(0))).mh = kohossz1%
    parr$(ux%(ux%(0))) = " ": narr$(ux%(ux%(0))) = "***":
          sor(sorszam%).msorszam = ux%(ux%(0))
          elo$(sorszam%) = elojele$
          abert$(sorszam%) = abe$
        Else
          For j% = 1 To ux%(0)
            If InStr(narr$(ux%(j%)), mnev$) > 0 Then
              sor(sorszam%).msorszam = ux%(j%)
              elo$(sorszam%) = elojele$
              abert$(sorszam%) = abe$
              Exit For
            End If
          Next
        End If
        If sor(sorszam%).msorszam = 0 Then
          kifp% = kifp% + 1
          kif$(kifp%) = fordit$(mnev$)
          sor(sorszam%).ksorsz = kifp%
          elo$(sorszam%) = elojele$
          abert$(sorszam%) = abe$
        End If
        Line Input #fi19, x$
      End If
    Loop
  End If
  Close fi19
  '--- értelmezés vége
  '--- lista keszitese
  fi20 = FreeFile
  Open lf$ For Output As #fi20
  fi1 = FreeFile
  Open fil$ For Binary As #fi1
  rc& = Int(LOF(fi1) / rho&)
  If rc& = 0 Then
    Call mess(langmodul(73), 3, 0, langmodul(74), valasz%)
    'MsgBox langmodul(73), 48, langmodul(74)
    Close fi20: Close fi1
    listhiba% = 1
    Exit Sub
  End If
  form1.Label5.Visible = True
  form1.ProgressBar2.Visible = True
  form1.ProgressBar2.Max = 100
  '--- cimsor és fejlécek kiírása
  irec$ = dbxkey("INST", "INST")
  hed1$ = Trim(Mid$(irec$, 5, 60)) + " " + "(AU2003/" + regszam$ + ")" + " " + dinfo$ + "/" + Trim$(ugyintezo$)
  Print #fi20, mssor$
  Print #fi20, "CM" + fejt$(1)
  Print #fi20, "FL" + hed1$
  For j1% = 2 To fejp%
    If j1% = 3 Then
      dsh% = Len(Trim(fejt$(2)))
      dss$ = Trim(fejt$(3))
      If Len(dss$) < dsh% Then
        Print #fi20, "FL" + dss$
      Else
        dsp% = dsh%
        For j111% = dsh% To 1 Step -1
          If Mid$(dss$, j111%, 1) = "," Then dsp% = j111%: Exit For
        Next
        ds1$ = Left$(dss$, dsp%)
        Print #fi20, "FL" + ds1$
        ds1$ = Trim(Mid$(dss$, dsp% + 1))
        Print #fi20, "FL" + ds1$
      End If
    Else
      Print #fi20, "FL" + fejt$(j1%)
    End If
  Next
  '--- elso rekord betöltése
  idx& = 1
  Seek #fi1, (idx& - 1) * rho& + 1
  rek$ = Space(rho&): Get #fi1, , rek$
  For j% = 1 To kontp% - 1
    konert$(j%) = Mid$(rek$, kontomb(j%).kopoz, kontomb(j%).kohossz)
    hasert$(j%) = konert$(j%)
  Next
  ele& = 0
  tetelso& = 1
  Do
    For i% = kontp% - 1 To 1 Step -1
      If konert$(i%) <> hasert$(i%) Then
        '--- kontrolszakitas
        tetelso& = 1: ele& = 1
        For j% = 1 To i%
          '--- összesen sor feltöltése
          b$ = Space$(Len(srr$))
          Mid$(b$, 1) = kosor$(j%)
          For j1% = 1 To spp%
            pozic% = sor(j1%).mspoz: forma$ = Trim$(sor(j1%).fmt)
            If Mid$(kosor$(j%), pozic%, 1) = "*" Then
              mhx% = Len(forma$)
              Mid$(b$, pozic%, mhx%) = Mid$(but$, pozic%, mhx%)
            Else
              If Mid$(kosor$(j%), pozic%, 1) <> "#" Then forma$ = " "
              mhx% = Len(forma$)
              If Left$(forma$, 1) = "#" Then
                cur@ = gyujto@(j%, j1%)
                c1@ = Abs(cur@)
                mezir$ = Right$(Space$(20) + Format$(c1@, forma$), mhx%)
                If cur@ < 0 Then
                  mezir$ = mezir$ + "-": mhx% = mhx% + 1
                End If
                Mid$(b$, pozic%, mhx%) = mezir$
              End If
            End If
          Next
          '--- sor kiírása
          '--- aláhúzás elõtte
          If modpar% = 0 Or j% >= modpar% Then
            If kontomb(j%).car <> "0" Then
              If szelesseg% <> 0 Then
                Print #fi20, "TS" + String$(szelesseg%, kontomb(j%).car)
              Else
                Print #fi20, "TS" + String$(Len(b$) - 2, kontomb(j%).car)
              End If
            End If
            xa$ = "K" + Trim(Str$(j%))
            If sortores% = 0 Then
              Print #fi20, RTrim(xa$ + b$)
            Else
              If Trim(Left(b$, sortores%)) <> "" Then Print #fi20, RTrim(xa$ + Left(b$, sortores%))
              If Trim(Mid(b$, sortores% + 1)) <> "" Then Print #fi20, RTrim(xa$ + Mid(b$, sortores% + 1))
            End If
            If kontomb(j%).cax <> "0" Then
              If szelesseg% <> 0 Then
                Print #fi20, "TS" + String$(szelesseg%, kontomb(j%).cax)
              Else
                Print #fi20, "TS" + String$(Len(b$) - 2, kontomb(j%).cax)
              End If
            End If
            If j% = pagpar% Then
              zx$ = "UJLAP"
              Print #fi20, zx$
            End If
          End If
          '--- gyujtok feltoltese
          For k% = 1 To spp%
            gyujto@(j% + 1, k%) = gyujto@(j% + 1, k%) + gyujto@(j%, k%)
            gyujto@(j%, k%) = 0
          Next
          '--- kontroll attoltes
          konert$(j%) = hasert$(j%)
        Next
      End If
    Next
    '--- sor osszeallitasa
    irja% = 0
    b$ = srr$
    For j1% = 1 To spp%
      midi% = sor(j1%).msorszam
      If midi% <> 0 Then mezo$ = Mid$(rek$, mp(midi%).rp, mp(midi%).mh)
      pozic% = sor(j1%).mspoz: forma$ = Trim$(sor(j1%).fmt)
      mhx% = Len(forma$)
      If Left$(forma$, 1) = "X" Then
        If sor(j1%).flem = 0 Or tetelso& = 1 Then
          Mid$(b$, pozic%, Len(forma$)) = Left$(mezo$ + Space$(80), mhx%)
        Else
          Mid$(b$, pozic%, Len(forma$)) = Left$(Space$(80), mhx%)
        End If
      Else
        If sor(j1%).ksorsz <> 0 Then
          cur@ = 0
          '-- cur@=kifejezes
          cur@ = kszamit@(rek$, kif$(sor(j1%).ksorsz))
        Else
          cur@ = xval(mezo$)
          If eljp%(midi%) <> 0 Then
            If Mid$(rek$, eljp%(midi%), 1) = elje$(midi%) Then cur@ = -cur@
          End If
        End If
        If elo$(j1%) = "*" Or elo$(j1%) = "-" And cur@ < 0 Or elo$(j1%) = "+" And cur@ >= 0 Then
          If abert$(j1%) = "1" Then cur@ = Abs(cur@)
          gyujto@(1, j1%) = gyujto@(1, j1%) + cur@
          If sor(j1%).halm = 1 Then cur@ = gyujto@(1, j1%)
          c1@ = Abs(cur@)
          mezir$ = Right$(Space$(20) + Format(c1@, forma$), mhx%)
          If cur@ < 0 Then
            mezir$ = mezir$ + "-": mhx% = mhx% + 1
          End If
          Mid$(b$, pozic%, mhx%) = mezir$
        Else
          Mid$(b$, pozic%, mhx%) = Space$(mhx%)
        End If
        If cur@ <> 0 Then irja% = 1
      End If
    Next
    If tetelso& = 1 Then but$ = b$
    tetelso& = 0
    '--- tetelsor kiirasa
    If modpar% = 0 And linepar% > 0 Then
      If nullpar% = 0 Or irja% = 1 Then
        If sortores% = 0 Then
          Print #fi20, "TS" + b$
        Else
          b1b$ = Left(b$, sortores%)
          b2b$ = Mid(b$, sortores% + 1)
          If Trim(b1b$) <> "" Then Print #fi20, "TS" + b1b$
          If Trim(b2b$) <> "" Then Print #fi20, "TS" + b2b$
        End If
        If linepar% = 2 Then Print #fi20, "SPACE"
      End If
    End If
    '--- gyujtes
    idx& = idx& + 1
    If idx& <= rc& Then
      form1.ProgressBar2.Value = pscale(idx&, rc&)
      eo% = 0
      Seek #fi1, (idx& - 1) * rho& + 1
      '--- javítottam rh&
      rek$ = Space(rho&): Get #fi1, , rek$
      '--- torzsek olvasasa
      For j% = 1 To kontp% - 1
        hasert$(j%) = Mid$(rek$, kontomb(j%).kopoz, kontomb(j%).kohossz)
      Next
    Else
      eo% = 1
    End If
    DoEvents
  Loop While eo% = 0
  '--- vegoszesenek
  For j% = 1 To kontp%
    '--- kontrolsor irasa
    If modpar% = 0 Or j% >= modpar% Then
      b$ = Space$(Len(srr$))
      Mid$(b$, 1) = kosor$(j%)
      For j1% = 1 To spp%
        pozic% = sor(j1%).mspoz: forma$ = Trim$(sor(j1%).fmt)
        If Mid$(kosor$(j%), pozic%, 1) = "*" Then
          mhx% = Len(forma$)
          Mid$(b$, pozic%, mhx%) = Mid$(but$, pozic%, mhx%)
        Else
          If Mid$(kosor$(j%), pozic%, 1) <> "#" Then forma$ = " "
          mhx% = Len(forma$)
          If Left$(forma$, 1) = "#" Then
            cur@ = gyujto@(j%, j1%)
            c1@ = Abs(cur@)
            mezir$ = Right$(Space$(20) + Format(c1@, forma$), mhx%)
            If cur@ < 0 Then
              mezir$ = mezir$ + "-": mhx% = mhx% + 1
            End If
            Mid$(b$, pozic%, mhx%) = mezir$
          End If
        End If
      Next
      If kontomb(j%).car <> "0" Then
        If szelesseg% = 0 Then
          Print #fi20, "TS" + String$(Len(b$) - 2, kontomb(j%).car)
        Else
          Print #fi20, "TS" + String$(szelesseg%, kontomb(j%).car)
        End If
      End If
      xa$ = "K" + Trim(Str$(j%))
      If sortores% = 0 Then
        Print #fi20, RTrim(xa$ + b$)
      Else
        Print #fi20, RTrim(xa$ + Left(b$, sortores%))
        Print #fi20, RTrim(xa$ + Mid(b$, sortores% + 1))
      End If
      If kontomb(j%).cax <> "0" Then
        If szelesseg% = 0 Then
          Print #fi20, "TS" + String$(Len(b$) - 2, kontomb(j%).cax)
        Else
          Print #fi20, "TS" + String$(szelesseg%, kontomb(j%).cax)
        End If
      End If
    End If
    If j% = pagpar% Then
      zx$ = "UJLAP"
      Print #fi20, zx$
    End If
    If j% = kontp% Then
    Else
      '--- gyujtes
      For k% = 1 To spp%
        gyujto@(j% + 1, k%) = gyujto@(j% + 1, k%) + gyujto@(j%, k%)
        gyujto@(j%, k%) = 0
      Next
    End If
  Next
  Close fi1: Close fi20
  Exit Sub
hibakez:
  Call mess(langmodul(166), 2, 0, langmodul(165), valasz%)
  Resume Next
  Close fi1: Close fi20
  Exit Sub
End Sub
Public Function komhas(kod$, sorsz%, rec$, rekpoz%)
  '--- rekord értékének hasonlítás kom paraméterhez
  '--- kod relációk
  '---     numerikus =N,<N,<=N,>N,>=N
  '---     string    =S,<S,<=S,>S,>=S
  '--- kod intervallum
  '---     IN-numerikus, IS-string, ID-dátum
  '---
  atu$ = komt(sorsz%).komatr
  hhos% = Val(Mid$(atu$, 2))
  Select Case kod$
    Case "=S"
      hmez$ = UCase$(Trim(komt(sorsz%).komtol))
      hhos% = Len(hmez$)
      If hhos% > 0 Then
        kmez$ = UCase$(Mid$(rec$, rekpoz%, hhos%))
        If kmez$ = hmez$ Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case "<S"
      hmez$ = UCase$(Trim(komt(sorsz%).komtol))
      hhos% = Len(hmez$)
      If hhos% > 0 Then
        kmez$ = UCase$(Mid$(rec$, rekpoz%, hhos%))
        If kmez$ < hmez$ Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case "<=S"
      hmez$ = UCase$(Trim(komt(sorsz%).komtol))
      hhos% = Len(hmez$)
      If hhos% > 0 Then
        kmez$ = UCase$(Mid$(rec$, rekpoz%, hhos%))
        If kmez$ <= hmez$ Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case ">S"
      hmez$ = UCase$(Trim(komt(sorsz%).komtol))
      hhos% = Len(hmez$)
      If hhos% > 0 Then
        kmez$ = UCase$(Mid$(rec$, rekpoz%, hhos%))
        If kmez$ > hmez$ Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case ">=S"
      hmez$ = UCase$(Trim(komt(sorsz%).komtol))
      hhos% = Len(hmez$)
      If hhos% > 0 Then
        kmez$ = UCase$(Mid$(rec$, rekpoz%, hhos%))
        If kmez$ >= hmez$ Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case "=D"
      If Len(Trim$(komt(sorsz%).komtol)) = 6 Then
        hmez$ = UCase$(Trim(komt(sorsz%).komtol))
        kmez$ = UCase$(Mid$(rec$, rekpoz%, 6))
        If dtm(kmez$) = dtm(hmez$) Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case "<D"
      If Len(Trim$(komt(sorsz%).komtol)) = 6 Then
        hmez$ = UCase$(Trim(komt(sorsz%).komtol))
        kmez$ = UCase$(Mid$(rec$, rekpoz%, 6))
        If dtm(kmez$) < dtm(hmez$) Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case "<=D"
      If Len(Trim$(komt(sorsz%).komtol)) = 6 Then
        hmez$ = UCase$(Trim(komt(sorsz%).komtol))
        kmez$ = UCase$(Mid$(rec$, rekpoz%, 6))
        If dtm(kmez$) <= dtm(hmez$) Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case ">D"
      If Len(Trim$(komt(sorsz%).komtol)) = 6 Then
        hmez$ = UCase$(Trim(komt(sorsz%).komtol))
        kmez$ = UCase$(Mid$(rec$, rekpoz%, 6))
        If dtm(kmez$) > dtm(hmez$) Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case ">=D"
      If Len(Trim$(komt(sorsz%).komtol)) = 6 Then
        hmez$ = UCase$(Trim(komt(sorsz%).komtol))
        kmez$ = UCase$(Mid$(rec$, rekpoz%, 6))
        If dtm(kmez$) >= dtm(hmez$) Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case "=N"
      If Trim(komt(sorsz%).komtol) <> "" Then
        kert@ = xval(komt(sorsz%).komtol)
        hert@ = xval(Mid$(rec$, rekpoz%, hhos%))
        If hert@ = kert@ Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case "<N"
      If Trim(komt(sorsz%).komtol) <> "" Then
        kert@ = xval(komt(sorsz%).komtol)
        hert@ = xval(Mid$(rec$, rekpoz%, hhos%))
        If hert@ < kert@ Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case "<=N"
      If Trim(komt(sorsz%).komtol) <> "" Then
        kert@ = xval(komt(sorsz%).komtol)
        hert@ = xval(Mid$(rec$, rekpoz%, hhos%))
        If hert@ <= kert@ Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case ">N"
      If Trim(komt(sorsz%).komtol) <> "" Then
        kert@ = xval(komt(sorsz%).komtol)
        hert@ = xval(Mid$(rec$, rekpoz%, hhos%))
        If hert@ > kert@ Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case ">=N"
      If Trim(komt(sorsz%).komtol) <> "" Then
        kert@ = xval(komt(sorsz%).komtol)
        hert@ = xval(Mid$(rec$, rekpoz%, hhos%))
        If hert@ >= kert@ Then komhas = 1 Else komhas = 0
      Else
        komhas = 1
      End If
    Case "IS"
      hmez$ = UCase$(Trim(komt(sorsz%).komtol))
      hhos% = Len(hmez$)
      If hhos% > 0 Then
        kmez$ = UCase$(Mid$(rec$, rekpoz%, hhos%))
        If kmez$ < hmez$ Then komhas = 0: Exit Function
      End If
      hmez$ = UCase$(Trim(komt(sorsz%).komig))
      hhos% = Len(hmez$)
      If hhos% > 0 Then
        kmez$ = UCase$(Mid$(rec$, rekpoz%, hhos%))
        If kmez$ > hmez$ Then komhas = 0: Exit Function
      End If
      komhas = 1
    Case "ID"
      hmez$ = UCase$(Trim(komt(sorsz%).komtol))
      hhos% = Len(hmez$)
      If hhos% = 6 Then
        kmez$ = UCase$(Mid$(rec$, rekpoz%, hhos%))
        If dtm(kmez$) < dtm(hmez$) Then komhas = 0: Exit Function
      End If
      hmez$ = UCase$(Trim(komt(sorsz%).komig))
      hhos% = Len(hmez$)
      If hhos% = 6 Then
        kmez$ = UCase$(Mid$(rec$, rekpoz%, hhos%))
        If dtm(kmez$) > dtm(hmez$) Then komhas = 0: Exit Function
      End If
      komhas = 1
    Case "IN"
      If Trim(komt(sorsz%).komtol) <> "" Then
        hert@ = xval(komt(sorsz%).komtol)
        kert@ = xval(Mid$(rec$, rekpoz%, hhos%))
        If kert@ < hert@ Then komhas = 0: Exit Function
      End If
      If Trim(komt(sorsz%).komig) <> "" Then
        hert@ = xval(komt(sorsz%).komig)
        kert@ = xval(Mid$(rec$, rekpoz%, hhos%))
        If kert@ > hert@ Then komhas = 0: Exit Function
      End If
      komhas = 1
    Case Else
  End Select
End Function
Public Sub kodtablak(azo$, vakod$)
  Select Case azo$
    Case "FAFAJ": betukod$ = "FORDADO": Kodok.Show vbModal
    Case "FIZMJELL": betukod$ = "PFIZJEL": Kodok.Show vbModal
    Case "AFALAN": betukod$ = "PADOALA": Kodok.Show vbModal
    Case "FTIPUS": betukod$ = "FKSZTIP": Kodok.Show vbModal
    Case "FJELLEG": betukod$ = "FKSZJEL": Kodok.Show vbModal
    Case "FVEGYES": betukod$ = "FKSZVEGY": Kodok.Show vbModal
    Case "FSZVJEL": betukod$ = "FSZVJEL": Kodok.Show vbModal
    Case "FMSZJEL": betukod$ = "FMSZJEL": Kodok.Show vbModal
    Case "BIZTIPUS": betukod$ = "FKBZTIP": Kodok.Show vbModal
    Case "PAJELL": betukod$ = "PARTJEL": Kodok.Show vbModal
    Case "LETILT": betukod$ = "PARTLET": Kodok.Show vbModal
    Case "PAFATIP": betukod$ = "PAFATIP": Kodok.Show vbModal
    Case "PAFAJEL": betukod$ = "PAFAJEL": Kodok.Show vbModal
    Case "PAFAVI": betukod$ = "PAFAVI": Kodok.Show vbModal
    Case "TJELL": betukod$ = "TJELL": Kodok.Show vbModal
    Case "PVSTAT", "PSSTAT": betukod$ = "PVSZSTAT": Kodok.Show vbModal
    Case "PVNYIT", "PSNYIT": betukod$ = "PVSZNYIT": Kodok.Show vbModal
    Case "PELOTIP": betukod$ = "PELOTIP": Kodok.Show vbModal
    Case "PELOMUV": betukod$ = "PELOMUV": Kodok.Show vbModal
    Case "PELONYJ": betukod$ = "PELONYIT": Kodok.Show vbModal
    Case "PTMUV", "PTTBEKI": betukod$ = "PKTEMUV": Kodok.Show vbModal
    Case "PTKOD", "PTFKOD": betukod$ = "PKTEKOD": Kodok.Show vbModal
    Case "BNMUV", "BNFBEKI": betukod$ = "PBNKMUV": Kodok.Show vbModal
    Case "BNKOD", "BNFKOD": betukod$ = "PBNKKOD": Kodok.Show vbModal
    Case "KRKOD": betukod$ = "PKORKOD": Kodok.Show vbModal
    Case "KRAFAK": betukod$ = "PKORAFA": Kodok.Show vbModal
    Case "PSZFAJT": betukod$ = "PSZFAJT": Kodok.Show vbModal
    Case "PTSZTOR", "BNSZTOR", "KRSTOR", "PSZBSTRN": betukod$ = "SZTORNO": Kodok.Show vbModal
    Case "BNELKOD": betukod$ = "BNELKOD": Kodok.Show vbModal
    Case "FDEVIZ", "FKLTSGH", "FSZERV", "FMUNSZ": betukod$ = "KOTELEZO": Kodok.Show vbModal
    Case "KAMATI": betukod$ = "KAMATI": Kodok.Show vbModal
    Case "FELSZI": betukod$ = "FELSZI": Kodok.Show vbModal
    Case "IVKONV", "ISKONV", "ISDKRF", "ISKORF", "ISKTRF": betukod$ = "KONVARF": Kodok.Show vbModal
    Case "TKKEZEL": betukod = "TKKEZEL": Kodok.Show vbModal
    Case "KMOZIR": betukod$ = "KMOZIR": Kodok.Show vbModal
    Case "ESZTARK": betukod$ = "ESZTARK": Kodok.Show vbModal
    Case "ESZLETK": betukod$ = "ESZLETK": Kodok.Show vbModal
    Case "ESZECTIP": betukod$ = "ESZECTIP": Kodok.Show vbModal
    Case "ENRECSJ": betukod$ = "ENRECSJ": Kodok.Show vbModal
    Case "EFRGKOD": betukod$ = "EFRGKOD": Kodok.Show vbModal
    Case "PEGTIP": betukod$ = "PEGTIP": Kodok.Show vbModal
    Case "PEGJELL": betukod$ = "PEGJELL": Kodok.Show vbModal
    Case "PEGRSZ": betukod$ = "PEGRSZ": Kodok.Show vbModal
    Case "PEGKAPC": betukod$ = "PEGKAPC": Kodok.Show vbModal
    Case "KTJELL": betukod$ = "KTRMJELL": Kodok.Show vbModal
    Case "KTAJELL": betukod$ = "KTRMAJEL": Kodok.Show vbModal
    Case "KTKEZEL": betukod$ = "KTRMKESZ": Kodok.Show vbModal
    Case "KTRKAKOD": betukod$ = "KTRMAKOD": Kodok.Show vbModal
    Case "KTKEZME": betukod$ = "KTRMMTA": Kodok.Show vbModal
    Case "KTLETI": betukod$ = "KTRMLETI": Kodok.Show vbModal
    Case "KTJOVK": betukod$ = "KTRMJOVK": Kodok.Show vbModal
    Case "KTFRCH": betukod$ = "KTRMFRCH": Kodok.Show vbModal
    Case "KTFELT": betukod$ = "KTRMKASS": Kodok.Show vbModal
    Case "KCSPBAZ": betukod$ = "KCSPBAZ": Kodok.Show vbModal
    Case "KONFNALA": betukod$ = "KCSPBAZ": Kodok.Show vbModal
    Case "KONFNALD": betukod$ = "KCSPBAZ": Kodok.Show vbModal
    Case "KONFNALF": betukod$ = "KCSPBAZ": Kodok.Show vbModal
    Case "KCSPHAT": betukod$ = "KCSPHAT": Kodok.Show vbModal
    Case "KRJELL": betukod$ = "KRAKJELL": Kodok.Show vbModal
    Case "KMOXIR": betukod$ = "KMOXIR": Kodok.Show vbModal
    Case "KMBIZT": betukod$ = "KMOXBIZT": Kodok.Show vbModal
    Case "KUZLJM": betukod$ = "KUZLJM": Kodok.Show vbModal
    Case "KSZCSZB": betukod$ = "KSZCSZB": Kodok.Show vbModal
    Case "KSZCJENG": betukod$ = "KSZCJENG": Kodok.Show vbModal
    Case "ARMGAKTV": betukod$ = "ARMGAKTV": Kodok.Show vbModal
    Case "ARM2AKTV": betukod$ = "ARMGAKTV": Kodok.Show vbModal
    Case "KMESTAT": betukod$ = "KMESTAT": Kodok.Show vbModal
    Case "KMESZALM": betukod$ = "KMESZALM": Kodok.Show vbModal
    Case "KMESZBI": betukod$ = "KMESZBI": Kodok.Show vbModal
    Case "KMESTJ": betukod$ = "KMESTJ": Kodok.Show vbModal
    Case "KKBZFJT": betukod$ = "KKBZFJT": Kodok.Show vbModal
    Case "KKB3FJT": betukod$ = "KKBZFJT": Kodok.Show vbModal
    Case "KSZFAJT": betukod$ = "KSZFAJT": Kodok.Show vbModal
    Case "KONFKBEL": betukod$ = "KONFKBEL": Kodok.Show vbModal
    Case Else
  End Select
End Sub

Public Function ARfkonv(adat$, arf#, hossz%)
  adta@ = xval(adat$) * arf#
  ARfkonv = Right$(Space$(hossz%) + Format(adta@, "#############0.00"), hossz%)
End Function

Public Function deviza(foro@, dat$, bank$, devnem$, kod$)
  '--- forint összeg konvertálása devizára
  '--- kod$ V-vételi árfolyamon
  '---      K-közép árfolyamon
  '---      E-eladási árfolyamon
  dkod$ = bank$ + devnem$
  devrec$ = dbxkey("PDEV", dkod$)
  If devrec$ = "" Then deviza = -1: Exit Function
  dkod$ = bank$ + devnem$ + dat$
  arfrec$ = dbxkey("PDRF", dkod$)
  If arfrec$ = "" Then deviza = -1: Exit Function
  egyseg@ = xval(Mid$(arfrec$, 18, 6)): If egyseg@ = 0 Then egyseg@ = 1
  Select Case kod$
    Case "V"
      arf@ = xval(Mid$(arfrec$, 24, 10))
      If arf@ <> 0 Then
        If devnem$ = "HUF" And konyveldevnem$ <> "   " Then
          atf@ = (foro@ * egyseg@) / arf@
        Else
          atf@ = (foro@ * egyseg@) * arf@
        End If
        deviza = xval(Format(atf@, "##########0.00"))
      Else
        deviza = -1
      End If
    Case "K"
      arf@ = xval(Mid$(arfrec$, 34, 10))
      If arf@ <> 0 Then
        If devnem$ = "HUF" And konyveldevnem$ <> "   " Then
          atf@ = (foro@ * egyseg@) / arf@
        Else
          atf@ = (foro@ * egyseg@) * arf@
        End If
        deviza = xval(Format(atf@, "##########0.00"))
      Else
        deviza = -1
      End If
    Case "E"
      arf@ = xval(Mid$(arfrec$, 44, 10))
      If arf@ <> 0 Then
        If devnem$ = "HUF" And konyveldevnem$ <> "   " Then
          atf@ = (foro@ * egyseg@) / arf@
        Else
          atf@ = (foro@ * egyseg@) * arf@
        End If
        deviza = xval(Format(atf@, "##########0.00"))
      Else
        deviza = -1
      End If
    Case Else
  End Select
End Function

Public Sub devkesz(devo@, foro@, bank$, devnem$)
  '--- deviza készlet módosítása
  '--- devo és foro elõjelhelyesen
  dkod$ = bank$ + devnem$
  devrec$ = dbxkey("PDEV", dkod$)
  If devrec$ <> "" Then
    '--- deviza rekord létezik
    atf@ = xval(Mid$(devrec$, 12, 14)) + devo@
    Mid$(devrec$, 12, 14) = Right$(Space$(14) + Format(atf@, "##########0.00"), 14)
    atf@ = xval(Mid$(devrec$, 26, 14)) + foro@
    Mid$(devrec$, 26, 14) = Right$(Space$(14) + Format(atf@, "##########0.00"), 14)
    Call dbxki("PDEV", devrec$, ";", "", "", hiba%)
  Else
    If devo@ > 0 Then
      '--- deviza készlet létrejön
      devrec$ = Space$(50)
      Mid$(devrec$, 1, 11) = dkod$
      Mid$(devrec$, 12, 14) = ertszam(Str$(devo@), 14, 2)
      Mid$(devrec$, 26, 14) = ertszam(Str$(foro@), 14, 2)
      hiba% = 0
      Call dbxki("PDEV", devrec$, ";", "U", "", hiba%)
    End If
  End If
End Sub

Public Function forint(devo@, dat$, bank$, devnem$, kod$)
  '--- deviza összeg konvertálása forintra
  '--- kod$ V-vételi árfolyamon
  '---      K-közép árfolyamon
  '---      E-eladási árfolyamon
  '---      S-deviza készletrõl
  dkod$ = bank$ + devnem$
  devrec$ = dbxkey("PDEV", dkod$)
  If devrec$ = "" Then forint = -1: Exit Function
  dkod$ = bank$ + devnem$ + dat$
  arfrec$ = dbxkey("PDRF", dkod$)
  If arfrec$ = "" Then forint = -1: Exit Function
  egyseg@ = xval(Mid$(arfrec$, 18, 6)): If egyseg@ = 0 Then egyseg@ = 1
  Select Case kod$
    Case "V"
      arf@ = xval(Mid$(arfrec$, 24, 10))
      If devnem$ = "HUF" And konyveldevnem <> "   " Then
        If arf@ <> 0 Then atf@ = (devo@ / egyseg@) / arf@ Else atf@ = -1
      Else
        atf@ = (devo@ / egyseg@) * arf@
      End If
      forint = xval(Format(atf@, "##########0.00"))
    Case "K"
      arf@ = xval(Mid$(arfrec$, 34, 10))
      If devnem$ = "HUF" And konyveldevnem <> "   " Then
        If arf@ <> 0 Then atf@ = (devo@ / egyseg@) / arf@ Else atf@ = -1
      Else
        atf@ = (devo@ / egyseg@) * arf@
      End If
      forint = xval(Format(atf@, "##########0.00"))
    Case "E"
      arf@ = xval(Mid$(arfrec$, 44, 10))
      If devnem$ = "HUF" And konyveldevnem <> "   " Then
        If arf@ <> 0 Then atf@ = (devo@ / egyseg@) / arf@ Else atf@ = -1
      Else
        atf@ = (devo@ / egyseg@) * arf@
      End If
      forint = xval(Format(atf@, "##########0.00"))
    Case "S"
      devk@ = xval(Mid$(devrec$, 12, 14))
      fint@ = xval(Mid$(devrec$, 26, 14))
      If devk@ <> 0 Then
        If devnem$ = "HUF" And konyveldevnem <> "   " Then
          If fimt@ <> 0 Then
            arf@ = devk@ / fint@
            If arf@ <> 0 Then
              atf@ = devo@ / arf@
            Else
              forint = -1: Exit Function
            End If
          Else
            forint = -1: Exit Function
          End If
        Else
          arf@ = fint@ / devk@
          atf@ = devo@ * arf@
        End If
        forint = xval(Format(atf@, "##########0.00"))
      Else
        forint = -1: Exit Function
      End If
    Case Else
  End Select
End Function

Public Sub arfokul(osb@, xforintszamla@, xdevszamla@, xforintbank@, xdevbank@, vsmod$, arfkul@, ellenszamla$)
  '--- pénzforgalom árfolyamkülönbözetének meghatározása
  szarfolyam# = xforintszamla@ / xdevszamla@
  szarfolyamon@ = osb@ * szarfolyam#
  '--- banki árfolyamon
  barfolyam# = xforintbank@ / xdevbank@
  barfolyamon@ = osb@ * barfolyam#
  kul@ = barfolyamon@ - szarfolyamon@
  If vsmod$ = "V" Then
    If kul@ > 0 Then
      '--- vevõ árfolyamnyereség
      arfkul@ = kul@
      ellenszamla$ = Mid$(irec$, 538, 8)
    End If
    If kul@ < 0 Then
      '--- vevõ árfolyamveszteség
      arfkul@ = kul@
      ellenszamla$ = Mid$(irec$, 546, 8)
    End If
  Else
    If kul@ > 0 Then
      '--- száll.árfolyamveszteség
      arfkul@ = -kul@
      ellenszamla$ = Mid$(irec$, 562, 8)
    End If
    If kul@ < 0 Then
      '--- szállító árfolyamnyereség
      arfkul@ = -kul@
      ellenszamla$ = Mid$(irec$, 554, 8)
    End If
  End If
End Sub

Public Sub forintosit(devpoz%, devnempoz%, forintpoz%, bankszamla$, erteknap$, arfkod$, tablakod$, vsor&)
  '--- deviza konverzió adatrögzítés közben
  '--- a deviza összeget a forintpozicióra be is írja
  '--- arfkod=V-vételi K-közép E-eladási
  '--- tablakod$=T-táblázat V-vektor
  If tablakod$ = "T" Then
    '--- táblázat
    dvo@ = xval(Tabla.MSFlexGrid1.TextMatrix(vsor&, devpoz%))
    fto@ = xval(Tabla.MSFlexGrid1.TextMatrix(vsor&, forintpoz%))
    If dvo@ = 0 Then Tabla.MSFlexGrid1.TextMatrix(vsor&, devnempoz%) = "   "
    dnm$ = Tabla.MSFlexGrid1.TextMatrix(vsor&, devnempoz%)
  Else
    '--- vektor
    dvo@ = xval(Vektor.MSFlexGrid1.TextMatrix(devpoz%, 1))
    fto@ = xval(Vektor.MSFlexGrid1.TextMatrix(forintpoz%, 1))
    If dvo@ = 0 Then Vektor.MSFlexGrid1.TextMatrix(devnempoz%, 1) = "   "
    dnm$ = Vektor.MSFlexGrid1.TextMatrix(devnempoz%, 1)
  End If
  '---
  If Trim$(dnm$) = "" Then
    '--- forint összeg van beütve
    dvo@ = 0
    If tablakod$ = "T" Then
      Tabla.MSFlexGrid1.TextMatrix(vsor&, devpoz%) = Space$(14)
    Else
      Vektor.MSFlexGrid1.TextMatrix(devpoz%, 1) = Space$(14)
    End If
  Else
    dkod$ = bankszamla$ + dnm$
    devrec$ = dbxkey("PDEV", dkod$)
    If dnm$ = "HUF" And konyveldat$ <> "   " Then
      If devrec$ = "" Then
        dkod$ = Mid$(irec$, 470, 8) + dnm$
        devrec$ = dbxkey("PDEV", dkod$)
      End If
    End If
    If devrec$ = "" Then
      '--- nincs pdev rekord
    Else
      If arfkod$ = "S" Then
        '--- deviza készletrõl
        devk@ = xval(Mid$(devrec$, 12, 14))
        fint@ = xval(Mid$(devrec$, 26, 14))
        If devk@ <> 0 Then
          If dnm$ = "HUF" And konyveldevnem <> "   " Then
            If fint@ <> 0 Then
              arf@ = devk@ / fint@
              If arf@ <> 0 Then
                atf@ = dvo@ / arf@
                fto@ = xval(Format(atf@, "##########0.00"))
              Else
                fto@ = 0
              End If
            Else
              fto@ = 0
            End If
          Else
            arf@ = fint@ / devk@
            atf@ = dvo@ * arf@
            fto@ = xval(Format(atf@, "##########0.00"))
          End If
        Else
          fto@ = 0
        End If
      Else
        '--- arfolyamokból
        dkod$ = bankszamla$ + dnm$ + erteknap$
        arfrec$ = dbxkey("PDRF", dkod$)
        If dnm$ = "HUF" And konyveldat$ <> "   " Then
          If arfrec$ = "" Then
            dkod$ = Mid$(irec$, 470, 8) + dnm$ + erteknap$
            arfrec$ = dbxkey("PDRF", dkod$)
          End If
        End If
        If arfrec$ = "" Then
          '--- nincs árfolyam
        Else
          egyseg@ = xval(Mid$(arfrec$, 18, 6))
          If egyseg@ = 0 Then egyseg@ = 1
          Select Case arfkod$
            Case "V"
              arf@ = xval(Mid$(arfrec$, 24, 10))
              If arf@ <> 0 Then
                If dnm$ = "HUF" And konyveldevnem <> "   " Then
                  atf@ = (dvo@ / egyseg@) / arf@
                Else
                  atf@ = (dvo@ / egyseg@) * arf@
                End If
                fto@ = xval(Format(atf@, "##########0.00"))
              Else
                fto@ = 0
              End If
            Case "K"
              arf@ = xval(Mid$(arfrec$, 34, 10))
              If arf@ <> 0 Then
                If dnm$ = "HUF" And konyveldevnem <> "   " Then
                  atf@ = (dvo@ / egyseg@) / arf@
                Else
                  atf@ = (dvo@ / egyseg@) * arf@
                End If
                fto@ = xval(Format(atf@, "##########0.00"))
              Else
                fto@ = 0
              End If
            Case "E"
              arf@ = xval(Mid$(arfrec$, 44, 10))
              If arf@ <> 0 Then
                If dnm$ = "HUF" And konyveldevnem <> "   " Then
                  atf@ = (dvo@ / egyseg@) / arf@
                Else
                  atf@ = (dvo@ / egyseg@) * arf@
                End If
                fto@ = xval(Format(atf@, "##########0.00"))
              Else
                fto@ = 0
              End If
            Case Else
          End Select
        End If
      End If
    End If
  End If
  If tablakod$ = "T" Then
    '--- táblázat
    If fto@ = 0 Then
      Tabla.MSFlexGrid1.TextMatrix(vsor&, forintpoz%) = Space$(14)
    Else
      Tabla.MSFlexGrid1.TextMatrix(vsor&, forintpoz%) = ertszam(Str$(fto@), 14, 2)
    End If
  Else
    '--- vektor
    If fto@ = 0 Then
      Vektor.MSFlexGrid1.TextMatrix(forintpoz%, 1) = Space$(14)
      Vektor.MSFlexGrid1.TextMatrix(2, 1) = Space$(15)
    Else
      Vektor.MSFlexGrid1.TextMatrix(forintpoz%, 1) = ertszam(Str$(fto@), 14, 2)
      Vektor.MSFlexGrid1.TextMatrix(2, 1) = "Árf.:" + Str$(arf@)
    End If
  End If
End Sub

Public Sub devszamegy(szrec$, vsmod$, fnap$, egyft@, egydev@, dnem$)
  '--- devizás számla egyenlegének meghatározása
  knap$ = fnap$
  kfnap$ = fnap$
  ysdev$ = Mid$(szrec$, 92, 3)
  If dtm(Mid$(szrec$, 58, 6)) <= dtm(fnap$) And (Mid$(szrec$, 166, 1) = " " Or dtm(Mid$(szrec$, 167, 6)) > fnap$) Then
    egyft@ = xval(Mid$(szrec$, 78, 14))
    egydev@ = xval(Mid$(szrec$, 95, 14))
    dnem$ = Mid$(szrec$, 92, 3)
  Else
    egyft@ = 0
    egydev@ = 0
    dnem$ = Mid$(szrec$, 92, 3)
  End If
  '--- helyesbítések
  hkod$ = vsmod$ + Mid$(szrec$, 1, 7)
  hrec$ = dbxkey("PSHL", hkod$)
  If hrec$ <> "" Then
    For j11% = 1 To 3
      h1rec$ = Mid$(hrec$, (j11% - 1) * 790 + 1, 780)
      If dtm(Mid$(h1rec$, 30, 6)) <= dtm(fnap$) And (Mid$(h1rec$, 86, 1) = " " Or dtm(Mid$(h1rec$, 87, 6)) > fnap$) Then
        If Mid$(h1rec$, 86, 1) <> "S" Then
          fto@ = xval(Mid$(h1rec$, 36, 14))
          dvo@ = xval(Mid$(h1rec$, 53, 14))
          egyft@ = egyft@ + fto@
          egydev@ = egydev@ + dvo@
        End If
      End If
    Next
  End If
  '--- kiegyenlítések
  tullog% = 0
  elem$ = Mid$(szrec$, (10 - 1) * 35 + 930, 35)
  If Trim(elem$) <> "" Then
    iikta$ = Mid$(szrec$, 1, 7)
    If vsmod$ = "V" Then pvskrec$ = dbxkey("PVSK", iikta$) Else pvskrec$ = dbxkey("PSSK", iikta$)
    If pvskrec$ <> "" Then
      elem$ = Mid$(pvskrec$, (20 - 1) * 35 + 2130, 35)
      If Trim(elem$) <> "" Then tullog% = 1
    Else
      tullog% = 1
    End If
  End If
  If tullog% = 0 Then
    '--- egyszerûen kiolvassuk számlából
    kikdb% = 0
    For j11% = 1 To 10
      elem$ = Mid$(szrec$, (j11% - 1) * 35 + 930, 35)
      If Trim(elem$) <> "" Then
        kikdb% = kikdb% + 1
        pikt$ = Mid$(elem$, 29, 7)
        fkrec$ = dbxkey("FKTE", pikt$)
        If fkrec$ <> "" Then
          ko@ = xval(Mid$(elem$, 1, 14))
          ybdev$ = Mid$(fkrec$, 146, 3)
          If ysdev$ <> ybdev$ Then
            pfikt$ = Mid$(fkrec$, 185, 7)
            ykxrec$ = dbxkey("PXRF", pfikt)
            If ykxrec$ <> "" Then
              If Mid$(ykxrec$, 11, 3) = ysdev$ And Mid$(ykxrec$, 8, 3) = ybdev$ Then
                ykxarf# = xval(Mid$(ykxrec$, 14, 14))
                ydvo@ = xval(Mid$(fkrec$, 149, 14)) * ykxarf#
                Mid$(fkrec$, 146, 3) = ysdev$
                Mid$(fkrec$, 149, 14) = ertszam(Str(ydvo@), 14, 2)
              End If
            End If
          End If
          If dtm(Mid$(fkrec$, 49, 6)) <= dtm(fnap$) And Mid$(fkrec$, 61, 1) = " " Then
            fto@ = xval(Mid$(fkrec$, 132, 14))
            dvo@ = xval(Mid$(fkrec$, 149, 14))
            If ko@ < 0 Then fto@ = -fto@: dvo@ = -dvo@
            egyft@ = egyft@ - fto@
            egydev@ = egydev@ - dvo@
          End If
        End If
      End If
    Next
    If kikdb% = 10 Then
      iikta$ = Mid$(szrec$, 1, 7)
      If vsmod$ = "V" Then pvskrec$ = dbxkey("PVSK", iikta$) Else pvskrec$ = dbxkey("PSSK", iikta$)
      If pvskrec$ <> "" Then
        For j11% = 1 To 20
          elem$ = Mid$(pvskrec$, (j11% - 1) * 35 + 2130, 35)
          If Trim(elem$) <> "" Then
            pikt$ = Mid$(elem$, 29, 7)
            fkrec$ = dbxkey("FKTE", pikt$)
            ko@ = xval(Mid$(elem$, 1, 14))
            If fkrec$ <> "" Then
              ybdev$ = Mid$(fkrec$, 146, 3)
              If ysdev$ <> ybdev$ Then
                pfikt$ = Mid$(fkrec$, 185, 7)
                ykxrec$ = dbxkey("PXRF", pfikt)
                If ykxrec$ <> "" Then
                  If Mid$(ykxrec$, 11, 3) = ysdev$ And Mid$(ykxrec$, 8, 3) = ybdev$ Then
                    ykxarf# = xval(Mid$(ykxrec$, 14, 14))
                    ydvo@ = xval(Mid$(fkrec$, 149, 14)) * ykxarf#
                    Mid$(fkrec$, 146, 3) = ysdev$
                    Mid$(fkrec$, 149, 14) = ertszam(Str(ydvo@), 14, 2)
                  End If
                End If
              End If
              If dtm(Mid$(fkrec$, 49, 6)) <= dtm(fnap$) And Mid$(fkrec$, 61, 1) = " " Then
                fto@ = xval(Mid$(fkrec$, 132, 14))
                dvo@ = xval(Mid$(fkrec$, 149, 14))
                If ko@ < 0 Then fto@ = -fto@: dvo@ = -dvo@
                egyft@ = egyft@ - fto@
                egydev@ = egydev@ - dvo@
              End If
            End If
          End If
        Next
      End If
    End If
  Else
    '--- lassú módszer a partner láncokról szedjük le
    vdfi = FreeFile
    Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #vdfi
    vpkod$ = Mid$(szrec$, 38, 15)
    vpartrec$ = dbxkey("PART", vpkod$)
    pktecim& = xval(Mid$(vpartrec$, 742, 10))
    pbnkcim& = xval(Mid$(vpartrec$, 752, 10))
    pkorcim& = xval(Mid$(vpartrec$, 762, 10))
    '--- pénztár
    Do While pktecim& > 0
      pkterec$ = Space(270)
      Get #vdfi, pktecim& + 9, pkterec$
      vkell% = 0
      If dtm(Mid$(pkterec$, 16, 6)) <= dtm(fnap$) And Mid$(pkterec$, 192, 1) <> "S" Then
        If vsmod$ = "V" Then
          If Mid$(pkterec$, 55, 1) = "V" And Mid$(pkterec$, 87, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        Else
          If Mid$(pkterec$, 55, 1) = "S" And Mid$(pkterec$, 94, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        End If
        If vkell% = 1 Then
          fto@ = xval(Mid$(pkterec$, 56, 14))
          dvo@ = xval(Mid$(pkterec$, 73, 14))
          If vsmod$ = "V" And Mid$(pkterec$, 8, 1) = "B" Or vsmod$ = "S" And Mid$(pkterec$, 8, 1) = "K" Then
            egyft@ = egyft@ - fto@
            egydev@ = egydev@ - dvo@
          Else
            egyft@ = egyft@ + fto@
            egydev@ = egydev@ + dvo@
          End If
        End If
      End If
      pktecim& = xval(Mid$(pkterec$, 220, 10))
    Loop
    '--- bank
    Do While pbnkcim& > 0
      pbnkrec$ = Space(270)
      Get #vdfi, pbnkcim& + 9, pbnkrec$
      ybdev$ = Mid$(pbnkrec$, 70, 3)
      If ysdev$ <> ybdev$ Then
        pfikt$ = Mid$(pbnkrec$, 1, 7)
        ykxrec$ = dbxkey("PXRF", pfikt$)
        If ykxrec$ <> "" Then
          If Mid$(ykxrec$, 11, 3) = ysdev$ And Mid$(ykxrec$, 8, 3) = ybdev$ Then
            ykxarf# = xval(Mid$(ykxrec$, 14, 14))
            ydvo@ = xval(Mid$(pbnkrec$, 73, 14)) * ykxarf#
            Mid$(pbnkrec$, 70, 3) = ysdev$
            Mid$(pbnkrec$, 73, 14) = ertszam(Str(ydvo@), 14, 2)
          End If
        End If
      End If
      vkell% = 0
      If dtm(Mid$(pbnkrec$, 16, 6)) <= dtm(knap$) And Mid$(pbnkrec$, 192, 1) <> "S" Then
        If vsmod$ = "V" Then
          If Mid$(pbnkrec$, 55, 1) = "V" And Mid$(pbnkrec$, 87, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        Else
          If Mid$(pbnkrec$, 55, 1) = "S" And Mid$(pbnkrec$, 94, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        End If
        If vkell% = 1 Then
          fto@ = xval(Mid$(pbnkrec$, 56, 14))
          dvo@ = xval(Mid$(pbnkrec$, 73, 14))
          If vsmod$ = "V" And Mid$(pbnkrec$, 8, 1) = "J" Or vsmod$ = "S" And Mid$(pbnkrec$, 8, 1) = "T" Then
            egyft@ = egyft@ - fto@
            egydev@ = egydev@ - dvo@
          Else
            egyft@ = egyft@ + fto@
            egydev@ = egydev@ + dvo@
          End If
        End If
      End If
      pbnkcim& = xval(Mid$(pbnkrec$, 220, 10))
    Loop
    '--- korrekció
    Do While pkorcim& > 0
      pkorrec$ = Space(600)
      Get #vdfi, pkorcim& + 9, pkorrec$
      vkell% = 0
      If dtm(Mid$(pkorrec$, 8, 6)) <= dtm(kfnap$) And Mid$(pkorrec$, 120, 1) <> "S" Then
        If vsmod$ = "V" Then
          If (Mid$(pkorrec$, 14, 1) = "K" Or Mid$(pkorrec$, 14, 1) = "V") And Mid$(pkorrec$, 31, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        Else
          If (Mid$(pkorrec$, 14, 1) = "K" Or Mid$(pkorrec$, 14, 1) = "S") And Mid$(pkorrec$, 38, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        End If
        If vkell% = 1 Then
          fto@ = xval(Mid$(pkorrec$, 45, 14))
          dvo@ = xval(Mid$(pkorrec$, 62, 14))
          egyft@ = egyft@ - fto@
          egydev@ = egydev@ - dvo@
        End If
      End If
      pkorcim& = xval(Mid$(pkorrec$, 135, 10))
    Loop
    Close vdfi
  End If
  '--- elõlegbeszámítások
  For j11% = 1 To 5
    elem$ = Mid$(szrec$, (j11% - 1) * 43 + 1280, 43)
    pikt$ = Mid$(elem$, 37, 7)
    pelorec$ = dbxkey("PELO", pikt$)
    If pelorec$ <> "" Then
      If dtm(Mid$(pelorec$, 38, 6)) <= dtm(fnap$) And Mid$(pelorec$, 90, 1) = " " Then
        ko@ = xval(Mid$(elem$, 23, 14))
        fto@ = xval(Mid$(pelorec$, 44, 14))
        dvo@ = xval(Mid$(pelorec$, 61, 14))
        If ko@ > 0 Then fto@ = -fto@: dvo@ = -dvo@
        egyft@ = egyft@ - fto@
        egydev@ = egydev@ - dvo@
      End If
    End If
  Next
End Sub
Public Sub szamlaegyenleg(szrec$, osszeg@, helybit@, kiegy@, egyenleg@, vsmod$, fnap$, kfnap$, forintegyenleg@)
  '--- szállító illetve vevõ számla egyenlegének meghatároztása
  ysdev$ = Mid$(szrec$, 92, 3)
  If dtm(Mid$(szrec$, 58, 6)) <= dtm(fnap$) And (Mid$(szrec$, 166, 1) = " " Or dtm(Mid$(szrec$, 167, 6)) > fnap$) Then
    fto@ = xval(Mid$(szrec$, 78, 14))
    dvo@ = xval(Mid$(szrec$, 95, 14))
  Else
    fto@ = 0
    dvo@ = 0
  End If
  dnm$ = Mid$(szrec$, 92, 3)
  'If dnm$ <> "   " And dnm$ <> langmodul(155) Then
  
  If dnm$ <> "   " And dnm$ <> Mid$(irec$, 666, 3) Then
    devi% = 1
    egyenleg@ = dvo@: osszeg@ = dvo@: kiegy@ = 0
    forintegyenleg@ = fto@
  Else
    devi% = 0
    egyenleg@ = fto@: osszeg@ = fto@: kiegy@ = 0
  End If
  '--- helyesbítések
  hkod$ = vsmod$ + Mid$(szrec$, 1, 7)
  hrec$ = dbxkey("PSHL", hkod$)
  helybit@ = 0
  If hrec$ <> "" Then
    For j11% = 1 To 3
      h1rec$ = Mid$(hrec$, (j11% - 1) * 790 + 1, 780)
      If dtm(Mid$(h1rec$, 30, 6)) <= dtm(fnap$) And (Mid$(h1rec$, 86, 1) = " " Or dtm(Mid$(h1rec$, 87, 6)) > fnap$) Then
        If Mid$(h1rec$, 86, 1) <> "S" Then
          fto@ = xval(Mid$(h1rec$, 36, 14))
          dvo@ = xval(Mid$(h1rec$, 53, 14))
          If devi% = 0 Then
            egyenleg@ = egyenleg@ + fto@
            osszeg@ = osszeg@ + fto@
            helybit@ = helybit@ + fto@
          Else
            forintegyenleg@ = forintegyenleg + fto@
            egyenleg@ = egyenleg@ + dvo@
            osszeg@ = osszeg@ + dvo@
            helybit@ = helybit@ + dvo@
          End If
        End If
      End If
    Next
  End If
  '--- kiegyenlítések
  tullog% = 0
  elem$ = Mid$(szrec$, (10 - 1) * 35 + 930, 35)
  If Trim(elem$) <> "" Then
    iikta$ = Mid$(szrec$, 1, 7)
    If vsmod$ = "V" Then pvskrec$ = dbxkey("PVSK", iikta$) Else pvskrec$ = dbxkey("PSSK", iikta$)
    If pvskrec$ <> "" Then
      elem$ = Mid$(pvskrec$, (20 - 1) * 35 + 2130, 35)
      If Trim(elem$) <> "" Then tullog% = 1
    Else
      tullog% = 1
    End If
  End If
  If tullog% = 0 Then
    '--- egyszerûen kiolvassuk számlából
    kikdb% = 0
    For j11% = 1 To 10
      elem$ = Mid$(szrec$, (j11% - 1) * 35 + 930, 35)
      If Trim(elem$) <> "" Then
        kikdb% = kikdb% + 1
        pikt$ = Mid$(elem$, 29, 7)
        fkrec$ = dbxkey("FKTE", pikt$)
        ko@ = xval(Mid$(elem$, 1, 14))
        ybdev$ = Mid$(fkrec$, 146, 3)
        If ysdev$ <> ybdev$ Then
          pfikt$ = Mid$(fkrec$, 185, 7)
          ykxrec$ = dbxkey("PXRF", pfikt)
          If ykxrec$ <> "" Then
            If Mid$(ykxrec$, 11, 3) = ysdev$ And Mid$(ykxrec$, 8, 3) = ybdev$ Then
              ykxarf# = xval(Mid$(ykxrec$, 14, 14))
              ydvo@ = xval(Mid$(fkrec$, 149, 14)) * ykxarf#
              Mid$(fkrec$, 146, 3) = ysdev$
              Mid$(fkrec$, 149, 14) = ertszam(Str(ydvo@), 14, 2)
            End If
          End If
        End If
        If fkrec$ <> "" Then
          If dtm(Mid$(fkrec$, 49, 6)) <= dtm(kfnap$) And Mid$(fkrec$, 61, 1) = " " Then
            fto@ = xval(Mid$(fkrec$, 132, 14))
            dvo@ = xval(Mid$(fkrec$, 149, 14))
            If ko@ < 0 Then fto@ = -fto@: dvo@ = -dvo@
            If devi% = 0 Then
              egyenleg@ = egyenleg@ - fto@
              kiegy@ = kiegy@ + fto@
            Else
              forintegyenleg@ = forintegyenleg@ - fto@
              egyenleg@ = egyenleg@ - dvo@
              kiegy@ = kiegy@ + dvo@
            End If
          End If
        End If
      End If
    Next
    If kikdb% = 10 Then
      iikta$ = Mid$(szrec$, 1, 7)
      If vsmod$ = "V" Then pvskrec$ = dbxkey("PVSK", iikta$) Else pvskrec$ = dbxkey("PSSK", iikta$)
      If pvskrec$ <> "" Then
        For j11% = 1 To 20
          elem$ = Mid$(pvskrec$, (j11% - 1) * 35 + 2130, 35)
          If Trim(elem$) <> "" Then
            pikt$ = Mid$(elem$, 29, 7)
            fkrec$ = dbxkey("FKTE", pikt$)
            If fkrec$ <> "" Then
              ko@ = xval(Mid$(elem$, 1, 14))
              ybdev$ = Mid$(fkrec$, 146, 3)
              If ysdev$ <> ybdev$ Then
                pfikt$ = Mid$(fkrec$, 185, 7)
                ykxrec$ = dbxkey("PXRF", pfikt$)
                If ykxrec$ <> "" Then
                  If Mid$(ykxrec$, 11, 3) = ysdev$ And Mid$(ykxrec$, 8, 3) = ybdev$ Then
                    ykxarf# = xval(Mid$(ykxrec$, 14, 14))
                    ydvo@ = xval(Mid$(fkrec$, 149, 14)) * ykxarf#
                    Mid$(fkrec$, 146, 3) = ysdev$
                    Mid$(fkrec$, 149, 14) = ertszam(Str(ydvo@), 14, 2)
                  End If
                End If
              End If
              If dtm(Mid$(fkrec$, 49, 6)) <= dtm(kfnap$) And Mid$(fkrec$, 61, 1) = " " Then
                fto@ = xval(Mid$(fkrec$, 132, 14))
                dvo@ = xval(Mid$(fkrec$, 149, 14))
                If ko@ < 0 Then fto@ = -fto@: dvo@ = -dvo@
                If devi% = 0 Then
                  egyenleg@ = egyenleg@ - fto@
                  kiegy@ = kiegy@ + fto@
                Else
                  forintegyenleg@ = forintegyenleg@ - fto@
                  egyenleg@ = egyenleg@ - dvo@
                  kiegy@ = kiegy@ + dvo@
                End If
              End If
            End If
          End If
        Next
      End If
    End If
  Else
    '--- lassú módszer a partner láncokról szedjük le
    vdfi = FreeFile
    Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #vdfi
    vpkod$ = Mid$(szrec$, 38, 15)
    vpartrec$ = dbxkey("PART", vpkod$)
    pktecim& = xval(Mid$(vpartrec$, 742, 10))
    pbnkcim& = xval(Mid$(vpartrec$, 752, 10))
    pkorcim& = xval(Mid$(vpartrec$, 762, 10))
    '--- pénztár
    Do While pktecim& > 0
      pkterec$ = Space(270)
      Get #vdfi, pktecim& + 9, pkterec$
      vkell% = 0
      If dtm(Mid$(pkterec$, 16, 6)) <= dtm(kfnap$) And Mid$(pkterec$, 192, 1) <> "S" Then
        If vsmod$ = "V" Then
          If Mid$(pkterec$, 55, 1) = "V" And Mid$(pkterec$, 87, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        Else
          If Mid$(pkterec$, 55, 1) = "S" And Mid$(pkterec$, 94, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        End If
        If vkell% = 1 Then
          fto@ = xval(Mid$(pkterec$, 56, 14))
          dvo@ = xval(Mid$(pkterec$, 73, 14))
          If vsmod$ = "V" And Mid$(pkterec$, 8, 1) = "B" Or vsmod$ = "S" And Mid$(pkterec$, 8, 1) = "K" Then
            If devi% = 0 Then
              egyenleg@ = egyenleg@ - fto@
              kiegy@ = kiegy@ + fto@
            Else
              forintegyenleg@ = forintegyenleg@ - fto@
              egyenleg@ = egyenleg@ - dvo@
              kiegy@ = kiegy@ + dvo@
            End If
          Else
            If devi% = 0 Then
              egyenleg@ = egyenleg@ + fto@
              kiegy@ = kiegy@ - fto@
            Else
              forintegyenleg@ = forintegyenleg@ + fto@
              egyenleg@ = egyenleg@ + dvo@
              kiegy@ = kiegy@ - dvo@
            End If
          End If
        End If
      End If
      pktecim& = xval(Mid$(pkterec$, 220, 10))
    Loop
    '--- bank
    Do While pbnkcim& > 0
      pbnkrec$ = Space(270)
      Get #vdfi, pbnkcim& + 9, pbnkrec$
      ybdev$ = Mid$(pbnkrec$, 70, 3)
      If ysdev$ <> ybdev$ Then
        pfikt$ = Mid$(pbnkrec$, 1, 7)
        ykxrec$ = dbxkey("PXRF", pfikt$)
        If ykxrec$ <> "" Then
          If Mid$(ykxrec$, 11, 3) = ysdev$ And Mid$(ykxrec$, 8, 3) = ybdev$ Then
            ykxarf# = xval(Mid$(ykxrec$, 14, 14))
            ydvo@ = xval(Mid$(pbnkrec$, 73, 14)) * ykxarf#
            Mid$(pbnkrec$, 70, 3) = ysdev$
            Mid$(pbnkrec$, 73, 14) = ertszam(Str(ydvo@), 14, 2)
          End If
        End If
      End If
      vkell% = 0
      If dtm(Mid$(pbnkrec$, 16, 6)) <= dtm(kfnap$) And Mid$(pbnkrec$, 192, 1) <> "S" Then
        If vsmod$ = "V" Then
          If Mid$(pbnkrec$, 55, 1) = "V" And Mid$(pbnkrec$, 87, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        Else
          If Mid$(pbnkrec$, 55, 1) = "S" And Mid$(pbnkrec$, 94, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        End If
        If vkell% = 1 Then
          fto@ = xval(Mid$(pbnkrec$, 56, 14))
          dvo@ = xval(Mid$(pbnkrec$, 73, 14))
          If vsmod$ = "V" And Mid$(pbnkrec$, 8, 1) = "J" Or vsmod$ = "S" And Mid$(pbnkrec$, 8, 1) = "T" Then
            If devi% = 0 Then
              egyenleg@ = egyenleg@ - fto@
              kiegy@ = kiegy@ + fto@
            Else
              forintegyenleg@ = forintegyenleg@ - fto@
              egyenleg@ = egyenleg@ - dvo@
              kiegy@ = kiegy@ + dvo@
            End If
          Else
            If devi% = 0 Then
              egyenleg@ = egyenleg@ + fto@
              kiegy@ = kiegy@ - fto@
            Else
              forintegyenleg@ = forintegyenleg@ + fto@
              egyenleg@ = egyenleg@ + dvo@
              kiegy@ = kiegy@ - dvo@
            End If
          End If
        End If
      End If
      pbnkcim& = xval(Mid$(pbnkrec$, 220, 10))
    Loop
    '--- korrekció
    Do While pkorcim& > 0
      pkorrec$ = Space(600)
      Get #vdfi, pkorcim& + 9, pkorrec$
      vkell% = 0
      If dtm(Mid$(pkorrec$, 8, 6)) <= dtm(kfnap$) And Mid$(pkorrec$, 120, 1) <> "S" Then
        If vsmod$ = "V" Then
          If (Mid$(pkorrec$, 14, 1) = "K" Or Mid$(pkorrec$, 14, 1) = "V") And Mid$(pkorrec$, 31, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        Else
          If (Mid$(pkorrec$, 14, 1) = "K" Or Mid$(pkorrec$, 14, 1) = "S") And Mid$(pkorrec$, 38, 7) = Mid$(szrec$, 1, 7) Then vkell% = 1
        End If
        If vkell% = 1 Then
          fto@ = xval(Mid$(pkorrec$, 45, 14))
          dvo@ = xval(Mid$(pkorrec$, 62, 14))
          If devi% = 0 Then
            egyenleg@ = egyenleg@ - fto@
            kiegy@ = kiegy@ + fto@
          Else
            forintegyenleg@ = forintegyenleg@ - fto@
            egyenleg@ = egyenleg@ - dvo@
            kiegy@ = kiegy@ + dvo@
          End If
        End If
      End If
      pkorcim& = xval(Mid$(pkorrec$, 135, 10))
    Loop
    Close vdfi
  End If
  '--- elõlegbeszámítások
  For j11% = 1 To 5
    elem$ = Mid$(szrec$, (j11% - 1) * 43 + 1280, 43)
    If Trim(elem$) <> "" Then
      pikt$ = Mid$(elem$, 37, 7)
      pelorec$ = dbxkey("PELO", pikt$)
      If pelorec$ <> "" Then
        If dtm(Mid$(pelorec$, 38, 6)) <= dtm(kfnap$) And Mid$(pelorec$, 90, 1) = " " Then
          ko@ = xval(Mid$(elem$, 23, 14))
          fto@ = xval(Mid$(pelorec$, 44, 14))
          dvo@ = xval(Mid$(pelorec$, 61, 14))
          If ko@ > 0 Then fto@ = -fto@: dvo@ = -dvo@
          If devi% = 0 Then
            egyenleg@ = egyenleg@ - fto@
            kiegy@ = kiegy@ + fto@
          Else
            forintegyenleg@ = forintegyenleg@ - fto@
            egyenleg@ = egyenleg@ - dvo@
            kiegy@ = kiegy@ + dvo@
          End If
        End If
      End If
    End If
  Next
End Sub

Public Function pscale(nval&, nmax&)
  '--- nval értékének beállítása 1 és 100 közé
  nmun& = nval& / (nmax& / 100) + 1
  If nmun& > 100 Then nmun& = 100
  pscale = nmun&
  If pscale < 1 Then pscale = 1
End Function

Public Function elozoho(dat$)
  ev% = Val(Mid$(dat$, 1, 2))
  ho% = Val(Mid$(dat$, 3, 2))
  If ho% = 1 Then
    If ev% = 0 Then ev% = 99 Else ev% = ev% - 1
    ho% = 12
  Else
    ho% = ho% - 1
  End If
  elozoho = Right("00" + Trim(Str(ev%)), 2) + Right("00" + Trim(Str(ho%)), 2) + "01"
End Function

Public Function afaosszes(rec$, tip$)
  '--- Áfa tartalom meghatározása
  '--- tip=VS vevõszámla
  '---     SS szállító számla
  '---     VH vevõ helyesbítõ
  '---     SH szállító helyesbítõ
  '---     VE vevõ elõleg
  '---     SE szállító elõleg
  Select Case tip$
    Case "VS", "SS"
      '--- számla áfája
      afa@ = 0
      For i97% = 1 To 5
        elem$ = Mid$(rec$, (i97% - 1) * 30 + 250, 30)
        afa1@ = xval(Mid$(elem$, 17, 14))
        If afa1@ <> 0 Then
          afkod$ = Mid$(elem$, 1, 2)
          afrec$ = dbxkey("PAFA", afkod$)
          If Left$(tip$, 1) = "S" And (Mid$(afrec$, 40, 2) = "IU" Or Mid$(afrec$, 40, 2) = "BR" Or Mid$(afrec$, 42, 1) = "N") Then
          Else
            If Left$(tip$, 1) = "S" And Mid$(afrec$, 42, 1) = "R" Then
              szaz@ = xval(Mid$(afrec$, 43, 6))
              afa1@ = afa1@ * (szaz@ / 100)
              afa@ = afa@ + afa1@
            Else
              afa@ = afa@ + afa1@
            End If
          End If
        End If
      Next
    Case "VH", "SH"
      '--- helyesbítõ áfája
      afa@ = 0
      For i97% = 1 To 5
        elem$ = Mid$(rec$, (i97% - 1) * 30 + 110, 30)
        afa1@ = xval(Mid$(elem$, 17, 14))
        If afa1@ <> 0 Then
          afkod$ = Mid$(elem$, 1, 2)
          afrec$ = dbxkey("PAFA", afkod$)
          If Left$(tip$, 1) = "S" And (Mid$(afrec$, 40, 2) = "IU" Or Mid$(afrec$, 40, 2) = "BR" Or Mid$(afrec$, 42, 1) = "N") Then
          Else
            If Left$(tip$, 1) = "S" And Mid$(afrec$, 42, 1) = "R" Then
              szaz@ = xval(Mid$(afrec$, 43, 6))
              afa1@ = afa1@ * (szaz@ / 100)
              afa@ = afa@ + afa1@
            Else
              afa@ = afa@ + afa1@
            End If
          End If
        End If
      Next
    Case "VE", "SE"
      afa@ = 0
      For i97% = 1 To 5
        elem$ = Mid$(rec$, (i97% - 1) * 30 + 230, 30)
        afa1@ = xval(Mid$(elem$, 17, 14))
        If afa1@ <> 0 Then
          afkod$ = Mid$(elem$, 1, 2)
          afrec$ = dbxkey("PAFA", afkod$)
          If Left$(tip$, 1) = "S" And (Mid$(afrec$, 40, 2) = "IU" Or Mid$(afrec$, 40, 2) = "BR" Or Mid$(afrec$, 42, 1) = "N") Then
          Else
            If Left$(tip$, 1) = "S" And Mid$(afrec$, 42, 1) = "R" Then
              szaz@ = xval(Mid$(afrec$, 43, 6))
              afa1@ = afa1@ * (szaz@ / 100)
              afa@ = afa@ + afa1@
            Else
              afa@ = afa@ + afa1@
            End If
          End If
        End If
      Next
    Case Else
  End Select
  afaosszes = afa@
End Function

Public Function afadatum(dat$)
  ev% = Val(Mid$(dat$, 1, 2))
  ho% = Val(Mid$(dat$, 3, 2))
  If ho% = 12 Then
    If ev% = 99 Then ev% = 0 Else ev% = ev% + 1
    ho% = 1
  Else
    ho% = ho% + 1
  End If
  afadatum = Right("00" + Trim(Str(ev%)), 2) + Right("00" + Trim(Str(ho%)), 2) + "19"
End Function

Public Function hetsor%(dat$)
  '--- hét sorszámának meghatározása
  evy% = Val(Mid$(dat$, 1, 2))
  hoy% = Val(Mid$(dat$, 3, 2))
  nay% = Val(Mid$(dat$, 5, 2))
  hanyadiknap% = DateSerial(evy%, hoy%, nay%) - DateSerial(evy% - 1, 12, 31)
  dax$ = "01-01-20" + Mid$(dat$, 1, 2)
  evelso% = WeekDay(dax$, vbMonday)
  elsohetnapjai% = 8 - evelso%
  If hanyadiknap% <= elsohetnapjai% Then
    hetsor% = 1
  Else
    tovabbinapok% = hanyadiknap% - elsohetnapjai%
    tovabbihetek% = Int(tovabbinapok% / 7)
    If tovabbinapok% Mod 7 <> 0 Then
      hetsor% = tovabbihetek% + 2
    Else
      hetsor% = tovabbihetek% + 1
    End If
  End If
End Function

Public Function okonver$(k$)
  '--- string konvertálása Windows -> 852
  z$ = k$
  q2$ = " " + Chr$(160) + Chr$(130) + Chr$(161) + Chr$(162) + Chr$(148) + Chr$(139) + Chr$(163) + Chr$(129) + Chr$(251)
  q2$ = q2$ + Chr$(181) + Chr$(144) + Chr$(214) + Chr$(224) + Chr$(153) + Chr$(138) + Chr$(233) + Chr$(154) + Chr$(235)
  q1$ = " áéíóöõúüûÁÉÍÓÖÕÚÜÛ"
  i1% = Len(k$)
  For i2% = 1 To i1%
    For i3% = 1 To Len(q1$)
      If Mid$(z$, i2%, 1) = Mid$(q1$, i3%, 1) Then
        Mid$(z$, i2%, 1) = Mid$(q2$, i3%, 1): Exit For
      End If
    Next
  Next
  okonver$ = z$
End Function

Public Sub titkosit(inpfil$, outfil$, timod$, tablafil$)
  '--- fájl tartalmának titkosítása és dekódolása
  '--- timod$=T titkosít
  '--- timod$=D dekódol
  '--- tablafil$ a kódtáblaként használandó fájl neve
  Dim kodtabla$(10)
  tofil = FreeFile
  Open outfil$ For Output As #tofil
  Close tofil
  If timod$ = "T" Then
    '--- fájl tikosítása
    kts& = Int(90 * Rnd + 1)
    tafil = FreeFile
    Open "c:\auwin\" + tablafil$ For Binary Shared As #tafil
    For ti& = 1 To 10
      kcim& = (kts& + ti& - 2) * 256 + 1
      kodtab$ = Space(256)
      Get #tafil, 1, kodtab$
      kodtabla$(ti&) = kodtab$
    Next
    Close tafil
    tifil = FreeFile
    Open inpfil$ For Binary Shared As #tifil
    tic& = LOF(tifil)
    tofil = FreeFile
    Open outfil$ For Binary As #tofil
    toc& = 1
    kgst% = "19511"
    Put #tofil, toc&, kgst%
    toc& = 3
    ktxx% = kts&
    Put #tofil, toc&, ktss%
    toc& = toc& + 2
    For ti& = 1 To tic&
      x$ = " "
      Get #tifil, ti&, x$
      tindex% = ti& Mod 10 + 1
      poz% = InStr(kodtabla$(tindex%), x$)
      ocar$ = Chr$(poz% - 1)
      Put #tofil, toc&, ocar$
      toc& = toc& + 1
    Next
    Close tifil: Close tofil
  Else
    '--- dekódolás
    tifil = FreeFile
    Open inpfil$ For Binary Shared As #tifil
    tic& = LOF(tifil)
    Get #tifil, 3, ktss%
    kts& = ktss%
    tofil = FreeFile
    Open outfil$ For Binary As #tofil
    toc& = 1
    tafil = FreeFile
    Open "c:\auwin\" + tablafil$ For Binary Shared As #tafil
    For ti& = 1 To 10
      kcim& = (kts& + ti& - 2) * 256 + 1
      kodtab$ = Space(256)
      Get #tafil, 1, kodtab$
      kodtabla$(ti&) = kodtab$
    Next
    Close tafil
    For ti& = 1 To tic& - 4
      x$ = " "
      Get #tifil, ti& + 4, x$
      tindex% = ti& Mod 10 + 1
      poz% = Asc(x$) + 1
      ocar$ = Mid$(kodtabla$(tindex%), poz%, 1)
      Put #tofil, toc&, ocar$
      toc& = toc& + 1
    Next
    Close tifil: Close tofil
  End If
End Sub

Public Function kodinverz$(sr$)
  '--- string invertálás 256 ra
  lc& = Len(sr$)
  koc$ = ""
  For li1% = 1 To sr$
    xa% = Asc(Mid(sr$, li1%, 1))
    koc$ = koc$ + Chr$(255 - xa%)
  Next
  kodinverz$ = koc$
End Function

Public Function jodatum%(dat$)
  If Len(dat$) = 6 Then
    ee% = xval(Mid$(dat$, 1, 2))
    hh% = xval(Mid$(dat$, 3, 2))
    nn% = xval(Mid$(dat$, 5, 2))
    If ee% > 0 And ee% <= 39 Then
      If hh% > 0 And hh% < 13 Then
        Select Case hh%
          Case 1, 3, 5, 7, 8, 10, 12
            If nn% >= 1 And nn% <= 31 Then jodatum% = 1: Exit Function
          Case 4, 6, 9, 11
            If nn% >= 1 And nn% <= 30 Then jodatum% = 1: Exit Function
          Case 2
            If ee% Mod 4 = 0 Then vv% = 29 Else vv% = 28
            If nn% >= 1 And nn% <= vv% Then jodatum% = 1: Exit Function
          Case Else
        End Select
      End If
    End If
  End If
  jodatum% = 0
End Function

Public Function xleft$(mez$, hossz%)
  xleft$ = Left(Trim(mez$) + Space(hossz%), hossz%)
End Function

Public Function xright$(mez$, hossz%)
  xright$ = Right$(Space(hossz%) + Trim(mez$), hossz%)
End Function

Public Sub afakulcstolt(irec$)
  afakulcsokdb = Val(Mid$(irec$, 720, 1))
  Select Case afakulcsokdb
    Case 0
      afakulcsok(1) = 5: afakulcsok(2) = 15: afakulcsok(3) = 25
      afakulcsokdb = 3
    Case 1
      afakulcsok(1) = xval(Mid(irec$, 721, 6))
    Case 2
      afakulcsok(1) = xval(Mid(irec$, 721, 6))
      afakulcsok(2) = xval(Mid(irec$, 727, 6))
    Case 3
      afakulcsok(1) = xval(Mid(irec$, 721, 6))
      afakulcsok(2) = xval(Mid(irec$, 727, 6))
      afakulcsok(3) = xval(Mid(irec$, 733, 6))
    Case Is > 3
      afakulcsok(1) = xval(Mid(irec$, 721, 6))
      afakulcsok(2) = xval(Mid(irec$, 727, 6))
      afakulcsok(3) = xval(Mid(irec$, 733, 6))
      afakulcsok(4) = xval(Mid(irec$, 739, 6))
    Case Else
  End Select
End Sub

Public Function datfor(dat$)
  datfor = Right(dat$, 2) + Mid(dat$, 3, 2) + Left(dat$, 2)
End Function

Public Function dat6(dat$)
  If langhun% > 1 Then
    dat6 = Right(dat$, 2) + Mid(dat$, 3, 2) + Left(dat$, 2)
  Else
    dat6 = dat$
  End If
End Function

Public Function banktagol$(bsz$)
  Select Case langhun%
    Case 1
      bxsz$ = Trim(bsz$)
      If bxsz$ = "" Then
        banktagol$ = ""
      Else
        If Len(bxsz$) > 8 Then
          If Len(bxsz$) > 16 Then
            banktagol$ = Left(bxsz$, 8) + "-" + Mid$(bxsz$, 9, 8) + "-" + Mid$(bxsz$, 17)
          Else
            banktagol$ = Left(bxsz$, 8) + "-" + Mid$(bxsz$, 9)
          End If
        Else
          banktagol$ = bxsz$
        End If
      End If
    Case 2
    Case 3
    Case 4
      bxsz$ = Trim(bsz$)
      If bxsz$ = "" Then
        banktagol$ = ""
      Else
        If Len(bxsz$) > 3 Then
          If Len(bxsz$) > 16 Then
            banktagol$ = Left(bxsz$, 3) + "-" + Mid$(bxsz$, 4, 13) + "-" + Mid$(bxsz$, 17)
          Else
            banktagol$ = Left(bxsz$, 3) + "-" + Mid$(bxsz$, 4)
          End If
        Else
          banktagol$ = bxsz$
        End If
      End If
    Case Else
      bxsz$ = Trim(bsz$)
      If bxsz$ = "" Then
        banktagol$ = ""
      Else
        If Len(bxsz$) > 8 Then
          If Len(bxsz$) > 16 Then
            banktagol$ = Left(bxsz$, 8) + "-" + Mid$(bxsz$, 9, 8) + "-" + Mid$(bxsz$, 17)
          Else
            banktagol$ = Left(bxsz$, 8) + "-" + Mid$(bxsz$, 9)
          End If
        Else
          banktagol$ = bxsz$
        End If
      End If
  End Select
End Function

Public Function binker&(tomb$(), tombelemszam&, azonosito$, tol%, ig%)
  '--- bináris keresés tömbben
  also& = 1: felso& = tombelemszam&
  If also& <= felso& Then
    bj$ = Mid$(tomb$(felso&), tol%, ig%)
    If azonosito$ = bj$ Then binker& = felso&: Exit Function
    bj$ = Mid$(tomb(also&), tol%, ig%)
    If azonosito$ = bj$ Then binker& = also&: Exit Function
    Do
      kozep& = Int((felso& + also&) / 2)
      bj$ = Mid$(tomb(kozep&), tol%, ig%)
      If azonosito$ = bj$ Then
        binker& = kozep&: Exit Function
      Else
        If azonosito$ < bj$ Then
          felso& = kozep&
        Else
          also& = kozep&
        End If
      End If
    Loop While also& + 1 < felso&
  End If
  binker& = 0
End Function

Public Function binolv$(tomb$(), tombelemszam&, azonosito$, tol%, ig%)
  wyx& = binker(tomb(), tombelemszam&, azonosito$, tol%, ig%)
  If wyx& > 0 Then binolv = tomb$(wyx&) Else binolv = ""
End Function

Public Sub binind(tomb$(), tombelemszam&, azonosito$, tol%, ig%, wyxrec$, wyx&)
  wyx& = binker(tomb(), tombelemszam&, azonosito$, tol%, ig%)
  If wyx& > 0 Then wyxrec$ = tomb$(wyx&) Else wyxrec$ = ""
End Sub

Public Sub mess(szoveg$, kod%, kod2%, fejlec$, valasz%)
  '--- msgbox helyettesítõ
  '--- 1 Kritikus hiba   (ok gomb)
  '--- 2 Hiba            (ok gomb)
  '--- 3 Figyelmeztetés  (ok gomb)
  '--- 4 Információ      (ok gomb)
  '--- 5 Kérdés          (igen, mégsem gomb)
  '--- 6 Siker           (ok gomb)
  If InStr(szoveg$, Chr$(13)) = 0 Then
    Mesbox.Height = 1682
    Mesbox.Label1.Height = 432
    If Len(szoveg$) > 80 Then
      Mesbox.Width = 9000
      Mesbox.Label1.Width = 8000
      Mesbox.Command1.Left = 2780: Mesbox.Command1.Top = 720
      Mesbox.Command2.Left = 4460: Mesbox.Command2.Top = 720
      Mesbox.Command3.Left = 3500: Mesbox.Command3.Top = 720
    Else
      Mesbox.Width = 5196
      Mesbox.Label1.Width = 4296
      Mesbox.Command1.Left = 840: Mesbox.Command1.Top = 720
      Mesbox.Command2.Left = 2520: Mesbox.Command2.Top = 720
      Mesbox.Command3.Left = 1680: Mesbox.Command3.Top = 720
    End If
  Else
    okdbx% = 0
    For j88% = 1 To Len(szoveg$)
      If Mid$(szoveg$, j88%, 1) = Chr$(13) Then okdbx% = okdbx% + 1
    Next
    Mesbox.Width = 9000: Mesbox.Label1.Width = 8000
    If okdbx% > 2 Then
      Mesbox.Height = 2700
      Mesbox.Label1.Height = 1450
      Mesbox.Command1.Left = 2780: Mesbox.Command1.Top = 1740
      Mesbox.Command2.Left = 4460: Mesbox.Command2.Top = 1740
      Mesbox.Command3.Left = 3500: Mesbox.Command3.Top = 1740
    Else
      Mesbox.Height = 1900
      Mesbox.Label1.Height = 800
      Mesbox.Command1.Left = 2780: Mesbox.Command1.Top = 1000
      Mesbox.Command2.Left = 4460: Mesbox.Command2.Top = 1000
      Mesbox.Command3.Left = 3500: Mesbox.Command3.Top = 1000
    End If
  End If
  Mesbox.Caption = fejlec$
  Mesbox.Label1.Caption = szoveg$
  Select Case kod%
    Case 1
      '--- kritikus hiba
      Mesbox.Image1.Visible = True: Mesbox.Image1.Left = 170: Mesbox.Image1.Top = 130
      Mesbox.Command3.Visible = True
    Case 2
      '--- hiba
      Mesbox.Image6.Visible = True: Mesbox.Image6.Left = 120: Mesbox.Image6.Top = 100
      Mesbox.Command3.Visible = True
    Case 3
      '--- figyelmeztetés
      Mesbox.Image2.Visible = True: Mesbox.Image2.Left = 170: Mesbox.Image2.Top = 130
      Mesbox.Command3.Visible = True
    Case 4
      '--- információ
      Mesbox.Image4.Visible = True: Mesbox.Image4.Left = 170: Mesbox.Image4.Top = 170
      Mesbox.Command3.Visible = True
    Case 5
      '--- kérdés
      Mesbox.Image3.Visible = True: Mesbox.Image3.Left = 170: Mesbox.Image3.Top = 170
      Mesbox.Command1.Visible = True: Mesbox.Command2.Visible = True
      If kod2% <> 0 Then
        Mesbox.Image3.Top = 600
        Mesbox.Image3.Visible = True
        Select Case kod2%
          Case 1: Mesbox.Image1.Visible = True: Mesbox.Image1.Left = 170: Mesbox.Image1.Top = 130
          Case 2: Mesbox.Image6.Visible = True: Mesbox.Image6.Left = 130: Mesbox.Image6.Top = 100
          Case 3: Mesbox.Image2.Visible = True: Mesbox.Image2.Left = 170: Mesbox.Image2.Top = 130
          Case 4: Mesbox.Image4.Visible = True: Mesbox.Image4.Left = 170: Mesbox.Image4.Top = 130
          Case Else
        End Select
      End If
    Case 6
      '--- siker üzenet
      Mesbox.Image5.Visible = True: Mesbox.Image5.Left = 170: Mesbox.Image5.Top = 150
      Mesbox.Command3.Visible = True
    Case Else
  End Select
  Mesbox.Show vbModal
  valasz% = msgvalasz
End Sub
Public Function ujobjektum(adatbaziskod$, objektumkod$)
  '--- új objektum beléptetése
  rrepar$ = adatbaziskod$ + "/" + objektumkod$ + "/" + terminal$ + task$ + "/" + auditorutvonal$
  r = Shell(programutvonal$ + "dbx4-new.exe " + rrepar$, vbNormalFocus)
End Function
Public Function expdatki(dat$)
  If Trim(dat$) = "" Then expdatki = Space(10): Exit Function
  If Mid$(dat$, 1, 2) > "39" Then evv$ = "19" Else evv$ = "20"
  expdatki = Mid$(dat$, 5, 2) + "." + Mid$(dat$, 3, 2) + "." + evv$ + Mid$(dat$, 1, 2)
End Function

Public Function partnercim$(partrec$)
  yx$ = Trim(Mid$(partrec$, 106, 8)) + " " + Trim(Mid$(partrec$, 114, 30)) + ", " + Trim(Mid$(partrec$, 144, 30)) + " " + Trim(Mid$(partrec$, 174, 10))
  orszag$ = orszagkod$(Mid$(partrec$, 780, 2))
  If Trim(orszag$) = "" Then orszag$ = Trim(Mid$(partrec$, 76, 30))
  partnercim$ = yx$ + " " + orszag$
End Function

Public Function orszagkod$(oko$)
  Dim ogk$(30, 2)
  ogk$(1, 1) = "HU": ogk$(1, 2) = "Magyarország"
  ogk$(2, 1) = "AT": ogk$(2, 2) = "Ausztria"
  ogk$(3, 1) = "BG": ogk$(3, 2) = "Belgium"
  ogk$(4, 1) = "DK": ogk$(4, 2) = "Dánia"
  ogk$(5, 1) = "GB": ogk$(5, 2) = "Nagy-Britannia"
  ogk$(6, 1) = "FI": ogk$(6, 2) = "Finnország"
  ogk$(7, 1) = "FR": ogk$(7, 2) = "Franciaország"
  ogk$(8, 1) = "DE": ogk$(8, 2) = "Németország"
  ogk$(9, 1) = "MT": ogk$(9, 2) = "Málta"
  ogk$(10, 1) = "SK": ogk$(10, 2) = "Szlovákia"
  ogk$(11, 1) = "SI": ogk$(11, 2) = "Szlovénia"
  ogk$(12, 1) = "NL": ogk$(12, 2) = "Hollandia"
  ogk$(13, 1) = "ES": ogk$(13, 2) = "Spanyolország"
  ogk$(14, 1) = "CY": ogk$(14, 2) = "Ciprus"
  ogk$(15, 1) = "CZ": ogk$(15, 2) = "Csehország"
  ogk$(16, 1) = "EE": ogk$(16, 2) = "Észtország"
  ogk$(17, 1) = "PL": ogk$(17, 2) = "Lengyelország"
  ogk$(18, 1) = "LT": ogk$(18, 2) = "Lettország"
  ogk$(19, 1) = "LV": ogk$(19, 2) = "Litvánia"
  ogk$(20, 1) = "EL": ogk$(20, 2) = "Görögország"
  ogk$(21, 1) = "IE": ogk$(21, 2) = "Írország"
  ogk$(22, 1) = "IT": ogk$(22, 2) = "Olaszország"
  ogk$(23, 1) = "LU": ogk$(23, 2) = "Luxemburg"
  ogk$(24, 1) = "PT": ogk$(24, 2) = "Portugália"
  ogk$(25, 1) = "SE": ogk$(25, 2) = "Svédország"
  For i1i% = 1 To 25
    If ogk$(i1i%, 1) = oko$ Then orszagkod$ = ogk$(i1i%, 2): Exit Function
  Next
  orszagkod$ = ""
End Function

Public Sub torzsrbe(objazon$, ttm$(), ttc&(), ttmp&, para%(), paradb&)
  '--- törzsállomány részleges beolvasása tömbbe
  ttmp& = 0
  ob% = obsorszama(objazon)
  dbxn$ = dbxneve(objazon)
  w1% = OBJTAB(ob%).obi(1)
  indn$ = RTrim$(INDTAB(w1%).indnev)
  ih& = ADATAB(INDTAB(w1%).adatsorsz).adatho + 5
  rh& = OBJTAB(ob%).rekhossz
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  rc& = Int(LOF(ndfi) / ih&)
  If rc& = 0 Then Exit Sub
  For isa% = 1 To paradb&: para%(isa%, 0) = 0: Next
  For iti& = 1 To rc&
    DoEvents
    Get #ndfi, (iti& - 1) * ih& + 1, cci&
    r$ = Space(rh&)
    Get #dbfi, cci& + 9, r$
    s$ = ""
    For isa% = 1 To paradb&
      s1$ = Mid$(r$, para%(isa%, 1), para%(isa%, 2))
      s1l% = Len(Trim(s1$)): If s1l% > para%(isa%, 0) Then para%(isa%, 0) = s1l%
      s$ = s$ + s1$
    Next
    ttmp& = ttmp& + 1
    ReDim Preserve ttm$(1 To ttmp&)
    ReDim Preserve ttc&(1 To ttmp&)
    ttm$(ttmp&) = s$
    ttc&(ttmp&) = cci&
  Next
  Close dbfi: Close ndfi
End Sub
Public Sub torzsbe(objazon$, ttm$(), ttmp&)
  '--- torzsállomány beolvasása tömbbe
  ttmp& = 0
  ob% = obsorszama(objazon)
  dbxn$ = dbxneve(objazon)
  w1% = OBJTAB(ob%).obi(1)
  indn$ = RTrim$(INDTAB(w1%).indnev)
  ih& = ADATAB(INDTAB(w1%).adatsorsz).adatho + 5
  rh& = OBJTAB(ob%).rekhossz
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  rc& = Int(LOF(ndfi) / ih&)
  If rc& = 0 Then Exit Sub
  For iti& = 1 To rc&
    DoEvents
    Get #ndfi, (iti& - 1) * ih& + 1, cci&
    r$ = Space(rh&)
    Get #dbfi, cci& + 9, r$
    ttmp& = ttmp& + 1
    ReDim Preserve ttm$(1 To ttmp&)
    ttm$(ttmp&) = r$
  Next
  Close dbfi: Close ndfi
End Sub

Public Sub torzsinput(objazon$, ttm$(), ttc&(), ttmp&, rrekod$, rskp%, rsho%)
  Call torzsbe(objazon$, ttm$(), ttmp&)
  If ttmp& > 0 Then
    If Trim(rrekod$) <> "" Then
      ReDim ttc(1 To ttmp)
      Call qsortr(ttm$(), ttc&(), ttmp&, rrekod$, rskp%, rsho%)
    End If
  End If
End Sub

Public Function torzskeres&(ttm$(), ttmp&, azonosito$, hossz%)
  '--- soros keresés tömbben
  torzskeres& = 0
  If ttmp& = 0 Then Exit Function
  For iti& = 1 To ttmp&
    If Mid$(ttm$(iti&), 1, hossz%) = azonosito$ Then torzskeres& = iti&: Exit Function
  Next
End Function

Public Sub grafik(szoveg$, grt@(), vonaldarab%, vonalszinek&(), vonalnevek$(), xmegys$, ymegys$, xelem%, xskala$, xkezdet$)
  '--- xskala$ N-nap H-hét E-érték  xketdet$=kezdõ érték
  Grafikon.Label1.Caption = szoveg$
  '--- címkéket kirakni
  lsta% = 700
  For i88% = 0 To 5
    If i88% + 1 <= vonaldarab% Then
      Grafikon.Line3(i88).Visible = True
      Grafikon.Line3(i88).X1 = lsta%
      Grafikon.Line3(i88).X2 = lsta% + 400: lsta% = lsta% + 500
      Grafikon.Line3(i88).BorderColor = vonalszinek&(i88% + 1)
      Grafikon.Line3(i88).BorderWidth = 2
      Grafikon.Label2(i88).Visible = True
      Grafikon.Label2(i88).Left = lsta%
      Grafikon.Label2(i88).Caption = Trim(vonalnevek$(i88 + 1))
      lw% = Grafikon.TextWidth(Trim(vonalnevek$(i88 + 1))) + 50
      Grafikon.Label2(i88).Width = lw%
      lsta% = lsta% + lw% + 50
    Else
      Grafikon.Line3(i88).Visible = False
      Grafikon.Label2(i88).Visible = False
    End If
  Next
  ymax# = 0
  For i99% = 1 To vonaldarab%
    For i88% = 1 To xelem%
      If grt@(i88%, i99%) > ymax# Then ymax# = grt@(i88%, i99%)
    Next
  Next
  If ymax# <> 0 Then yegys# = 5000 / ymax# Else yegys# = 0
  xegys& = Int(9000 / xelem%)
  '--- függõleges léc kalibrálása
  xz1@ = Int(ymax# / 10)
  Grafikon.Line1.Y1 = 6300 - yegys# * ymax#
  Grafikon.Line2.X2 = 1000 + xegys& * xelem%
  yynd& = 0
  For i88% = 1 To 10
    yynd& = xz1@ * i88 * yegys#
    Grafikon.Line4(i88% - 1).Y1 = 6300 - yynd&
    Grafikon.Line4(i88% - 1).Y2 = 6300 - yynd&
    Grafikon.Line4(i88% - 1).X2 = 1000 + xegys& * xelem%
    Grafikon.Line4(i88% - 1).BorderColor = RGB(180, 180, 180)
    Grafikon.Label3(i88% - 1).Top = 6200 - yynd&
    Grafikon.Label3(i88% - 1).Caption = Str(Int(xz1@ * i88%)) + " " + ymegys$
  Next
  Grafikon.Line4(10).Y1 = 6300 - ymax# * yegys#
  Grafikon.Line4(10).Y2 = 6300 - ymax# * yegys#
  Grafikon.Line4(10).X2 = 1000 + xegys& * xelem%
  Grafikon.Line4(10).BorderColor = RGB(180, 180, 180)
  Grafikon.Label3(10).Top = 6200 - ymax# * yegys#
  Grafikon.Label3(10).Caption = Str(Int(ymax#)) + " " + ymegys$
  '--- vizszintes léc kalibrálása
  Grafikon.MSFlexGrid1.Clear
  Grafikon.MSFlexGrid1.Top = 6300 - ymax# * yegys#
  Grafikon.MSFlexGrid1.Height = ymax# * yegys#
  Grafikon.MSFlexGrid1.Left = Grafikon.Line2.X2 + 50
  Grafikon.MSFlexGrid1.Width = 11100 - xegys& * xelem%
  Select Case xskala$
    Case "N"
      Grafikon.MSFlexGrid1.Rows = xelem% + 1
      Grafikon.MSFlexGrid1.Cols = vonaldarab% + 1
      Grafikon.MSFlexGrid1.ColWidth(0) = 800
      xco& = 0: iindex% = 0
      dd$ = xkezdet$
      For i89% = 1 To vonaldarab%
        Grafikon.MSFlexGrid1.TextMatrix(0, i89%) = vonalnevek$(i89%)
      Next
      For i88% = 1 To xelem%
        dd$ = novdat(dd$)
        Grafikon.MSFlexGrid1.TextMatrix(i88, 0) = datki(dd$)
        For i89% = 1 To vonaldarab%
          Grafikon.MSFlexGrid1.Row = i88%: Grafikon.MSFlexGrid1.Col = i89%
          Grafikon.MSFlexGrid1.CellForeColor = vonalszinek(i89%)
          Grafikon.MSFlexGrid1.TextMatrix(i88, i89%) = grt@(i88%, i89%)
        Next
        If Mid$(dd$, 5, 2) = "01" Or i88% = xelem% Or i88% = 1 Then
          Grafikon.Line5(iindex).X1 = 1000 + xco&
          Grafikon.Line5(iindex).X2 = 1000 + xco&
          Grafikon.Line5(iindex).Y1 = 6300 - ymax# * yegys#
          Grafikon.Line5(iindex).BorderColor = RGB(180, 180, 180)
          Grafikon.Label5(iindex).Left = 800 + xco&
          If iindex Mod 2 = 0 Then
            Grafikon.Line5(iindex).Y2 = 6350
            Grafikon.Label5(iindex).Top = 6400
          Else
            Grafikon.Line5(iindex).Y2 = 6550
            Grafikon.Label5(iindex).Top = 6600
          End If
          Grafikon.Label5(iindex).Caption = datki(dd$)
          iindex% = iindex% + 1
        End If
        xco& = xco& + xegys&
      Next
      'Grafikon.Label5(0).Visible = False: Grafikon.Line5(0).Visible = False
      For i88% = iindex% To 25
        Grafikon.Label5(i88%).Visible = False: Grafikon.Line5(i88%).Visible = False
      Next
    Case "H"
    Case "E"
    Case Else
  End Select
  'Grafikon.Label3.Caption = Trim(Str(Int(ymax#))) + " " + ymegys$
  Grafikon.Label4.Caption = Trim(Str(Int(xelem%))) + " " + xmegys$
  Grafikon.Show
  For i99% = 1 To vonaldarab%
    elox% = 1000: eloy% = 6290
    For i88% = 1 To xelem%
      If i88% = xelem% Then
        vegx% = elox% + xegys&
        'vegx% = xelem% * xegys&
      Else
        vegx% = elox% + xegys&
      End If
      If grt@(i88%, i99%) = ymax# Then
        vegy% = 6290 - Int(yegys# * grt@(i88%, i99%))
      Else
        vegy% = 6290 - Int(yegys# * grt@(i88%, i99%))
      End If
      Grafikon.DrawWidth = 2
      Grafikon.Line (elox%, eloy%)-(vegx%, vegy%), vonalszinek&(i99%)
      elox% = vegx%: eloy% = vegy%
    Next
  Next
End Sub

Public Sub waitforfile(utvo$, fnev$)
  Do
    DoEvents
    s$ = UCase(Dir(utvo$ + fnev$))
  Loop While s$ <> UCase(fnev$)
  sm1& = FileLen(utvo$ + fnev$)
  Do
    DoEvents
    If UCase(utvo$) = "C:\ECOSERV\" Then
      Call waitsec(10)
    Else
      Call waitsec(2)
    End If
    sm& = FileLen(utvo$ + fnev$)
    If sm1& = sm& Then Exit Do
    sm1& = sm&
  Loop While sm& > 0
End Sub
Public Sub xmerge(infil1$, infil2$, outfil$, rho%, kulcsk%, kulcsh%)
  '--- két fájl mergelése
  Xmerg.Label1.Caption = infil1$
  Xmerg.Label2.Caption = infil2$
  Xmerg.ProgressBar1.Min = 0
  Xmerg.ProgressBar1.Max = 100
  Xmerg.ProgressBar2.Min = 0
  Xmerg.ProgressBar2.Max = 100
  Xmerg.Show
  mfi1 = FreeFile
  Open infil1 For Binary As mfi1
  mfi2 = FreeFile
  Open infil2 For Binary As mfi2
  mfi3 = FreeFile
  Open outfil$ For Output As #mfi3
  rc1& = Int(LOF(mfi1) / 55)
  rc2& = Int(LOF(mfi2) / 55)
  If rc1& = 0 Then
    kulcs1$ = String(kulcsh%, Chr(255))
  Else
    poz1& = 1
    r1$ = Space(53)
    Get #mfi1, (poz1& - 1) * 55 + 1, r1$
    kulcs1$ = Mid$(r1$, kulcsk%, kulcsh%)
  End If
  If rc2& = 0 Then
    kulcs2$ = String(kulcsh%, Chr(255))
  Else
    poz2& = 1
    r2$ = Space(53)
    Get #mfi2, (poz2& - 1) * 55 + 1, r2$
    kulcs2$ = Mid$(r2$, kulcsk%, kulcsh%)
  End If
  Do
    DoEvents
    If kulcs1$ = String(kulcsh%, Chr(255)) And kulcs2$ = String(kulcsh%, Chr(255)) Then Exit Do
    If kulcs1$ <= kulcs2$ Then
      '--- 1.ir és olvas
      Print #mfi3, r1$
      poz1& = poz1& + 1
      If poz1& > rc1& Then
        kulcs1$ = String(kulcsh%, Chr(255))
      Else
        Xmerg.ProgressBar1.Value = pscale(poz1&, rc1&)
        r1$ = Space(53)
        Get #mfi1, (poz1& - 1) * 55 + 1, r1$
        kulcs1$ = Mid$(r1$, kulcsk%, kulcsh%)
      End If
    Else
      '--- 2.ir és olvas
      Print #mfi3, r2$
      poz2& = poz2& + 1
      If poz2& > rc2& Then
        kulcs2$ = String(kulcsh%, Chr(255))
      Else
        Xmerg.ProgressBar2.Value = pscale(poz2&, rc2&)
        r2$ = Space(53)
        Get #mfi2, (poz2& - 1) * 55 + 1, r2$
        kulcs2$ = Mid$(r2$, kulcsk%, kulcsh%)
      End If
    End If
  Loop While poz1& <> -1
  Close mfi1: Close mfi2: Close mfi3
  Xmerg.Hide
End Sub

Public Function szerkod$(szlasz$, erak$)
  If Len(szlasz$) >= 3 Then
    szlakezd$ = UCase(Left(szlasz$, 3))
  Else
    szerkod = "": Exit Function
  End If
  Select Case szlakezd
    Case "MER"
      Select Case erak$
        Case "0002": szerkod = "0002    "
        Case "0004": szerkod = "0004    "
        Case Else: szerkod = "0001    "
      End Select
    Case "ADR": szerkod = "1004    "
    Case "KOM": szerkod = "1002    "
    Case "CSE": szerkod = "1003    "
    Case "MOH": szerkod = "1001    "
    Case "ZEG": szerkod = "1005    "
    Case "SZV": szerkod = "1006    "
    Case "NAG": szerkod = "1007    "
    Case "TEJ": szerkod = "0005    "
    Case "5KE": szerkod = "0001    "
    Case "ORI": szerkod = "2001    "
    Case "ROK": szerkod = "2002    "
    Case "UJH": szerkod = "2003    "
    Case "SAS": szerkod = "2004    "
    Case "HID": szerkod = "2006    "
    Case "SZI": szerkod = "2005    "
    Case Else: szerkod = ""
  End Select
End Function

Public Sub v1ablak(objazon, ablaksz%, rec$, ttop&, lleft&, sszel&, mmag&)
  '--- vektoros adatok bevitele
  '--- ablakszám az objektumon belül ablakszám=0 esetén teljes objektum
  '--- rec$-ban a rekord
  '--- tablazat mérete és poziciója
  '--- ttop,lleft,mmag,sszel twip-ben
  charpertwipa% = 120
  charpertwip% = 120
  form1.MSFlexGrid1.Clear
  form1.MSFlexGrid1.Cols = 2
  w1& = obsorszama(objazon)
  w2& = OBJTAB(w1&).oba(ablaksz%)
  wodarab& = ABLTAB(w2&).adatsorsz(0)
  foindex% = INDTAB(OBJTAB(w1&).obi(1)).adatsorsz
  form1.MSFlexGrid1.TextMatrix(0, 0) = langmodul(75)
  form1.MSFlexGrid1.TextMatrix(0, 1) = langmodul(76)
  mmax% = 12: hxmax% = 0
  mhmax% = 10
  odarab& = 0
  'form1.Font.Name = "Microsoft Sans Serif"
  'form1.Font.Size = 8
  form1.MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  For i1& = 1 To wodarab&
    ne$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatnev)
    ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).attr)
    If InStr(ar$, "R") = 0 Then
      odarab& = odarab& + 1
      form1.MSFlexGrid1.Rows = odarab& + 2
      ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).attr)
      mh% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatho
      kp% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatkp
      form1.MSFlexGrid1.TextMatrix(odarab&, 0) = ne$
      w3& = Len(ne$)
      hxh% = form1.TextWidth(ne$) + 100
      If hxmax% < hxh% Then hxmax% = hxh%
      form1.MSFlexGrid1.ColWidth(0) = hxmax%
      If mh% > mhmax% Then h1% = mh% * charpertwip%: mhmax% = mh% Else h1% = mhmax% * charpertwip%
      form1.MSFlexGrid1.ColWidth(1) = h1%
      form1.MSFlexGrid1.Width = hxmax% + mhmax% * charpertwip% + 70
      'form1.Width = form1.MSFlexGrid1.Width + 300
      'VA1.Caption = ABLTAB(w2&).fejlec
      form1.MSFlexGrid1.ColSel = 1
      form1.MSFlexGrid1.RowSel = odarab&
      If InStr(ar$, "J") > 0 And InStr(ar$, "NZJ") = 0 Then
        form1.MSFlexGrid1.CellAlignment = flexAlignLeftCenter
      Else
        form1.MSFlexGrid1.CellAlignment = flexAlignLeftCenter
      End If
      amezo$ = Trim$(Mid$(rec$, kp%, mh%))
      form1.MSFlexGrid1.TextMatrix(odarab&, 1) = amezo$
      mtb$(odarab&) = ar$: mho%(odarab&) = mh%: mkp(odarab&) = kp%
      mesor(odarab&) = ABLTAB(w2&).adatsorsz(i1&)
    End If
  Next
  If sszel& <> 0 Then form1.MSFlexGrid1.Width = sszel& - 300
  If ttop& <> 0 Then
    form1.MSFlexGrid1.Top = ttop&
  Else
    form1.MSFlexGrid1.Top = 1500
  End If
  If lleft& <> 0 Then
    form1.MSFlexGrid1.Left = lleft&
  Else
    form1.MSFlexGrid1.Left = 100
  End If
  'If VA1.Left + VA1.Width > form1.Width Then VA1.Left = form1.Width - VA1.Width
  If mmag& <> 0 Then mag& = mmag& Else mag& = odarab& * 240 + 2200
  tetf& = form1.MSFlexGrid1.Top
  If tetf& + mag& > 8700 Then mag& = 8700 - tetf&
  form1.MSFlexGrid1.Height = mag& - 1870
  form1.MSFlexGrid1.Visible = True
End Sub

Public Function merkaraktar$(kod$)
  '--- Merkatimpex, 5-ker raktár kódból név
  Select Case kod$             '12345678901234567890"
    Case "0001": merkaraktar = "Merkatimpex központ "
    Case "1001": merkaraktar = "Mohácsi lerakat     "
    Case "1002": merkaraktar = "Komlói lerakat      "
    Case "1003": merkaraktar = "Csernecky italdepo  "
    Case "1004": merkaraktar = "Adrienn C+C         "
    Case "1005": merkaraktar = "Zalaegerszegi depo  "
    Case "1006": merkaraktar = "Szigetvári depo     "
    Case "2001": merkaraktar = "Óriás diszkont      "
    Case "2002": merkaraktar = "Rókus diszkont      "
    Case "2003": merkaraktar = "Újhegyi CBA         "
    Case "2004": merkaraktar = "Sásdi CBA           "
    Case "2005": merkaraktar = "Szigetvári tanbolt  "
    Case "2006": merkaraktar = "Hidasi CBA          "
    Case "2007": merkaraktar = "Pécsváradi CBA      "
    Case "2008": merkaraktar = "Villányi CBA        "
    Case Else: merkaraktar = Space(20)
  End Select
End Function


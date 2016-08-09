Attribute VB_Name = "WRTBIZ"
'--- NAV-os számlázá struktúrái és rutinjai
Type wcimtyp
  wpkod As String
  Wpnev As String
  wirszam As String
  wtelep As String
  wker As String
  koztern As String
  kozterj As String
  hazsz As String
  epul As String
  lepcs As String
  szint As String
  ajto As String
  wradosz As String
  wrcsopado As String
  wreuadosz As String
  wcim As String
  wcic1 As String
  wcic2 As String
End Type
Public wxi As wcimtyp
Public wxp As wcimtyp
Public wxk As wcimtyp
Public wrtbsel$(), wrtbcimm&(), wrtbdb&
Public wrtbz$, wrttet$(), wrtetdb%
Public wszamlatipus$, wszamlafajta$, wrnsrec$, tetnevmt$(200), wbeszeloikt$(200)
Public wrtszamlacim$, wrtszamlaszam$, wrterszamlaszam$, wrtsztornoszam$, wrtuzemmod$, welolegszamla%, welolegkontir$
Public utftab%(30, 3), utfdb%, kodasc$, kodutf$
Public xmlfil, navsor$, wbadoszam8$, wexportszamla%

Public Sub wrtinit()
  wrtbz$ = Space(3000)
  ReDim wrttet(1 To 1)
  wrtetdb = 0
End Sub

Public Sub wrtbir(wrbkp%, wrbho%, wmezo$)
  '--- karakteres mezõ írása wrtbz-be (WRTB)
  If Trim(wmezo$) <> "" Then
    If wrbkp% > 0 And wrbho% > 0 Then
      Mid$(wrtbz, wrbkp%, wrbho%) = wmezo$
    End If
  End If
End Sub

Public Sub wrtnum(wrbkp%, wrbho%, wernum@)
  '--- 2 tizedes numerikus mezõ írása wrtbz-be (WRTB)
  If wernum@ <> 0 Then
    Mid$(wrtbz, wrbkp%, wrbho%) = Right(Space(wrbho%) + ertszam(Str(wernum@), 14, 2), wrbho%)
  Else
    Mid$(wrtbz, wrbkp%, wrbho%) = Space(wrbho%)
  End If
End Sub

Public Sub wrttir(wujsor%, wrtkp%, wrtho%, wmezo$)
  '--- karakteres mezõ írása wrtt-be
  If wujsor% = 1 Then
    wrtetdb = wrtetdb + 1
    ReDim Preserve wrttet(1 To wrtetdb)
    wrttet(wrtetdb) = Space(300)
  End If
  If Trim(wmezo$) <> "" Then
    If wrtkp% > 0 And wrtho% > 0 Then
      Mid$(wrttet(wrtetdb), wrtkp%, wrtho%) = wmezo$
    End If
  End If
End Sub

Public Sub wrtbcim(objektum$, wr$)
  '--- wxi, wxp, wxk strukúrák feltöltése WRNS objektumból
  Select Case objektum
    Case "INST"
      wxi.wpkod = ""
      wxi.Wpnev = Trim(Mid$(wr$, 5, 60))
      wxi.wirszam = Trim(Mid$(wr$, 95, 8))
      wxi.wtelep = Trim(Mid$(wr$, 103, 30))
      wxi.wker = Trim(Mid$(wr$, 901, 10))
      wxi.koztern = Trim(Mid$(wr$, 133, 30))
      wxi.kozterj = Trim(Mid$(wr$, 911, 10))
      wxi.hazsz = Trim(Mid$(wr$, 163, 10))
      wxi.epul = Trim(Mid$(wr$, 921, 10))
      wxi.lepcs = Trim(Mid$(wr$, 931, 10))
      wxi.szint = Trim(Mid$(wr$, 941, 10))
      wxi.ajto = Trim(Mid$(wr$, 951, 10))
      wxi.wradosz = Trim(Mid$(wr$, 173, 15))
      wxi.wrcsopado = ""
      wxi.wreuadosz = Trim(Mid$(wr$, 188, 15))
      wxi.wcim = wxi.wirszam + " " + wxi.wtelep
      If wxi.wker <> "" Then wxi.wcim = wxi.wcim + " " + wxi.wker + ", " Else wxi.wcim = wxi.wcim + ", "
      wxi.wcim = wxi.wcim + wxi.koztern + " " + wxi.kozterj + " "
      If wxi.hazsz <> "" Then wxi.wcim = wxi.wcim + wxi.hazsz + "."
      If wxi.epul <> "" Then wxi.wcim = wxi.wcim + wxi.epul + ".ép."
      If wxi.lepcs <> "" Then wxi.wcim = wxi.wcim + wxi.lepcs + ".lh."
      If wxi.szint <> "" Then wxi.wcim = wxi.wcim + wxi.szint + ".em."
      If wxi.ajto <> "" Then wxi.wcim = wxi.wcim + wxi.ajto + ".ajtó"
      wxi.wcim = Trim(wxi.wcim)
      Call wrtbir(100, 12, wxi.wradosz)
      Call wrtbir(124, 20, wxi.wreuadosz)
      Call wrtbir(226, 60, wxi.Wpnev)
      Call wrtbir(286, 10, wxi.wirszam)
      Call wrtbir(296, 30, wxi.wtelep)
      Call wrtbir(326, 10, wxi.wker)
      Call wrtbir(336, 30, wxi.koztern)
      Call wrtbir(366, 10, wxi.kozterj)
      Call wrtbir(376, 10, wxi.hazsz)
      Call wrtbir(386, 10, wxi.epul)
      Call wrtbir(396, 10, wxi.lepcs)
      Call wrtbir(406, 10, wxi.szint)
      Call wrtbir(416, 10, wxi.ajto)
    Case "PART"
      wxp.wpkod = Mid$(wr$, 1, 15)
      wxp.Wpnev = Trim(Mid$(wr$, 16, 60))
      wxp.wirszam = Trim(Mid$(wr$, 106, 8))
      wxp.wtelep = Trim(Mid$(wr$, 114, 30))
      wxp.wker = Trim(Mid$(wr$, 543, 10))
      wxp.koztern = Trim(Mid$(wr$, 144, 30))
      wxp.kozterj = Trim(Mid$(wr$, 553, 10))
      wxp.hazsz = Trim(Mid$(wr$, 174, 10))
      wxp.epul = Trim(Mid$(wr$, 563, 10))
      wxp.lepcs = Trim(Mid$(wr$, 573, 10))
      wxp.szint = Trim(Mid$(wr$, 583, 10))
      wxp.ajto = Trim(Mid$(wr$, 593, 10))
      wxp.wradosz = Trim(Mid$(wr$, 184, 15))
      wxp.wrcsopado = Trim(Mid$(wr$, 855, 15))
      wxp.wreuadosz = Trim(Mid$(wr$, 199, 15))
      wxp.wcic1 = wxp.wirszam + " " + wxp.wtelep
      wxp.wcic2 = ""
      If wxp.wker <> "" Then wxp.wcic1 = wxp.wcic1 + " " + wxp.wker + ", " Else wxp.wcic1 = wxp.wcic1 + ", "
      wxp.wcic2 = wxp.wcic2 + wxp.koztern + " " + wxp.kozterj + " "
      If wxp.hazsz <> "" Then wxp.wcic2 = wxp.wcic2 + wxp.hazsz + "."
      If wxp.epul <> "" Then wxp.wcic2 = wxp.wcic2 + wxp.epul + ".ép."
      If wxp.lepcs <> "" Then wxp.wcic2 = wxp.wcic2 + wxp.lepcs + ".lh."
      If wxp.szint <> "" Then wxp.wcic2 = wxp.wcic2 + wxp.szint + ".em."
      If wxp.ajto <> "" Then wxp.wcic2 = wxp.wcic2 + wxp.ajto + ".ajtó"
      wxp.wcim = Trim(wxp.wcic1 + wxp.wcic2)
      Call wrtbir(441, 12, wxp.wradosz)
      Call wrtbir(465, 20, wxp.wreuadosz)
      Call wrtbir(426, 15, wxp.wpkod)
      Call wrtbir(505, 60, wxp.Wpnev)
      Call wrtbir(565, 10, wxp.wirszam)
      Call wrtbir(575, 30, wxp.wtelep)
      Call wrtbir(605, 10, wxp.wker)
      Call wrtbir(615, 30, wxp.koztern)
      Call wrtbir(645, 10, wxp.kozterj)
      Call wrtbir(655, 10, wxp.hazsz)
      Call wrtbir(665, 10, wxp.epul)
      Call wrtbir(675, 10, wxp.lepcs)
      Call wrtbir(685, 10, wxp.szint)
      Call wrtbir(695, 10, wxp.ajto)
    Case "WRNS"
      If wr$ = "" Then Exit Sub
      wxk.Wpnev = Trim(Mid$(wr$, 63, 60))
      wxk.wirszam = Trim(Mid$(wr$, 123, 10))
      wxk.wtelep = Trim(Mid$(wr$, 133, 30))
      wxk.wker = Trim(Mid$(wr$, 163, 10))
      wxk.koztern = Trim(Mid$(wr$, 173, 30))
      wxk.kozterj = Trim(Mid$(wr$, 203, 10))
      wxk.hazsz = Trim(Mid$(wr$, 213, 10))
      wxk.epul = Trim(Mid$(wr$, 223, 10))
      wxk.lepcs = Trim(Mid$(wr$, 233, 10))
      wxk.szint = Trim(Mid$(wr$, 243, 10))
      wxk.ajto = Trim(Mid$(wr$, 253, 10))
      wxk.wradosz = Trim(Mid$(wr$, 51, 12))
      wxk.wcim = wxp.wirszam + " " + wxp.wtelep
      If wxk.wker <> "" Then wxk.wcim = wxk.wcim + " " + wxk.wker + ", " Else wxk.wcim = wxk.wcim + ", "
      wxk.wcim = wxk.wcim + wxk.koztern + " " + wxk.kozterj + " "
      If wxk.hazsz <> "" Then wxk.wcim = wxk.wcim + wxk.hazsz + "."
      If wxk.epul <> "" Then wxk.wcim = wxk.wcim + wxk.epul + ".ép."
      If wxk.lepcs <> "" Then wxk.wcim = wxk.wcim + wxk.lepcs + ".lh."
      If wxk.szint <> "" Then wxk.wcim = wxk.wcim + wxk.szint + ".em."
      If wxk.ajto <> "" Then wxk.wcim = wxk.wcim + wxk.ajto + ".ajtó"
      wxk.wcim = Trim(wxk.wcim)
      Call wrtbir(441, 12, wxk.wradosz)
      Call wrtbir(717, 60, wxk.Wpnev)
      Call wrtbir(777, 10, wxk.wirszam)
      Call wrtbir(787, 30, wxk.wtelep)
      Call wrtbir(817, 10, wxk.wker)
      Call wrtbir(827, 30, wxk.koztern)
      Call wrtbir(857, 10, wxk.kozterj)
      Call wrtbir(867, 10, wxk.hazsz)
      Call wrtbir(877, 10, wxk.epul)
      Call wrtbir(887, 10, wxk.lepcs)
      Call wrtbir(897, 10, wxk.szint)
      Call wrtbir(907, 10, wxk.ajto)
    Case Else
  End Select
End Sub
Public Sub wrbinp(ok%)
  '--- wrtb beolvasása text fájlból
  ok% = 0
  wfil = FreeFile
  wfin$ = listautvonal$ + terminal$ + task$ + "AUW.WLI"
  Open wfin$ For Input As #wfil
  Line Input #wfil, wrtbz
  wrtetdb = 0
  ReDim wrttet(1 To 1)
  Do
    Line Input #wfil, b$
    wrtetdb = wrtetdb + 1
    ReDim Preserve wrttet(1 To wrtetdb)
    wrttet(wrtetdb) = b$
  Loop While Not EOF(wfil)
  Close wfil
  ok% = 1
End Sub

Public Sub wrbout(ok%)
  '--- wrtb kiírása textfájlba
  ok% = 0
  If wrtetdb > 0 Then
    wfil = FreeFile
    wfin$ = listautvonal$ + terminal$ + task$ + "AUW.WLI"
    Open wfin$ For Output As #wfil
    Mid$(wrtbz, 11, 3) = Right$("000" + Trim(Str(wrtetdb)), 3)
    Print #wfil, wrtbz
    For i31% = 1 To wrtetdb
      b$ = wrttet(i31%)
      Print #wfil, b$
    Next
    Close wfil
    ok% = 1
  End If
End Sub

Public Function wrtujszamlaszam$(fajtaa$)
  '--- új számlaszámok kiadása (ha nem az auwszlahivas rutinban adjuk meg)
  Select Case fajtaa$
    Case "E", "U", "J"
      '--- pénzügy és kereskedelem belföldi elõleg és normál számla
      Mid$(wrnsrec, 14, 6) = novel(wrnsrec$, 14, 6): wrtszamlaszam$ = Mid$(wrnsrec, 10, 10)
    
    Case "X", "Z"
      '--- pénzügy és kereskedelem közösségi és export számla
      Mid$(wrnsrec, 24, 6) = novel(wrnsrec$, 24, 6): wrtszamlaszam$ = Mid$(wrnsrec, 20, 10)
    Case "K"
      '--- kereskedelem bolti számla
      Mid$(wrnsrec, 34, 6) = novel(wrnsrec$, 34, 6): wrtszamlaszam$ = Mid$(wrnsrec, 30, 10)
    Case Else
      wrtszamlaszam$ = ""
  End Select
  wrtujszamlaszam = wrtszamlaszam
End Function

Public Sub auwszlahivas(uzemmod$, fajtaa$, wrtszamlaszam$, cim$)
  '--- az auw-szla program hívása
  wfil = FreeFile
  wfin$ = listautvonal$ + terminal$ + task$ + "AUW.WLH"
  Open wfin For Output As #wfil
  s$ = "CIM=" + cim$: Print #wfil, s$
  s$ = "MOD=" + uzemmod$: Print #wfil, s$
  s$ = "FAJ=" + fajtaa$: Print #wfil, s$
  s$ = "SSZ=" + wrtszamlaszam$: Print #wfil, s$
  If wrtsztornoszam <> "" Then
    s$ = "STS=" + wrtsztornoszam: Print #wfil, s$
  End If
  Close wfil
  Shell programutvonal$ + "auw-szla " + terminal$ + task$ + "/" + listautvonal$, vbHide
End Sub

Public Sub auwszlavalasz(wszamlaszam$)
  '--- az auw-szla program válaszának (WLE) beolvasása
  wfil = FreeFile
  Open listautvonal$ + terminal$ + task$ + "AUW.WLE" For Binary Shared As #wfil
  fm& = LOF(1)
  Close wfil
  If fm& = 0 Then
    wszamlaszam$ = ""
  Else
    wfil = FreeFile
    Open listautvonal$ + terminal$ + task$ + "AUW.WLE" For Input Shared As #wfil
    Line Input #wfil, wszamlaszam$
    Close wfil
  End If
End Sub

Public Sub wrtbeolvas(szamlaszam$, ok%)
  '--- adott számla beolvasása auwszla/wrtb,wrtt -bõl
  wrtbz = dbxkey("WRTB", szamlaszam$)
  If wrtbz <> "" Then
    wrtetdb = xval(Mid$(wrtbz, 11, 3))
    If wrtetdb > 0 Then
      ReDim wrttet(1 To wrtetdb)
      For i41% = 1 To wrtetdb
        inxed$ = Left(wrtbz, 10) + Right$("000" + Trim(Str(i41%)), 3)
        sr$ = dbxkey("WRTT", inxed$)
        If sr$ <> "" Then wrttet(i41%) = sr$
      Next
    End If
    ok% = 1
  Else
    ok% = 0
  End If
End Sub

Public Sub wrtbujir(szamlaszam$, ok%)
  '--- új bizonylat kíírása auwszla/wrtb,wrtt -bõl
  Mid$(wrtbz, 11, 3) = Right$("000" + Trim(Str(wrtetdb)), 3)
  Call dbxki("WRTB", wrtbz, ";", "U", "", hiba%)
  If wrtetdb > 0 Then
    For i31% = 1 To wrtetdb
      sr$ = wrttet(i31%)
      Mid$(sr$, 1, 10) = szamlaszam$
      Mid$(sr$, 11, 3) = Right("000" + Trim(Str(i31%)), 3)
      Call dbxki("WRTT", sr$, ";", "U", "", hiba%)
    Next
  End If
End Sub

Public Sub wrtbstorno(ok%)
  '--- elõjelek megváltoztatása (sztornózés) wrtbz és wrttet srujtúrákban
  Call wrtbmezostorno(1506, 14)
  Call wrtbmezostorno(1520, 14)
  Call wrtbmezostorno(1534, 14)
  Call wrtbmezostorno(1554, 14)
  Call wrtbmezostorno(1668, 14)
  Call wrtbmezostorno(1682, 14)
  Call wrtbmezostorno(1702, 14)
  Call wrtbmezostorno(1716, 14)
  Call wrtbmezostorno(1730, 14)
  Call wrtbmezostorno(1750, 14)
  Call wrtbmezostorno(1764, 14)
  Call wrtbmezostorno(1778, 14)
  Call wrtbmezostorno(1798, 14)
  Call wrtbmezostorno(1812, 14)
  Call wrtbmezostorno(1826, 14)
  Call wrtbmezostorno(1840, 14)
  Call wrtbmezostorno(1854, 14)
  Call wrtbmezostorno(1868, 14)
  Call wrtbmezostorno(1888, 14)
  Call wrtbmezostorno(1902, 14)
  For i31% = 1 To wrtetdb
    Call wrttmezostorno(i31%, 81, 12)
    Call wrttmezostorno(i31%, 220, 14)
    Call wrttmezostorno(i31%, 244, 14)
    Call wrttmezostorno(i31%, 220, 14)
    Call wrttmezostorno(i31%, 268, 14)
  Next
End Sub

Public Sub wrtbmezostorno(wmkp%, wmho%)
  numa@ = xval(Mid$(wrtbz, wmkp%, wmho%))
  If numa@ <> 0 Then
    numa@ = -numa@
    Mid$(wrtbz, wmkp%, wmho%) = ertszam(Str(numa@), wmho%, 2)
  End If
End Sub

Public Sub wrttmezostorno(si%, wmkp%, wmho%)
  numa@ = xval(Mid$(wrttet(si%), wmkp%, wmho%))
  If numa@ <> 0 Then
    numa@ = -numa@
    Mid$(wrttet(si%), wmkp%, wmho%) = ertszam(Str(numa@), wmho%, 2)
  End If
End Sub
Public Sub elolegnyilv(pszbrec$, psztrec$, termrec$)
  '--- 2016.01.01  elõlegnyilvántartás
  pelvrec$ = Space(650)
  Mid$(pelvrec$, 224, 1) = "E"
  Mid$(pelvrec$, 225, 1) = "V"
  Mid$(pelvrec$, 8, 15) = Mid$(pszbrec$, 1, 10)
  Mid$(pelvrec$, 23, 15) = Mid$(pszbrec$, 61, 15)
  Mid$(pelvrec$, 38, 6) = Mid$(pszbrec$, 84, 6)
  Mid$(pelvrec$, 244, 2) = Mid$(psztrec$, 81, 2)
  wafakul$ = Mid$(torolvas("PAFA", Mid$(psztrec$, 81, 2), 1, 50), 33, 6)
  Mid$(pelvrec$, 246, 6) = wafakul$
  Mid$(pelvrec$, 58, 3) = Mid$(pszbrec$, 98, 3)
  Mid$(pelvrec$, 266, 15) = Mid$(psztrec$, 107, 15)
  Mid$(pelvrec$, 281, 14) = "  " + Mid$(psztrec$, 69, 12)
  Mid$(pelvrec$, 295, 14) = "  " + Mid$(psztrec$, 21, 12)
  Mid$(pelvrec$, 309, 6) = Mid$(psztrec$, 33, 6)
  Mid$(pelvrec$, 315, 120) = Mid$(termrec$, 16, 60) + Mid$(termrec$, 196, 60)
  If xval(Mid$(pelvrec$, 295, 14)) <> 0 Then
    wneert@ = xval(Mid$(pelvrec$, 295, 14)) * xval(Mid$(pelvrec$, 281, 14))
  Else
    wneert@ = xval(Mid$(pelvrec$, 281, 14))
  End If
  wafakulcs@ = xval(wafakul$)
  If wafakulcs@ > 0 Then
    wafaosz@ = (wneert@ * wafakulcs@) / 100
    If Trim(Mid$(pelvrec$, 58, 3)) <> "" Then
      wafaosz@ = xval(ertszam(Str(wafaosz@), 14, 2))
    Else
      wafaosz@ = xval(ertszam(Str(wafaosz@), 14, 0))
    End If
  Else
    wafaosz@ = 0
    wafalap@ = wneert@
  End If
  wbruttoo@ = wneert@ + wafaosz@
  Mid$(pelvrec$, 61, 14) = Mid$(psztrec$, 101, 10)
  If Trim(Mid(pszbrec$, 98, 3)) <> "" Then
    Mid$(pelvrec$, 44, 14) = ertszam(Str(wbruttoo@), 14, 2)
    Mid$(pelvrec$, 230, 14) = ertszam(Str(wafalap@), 14, 2)
    Mid$(pelvrec$, 252, 14) = ertszam(Str(wafaosz@), 14, 2)
  Else
    Mid$(pelvrec$, 44, 14) = ertszam(Str(wbruttoo@), 14, 2)
    Mid$(pelvrec$, 230, 14) = ertszam(Str(wafalap@), 14, 2)
    Mid$(pelvrec$, 252, 14) = ertszam(Str(wafaosz@), 14, 0)
  End If
  Mid$(pelvrec$, 491, 8) = Mid$(psztrec$, 83, 8)
  Call dbxki("PELV", pelvrec$, ";", "U", "G", hiba%)
End Sub
                
Public Sub elolegbeszamit(psztrec$, stjel$, beszikt$)
  '--- elõlegbeszámítás 2016.01.01
  '--- stjel="S" storno egyébként beszámítás
  If beszikt$ = "" Then
    beszikt$ = Mid$(psztrec$, 150, 7)
  End If
  If beszikt$ <> "" Then
    pelvrec$ = dbxkey("PELV", beszikt$)
    If pelvrec$ <> "" Then
      If xval(Mid$(psztrec$, 21, 12)) <> 0 Then
        If stjel$ = "S" Then
          wmenny@ = -xval(Mid$(psztrec$, 21, 12))
          wneert@ = -xval(Mid$(psztrec$, 21, 12)) * xval(Mid$(psztrec$, 69, 12))
        Else
          wmenny@ = xval(Mid$(psztrec$, 21, 12))
          wneert@ = xval(Mid$(psztrec$, 21, 12)) * xval(Mid$(psztrec$, 69, 12))
        End If
      Else
        If stjel$ = "S" Then
          wmenny@ = 0
          wneert@ = -xval(Mid$(psztrec$, 69, 12))
        Else
          wmenny@ = 0
          wneert@ = xval(Mid$(psztrec$, 69, 12))
        End If
      End If
      wafakul$ = Mid$(torolvas("PAFA", Mid$(psztrec$, 81, 2), 1, 50), 33, 6)
      wafakulcs@ = xval(wafakul$)
      If wafakulcs@ > 0 Then
        wafaosz@ = (wneert@ * wafakulcs@) / 100
        If Trim(Mid$(pelvrec$, 58, 3)) <> "" Then
          wafaosz@ = xval(ertszam(Str(wafaosz@), 14, 2))
        Else
          wafaosz@ = xval(ertszam(Str(wafaosz@), 14, 0))
        End If
      Else
        wafaosz@ = 0
        wafalap@ = wneert@
      End If
      wbruttoo@ = wneert@ + wafaosz@
      wbruttoo@ = xval(Mid$(pelvrec$, 449, 14)) - wbruttoo@
      wmenny@ = xval(Mid$(pelvrec$, 435, 14)) - wmenny@
      wneert@ = xval(Mid$(pelvrec$, 463, 14)) - wneert@
      wafaosz@ = xval(Mid$(pelvrec$, 477, 14)) - wafaosz@
      Mid$(pelvrec$, 435, 14) = ertszam(Str(wmenny@), 14, 2)
      Mid$(pelvrec$, 449, 14) = ertszam(Str(wbruttoo@), 14, 2)
      Mid$(pelvrec$, 463, 14) = ertszam(Str(wneert), 14, 2)
      If Trim(Mid(pszbrec$, 98, 3)) <> "" Then
        Mid$(pelvrec$, 477, 14) = ertszam(Str(wafaosz@), 14, 2)
      Else
        Mid$(pelvrec$, 477, 14) = ertszam(Str(wafaosz@), 14, 0)
      End If
      Call dbxki("PELV", pelvrec$, ";", "", "", hiba%)
    End If
  End If
End Sub

Public Sub welolegszamlasztorno(weszamla$)
  fil1 = FreeFile
  Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #fil1
  fil2 = FreeFile
  Open auditorutvonal$ + "auw-pelv.ndx" For Binary Shared As #fil2
  rcee& = Int(LOF(fil2) / 12)
  For i61& = 1 To rcee&
    Get #fil2, (i61& - 1) * 12 + 1, rcim&
    r$ = Space(650)
    Get #fil1, rcim& + 9, r$
    If Trim(Mid$(r$, 8, 15)) = weszamla$ Then
      Mid$(r$, 90, 1) = "S"
      Put #fil1, rcim& + 9, r$
    End If
  Next
  Close fil1: Close fil2
End Sub

Public Sub wbutfinit()
  '--- latin2 utf-8 konverziós kódtábla
  utfdb = 18
  utftab(1, 1) = 193: utftab(1, 2) = 195: utftab(1, 3) = 129 '--- Á
  utftab(2, 1) = 201: utftab(2, 2) = 195: utftab(2, 3) = 137 '--- É
  utftab(3, 1) = 205: utftab(3, 2) = 195: utftab(3, 3) = 141 '--- Í
  utftab(4, 1) = 211: utftab(4, 2) = 195: utftab(4, 3) = 147 '--- Ó
  utftab(5, 1) = 214: utftab(5, 2) = 195: utftab(5, 3) = 150 '--- Ö
  utftab(6, 1) = 213: utftab(6, 2) = 197: utftab(6, 3) = 144 '--- Õ
  utftab(7, 1) = 218: utftab(7, 2) = 195: utftab(7, 3) = 154 '--- Ú
  utftab(8, 1) = 220: utftab(8, 2) = 195: utftab(8, 3) = 156 '--- Ü
  utftab(9, 1) = 219: utftab(9, 2) = 197: utftab(9, 3) = 176 '--- Û
  utftab(10, 1) = 225: utftab(10, 2) = 195: utftab(10, 3) = 161 '--- á
  utftab(11, 1) = 233: utftab(11, 2) = 195: utftab(11, 3) = 169 '--- é
  utftab(12, 1) = 237: utftab(12, 2) = 195: utftab(12, 3) = 173 '--- í
  utftab(13, 1) = 243: utftab(13, 2) = 195: utftab(13, 3) = 179 '--- ó
  utftab(14, 1) = 246: utftab(14, 2) = 195: utftab(14, 3) = 182 '--- ö
  utftab(15, 1) = 245: utftab(15, 2) = 197: utftab(15, 3) = 145 '--- õ
  utftab(16, 1) = 250: utftab(16, 2) = 195: utftab(16, 3) = 186 '--- ú
  utftab(17, 1) = 252: utftab(17, 2) = 195: utftab(17, 3) = 188 '--- ü
  utftab(18, 1) = 251: utftab(18, 2) = 197: utftab(18, 3) = 177 '--- û
  kodasc = "": kodutf = ""
  For i99% = 1 To utfdb
    kodasc$ = kodasc$ + Chr$(utftab(i99%, 1))
    kodutf$ = kodutf$ + Chr$(utftab(i99%, 2)) + Chr$(utftab(i99%, 3))
  Next
End Sub

Public Function wbascutf$(szov$)
  '--- konvertálás Latin-2 -> UTF-8
  ilh& = Len(szov$): utfb$ = ""
  If ilh& > 0 Then
    For ilx& = 1 To ilh&
      car$ = Mid$(szov$, ilx&, 1)
      poz1% = InStr(kodasc$, car$)
      If poz1% > 0 Then
        utfb$ = utfb$ + Mid$(kodutf$, (poz1% - 1) * 2 + 1, 2)
      Else
        utfb$ = utfb$ + car$
      End If
    Next
    wbascutf = utfb$
  Else
    wbascutf = ""
  End If
End Function

Public Function wbutfasc$(szov$)
  '--- kovertálás UTF-8 -> Latin-2
  ilh& = Len(szov$): utfb$ = ""
  If ilh& > 0 Then
    ilx& = 1
    Do
      car$ = Mid$(szov$, ilx&, 1)
      If ilx& < ilh& Then
        car2$ = Mid$(szov$, ilx&, 2)
        If car$ = Chr$(195) Or car$ = Chr$(197) Then
          poz1% = InStr(kodutf$, car2$)
          If poz1% > 0 Then
            innx% = Int(poz1 / 2) + 1
            utfb$ = utfb$ + Mid$(kodasc$, innx%, 1): ilx& = ilx& + 2
          Else
            utfb$ = utfb$ + car$: ilx& = ilx& + 1
          End If
        Else
          utfb$ = utfb$ + car$: ilx& = ilx& + 1
        End If
      Else
        utfb$ = utfb$ + car$: ilx& = ilx& + 1
      End If
    Loop While ilx& <= ilh&
    wbutfasc = utfb$
  Else
    wbutfasc = ""
  End If
End Function

Public Sub xmlopen(wbutvonal$, adoszam8$, sorszam4$, aszdatuma$)
  xmlfilnev$ = "NAV_" + adoszam8 + "_" + sorszam4 + "_" + wbedatki(aszdatuma$) + ".xml"
  xmlfil = FreeFile
  Open auditorutvonal$ + xmlfilnev$ For Output As #xmlfil
  wqwq$ = "<?xml version=""1.0"" encoding=""utf-8""?>"
  Print #xmlfil, wqwq$;
End Sub

Public Sub xmlir(mark$, xt$, ujel$)
  bbx$ = Trim(xt$)
  If bbx$ = "" Then
    sw$ = "<" + mark$ + ">"
  Else
    If ujel$ = "U" Then
      sw$ = "<" + mark$ + ">" + wbascutf(bbx$) + "</" + mark$ + ">"
    Else
      sw$ = "<" + mark$ + ">" + bbx$ + "</" + mark$ + ">"
    End If
  End If
  If ujel$ = "U" Then
    Print #xmlfil, sw$
  Else
    Print #xmlfil, sw$
  End If
End Sub

Public Sub xmlclos()
  Close xmlfil
End Sub

Public Function wbedatki$(dat$)
  wx$ = "20" + Mid$(dat$, 1, 2) + "-" + Mid$(dat$, 3, 2) + "-" + Mid$(dat$, 5, 2)
  wbedatki = wx$
End Function

Public Function wbveszpont(szov$)
  wx$ = szov$
  poz1% = InStr(wx$, ",")
  If poz1% > 0 Then
    If poz1% = Len(wx$) Then
      wx$ = Mid$(wx$, 1, poz1% - 1)
    Else
      Mid$(wx$, poz1%, 1) = "."
    End If
  End If
  wbveszpont = wx$
End Function

Public Function wrtbzmj%(kezdo%, hossz%)
  If Trim(Mid(wrtbz, kezdo%, hossz%)) = "" Then
    wrtbzmj = 0
  Else
    wrtbzmj = 1
  End If
End Function

Public Function wrtbzmcar$(kezdo%, hossz%)
  wrtbzmcar = Trim(Mid(wrtbz, kezdo%, hossz%))
End Function

Public Function wrtbzmnum$(kezdo%, hossz%)
  wrtbzmxx@ = xval(Trim(Mid(wrtbz, kezdo%, hossz%)))
  If wrtbzmxx@ <> 0 Then
    wrtbzmnum = wbveszpont(Trim(Mid(wrtbz, kezdo%, hossz%)))
  Else
    wrtbzmnum = ""
  End If
End Function

Public Function wrtttmj%(wttr$, kezdo%, hossz%)
  If Trim(Mid(wttr$, kezdo%, hossz%)) = "" Then
    wrtttmj = 0
  Else
    wrtttmj = 1
  End If
End Function

Public Function wrtttmcar$(wttr$, kezdo%, hossz%)
  wrtttmcar = Trim(Mid(wttr$, kezdo%, hossz%))
End Function

Public Function wrtttmnum$(wttr$, kezdo%, hossz%)
  wrtbzmxx@ = xval(Trim(Mid(wttr$, kezdo%, hossz%)))
  If wrtbzmxx@ <> 0 Then
    wrtttmnum = wbveszpont(Trim(Mid(wttr$, kezdo%, hossz%)))
  Else
    wrtttmnum = ""
  End If
End Function

Public Function wbujcim$(rec$, wjel$)
  If wjel$ = "P" Then
    kerulet$ = Trim(Mid$(rec$, 543, 10))
    kozter$ = Trim(Mid$(rec$, 144, 30))
    kozterj$ = Trim(Mid$(rec$, 553, 10))
    haz$ = Trim(Mid$(rec$, 174, 10))
    epulet$ = Trim(Mid$(rec$, 563, 10))
    lepcso$ = Trim(Mid$(rec$, 573, 10))
    szint$ = Trim(Mid$(rec$, 583, 10))
    ajto$ = Trim(Mid$(rec$, 593, 10))
    sx$ = ""
    If kerulet$ <> "" Then sx$ = sx$ + kerulet$ + " "
    If kozter$ <> "" Then sx$ = sx$ + kozter$ + " "
    If kozterj$ <> "" Then sx$ = sx$ + kozterj$ + " "
    If haz$ <> "" Then sx$ = sx$ + haz$ + " "
    If epulet$ <> "" Then sx$ = sx$ + epulet$ + " "
    If lepcso$ <> "" Then sx$ = sx$ + lepcso$ + " "
    If szint$ <> "" Then sx$ = sx$ + szint$ + " "
    If ajto$ <> "" Then sx$ = sx$ + ajto$ + " "
    wbujcim = Trim(sx$)
  Else
    kerulet$ = Trim(Mid$(rec$, 901, 10))
    kozter$ = Trim(Mid$(rec$, 133, 30))
    kozterj$ = Trim(Mid$(rec$, 911, 10))
    haz$ = Trim(Mid$(rec$, 163, 10))
    epulet$ = Trim(Mid$(rec$, 921, 10))
    lepcso$ = Trim(Mid$(rec$, 931, 10))
    szint$ = Trim(Mid$(rec$, 941, 10))
    ajto$ = Trim(Mid$(rec$, 951, 10))
    sx$ = ""
    If kerulet$ <> "" Then sx$ = sx$ + kerulet$ + " "
    If kozter$ <> "" Then sx$ = sx$ + kozter$ + " "
    If kozterj$ <> "" Then sx$ = sx$ + kozterj$ + " "
    If haz$ <> "" Then sx$ = sx$ + haz$ + " "
    If epulet$ <> "" Then sx$ = sx$ + epulet$ + " "
    If lepcso$ <> "" Then sx$ = sx$ + lepcso$ + " "
    If szint$ <> "" Then sx$ = sx$ + szint$ + " "
    If ajto$ <> "" Then sx$ = sx$ + ajto$ + " "
    wbujcim = Trim(sx$)
  End If
End Function

Public Function kepvadoszam$(rec$)
  kepvadoszam = Trim(Mid$(rec$, 51, 12))
End Function

Public Function kepvisneve$(rec$)
  kepvisneve = Trim(Mid$(rec$, 63, 60))
End Function

Public Function wbteljescim(rec$, wjel$)
  If wjel$ = "K" Then
    irszam$ = Trim(Mid$(rec$, 113, 10))
    telepu$ = Trim(Mid$(rec$, 133, 30))
    kerulet$ = Trim(Mid$(rec$, 163, 10))
    kozter$ = Trim(Mid$(rec$, 173, 30))
    kozterj$ = Trim(Mid$(rec$, 203, 10))
    haz$ = Trim(Mid$(rec$, 213, 10))
    epulet$ = Trim(Mid$(rec$, 223, 10))
    lepcso$ = Trim(Mid$(rec$, 233, 10))
    szint$ = Trim(Mid$(rec$, 243, 10))
    ajto$ = Trim(Mid$(rec$, 253, 10))
    sx$ = Trim(irszam$) + " " + Trim(telepu$) + " "
    If kerulet$ <> "" Then sx$ = sx$ + kerulet$ + " "
    If kozter$ <> "" Then sx$ = sx$ + kozter$ + " "
    If kozterj$ <> "" Then sx$ = sx$ + kozterj$ + " "
    If haz$ <> "" Then sx$ = sx$ + haz$ + " "
    If epulet$ <> "" Then sx$ = sx$ + epulet$ + " "
    If lepcso$ <> "" Then sx$ = sx$ + lepcso$ + " "
    If szint$ <> "" Then sx$ = sx$ + szint$ + " "
    If ajto$ <> "" Then sx$ = sx$ + ajto$ + " "
    wbteljescim = Trim(sx$)
  Else
    If wjel$ = "P" Then
      irszam$ = Trim(Mid$(rec$, 106, 8))
      telepu$ = Trim(Mid$(rec$, 114, 30))
      kerulet$ = Trim(Mid$(rec$, 543, 10))
      kozter$ = Trim(Mid$(rec$, 144, 30))
      kozterj$ = Trim(Mid$(rec$, 553, 10))
      haz$ = Trim(Mid$(rec$, 174, 10))
      epulet$ = Trim(Mid$(rec$, 563, 10))
      lepcso$ = Trim(Mid$(rec$, 573, 10))
      szint$ = Trim(Mid$(rec$, 583, 10))
      ajto$ = Trim(Mid$(rec$, 583, 10))
      sx$ = Trim(irszam$) + " " + Trim(telepu$) + " "
      If kerulet$ <> "" Then sx$ = sx$ + kerulet$ + " "
      If kozter$ <> "" Then sx$ = sx$ + kozter$ + " "
      If kozterj$ <> "" Then sx$ = sx$ + kozterj$ + " "
      If haz$ <> "" Then sx$ = sx$ + haz$ + " "
      If epulet$ <> "" Then sx$ = sx$ + epulet$ + " "
      If lepcso$ <> "" Then sx$ = sx$ + lepcso$ + " "
      If szint$ <> "" Then sx$ = sx$ + szint$ + " "
      If ajto$ <> "" Then sx$ = sx$ + ajto$ + " "
      wbteljescim = Trim(sx$)
    Else
      irszam$ = Trim(Mid$(rec$, 95, 8))
      telepu$ = Trim(Mid$(rec$, 103, 30))
      kerulet$ = Trim(Mid$(rec$, 901, 10))
      kozter$ = Trim(Mid$(rec$, 133, 30))
      kozterj$ = Trim(Mid$(rec$, 911, 10))
      haz$ = Trim(Mid$(rec$, 163, 10))
      epulet$ = Trim(Mid$(rec$, 921, 10))
      lepcso$ = Trim(Mid$(rec$, 931, 10))
      szint$ = Trim(Mid$(rec$, 941, 10))
      ajto$ = Trim(Mid$(rec$, 951, 10))
      sx$ = Trim(irszam$) + " " + Trim(telepu$) + " "
      If kerulet$ <> "" Then sx$ = sx$ + kerulet$ + " "
      If kozter$ <> "" Then sx$ = sx$ + kozter$ + " "
      If kozterj$ <> "" Then sx$ = sx$ + kozterj$ + " "
      If haz$ <> "" Then sx$ = sx$ + haz$ + " "
      If epulet$ <> "" Then sx$ = sx$ + epulet$ + " "
      If lepcso$ <> "" Then sx$ = sx$ + lepcso$ + " "
      If szint$ <> "" Then sx$ = sx$ + szint$ + " "
      If ajto$ <> "" Then sx$ = sx$ + ajto$ + " "
      wbteljescim = Trim(sx$)
    End If
  End If
End Function

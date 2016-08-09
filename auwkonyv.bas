Attribute VB_Name = "auwkonyv"
Public szampkod$, szampmod$, szamprec$, szampdarab%, glomegikt$, globado$, globpartner$
Public szampcim&(30000), szampkie$(500), szampkiedb%, globmukeng$, globmukjel$
Public xbank@, xfennmarad@, xfelossz@, xoegy@, konyveldevnem$
Public xdevbank@, xforintbank@, xaktkie@, xegyenleg@, xdevnembank$
Public xdevszamla@, xforintszamla@, hsxrec$, hsxrecdb&
Dim thtomb$(), tbizt$()
Public ckodok$(3000), pkodok$(3000), ellenszov$
Public zstatusz$, zszcim$
Public welobepartrec$, wujeloleg%, welobemenny@, welobebrutto@, welobeikt$

Type pttyp
  pttkod As String * 15
  pttnev As String * 40
  pttkamsz As Currency           '--- kamatláb
  pttszfrg As Currency           '--- számla forgalom
  pttpfrg As Currency            '--- pénzforgalom
  pttnyall As Currency           '--- nyitó állomány
  pttzall As Currency            '--- záró állomány
  pttszdb As Integer             '--- számla darab
  pttkidb As Integer             '--- kiegyenlítés darab
  pttatfiz As Currency           '--- átlagos kiegyenlítés (nap)
  pttatkes As Currency           '--- átlagos késedelem (nap)
  pttatall As Currency           '--- átlagos állomány
  ptttkam As Currency            '--- terhelhetõ kamat
  pttvkam As Currency            '--- virtuális kamat
  pttzdb As Integer              '--- idõszak végén nyitott számlák száma
  pttall(366) As Currency
End Type
Type pftyp
  pftdat As String * 10
  pftszkpbe As Currency
  pftegkpbe As Currency
  pftkpbe As Currency
  pftszkpki As Currency
  pftegkpki As Currency
  pftkpki As Currency
  pftszbnbe As Currency
  pftegbnbe As Currency
  pftbnbe As Currency
  pftszbnki As Currency
  pftegbnki As Currency
  pftbnki As Currency
End Type

Public Function cbaaraz@(tkod$, pkod$, datum$, arsor%, disz$)
  '--- CBA árazó
  '--- alapár meghatározás
  ktrmrec$ = dbxkey("KTRM", tkod$)
  If arsor% = 0 Then prec$ = dbxkey("PART", pkod$)
  If disz$ = "D" Then
    '--- disztribúciós áron
    cbaaraz@ = xval(Mid$(ktrmrec$, 893, 14)): Exit Function
  Else
    If arsor% <> 0 Then
      iarkat% = arsor%
      alapar@ = xval(Mid$(ktrmrec$, (iarkat% - 1) * 14 + 580, 14))
    Else
      If Mid$(partrec$, 782, 1) = " " Then iarkat% = 1 Else iarkat% = xval(Mid$(partrec$, 782, 1))
      If iarkat% = 0 Then
        cbaaraz@ = xval(Mid$(ktrmrec$, 1274, 14)): Exit Function
      Else
        alapar@ = xval(Mid$(ktrmrec$, (iarkat% - 1) * 14 + 580, 14))
      End If
    End If
  End If
  '--- akció
  akctrec$ = dbxkey("AKCT", Mid$(ktrmrec$, 1, 15))
  If akctrec$ <> "" Then
    For i77% = 1 To 5
      aelem$ = Mid$(akctrec$, (i77% - 1) * 140 + 100, 140)
      akcf$ = Mid$(aelem$, 1, 7)
      akcfrec$ = dbxkey("AKCF", akcf$)
      dat1$ = Mid$(akcfrec$, 41, 6)
      dat2$ = Mid$(akcfrec$, 47, 6)
      If dtm(datum$) >= dtm(dat1$) And dtm(datum$) <= dtm(dat2$) Then
      End If
      If Trim(aelem$) <> "" Then
        dat1$ = Mid$(aelem$, 23, 6)
        dat2$ = Mid$(aelem$, 29, 6)
        If dtm(maidatum$) >= dtm(dat1$) And dtm(maidatum$) <= dtm(dat2$) Then
          If iarkat% = 0 Then
            cbaaraz@ = xval(Mid$(ktrmrec$, 1274, 14)): Exit Function
          Else
            akciar@ = xval(Mid$(aelem$, (iarkat% - 1) * 10 + 45, 10))
            If akciar@ <> 0 And akciar@ < alapar@ Then alapar@ = akciar@
          End If
        End If
      End If
    Next
  End If
  cbaaraz@ = alapar@
End Function

Public Function beszaraz@(tkod$, pkod$, datum$)
  '--- szerzõdött beszerzési ár akcióval
  ktrmrec$ = dbxkey("KTRM", tkod$)
  prec$ = dbxkey("PART", pkod$)
  beszar@ = xval(Mid$(ktrmrec$, 1025, 12))
  refar@ = xval(Mid$(ktrmrec$, 1274, 14))
  '--- akciós ár van-e
  alapar@ = 0
  akctrec$ = dbxkey("AKCT", Mid$(ktrmrec$, 1, 15))
  If akctrec$ <> "" Then
    For i77% = 1 To 5
      aelem$ = Mid$(akctrec$, (i77% - 1) * 140 + 100, 140)
      If Trim(aelem$) <> "" Then
        akciar@ = xval(Mid$(aelem, 105, 10))
        If akciar@ <> 0 Then
          akcf$ = Mid$(aelem$, 1, 7)
          akcfrec$ = dbxkey("AKCF", akcf$)
          dat1$ = Mid$(akcfrec$, 41, 6)
          dat2$ = Mid$(akcfrec$, 47, 6)
          If dtm(datum$) >= dtm(dat1$) And dtm(datum$) <= dtm(dat2$) Then
            If akciar@ <> 0 And (alapar@ = 0 Or akciar@ < alapar@) Then alapar@ = akciar@
          End If
        End If
      End If
    Next
  End If
  If alapar@ <> 0 Then beszaraz@ = alapar@: Exit Function
  If prec$ <> "" Then
    crmind$ = pkod$ + tkod$ + "S"
    crmrec$ = dbxkey("CRMG", crmind$)
    If crmrec$ <> "" Then
      If UCase(Mid$(crmrec$, 62, 1)) <> "N" Then
        dat1$ = Mid$(crmrec$, 32, 6)
        dat2$ = Mid$(crmrec$, 38, 6)
        If dtm(datum$) >= dtm(dat1$) And dtm(datum$) <= dtm(dat2$) Then
          alapar@ = xval(Mid(crmrec$, 50, 12))
          If alapar@ <> 0 Then beszaraz@ = alapar@: Exit Function
        End If
      End If
    End If
  End If
  '--- ha idáig eljutott, extra ár vagy listaar
  If beszar@ <> 0 Then
    beszaraz@ = beszar@
  Else
    beszaraz@ = refar@
  End If
End Function

Public Sub pktetk(rec$)
  '--- lerakat pénztár könyvelés
  ptfok$ = Mid$(rec$, 22, 8)
  kod$ = Mid$(rec$, 55, 1)
  If kod$ <> "X" And kod$ <> "V" And kod$ <> "E" Then Exit Sub
  irany$ = Mid$(rec$, 8, 1)
  osszeg@ = xval(Mid$(rec$, 56, 14))
  If osszeg@ < 0 Then
    Mid$(rec$, 56, 14) = ertszam(Str(Abs(osszeg@)), 14, 2)
    If irany$ = "B" Then irany$ = "K" Else irany$ = "B"
    Mid$(rec$, 8, 1) = irany$
  End If
  irec$ = dbxkey("INST", "INST")
  '--- pénztárbiz szám
  ptbizszam$ = Mid$(rec$, 9, 7)
  '--- könyvelési biz szám
  fbizszam$ = novel(irec$, 328, 7)
  Call dbxki("INST", irec$, ";", "", "", hiba%)
  Call dbxki("PKTE", rec$, ";", "U", "G", hiba%)
  '--- partner kód
  pkod$ = Mid$(rec$, 108, 15)
  If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod$) Else partrec$ = ""
  '--- könyvelési számlák megállapítása
  ymkod$ = "9999"
  Select Case kod$
    Case "X"
      ymkod$ = "3033"
      If irany$ = "B" Then
        tsy$ = ptfok$
        ksy$ = Mid$(rec$, 146, 8)
        kky$ = Mid$(rec$, 154, 8)
        svy$ = Mid$(rec$, 162, 8)
        msy$ = Mid$(rec$, 170, 8)
      Else
        ksy$ = ptfok$
        tsy$ = Mid$(rec$, 146, 8)
        tky$ = Mid$(rec$, 154, 8)
        svy$ = Mid$(rec$, 162, 8)
        msy$ = Mid$(rec$, 170, 8)
      End If
    Case "V", "E"
      ymkod$ = "1133"
      If irany$ = "K" Then
        ksy$ = ptfok$
        tsy$ = Mid$(partrec$, 298, 8)
      Else
        tsy$ = ptfok$
        ksy$ = Mid$(partrec$, 298, 8)
      End If
      tky$ = Space$(8)
      kky$ = Space$(8)
      svy$ = Space$(8)
      msy$ = Space$(8)
    Case Else
  End Select
  '--- könyvelési rekord feltöltése
  trec$ = Space$(320)
  Mid$(trec$, 8, 7) = fbizszam$
  Mid$(trec$, 15, 7) = ptbizszam$
  Mid$(trec$, 30, 15) = Mid$(rec$, 123, 15)
  Mid$(trec$, 45, 4) = "TPPP"
  Mid$(trec$, 49, 6) = Mid$(rec$, 16, 6)
  Mid$(trec$, 55, 6) = maidatum$
  Mid$(trec$, 68, 8) = ugyintezo$
  Mid$(trec$, 100, 8) = tsy$
  Mid$(trec$, 108, 8) = tky$
  Mid$(trec$, 116, 8) = ksy$
  Mid$(trec$, 124, 8) = kky$
  Mid$(trec$, 84, 8) = svy$
  Mid$(trec$, 92, 8) = msy$
  Mid$(trec$, 132, 31) = Mid$(rec$, 56, 31)
  Mid$(trec$, 192, 60) = ""
  Mid$(trec$, 163, 15) = pkod$
  Mid$(trec$, 185, 7) = Mid$(rec$, 1, 7)
  Mid$(trec$, 252, 25) = Mid$(rec$, 30, 25)
  Mid$(trec$, 277, 4) = ymkod$
  '--- rekord felirasa
  Call fktkonyvel(trec$, "U")
  Mid$(rec$, 207, 7) = Mid$(trec$, 1, 7)
  Call dbxki("PKTE", rec$, ";", "", "", hiba%)
  '--- pénztár és fkte lekönyvelve
  '--- folyószámla kiegyenlítls kezelése
  vikt$ = Mid$(rec$, 87, 7)
  eikt$ = Mid(rec$, 101, 7)
  Select Case kod$
    Case "X"
      Mid$(rec$, 215, 2) = "09"
      Call dbxki("PKTE", rec$, ";", "", "", hiba%)
    Case "E"
      '--- vevõ kiegyenlítés
      '--- vevõ egyenleg
      '--- partner láncrafûzés
      If Trim(eikt) <> "" Then
        erec$ = dbxkey("PELO", eikt$)
        If partrec$ <> "" And erec$ <> "" Then
          '--- partner egyenleg módosítása
          ooo@ = xval(Mid$(rec$, 56, 14))
          Call kivvon(partrec$, 659, 14, ooo@, 2)
          '--- partner láncrafûzése
          w1% = obsorszama("PKTE")
          aktucim& = OBJTAB(w1%).obcim
          Call lancra("AUWSZAMV", "PTPART", partrec$, aktucim&, rec$)
          '--- pénzforgalmi könyvelési tétel könyvelése
          Mid$(rec$, 215, 2) = "01"
          Call dbxki("PKTE", rec$, ";", "", "", hiba%)
          '--- kiegyenlítés beírása a számlába
          Mid$(erec$, 105, 7) = Mid$(rec$, 1, 7)
          Mid$(erec$, 127, 7) = Mid$(trec$, 1, 7)
          Mid$(erec$, 119, 8) = Mid$(rec$, 22, 8)
          Call dbxki("PELO", erec$, ";", "", "", hiba%)
        End If
      End If
    Case "V"
      '--- vevõ kiegyenlítés
      '--- vevõ egyenleg
      '--- partner láncrafûzés
      If Trim(vikt) <> "" Then
        vrec$ = dbxkey("PVSZ", vikt$)
        If partrec$ <> "" And vrec$ <> "" Then
          '--- partner egyenleg módosítása
          ooo@ = xval(Mid$(rec$, 56, 14))
          Call kivvon(partrec$, 659, 14, ooo@, 2)
          '--- partner láncrafûzése
          w1% = obsorszama("PKTE")
          aktucim& = OBJTAB(w1%).obcim
          Call lancra("AUWSZAMV", "PTPART", partrec$, aktucim&, rec$)
          '--- pénzforgalmi könyvelési tétel könyvelése
          Mid$(rec$, 215, 2) = "01"
          Call dbxki("PKTE", rec$, ";", "", "", hiba%)
          '--- kiegyenlítés beírása a számlába
          tal2% = 0
          For i2% = 1 To 10
            elem1$ = Mid$(vrec$, (i2% - 1) * 35 + 930, 35)
            If Trim$(elem1$) = "" Then tal2% = i2%: Exit For
          Next
          If tal2% > 0 Then
            elem1$ = Space$(35)
            If irany$ = "K" Then
              Mid$(elem1$, 1, 14) = Right$(Space$(14) + Format(-ooo@, "##########0.00"), 14)
            Else
              Mid$(elem1$, 1, 14) = Right$(Space$(14) + Format(ooo@, "##########0.00"), 14)
            End If
            Mid$(elem1$, 15, 8) = Mid$(rec$, 22, 8)
            Mid$(elem1$, 23, 6) = Mid$(rec$, 16, 6)
            Mid$(elem1$, 29, 7) = Mid$(trec$, 1, 7)
            Mid$(vrec$, (tal2% - 1) * 35 + 930, 35) = elem1$
            Call dbxki("PVSZ", vrec$, ";", "", "", hiba%)
          End If
        End If
      End If
    Case Else
  End Select
End Sub

Public Sub pktekonyvel(rec$, ptfok$, pjel%, ptbizszam$)
  '--- pénztári tétel könyvelése csak V,S,M,X
  '--- pjel%=0 pénztárbiz szám megvan
  '--- pjel%=1 pénztárbiz.szám kell
  kod$ = Mid$(rec$, 55, 1)
  irany$ = Mid$(rec$, 8, 1)
  osszeg@ = xval(Mid$(rec$, 56, 14))
  If osszeg@ < 0 Then
    Mid$(rec$, 56, 14) = ertszam(Str(Abs(osszeg@)), 14, 2)
    If irany$ = "B" Then irany$ = "K" Else irany$ = "B"
    Mid$(rec$, 8, 1) = irany$
  End If
  irec$ = dbxkey("INST", "INST")
  '--- pénztárbiz szám
  If pjel% = 1 Then
    If irany$ = "B" Then
      bizszam$ = novel(irec$, 307, 7)
      Mid$(irec$, 307, 7) = bizszam$
    Else
      bizszam$ = novel(irec$, 314, 7)
      Mid$(irec$, 314, 7) = bizszam$
    End If
    ptbizszam$ = bizszam$
    Call dbxki("INST", irec$, ";", "", "", hiba%)
    Mid$(rec$, 9, 7) = ptbizszam$
  Else
    ptbizszam$ = Mid$(rec$, 9, 7)
  End If
  '--- könyvelési biz szám
  fbizszam$ = novel(irec$, 328, 7)
  Call dbxki("INST", irec$, ";", "", "", hiba%)
  Call dbxki("PKTE", rec$, ";", "U", "G", iba%)
  '--- partner kód
  vikt$ = Mid$(rec$, 87, 7)
  sikt$ = Mid$(rec$, 94, 7)
  pkod$ = Mid$(rec$, 108, 15)
  If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod$) Else partrec$ = ""
  '--- könyvelési számlák megállapítása
  ymkod$ = "9999"
  Select Case kod$
    Case "X"
      ymkod$ = "3033"
      If irany$ = "B" Then
        tsy$ = ptfok$
        ksy$ = Mid$(rec$, 146, 8)
        kky$ = Mid$(rec$, 154, 8)
        svy$ = Mid$(rec$, 162, 8)
        msy$ = Mid$(rec$, 170, 8)
      Else
        ksy$ = ptfok$
        tsy$ = Mid$(rec$, 146, 8)
        tky$ = Mid$(rec$, 154, 8)
        svy$ = Mid$(rec$, 162, 8)
        msy$ = Mid$(rec$, 170, 8)
      End If
    Case "V"
      ymkod$ = "1133"
      If irany$ = "K" Then
        ksy$ = ptfok$
        tsy$ = Mid$(partrec$, 298, 8)
      Else
        tsy$ = ptfok$
        ksy$ = Mid$(partrec$, 298, 8)
      End If
      tky$ = Space$(8)
      kky$ = Space$(8)
      svy$ = Space$(8)
      msy$ = Space$(8)
    Case "S"
      ymkod$ = "2133"
      If irany$ = "B" Then
        tsy$ = ptfok$
        ksy$ = Mid$(partrec$, 306, 8)
      Else
        ksy$ = ptfok$
        tsy$ = Mid$(partrec$, 306, 8)
      End If
      tky$ = Space$(8)
      kky$ = Space$(8)
      svy$ = Space$(8)
      msy$ = Space$(8)
    Case Else
  End Select
  '--- könyvelési rekord feltöltése
  trec$ = Space$(320)
  Mid$(trec$, 8, 7) = fbizszam$
  Mid$(trec$, 15, 7) = ptbizszam$
  Mid$(trec$, 30, 15) = Mid$(rec$, 123, 15)
  Mid$(trec$, 45, 4) = "TPPP"
  Mid$(trec$, 49, 6) = Mid$(rec$, 16, 6)
  Mid$(trec$, 55, 6) = maidatum$
  Mid$(trec$, 68, 8) = ugyintezo$
  Mid$(trec$, 100, 8) = tsy$
  Mid$(trec$, 108, 8) = tky$
  Mid$(trec$, 116, 8) = ksy$
  Mid$(trec$, 124, 8) = kky$
  Mid$(trec$, 84, 8) = svy$
  Mid$(trec$, 92, 8) = msy$
  Mid$(trec$, 132, 31) = Mid$(rec$, 56, 31)
  Mid$(trec$, 192, 60) = ""
  Mid$(trec$, 163, 15) = pkod$
  Mid$(trec$, 185, 7) = Mid$(rec$, 1, 7)
  Mid$(trec$, 252, 25) = Mid$(rec$, 30, 25)
  Mid$(trec$, 277, 4) = ymkod$
  '--- rekord felirasa
  Call fktkonyvel(trec$, "U")
  Mid$(rec$, 207, 7) = Mid$(trec$, 1, 7)
  '--- pénztár és fkte lekönyvelve
  '--- folyószámla kiegyenlítls kezelése
  Select Case kod$
    Case "X"
      Mid$(rec$, 215, 2) = "09"
      Call dbxki("PKTE", rec$, ";", "", "", hiba%)
    Case "V"
      '--- vevõ kiegyenlítés
      '--- vevõ egyenleg
      '--- partner láncrafûzés
      vrec$ = dbxkey("PVSZ", vikt$)
      If partrec$ <> "" And vrec$ <> "" Then
        '--- partner egyenleg módosítása
        ooo@ = xval(Mid$(rec$, 56, 14))
        Call kivvon(partrec$, 659, 14, ooo@, 2)
        '--- partner láncrafûzése
        w1% = obsorszama("PKTE")
        aktucim& = OBJTAB(w1%).obcim
        Call lancra("AUWSZAMV", "PTPART", partrec$, aktucim&, rec$)
        '--- pénzforgalmi könyvelési tétel könyvelése
        Mid$(rec$, 215, 2) = "01"
        Call dbxki("PKTE", rec$, ";", "", "", hiba%)
        '--- kiegyenlítés beírása a számlába
        tal2% = 0
        For i2% = 1 To 10
          elem1$ = Mid$(vrec$, (i2% - 1) * 35 + 930, 35)
          If Trim$(elem1$) = "" Then tal2% = i2%: Exit For
        Next
        If tal2% > 0 Then
          elem1$ = Space$(35)
          If irany$ = "K" Then
            Mid$(elem1$, 1, 14) = Right$(Space$(14) + Format(-ooo@, "##########0.00"), 14)
          Else
            Mid$(elem1$, 1, 14) = Right$(Space$(14) + Format(ooo@, "##########0.00"), 14)
          End If
          Mid$(elem1$, 15, 8) = Mid$(rec$, 22, 8)
          Mid$(elem1$, 23, 6) = Mid$(rec$, 16, 6)
          Mid$(elem1$, 29, 7) = Mid$(trec$, 1, 7)
          Mid$(vrec$, (tal2% - 1) * 35 + 930, 35) = elem1$
          Call dbxki("PVSZ", vrec$, ";", "", "", hiba%)
        End If
      End If
    Case "S"
      '--- szállító kiegyenlítés
      '--- partner láncrafûzés
      srec$ = dbxkey("PSSZ", sikt$)
      If partrec$ <> "" And srec$ <> "" Then
        '--- partner egyenleg módosítása
        ooo@ = xval(Mid$(rec$, 56, 14))
        '--- partner láncrafûzése
        w1% = obsorszama("PKTE")
        aktucim& = OBJTAB(w1%).obcim
        Call lancra("AUWSZAMV", "PTPART", partrec$, aktucim&, rec$)
        '--- pénzforgalmi könyvelési tétel könyvelése
        Mid$(rec$, 215, 2) = "02"
        Call dbxki("PKTE", rec$, ";", "", "", hiba%)
        '--- kiegyenlítés beírása a számlába
        tal2% = 0
        For i2% = 1 To 10
          elem1$ = Mid$(srec$, (i2% - 1) * 35 + 930, 35)
          If Trim$(elem1$) = "" Then tal2% = i2%: Exit For
        Next
        If tal2% > 0 Then
          elem1$ = Space$(35)
          If irany$ = "B" Then
            Mid$(elem1$, 1, 14) = Right$(Space$(14) + Format(-ooo@, "##########0.00"), 14)
          Else
            Mid$(elem1$, 1, 14) = Right$(Space$(14) + Format(ooo@, "##########0.00"), 14)
          End If
          Mid$(elem1$, 15, 8) = Mid$(rec$, 22, 8)
          Mid$(elem1$, 23, 6) = Mid$(rec$, 16, 6)
          Mid$(elem1$, 29, 7) = Mid$(trec$, 1, 7)
          Mid$(srec$, (tal2% - 1) * 35 + 930, 35) = elem1$
          Call dbxki("PSSZ", srec$, ";", "", "", hiba%)
        End If
      End If
    Case Else
  End Select
End Sub

Public Sub pktesztrn(rec$)
  '--- pénztári tétel sztornó
  
End Sub

Public Function szegyen@(r$, fnap$, kfnap$)
  '--- számla egyenleg
  If dtm(fnap$) >= dtm(Mid$(r$, 58, 6)) Then
    szosz@ = xval(Mid$(r$, 78, 14))
  Else
    szosz@ = 0
  End If
  For i91% = 1 To 10
    elem$ = Mid$(r$, (i91% - 1) * 35 + 930, 35)
    If dtm(kfnap$) >= dtm(Mid$(elem$, 23, 6)) Then
      kie@ = xval(Mid$(elem$, 1, 14))
      szosz@ = szosz@ - kie@
    End If
  Next
  szegyen@ = szosz@
End Function

Public Sub kiegolvas(vsj$, vsikt$, kiegrec$)
  If vsj$ = "V" Then
    kiegrec$ = dbxkey("PVSK", vsikt$)
  Else
    kiegrec$ = dbxkey("PSSK", vsikt$)
  End If
End Sub
Public Function arazo@(partkod$, termkod$, minta$, datum$, disztribucio$)
  '--- disztribúciós árképzés
  If disztribucio$ = "D" Then
    '--- disztribúciós árképzés
    partrec$ = dbxkey("PART", partkod$)
    If partrec$ <> "" Then
      If Mid$(partrec$, 799, 1) = "D" Then
        termrec$ = dbxkey("KTRM", termkod$)
        dar@ = xval(Mid$(termrec$, 893, 14))
        If dar@ <> 0 Then arazo@ = dar@: Exit Function
      End If
    End If
  End If
  '--- kereskedelmi árképzés
  aktara@ = 2000000000
  partrec$ = dbxkey("PART", partkod$)
  If partrec$ <> "" Then
    parteng1$ = Trim(Mid$(partrec$, 333, 6))
    If Right$(parteng1$, 1) = "-" Then
      parteng@ = -xval(Left$(parteng1$, Len(parteng1$) - 1))
    Else
      parteng@ = xval(parteng1$)
    End If
    termrec$ = dbxkey("KTRM", termkod$)
    armkod$ = termkod$ + partkod$
    armgrec$ = dbxkey("ARMG", armkod$)
    vanar% = 0: vaneng% = 0
    If armgrec$ = "" Then
    Else
      For i91% = 20 To 1 Step -1
        arelem$ = Mid(armgrec$, (i91% - 1) * 40 + 200, 40)
        If Trim(arelem$) <> "" Then
          If dtm(datum$) >= dtm(Mid$(arelem$, 1, 6)) And dtm(datum$) <= dtm(Mid$(arelem$, 7, 6)) Then
            If Mid$(arelem$, 33, 8) = Space$(8) Or Mid$(arelem$, 33, 8) = minta$ Then
              ar@ = xval(Mid$(arelem$, 13, 14))
              speceng@ = xval(Mid$(arelem$, 27, 6))
              If ar@ <> 0 Then
                If ar@ < aktara@ Then
                  aktara@ = ar@
                  arazo@ = ar@
                  GoTo arkilep
                End If
              End If
              If speceng@ <> 0 Then vaneng% = 1
            End If
          End If
        End If
      Next
    End If
    If Mid$(partrec$, 782, 1) = " " Then artip% = 1 Else artip% = xval(Mid$(partrec$, 782, 1))
    If artip% = 0 Then
      alapar@ = xval(Mid$(termrec$, 1274, 14))
    Else
      alapar@ = xval(Mid$(termrec$, (artip% - 1) * 14 + 580, 14))
    End If
    If vaneng% = 0 Then eng@ = parteng@ Else eng@ = speceng@
    If eng@ <> 0 Then
      ujar@ = alapar@ + alapar@ * eng@ / 100
      If ujar@ < aktara@ Then aktara@ = ujar@
      'arazo@ = xval(ertszam(Str(ujar@), 14, 2))
      arazo@ = xval(ertszam(Str(aktara@), 14, 2))
    Else
      If alapar@ < aktara@ Then aktara@ = alapar@
      arazo@ = aktara@
      'arazo@ = alapar@
    End If
  Else
    arazo = 0
    Exit Function
  End If
arkilep:
  '--- akciók kezelése
  gyartokod$ = Mid$(termrec$, 860, 15)
  tcsopkod$ = Mid$(termrec$, 438, 4)
  vkat$ = Mid$(partrec$, 314, 1)
  afil = FreeFile
  Open auditorutvonal$ + "akcio.par" For Binary Shared As #afil
  afim& = LOF(afil)
  Close afil
  If afim& < 10 Then Exit Function
  afil = FreeFile
  Open auditorutvonal$ + "akcio.par" For Input Shared As #afil
  Do
    Line Input #afil, aw$
    rendben% = 1
    If Mid$(aw$, 1, 1) <> "N" And Mid$(aw$, 1, 1) <> "M" Then rendben% = 0
    tttx$ = Trim(Mid(aw$, 2, 15)): tttl% = Len(tttx$)
    If tttl% > 0 And rendben% = 1 Then
      If Left(termkod$, tttl%) <> tttx$ Then rendben% = 0
    End If
    cccx$ = Trim(Mid(aw$, 17, 4)): cccl% = Len(cccx$)
    If cccl% > 0 And rendben% = 1 Then
      If Left(tcsopkod$, cccl%) <> cccx$ Then rendben% = 0
    End If
    pppx$ = Trim(Mid(aw$, 36, 15)): pppl% = Len(pppx$)
    If pppl% > 0 And rendben% = 1 Then
      If Left(partkod$, pppl%) <> pppx$ Then rendben% = 0
    End If
    gggx$ = Trim(Mid(aw$, 21, 15)): gggl% = Len(gggx$)
    If gggl% > 0 And rendben% = 1 Then
      If Left(gyartokod$, gggl%) <> gggx$ Then rendben% = 0
    End If
    vvvx$ = Trim(Mid(aw$, 51, 1)): vvvl% = Len(vvvx$)
    If vvvl% > 0 And rendben% = 1 Then
      If Left(vkat$, vvvl%) <> vvvx$ Then rendben% = 0
    End If
    If rendben% = 1 Then
      d1$ = Mid$(aw$, 52, 6): d2$ = Mid$(aw$, 58, 6)
      If dtm(d1$) <= dtm(datum$) And dtm(d2$) >= dtm(datum$) Then
        ar@ = xval(Mid$(aw$, 64, 12))
        speceng@ = xval(Mid$(aw$, 76, 12))
        If ar@ > 0 Then
          If ar@ < aktara@ Then aktara@ = ar@
          arazo@ = aktara@
          Close afil
          If aktara@ = 2000000000 Then arazo@ = 0
          Exit Function
        Else
          If speceng@ <> 0 Then
            If Mid$(partrec$, 782, 1) = " " Then artip% = 1 Else artip% = xval(Mid$(partrec$, 782, 1))
            If artip% = 0 Then
              alapar@ = xval(Mid$(termrec$, 1274, 14))
            Else
              alapar@ = xval(Mid$(termrec$, (artip% - 1) * 14 + 580, 14))
            End If
            ujar@ = alapar@ + alapar@ * speceng@ / 100
            If ujar@ < aktara@ Then aktara@ = ujar@
            arazo@ = xval(ertszam(Str(aktara@), 14, 2))
            'arazo@ = xval(ertszam(Str(ujar@), 14, 2))
            Close afil
            If aktara@ = 2000000000 Then arazo@ = 0
            Exit Function
          End If
        End If
      End If
    End If
  Loop While Not EOF(afil)
  Close afil
  If aktara@ = 2000000000 Then arazo@ = 0
End Function

Public Sub komfoglal(jel$, tkod$, menny@)
  '--- foglalás komissiózható készletbõl
  meo@ = menny@
  If jel$ = "F" Then
    For i91% = 1 To komiraktdb
      raktarkod$ = komirakt(i91%)
      Call rkeszlet(tkod$, "", raktartkod$, keszme@, foglme@)
      If keszme@ - foglme@ >= meo@ Then
        Call foglal("F", raktarkod$, tkod$, "", meo@)
        Exit Sub
      Else
        If keszme - foglme@ > 0 Then
          xx@ = keszme@ - foglme@
          Call foglal("F", raktarkod$, tkod$, "", xx@)
          meo@ = meo@ - xx@
        Else
        End If
      End If
      If meo@ <= 0 Then Exit Sub
    Next
  Else
    For i91% = 1 To komiraktdb
      raktarkod$ = komirakt(i91%)
      Call rkeszlet(tkod$, "", raktartkod$, keszme@, foglme@)
      If foglme@ >= meo@ Then
        Call foglal("N", raktarkod$, tkod$, "", meo@)
        Exit Sub
      Else
        If foglme@ > 0 Then
          xx@ = foglme@
          Call foglal("N", raktarkod$, tkod$, "", xx@)
          meo@ = meo@ - xx@
        Else
        End If
      End If
      If meo@ <= 0 Then Exit Sub
    Next
  End If
End Sub

Public Sub foglal(jel$, rakod$, tkod$, minta$, menny@)
  '--- foglalás jel=F-foglal =N-felszabadit
  rkszind$ = rakod$ + tkod$
  rkszrec$ = dbxkey("RKSZ", rkszind)
  If rkszrec$ <> "" Then
    If minta$ = "" Or minta$ = Space$(8) Then
      keszme@ = xval(Mid$(rkszrec$, 20, 12))
      foglme@ = xval(Mid$(rkszrec$, 32, 12))
      If jel$ = "F" Then
        If menny@ > keszme@ - foglme@ Then menny@ = keszme@ - foglme@
        foglme@ = foglme@ + menny@
        Mid$(rkszrec$, 32, 12) = ertszam(Str(foglme@), 12, 3)
        If xval(Mid$(rkszrec$, 32, 12)) < 0 Then Mid$(rkszrec$, 32, 12) = Space$(12)
      Else
        foglme@ = foglme@ - menny@: If foglme@ < 0 Then foglme@ = 0
        Mid$(rkszrec$, 32, 12) = ertszam(Str(foglme@), 12, 3)
        If xval(Mid$(rkszrec$, 32, 12)) < 0 Then Mid$(rkszrec$, 32, 12) = Space$(12)
      End If
    Else
      For i91% = 1 To 300
        elem$ = Mid$(rkszrec$, (i91% - 1) * 28 + 200, 28)
        If Mid$(elem$, 1, 8) = minta$ Then
          keszme@ = xval(Mid$(elem$, 9, 10))
          foglme@ = xval(Mid$(elem$, 19, 10))
          If jel$ = "F" Then
            If menny@ > keszme@ - foglme@ Then menny@ = keszme@ - foglme@
            foglme@ = foglme@ + menny@
            Mid$(elem$, 19, 10) = ertszam(Str(foglme@), 10, 2)
            If xval(Mid$(elem$, 19, 10)) < 0 Then Mid$(elem$, 19, 10) = Space$(10)
            Mid$(rkszrec$, (i91% - 1) * 28 + 200, 28) = elem$
          Else
            foglme@ = foglme@ - menny@: If foglme@ < 0 Then foglme@ = 0
            Mid$(elem$, 19, 10) = ertszam(Str(foglme@), 10, 2)
            If xval(Mid$(elem$, 19, 10)) < 0 Then Mid$(elem$, 19, 10) = Space$(10)
            Mid$(rkszrec$, (i91% - 1) * 28 + 200, 28) = elem$
          End If
          Exit For
        End If
      Next
    End If
    Call dbxki("RKSZ", rkszrec$, ";", "", "", hiba%)
  End If
  termrec$ = dbxkey("KTRM", tkod$)
  If termrec$ <> "" Then
    If jel$ = "F" Then
      Call hozzad(termrec$, 762, 14, menny@, 3)
    Else
      Call kivvon(termrec$, 762, 14, menny@, 3)
    End If
    Call dbxki("KTRM", termrec$, ";", "", "", hiba%)
  End If
End Sub

Public Sub cplckeszlet(tkod$, keszme@, rakta$)
  termrec$ = dbxkey("KTRM", tkod$)
  If rakta$ = "0002" Then
    keszme@ = xval(Mid$(termrec$, 955, 14)) + xval(Mid$(termrec$, 941, 14))
  Else
    keszme@ = xval(Mid$(termrec$, 955, 14))
  End If
End Sub

Public Sub komkeszlet(tkod$, keszme@, foglme@, cpjel$, zpjel$)
  '--- komissiózható készlet
  '--- cpjel$=C C+C és vegyi készlet, egyébként üres
  If cpjel$ = "" Then
    If zpjel$ = "" Then
      termrec$ = dbxkey("KTRM", tkod$)
      keszme@ = xval(Mid$(termrec$, 983, 14))
      foglme@ = xval(Mid$(termrec$, 762, 14))
    Else
      termrec$ = dbxkey("KTRM", tkod$)
      keszme@ = xval(Mid$(termrec$, 1001, 12))
      foglme@ = 0
    End If
  Else
    termrec$ = dbxkey("KTRM", tkod$)
    keszme@ = xval(Mid$(termrec$, 955, 14)) + xval(Mid$(termrec$, 941, 14))
    foglme@ = 0
  End If
End Sub

Public Sub rkeszlet(tkod$, minta$, rakod$, keszme@, foglme@)
  '--- kereskedelmi készlet raktárkészlete
  If telkom.Ltvan = 1 Then
    tirec$ = dbxkey("KTRM", tkod$)
    If tirec$ <> "" Then
      keszme@ = xval(Mid$(tirec$, 748, 14))
      foglme@ = 0
      Exit Sub
    End If
  Else
    rkszind$ = rakod$ + tkod$
    rkszrec$ = dbxkey("RKSZ", rkszind)
    If rkszrec$ <> "" Then
      If minta$ = "" Or minta$ = Space$(8) Then
        keszme@ = xval(Mid$(rkszrec$, 20, 12))
        foglme@ = xval(Mid$(rkszrec$, 32, 12))
        Exit Sub
      Else
        For i91% = 1 To 300
          elem$ = Mid$(rkszrec$, (i91% - 1) * 28 + 200, 28)
          If Mid$(elem$, 1, 8) = minta$ Then
            keszme@ = xval(Mid$(elem$, 9, 10))
            foglme@ = xval(Mid$(elem$, 19, 10))
            Exit Sub
          End If
        Next
      End If
    End If
  End If
  keszme@ = 0: foglme@ = 0
End Sub

Public Sub gkeszletvalt(irany$, frec$, merlmod$, tvagyf$, dat$)
  '--- tapado gongyöleges készletváltozás merka c+c, 5ker
  '--- irany= B-bevét K-kiadás
  '--- frec = a tétel rekord kkft vagy ertt
  '--- merlmod = K- változtatja az átlagárat egyébként nem
  '--- tvagyf  = T-kkft rekord F-ertt rekord
  If dtm(dat$) <= dtm(kerlezardat$) Then Exit Sub
  tapgo$ = ""
  If tvagyf$ = "T" Then
    '--- kkft
    menny@ = xval(Mid$(frec$, 71, 12))
    'fogl@ = xval(Mid$(frec$, 83, 12))
    bear@ = xval(Mid$(frec$, 59, 12))
    raktarkod$ = Mid$(frec$, 24, 4)
    tkod$ = Mid$(frec$, 36, 15)
    tapgo$ = Mid$(frec$, 116, 1)
  Else
    '--- ertt
    tkod$ = Mid$(frec$, 28, 15)
    If irany$ = "B" Then
      menny@ = Abs(xval(Mid$(frec$, 67, 12)))
    Else
      menny@ = -(Abs(xval(Mid$(frec$, 67, 12))))
    End If
    fogl@ = 0
    bear@ = xval(Mid$(frec$, 79, 12))
    raktarkod$ = Mid$(frec$, 12, 4)
    trec$ = dbxkey("KTRM", tkod$)
    If Trim(Mid$(trec$, 1067, 15)) <> "" Then tapgo$ = "G"
  End If
  trec$ = dbxkey("KTRM", tkod$)
  xcikarak@ = xval(Mid$(trec$, 552, 14))
  xcikjell$ = Mid$(trec$, 442, 1)
  xcrakj$ = Mid$(trec$, 907, 1)
  xckommi@ = xval(Mid$(trec$, 983, 14))
  xckeszl@ = xval(Mid$(trec$, 748, 14))
  xcfogl@ = xval(Mid$(trec$, 762, 14))
  xcikmenny@ = xval(Mid$(trec$, 748, 14))
  xckozp@ = xval(Mid$(trec$, 927, 314))
  xcvegyi@ = xval(Mid$(trec$, 941, 14))
  xczolds@ = xval(Mid$(trec$, 1001, 12))
  xccplc@ = xval(Mid$(trec$, 955, 14))
  xcgongy@ = xval(Mid$(trec$, 969, 14))
  If merlmod$ = "K" Then
    '--- átlagár számítása
    ujmenny@ = xcikmenny@ + menny@
    ujert@ = xcikmenny@ * xcikarak@ + menny@ * bear@
    If ujmenny@ <> 0 Then ujar@ = ujert@ / ujmenny@ Else ujar@ = 0
    '--- ár számított
    xcikarak@ = ujar@
    '--- mennyiség
    xcikmenny@ = ujmenny@
  Else
    ujmenny@ = xcikmenny@ + menny@
    '--- ár berakott
    ujar@ = xcikarak@
    '--- mennyiség
    xcikmenny@ = ujmenny@
    '--- frec visszírása
    If tvagyf$ = "T" Then Mid$(frec$, 59, 12) = ertszam(Str(ujar@), 12, 2) Else Mid$(frec$, 79, 12) = ertszam(Str(ujar@), 12, 2)
  End If
  '--- az aktuális készletbe minden telephelyi készlet
  xckeszl@ = xckeszl@ + menny@
  xcfogl@ = xcfogl@ + fogl@
  rkkd$ = xcrakj$
  If rkkd$ = "I" Or rkkd$ = "E" Or rkkd$ = "D" Then rkkd$ = "K"
  If Mid$(ikonfrec$, 300, 4) = Mid$(ikonfrec$, 280, 4) Then
    '--- lerakat, minden áru és gyöngyöleg
    '--- a központi készlet és a c+c készlet ugyanaz
    xckozp@ = xckozp@ + menny@
    xckommi@ = xckommi@ + menny@
    xccplc@ = xccplc@ + menny@
  Else
    '--- központ
    If raktarkod$ = Mid$(ikonfrec$, 300, 4) Then
      '--- központi c+c raktár
      If rkkd$ = "V" Then
        '--- Vegyiáru
        xcvegyi@ = xcvegyi@ + menny@
        xckommi@ = xckommi@ + menny@
      Else
        '--- Egyéb áru, és göngyöleg
        xccplc@ = xccplc@ + menny@
      End If
    Else
      '--- központi készletek a helyükre mennek
      If tvagyf$ = "F" Then
        If zstatusz$ = "R" And zszcim$ = "00010009" Then
          '--- átadás C+C-nek bevételezés
          xccplc@ = xccplc@ - menny@
        End If
      End If
      If rkkd$ = "K" Or rkkd$ = "V" Then
        '--- komissiózható készlet
        xckommi@ = xckommi@ + menny@
      End If
      Select Case rkkd$
        Case "K"
          '--- minden áru, kivéve vegyi és üres göngyöleg
          xckozp@ = xckozp@ + menny@
        Case "V"
          '--- Vegyi áru
          xcvegyi@ = xcvegyi@ + menny@
        Case "Z"
          '--- zöldség
          xczolds@ = xczolds@ + menny@
        Case "G"
          '--- Üres göngyöleg
          If raktarkod$ = Mid$(ikonfrec$, 280, 4) Then
            '--- központi raktár
            xckommi@ = xckommi@ + menny@
            xckozp@ = xckozp@ + menny@
          Else
            '--- göngyöleg raktár
            xcgongy@ = xcgongy@ + menny@
          End If
        Case Else
      End Select
    End If
  End If
  Mid$(trec$, 552, 14) = ertszam(Str$(xcikarak@), 14, 2)
  '--- aktuális készlet
  If Mid$(ikonfrec$, 300, 4) = Mid$(ikonfrec$, 280, 4) Then
    ycakt@ = xckozp@
  Else
    ycakt@ = xckozp@ + xcvegyi@ + xczolds@ + xccplc@ + xcgongy@
  End If
  Mid$(trec$, 748, 14) = ertszam(Str$(ycakt@), 14, 3)
  'Mid$(trec$, 748, 14) = ertszam(Str$(xcikmenny@), 14, 3)
  '--- foglalt készlet
  Mid$(trec$, 762, 14) = ertszam(Str$(xcfogl@), 14, 3)
  '--- komissiózható készlet
  Mid$(trec$, 983, 14) = ertszam(Str$(xckommi@), 14, 3)
  '--- központi készlet
  Mid$(trec$, 927, 14) = ertszam(Str$(xckozp@), 14, 3)
  '--- vegyiáru készlet
  Mid$(trec$, 941, 14) = ertszam(Str$(xcvegyi@), 14, 3)
  '--- zoldseg készlet
  Mid$(trec$, 1001, 12) = ertszam(Str$(xczolds@), 12, 2)
  '--- c+c készlet
  Mid$(trec$, 955, 14) = ertszam(Str$(xccplc@), 14, 3)
  '--- göngyöleg készlet
  Mid$(trec$, 969, 14) = ertszam(Str$(xcgongy@), 14, 3)
  '--- ktrm visszaírása
  If kulsobolt = 1 Then
    If irany$ = "B" And merlmod$ = "K" And tvagyf$ = "T" Then
      Mid$(trec$, 1274, 14) = ertszam(Str(bear@), 14, 2)
      Mid$(trec$, 1025, 12) = ertszam(Str(bear@), 12, 2)
      Mid$(trec$, 566, 14) = ertszam(Str(bear@), 14, 2)
      Mid$(trec$, 552, 14) = ertszam(Str(ujar@), 14, 2)
    End If
  End If
  Call dbxki("KTRM", trec$, ";", "", "", hiba%)
'--- gyûjtõ göngyöleg kezelése
'--- tapadó göngyöleg kezelése
  If Trim(Mid$(trec$, 1067, 15)) <> "" Then
    xctg$ = Mid$(trec$, 1067, 15)
    gonrec$ = dbxkey("KTRM", xctg$)
  End If
  If gonrec$ <> "" Then
    '--- tapadó göngyöleg van
    xgurm@ = xval(Mid$(gonrec$, 1297, 4))
    If tapgo$ = "G" Then
      '--- göngyöleg
      If xgurm@ = 0 Then
        gmenny@ = menny@
      Else
        gmenny@ = menny@ / xgurm@
      End If
      gcikmenny@ = xval(Mid$(gonrec$, 748, 14))
      gcikarak@ = xval(Mid$(gonrec$, 552, 14))
      If merlmod$ = "K" Then
        'gear@ = xval(Mid$(frec$, 117, 10))
        'ujgmenny@ = gcikmenny@ + gmenny@
        'ujgert@ = gcikmenny@ * gcikarak@ + gmenny@ * gear@
        'If ujgmenny@ <> 0 Then ujgar@ = ujgert@ / ujgmenny@ Else ujgar@ = 0
        '--- göngyöleg ár számított
        'Mid$(gonrec$, 748, 14) = ertszam(Str(ujgmenny@), 14, 2)
        'Mid$(gonrec$, 552, 14) = ertszam(Str(ujgar@), 14, 2)
      Else
        'ujgmenny@ = gcikmenny@ + gmenny@
        'Mid$(gonrec$, 748, 14) = ertszam(Str(ujgmenny@), 14, 2)
        'Mid$(frec$, 117, 10) = ertszam(Str(gcikarak@), 10, 1)
        'If tvagyf$ = "T" Then Mid$(frec$, 117, 10) = ertszam(Str(ujar@), 12, 2) Else Mid$(frec$, 113, 8) = ertszam(Str(ujar@), 8, 1)
      End If
      'Call dbxki("KTRM", gonrec$, ";", "", "", hiba%)
    End If
  End If
  termrec$ = dbxkey("KTRM", tkod$)
End Sub
Public Sub keszletvalt(irany$, trec$, merlmod$)
  '--- kereskedelmi készlet változása
  '--- irany= B-bevét K-kiadás
  '--- trec$ a készletforgalmi rekord
  tkod$ = Mid$(trec$, 36, 15)
  rak$ = Mid(trec$, 24, 4)
  If telkom.Ltvan = 1 Then
    If rak$ <> telkom.Lraktar Then Exit Sub
    
  End If
  moz$ = Mid$(trec$, 21, 3)
  thely$ = Mid$(trec$, 28, 8)
  mint$ = Mid$(trec$, 51, 8)
  bear@ = xval(Mid$(trec$, 59, 12))
  utbear@ = bear@
  menny@ = xval(Mid(trec$, 71, 12))
  fmenny@ = xval(Mid$(trec$, 83, 12))
  If Mid$(trec$, 116, 1) = "G" Then
    zarmenny@ = 0
  Else
    zarmenny@ = xval(Mid$(trec$, 116, 12))
  End If
  '--- terméktörzs és mérlegelt átlagár
  termrec$ = dbxkey("KTRM", tkod$)
  rakrec$ = dbxkey("KRAK", rak$)
  If termrec$ <> "" And rakrec$ <> "" Then
    km@ = xval(Mid$(termrec$, 748, 14))
    kf@ = xval(Mid$(termrec$, 762, 14))
    nyar@ = xval(Mid$(termrec$, 552, 14))
    If (Mid$(termrec$, 442, 1) = "A" Or Mid$(termrec$, 442, 1) = "R") And merlmod$ = "K" Then
      '--- átlagár
      regiert@ = km@ * nyar@
      If menny@ < 0 Then
        fert@ = -menny@ * bear@
        ujert@ = regiert@ - fert@
      Else
        fert@ = menny@ * bear@
        ujert@ = regiert@ + fert@
      End If
      ujmenny@ = km@ + menny@
      If ujmenny@ > 0 Then
        ujar@ = ujert@ / ujmenny@
        Mid$(termrec$, 552, 14) = ertszam(Str(ujar), 14, 2)
      End If
      Mid$(termrec$, 566, 14) = ertszam(Str(utbear@), 14, 2)
    End If
    If telkom.Ltvan = 1 Then
      Call hozzad(termrec$, 748, 14, menny@, 3)
      Select Case telkom.Ltelindex
        Case 1: Call hozzad(termrec$, 983, 14, menny@, 3)
        Case 2: Call hozzad(termrec$, 927, 14, menny@, 3)
        Case 3: Call hozzad(termrec$, 941, 14, menny@, 3)
        Case 4: Call hozzad(termrec$, 955, 14, menny@, 3)
        Case Else
      End Select
    Else
      If komiraktdb > 0 Then
        If komirak(rak$) = 1 Then
          Call hozzad(termrec$, 748, 14, menny@, 3)
          Call hozzad(termrec$, 762, 14, fmenny@, 3)
        End If
      Else
        Call hozzad(termrec$, 748, 14, menny@, 3)
        Call hozzad(termrec$, 762, 14, fmenny@, 3)
      End If
    End If
    If xval(Mid$(termrec$, 762, 14)) < 0 Then Mid$(termrec$, 762, 14) = Space$(14)
    Call dbxki("KTRM", termrec$, ";", "", "", hiba%)
    If telkom.Ltvan = 1 Then Exit Sub
    '--- raktárkészlet
    azon$ = rak$ + tkod$
    rkszrec$ = dbxkey("RKSZ", azon$)
    If rkszrec$ = "" Then
      rkszrec$ = azon$ + Space$(8981)
      Call dbxki("RKSZ", rkszrec$, ";", "U", "", hiba%)
      '--- rksz felfuzese cikk lancra
      w1% = obsorszama("RKSZ")
      aktucim& = OBJTAB(w1%).obcim
      Call lancra("AUWKER", "RKSZCIK", termrec$, aktucim&, rkszrec$)
      w1% = obsorszama("RKSZ")
      aktucim& = OBJTAB(w1%).obcim
      Call lancra("AUWKER", "RKSZRAK", rakrec$, aktucim&, rkszrec$)
    End If
    Call hozzad(rkszrec$, 20, 12, menny@, 3)
    Call hozzad(rkszrec$, 32, 12, fmenny@, 3)
    If xval(Mid$(rkszrec$, 32, 12)) < 0 Then Mid$(rkszrec$, 32, 12) = Space$(12)
    Call hozzad(rkszrec$, 91, 12, zarmenny@, 3)
    If xval(Mid$(rkszrec$, 91, 12)) < 0 Then Mid$(rkszrec$, 91, 12) = Space$(12)
    '--- mintaszám kezelése
    If mint$ <> Space$(8) Then
      idxx% = 0
      For i371& = 1 To 300
        ele2$ = Mid(rkszrec$, (i371& - 1) * 28 + 200, 28)
        If Mid$(ele2$, 1, 8) = mint$ Then
          Call hozzad(ele2$, 9, 10, menny@, 2)
          Call hozzad(ele2$, 19, 10, fmenny@, 2)
          Mid$(rkszrec$, (i371& - 1) * 28 + 200, 28) = ele2$
          idxx% = i371&: Exit For
        Else
          If ele2$ = Space$(28) Then
            Mid$(ele2$, 1, 8) = mint$
            Call hozzad(ele2$, 9, 10, menny@, 2)
            Call hozzad(ele2$, 19, 10, fmenny@, 2)
            If xval(Mid$(ele2$, 19, 10)) < 0 Then Mid$(ele2$, 19, 10) = Space$(10)
            Mid$(rkszrec$, (i371& - 1) * 28 + 200, 28) = ele2$
            idxx% = i371&: Exit For
          End If
        End If
      Next
      If idxx% <> 0 Then
        If xval(Mid$(ele2$, 9, 10)) = 0 Then
          If idxx% < 300 Then
            For i371& = idxx% To 299
              Mid$(rkszrec$, (i371& - 1) * 28 + 200, 28) = Mid$(rkszrec$, i371& * 28 + 200, 28)
            Next
          End If
          Mid$(rkszrec$, (300 - 1) * 28 + 200, 28) = Space$(28)
        End If
      End If
    End If
    Call dbxki("RKSZ", rkszrec$, ";", "", "", hiba%)
    '--- tárhelyes készlet kezelése
    If thely$ <> Space$(8) Then
      For j1& = 1 To 10
        azon$ = tkod$ + mint$ + rak$ + Right$("00" + Trim(Str(j1&)), 2)
        tcrec$ = dbxkey("TRCK", azon$)
        If tcrec$ = "" Then
          tcrec$ = azon$ + Space$(3471)
          Call dbxki("TRCK", tcrec$, ";", "U", "", hiba%)
          w1% = obsorszama("TRCK")
          aktucim& = OBJTAB(w1%).obcim
          Call lancra("AUWKER", "TRCKCIK", termrec$, aktucim&, tcrec$)
        End If
        ReDim Preserve thtomb(j1& * 100)
        For i1& = 1 To 100
          ele2$ = Mid$(tcrec$, (i1& - 1) * 32 + 200, 32)
          thtomb$((j1& - 1) * 100 + i1&) = ele2$
          If ele2$ = Space$(32) Then vanures% = i1&: Exit For
        Next
        If vanures% > 0 Then berec% = j1&: Exit For
      Next
      talal% = 0
      darab& = vanures% - 1
      If vanures% > 1 Then
        For i1& = 1 To vanures% - 1
          If Mid$(thtomb$(i1&), 1, 8) = thely$ Then
            talal% = i1&: Exit For
          End If
        Next
      End If
      If talal% = 0 Then talal% = vanures%: darab& = darab& + 1
      ele2$ = thtomb$(talal%)
      Mid$(ele2$, 1, 8) = thely$
      Call hozzad(ele2$, 9, 12, menny@, 3)
      Call hozzad(ele2$, 21, 12, zarmenny@, 3)
      If xval(Mid$(ele2$, 21, 12)) < 0 Then Mid$(ele2$, 21, 12) = Space$(12)
      thtomb$(talal%) = ele2$
      If xval(Mid$(ele2$, 9, 12)) = 0 Then
        '--- tömöríteni
        If talal% < darab& Then
          For i1& = talal% To darab& - 1
            thtomb$(i1&) = thtomb$(i1& + 1)
          Next
        End If
        thtomb$(darab&) = Space$(32)
      End If
      '--- tömörített kiírása
      For j1& = 1 To berec%
        azon$ = tkod$ + mint$ + rak$ + Right$("00" + Trim(Str(j1&)), 2)
        tcrec$ = dbxkey("TRCK", azon$)
        Mid$(tcrec$, 200, 3301) = Space$(3301)
        For i1& = 1 To 100
          idxx% = (j1& - 1) * 100 + i1&
          If idxx% > darab& Then Exit For
          ele2$ = thtomb$(idxx%)
          Mid$(tcrec$, (i1& - 1) * 32 + 200, 32) = ele2$
          If idxx% = darab& Then Exit For
        Next
        Call dbxki("TRCK", tcrec$, ";", "", "", hiba%)
      Next
    End If
  End If
End Sub

Public Sub eszkozertek(eszkaz$, dat$, brertek@, snertek@, anertek@, stsecs@, atsecs@, fbrertek@, fnertek@, ftsecs@)
  Dim ecst@()
  '--- (nap,1) bruttó ertekváltozás az adott napon
  '--- (nap,2) nettó ertekváltozás az adott napon számv.
  '--- (nap,3) nettó ertekváltozás az adott napon adótrv.
  '--- (nap,4) napi tervszerinti ÉCS számv
  '--- (nap,5) napi tervszerinti ÉCS adótrv
  '--- (nap,6) fejlesztési tartalékra esõ a 3-asból 08.05.07
  '--- (nap,7) fejlesztési tartalékra esõ az 5-ösbõl 08.05.07
  '--- (nap,8) fejlesztési tartalékra esõ az 1-esbõl
  '--- az adott eszköz adott napi záró értékének meghatározása
  '--- brertek tárgynapi záró bruttó érték
  '--- fbrertek fejlesztési tartalék tárgynapi bruttó értéke
  '--- snertek tárgynapi nettó érték számv
  '--- anertek tárgynapi nettó érték adótrv.
  '--- fnertek fejl.tartalék tárgynapi nettó értéke
  '--- stsecs  tárgynapig elszámolható értékcsökkenés az utolsó feladás óta számv
  '--- atsecs  tárgynapig elszámolható értékcsökkenés az utolsó feladás óta adótörv.
  '--- ftcses  tárgynapig elszámolható értékcsökkenés az utolsó feladás óta fejl.tartalékból
  eszkrec$ = dbxkey("ESZK", eszkaz$)
  induldat$ = Mid$(eszkrec$, 742, 6)
  '--- utolsó feladás meghatározása
  utfelad$ = Mid$(irec$, 415, 6)
  If induldat$ <= utfelad$ Then induldat$ = novdat(utfelad$)
  If induldat$ > dat$ Then
     '--- 08.05.07
     brertek@ = 0: snertek@ = 0: anertek@ = 0: stsecs@ = 0: saecs@ = 0: sfecs@ = 0
     Exit Sub
  End If
  vcim& = xval(Mid(eszkrec$, 720, 10))
  nbrutto@ = xval(Mid$(eszkrec$, 661, 14))
  nsnetto@ = nbrutto@ - xval(Mid$(eszkrec$, 692, 14))
  '--- 08.05.07
  nanetto@ = nbrutto@ - xval(Mid$(eszkrec$, 706, 14))
  nfbrutto@ = xval(Mid$(eszkrec$, 750, 14))
  nfnetto@ = xval(Mid$(eszkrec$, 750, 14)) - xval(Mid$(eszkrec$, 764, 14))
  napindex% = napkul(induldat$, dat$) + 1
  ReDim ecst@(napindex% + 1, 8)
  '---
  napjai% = napindex%
  fi1 = FreeFile
  Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #fi1
  Do While vcim& > 0
    Seek #fi1, vcim& + 9
    frec$ = Space(200): Get #fi1, , frec$
    mkod$ = Mid$(frec$, 29, 2)
    er1@ = xval(Mid$(frec$, 51, 14))
    er2@ = xval(Mid$(frec$, 65, 14))
    er3@ = xval(Mid$(frec$, 79, 14))
    '--- 08.05.07
    fejlt$ = UCase(Mid$(frec$, 122, 1))
    er4@ = xval(Mid$(frec$, 174, 14))
    er5@ = xval(Mid$(frec$, 188, 12))
    '---
    If Mid$(frec$, 107, 1) <> "S" Then
      If dtm(Mid$(frec$, 23, 6)) <= dtm(utfelad$) Then
        '--- nyitó értékbe
        Select Case mkod$
          Case "BA", "KA"
            '--- aktiválás, nincs teendõ
          Case "FE", "RK"
            '--- felújítás, ráaktiválás
            nbrutto@ = nbrutto@ + er1@
            nsnetto@ = nsnetto@ + er1@
            nanetto@ = nanetto@ + er1@
            '--- 08.05.07
            If fejlt$ = "I" Then
              nfbrutto@ = nfbrutto@ + er1@
              nfnetto@ = nfnetto@ + er1@
            End If
          Case "ER", "AT", "AP", "SE", "KM"
            '--- kivezetés részben, vagy egészben
            nbrutto@ = nbrutto@ - er1@
            nsnetto@ = nsnetto@ - er2@
            nanetto@ = nanetto@ - er3@
            '--- 08.05.07
            nfnetto@ = nfnetto@ - er4@
            nfbrutto@ = nfbrutto@ - er5@
            '---
          Case "TS"
            '--- terv szerinti értékcsökkenés
            nsnetto@ = nsnetto@ - er2@
            nanetto@ = nanetto@ - er3@
            '--- 08.05.07
            nfnetto@ = nfnetto@ - er4@
            '---
          Case "TF", "EV"
            '--- értékvesztés, terven felüli értékcsökkenés
            nsnetto@ = nsnetto@ - er2@
            nanetto@ = nanetto@ - er3@
            '--- 08.05.07
            nfnetto@ = nfnetto@ - er4@
            '---
          Case "VI"
            '--- értékvesztés visszaírása
            nsnetto@ = nsnetto@ + er2@
            nanetto@ = nanetto@ + er3@
            '--- 08.05.07
            nfnetto@ = nfnetto@ + er4@
            '---
          Case Else
        End Select
      Else
        '--- tárgyidõszaki értékbe
        If dtm(Mid$(frec$, 23, 6)) <= dtm(dat$) Then
          napindex% = napkul(induldat$, Mid$(frec$, 23, 6)) + 1
          Select Case mkod$
            Case "BA", "KA"
              '--- aktiválás, nincs teendõ
            Case "FE", "RK"
              '--- felújítás, ráaktiválás
              ecst@(napindex%, 1) = ecst@(napindex%, 1) + er1@
              ecst@(napindex%, 2) = ecst@(napindex%, 2) + er1@
              ecst@(napindex%, 3) = ecst@(napindex%, 3) + er1@
              '--- 08.05.07
              If fejlt$ = "I" Then
                ecst@(napindex%, 8) = ecst@(napindex%, 8) + er1@
                ecst@(napindex%, 6) = ecst@(napindex%, 6) + er1@
              End If
            Case "ER", "AT", "AP", "SE", "KM"
              '--- kivezetés részben, vagy egészben
              ecst@(napindex% + 1, 1) = ecst@(napindex% + 1, 1) - er1@
              ecst@(napindex% + 1, 2) = ecst@(napindex% + 1, 2) - er2@
              ecst@(napindex% + 1, 3) = ecst@(napindex% + 1, 3) - er3@
              '--- 08.05.07
              ecst@(napindex% + 1, 6) = ecst@(napindex% + 1, 6) - er4@
              ecst@(napindex% + 1, 8) = ecst@(napindex% + 1, 8) - er5@
            Case "TS"
              '--- terv szerinti értékcsökkenés
              '--- csak ecstipus=K esetén lehetséges
              ecst@(napindex%, 2) = ecst@(napindex%, 2) - er2@
              ecst@(napindex%, 3) = ecst@(napindex%, 3) - er3@
              '--- 08.05.07
              ecst@(napindex%, 6) = ecst@(napindex%, 6) - er4@
              '---
              ecst@(napindex%, 4) = ecst@(napindex%, 4) + er2@
              ecst@(napindex%, 5) = ecst@(napindex%, 5) + er3@
              '--- 08.05.07
              ecst@(napindex%, 7) = ecst@(napindex%, 7) + er4@
            Case "TF", "EV"
              '--- értékvesztés, terven felüli értékcsökkenés
              ecst@(napindex%, 2) = ecst@(napindex%, 2) - er2@
              ecst@(napindex%, 3) = ecst@(napindex%, 3) - er3@
              '--- 08.05.07
              ecst@(napindex%, 6) = ecst@(napindex%, 6) - er4@
            Case "VI"
              '--- értékvesztés visszaírása
              ecst@(napindex%, 2) = ecst@(napindex%, 2) + er2@
              ecst@(napindex%, 3) = ecst@(napindex%, 3) + er3@
              '--- 08.05.07
              ecst@(napindex%, 6) = ecst@(napindex%, 6) + er4@
            Case Else
          End Select
        End If
      End If
    End If
    vcim& = Val(Mid$(frec$, 126, 10))
  Loop
  If nsnetto@ <> 0 Or nanetto@ <> 0 Then
    '--- ÉCS számítás
    nkod$ = Mid$(eszkrec$, 606, 8)
    nrec$ = dbxkey("ENRM", nkod$)
    If Mid$(nrec$, 39, 1) = "A" Then
      '--- 08.05.07
      stsecs@ = nsnetto@: atsecs@ = nanetto@: ftsecs@ = nfnetto@
      snertek@ = 0: anertek@ = 0: fnertek@ = 0
      GoTo kilepe
    End If
    adecskulcs@ = xval(Mid$(nrec$, 40, 6))
    szecskulcs@ = xval(Mid$(eszkrec$, 615, 6))
    ecstip$ = Mid$(eszkrec$, 614, 1)
    maradv@ = xval(Mid$(eszkrec$, 675, 14))
    evnapjai& = napkul(Mid$(dat$, 1, 2) + "0100", Mid$(dat$, 1, 2) + "1231")
    napibrutto@ = nbrutto@
    napisnetto@ = nsnetto@
    napianetto@ = nanetto@
    '--- 08.05.07
    napifnetto@ = nfnetto@
    napifbrutto@ = nfbrutto@
    stsecs@ = 0
    atsecs@ = 0
    ftsecs@ = 0
    '---
    For i1% = 1 To napjai%
      napibrutto@ = napibrutto@ + ecst@(i1%, 1)
      napisnetto@ = napisnetto@ + ecst@(i1%, 2)
      napianetto@ = napianetto@ + ecst@(i1%, 3)
      '--- 08.05.07
      napifnetto@ = napifnetto@ + ecst@(i1%, 6)
      napifbrutto@ = napifbrutto@ + ecst@(i1%, 8)
      '---
      '--- a napi ecs meghatárzása
      napbtto@ = napibrutto@
      napfbtto@ = napifbrutto@
      Select Case ecstip$
        Case "B"
          napszecs@ = ((napbtto@ * szecskulcs@) / evnapjai&) / 100
          napadecs@ = ((napbtto@ * adecskulcs@) / evnapjai&) / 100
          '--- 08.05.07
          napfdecs@ = ((napfbtto@ * adecskulcs@) / evnapjai&) / 100
        Case "E"
          napszecs@ = (((napbtto@ - maradv@) * szecskulcs@) / evnapjai&) / 100
          napadecs@ = ((napbtto@ * adecskulcs@) / evnapjai&) / 100
          '--- 08.05.07
          napfdecs@ = ((napfbtto@ * adecskulcs@) / evnapjai&) / 100
        Case "I"
          napszecs@ = ((napbtto@ * szecskulcs@) / evnapjai&) / 100
          napadecs@ = napszecs@
          '--- 08.05.07
          napfdecs@ = ((napfbtto@ * adecskulcs@) / evnapjai&) / 100
        Case "-"
          '--- 08.05.07
          napszecs@ = 0: napadecs@ = 0: napfdecs@ = 0
        Case "K"
          napszecs@ = ecst@(i1%, 4)
          napadecs@ = ecst@(i1%, 5)
          '--- 08.05.07
          napfdecs@ = ecst(i1%, 7)
      End Select
      If ecstip$ <> "-" And ecstip$ <> "K" Then
        If napisnetto@ - napszecs@ < maradv@ Then
          napszecs@ = napisnetto@ - maradv@
          If napszecs@ < 0 Then napszecs@ = 0
        End If
        If napianetto@ - napadecs@ < 0@ Then
          napadecs@ = napianetto@
          If napadecs@ < 0 Then napadecs@ = 0
        End If
        '--- 08.05.07
        If napifnetto@ - napfdecs@ < 0@ Then
          napfdecs@ = napifnetto@
          If napfdecs@ < 0 Then napfdecs@ = 0
        End If
        '---
      End If
      napisnetto@ = napisnetto@ - napszecs@
      napianetto@ = napianetto@ - napadecs@
      '--- 08.05.07
      napifnetto@ = napifnetto@ - napfdecs@
      '---
      stsecs@ = stsecs@ + napszecs@
      atsecs@ = atsecs@ + napadecs@
      '--- 08.05.07
      ftsecs@ = ftsecs@ + napfdecs@
      '---
    Next
    '--- kerekítési problémák kezelése
    brertek@ = napibrutto@
    snertek@ = napisnetto@
    anertek@ = napianetto@
    '--- 08.05.07
    fbrertek@ = napifbrutto@
    fnertek@ = napifnetto@
  Else
    '--- 08.05.07
    brertek@ = nbrutto@: snertek@ = nsnetto@: anertek@ = nanetto@: fnertek@ = nfnetto@: fbrertek@ = fbrutto@
    stsecs@ = 0: atsecs@ = 0: ftsecs@ = 0
  End If
kilepe:
  Close fi1
End Sub
Public Sub pszkonyvel(pszrec$, kiegrec$, vsmod$)
  '--- (PVSZ,PSSZ) vevõ és szállító folyószámla rekord könyvelése
  '--- vsmod=V vevõ =S szállító
  Dim arfolyam As Double
  
  If pszrec$ = "" Then Exit Sub
  If vsmod$ = "V" Then
    Call dbxki("PVSZ", pszrec$, ";", "U", "G", hiba%)
    If kiegrec$ <> "" Then
      Mid$(kiegrec$, 1, 7) = Mid$(pszrec$, 1, 7)
      Mid$(kiegrec$, 8, 2) = "UJ"
      Call dbxki("PVSK", kiegrec$, ";", "U", "", hiba%)
    End If
  Else
    Call dbxki("PSSZ", pszrec$, ";", "U", "G", hiba%)
    If kiegrec$ <> "" Then
      Mid$(kiegrec$, 1, 7) = Mid$(pszrec$, 1, 7)
      Mid$(kiegrec$, 8, 2) = "UJ"
      Call dbxki("PSSK", kiegrec$, ";", "U", "", hiba%)
    End If
  End If
  '--- partner egyenlegek
  pkod$ = Mid$(pszrec$, 38, 15)
  partrec$ = dbxkey("PART", pkod$)
  If partrec$ <> "" Then
    forintertek@ = xval(Mid$(pszrec$, 78, 14))
    devnem$ = Mid$(pszrec$, 92, 3)
    devizaertek@ = xval(Mid$(pszrec$, 95, 14))
    If devizaertek@ <> 0 Then arfolyam = forintertek@ / devizaertek@ Else arfolyam = 0
    If vsmod$ = "V" Then
      Call hozzad(partrec$, 631, 14, forintertek@, 2)
      Call hozzad(partrec$, 645, 14, forintertek@, 2)
      Call hozzad(partrec$, 659, 14, forintertek@, 2)
      '--- partner lánc
      w1% = obsorszama("PVSZ")
      aktucim& = OBJTAB(w1%).obcim
      Call lancra("AUWSZAMV", "PVPART", partrec$, aktucim&, pszrec$)
    Else
      w1% = obsorszama("PSSZ")
      aktucim& = OBJTAB(w1%).obcim
      Call lancra("AUWSZAMV", "PSPART", partrec$, aktucim&, pszrec$)
    End If
    '--- kontírozás könyvelése
    irec$ = dbxkey("INST", "INST")
    bizszam$ = novel(irec$, 328, 7)
    Mid$(irec$, 328, 7) = bizszam$
    Call dbxki("INST", irec$, ";", "", "", hiba%)
    For gei% = 1 To 2
      If gei% = 1 Then geip% = 10 Else geip% = 40
      For i99% = 1 To geip%
        If gei% = 2 And i99% = 1 And kiegrec$ = "" Then Exit For
        If gei% = 1 Then
          elem$ = Mid$(pszrec$, (i99% - 1) * 53 + 400, 53)
        Else
          elem$ = Mid$(kiegrec$, (i99% - 1) * 53 + 10, 53)
        End If
        If Trim$(elem$) <> "" Then
          '--- könyvelési tétel elõállítása
          trec$ = Space$(320)
          Mid$(trec$, 8, 7) = bizszam$
          Mid$(trec$, 15, 30) = Mid$(pszrec$, 8, 30)
          Mid$(trec$, 49, 6) = Mid$(pszrec$, 58, 6)
          If vsmod$ = "V" Then
            Mid$(trec$, 45, 4) = "TPV "
          Else
            Mid$(trec$, 45, 4) = "TPS "
          End If
          Mid$(trec$, 55, 6) = maidatum$
          Mid$(trec$, 68, 8) = ugyintezo$
          Mid$(trec$, 84, 16) = Mid$(elem$, 31, 16)
          If vsmod$ = "V" Then
            pfsz$ = Mid$(partrec$, 298, 8)
          Else
            pfsz$ = Mid$(partrec$, 306, 8)
          End If
          elko$ = Mid$(elem$, 15, 16)
          tkosz@ = xval(Mid$(elem$, 1, 14))
          hkosz@ = tkosz@
          If vsmod$ = "V" Then
            If tkosz@ < 0 Then
              tkosz@ = -tkosz@
              Mid$(trec$, 100, 16) = elko$
              Mid$(trec$, 116, 8) = pfsz$
            Else
              Mid$(trec$, 100, 8) = pfsz$
              Mid$(trec$, 116, 16) = elko$
            End If
          Else
            If tkosz@ < 0 Then
              tkosz@ = -tkosz@
              Mid$(trec$, 100, 8) = pfsz$
              Mid$(trec$, 116, 16) = elko$
            Else
              Mid$(trec$, 100, 16) = elko$
              Mid$(trec$, 116, 8) = pfsz$
            End If
          End If
          Mid$(trec$, 132, 14) = Right$(Space$(14) + Format(tkosz@, "##########0.00"), 14)
          Mid$(trec$, 146, 3) = devnem$
          If arfolyam <> 0 Then
            dever@ = tkosz@ / arfolyam
            Mid$(trec$, 149, 14) = Right$(Space$(14) + Format(dever@, "##########0.00"), 14)
          End If
          Mid$(trec$, 163, 15) = Mid$(pszrec$, 38, 15)
          Mid$(trec$, 178, 7) = Mid$(pszrec$, 1, 7)
          Mid$(trec$, 192, 60) = Mid$(pszrec$, 110, 30)
          If vsmod$ = "V" Then
            Mid$(trec$, 277, 4) = "1001"
          Else
            Mid$(trec$, 277, 4) = "2001"
          End If
          '--- rekord felirasa
          Call fktkonyvel(trec$, "U")
          Mid$(elem$, 47, 7) = Mid$(trec$, 1, 7)
          If gei% = 1 Then
            Mid$(pszrec$, (i99% - 1) * 53 + 400, 53) = elem$
          Else
            Mid$(kiegrec$, (i99% - 1) * 53 + 10, 53) = elem$
          End If
        End If
      Next
    Next
    '--- elolegbeszámítás könyvelése
    For i99% = 1 To 5
      elem$ = Mid$(pszrec$, (i99% - 1) * 43 + 1280, 43)
      If Trim$(elem$) <> "" Then
        konyveldat$ = Mid$(pszrec$, 58, 6)
        Call elobekonyvel(elem$, "V", konyveldat$)
        Mid$(pszrec$, (i99% - 1) * 43 + 1280, 43) = elem$
      End If
    Next
    If pszrec$ <> "" Then
      If vsmod$ = "V" Then
        Call dbxki("PVSZ", pszrec$, ";", "", "", hiba%)
        If kiegrec$ <> "" Then Call dbxki("PVSK", kiegrec$, ";", "", "", hiba%)
      Else
        Call dbxki("PSSZ", pszrec$, ";", "", "", hiba%)
        If kiegrec$ <> "" Then Call dbxki("PSSK", kiegrec$, ";", "", "", hiba%)
      End If
    End If
  End If
End Sub

Public Sub pszsztrn(szikt$, vsmod$, sztordat$, sztornoszamlaszam$)
  '--- (PVSZ,PSSZ) vevõ és szállító folyószámlarekord sztornója
  '--- szikt pvsz vagy pssz iktatója
  '--- pvsz,pssz sztornó, elõlegbesz.sztornó, helyesb.sztornó
  '--- pvsz,pssz,pelo(B),pshl szotrnózva és visszaírva
  If vsmod$ = "V" Then
    pszrec$ = dbxkey("PVSZ", szikt$)
    kiegrec$ = dbxkey("PVSK", szikt$)
  Else
    pszrec$ = dbxkey("PSSZ", szikt$)
    kiegrec$ = dbxkey("PSSK", szikt$)
  End If
  If pszrec$ <> "" Then
    '--- pvsz, pssz, sztornója
    Mid$(pszrec$, 166, 1) = "S"
    Mid$(pszrec$, 167, 6) = sztordat$
    Mid$(pszrec$, 181, 8) = ugyintezo$
    Mid$(pszrec$, 220, 15) = sztornoszamlaszam$
    forintertek@ = xval(Mid$(pszrec$, 78, 14))
    If vsmod$ = "V" Then
      Call dbxki("PVSZ", pszrec$, ";", "", "", hiba%)
    Else
      Call dbxki("PSSZ", pszrec$, ";", "", "", hiba%)
    End If
    pkod$ = Mid$(pszrec$, 38, 15)
    partrec$ = dbxkey("PART", pkod$)
    If partrec$ <> "" Then
      If vsmod$ = "V" Then
        Call kivvon(partrec$, 631, 14, forintertek@, 2)
        Call kivvon(partrec$, 645, 14, forintertek@, 2)
        Call kivvon(partrec$, 659, 14, forintertek@, 2)
      End If
      '--- kontirozas sztornója
      For i98% = 1 To 10
        elem$ = Mid$(pszrec$, (i98% - 1) * 53 + 400, 53)
        If Trim$(elem$) <> "" Then
          tikt$ = Mid(elem$, 47, 7)
          Call fktsztrn(tikt$, sztordat$)
        End If
      Next
      If kiegrec$ <> "" Then
        For i98% = 1 To 40
          elem$ = Mid$(kiegrec$, (i98% - 1) * 53 + 10, 53)
          If Trim$(elem$) <> "" Then
            tikt$ = Mid(elem$, 47, 7)
            Call fktsztrn(tikt$, sztordat$)
          End If
        Next
      End If
      '--- beszámított elõlegek sztornója
      For i98% = 1 To 5
        elem$ = Mid$(pszrec$, (i98% - 1) * 43 + 1280, 43)
        If Trim$(elem$) <> "" Then
          Call elobesztrn(elem$, vsmod$, sztordat$)
        End If
      Next
      '--- helyesbítések sztornója
      azon$ = vsmod$ + Mid$(pszrec$, 1, 7)
      herec$ = dbxkey("PSHL", azon$)
      If herec$ <> "" Then
        For i99% = 1 To 3
          hrec$ = Mid$(herec$, (i99% - 1) * 790 + 1, 790)
          If Trim$(Mid$(hrec$, 9, 15)) <> "" Then
            szikt$ = Mid$(pszrec$, 1, 7)
            Call hsztrn(pszrec$, i99%, vsmod$, sztordat$, "")
            Mid$(herec$, (i99% - 1) * 790 + 1, 790) = hrec$
          End If
        Next
        Call dbxki("PSHL", herec$, ";", "", "", hiba%)
      End If
    End If
  End If
End Sub

Public Sub pelokonyvel(pelorec$, vsmod$)
  '--- (PELO) pelorec-ben található elõlegszámla könyvelése
  '--- vsmod$=V vevõ S szállító
  '--- tesztelt!!!
  Dim arfolyam As Double
  If pelorec$ = "" Then Exit Sub
  Call dbxki("PELO", pelorec$, ";", "U", "G", hiba%)
  '--- partner egyenlegek
  pkod$ = Mid$(pelorec$, 23, 15)
  partrec$ = dbxkey("PART", pkod$)
  If partrec$ <> "" Then
    forintertek@ = xval(Mid$(pelorec$, 44, 14))
    devnem$ = Mid$(pelorec$, 58, 3)
    devizaertek@ = xval(Mid$(pelorec$, 61, 14))
    If devizaertek@ <> 0 Then arfolyam = forintertek@ / devizaertek@ Else arfolyam = 0
    If vsmod$ = "V" Then
      Call hozzad(partrec$, 617, 14, forintertek@, 2)
      '--- partner lánc
      w1% = obsorszama("PELO")
      aktucim& = OBJTAB(w1%).obcim
      Call lancra("AUWSZAMV", "PELOPART", partrec$, aktucim&, pelorec$)
    Else
      w1% = obsorszama("PELO")
      aktucim& = OBJTAB(w1%).obcim
      Call lancra("AUWSZAMV", "PELOPART", partrec$, aktucim&, pelorec$)
    End If
    irec$ = dbxkey("INST", "INST")
    bizszam$ = novel(irec$, 328, 7)
    Mid$(irec$, 328, 7) = bizszam$
    Call dbxki("INST", irec$, ";", "", "", hiba%)
    For i99% = 1 To 5
      elem$ = Mid$(pelorec$, (i99% - 1) * 53 + 380, 53)
      If Trim$(elem$) <> "" Then
        '--- könyvelési tétel elõállítása
        trec$ = Space$(320)
        Mid$(trec$, 8, 7) = bizszam$
        Mid$(trec$, 15, 15) = Mid$(pelorec$, 8, 15)
        Mid$(trec$, 49, 6) = Mid$(pelorec$, 38, 6)
        If vsmod$ = "V" Then
          Mid$(trec$, 277, 4) = "1021"
          Mid$(trec$, 45, 4) = "TPEV"
        Else
          Mid$(trec$, 277, 4) = "2021"
          Mid$(trec$, 45, 4) = "TPES"
        End If
        Mid$(trec$, 55, 6) = maidatum$
        Mid$(trec$, 68, 8) = ugyintezo$
        Mid$(trec$, 84, 16) = Mid$(elem$, 31, 16)
        If vsmod$ = "V" Then
          pfsz$ = Mid$(partrec$, 298, 8)
        Else
          pfsz$ = Mid$(partrec$, 306, 8)
        End If
        elko$ = Mid$(elem$, 15, 16)
        tkosz@ = xval(Mid$(elem$, 1, 14))
        hkosz@ = tkosz@
        If vsmod$ = "V" Then
          If tkosz@ < 0 Then
            tkosz@ = -tkosz@
            Mid$(trec$, 100, 16) = elko$
            Mid$(trec$, 116, 8) = pfsz$
          Else
            Mid$(trec$, 100, 8) = pfsz$
            Mid$(trec$, 116, 16) = elko$
          End If
        Else
          If tkosz@ < 0 Then
            tkosz@ = -tkosz@
            Mid$(trec$, 100, 8) = pfsz$
            Mid$(trec$, 116, 16) = elko$
          Else
            Mid$(trec$, 100, 16) = elko$
            Mid$(trec$, 116, 8) = pfsz$
          End If
        End If
        Mid$(trec$, 132, 14) = Right$(Space$(14) + Format(tkosz@, "##########0.00"), 14)
        Mid$(trec$, 146, 3) = devnem$
        If arfolyam <> 0 Then
          dever@ = tkosz@ / arfolyam
          Mid$(trec$, 149, 14) = Right$(Space$(14) + Format(dever@, "##########0.00"), 14)
        End If
        Mid$(trec$, 163, 15) = Mid$(pelorec$, 23, 15)
        Mid$(trec$, 178, 7) = Mid$(pelorec$, 1, 7)
        If vsmod$ = "V" Then
          Mid$(trec$, 192, 60) = langmodul(1)
        Else
          Mid$(trec$, 192, 60) = langmodul(2)
        End If
        '--- rekord felirasa
        Call fktkonyvel(trec$, "U")
        Mid$(elem$, 47, 7) = Mid$(trec$, 1, 7)
        Mid$(pelorec$, (i99% - 1) * 53 + 380, 53) = elem$
      End If
    Next
    Call dbxki("PELO", pelorec$, ";", "", "", hiba%)
  End If
End Sub

Public Sub pelosztorno(pikt$, vsmod$, sztordat$)
  '--- (PELO) pikt$ iktatójú elõlegszámla sztornója
  '--- vsmod=V vagy S
  '--- tesztelt!!!
  pelorec$ = dbxkey("PELO", pikt$)
  If pelorec$ <> "" Then
    vsmod$ = Mid$(pelorec$, 225, 1)
    '--- pvsz, pssz, sztornója
    Mid$(pelorec$, 90, 1) = "S"
    Mid$(pelorec$, 91, 6) = sztordat$
    Mid$(pelorec$, 97, 8) = ugyintezo$
    forintertek@ = xval(Mid$(pelorec$, 44, 14))
    Call dbxki("PELO", pelorec$, ";", "", "", hiba%)
    pkod$ = Mid$(pelorec$, 23, 15)
    partrec$ = dbxkey("PART", pkod$)
    If partrec$ <> "" Then
      If vsmod$ = "V" Then
        Call kivvon(partrec$, 617, 14, forintertek@, 2)
      End If
      '--- kontirozas sztornója
      For i98% = 1 To 5
        elem$ = Mid$(pelorec$, (i98% - 1) * 53 + 380, 53)
        If Trim$(elem$) <> "" Then
          tikt$ = Mid(elem$, 47, 7)
          Call fktsztrn(tikt$, sztordat$)
        End If
      Next
    End If
  End If
End Sub

Public Sub elobekonyvel(psel$, vsmod$, konyveldat$)
  '--- PSEL objektumban megadott elõlegbeszámítás könyvelése
  '--- devizás elõleg esetn psel-ben deviza összeg van
  '--- tesztelt!!!
  elikt$ = Mid$(psel$, 1, 7)
  elrec$ = dbxkey("PELO", elikt$)
  If elrec$ <> "" Then
    el1rec$ = elrec$
    elossz@ = xval(Mid$(elrec$, 44, 14))
    devnem$ = Mid$(elrec$, 58, 3)
    develossz@ = xval(Mid$(elrec$, 61, 14))
    If develossz@ <> 0 Then arfolyam@ = elossz@ / develossz@ Else arfolyam@ = 0
    beossz@ = xval(Mid$(psel$, 23, 14))
    If devnem$ <> "   " And arfolyam@ <> 0 Then
      bedevossz@ = xval(Mid$(psel$, 23, 14))
      beossz@ = bedevossz@ * arfolyam@
      If develossz@ = 0 Then arany@ = 1 Else arany@ = bedevossz@ / develossz@
    Else
      bedevossz@ = 0
      beossz@ = xval(Mid$(psel$, 23, 14))
      If elossz@ = 0 Then arany@ = 1 Else arany@ = beossz@ / elossz@
    End If
    '--- áfa arányosítása
    beossz@ = -beossz@
    bedevossz@ = -bedevossz@
    elso% = 0
    alapo@ = 0: afao@ = 0
    For i98% = 1 To 5
      afel$ = Mid$(elrec$, (i98% - 1) * 30 + 230, 30)
      If Trim$(afel$) <> "" Then
        If elso% = 0 Then elso% = i98%
        alap@ = xval(Mid$(afel$, 3, 14)) * arany@
        afa@ = xval(Mid$(afel$, 17, 14)) * arany@
        afaker% = xval(Mid$(irec$, 345, 1))
        If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
        afa@ = xval(Format(afa@, fst$))
        Mid$(afel$, 3, 14) = Right$(Space$(14) + Format(-alap@, "##########0.00"), 14)
        Mid$(afel$, 17, 14) = Right$(Space$(14) + Format(-afa@, "##########0.00"), 14)
        alapo@ = alapo@ + xval(Mid$(afel$, 3, 14))
        afao@ = afao@ + xval(Mid$(afel$, 17, 14))
        Mid$(elrec$, (i98% - 1) * 30 + 230, 30) = afel$
      End If
    Next
    If beossz@ <> alapo@ + afao@ Then
      kulon@ = beossz@ - (alapo@ + afao@)
      afel$ = Mid$(elrec$, (elso% - 1) * 30 + 230, 30)
      alap@ = xval(Mid$(afel$, 3, 14)) + kulon@
      Mid$(afel$, 3, 14) = Right$(Space$(14) + Format(alap@, "##########0.00"), 14)
      Mid$(elrec$, (elso% - 1) * 30 + 230, 30) = afel$
    End If
    '--- kontírozás arányosítás
    elso% = 0
    konosz@ = 0
    For i98 = 1 To 5
      afel$ = Mid$(elrec$, (i98% - 1) * 53 + 380, 53)
      If Trim$(afel$) <> "" Then
        'If elso% = 0 Then elso% = i98%
        elso% = i98%
        kono@ = xval(Mid$(afel$, 1, 14)) * arany@
        ' Eszi - kerekítés
        afaker% = xval(Mid$(irec$, 345, 1))
        If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
        kono@ = xval(Format(kono@, fst$))
        
        Mid$(afel$, 1, 14) = Right$(Space$(14) + Format(-kono@, "##########0.00"), 14)
        konosz@ = konosz@ + xval(Mid$(afel$, 1, 14))
        Mid$(elrec$, (i98% - 1) * 53 + 380, 53) = afel$
      End If
    Next
    If beossz@ <> konosz@ Then
      kulon@ = beossz@ - konosz@
      afel$ = Mid$(elrec$, (elso% - 1) * 53 + 380, 53)
      alap@ = xval(Mid$(afel$, 1, 14)) + kulon@
      Mid$(afel$, 1, 14) = Right$(Space$(14) + Format(alap@, "##########0.00"), 14)
      Mid$(elrec$, (elso% - 1) * 53 + 380, 53) = afel$
    End If
    '--- beszámítás beírása eredeti elõleg számlába
    If devnem$ <> "   " And arfolyam@ <> 0 Then
      Call hozzad(el1rec$, 134, 14, bedevossz@, 2)
    Else
      Call hozzad(el1rec$, 134, 14, beossz@, 2)
    End If
    Call dbxki("PELO", el1rec$, ";", "", "", hiba%)
    '--- elrec$ kiegészítése és kiírása
    Mid$(elrec$, 1, 7) = Space$(7)
    Mid$(elrec$, 134, 14) = Space(14)
    Mid$(elrec$, 225, 1) = vsmod$
    Mid$(elrec$, 224, 1) = "B"
    Mid$(elrec$, 38, 6) = konyveldat$
    '--- deviza konverzio
    Mid$(elrec$, 44, 14) = Right$(Space$(14) + Format(beossz@, "##########0.00"), 14)
    If devnem$ <> "   " And arfolyam@ <> 0 Then
      Mid$(elrec$, 61, 14) = ertszam(Str$(bedevossz@), 14, 2)
      Mid$(elrec$, 58, 3) = devnem$
    End If
    Mid$(elrec$, 76, 6) = maidatum$
    Mid$(elrec$, 82, 8) = ugyintezo$
    Call dbxki("PELO", elrec$, ";", "U", "G", hiba%)
    Mid$(psel$, 37, 7) = Mid$(elrec$, 1, 7)
    pkod$ = Mid$(elrec$, 23, 15)
    partrec$ = dbxkey("PART", pkod$)
    If partrec$ <> "" Then
      If vsmod$ = "V" Then
        oo@ = beossz@
        Call hozzad(partrec$, 617, 14, oo@, 2)
        Call hozzad(partrec$, 659, 14, oo@, 2)
      End If
      w1% = obsorszama("PELO")
      aktucim& = OBJTAB(w1%).obcim
      Call lancra("AUWSZAMV", "PELOPART", partrec$, aktucim&, elrec$)
      '--- az elrec$ tényleges könyvelése
      '--- könyvelési bizonylaszám kiadása
      irec$ = dbxkey("INST", "INST")
      bizszam$ = novel(irec$, 328, 7)
      Mid$(irec$, 328, 7) = bizszam$
      Call dbxki("INST", irec$, ";", "", "", hiba%)
      For i98% = 1 To 5
        elem$ = Mid$(elrec$, (i98% - 1) * 53 + 380, 53)
        If Trim$(elem$) <> "" Then
          '--- könyvelési tétel elõállítása
          trec$ = Space$(320)
          Mid$(trec$, 8, 7) = bizszam$
          Mid$(trec$, 15, 15) = Mid$(elrec$, 8, 15)
          Mid$(trec$, 49, 6) = Mid$(elrec$, 38, 6)
          If vsmod$ = "V" Then
            Mid$(trec$, 45, 4) = "TPEV"
            Mid$(trec$, 277, 4) = "1021"
          Else
            Mid$(trec$, 277, 4) = "2021"
            Mid$(trec$, 45, 4) = "TPES"
          End If
          Mid$(trec$, 55, 6) = maidatum$
          Mid$(trec$, 68, 8) = ugyintezo$
          Mid$(trec$, 84, 16) = Mid$(elem$, 31, 16)
          If vsmod$ = "V" Then
            pfsz$ = Mid$(partrec$, 298, 8)
          Else
            pfsz$ = Mid$(partrec$, 306, 8)
          End If
          elko$ = Mid$(elem$, 15, 16)
          tkosz@ = xval(Mid$(elem$, 1, 14))
          hkosz@ = tkosz@
          If vsmod$ = "V" Then
            If tkosz@ < 0 Then
              tkosz@ = -tkosz@
              Mid$(trec$, 100, 16) = elko$
              Mid$(trec$, 116, 8) = pfsz$
            Else
              Mid$(trec$, 100, 8) = pfsz$
              Mid$(trec$, 116, 16) = elko$
            End If
          Else
            If tkosz@ < 0 Then
              tkosz@ = -tkosz@
              Mid$(trec$, 100, 8) = pfsz$
              Mid$(trec$, 116, 16) = elko$
            Else
              Mid$(trec$, 100, 16) = elko$
              Mid$(trec$, 116, 8) = pfsz$
            End If
          End If
          Mid$(trec$, 132, 14) = Right$(Space$(14) + Format(tkosz@, "##########0.00"), 14)
          Mid$(trec$, 146, 3) = devnem$
          If arfolyam@ <> 0 Then
            dever@ = tkosz@ / arfolyam@
            Mid$(trec$, 149, 14) = Right$(Space$(14) + Format(dever@, "##########0.00"), 14)
          End If
          Mid$(trec$, 163, 15) = Mid$(elrec$, 23, 15)
          Mid$(trec$, 178, 7) = Mid$(elrec$, 1, 7)
          If vsmod$ = "V" Then
            Mid$(trec$, 192, 60) = "Vevõ elõleg beszámítás"
          Else
            Mid$(trec$, 192, 60) = "Szállító beszámítás"
          End If
          '--- rekord felirasa
          Call fktkonyvel(trec$, "U")
          Mid$(elem$, 47, 7) = Mid$(trec$, 1, 7)
          Mid$(elrec$, (i98% - 1) * 53 + 380, 53) = elem$
        End If
      Next
      '--- pelo visszaírása
      Call dbxki("PELO", elrec$, ";", "", "", hiba%)
    End If
  End If
End Sub

Public Sub elobesztrn(psel$, vsmod$, sztordat$)
  '--- PSEL objektumban megadott elõlegbeszámítás sztornója
  '--- tesztelt!!!
  elikt$ = Mid$(psel$, 37, 7)
  elrec$ = dbxkey("PELO", elikt$)
  beossz@ = xval(Mid$(elrec$, 44, 14))
  devnem$ = Mid$(elrec$, 58, 3)
  bedevossz@ = xval(Mid$(elrec$, 61, 14))
  el1ikt$ = Mid$(psel$, 1, 7)
  el1rec$ = dbxkey("PELO", el1ikt$)
  If el1rec$ <> "" Then
    If bedevossz@ <> 0 Then arfolyam@ = beossz@ / bedevossz@ Else arfolyam@ = 0
    If devnem$ <> "   " And arfolyam@ <> 0 Then
      Call kivvon(el1rec$, 134, 14, bedevossz@, 2)
    Else
      Call kivvon(el1rec$, 134, 14, beossz@, 2)
    End If
    Call dbxki("PELO", el1rec$, ";", "", "", hiba%)
    elrec$ = dbxkey("PELO", elikt$)
    If elrec$ <> "" Then
      Mid$(elrec$, 90, 1) = "S"
      Mid$(elrec$, 91, 6) = sztordat$
      Mid$(elrec$, 97, 8) = ugyintezo$
      Call dbxki("PELO", elrec$, ";", "", "", hiba%)
      pkod$ = Mid$(elrec$, 23, 15)
      partrec$ = dbxkey("PART", pkod$)
      If partrec$ <> "" Then
        If vsmod$ = "V" Then
          oo@ = beossz@
          Call kivvon(partrec$, 617, 14, oo@, 2)
          Call kivvon(partrec$, 659, 14, oo@, 2)
          Call dbxki("PART", partrec$, ";", "", "", hiba%)
        End If
      End If
      For i98% = 1 To 5
        elem$ = Mid$(elrec$, (i98% - 1) * 53 + 380, 53)
        If Trim$(elem$) <> "" Then
          tikt$ = Mid$(elem$, 47, 7)
          Call fktsztrn(tikt$, sztordat$)
        End If
      Next
    End If
  End If
End Sub

Public Sub fktkonyvel(trec$, jel$)
  '--- (FKTE) trec-ben megadott fõkönyvi tétel rekord könyvelése
  '--- jel=U új tétel írssal =V visszaírással
  '--- komplett trec kiegészítve és visszaírva
  '--- tesztelt!!!
  If trec$ = "" Then Exit Sub
  If jel$ = "U" Then
    Call dbxki("FKTE", trec$, ";", "U", "G", hiba%)
  Else
    Call dbxki("FKTE", trec$, ";", "", "", hiba%)
  End If
  '--- fksz egyenlegek halmozása és felfûzés
  oo@ = xval(Mid$(trec$, 132, 14))
  w1% = obsorszama("FKTE")
  aktucim& = OBJTAB(w1%).obcim
  '--- számlatükör egyenleg növelés és karton pointerek felfûzése
  '--- tartozik oldal
  tsy$ = Mid$(trec$, 100, 8)
  fkrec$ = dbxkey("FKSZ", tsy$)
  If fkrec$ <> "" Then
    Call hozzad(fkrec$, 318, 14, oo@, 2)
    Call lancra("AUWSZAMV", "TSZAMLA", fkrec$, aktucim&, trec$)
  End If
  '--- követel oldal
  ksy$ = Mid$(trec$, 116, 8)
  fkrec$ = dbxkey("FKSZ", ksy$)
  If fkrec$ <> "" Then
    Call hozzad(fkrec$, 332, 14, oo@, 2)
    Call lancra("AUWSZAMV", "KSZAMLA", fkrec$, aktucim&, trec$)
  End If
  '--- tartozik költséghely
  tky$ = Mid$(trec$, 108, 8)
  If Trim$(tky$) <> "" Then
    fkrec$ = dbxkey("FKSZ", tky$)
    If fkrec$ <> "" Then
      Call hozzad(fkrec$, 318, 14, oo@, 2)
      Call dbxki("FKSZ", fkrec$, ";", "", "", hiba%)
    End If
    'Call lancra("AUWSZAMV", "TKHELY", fkrec$, aktucim&, trec$)
    knemy$ = Mid$(irec$, 335, 8)
    fkrec$ = dbxkey("FKSZ", knemy$)
    If fkrec$ <> "" Then
      Call hozzad(fkrec$, 332, 14, oo@, 2)
      Call dbxki("FKSZ", fkrec$, ";", "", "", hiba%)
    End If
  End If
  '--- követel költséghely
  kky$ = Mid$(trec$, 124, 8)
  If Trim$(kky$) <> "" Then
    fkrec$ = dbxkey("FKSZ", kky$)
    If fkrec$ <> "" Then
      Call hozzad(fkrec$, 332, 14, oo@, 2)
      Call dbxki("FKSZ", fkrec$, ";", "", "", hiba%)
    End If
    'Call lancra("AUWSZAMV", "KKHELY", fkrec$, aktucim&, trec$)
    knemy$ = Mid$(irec$, 335, 8)
    fkrec$ = dbxkey("FKSZ", knemy$)
    If fkrec$ <> "" Then
      Call hozzad(fkrec$, 318, 14, oo@, 2)
      Call dbxki("FKSZ", fkrec$, ";", "", "", hiba%)
    End If
  End If
End Sub

Public Sub fktsztrn(tikt$, sztordat$)
  '--- (FKTE) tikt iktatójú fõkönyvi tétel rekord sztornója
  '--- tétel sztornó, halmozás (kivvon)
  '--- tesztelt!!!
  trec$ = dbxkey("FKTE", tikt$)
  If trec$ <> "" Then
    Mid$(trec$, 61, 1) = "S"
    Mid$(trec$, 62, 6) = sztordat$
    Mid$(trec$, 76, 8) = ugyintezo$
    Call dbxki("FKTE", trec$, ";", "", "", hiba%)
    '--- fksz egyenlegek halmozása és felfûzés
    oo@ = xval(Mid$(trec$, 132, 14))
    '--- tartozik oldal
    tsy$ = Mid$(trec$, 100, 8)
    fkrec$ = dbxkey("FKSZ", tsy$)
    If fkrec$ <> "" Then
      Call kivvon(fkrec$, 318, 14, oo@, 2)
      Call dbxki("FKSZ", fkrec$, ";", "", "", hiba%)
    End If
    '--- követel oldal
    ksy$ = Mid$(trec$, 116, 8)
    fkrec$ = dbxkey("FKSZ", ksy$)
    If fkrec$ <> "" Then
      Call kivvon(fkrec$, 332, 14, oo@, 2)
      Call dbxki("FKSZ", fkrec$, ";", "", "", hiba%)
    End If
    '--- tartozik költséghely
    tky$ = Mid$(trec$, 108, 8)
    If Trim$(tky$) <> "" Then
      fkrec$ = dbxkey("FKSZ", tky$)
      If fkrec$ <> "" Then
        Call kivvon(fkrec$, 318, 14, oo@, 2)
        Call dbxki("FKSZ", fkrec$, ";", "", "", hiba%)
      End If
      knemy$ = Mid$(irec$, 335, 8)
      fkrec$ = dbxkey("FKSZ", knemy$)
      If fkrec$ <> "" Then
        Call kivvon(fkrec$, 332, 14, oo@, 2)
        Call dbxki("FKSZ", fkrec$, ";", "", "", hiba%)
      End If
    End If
    '--- követel költséghely
    kky$ = Mid$(trec$, 124, 8)
    If Trim$(kky$) <> "" Then
      fkrec$ = dbxkey("FKSZ", kky$)
      If fkrec$ <> "" Then
        Call kivvon(fkrec$, 332, 14, oo@, 2)
        Call dbxki("FKSZ", fkrec$, ";", "", "", hiba%)
      End If
      knemy$ = Mid$(irec$, 335, 8)
      fkrec$ = dbxkey("FKSZ", knemy$)
      If fkrec$ <> "" Then
        Call kivvon(fkrec$, 318, 14, oo@, 2)
        Call dbxki("FKSZ", fkrec$, ";", "", "", hiba%)
      End If
    End If
  End If
End Sub

Public Sub hkonyvel(pszrec$, hrec$, vsmod$)
  '--- (PSHL) hrec-ben lévõ helyesbítõ könyvelése pszrec számlához
  '--- új számlahelyesbítõ
  '--- hrec tartalma összeállítva
  '--- hrecet újként ír, kontírozást könyvel, hrecet kiegészít és visszaír
  '--- komplett hrec és herec kiegészítve és visszaírva
  hazo$ = vsmod$ + Mid$(pszrec$, 1, 7)
  herec$ = dbxkey("PSHL", hazo$)
  If herec$ = "" Then
    herec$ = Space$(2400)
    Mid$(herec$, 1, 8) = hazo$
    Call dbxki("PSHL", herec$, ";", "U", "", hiba%)
  End If
  tal% = 0
  For i98% = 1 To 3
    h1rec$ = Mid$(hrec$, (i98% - 1) * 790 + 1, 790)
    szlasz$ = Mid$(h1rec$, 9, 15)
    If Trim$(szlasz$) = "" Then tal% = i98%
  Next
  If tal% > 0 Then
    Mid$(hrec$, 1, 8) = Mid$(herec$, 1, 8)
    Mid$(herec$, (tal% - 1) * 790 + 1, 790) = hrec$
    Call dbxki("PSHL", hrec$, ";", "", "", hiba%)
    pkod$ = Mid$(pszrec$, 38, 15)
    partrec$ = dbxkey("PART", pkod$)
    If partrec$ <> "" Then
      forintertek@ = xval(Mid$(hrec$, 36, 14))
      devnem$ = Mid$(hrec$, 50, 3)
      devizaertek@ = xval(Mid$(hrec$, 53, 14))
      If devizaertek@ <> 0 Then arfolyam@ = forintertek@ / devizaertek@ Else arfolyam@ = 0
      If vsmod$ = "V" Then
        Call hozzad(partrec$, 631, 14, forintertek@, 2)
        Call hozzad(partrec$, 645, 14, forintertek@, 2)
        Call hozzad(partrec$, 659, 14, forintertek@, 2)
      End If
      irec$ = dbxkey("INST", "INST")
      bizszam$ = novel(irec$, 328, 7)
      Mid$(irec$, 328, 7) = bizszam$
      Call dbxki("INST", irec$, ";", "", "", hiba%)
      For i99% = 1 To 10
        elem$ = Mid$(hrec$, (i99% - 1) * 53 + 260, 53)
        If Trim$(elem$) <> "" Then
          '--- könyvelési tétel elõállítása
          trec$ = Space$(320)
          Mid$(trec$, 8, 7) = bizszam$
          Mid$(trec$, 15, 15) = Mid$(hrec$, 9, 15)
          Mid$(trec$, 49, 6) = Mid$(hrec$, 30, 6)
          If vsmod$ = "V" Then
            Mid$(trec$, 45, 4) = "TPVH"
            Mid$(trec$, 277, 4) = "1011"
          Else
            Mid$(trec$, 45, 4) = "TPSH"
            Mid$(trec$, 277, 4) = "2011"
          End If
          Mid$(trec$, 55, 6) = maidatum$
          Mid$(trec$, 68, 8) = ugyintezo$
          Mid$(trec$, 84, 16) = Mid$(elem$, 31, 16)
          If vsmod$ = "V" Then
            pfsz$ = Mid$(partrec$, 298, 8)
          Else
            pfsz$ = Mid$(partrec$, 306, 8)
          End If
          elko$ = Mid$(elem$, 15, 16)
          tkosz@ = xval(Mid$(elem$, 1, 14))
          hkosz@ = tkosz@
          If vsmod$ = "V" Then
            If tkosz@ < 0 Then
              tkosz@ = -tkosz@
              Mid$(trec$, 100, 16) = elko$
              Mid$(trec$, 116, 8) = pfsz$
            Else
              Mid$(trec$, 100, 8) = pfsz$
              Mid$(trec$, 116, 16) = elko$
            End If
          Else
            If tkosz@ < 0 Then
              tkosz@ = -tkosz@
              Mid$(trec$, 100, 8) = pfsz$
              Mid$(trec$, 116, 16) = elko$
            Else
              Mid$(trec$, 100, 16) = elko$
              Mid$(trec$, 116, 8) = pfsz$
            End If
          End If
          Mid$(trec$, 132, 14) = Right$(Space$(14) + Format(tkosz@, "##########0.00"), 14)
          Mid$(trec$, 146, 3) = devnem$
          If arfolyam@ <> 0 Then
            dever@ = tkosz@ / arfolyam@
            Mid$(trec$, 149, 14) = Right$(Space$(14) + Format(dever@, "##########0.00"), 14)
          End If
          Mid$(trec$, 163, 15) = Mid$(pszrec$, 38, 15)
          Mid$(trec$, 178, 7) = Mid$(pszrec$, 1, 7)
          Mid$(trec$, 192, 60) = langmodul(3)
          '--- rekord felirasa
          Call fktkonyvel(trec$, "U")
          Mid$(elem$, 47, 7) = Mid$(trec$, 1, 7)
          Mid$(hrec$, (i99% - 1) * 53 + 260, 53) = elem$
        End If
      Next
      Mid$(herec$, (tal% - 1) * 790 + 1, 790) = hrec$
      Call dbxki("PSHL", hrec$, ";", "", "", hiba%)
    End If
  End If
End Sub

Public Sub hsztrn(szikt$, sorszam%, vsmod$, sztordat$, sztornoszamlaszam$)
  '--- (PSHL) szikt iktatójú számla sorszam%-adik helysbítõjének sztornója
  '--- tesztelt!!!
  If vsmod$ = "V" Then
    pszrec$ = dbxkey("PVSZ", szikt$)
  Else
    pszrec$ = dbxkey("PSSZ", szikt$)
  End If
  If pszrec$ <> "" Then
    hazo$ = vsmod$ + Mid$(pszrec$, 1, 7)
    herec$ = dbxkey("PSHL", hazo$)
    If herec$ <> "" Then
      hrec$ = Mid$(herec$, (sorszam% - 1) * 790 + 1, 790)
      Mid$(hrec$, 86, 1) = "S"
      Mid$(hrec$, 87, 6) = sztordat$
      Mid$(hrec$, 93, 8) = ugyintezo$
      forintertek@ = xval(Mid$(hrec$, 36, 14))
      Mid$(herec$, (sorszam% - 1) * 790 + 1, 790) = hrec$
      Call dbxki("PSHL", herec$, ";", "", "", hiba%)
      pkod$ = Mid$(pszrec$, 38, 15)
      partrec$ = dbxkey("PART", pkod$)
      If partrec$ <> "" Then
        If vsmod$ = "V" Then
          Call kivvon(partrec$, 631, 14, forintertek@, 2)
          Call kivvon(partrec$, 645, 14, forintertek@, 2)
          Call kivvon(partrec$, 659, 14, forintertek@, 2)
        End If
      End If
      Call dbxki("PART", partrec$, ";", "", "", hiba%)
      '--- kontirozas sztornója
      For i98% = 1 To 10
        elem$ = Mid$(hrec$, (i98% - 1) * 53 + 260, 53)
        If Trim$(elem$) <> "" Then
          tikt$ = Mid(elem$, 47, 7)
          Call fktsztrn(tikt$, sztordat$)
        End If
      Next
    End If
  End If
End Sub

Public Sub korrkonyvel(krec$, vsmod$)
  '--- (PKOR) krec-ben levõ korrekció könyvelése
  '--- vsmod=V-vevõ S-szállító K-kompenzáció
  '--- tesztelt!!!
  If krec$ = "" Then Exit Sub
  Call dbxki("PKOR", krec$, ";", "U", "G", hiba%)
  forintertek@ = xval(Mid$(krec$, 45, 14))
  devnem$ = Mid$(krec$, 59, 3)
  devizaertek@ = xval(Mid$(krec$, 62, 14))
  If devizaertek@ <> 0 Then arfolyam@ = forintertek@ / devizaertek@ Else arfolyam@ = 0
  pkod$ = Mid$(krec$, 16, 15)
  partrec$ = dbxkey("PART", pkod$)
  If partrec$ <> "" Then
    oo@ = xval(Mid$(krec$, 45, 14))
    If vsmod$ = "V" Or vsmod$ = "K" Then
      Call kivvon(partrec$, 659, 14, oo@, 2)
    End If
    w1% = obsorszama("PKOR")
    aktucim& = OBJTAB(w1%).obcim
    Call lancra("AUWSZAMV", "KRPART", partrec$, aktucim&, krec$)
    '--- kontírozások rögzítés, fksz egyenleg, felfûzés
    irec$ = dbxkey("INST", "INST")
    bizszam$ = novel(irec$, 328, 7)
    Mid$(irec$, 328, 7) = bizszam$
    Call dbxki("INST", irec$, ";", "", "", hiba%)
    For i1% = 1 To 5
      elem$ = Mid$(krec$, (i1% - 1) * 53 + 310, 53)
      If Trim$(elem$) <> "" Then
        '--- könyvelési tétel elõállítása
        trec$ = Space$(320)
        Mid$(trec$, 8, 7) = bizszam$
        Mid$(trec$, 15, 15) = Mid$(krec$, 1, 7)
        Mid$(trec$, 49, 6) = Mid$(krec$, 8, 6)
        Mid$(trec$, 45, 4) = "TPPK"
        Mid$(trec$, 55, 6) = maidatum$
        Mid$(trec$, 68, 8) = ugyintezo$
        Mid$(trec$, 84, 16) = Mid$(elem$, 31, 16)
        If vsmod$ = "V" Or vsmod$ = "K" Then
          pfsz$ = Mid$(partrec$, 298, 8)
          elko$ = Mid$(elem$, 15, 16)
          tkosz@ = xval(Mid$(elem$, 1, 14))
          hkosz@ = tkosz@
          If tkosz@ >= 0 Then
            Mid$(trec$, 100, 16) = elko$
            Mid$(trec$, 116, 8) = pfsz$
          Else
            tkosz@ = -tkosz@
            Mid$(trec$, 100, 8) = pfsz$
            Mid$(trec$, 116, 16) = elko$
          End If
          Mid$(trec$, 146, 3) = devnem$
          If arfolyam@ <> 0 Then
            dever@ = tkosz@ / arfolyam@
            Mid$(trec$, 149, 14) = Right$(Space$(14) + Format(dever@, "##########0.00"), 14)
          Else
            dever@ = 0
          End If
          Mid$(trec$, 132, 14) = Right$(Space$(14) + Format(tkosz@, "##########0.00"), 14)
          Mid$(trec$, 163, 15) = Mid$(krec$, 16, 15)
          Mid$(trec$, 178, 7) = Mid$(krec$, 31, 7)
          Mid$(trec$, 277, 4) = "1153"
          If vsmod$ = "V" Then
            Mid$(trec$, 192, 60) = langmodul(4)
          Else
            Mid$(trec$, 192, 60) = langmodul(5)
          End If
          '--- rekord felirasa
          Call fktkonyvel(trec$, "U")
          Mid$(elem$, 47, 7) = Mid$(trec$, 1, 7)
          '--- iktató pkor-ba
          Mid$(krec$, (i1% - 1) * 53 + 310, 53) = elem$
          '--- kiegyenlítés felvezetése pvsz-re
          vsikt$ = Mid$(krec$, 31, 7)
          vsrec$ = dbxkey("PVSZ", vsikt$)
          If vsrec$ <> "" Then
            tal% = 0
            For i11% = 1 To 10
              elem2$ = Mid$(vsrec$, (i11% - 1) * 35 + 930, 35)
              If Trim$(elem2$) = "" Then tal% = i11%: Exit For
            Next
            If tal% > 0 Then
              elem2$ = Space$(35)
              Mid$(elem2$, 1, 14) = Mid$(elem$, 1, 14)
              Mid$(elem2$, 1, 14) = Mid$(elem$, 1, 14)
              Mid$(elem2$, 15, 8) = Mid$(elem$, 15, 8)
              Mid$(elem2$, 23, 6) = Mid$(krec$, 8, 6)
              Mid$(elem2$, 29, 7) = Mid$(trec$, 1, 7)
              Mid$(vsrec$, (tal% - 1) * 35 + 930, 35) = elem2$
              Call dbxki("PVSZ", vsrec$, ";", "", "", hiba%)
            Else
              vkiegrec$ = dbxkey("PVSK", vsikt$)
              If vkiegrec$ = "" Then
                vkiegrec$ = Space(3000)
                Mid$(vkiegrec$, 8, 2) = "UJ"
                Mid$(vkiegrec$, 1, 7) = vsikt$
                Call dbxki("PVSK", vkiegrec$, ";", "U", "", hiba%)
                vkiegrec$ = dbxkey("PVSK", vsikt$)
              End If
              For i11% = 1 To 20
                elem2$ = Mid$(vkiegrec$, (i11% - 1) * 35 + 2130, 35)
                If Trim$(elem2$) = "" Then tal% = i11%: Exit For
              Next
              If tal% > 0 Then
                elem2$ = Space$(35)
                Mid$(elem2$, 1, 14) = Mid$(elem$, 1, 14)
                Mid$(elem2$, 1, 14) = Mid$(elem$, 1, 14)
                Mid$(elem2$, 15, 8) = Mid$(elem$, 15, 8)
                Mid$(elem2$, 23, 6) = Mid$(krec$, 8, 6)
                Mid$(elem2$, 29, 7) = Mid$(trec$, 1, 7)
                Mid$(vkiegrec$, (tal% - 1) * 35 + 2130, 35) = elem2$
                Mid$(vkiegrec$, 8, 2) = "UJ"
                Call dbxki("PVSK", vkiegrec$, ";", "", "", hiba%)
              End If
            End If
          End If
        End If
        If vsmod$ = "S" Or vsmod$ = "K" Then
          pfsz$ = Mid$(partrec$, 306, 8)
          elko$ = Mid$(elem$, 15, 16)
          tkosz@ = xval(Mid$(elem$, 1, 14))
          hkosz@ = tkosz@
          If tkosz@ < 0 Then
            tkosz@ = -tkosz@
            Mid$(trec$, 100, 16) = elko$
            Mid$(trec$, 116, 8) = pfsz$
          Else
            Mid$(trec$, 100, 8) = pfsz$
            Mid$(trec$, 116, 16) = elko$
          End If
          Mid$(trec$, 146, 3) = devnem$
          If arfolyam@ <> 0 Then
            dever@ = tkosz@ / arfolyam@
            Mid$(trec$, 149, 14) = Right$(Space$(14) + Format(dever@, "##########0.00"), 14)
          Else
            dever@ = 0
          End If
          Mid$(trec$, 132, 14) = Right$(Space$(14) + Format(tkosz@, "##########0.00"), 14)
          Mid$(trec$, 163, 15) = Mid$(krec$, 16, 15)
          Mid$(trec$, 178, 7) = Mid$(krec$, 38, 7)
          If vsmod$ = "S" Then
            Mid$(trec$, 192, 60) = langmodul(6)
          Else
            Mid$(trec$, 192, 60) = langmodul(5)
          End If
          '--- rekord felirasa
          Mid$(trec$, 277, 4) = "2153"
          Call fktkonyvel(trec$, "U")
          If vsmod$ = "S" Then
            Mid$(elem$, 47, 7) = Mid$(trec$, 1, 7)
            '--- iktató pkor-ba
            Mid$(krec$, (i1% - 1) * 53 + 310, 53) = elem$
          End If
          '--- kiegyenlítés felvezetése pssz-re
          vsikt$ = Mid$(krec$, 38, 7)
          vsrec$ = dbxkey("PSSZ", vsikt$)
          If vsrec$ <> "" Then
            tal% = 0
            For i11% = 1 To 10
              elem2$ = Mid$(vsrec$, (i11% - 1) * 35 + 930, 35)
              If Trim$(elem2$) = "" Then tal% = i11%: Exit For
            Next
            If tal% > 0 Then
              elem2$ = Space$(35)
              Mid$(elem2$, 1, 14) = Mid$(elem$, 1, 14)
              Mid$(elem2$, 15, 8) = Mid$(elem$, 15, 8)
              Mid$(elem2$, 23, 6) = Mid$(krec$, 8, 6)
              Mid$(elem2$, 29, 7) = Mid$(trec$, 1, 7)
              Mid$(vsrec$, (tal% - 1) * 35 + 930, 35) = elem2$
              Call dbxki("PSSZ", vsrec$, ";", "", "", hiba%)
            Else
              skiegrec$ = dbxkey("PSSK", vsikt$)
              If skiegrec$ <> "" Then
                skiegrec$ = Space(3000)
                Mid$(skiegrec$, 8, 2) = "UJ"
                Mid$(skiegrec$, 1, 7) = vsikt$
                Call dbxki("PSSK", skiegrec$, ";", "U", "", hiba%)
                skiegrec$ = dbxkey("PSSK", vsikt$)
              End If
              For i11% = 1 To 20
                elem2$ = Mid$(skiegrec$, (i11% - 1) * 35 + 2130, 35)
                If Trim$(elem2$) = "" Then tal% = i11%: Exit For
              Next
              If tal% > 0 Then
                elem2$ = Space$(35)
                Mid$(elem2$, 1, 14) = Mid$(elem$, 1, 14)
                Mid$(elem2$, 1, 14) = Mid$(elem$, 1, 14)
                Mid$(elem2$, 15, 8) = Mid$(elem$, 15, 8)
                Mid$(elem2$, 23, 6) = Mid$(krec$, 8, 6)
                Mid$(elem2$, 29, 7) = Mid$(trec$, 1, 7)
                Mid$(skiegrec$, (tal% - 1) * 35 + 2130, 35) = elem2$
                Call dbxki("PSSK", skiegrec$, ";", "", "", hiba%)
              End If
            End If
          End If
        End If
      End If
    Next
    '--- pelo visszaírása
    Call dbxki("PKOR", krec$, ";", "", "", hiba%)
  End If
End Sub

Public Sub korrsztorno(kikt$, vsmod$, sztordat$)
  '--- (PKOR) kikt iktatójú korrekció sztornója
  '--- tesztelt!!!
  rec$ = dbxkey("PKOR", kikt$)
  If rec$ <> "" Then
    Mid$(rec$, 120, 1) = "S"
    Mid$(rec$, 121, 6) = sztordat$
    Mid$(rec$, 127, 8) = ugyintezo$
    Call dbxki("PKOR", rec$, ";", "", "", hiba%)
    '--- partner egyenleg
    pkod$ = Mid$(rec$, 16, 15)
    partrec$ = dbxkey("PART", pkod$)
    If partrec$ <> "" Then
      oo@ = xval(Mid$(rec$, 45, 14))
      If vsmod$ = "V" Or vsmod$ = "K" Then
        Call hozzad(partrec$, 659, 14, oo@, 2)
      End If
      Call dbxki("PART", partrec$, ";", "", "", hiba%)
      '--- könyvelési tételek sztornója
      For i1% = 1 To 5
        elem$ = Mid$(rec$, (i1% - 1) * 53 + 310, 53)
        If Trim$(elem$) <> "" Then
          tikt$ = Mid$(elem$, 47, 7)
          Call fktsztrn(tikt$, sztordat$)
          If vsmod$ = "K" Then
            tik1& = Val(tikt$) + 1
            tikt1$ = Right$("0000000" + Trim$(Str$(tik1&)), 7)
            Call fktsztrn(tikt1$, sztordat$)
          End If
          If vsmod$ = "V" Or vsmod$ = "K" Then
            vsikt$ = Mid$(rec$, 31, 7)
            vsrec$ = dbxkey("PVSZ", vsikt$)
            vkiegrec$ = dbxkey("PVSK", vsikt$)
            If vsrec$ <> "" Then
              tal% = 0
              For i11% = 1 To 10
                elem2$ = Mid$(vsrec$, (i11% - 1) * 35 + 930, 35)
                If Mid$(elem2$, 29, 7) = tikt$ Then
                  Mid$(vsrec$, (i11% - 1) * 35 + 930, 35) = Space$(35)
                  Call dbxki("PVSZ", vsrec$, ";", "", "", hiba%)
                  tal% = i11%
                  Exit For
                End If
              Next
              If vkiegrec$ <> "" And tal% = 0 Then
                For i11% = 1 To 20
                  elem2$ = Mid$(vkiegrec$, (i11% - 1) * 35 + 2130, 35)
                  If Mid$(elem2$, 29, 7) = tikt$ Then
                    Mid$(vkiegrec$, (i11% - 1) * 35 + 2130, 35) = Space$(35)
                    Mid$(vkiegrec$, 8, 2) = "UJ"
                    Call dbxki("PVSK", vkiegrec$, ";", "", "", hiba%)
                    tal% = i11%
                    Exit For
                  End If
                Next
              End If
            End If
          End If
          If vsmod$ = "S" Or vsmod$ = "K" Then
            vsikt$ = Mid$(rec$, 38, 7)
            vsrec$ = dbxkey("PSSZ", vsikt$)
            vkiegrec$ = dbxkey("PSSK", vsikt$)
            If vsrec$ <> "" Then
              tal% = 0
              For i11% = 1 To 10
                elem2$ = Mid$(vsrec$, (i11% - 1) * 35 + 930, 35)
                If vsmod$ = "S" And Mid$(elem2$, 29, 7) = tikt$ Or vsmod$ = "K" And Mid$(elem2$, 29, 7) = tikt1$ Then
                  Mid$(vsrec$, (i11% - 1) * 35 + 930, 35) = Space$(35)
                  Call dbxki("PSSZ", vsrec$, ";", "", "", hiba%)
                  tal% = i11%
                  Exit For
                End If
              Next
              If vkiegrec$ <> "" And tal% = 0 Then
                For i11% = 1 To 20
                  elem2$ = Mid$(vkiegrec$, (i11% - 1) * 35 + 2130, 35)
                  If Mid$(elem2$, 29, 7) = tikt$ Then
                     Mid$(vkiegrec$, (i11% - 1) * 35 + 2130, 35) = Space$(35)
                     Call dbxki("PSSK", vkiegrec$, ";", "", "", hiba%)
                     tal% = i11%
                     Exit For
                  End If
                Next
              End If
            End If
          End If
        End If
      Next
    End If
  End If
End Sub

Public Function fktellen(trec$)
  '--- könyvelési tétel ellenõrzése
  szerv$ = Mid$(trec$, 84, 8)
  musz$ = Mid$(trec$, 92, 8)
  tsz$ = Mid$(trec$, 100, 8)
  tkh$ = Mid(trec$, 108, 8)
  ksz$ = Mid$(trec$, 116, 8)
  kkh$ = Mid$(trec$, 124, 8)
  osz@ = xval(Mid$(trec$, 132, 14))
  dev$ = Mid$(trec$, 146, 3)
  devosz@ = xval(Mid$(trec$, 149, 14))
  If Trim(szerv$) <> "" Then
    xxx$ = dbxkey("FSZV", szerv$)
    If xxx$ = "" Then
      fktellen = 1: ellenszov$ = langmodul(7)
      Exit Function
    End If
  End If
  If ikonfrec$ = "" Then
    If Trim(musz$) <> "" Then
      xxx$ = dbxkey("FMSZ", musz$)
      If xxx$ = "" Then
        fktellen = 2: ellenszov$ = langmodul(8)
        Exit Function
      End If
    End If
  End If
  '--- tartozik oldal
  If Trim(tsz$) <> "" Then
    fkrec$ = dbxkey("FKSZ", tsz$)
    If fkrec$ = "" Then
      fktellen = 3: ellenszov$ = langmodul(9): Exit Function
    Else
      '--- gyûjtõ
      If Mid$(fkrec$, 76, 1) = "*" Then fktellen = 4: ellenszov$ = langmodul(10): Exit Function
      '--- könyvelés engedélyezése
      If Mid$(fkrec$, 71, 1) = "L" Then fktellen = 5: ellenszov$ = langmodul(11): Exit Function
      If Mid$(fkrec$, 71, 1) = "V" Then fktellen = 5: ellenszov$ = langmodul(12): Exit Function
      '--- deviza használat
      If Mid$(fkrec$, 72, 1) = "K" And dev$ = "   " Then fktellen = 6: ellenszov$ = langmodul(13): Exit Function
      If Mid$(fkrec$, 72, 1) = "T" And dev$ <> "   " Then fktellen = 6: ellenszov$ = langmodul(14): Exit Function
      '--- költséghely használat
      If Mid$(fkrec$, 73, 1) = "K" And Trim$(tkh$) = "" Then fktellen = 6: ellenszov$ = langmodul(15): Exit Function
      If Mid$(fkrec$, 73, 1) = "T" And Trim$(tkh$) <> "" Then fktellen = 6: ellenszov$ = langmodul(16): Exit Function
      '--- szerv.egység használat
      If Mid$(fkrec$, 74, 1) = "K" And Trim$(szerv$) = "" Then fktellen = 7: ellenszov$ = langmodul(17): Exit Function
      If Mid$(fkrec$, 74, 1) = "T" And Trim$(szerv$) <> "" Then fktellen = 7: ellenszov$ = langmodul(18): Exit Function
      '--- munkaszám használat
      If Mid$(fkrec$, 75, 1) = "K" And Trim$(musz$) = "" Then fktellen = 8: ellenszov$ = langmodul(19): Exit Function
      If Mid$(fkrec$, 75, 1) = "T" And Trim$(musz$) <> "" Then fktellen = 8: ellenszov$ = langmodul(20): Exit Function
    End If
  Else
    fktellen = 9: ellenszov$ = langmodul(21): Exit Function
  End If
  '--- követel oldal
  If Trim(ksz$) <> "" Then
    fkrec$ = dbxkey("FKSZ", ksz$)
    If fkrec$ = "" Then
      fktellen = 3: ellenszov$ = langmodul(22): Exit Function
    Else
      '--- gyûjtõ
      If Mid$(fkrec$, 76, 1) = "*" Then fktellen = 4: ellenszov$ = langmodul(23): Exit Function
      '--- könyvelés engedélyezése
      If Mid$(fkrec$, 71, 1) = "L" Then fktellen = 5: ellenszov$ = langmodul(24): Exit Function
      If Mid$(fkrec$, 71, 1) = "V" Then fktellen = 5: ellenszov$ = langmodul(25): Exit Function
      '--- deviza használat
      If Mid$(fkrec$, 72, 1) = "K" And dev$ = "   " Then fktellen = 6: ellenszov$ = langmodul(26): Exit Function
      If Mid$(fkrec$, 72, 1) = "T" And dev$ <> "   " Then fktellen = 6: ellenszov$ = langmodul(27): Exit Function
      '--- költséghely használat
      If Mid$(fkrec$, 73, 1) = "K" And Trim$(kkh$) = "" Then fktellen = 6: ellenszov$ = langmodul(28): Exit Function
      If Mid$(fkrec$, 73, 1) = "T" And Trim$(kkh$) <> "" Then fktellen = 6: ellenszov$ = langmodul(29): Exit Function
      '--- szerv.egység használat
      If Mid$(fkrec$, 74, 1) = "K" And Trim$(szerv$) = "" Then fktellen = 7: ellenszov$ = langmodul(30): Exit Function
      If Mid$(fkrec$, 74, 1) = "T" And Trim$(szerv$) <> "" Then fktellen = 7: ellenszov$ = langmodul(31): Exit Function
      '--- munkaszám használat
      If Mid$(fkrec$, 75, 1) = "K" And Trim$(musz$) = "" Then fktellen = 8: ellenszov$ = langmodul(32): Exit Function
      If Mid$(fkrec$, 75, 1) = "T" And Trim$(musz$) <> "" Then fktellen = 8: ellenszov$ = langmodul(33): Exit Function
    End If
  Else
    fktellen = 9: ellenszov$ = langmodul(34): Exit Function
  End If
  '--- tartozik költséghely
  If Trim(tkh$) <> "" Then
    fkrec$ = dbxkey("FKSZ", tkh$)
    If fkrec$ = "" Then
      fktellen = 1: ellenszov$ = langmodul(35): Exit Function
    Else
      '--- gyûjtõ
      If Mid$(fkrec$, 76, 1) = "*" Then fktellen = 1: ellenszov$ = langmodul(36): Exit Function
      If Mid$(fkrec$, 69, 1) <> "K" Then fktellen = 1: ellenszöv$ = langmodul(37): Exit Function
    End If
  End If
  If Trim(kkh$) <> "" Then
    fkrec$ = dbxkey("FKSZ", kkh$)
    If fkrec$ = "" Then
      fktellen = 1: ellenszov$ = langmodul(38): Exit Function
    Else
      '--- gyûjtõ
      If Mid$(fkrec$, 76, 1) = "*" Then fktellen = 1: ellenszov$ = langmodul(39): Exit Function
      If Mid$(fkrec$, 69, 1) <> "K" Then fktellen = 1: ellenszöv$ = langmodul(40): Exit Function
    End If
  End If
  If osz@ < 0 Or devosz@ < 0 Then fktellen = 1: ellenszov$ = langmodul(41): Exit Function
  If devosz@ <> 0 And dev$ = "   " Then fktellen = 1: ellenszov$ = langmodul(42): Exit Function
  fktellen = 0
End Function

Public Function fellen(fksz$, khely$, szerv$, munkaszam$, deviban%, vegyes%, jellegek$, tipusok$)
  '--- fõkönyvi szám használatának ellenõrzése
  '--- fksz-a fõkönyvi szám
  '--- khely-a költséghely
  '--- szerv-a szervezeti egység
  '--- munkaszam-a munkaszám
  '--- deviban=1 ha ven hozzá devizanem és érték
  '--- vegyes=1, ha vegyes bizonylatot ellenõrzünk
  '--- jellegek-a megengedett fksz jellegek
  '--- tipusok-a megengedett fksz típusok
  '--- fellen=1 hibás, egyébként fellen=0
  '--- tesztelt!!!
  fkrec$ = dbxkey("FKSZ", fksz$)
  If fkrec$ = "" Then
    Call mess(fksz$ + " " + langmodul(44), 2, 0, langmodul(43), valasz%)
    'MsgBox fksz$ + " " + langmodul(44), 48, langmodul(43)
    fellen = 1: Exit Function
  Else
    '--- gyûjtõszámla ellenõrzése
    If Mid$(fkrec$, 76, 1) = "*" Then
      Call mess(fksz$ + " " + langmodul(45), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(45), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    '--- típusok ellenõrzése
    If InStr(tipusok$, Mid$(fkrec$, 69, 1)) = 0 Then
      Call mess(fksz$ + " " + langmodul(46), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(46), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    '--- jellegek ellenõrzése
    If InStr(jellegek$, Mid$(fkrec$, 70, 1)) = 0 Then
      Call mess(fksz$ + " " + langmodul(47), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(47), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    '--- könyvelés engedélyezésének az ellenõrzése
    If Mid$(fkrec$, 71, 1) = "L" Then
      Call mess(fksz$ + " " + langmodul(48), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(48), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    If Mid$(fkrec$, 71, 1) = "V" And vegyes% = 0 Then
      Call mess(fksz$ + " " + langmodul(49), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(49), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    If Mid$(fkrec$, 71, 1) = "N" And vegyes% = 1 Then
      Call mess(fksz$ + " " + langmodul(50), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(50), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    '--- deviza használat ellenõrzése
    If Mid$(fkrec$, 72, 1) = "K" And deviban% = 0 Then
      Call mess(fksz$ + " " + langmodul(51), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(51), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    If Mid$(fkrec$, 72, 1) = "T" And deviban% = 1 Then
      Call mess(fksz$ + " " + langmodul(52), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(52), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    '--- költséghely ellenõrzése
    If Mid$(fkrec$, 73, 1) = "K" And Trim$(khely$) = "" Then
      Call mess(fksz$ + " " + langmodul(53), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(53), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    If Mid$(fkrec$, 73, 1) = "T" And Trim$(khely$) <> "" Then
      Call mess(fksz$ + " " + langmodul(54), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(54), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    If Trim$(khely$) <> "" Then
      kk$ = Left$(khely$ + Space$(8), 8)
      fhrec$ = dbxkey("FKSZ", kk$)
      If fhrec$ = "" Then
        Call mess(fksz$ + " " + langmodul(55) + " " + kk$ + " " + langmodul(56), 2, 0, langmodul(43), valasz%)
        'MsgBox fksz$ + " " + langmodul(55) + " " + kk$ + " " + langmodul(56), 48, langmodul(43)
        fellen = 1: Exit Function
      Else
        If Mid$(fhrec$, 69, 1) <> "K" Then
          Call mess(fksz$ + " " + langmodul(55) + " " + kk$ + " " + langmodul(57), 2, 0, langmodul(43), valasz%)
          'MsgBox fksz$ + " " + langmodul(55) + " " + kk$ + " " + langmodul(57), 48, langmodul(43)
          fellen = 1: Exit Function
        End If
        If Mid$(fhrec$, 76, 1) = "*" Then
          Call mess(fksz$ + " " + langmodul(55) + " " + kk$ + " " + langmodul(58), 2, 0, langmodul(43), valasz%)
          'MsgBox fksz$ + " " + langmodul(55) + " " + kk$ + " " + langmodul(58), 48, langmodul(43)
          fellen = 1: Exit Function
        End If
      End If
    End If
    '--- szervezeti egység ellenõrzése
    If Mid$(fkrec$, 74, 1) = "K" And Trim$(szerv$) = "" Then
      Call mess(fksz$ + " " + langmodul(59), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(59), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    If Mid$(fkrec$, 74, 1) = "T" And Trim$(szerv$) <> "" Then
      Call mess(fksz$ + " " + langmodul(60), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(60), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    '--- munkaszám ellenõrzése
    If Mid$(fkrec$, 75, 1) = "K" And Trim$(munkaszam$) = "" Then
      Call mess(fksz$ + " " + langmodul(61), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(61), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
    If Mid$(fkrec$, 75, 1) = "T" And Trim$(munkaszam$) <> "" Then
      Call mess(fksz$ + " " + langmodul(62), 2, 0, langmodul(43), valasz%)
      'MsgBox fksz$ + " " + langmodul(62), 48, langmodul(43)
      fellen = 1: Exit Function
    End If
  End If
  fellen = 0
End Function

Public Function konell%(tsy$, tky$, svy$, msy$, dev$, vegyes%, sor%, szoveg$)
  '--- tesztelt!!!
  If Trim$(tsy$) = "" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + szoveg$ + " " + langmodul(66), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + szoveg$ + " " + langmodul(66), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  fkrec$ = dbxkey("FKSZ", tsy$)
  If Mid$(fkrec$, 71, 1) = "L" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(64), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(64), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If vegyes% = 1 And Mid$(fkrec$, 71, 1) = "N" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(65), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(65), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If Mid$(fkrec$, 72, 1) = "K" And Trim$(dev$) = "" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(51), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(51), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If Mid$(fkrec$, 72, 1) = "T" And Trim$(dev$) <> "" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(52), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(52), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If Mid$(fkrec$, 73, 1) = "K" And Trim$(tky$) = "" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(53), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(53), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If Mid$(fkrec$, 73, 1) = "T" And Trim$(tky$) <> "" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(54), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(54), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If Mid$(fkrec$, 74, 1) = "K" And Trim$(svy$) = "" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(59), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(59), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If Mid$(fkrec$, 74, 1) = "T" And Trim$(svy$) <> "" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(60), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(60), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If Mid$(fkrec$, 75, 1) = "K" And Trim$(msy$) = "" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(61), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(61), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If Mid$(fkrec$, 75, 1) = "T" And Trim$(msy$) <> "" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(62), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(62), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If Mid$(fkrec$, 76, 1) = "*" Then
    Call mess(Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(67), 2, 0, langmodul(68), valasz%)
    'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tsy$ + " " + langmodul(67), 48, langmodul(68)
    konell% = 1: Exit Function
  End If
  If Trim$(tky$) <> "" Then
    fkrec$ = dbxkey("FKSZ", tky$)
    If Mid$(fkrec$, 76, 1) = "*" Then
      Call mess(Str$(sor%) + "." + langmodul(63) + " " + tky$ + " " + langmodul(67), 2, 0, langmodul(68), valasz%)
      'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tky$ + " " + langmodul(67), 48, langmodul(68)
      konell% = 1: Exit Function
    End If
    If Mid$(fkrec$, 69, 1) <> "K" Then
      Call mess(Str$(sor%) + "." + langmodul(63) + " " + tky$ + " " + langmodul(57), 2, 0, langmodul(68), valasz%)
      'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tky$ + " " + langmodul(57), 48, langmodul(68)
      konell% = 1: Exit Function
    End If
    If Mid$(fkrec$, 71, 1) = "L" Then
      Call mess(Str$(sor%) + "." + langmodul(63) + " " + tky$ + " " + langmodul(48), 2, 0, langmodul(68), valasz%)
      'MsgBox Str$(sor%) + "." + langmodul(63) + " " + tky$ + " " + langmodul(48), 48, langmodul(68)
      konell% = 1: Exit Function
    End If
  End If
  konell% = 0
End Function

Public Sub turaadat(turaik$, nsuly@, brsuly@, okdb&, orekdb&, tbizt$(), tbizdb&, tszladb&, tszlevdb&, tkp&, nert@, bert@, umod$)
  '--- túra adatainak begyûjtése
  Dim fbizt$(), fbizdb&
  turec$ = dbxkey("KOMS", turaik$)
  nsuly@ = 0: brsuly@ = 0: okdb& = 0: orekdb& = 0: nert@ = 0: bert@ = 0
  tbizdb& = 0: tszladb& = 0: tszlevdb& = 0: tkpdb& = 0
  If turec$ <> "" Then
    For i96% = 1 To 25
      fuvikt$ = Mid$(turec$, (i96% - 1) * 7 + 200, 7)
      If Trim(fuvikt$) <> "" Then
        Call fuvaradat(fuvikt$, nettosuly@, bruttosuly@, okartondb&, orekeszdb&, fbizt$(), fbizdb&, oszladb&, oszlevdb&, okpdb&, onettoertek@, obruttoertek@, umod$)
        nsuly@ = nsuly@ + nettosuly@
        brsuly@ = brsuly@ + bruttosuly@
        okdb& = okdb& + okartondb&
        orekdb& = orekdb& + orekeszdb&
        nert@ = nert@ + onettoertek@
        bert@ = bert@ + obruttoertek@
        tszladb = tszladb + oszladb
        tszlevdb = tszlevdb + oszlevdb
        tkp = tkp + okpdb
        If fbizdb& > 0 Then
          ubizdb& = tbizdb& + fbizdb&
          ReDim Preserve tbizt(1 To ubizdb&)
          For i95% = 1 To fbizdb&
            tbizt(tbizdb& + i95%) = fbizt(i95%)
          Next
          tbizdb& = ubizdb&
        End If
      End If
    Next
  End If
End Sub

Public Sub fuvaradat(fuvarikt$, nettosuly@, bruttosuly@, okartondb&, orekeszdb&, fbizt$(), fbizdb&, oszladb&, oszlevdb&, okpdb&, onettoertek@, obruttoertek@, umod$)
  '--- fuvar adatainak a begyûjtése
  fuvrec$ = dbxkey("FUVA", fuvarikt)
  nettosuly@ = 0: bruttosuly@ = 0: okartondb& = 0: orekeszdb& = 0
  onettoertek@ = 0: obruttoertek@ = 0
  fbizdb& = 0: oszladb& = 0: oszlevdb& = 0: okpdb& = 0
  If fuvrec$ <> "" Then
    For i99% = 1 To 40
      megrik$ = Mid$(fuvrec$, (i99% - 1) * 50 + 300, 7)
      If Trim(megrik$) <> "" Then
        Call megadat(megrik$, mnsuly@, mbsuly@, mkdb&, mrekdb&, statusz$, fizmod$, mnert@, mbert@, umod$)
        If statusz$ <> "*" Then
          nettosuly@ = nettosuly@ + mnsuly@
          bruttosuly@ = bruttosuly@ + mbsuly@
          okartondb& = okartondb& + mkdb&
          orekeszdb& = orekeszdb& + mrekdb&
          onettoertek@ = onettoertek@ + mnert@
          obruttoertek@ = obruttoertek@ + mbert@
          fbizdb = fbizdb + 1
          ReDim Preserve fbizt(1 To fbizdb)
          fbizt(fbizdb) = megrik$
          Select Case statusz$
            Case "S": oszladb& = oszladb& + 1
            Case "R", "L": oszlevdb& = oszlevdb& + 1
            Case Else
          End Select
          If fizmod$ = "01" Or fizmod$ = "02" Or fizmod$ = "05" Then okpdb& = okpdb& + 1
        End If
      End If
    Next
  End If
End Sub

Public Sub megadat(megrik$, nsuly@, brsuly@, okdb&, orekdb&, statusz$, fizmod$, nert@, bert@, umod$)
  '--- megrendelés adatainak begyûjtése
  '--- umod=0 minden 1-súly nélkül 2-súly és rekeszek nélkül
  nsuly@ = 0: brsuly@ = 0: okdb& = 0: orekdb& = 0: nert@ = 0: bert@ = 0
  megrec$ = dbxkey("ERTB", megrik$)
  If megrec$ <> "" Then
    statusz$ = Mid$(megrec$, 26, 1): If Mid$(megrec$, 201, 1) = "S" Then statusz$ = "*"
    nert@ = xval(Mid$(megrec$, 145, 10))
    bert@ = xval(Mid$(megrec$, 165, 10))
    fizmod$ = Mid$(megrec$, 242, 2)
    '--- kartonok, rekeszek és sûlyok
    If umod$ = "0" Or umod$ = "1" Then
      xdb& = xval(Mid$(megrec$, 189, 4))
      For i92% = 1 To xdb&
        DoEvents
        azo$ = megrik$ + Right("0000" + Trim(Str(i92%)), 4)
        sr$ = dbxkey("ERTT", azo$)
        If sr$ <> "" Then
          menny@ = xval(Mid$(sr$, 67, 12))
          rek& = xval(Mid$(sr$, 91, 4))
          If Mid$(sr$, 23, 1) <> "S" Then
            tkod$ = Mid$(sr$, 28, 15)
            ktrmrec$ = dbxkey("KTRM", tkod$)
            kartondb& = xval(Mid$(ktrmrec$, 1226, 7))
            If kartondb& > 0 Then
              If rek& > 0 Then
                orekdb& = orekdb& + rek&
                okdb& = okdb& + rek&
              Else
                If kartondb > 1 Then
                  okdb& = okdb& + menny@ / kartondb&
                End If
              End If
            End If
            If umod$ = "0" Then
              '--- súlyok
              esuly@ = xval(Mid$(ktrmrec$, 834, 12))
              nsuly@ = nsuly@ + menny@ * esuly@
              brsuly@ = brsuly@ + menny@ * esuly@
              tg$ = Mid$(ktrmrec$, 1067, 15)
              gg$ = Mid$(ktrmrec$, 908, 15)
              If Trim(tg$) <> "" Then
                gtrmrec$ = dbxkey("KTRM", tg$)
                If gtrmrec$ <> "" Then
                  tsuly@ = xval(Mid$(gtrmrec$, 834, 12))
                  brsuly@ = brsuly@ + menny@ * tsuly@
                End If
              End If
              If Trim(gg$) <> "" And rek& <> 0 Then
                rtrmrec$ = dbxkey("KTRM", gg$)
                If rtrmrec$ <> "" Then
                  rsuly@ = xval(Mid$(rtrmrec$, 834, 12))
                  brsuly@ = brsuly@ + rek& * rek&
                End If
              End If
            End If
          End If
        End If
      Next
    End If
  End If
End Sub

Public Function szamlaiktato$(qszlasz$)
  iaqwfil = FreeFile
  Open auditorutvonal$ + "auw-pvsz.ndx" For Binary Shared As #iaqwfil
  iaqwdb& = Int(LOF(iaqwfil) / 12)
  If iaqwdb& > hsxrecdb& Then
    iaqwfil1 = FreeFile
    Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #iaqwfil1
  End If
  If hsxrecdb& > 0 Then
    For iaqw& = 1 To hsxrecdb&
      iasz$ = Mid$(hsxrec$, (iaqw& - 1) * 15 + 1, 15)
      If iasz$ = qszlasz$ Then
        iaqwfil = FreeFile
        Open auditorutvonal$ + "auw-pvsz.ndx" For Binary Shared As #iaqwfil
        Ikt$ = Space(7)
        Get #iaqwfil, (iaqw& - 1) * 12 + 5, Ikt$
        szamlaiktato = Ikt$
        Close iaqwfil: Close iaqwfil1
        Exit Function
      End If
    Next
  End If
  If iaqwdb& > hsxrecdb& Then
    For iaqw& = hsxrecdb& + 1 To iaqwdb&
      Ikt$ = Space(7)
      Get #iaqwfil, (iaqw& - 1) * 12 + 1, rcim&
      iar$ = Space(50)
      Get #iaqwfil1, rcim& + 9, iar$
      If Mid$(iar$, 8, 15) = qszlasz$ Then
        szamlaiktato = Mid$(iar$, 1, 7)
        Close iaqwfil: Close iaqwfil1
        Exit Function
      End If
    Next
  End If
  Close iaqwfil: Close iaqwfil1
  szamlaiktato$ = ""
End Function

Public Sub compenz(a1ikt$, a2ikt$, oszeg@, elszla$, koriktat$)
  '--- vevõ számla - jóváíró számla összevezetés
  '--- a1ikt$=alapszámla, a2ikt$=módosító számla
  If oszeg@ <> 0 Then
    ssrec$ = dbxkey("PVSZ", a2ikt$)
    konyveldat$ = Mid$(ssrec$, 58, 6)
    korpkod$ = Mid$(ssrec$, 38, 15)
    '--- módosító számla
    korec$ = Space$(600)
    Mid$(korec$, 8, 6) = konyveldat$
    Mid$(korec$, 14, 1) = "V"
    Mid$(korec$, 15, 1) = "N"
    Mid$(korec$, 16, 15) = korpkod$
    Mid$(korec$, 31, 7) = a2ikt$
    Mid$(korec$, 45, 14) = ertszam(Str(oszeg@), 14, 2)
    Mid$(korec$, 76, 30) = "Automatikus korrekció"
    Mid$(korec$, 106, 6) = maidatum$
    Mid$(korec$, 112, 8) = ugyintezo$
    elem$ = Space$(53)
    Mid$(elem$, 1, 14) = ertszam(Str(oszeg@), 14, 2)
    Mid$(elem$, 15, 8) = elszla
    Mid$(korec$, 310, 53) = elem$
    Call korrkonyvel(korec$, "V")
    koriktat$ = Mid$(korec$, 1, 7)
    '--- alapszámla
    korec$ = Space$(600)
    Mid$(korec$, 8, 6) = konyveldat$
    Mid$(korec$, 14, 1) = "V"
    Mid$(korec$, 15, 1) = "N"
    Mid$(korec$, 16, 15) = korpkod$
    Mid$(korec$, 31, 7) = a1ikt$
    Mid$(korec$, 45, 14) = ertszam(Str(-oszeg@), 14, 2)
    Mid$(korec$, 76, 30) = "Automatikus korrekció"
    Mid$(korec$, 106, 6) = maidatum$
    Mid$(korec$, 112, 8) = ugyintezo$
    elem$ = Space$(53)
    Mid$(elem$, 1, 14) = ertszam(Str(-oszeg@), 14, 2)
    Mid$(elem$, 15, 8) = elszla
    Mid$(korec$, 310, 53) = elem$
    Call korrkonyvel(korec$, "V")
  End If
End Sub
Public Sub compsztorno(krika$)
  '--- egymást követõ kompenzációs tételek sztornója
  If Mid$(krika$, 1, 3) = "Ko:" Then
    krik$ = Mid$(krika$, 4, 7)
    Call korrsztorno(krik$, "V", maidatum$)
    kri1$ = xval(krik$) + 1
    krik$ = Right("0000000" + Trim(Str(kri1$)), 7)
    Call korrsztorno(krik$, "V", maidatum$)
  End If
End Sub

Public Sub szamlasorertek(termekkod$, mennyiseg@, bruttoegysegar@, cnettoertek@, cafamentes@, cafa05@, cafa20@, cafaadojegyes@)
  ertker% = xval(Mid$(irec$, 344, 1))
  afaker% = xval(Mid$(irec$, 345, 1))
  If ertker% = 0 Then fste$ = "############0" Else fste$ = "#############0." + String(ertker%, "0")
  If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
  If Len(termekkod$) = 2 Then
    afakod$ = termekkod$: If afkulcs@ = 99 Then afkulcs@ = 0
    afkulcs@ = xval(afakod$)
  Else
    xtermrec$ = dbxkey("KTRM", termekkod$)
    afakod$ = Mid$(xtermrec$, 706, 2)
    afkulcs@ = xval(afakod$): If afkulcs@ = 99 Then afkulcs@ = 0
  End If
  netar@ = xval(ertszam(Str(bruttoegysegar@ / (1 + afkulcs@ / 100)), 10, 2))
  nert@ = xval(Format(netar@ * mennyiseg@, fste$))
  cnettoertek@ = cnettoertek@ + nert@
  '--- termék áfája
  If afakod$ = "05" Then
    cafa05@ = cafa05@ + nert@ * 0.05
  Else
    If afakod$ = "20" Then
      cafa20@ = cafa20@ + nert@ * 0.2
    End If
  End If
End Sub

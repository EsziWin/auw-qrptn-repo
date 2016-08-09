Attribute VB_Name = "ListMod"

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
Dim ux%(1001), param$(20)
Dim mp(1001) As MTYP
Dim parr$(1001)
Dim narr$(1001)
Dim sor(30) As sorleir     '-sor mezoi tomb pointer=spp%
Dim fejt$(20)              '-fejlecek pointer=fejp%
Dim kif$(10)               '-kifejezesek pointer=kifp%
                           '-srr$ a listasor
Dim elo$(30), abert$(30)
Dim eljp%(1001), elje$(1001)



Public Sub Nylistazo(tdef$, listanev$, komment$, listhiba%)
  '--- TBT kiterjesztésû fájl interpretálása
  On Error GoTo hibakez
  listhiba% = 0
  soso% = 100
  szelesseg% = 0
  sortores% = 0
  ' Eszi
  For i% = 1 To 10
   For j% = 1 To 20
     gyujto@(i%, j%) = 0
   Next
  Next
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
        If hal% = 1 Then
          Mid$(maszk$, 1, 1) = "#": sor(spp%).halm = 1
        Else
          sor(spp%).halm = 0
        End If
        If fle% = 1 Then
          Mid$(maszk$, 1, 1) = "X": sor(spp%).flem = 1
        Else
          sor(spp%).flem = 0
        End If
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
                If j1% = 7 Then
                  wforma$ = forma$
                  wpozic% = pozic%
                  wmhx% = mhx%
                End If
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
             
            If j% = 1 Then
              tip$ = Mid(b$, 15, 1)
              If tip$ = "N" Or tip$ = "K" Then
            
                Call kerekit510(gyujto@(j%, 8), kerekossz@, kerek@, "K")
                If Not gyujto@(j%, 8) = kerekossz@ Then
                  ss$ = Space$(Len(srr$))
                  Mid$(ss$, 1, 20) = "Kerekítés:"
                  'mezir$ = Right$(Space$(20) + Format$(kerek@, wforma$), wmhx%)
                  mezir$ = Right$(Space$(20) + Format$(kerek@, "############0.00"), 16)
                  'If kerek@ < 0 Then
                  '  mezir$ = mezir$ + "-":
                  'End If
  '                Mid$(ss$, wpozic%, wmhx%) = mezir$
                  Mid$(ss$, 81, 16) = mezir$

                  Print #fi20, "TS" + ss$
                  gyujto@(j%, 8) = gyujto@(j%, 8) + kerek@
                End If
              End If
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
      If j% = 1 Then
              tip$ = Mid(b$, 15, 1)
              If tip$ = "N" Or tip$ = "K" Then
            
                Call kerekit510(gyujto@(j%, 8), kerekossz@, kerek@, "K")
                If Not gyujto@(j%, 7) = kerekossz@ Then
                  ss$ = Space$(Len(srr$))
                  Mid$(ss$, 1, 20) = "Kerekítés:"
                  mezir$ = Right$(Space$(20) + Format$(kerek@, "############0.00"), 16)
                  'If kerek@ < 0 Then
                  '  mezir$ = mezir$ + "-":
                  'End If
                  Mid$(ss$, 81, 16) = mezir$

                  Print #fi20, "TS" + ss$
                  gyujto@(j%, 8) = gyujto@(j%, 8) + kerek@
                End If
              End If
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
  Call mess(langmodul(166) + Str(Err()), 2, 0, langmodul(165), valasz%)
 ' Resume Next
  Close fi1: Close fi20
  Exit Sub
End Sub


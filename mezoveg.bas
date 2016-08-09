Attribute VB_Name = "mezoveg"
Public meszovtomb$(1001, 5), vtsztomb$(1001), mozrec$, krakrec$, akcirec$(5000), akcidb&
Public forintja@(1001), zarfo#, nemarazni%, xvskod$, stax%, tablamennyiseg@, jovadatum$
Public rtomb$
Public Sub mezoeleje(vsor%, voszl%)
  Select Case programnev$
    Case "AUW-RTRM"
      If objektum$ = "ARE1" Then
        vs& = vsor%: vo& = voszl%
        form1.List3.Clear
        arakt$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(4), 4)
        rarec$ = dbxkey("KRAK", arakt$)
        form1.List3.AddItem Mid$(rarec$, 5, 60)
      End If
      If objektum$ = "AKCX" Then
        vs& = vsor%: vo& = voszl%
        form1.List3.Clear
        aikt$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(7), 7)
        afrec$ = dbxkey("AKCF", aikt$)
        form1.List3.AddItem Mid$(afrec$, 8, 30)
        form1.List3.AddItem " "
        If Mid$(afrec$, 38, 1) = "I" Then
          form1.List3.AddItem "Szállítói: " + datki(Mid$(afrec$, 41, 6)) + "-" + datki(Mid$(afrec$, 47, 6))
        End If
        If Mid$(afrec$, 39, 1) = "I" Then
          form1.List3.AddItem "Nagyker: " + datki(Mid$(afrec$, 53, 6)) + "-" + datki(Mid$(afrec$, 59, 6))
        End If
        If Mid$(afrec$, 40, 1) = "I" Then
          form1.List3.AddItem "Fogyasztói:" + datki(Mid$(afrec$, 65, 6)) + "-" + datki(Mid$(afrec$, 71, 6))
        End If
        form1.List3.AddItem " "
        If Trim(Mid$(afrec$, 77, 15)) <> "" Then
          pkod$ = Mid$(afrec$, 77, 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            form1.List3.AddItem "Gyártó:"
            form1.List3.AddItem "  " + Trim(Mid$(partrec$, 16, 50))
          End If
        End If
        If Trim(Mid$(afrec$, 92, 15)) <> "" Then
          pkod$ = Mid$(afrec$, 92, 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            form1.List3.AddItem "Szállító:"
            form1.List3.AddItem "  " + Trim(Mid$(partrec$, 16, 50))
          End If
        End If
      End If
    Case "AUW-ARMG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "CRM1" Then
        form1.List2.Clear
        tkkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
        If tkkod$ <> Space(15) Then
          termrec$ = dbxkey("KTRM", tkkod$)
          If termrec$ <> "" Then
            form1.List2.AddItem "Termék: " + Mid$(termrec$, 16, 60)
            xx$ = "": zz$ = ""
            For i81% = 1 To 7
              nv$ = Trim(Mid$(ikonfrec$, (i81% - 1) * 27 + 5, 20))
              ar$ = Trim(Mid$(termrec$, (i81% - 1) * 14 + 580, 14))
              If i81% < 5 Then
                xx$ = xx$ + nv$ + ":" + ar$ + "     "
              Else
                yy$ = yy$ + nv$ + ":" + ar$ + "     "
              End If
            Next
            nv$ = "Disztr.ár:"
            ar$ = Trim(Mid$(termrec$, 893, 14))
            yy$ = yy$ + nv$ + ":" + ar$ + "     "
            nv$ = "Fogy.ár:"
            ar$ = Trim(Mid$(termrec$, 678, 14))
            yy$ = yy$ + nv$ + ":" + ar$
            form1.List2.AddItem xx$
            form1.List2.AddItem yy$
          End If
        End If
      End If
      If objektum$ = "AKCS" Then
        form1.List1.Clear
        form1.List2.Clear
        tkkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
        If tkkod$ <> Space(15) Then
          termrec$ = dbxkey("KTRM", tkkod$)
          If termrec$ <> "" Then
            form1.List1.AddItem "Termék: " + Mid$(termrec$, 16, 60)
            xx$ = "": zz$ = ""
            nv$ = "Besz.ár:"
            ar$ = Trim(Mid$(termrec$, 1025, 12))
            xx$ = xx$ + nv$ + ":" + ar$ + "     "
            For i81% = 1 To 6
              nv$ = Trim(Mid$(ikonfrec$, (i81% - 1) * 27 + 5, 20))
              ar$ = Trim(Mid$(termrec$, (i81% - 1) * 14 + 580, 14))
              If i81% < 5 Then
                xx$ = xx$ + nv$ + ":" + ar$ + "     "
              Else
                yy$ = yy$ + nv$ + ":" + ar$ + "     "
              End If
            Next
            nv$ = "Disztr.ár:"
            ar$ = Trim(Mid$(termrec$, 893, 14))
            yy$ = yy$ + nv$ + ":" + ar$ + "     "
            nv$ = "Fogy.ár:"
            ar$ = Trim(Mid$(termrec$, 678, 14))
            yy$ = yy$ + nv$ + ":" + ar$
            form1.List1.AddItem xx$
            form1.List1.AddItem yy$
          End If
        End If
      End If
    Case "AUW-JOV"
      Call jmezoeleje(vsor%, voszl%)
    Case "AUW-DOK"
    Case "AUW-RSMG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KSM1" Then
        If vsorszam& = 0 Then
          mezohiba% = -1
          Call mezovege(vsor%, voszl%, mezohiba%, 0)
        End If
        'If vs& <> vsorszam& Then
          form1.MSFlexGrid4.Visible = False
          vsorszam& = vs&
          termkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          minta$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
          If Trim(termkod$) <> "" Then
            termrec$ = dbxkey("KTRM", termkod$)
            If minta$ = Space$(8) Then minta$ = ""
            form1.List1.Clear
            'form1.List1.AddItem langmodul(85) + ":" + Trim(termkod$) + "  " + langmodul(86) + ":" + vevocikk$
            'form1.List1.AddItem "Termék kód:" + Trim(termkod$) + "  Vevõi cikkszám:" + vevocikk$
            form1.List1.AddItem Mid$(termrec$, 16, 60)
            form1.List1.AddItem Mid$(termrec$, 196, 60)
            'form1.List1.AddItem langmodul(87) + " :" + Mid$(termrec$, 552, 14)
            form1.List1.AddItem "Utolsó beszerzési ár :" + Mid$(termrec$, 566, 14)
            rdb% = Len(rtomb$) / 5
            For i% = 1 To rdb%
              raktkod$ = Mid$(rtomb$, (i% - 1) * 5 + 1, 4)
              rkszind$ = raktkod$ + Mid$(termrec$, 1, 15)
              rkszrec$ = dbxkey("RKSZ", rkszind$)
              If Val(Mid$(rkszrec$, 20, 14)) > 0 Then
                form1.List1.AddItem raktkod$ + " raktár készlet: " + Mid$(rkszrec$, 20, 14)
              End If
            Next

          Else
            form1.List1.Clear
          End If
        'End If
      End If
    Case "AUW-RMEG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KME1" Then
        If vsorszam& = 0 Then
          mezohiba% = -1
          Call mezovege(vsor%, voszl%, mezohiba%, 0)
        End If
        If vs& <> vsorszam& Then
          form1.MSFlexGrid4.Visible = False
          vsorszam& = vs&
          termkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          minta$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
          If Trim(termkod$) <> "" Then
            termrec$ = dbxkey("KTRM", termkod$)
            armind$ = Mid$(termrec$, 1, 15) + Mid$(partrec$, 1, 15)
            armgrec$ = dbxkey("ARMG", armind$)
            vevocikk$ = ""
            If armgrec$ <> "" Then vevocikk$ = Trim(Mid(armgrec$, 31, 15))
            If minta$ = Space$(8) Then minta$ = ""
            form1.List1.Clear
            form1.List1.AddItem langmodul(85) + ":" + Trim(termkod$) + "  " + langmodul(86) + ":" + vevocikk$
            'form1.List1.AddItem "Termék kód:" + Trim(termkod$) + "  Vevõi cikkszám:" + vevocikk$
            form1.List1.AddItem Mid$(termrec$, 16, 60)
            form1.List1.AddItem Mid$(termrec$, 196, 60)
            form1.List1.AddItem langmodul(88) + " 1:" + Mid$(termrec$, 580, 14)
            'form1.List1.AddItem "Listaár 1:" + Mid$(termrec$, 580, 14)
            If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
              Call rkeszlet(termkod$, minta$, Mid$(rakrec$, 1, 4), keszme@, foglme@)
              form1.List1.AddItem langmodul(89) + ":" + ertszam(Str(keszme@), 14, 2) + " " + langmodul(90) + ":" + ertszam(Str(foglme@), 14, 2) + " " + langmodul(91) + ":" + ertszam(Str(keszme@ - foglme@), 14, 2)
              'form1.List1.AddItem "Raktárkészlet:" + ertszam(Str(keszme@), 14, 2) + " Foglalt készlet:" + ertszam(Str(foglme@), 14, 2) + " Szabad készlet:" + ertszam(Str(keszme@ - foglme@), 14, 2)
              form1.List1.AddItem langmodul(92) + ":" + Mid$(termrec$, 542, 10)
              'form1.List1.AddItem "Rendelhetõ mennyiség:" + Mid$(termrec$, 542, 10)
              If Mid$(termrec$, 443, 1) = "M" And Mid$(rakrec$, 67, 1) = "I" Then
                form1.List1.AddItem langmodul(93)
                'form1.List1.AddItem "Mintaszám kötelezõ"
                '--- mintaszám táblázat mgejlenítése
                rkszind$ = Mid$(rakrec$, 1, 4) + Mid$(termrec$, 1, 15)
                rkszrec$ = dbxkey("RKSZ", rkszind$)
                form1.MSFlexGrid4.Visible = False
              Else
                form1.MSFlexGrid4.Visible = False
                form1.List1.AddItem langmodul(94)
                'form1.List1.AddItem "Mintaszám tilos"
              End If
            Else
              form1.List1.AddItem langmodul(95)
              'form1.List1.AddItem "Nincs készletnyilvántartás"
            End If
          Else
            form1.List1.Clear
          End If
        End If
      End If
    Case "AUW-REGY", "AUW-CEGY", "AUW-CVIS"
      vs& = vsor%: vo& = voszl%
      If (objektum$ = "KKF4" Or objektum$ = "KKF3" Or objektum$ = "KKF1") And programnev$ = "AUW-CEGY" Then
        form1.Label16.Visible = True
        form1.Label17.Visible = True
        form1.Label18.Visible = True
        If objektum$ = "KKF3" Then
          Call kiszamolja(biznetto@, bizbrutto@, fogyert@, 1)
        Else
          If objektum$ = "KKF1" Then
            Call kiszamolja(biznetto@, bizbrutto@, fogyert@, 2)
          Else
            Call kiszamolja(biznetto@, bizbrutto@, fogyert@, 3)
          End If
        End If
        form1.Label16.Caption = " Nettó érték:" + Chr(13) + ertszamx(Str(biznetto@), 17, 2)
        form1.Label17.Caption = " Bruttó érték:" + Chr(13) + ertszamx(Str(bizbrutto@), 17, 2)
        form1.Label18.Caption = " Fogy. érték:" + Chr(13) + ertszamx(Str(fogyert@), 17, 2)
      End If
      If objektum$ = "KKF4" Then
        '--- tapadó göngyöleges
        'If vo& > 2 Then vo& = 2
        If vsorszam& = 0 Then
          mezohiba% = -1
          Call mezovege(vsor%, voszl%, mezohiba%, 0)
        End If
        If vs& <> vsorszam& Then
          vsorszam& = vs&
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
              form1.List2.Clear
              form1.List2.AddItem Mid$(termrec$, 16, 60)
              kartondb% = xval(Mid$(termrec$, 1226, 7))
              If kartondb% > 1 Then
                xyx$ = "K" + Trim(Str(kartondb%))
              Else
                xyx$ = " "
              End If
              form1.List2.AddItem xyx$
              form1.List2.AddItem "Nyilv.ár: " + Trim(Mid$(termrec$, 554, 12)) + " Ref.ár: " + Trim(Mid$(termrec$, 1276, 12)) + " Nagyk.ár: " + Trim(Mid$(termrec$, 582, 12)) + " Fogy.ár: " + Trim(Mid$(termrec$, 680, 12))
              form1.List2.AddItem "Akt.készl: " + Trim(Mid$(termrec$, 750, 12)) + " Közp.rak: " + Trim(Mid$(termrec$, 929, 12)) + " Vegyi: " + Trim(Mid$(termrec$, 943, 12)) + " C+C: " + Trim(Mid$(termrec$, 957, 12))
            Else
              form1.List2.Clear
            End If
          Else
            form1.List2.Clear
          End If
        End If
      End If
      If objektum$ = "KKF2" Then
        If vsorszam& = 0 Then
          mezohiba% = -1
          Call mezovege(vsor%, voszl%, mezohiba%, 0)
        End If
        If vs& <> vsorszam& Then
          vsorszam& = vs&
          If programnev$ = "AUW-CEGY" Then
          End If
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(15), 15)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
              If programnev$ = "AUW-CEGY" Then
                form1.List2.Clear
                form1.List2.AddItem Mid$(termrec$, 16, 60)
                kartondb% = xval(Mid$(termrec$, 1226, 7))
                If kartondb% > 1 Then
                  xyx$ = "K" + Trim(Str(kartondb%))
                Else
                  xyx$ = " "
                End If
                form1.List2.AddItem xyx$
                form1.List2.AddItem "Nyilv.ár: " + Trim(Mid$(termrec$, 554, 12)) + " Ref.ár: " + Trim(Mid$(termrec$, 1276, 12)) + " Nagyk.ár: " + Trim(Mid$(termrec$, 582, 12)) + " Fogy.ár: " + Trim(Mid$(termrec$, 680, 12))
                form1.List2.AddItem "Akt.készl: " + Trim(Mid$(termrec$, 750, 12)) + " Közp.rak: " + Trim(Mid$(termrec$, 929, 12)) + " Vegyi: " + Trim(Mid$(termrec$, 943, 12)) + " C+C: " + Trim(Mid$(termrec$, 957, 12))
              Else
                form1.List2.Clear
                form1.List2.AddItem Mid$(termrec$, 16, 60)
                form1.List2.AddItem Mid$(termrec$, 196, 60)
                form1.List2.AddItem langmodul(87) + ":" + Mid$(termrec$, 552, 14)
                'form1.List2.AddItem "Nyilv.ár:" + Mid$(termrec$, 552, 14)
              End If
            Else
              form1.List2.Clear
            End If
          Else
            form1.List2.Clear
          End If
        End If
      End If
      If objektum$ = "KKF1" Or objektum$ = "KKF3" Then
        If vsorszam& = 0 Then
          mezohiba% = -1
          Call mezovege(vsor%, voszl%, mezohiba%, 0)
        End If
        If vs& <> vsorszam& Then
          vsorszam& = vs&
          If objektum$ = "KXF1" Then
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
          Else
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          End If
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
              form1.List2.Clear
              form1.List2.AddItem Mid$(termrec$, 16, 60)
              form1.List2.AddItem Mid$(termrec$, 196, 60)
              Select Case Mid$(termrec$, 442, 1)
                Case "A", "R", "F", "K"
                  If Mid$(termrec$, 443, 1) <> "N" Then
                    form1.List2.AddItem langmodul(96) + ": " + Trim$(ertszam(Mid$(termrec$, 748, 14), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6)) + " " + langmodul(97) + ":" + Trim$(ertszam(Mid$(termrec$, 762, 14), 14, 3)) + " " + langmodul(87) + ":" + Trim(ertszam(Mid$(termrec$, 552, 14), 14, 2))
                    'form1.List2.AddItem "Aktuális készlet: " + Trim$(ertszam(Mid$(termrec$, 748, 14), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6)) + " Foglalt:" + Trim$(ertszam(Mid$(termrec$, 762, 14), 14, 3)) + " Nyilv.ár:" + Trim(ertszam(Mid$(termrec$, 552, 14), 14, 2))
                  Else
                    Call mess(langmodul(98), 2, 0, langmodul(99), valasz%)
                    'MsgBox langmodul(98), 48, langmodul(99)
                    'MsgBox "Készletkezelés letiltva", 48, "Termék hiba!"
                    form1.List2.Clear
                  End If
                Case "S"
                  Call mess(langmodul(100), 2, 0, langmodul(99), valasz%)
                  'MsgBox langmodul(100), 48, langmodul(99)
                  'MsgBox "Szolgáltatás", 48, "Termék hiba!"
                  form1.List2.Clear
                Case Else
              End Select
            Else
              form1.List2.Clear
            End If
          Else
            form1.List2.Clear
          End If
        End If
      End If
    Case "AUW-KFRG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KFTT" Then
        If vsorszam& = 0 Then
          mezohiba% = -1
          Call mezovege(vsor%, voszl%, mezohiba%, 0)
        End If
        If vs& <> vsorszam& Then
          vsorszam& = vs&
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 6)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("PTRM", tkod$)
            If termrec$ <> "" Then
              form1.List2.Visible = True
              form1.List2.Clear
              form1.List2.AddItem Mid$(termrec$, 7, 60)
              form1.List2.AddItem Mid$(termrec$, 67, 60)
              Select Case Mid$(termrec$, 127, 1)
                Case "A"
                  If Mid$(termrec$, 187, 1) = "I" Then
                    form1.List2.AddItem langmodul(96) + ": " + Trim$(ertszam(Mid$(termrec$, 315, 14), 14, 3)) + " " + Trim(Mid$(termrec$, 140, 6)) + " " + langmodul(87) + ":" + Trim(ertszam(Mid$(termrec$, 301, 14), 14, 2))
                    'form1.List2.AddItem "Aktuális készlet: " + Trim$(ertszam(Mid$(termrec$, 315, 14), 14, 3)) + " " + Trim(Mid$(termrec$, 140, 6)) + " Nyilv.ár:" + Trim(ertszam(Mid$(termrec$, 301, 14), 14, 2))
                  Else
                    Call mess(langmodul(98), 2, 0, langmodul(99), valasz%)
                    'MsgBox langmodul(98), 48, langmodul(99)
                    'MsgBox "Készletkezelés letiltva", 48, "Termék hiba!"
                    form1.List2.Clear
                  End If
                Case "F", "S"
                  Call mess(langmodul(100), 2, 0, langmodul(99), valasz%)
                  'MsgBox langmodul(100), 48, langmodul(99)
                  'MsgBox "Szolgáltatás", 48, "Termék hiba!"
                  form1.List2.Clear
                Case Else
              End Select
            Else
              form1.List2.Clear
            End If
          Else
            form1.List2.Clear
          End If
        End If
      End If
    Case "AUW-PSZL", "AUW-SPDC"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "PSZL" Or objektum$ = "PSZ2" Then
        If vsorszam& = 0 Then
          mezohiba% = -1
          Call mezovege(vsor%, voszl%, mezohiba%, 0)
        End If
        If vs& <> vsorszam& Then
          vsorszam& = vs&
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 6)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("PTRM", tkod$)
            If termrec$ <> "" Then
              form1.List2.Clear
              If Trim(Mid$(termrec$, 7, 60)) = "" Then
                form1.List2.AddItem meszovtomb$(vs&, 1)
                form1.List2.AddItem "   "
              Else
                form1.List2.AddItem Mid$(termrec$, 7, 60)
                form1.List2.AddItem Mid$(termrec$, 67, 60)
              End If
              Select Case Mid$(termrec$, 127, 1)
                Case "A"
                  If Mid$(termrec$, 187, 1) = "I" Then
                    form1.List2.AddItem langmodul(101) + " " + langmodul(102) + ". " + langmodul(96) + ": " + Trim$(ertszam(Mid$(termrec$, 315, 14), 14, 3)) + " " + Trim(Mid$(termrec$, 140, 6))
                    'form1.List2.AddItem "Áru, anyag készletkezeléssel. Aktuális készlet: " + Trim$(ertszam(Mid$(termrec$, 315, 14), 14, 3)) + " " + Trim(Mid$(termrec$, 140, 6))
                  Else
                    form1.List2.AddItem langmodul(101) + " " + langmodul(103) + "."
                    'form1.List2.AddItem "Áru, anyag készletkezelés nélkül."
                  End If
                Case "F", "S"
                  form1.List2.AddItem langmodul(100) + "."
                  'form1.List2.AddItem "Szolgáltatás."
                Case Else
              End Select
            Else
              form1.List2.Clear
            End If
          Else
            form1.List2.Clear
          End If
        End If
      End If
    Case "AUW-CJOV"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KSZL" Or objektum$ = "KSVL" Then
        If vsorszam& = 0 Then
          mezohiba% = -1
          Call mezovege(vsor%, voszl%, mezohiba%, 0)
        End If
        If vs& <> vsorszam& Then
          vsorszam& = vs&
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
              form1.List2.Clear
              form1.List2.AddItem Mid$(termrec$, 16, 60)
              form1.List2.AddItem Mid$(termrec$, 196, 60)
              tagon$ = Mid$(termrec$, 1067, 15)
              If Trim(tagon$) <> "" Then
                gonrec$ = dbxkey("KTRM", tagon$)
                form1.List2.AddItem Mid$(gonrec$, 16, 60)
              End If
              'minta$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
              Call rkeszlet(tkod$, "", Mid$(krakrec$, 1, 4), keszme@, foglme@)
            Else
              form1.List2.Clear
            End If
          Else
            form1.List2.Clear
          End If
        End If
      End If
    Case "AUW-CSZL", "AUW-RSZL", "AUW-RJOV", "AUW-RSZS"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KSZL" Or objektum$ = "KSVL" Then
        If vsorszam& = 0 Then
          mezohiba% = -1
          Call mezovege(vsor%, voszl%, mezohiba%, 0)
        End If
        If vs& <> vsorszam& Then
          vsorszam& = vs&
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
              form1.List2.Clear
              form1.List2.AddItem Mid$(termrec$, 16, 60)
              form1.List2.AddItem Mid$(termrec$, 196, 60)
              minta$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
              If programnev$ = "AUW-CSZL" Then
                keszme@ = xval(Mid$(termrec$, 748, 14))
              Else
                Call rkeszlet(tkod$, minta$, Mid$(krakrec$, 1, 4), keszme@, foglme@)
              End If
              Select Case Mid$(termrec$, 442, 1)
                Case "R"
                  If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
                    form1.List2.AddItem langmodul(104) + " " + langmodul(102) + ". " + langmodul(96) + ": " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                  Else
                    form1.List2.AddItem langmodul(104) + " " + langmodul(103) + "."
                    'form1.List2.AddItem "Áru készletkezelés nélkül."
                  End If
                Case "K", "F"
                  If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
                    form1.List2.AddItem langmodul(106) + " " + langmodul(102) + ". " + langmodul(96) + ": " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                    'form1.List2.AddItem "Saját termék készletkezeléssel. Aktuális készlet: " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                  Else
                    form1.List2.AddItem langmodul(106) + " " + langmodul(103) + "."
                    'form1.List2.AddItem "Saját termék készletkezelés nélkül."
                  End If
                Case "A"
                  If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
                    form1.List2.AddItem langmodul(105) + " " + langmodul(102) + ". " + langmodul(96) + ": " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                    'form1.List2.AddItem "Anyag készletkezeléssel. Aktuális készlet: " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                  Else
                    form1.List2.AddItem langmodul(105) + " " + langmodul(103) + "."
                    'form1.List2.AddItem "Anyag készletkezelés nélkül."
                  End If
                Case "S"
                  form1.List2.AddItem langmodul(100) + "."
                  'form1.List2.AddItem "Szolgáltatás."
                Case Else
              End Select
            Else
              form1.List2.Clear
            End If
          Else
            form1.List2.Clear
          End If
        End If
      End If
    Case Else
  End Select
End Sub

Public Sub mezovege2(vsor%, voszl%, mezohiba%, abmod%)
  '--- adatvége teendõk beépítése 2.
  '--- objektum$ az obj.azonosító public
  '--- abmod%=1 ablak abmod%=0 táblázat
  '--- ablaksorszam% az ablak sorszáma public
  If mezohiba% <> -1 Then mezohiba% = 0
  Select Case programnev$
    Case "AUW-RFVU"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "FUVA" Then
        If vs& = 4 Then
          gkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(4, 1) + Space$(10), 10)
          grec$ = dbxkey("SPGJ", gkod$)
          If grec$ <> "" Then
            Vektor.MSFlexGrid1.TextMatrix(5, 1) = Mid$(grec$, 300, 40)
          End If
        End If
      End If
      If objektum$ = "FUVB" Then
        If vo& = 1 Then
          megi$ = Trim(Tabla.MSFlexGrid1.TextMatrix(vs&, 1))
          megrec$ = dbxkey("ERTB", megi$)
          kellx% = 1
          If Len(Trim(Mid$(megrec$, 216, 8))) = 7 Then
            Call mess("Ez a megrendelés már fuvaron van. Iktató:" + Mid$(megrec$, 216, 7) + Chr(13) + Chr(13) + "Mégis folytatni akarja?", 5, 3, "Figyelmeztetés", valasz%)
            If valasz% = 0 Then kellx% = 0
          End If
          If kellx% = 1 Then
            If Mid$(megrec$, 26, 1) <> "R" And Mid$(megrec$, 26, 1) <> "S" And Mid$(megrec$, 26, 1) <> "L" Then
              Call mess("Ez a státusz nem rakható fuvarra!", 3, 0, "Hiba", valasz%)
            Else
              If Mid$(megrec$, 26, 1) = "R" Or Mid$(megrec$, 26, 1) = "L" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 2) = Trim(Mid$(megrec$, 155, 10))
              Else
                Tabla.MSFlexGrid1.TextMatrix(vs&, 2) = Trim(Mid$(megrec$, 135, 10))
              End If
              Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Trim(Mid$(megrec$, 45, 30))
              szkod$ = Mid$(megrec$, 105, 15)
              szrec$ = dbxkey("KSZC", szkod$)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = postacim(szrec$, 121)
            End If
          End If
        End If
      End If
    Case "AUW-CJOV"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KSZF" Then
        If vs& = 1 Then
          Skod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
          szcrec$ = dbxkey("KSZC", Skod$)
          If szcrec$ <> "" Then
            Vektor.MSFlexGrid1.TextMatrix(2, 1) = Mid$(szcrec$, 16, 15)
          End If
        End If
        If vs& = 2 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(2, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            Vektor.MSFlexGrid1.TextMatrix(7, 1) = Mid$(partrec$, 328, 2)
          End If
        End If
      End If
      If objektum$ = "KSVL" Then
        If vo& = 1 And mezohiba% <> -1 Then
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
              form1.List2.Clear
              form1.List2.AddItem Mid$(termrec$, 16, 60)
              form1.List2.AddItem Mid$(termrec$, 196, 60)
              tagon$ = Mid$(termrec$, 1067, 15)
              If Trim(tagon$) <> "" Then
                gonrec$ = dbxkey("KTRM", tagon$)
                form1.List2.AddItem Mid$(gonrec$, 16, 60)
                If arkateg = 2 Or arkateg = 3 Then gonkateg% = 2 Else gonkateg% = 1
                gelar@ = cbaaraz(Mid$(gonrec$, 1, 15), Mid$(partrec$, 1, 15), maidatum$, gonkateg, szx$)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 2) = ertszam(Str$(gelar@), 12, 2)
              End If
              Call rkeszlet(tkod$, minta$, Mid$(krakrec$, 1, 4), keszme@, foglme@)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = Mid$(termrec$, 484, 6)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = Trim(Mid$(termrec$, 552, 14))
              Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = Mid$(termrec$, 706, 2)
              ellszi$ = Mid$(termrec$, 716, 24)
              If Trim(ellszi$) = "" Then ellszi$ = Mid$(krakrec$, 151, 24)
              If Trim(ellszi$) <> "" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = Mid$(ellszi$, 1, 8)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = Mid$(ellszi$, 9, 8)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 10) = Mid$(ellszi$, 17, 8)
              End If
              If Trim(Mid$(partrec$, 783, 8)) <> "" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = Mid$(partrec$, 783, 8)
              End If
              elar@ = cbaaraz(Mid$(termrec$, 1, 15), Mid$(partrec$, 1, 15), maidatum$, arkateg, diszx$)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 6) = ertszam(Str$(elar@), 12, 2)
            End If
          End If
        End If
        If vo& = 1 Or vo& = 2 Or vo& = 3 Or vo& = 5 Or vo& = 6 Or vo& = 7 Then
          '--- számlaéerék kiszámítása
          elert@ = 0: afa@ = 0
          For i13% = 1 To 999
            tkod$ = Tabla.MSFlexGrid1.TextMatrix(i13%, 1)
            If Trim$(tkod$) <> "" Then
              afakod$ = Tabla.MSFlexGrid1.TextMatrix(i13%, 7)
              If afakod$ = utafakod$ Then
                afakulcs@ = utafakulcs@
              Else
                afrec$ = dbxkey("PAFA", afakod$)
                afakulcs@ = xval(Mid$(afrec$, 33, 6))
                utafakod$ = afakod$
                utafakulcs@ = afakulcs@
              End If
              menny@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 3))
              gonar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 2))
              elar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 6))
              If menny@ <> 0 Then
                alaposz@ = (elar + gonar) * menny@
                elert@ = elert@ + (elar@ + gonar@) * menny@
              End If
              afaosz@ = (alaposz@ * afakulcs@) / 100
              afaker% = xval(Mid$(irec$, 345, 1))
              If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
              afaosz@ = xval(Format(afaosz@, fst$))
              afa@ = afa@ + afaosz@
            End If
          Next
          form1.List3.Clear
          form1.List3.AddItem langmodul(127) + ":" + ertszam(Str$(elert@), 14, 2)
          form1.List3.AddItem langmodul(128) + ":" + ertszam(Str$(elert@ + afa@), 14, 2)
        End If
      End If
    Case "AUW-ARMG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "CRM1" Then
        If vo& = 1 Then
          form1.List2.Clear
          tkkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If tkkod$ <> Space(15) Then
            If vs& > 1 Then
              For i95% = 1 To vs& - 1
                tkkod1$ = Left(Tabla.MSFlexGrid1.TextMatrix(i95%, 1) + Space$(15), 15)
                If tkkod1$ = tkkod$ Then
                  Call mess("Egy termék kód csak egyszer szerepelhet!", 3, 0, "Hiba", valasz%)
                  mezohiba% = 1
                End If
              Next
            End If
            termrec$ = dbxkey("KTRM", tkkod$)
            If termrec$ <> "" Then
              form1.List2.AddItem "Termék: " + Mid$(termrec$, 16, 60)
              xx$ = "": zz$ = ""
              For i81% = 1 To 7
                nv$ = Trim(Mid$(ikonfrec$, (i81% - 1) * 27 + 5, 20))
                ar$ = Trim(Mid$(termrec$, (i81% - 1) * 14 + 580, 14))
                If i81% < 5 Then
                  xx$ = xx$ + nv$ + ":" + ar$ + "     "
                Else
                  yy$ = yy$ + nv$ + ":" + ar$ + "     "
                End If
              Next
              nv$ = "Disztr.ár:"
              ar$ = Trim(Mid$(termrec$, 893, 14))
              yy$ = yy$ + nv$ + ":" + ar$ + "     "
              nv$ = "Fogy.ár:"
              ar$ = Trim(Mid$(termrec$, 678, 14))
              yy$ = yy$ + nv$ + ":" + ar$
              form1.List2.AddItem xx$
              form1.List2.AddItem yy$
            End If
          Else
            form1.List2.Clear
          End If
        End If
      End If
      If objektum$ = "ARKZ" Then
        If vs& = 1 Then
          tkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          termrec$ = dbxkey("KTRM", tkod$)
          Vektor.MSFlexGrid1.TextMatrix(2, 1) = Mid$(termrec$, 16, 40)
        End If
        If vs& = 6 Then
          ref@ = xval(Vektor.MSFlexGrid1.TextMatrix(6, 1))
          tkkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
          If tkkod$ <> Space(15) Then
            termrec$ = dbxkey("KTRM", tkkod$)
            If termrec$ <> "" Then
              tcsk$ = Mid$(termrec$, 438, 4)
              tcsrec$ = dbxkey("KCSP", tcsk$)
              utb@ = xval(Mid$(termrec$, 566, 14))
              bea@ = xval(Mid$(termrec$, 552, 14))
              If Mid$(ikonfrec$, 194, 1) = "B" Then alap@ = bea@
              If Mid$(ikonfrec$, 194, 1) = "U" Then alap@ = utb@
              If Mid$(ikonfrec$, 194, 1) = "R" Then alap@ = ref@
              For i81% = 1 To 6
                szorz@ = 1
                If tcsrec$ <> "" Then
                  szorz@ = (100 + xval(Mid(tcsrec$, (i81% - 1) * 10 + 75, 10))) / 100
                End If
                If szorz@ = 1 Then
                  szorz@ = xval(Mid(ikonfrec$, (i81% - 1) * 27 + 25, 7))
                End If
                uaa@ = alap@ * szorz@
                Vektor.MSFlexGrid1.TextMatrix(i81% + 6, 1) = ertszam(Str(uaa@), 10, 2)
              Next
              If Mid$(ikonfrec$, 202, 1) = "B" Then alap@ = bea@
              If Mid$(ikonfrec$, 202, 1) = "U" Then alap@ = utb@
              If Mid$(ikonfrec$, 202, 1) = "R" Then alap@ = ref@
              If Mid$(termrec$, 1247, 1) = "I" Then
                szorz@ = xval(Mid(ikonfrec$, 195, 7))
                uaa@ = alap@ * szorz@
                Vektor.MSFlexGrid1.TextMatrix(13, 1) = ertszam(Str(uaa@), 10, 2)
              Else
                Vektor.MSFlexGrid1.TextMatrix(13, 1) = ""
              End If
              If Mid$(ikonfrec$, 210, 1) = "B" Then alap@ = bea@
              If Mid$(ikonfrec$, 210, 1) = "U" Then alap@ = utb@
              If Mid$(ikonfrec$, 210, 1) = "R" Then alap@ = ref@
              If tcsrec$ <> "" Then
                szorz@ = (100 + xval(Mid(tcsrec$, 145, 10))) / 100
              Else
                szorz@ = 1
              End If
              If szorz@ = 1 Then szorz@ = xval(Mid(ikonfrec$, 203, 7))
              uaa@ = afasbrutto(alap@ * szorz@, termrec$)
              Vektor.MSFlexGrid1.TextMatrix(14, 1) = ertszam(Str(uaa@), 10, 2)
            End If
          End If
        End If
      End If
      If objektum$ = "AKCF" Then
        If vs& = 11 Or vs& = 12 Then
          gykod$ = Left(Vektor.MSFlexGrid1.TextMatrix(11, 1) + Space$(15), 15)
          szkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(12, 1) + Space$(15), 15)
          form1.List4.Clear
          If gykod$ <> Space(15) Then
            '--- gyártó
            partrec$ = dbxkey("PART", gykod$)
            If partrec$ <> "" Then
              form1.List4.AddItem "Gyártó: " + Mid$(partrec$, 16, 60)
            End If
          End If
          If partrec$ = "" Or Trim(gykod$) = "" Then
            form1.List4.AddItem "Gyártóhoz nem kötött."
          End If
          If szkod$ <> Space(15) Then
            '--- szállító
            partrec$ = dbxkey("PART", szkod$)
            If partrec$ <> "" Then
              form1.List4.AddItem "Szállító: " + Mid$(partrec$, 16, 60)
            End If
          End If
          If partrec$ = "" Or Trim(szkod$) = "" Then
            form1.List4.AddItem "Szállítóhoz  nem kötött."
          End If
        End If
      End If
      If objektum$ = "AKCS" Then
        If vo& = 3 Then
          ref@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 3))
          tkkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If tkkod$ <> Space(15) Then
            termrec$ = dbxkey("KTRM", tkkod$)
            utb@ = xval(Mid$(termrec$, 566, 14))
            bea@ = xval(Mid$(termrec$, 552, 14))
            If termrec$ <> "" Then
              tcsk$ = Mid$(termrec$, 438, 4)
              tcsrec$ = dbxkey("KCSP", tcsk$)
              '--- árkalkuláció
              If Mid$(ikonfrec$, 194, 1) = "B" Then alap@ = bea@
              If Mid$(ikonfrec$, 194, 1) = "U" Then alap@ = utb@
              If Mid$(ikonfrec$, 194, 1) = "R" Then alap@ = ref@
              For i81% = 1 To 6
                szorz@ = 1
                If tcsrec$ <> "" Then
                  szorz@ = (100 + xval(Mid(tcsrec$, (i81% - 1) * 10 + 75, 10))) / 100
                End If
                If szorz@ = 1 Then
                  szorz@ = xval(Mid(ikonfrec$, (i81% - 1) * 27 + 25, 7))
                End If
                uaa@ = alap@ * szorz@
                Tabla.MSFlexGrid1.TextMatrix(vs&, i81% + 3) = ertszam(Str(uaa@), 10, 2)
              Next
              If Mid$(ikonfrec$, 202, 1) = "B" Then alap@ = bea@
              If Mid$(ikonfrec$, 202, 1) = "U" Then alap@ = utb@
              If Mid$(ikonfrec$, 202, 1) = "R" Then alap@ = ref@
              If Mid$(termrec$, 1247, 1) = "I" Then
                szorz@ = xval(Mid(ikonfrec$, 195, 7))
                uaa@ = alap@ * szorz@
                Tabla.MSFlexGrid1.TextMatrix(vs&, 10) = ertszam(Str(uaa@), 10, 2)
              Else
                Tabla.MSFlexGrid1.TextMatrix(vs&, 10) = ""
              End If
              If Mid$(ikonfrec$, 210, 1) = "B" Then alap@ = bea@
              If Mid$(ikonfrec$, 210, 1) = "U" Then alap@ = utb@
              If Mid$(ikonfrec$, 210, 1) = "R" Then alap@ = ref@
              If tcsrec$ <> "" Then
                szorz@ = (100 + xval(Mid(tcsrec$, 145, 10))) / 100
              Else
                szorz@ = 1
              End If
              If szorz@ = 1 Then szorz@ = xval(Mid(ikonfrec$, 203, 7))
              uaa@ = afasbrutto(alap@ * szorz@, termrec$)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 11) = ertszam(Str(uaa@), 10, 0)
            End If
          End If
        End If
        If vo& = 1 Then
          form1.List1.Clear
          form1.List2.Clear
          tkkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If tkkod$ <> Space(15) Then
            If vs& > 1 Then
              For i95% = 1 To vs& - 1
                tkkod1$ = Left(Tabla.MSFlexGrid1.TextMatrix(i95%, 1) + Space$(15), 15)
                If tkkod1$ = tkkod$ Then
                  Call mess("Egy termék kód csak egyszer szerepelhet!", 3, 0, "Hiba", valasz%)
                  mezohiba% = 1
                End If
              Next
            End If
            termrec$ = dbxkey("KTRM", tkkod$)
            If termrec$ <> "" Then
              form1.List1.AddItem "Termék: " + Mid$(termrec$, 16, 60)
              xx$ = "": zz$ = ""
              For i81% = 1 To 6
                nv$ = Trim(Mid$(ikonfrec$, (i81% - 1) * 27 + 5, 20))
                ar$ = Trim(Mid$(termrec$, (i81% - 1) * 14 + 580, 14))
                If i81% < 5 Then
                  xx$ = xx$ + nv$ + ":" + ar$ + "     "
                Else
                  yy$ = yy$ + nv$ + ":" + ar$ + "     "
                End If
              Next
              nv$ = "Disztr.ár:"
              ar$ = Trim(Mid$(termrec$, 893, 14))
              yy$ = yy$ + nv$ + ":" + ar$ + "     "
              nv$ = "Fogy.ár:"
              ar$ = Trim(Mid$(termrec$, 678, 14))
              yy$ = yy$ + nv$ + ":" + ar$
              form1.List1.AddItem xx$
              form1.List1.AddItem yy$
            End If
          End If
        End If
      End If
    Case "AUW-KOZR"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KOZR" And vs& = 2 And ablaksorszam = 1 Then
        partkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
        partrec$ = dbxkey("PART", partkod$)
        Vektor.MSFlexGrid1.TextMatrix(3, 1) = Mid$(partrec$, 16, 60)
        Vektor.MSFlexGrid1.TextMatrix(4, 1) = postacim(partrec$, 106)
        Vektor.MSFlexGrid1.TextMatrix(10, 1) = Mid$(partrec$, 184, 15)
        Vektor.MSFlexGrid1.TextMatrix(11, 1) = Mid$(partrec$, 214, 30)
        Vektor.MSFlexGrid1.TextMatrix(12, 1) = banktagol(Mid$(partrec$, 244, 24))
      End If
      If objektum$ = "KOZR" And vs& = 20 And ablaksorszam = 1 Then
        uzkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
        uzrec$ = dbxkey("KUZL", uzkod$)
        Vektor.MSFlexGrid1.TextMatrix(21, 1) = Mid$(uzrec$, 9, 60)
      End If
      If objektum$ = "KOZR" And vs& = 23 And ablaksorszam = 1 Then
        uzkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
        uzrec$ = dbxkey("KUZL", uzkod$)
        Vektor.MSFlexGrid1.TextMatrix(24, 1) = Mid$(uzrec$, 9, 60)
      End If
      If objektum$ = "KOZR" And vs& = 4 And ablaksorszam = 3 Then
        tkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(6), 6)
        trec$ = dbxkey("PTRM", tkod$)
        Vektor.MSFlexGrid1.TextMatrix(5, 1) = Mid$(trec$, 7, 60)
        Vektor.MSFlexGrid1.TextMatrix(13, 1) = Mid$(trec$, 140, 6)
        Vektor.MSFlexGrid1.TextMatrix(14, 1) = Mid$(trec$, 146, 12)
      End If
      If objektum$ = "KOZR" And vs& = 20 And ablaksorszam = 3 Then
        tkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(6), 6)
        trec$ = dbxkey("PTRM", tkod$)
        Vektor.MSFlexGrid1.TextMatrix(21, 1) = Mid$(trec$, 7, 60)
        Vektor.MSFlexGrid1.TextMatrix(23, 1) = Mid$(trec$, 140, 6)
        Vektor.MSFlexGrid1.TextMatrix(24, 1) = Mid$(trec$, 146, 12)
      End If
    Case "AUW-GSML"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "GSMF" And vs& = 3 Then
        szckod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
        If Trim(szckod$) <> "" Then
          szcrec$ = dbxkey("KSZC", szckod$)
          Vektor.MSFlexGrid1.TextMatrix(4, 1) = Mid$(szcrec$, 31, 60)
          Vektor.MSFlexGrid1.TextMatrix(5, 1) = Mid$(szcrec$, 16, 15)
        End If
      End If
      If objektum$ = "GSMF" And vs& = 5 Then
        partkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
        partrec$ = dbxkey("PART", partkod$)
        Vektor.MSFlexGrid1.TextMatrix(6, 1) = Mid$(partrec$, 16, 60)
        Vektor.MSFlexGrid1.TextMatrix(7, 1) = Trim(Mid$(partrec$, 106, 8)) + " " + Trim(Mid$(partrec$, 114, 30)) + " " + Trim(Mid$(partrec$, 144, 30)) + " " + Trim(Mid$(partrec$, 174, 10))
      End If
      If objektum$ = "GSMF" And vs& = 14 Then
        gepkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(20), 20)
        geprec$ = dbxkey("GGEP", gepkod$)
        Vektor.MSFlexGrid1.TextMatrix(15, 1) = Mid$(geprec$, 16, 60)
        Vektor.MSFlexGrid1.TextMatrix(16, 1) = Mid$(geprec$, 141, 40)
        Vektor.MSFlexGrid1.TextMatrix(17, 1) = Mid$(geprec$, 181, 40)
        Vektor.MSFlexGrid1.TextMatrix(18, 1) = Mid$(geprec$, 221, 30)
        Vektor.MSFlexGrid1.TextMatrix(19, 1) = Mid$(geprec$, 251, 30)
        Vektor.MSFlexGrid1.TextMatrix(20, 1) = Mid$(geprec$, 281, 9)
      End If
      If objektum$ = "GSM1" Then
        If vo& = 3 Then
          termkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
          If Trim(termkod$) <> "" Then
            tko$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 4) + Space$(15), 15)
            If Trim(tko$) <> "" Then
              Call mess(langmodul(107), 2, 0, langmodul(108), valasz%)
              'MsgBox langmodul(107), 48, langmodul(108)
              'MsgBox "Vagy anyag vagy mûvelet", 48, "Cikkszám hiba"
              mezohiba = 1
            Else
              termrec$ = dbxkey("KTRM", termkod$)
              If termrec$ <> "" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = Mid$(termrec$, 484, 6)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = Mid$(termrec$, 582, 12)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = Mid$(termrec$, 16, 40)
              End If
            End If
          End If
        End If
        If vo& = 4 Then
          termkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 4) + Space$(15), 15)
          If Trim(termkod$) <> "" Then
            tko$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
            If Trim(tko$) <> "" Then
              Call mess(langmodul(107), 2, 0, langmodul(109), valasz%)
              'MsgBox langmodul(107), 48, langmodul(109)
              'MsgBox "Vagy anyag vagy mûvelet", 48, "Mûvelet hiba"
              mezohiba = 1
            Else
              termrec$ = dbxkey("GYMV", termkod$)
              If termrec$ <> "" Then
                egysegar@ = xval(Mid$(termrec$, 394, 12))
                If egysegar@ <> 0 Then
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = Mid$(termrec$, 376, 6)
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = Mid$(termrec$, 394, 12)
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = Mid$(termrec$, 16, 40)
                Else
                  szakod$ = Mid$(termrec$, 389, 5)
                  If Trim(szakod$) <> "" Then
                    szakrec$ = dbxkey("SZAK", szakod$)
                    If szakrec$ <> "" Then
                      Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = Mid$(szakrec$, 66, 10)
                    End If
                  End If
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = langmodul(110)
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = Mid$(termrec$, 16, 40)
                End If
              End If
            End If
          End If
        End If
      End If
    Case "AUW-RJAR"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "RJA1" Then
        If vs& = 1 Then
          szckod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          szcrec$ = dbxkey("KSZC", szckod$)
          pkod$ = Mid$(szcrec$, 16, 15)
          sznev$ = Mid$(szcrec$, 31, 60)
          szcim$ = Trim(Mid$(szcrec$, 121, 8)) + " " + Trim(Mid$(szcrec$, 129, 30)) + ", " + Trim(Mid$(szcrec$, 159, 30)) + " " + Trim(Mid$(szcrec$, 189, 10))
          Vektor.MSFlexGrid1.TextMatrix(2, 1) = Mid$(szcrec$, 16, 15)
          Vektor.MSFlexGrid1.TextMatrix(3, 1) = sznev$
          Vektor.MSFlexGrid1.TextMatrix(4, 1) = Left(szcim$ + Space(60), 60)
        End If
      End If
    Case Else
  End Select
  If mezohiba% = -1 Then mezohiba% = 1
End Sub
Public Sub mezovege(vsor%, voszl%, mezohiba%, abmod%)
  '--- adatvége teendõk beépítése
  '--- objektum$ az obj.azonosító public
  '--- abmod%=1 ablak abmod%=0 táblázat
  '--- ablaksorszam% az ablak sorszáma public
  If programnev$ = "AUW-RFVU" Or programnev$ = "AUW-CJOV" Or programnev$ = "AUW-ARMG" Or programnev$ = "AUW-KOZR" Or programnev = "AUW-GSML" Or progamnev = "AUW-RJAR" Then
    Call mezovege2(vsor%, voszl%, mezohiba%, abmod%)
    Exit Sub
  Else
   
    If Left(programnev$, 5) = "AUW-Q" Then
      Call mezovege3(vsor%, voszl%, mezohiba%, abmod%)
      Exit Sub
    End If
  End If
  If mezohiba% <> -1 Then mezohiba% = 0
  Select Case programnev$
    Case "AUW-RTRM"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "RCPS" Then
        If vo& = 1 Then
          termkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          termrec$ = dbxkey("KTRM", termkod$)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 2) = Mid$(termrec$, 16, 60)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = Mid$(termrec$, 484, 6)
        End If
      End If
      If objektum$ = "REAN" Then
        If vs& = 2 Then
          termkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          termrec$ = dbxkey("KTRM", termkod$)
          Vektor.MSFlexGrid1.TextMatrix(3, 1) = Mid$(termrec$, 1250, 24)
        End If
      End If
    Case "AUW-PSMG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "PMEG" And vs& = 5 Then
        pkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
        prec$ = dbxkey("PART", pkod$)
        eng@ = xval(Mid$(prec$, 333, 6))
        ptrmkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(6), 6)
        ptrmrec$ = dbxkey("PTRM", ptrmkod$)
        Vektor.MSFlexGrid1.TextMatrix(6, 1) = Trim(Mid$(ptrmrec$, 7, 60))
        erar@ = xval(Mid$(ptrmrec$, 146, 12))
        If eng@ < 0 Then egar@ = erar@ - (erar@ * Abs(eng@) / 100) Else egar@ = erar@ + (erar@ * Abs(eng@) / 100)
        Vektor.MSFlexGrid1.TextMatrix(7, 1) = ertszam(Str(egar@), 14, 2)
        Vektor.MSFlexGrid1.TextMatrix(8, 1) = Mid$(ptrmrec$, 158, 3)
        Vektor.MSFlexGrid1.TextMatrix(10, 1) = Mid$(ptrmrec$, 140, 6)
      End If
    Case "AUW-JOV"
      Call jmezovege(vsor%, voszl%, mezohiba%, abmod%)
    Case "AUW-GYGT"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "GGEP" And vs& = 9 Then
        partkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
        If Trim(partkod$) <> "" Then
          partrec$ = dbxkey("PART", partkod$)
          Vektor.MSFlexGrid1.TextMatrix(10, 1) = Mid$(partrec$, 16, 60)
          Vektor.MSFlexGrid1.TextMatrix(11, 1) = Trim(Mid$(partrec$, 106, 8)) + " " + Trim(Mid$(partrec$, 114, 30)) + " " + Trim(Mid$(partrec$, 144, 30)) + " " + Trim(Mid$(partrec$, 174, 10))
        End If
      End If
    Case "AUW-GYSK"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "GYAA" And vs& = 2 Then
        termkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
        If Trim(termkod$) <> "" Then
          termrec$ = dbxkey("KTRM", termkod$)
          form1.List2.Clear
          form1.List2.AddItem termkod$
          form1.List2.AddItem Trim(Mid$(termrec$, 16, 60)) + " " + Trim(Mid$(termrec$, 484, 6))
          form1.List2.AddItem Mid$(termrec$, 196, 60)
          form1.List2.AddItem Mid$(termrec$, 256, 60)
          form1.List2.AddItem Mid$(termrec$, 316, 60)
          form1.List2.AddItem Mid$(termrec$, 376, 60)
          xw1$ = Trim(Vektor.MSFlexGrid1.TextMatrix(6, 1))
          If xw1$ = "" Then Vektor.MSFlexGrid1.TextMatrix(6, 1) = Mid$(termrec$, 196, 60)
          xw1$ = Trim(Vektor.MSFlexGrid1.TextMatrix(7, 1))
          If xw1$ = "" Then Vektor.MSFlexGrid1.TextMatrix(7, 1) = Mid$(termrec$, 256, 60)
          xw1$ = Trim(Vektor.MSFlexGrid1.TextMatrix(8, 1))
          If xw1$ = "" Then Vektor.MSFlexGrid1.TextMatrix(8, 1) = Mid$(termrec$, 316, 60)
          xw1$ = Trim(Vektor.MSFlexGrid1.TextMatrix(9, 1))
          If xw1$ = "" Then Vektor.MSFlexGrid1.TextMatrix(9, 1) = Mid$(termrec$, 376, 60)
        End If
      End If
    Case "AUW-DOK"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "DOKU" And ablaksorszam% = 1 Then
        If vs& = 4 Then
          partkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If Trim(partkod$) <> "" Then
            partrec$ = dbxkey("PART", partkod$)
            Vektor.MSFlexGrid1.TextMatrix(9, 1) = Mid$(partrec$, 16, 60)
          End If
        End If
        If vs& = 7 Then
          partkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If Trim(partkod$) <> "" Then
            partrec$ = dbxkey("KTRM", partkod$)
            Vektor.MSFlexGrid1.TextMatrix(9, 1) = Mid$(partrec$, 16, 60)
          End If
        End If
        If vs& = 8 Then
          partkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If Trim(partkod$) <> "" Then
            partrec$ = dbxkey("ESZK", partkod$)
            Vektor.MSFlexGrid1.TextMatrix(9, 1) = Mid$(partrec$, 16, 60)
          End If
        End If
      End If
    Case "AUW-RSMG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KSMG" And ablaksorszam% = 1 Then
        If vs& = 1 Then
          partkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", partkod$)
          Vektor.MSFlexGrid1.TextMatrix(2, 1) = Mid$(partrec$, 16, 60)
          Vektor.MSFlexGrid1.TextMatrix(3, 1) = Mid$(partrec$, 76, 30)
          Vektor.MSFlexGrid1.TextMatrix(4, 1) = Mid$(partrec$, 106, 8)
          Vektor.MSFlexGrid1.TextMatrix(5, 1) = Mid$(partrec$, 114, 30)
          Vektor.MSFlexGrid1.TextMatrix(6, 1) = Mid$(partrec$, 144, 30)
          Vektor.MSFlexGrid1.TextMatrix(7, 1) = Mid$(partrec$, 174, 10)
        End If
      End If
      If objektum$ = "KSM1" Then
        tetelszam% = 0
        ertek@ = 0
        For i93% = 1 To 100
          If xval(Tabla.MSFlexGrid1.TextMatrix(i93%, 3)) <> 0 Then tetelszam% = tetelszam% + 1
          ertek@ = ertek@ + xval(Tabla.MSFlexGrid1.TextMatrix(i93%, 3)) * xval(Tabla.MSFlexGrid1.TextMatrix(i93%, 4))
        Next
        form1.List2.Clear
        form1.List2.AddItem langmodul(111) + ":" + Str(tetelszam%)
        'form1.List2.AddItem "Tételszám:" + Str(tetelszam%)
        form1.List2.AddItem langmodul(112) + ":" + ertszam(Str(ertek), 14, 2) + " " + langmodul(113)
        'form1.List2.AddItem "Összes nettó érték:" + ertszam(Str(ertek), 14, 2) + " Ft"
        If vo& = 1 Then
          termkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If Trim(termkod$) <> "" Then
            termrec$ = dbxkey("KTRM", termkod$)
            refar@ = xval(Mid$(termrec$, 1274, 14))
            form1.List1.Clear
            form1.List1.AddItem langmodul(85) + ":" + Trim(termkod$) + "  " + langmodul(86) + ":" + vevocikk$
            'form1.List1.AddItem "Termék kód:" + Trim(termkod$) + "  Vevõi cikkszám:" + vevocikk$
            form1.List1.AddItem Mid$(termrec$, 16, 60)
            form1.List1.AddItem Mid$(termrec$, 196, 60)
            form1.List1.AddItem langmodul(87) + " :" + Mid$(termrec$, 552, 14)
            'form1.List1.AddItem "Nyilv.ár :" + Mid$(termrec$, 552, 14)
            If Val(Tabla.MSFlexGrid1.TextMatrix(vs&, 4)) = 0 Then
              Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = ertszam(Mid$(termrec$, 566, 14), 12, 2)
            End If
          End If
        End If
      End If
    Case "AUW-RMEG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KMEG" And ablaksorszam% = 1 Then
        If vs& = 2 Then
          szckod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If Trim(szckod$) <> "" Then
            kszcrec$ = dbxkey("KSZC", szckod$)
            partkod$ = Mid$(kszcrec$, 16, 15)
            partrec$ = dbxkey("PART", partkod$)
            Vektor.MSFlexGrid1.TextMatrix(3, 1) = partkod$
            Vektor.MSFlexGrid1.TextMatrix(4, 1) = "    "
            Vektor.MSFlexGrid1.TextMatrix(5, 1) = Mid$(kszcrec$, 31, 60)
            Vektor.MSFlexGrid1.TextMatrix(6, 1) = Mid$(kszcrec$, 91, 30)
            Vektor.MSFlexGrid1.TextMatrix(7, 1) = Mid$(kszcrec$, 121, 8)
            Vektor.MSFlexGrid1.TextMatrix(8, 1) = Mid$(kszcrec$, 129, 30)
            Vektor.MSFlexGrid1.TextMatrix(9, 1) = Mid$(kszcrec$, 159, 30)
            Vektor.MSFlexGrid1.TextMatrix(10, 1) = Mid$(kszcrec$, 189, 10)
            Vektor.MSFlexGrid1.TextMatrix(11, 1) = Mid$(kszcrec$, 200, 8)
            Vektor.MSFlexGrid1.TextMatrix(17, 1) = Mid$(kszcrec$, 199, 1)
            Vektor.MSFlexGrid1.TextMatrix(19, 1) = Mid$(partrec$, 782, 1)
            Vektor.MSFlexGrid1.TextMatrix(20, 1) = Mid$(partrec$, 328, 2)
            Vektor.MSFlexGrid1.TextMatrix(21, 1) = Mid$(partrec$, 330, 3)
          End If
        End If
        If vs& = 3 Then
          If Trim(Vektor.MSFlexGrid1.TextMatrix(2, 1)) = "" Then
            partkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
            If Trim(partkod$) <> "" Then
              partrec$ = dbxkey("PART", partkod$)
              Vektor.MSFlexGrid1.TextMatrix(4, 1) = "    "
              Vektor.MSFlexGrid1.TextMatrix(5, 1) = Mid$(partrec$, 16, 60)
              Vektor.MSFlexGrid1.TextMatrix(6, 1) = Mid$(partrec$, 76, 30)
              Vektor.MSFlexGrid1.TextMatrix(7, 1) = Mid$(partrec$, 106, 8)
              Vektor.MSFlexGrid1.TextMatrix(8, 1) = Mid$(partrec$, 114, 30)
              Vektor.MSFlexGrid1.TextMatrix(9, 1) = Mid$(partrec$, 144, 30)
              Vektor.MSFlexGrid1.TextMatrix(10, 1) = Mid$(partrec$, 174, 10)
              Vektor.MSFlexGrid1.TextMatrix(11, 1) = Mid$(partrec$, 319, 8)
              Vektor.MSFlexGrid1.TextMatrix(17, 1) = "S"
              Vektor.MSFlexGrid1.TextMatrix(19, 1) = Mid$(partrec$, 782, 1)
              Vektor.MSFlexGrid1.TextMatrix(20, 1) = Mid$(partrec$, 328, 2)
              Vektor.MSFlexGrid1.TextMatrix(21, 1) = Mid$(partrec$, 330, 3)
            End If
          End If
        End If
        If vs& = 4 Then
          If Trim(Vektor.MSFlexGrid1.TextMatrix(4, 1)) <> "" Then
            krakkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(4), 4)
            If Trim(krakkod$) <> "" Then
              krakrec$ = dbxkey("KRAK", krakkod$)
              Vektor.MSFlexGrid1.TextMatrix(2, 1) = Space$(15)
              Vektor.MSFlexGrid1.TextMatrix(3, 1) = Space$(15)
              Vektor.MSFlexGrid1.TextMatrix(5, 1) = Mid$(krakrec$, 5, 60)
              Vektor.MSFlexGrid1.TextMatrix(6, 1) = Space$(30)
              Vektor.MSFlexGrid1.TextMatrix(7, 1) = Space$(8)
              Vektor.MSFlexGrid1.TextMatrix(8, 1) = Space$(30)
              Vektor.MSFlexGrid1.TextMatrix(9, 1) = Space$(30)
              Vektor.MSFlexGrid1.TextMatrix(10, 1) = Space$(10)
              Vektor.MSFlexGrid1.TextMatrix(11, 1) = Space$(8)
              Vektor.MSFlexGrid1.TextMatrix(17, 1) = "L"
              Vektor.MSFlexGrid1.TextMatrix(19, 1) = " "
              Vektor.MSFlexGrid1.TextMatrix(20, 1) = "  "
              Vektor.MSFlexGrid1.TextMatrix(21, 1) = "   "
            End If
          End If
        End If
      End If
      If objektum$ = "KME1" Then
        tetelszam% = 0
        ertek@ = 0
        For i93% = 1 To 200
          If xval(Tabla.MSFlexGrid1.TextMatrix(i93%, 3)) <> 0 Then tetelszam% = tetelszam% + 1
          ertek@ = ertek@ + xval(Tabla.MSFlexGrid1.TextMatrix(i93%, 3)) * xval(Tabla.MSFlexGrid1.TextMatrix(i93%, 4))
        Next
        form1.List2.Clear
        'form1.List2.AddItem " "
        If Mid$(form1.kmegrec, 54, 4) = "    " Then
          form1.List2.AddItem langmodul(114)
          'form1.List2.AddItem "Külsõ megrendelés"
        Else
          form1.List2.AddItem langmodul(115)
        End If
        If Mid$(form1.kmegrec, 192, 1) = "D" Then
          form1.List2.AddItem langmodul(116)
        Else
          form1.List2.AddItem langmodul(117)
          'form1.List2.AddItem "Nem foglal készletet!"
        End If
        form1.List2.AddItem langmodul(111) + ":" + Str(tetelszam%)
        'form1.List2.AddItem "Tételszám:" + Str(tetelszam%)
        form1.List2.AddItem langmodul(112) + ":" + ertszam(Str(ertek), 14, 2) + " " + langmodul(113)
        'form1.List2.AddItem "Összes nettó érték:" + ertszam(Str(ertek), 14, 2) + " Ft"
        If vo& = 1 Then
          termkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          If Trim(termkod$) <> "" Then
            termrec$ = dbxkey("KTRM", termkod$)
            armind$ = Mid$(termrec$, 1, 15) + Mid$(partrec$, 1, 15)
            armgrec$ = dbxkey("ARMG", armind$)
            vevocikk$ = ""
            If armgrec$ <> "" Then vevocikk$ = Trim(Mid(armgrec$, 31, 15))
            form1.List1.Clear
            form1.List1.AddItem langmodul(85) + ":" + Trim(termkod$) + "  " + langmodul(86) + ":" + vevocikk$
            'form1.List1.AddItem "Termék kód:" + Trim(termkod$) + "  Vevõi cikkszám:" + vevocikk$
            form1.List1.AddItem Mid$(termrec$, 16, 60)
            form1.List1.AddItem Mid$(termrec$, 196, 60)
            form1.List1.AddItem langmodul(88) + " 1:" + Mid$(termrec$, 580, 14)
            'form1.List1.AddItem "Listaár 1:" + Mid$(termrec$, 580, 14)
            If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
              Call rkeszlet(termkod$, "", Mid$(rakrec$, 1, 4), keszme@, foglme@)
              form1.List1.AddItem langmodul(89) + ":" + ertszam(Str(keszme@), 14, 2) + " " + langmodul(90) + ":" + ertszam(Str(foglme@), 14, 2) + " " + langmodul(91) + ":" + ertszam(Str(keszme@ - foglme@), 14, 2)
              'form1.List1.AddItem "Raktárkészlet:" + ertszam(Str(keszme@), 14, 2) + " Foglalt készlet:" + ertszam(Str(foglme@), 14, 2) + " Szabad készlet:" + ertszam(Str(keszme@ - foglme@), 14, 2)
              form1.List1.AddItem langmodul(92) + ":" + Mid$(termrec$, 542, 10)
              'form1.List1.AddItem "Rendelhetõ mennyiség:" + Mid$(termrec$, 542, 10)
              If Mid$(termrec$, 443, 1) = "M" And Mid$(rakrec$, 67, 1) = "I" Then
                form1.List1.AddItem langmodul(93)
                'form1.List1.AddItem "Mintaszám kötelezõ"
                '--- mintaszám táblázat mgejlenítése
                rkszind$ = Mid$(rakrec$, 1, 4) + Mid$(termrec$, 1, 15)
                rkszrec$ = dbxkey("RKSZ", rkszind$)
                form1.MSFlexGrid4.Clear
                form1.MSFlexGrid4.Cols = 2: form1.MSFlexGrid4.FixedCols = 0
                form1.MSFlexGrid4.Rows = 2: form1.MSFlexGrid4.FixedRows = 1
                form1.MSFlexGrid4.TextMatrix(0, 0) = langmodul(118)
                form1.MSFlexGrid4.ColWidth(0) = 1100
                form1.MSFlexGrid4.TextMatrix(0, 1) = langmodul(92)
                form1.MSFlexGrid4.ColWidth(1) = 2000
                For i911% = 1 To 300
                  ele3$ = Mid$(rkszrec$, (i911% - 1) * 28 + 200, 28)
                  If xval(Mid$(ele3$, 9, 10)) - xval(Mid$(ele3$, 19, 10)) > 0 Then
                    form1.MSFlexGrid4.TextMatrix(form1.MSFlexGrid4.Rows - 1, 0) = Mid(ele3$, 1, 8)
                    form1.MSFlexGrid4.TextMatrix(form1.MSFlexGrid4.Rows - 1, 1) = xval(Mid$(ele3$, 9, 10)) - xval(Mid$(ele3$, 19, 10))
                    form1.MSFlexGrid4.Rows = form1.MSFlexGrid4.Rows + 1
                  End If
                Next
                If form1.MSFlexGrid4.Rows > 2 Then form1.MSFlexGrid4.Rows = form1.MSFlexGrid4.Rows - 1
                form1.MSFlexGrid4.Visible = True
              Else
                form1.MSFlexGrid4.Visible = False
                form1.List1.AddItem langmodul(94)
                'form1.List1.AddItem "Mintaszám tilos"
              End If
            Else
              form1.List1.AddItem langmodul(95)
              'form1.List1.AddItem "Nincs készletnyilvántartás"
            End If
            '--- árazás
            minta$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
            diszt$ = Trim(Tabla.MSFlexGrid1.TextMatrix(vs&, 6))
            egysar@ = arazo@(Mid$(partrec$, 1, 15), termkod$, minta$, erteknap$, diszt$)
            Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = ertszam(Str(egysar), 12, 2)
          End If
        End If
        If vo& = 2 Then
          form1.MSFlexGrid4.Visible = False
          termkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          minta$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
          If Trim(termkod$) <> "" And Trim(minta$) <> "" Then
            termrec$ = dbxkey("KTRM", termkod$)
            form1.List1.Clear
            form1.List1.AddItem Mid$(termrec$, 16, 60)
            form1.List1.AddItem Mid$(termrec$, 196, 60)
            form1.List1.AddItem langmodul(88) + " 1:" + Mid$(termrec$, 580, 14)
            'form1.List1.AddItem "Listaár 1:" + Mid$(termrec$, 580, 14)
            If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
              Call rkeszlet(termkod$, minta$, Mid$(rakrec$, 1, 4), keszme@, foglme@)
              form1.List1.AddItem langmodul(89) + ":" + ertszam(Str(keszme@), 14, 2) + " " + langmodul(90) + ":" + ertszam(Str(foglme@), 14, 2) + " " + langmodul(91) + ":" + ertszam(Str(keszme@ - foglme@), 14, 2)
              'form1.List1.AddItem "Raktárkészlet:" + ertszam(Str(keszme@), 14, 2) + " Foglalt készlet:" + ertszam(Str(foglme@), 14, 2) + " Szabad készlet:" + ertszam(Str(keszme@ - foglme@), 14, 2)
            Else
              form1.List1.AddItem langmodul(95)
              'form1.List1.AddItem "Nincs készletnyilvántartás"
            End If
            '--- árazás
            minta$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
            diszt$ = Trim(Tabla.MSFlexGrid1.TextMatrix(vs&, 6))
            egysar@ = arazo@(Mid$(partrec$, 1, 15), termkod$, minta$, erteknap$, diszt$)
            Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = ertszam(Str(egysar), 12, 2)
          End If
        End If
      End If
    Case "AUW-REGY", "AUW-CEGY", "AUW-CVIS"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KKB3" Then
        If vs& = 6 Then
          pkod$ = Trim(Vektor.MSFlexGrid1.TextMatrix(vs&, 1))
          If pkod$ <> "" Then
            paxrec$ = dbxkey("PART", pkod$)
            If paxrec$ <> "" Then
              If programnev$ <> "AUW-CVIS" Then
                form1.List2.Clear
                form1.List2.Visible = True
                form1.List2.AddItem Mid$(paxrec$, 16, 60)
                form1.List2.AddItem postacim(paxrec$, 106)
                form1.Text3.Text = Trim(form1.Text3.Text) + Chr(13) + Chr(10) + Trim(Mid$(paxrec$, 16, 60))
              End If
            End If
          End If
        End If
        If vs& = 5 Then
          mkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + "   ", 3)
          mozrec$ = dbxkey("KMOX", mkod$)
          If mozrec$ <> "" Then
            If programnev$ <> "AUW-CVIS" Then
              form1.Text5.Text = mkod$ + " " + Mid$(mozrec$, 4, 30)
              form1.Text5.Visible = True
            End If
          End If
        End If
        If vs& = 4 Then
          rkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + "    ", 4)
          rakrec$ = dbxkey("KRAK", rkod$)
          If rakrec$ <> "" Then
            sz$ = ""
            If Mid$(rakrec$, 66, 1) = "I" Then sz$ = sz$ + " " + langmodul(120)
            If Mid$(rakrec$, 67, 1) = "I" Then sz$ = sz$ + " " + langmodul(121)
            If programnev$ <> "AUW-CVIS" Then
              form1.Text3.Text = rkod$ + " " + Trim(Mid$(rakrec$, 5, 60)) + sz$
              form1.Text3.Visible = True
            End If
          End If
        End If
      End If
      If objektum$ = "KKF1" Or objektum$ = "KKF3" Then
        If objektum$ = "KKF1" Then
          If vo& = 1 Then
            mkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + "   ", 3)
            mozrec$ = dbxkey("KMOX", mkod$)
            If mozrec$ <> "" Then
              If programnev$ <> "AUW-CVIS" Then
                form1.Text5.Text = mkod$ + " " + Mid$(mozrec$, 4, 30)
                form1.Text5.Visible = True
              End If
            End If
          End If
          If vo& = 2 Then
            rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
            rakrec$ = dbxkey("KRAK", rkod$)
            If rakrec$ <> "" Then
              sz$ = ""
              If Mid$(rakrec$, 66, 1) = "I" Then sz$ = sz$ + " " + langmodul(120)
              If Mid$(rakrec$, 67, 1) = "I" Then sz$ = sz$ + " " + langmodul(121)
              If programnev$ <> "AUW-CVIS" Then
                form1.Text3.Text = rkod$ + " " + Trim(Mid$(rakrec$, 5, 60)) + sz$
                form1.Text3.Visible = True
              End If
            End If
          End If
          If vo& = 4 Then
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
            If Trim$(tkod$) <> "" Then
              termrec$ = dbxkey("KTRM", tkod$)
              If termrec$ <> "" Then
                '--- táblás
                If Mid$(termrec$, 443, 1) = "M" And Mid$(termrec$, 849, 1) = "T" Then
                  valtoszam@ = xval(Mid$(termrec$, 850, 7))
                  mintaszam$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 4) + Space$(8), 8)
                  tszel@ = xval(Mid$(mintaszam$, 1, 4)) / 100
                  tmag@ = xval(Mid$(mintaszam$, 5, 4)) / 100
                  kilodb@ = tszel@ * tmag@ * valtoszam@
                  akesz@ = xval(Mid$(termrec$, 748, 14))
                  If kilodb@ <> 0 Then adarab@ = akesz@ / kilodb@ Else adarab@ = 0
                  Tblanyg.Label4 = Left(mintaszam$, 4) + "x" + Mid$(mintaszam$, 5, 4)
                  Tblanyg.Label6 = Str(valtoszam@)
                  Tblanyg.Label8 = Str(kilodb@)
                  Tblanyg.Label13 = Str(akesz@) + " kg /" + Str(adarab@) + " db"
                  Tblanyg.Show vbModal
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = ertszam(Str(tablamennyiseg@ * kilodb@), 12, 3)
                End If
              End If
            End If
          End If
        End If
        If objektum$ = "KKF3" And vo& = 3 And Arszorzo1 = 1 Then
          Arszorzo.Text1.Text = ""
          Arszorzo.Text2.Text = ""
          Arszorzo.Text3.Text = ""
          Arszorzo.Label5.Caption = ""
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          termrec$ = dbxkey("KTRM", tkod$)
          tcs$ = Mid$(termrec$, 438, 4)
          tcsrec$ = dbxkey("KCSP", tcs$)
          If tcsrec$ <> "" Then
            Arszorzo.Text2.Text = Trim(Mid$(tcsrec$, 145, 10))
          End If
          bear@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 3))
          Arszorzo.Label1.Caption = Mid$(termrec$, 16, 40)
          Arszorzo.Label5.Caption = "Elõzõ fogy ár: " + Trim(Mid$(termrec$, 678, 14))
          Arszorzo.Text1.Text = Str(bear@)
          Arszorzo.Show vbModal
          If wrogzites = 1 Then
            Call dbxtrkezd("KTRM")
            Mid$(termrec$, 678, 14) = Right(Space(14) + Trim(Arszorzo.Text3.Text), 14)
            Call dbxki("KTRM", termrec$, ";", "", "", hiba%)
            Call dbxtrvege
          End If
        End If
        If objektum$ = "KKF1" And vo& = 3 Or objektum$ = "KKF3" And vo& = 1 Then
          form1.List2.Clear
          If objektum$ = "KKF1" Then
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
          Else
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          End If
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
              If programnev$ = "AUW-CEGY" Or programnev$ = "AUW-CVIS" Then
                refar@ = beszaraz(tkod, arpartkod$, arbizdatum$)
                If objektum$ = "KXF1" Then
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = ertszam(Str(refar@), 12, 2)
                Else
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = ertszam(Str(refar@), 12, 2)
                End If
                gonkod$ = Mid$(termrec$, 1067, 15)
                If Trim(gonkod$) <> "" Then
                  gonrec$ = dbxkey("KTRM", gonkod$)
                  If gonrec$ <> "" Then
                    gonrefar@ = beszaraz(gonkod$, arpartkod$, arbizdatum$)
                    Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = ertszam(Str(gonrefar@), 8, 2)
                  End If
                End If
                form1.List2.Clear
                form1.List2.AddItem Mid$(termrec$, 16, 60)
                kartondb% = xval(Mid$(termrec$, 1226, 7))
                If kartondb% > 1 Then
                  xyx$ = "K" + Trim(Str(kartondb%))
                Else
                  xyx$ = " "
                End If
                form1.List2.AddItem xyx$
                form1.List2.AddItem "Nyilv.ár: " + Trim(Mid$(termrec$, 554, 12)) + " Ref.ár: " + Trim(Mid$(termrec$, 1276, 12)) + " Nagyk.ár: " + Trim(Mid$(termrec$, 582, 12)) + " Fogy.ár: " + Trim(Mid$(termrec$, 680, 12))
                form1.List2.AddItem "Akt.készl: " + Trim(Mid$(termrec$, 750, 12)) + " Közp.rak: " + Trim(Mid$(termrec$, 929, 12)) + " Vegyi: " + Trim(Mid$(termrec$, 943, 12)) + " C+C: " + Trim(Mid$(termrec$, 957, 12))
              Else
                form1.List2.Clear
                form1.List2.AddItem Mid$(termrec$, 16, 60)
                form1.List2.AddItem Mid$(termrec$, 196, 60)
                Select Case Mid$(termrec$, 442, 1)
                  Case "A", "R", "F", "K"
                    If Mid$(termrec$, 443, 1) <> "N" Then
                      form1.List2.AddItem langmodul(96) + ": " + Trim$(ertszam(Mid$(termrec$, 748, 14), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6)) + " " + langmodul(97) + ":" + Trim$(ertszam(Mid$(termrec$, 762, 14), 14, 3)) + " " + langmodul(87) + ":" + Trim(ertszam(Mid$(termrec$, 552, 14), 14, 2))
                    Else
                      Call mess(langmodul(98), 2, 0, langmodul(99), valasz%)
                      form1.List2.Clear
                    End If
                  Case "S"
                    Call mess(langmodul(100), 2, 0, langmodul(99), valasz%)
                    form1.List2.Clear
                  Case Else
                End Select
              End If
            Else
              form1.List2.Clear
            End If
          Else
            form1.List2.Clear
          End If
        End If
      End If
    Case "AUW-KFRG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KFBZ" Then
        If vs& = 3 Then
          mkod$ = Vektor.MSFlexGrid1.TextMatrix(vs&, 1)
          mozrec$ = dbxkey("KMOZ", mkod$)
          form1.Text5.Text = Mid$(mozrec$, 4, 30)
          form1.Text5.Visible = True
        End If
      End If
      If objektum$ = "KFTT" Then
        form1.List2.Clear
        tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 6)
        If Trim$(tkod$) <> "" Then
          If vo& = 3 Then
            termrec$ = dbxkey("PTRM", tkod$)
            km@ = xval(Mid$(termrec$, 315, 14))
            If Mid$(mozrec$, 34, 1) = "B" Then
              If xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 3)) = 0 Then
                Call mess(langmodul(122), 2, 0, langmodul(123), valasz%)
                'MsgBox langmodul(122), 48, langmodul(123)
                'MsgBox "Beszerzési ár kötelezõ!", 48, "Készlet hiba"
                mezohiba% = 1
              End If
            End If
          End If
          If vo& = 2 Then
            termrec$ = dbxkey("PTRM", tkod$)
            km@ = xval(Mid$(termrec$, 315, 14))
            If Mid$(mozrec$, 34, 1) = "K" Then
              If km@ < xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 2)) Then
                Call mess(langmodul(124), 2, 0, langmodul(123), valasz%)
                'MsgBox langmodul(124), 48, langmodul(123)
                'MsgBox "Kiadáshoz a készlet kevés!", 48, "Készlet hiba"
                mezohiba% = 1
              End If
            End If
          End If
          If vo& = 1 Or vo& = 2 Then
            termrec$ = dbxkey("PTRM", tkod$)
            form1.List2.AddItem Mid$(termrec$, 7, 60)
            form1.List2.AddItem Mid$(termrec$, 67, 60)
            Select Case Mid$(termrec$, 127, 1)
              Case "A"
                If Mid$(termrec$, 187, 1) = "I" Then
                  form1.List2.AddItem langmodul(96) + ": " + Trim$(ertszam(Mid$(termrec$, 315, 14), 14, 3)) + " " + Trim(Mid$(termrec$, 140, 6)) + " " + langmodul(87) + ":" + Trim(ertszam(Mid$(termrec$, 301, 14), 14, 2))
                  'form1.List2.AddItem "Aktuális készlet: " + Trim$(ertszam(Mid$(termrec$, 315, 14), 14, 3)) + " " + Trim(Mid$(termrec$, 140, 6)) + " Nyilv.ár:" + Trim(ertszam(Mid$(termrec$, 301, 14), 14, 2))
                Else
                  Call mess(langmodul(98), 2, 0, langmodul(99), valasz%)
                  'MsgBox langmodul(98), 48, langmodul(99)
                  'MsgBox "Készletkezelés letiltva", 48, "Termék hiba!"
                  form1.List2.Clear
                  mezohiba% = 1
                End If
              Case "F", "S"
                Call mess(langmodul(100), 2, 0, langmodul(99), valasz%)
                'MsgBox langmodul(100), 48, langmodul(99)
                'MsgBox "Szolgáltatás", 48, "Termék hiba!"
                form1.List2.Clear
                mezohiba% = 1
              Case Else
            End Select
          End If
        End If
      End If
    Case "AUW-RJOV"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KSZF" Then
        If vs& = 1 Then
          Skod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
          szcrec$ = dbxkey("KSZC", Skod$)
          If szcrec$ <> "" Then
            Vektor.MSFlexGrid1.TextMatrix(2, 1) = Mid$(szcrec$, 16, 15)
          End If
        End If
        If vs& = 2 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(2, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            Vektor.MSFlexGrid1.TextMatrix(7, 1) = Mid$(partrec$, 328, 2)
          End If
        End If
      End If
      If objektum$ = "KSVL" Then
        If vo& = 1 And mezohiba% <> -1 Then
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          minta$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(15), 8)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
              form1.List2.Clear
              form1.List2.AddItem Mid$(termrec$, 16, 60)
              form1.List2.AddItem Mid$(termrec$, 196, 60)
              minta$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
              Call rkeszlet(tkod$, minta$, Mid$(krakrec$, 1, 4), keszme@, foglme@)
              Select Case Mid$(termrec$, 442, 1)
                Case "R"
                  If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
                    form1.List2.AddItem langmodul(104) + " " + langmodul(102) + ". " + langmodul(96) + ": " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                    'form1.List2.AddItem "Áru készletkezeléssel. Aktuális készlet: " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                  Else
                    form1.List2.AddItem langmodul(104) + " " + langmodul(103) + "."
                    'form1.List2.AddItem "Áru készletkezelés nélkül."
                  End If
                Case "K", "F"
                  If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
                    form1.List2.AddItem langmodul(106) + " " + langmodul(102) + ". " + langmodul(96) + ": " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                    'form1.List2.AddItem "Saját termék készletkezeléssel. Aktuális készlet: " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                  Else
                    form1.List2.AddItem langmodul(106) + " " + langmodul(103) + "."
                    'form1.List2.AddItem "Saját termék készletkezelés nélkül."
                  End If
                Case "A"
                  If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
                    form1.List2.AddItem langmodul(105) + " " + langmodul(102) + ". " + langmodul(96) + ": " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                    'form1.List2.AddItem "Anyag készletkezeléssel. Aktuális készlet: " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                  Else
                    form1.List2.AddItem langmodul(105) + " " + langmodul(103) + "."
                    'form1.List2.AddItem "Anyag készletkezelés nélkül."
                  End If
                Case "S"
                  form1.List2.AddItem langmodul(100) + "."
                  'form1.List2.AddItem "Szolgáltatás."
                Case Else
              End Select
              Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = Mid$(termrec$, 484, 6)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = Trim(Mid$(termrec$, 552, 14))
              Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = Mid$(termrec$, 706, 2)
              ellszi$ = Mid$(termrec$, 716, 24)
              If Trim(ellszi$) = "" Then ellszi$ = Mid$(krakrec$, 151, 24)
              If Trim(ellszi$) <> "" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = Mid$(ellszi$, 1, 8)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = Mid$(ellszi$, 9, 8)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 10) = Mid$(ellszi$, 17, 8)
              End If
              If Trim(Mid$(partrec$, 783, 8)) <> "" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = Mid$(partrec$, 783, 8)
              End If
              elar@ = arazo(Mid$(partrec$, 1, 15), Mid$(termrec$, 1, 15), minta$, maidatum$, diszx$)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 6) = ertszam(Str$(elar@), 12, 2)
            End If
          End If
        End If
      End If
    Case "AUW-CSZL", "AUW-RSZL", "AUW-RSZS"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KSZF" Then
        If vs& = 1 Then
          Skod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
          szcrec$ = dbxkey("KSZC", Skod$)
          If szcrec$ <> "" Then
            If Trim(Mid(szcrec$, 200, 8)) <> "" Then
              If Trim(Vektor.MSFlexGrid1.TextMatrix(12, 1)) = "" Then
                Vektor.MSFlexGrid1.TextMatrix(12, 1) = Mid(szcrec$, 200, 8)
              End If
            End If
            Vektor.MSFlexGrid1.TextMatrix(2, 1) = Mid$(szcrec$, 16, 15)
          End If
        End If
        If vs& = 2 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(2, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            If Trim(Mid(partrec$, 319, 8)) <> "" Then
              If Trim(Vektor.MSFlexGrid1.TextMatrix(12, 1)) = "" Then
                Vektor.MSFlexGrid1.TextMatrix(12, 1) = Mid(partrec$, 319, 8)
              End If
            End If
            pjelleg$ = Mid$(partrec$, 700, 2)
            If pjelleg$ = "KE" Then Vektor.MSFlexGrid1.TextMatrix(3, 1) = "ES"
            If pjelleg$ = "KG" Then Vektor.MSFlexGrid1.TextMatrix(3, 1) = "XS"
            Vektor.MSFlexGrid1.TextMatrix(7, 1) = Mid$(partrec$, 328, 2)
          End If
        End If
        If vs& = 4 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(2, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            napok% = Val(Mid$(partrec$, 330, 3))
            fidat$ = Vektor.MSFlexGrid1.TextMatrix(4, 1)
            If napok% > 0 Then
              For i13% = 1 To napok%
                xxxx$ = novdat(fidat$)
                fidat$ = xxxx$
              Next
            End If
            Vektor.MSFlexGrid1.TextMatrix(6, 1) = fidat$
          End If
        End If
        If vs& = 5 Then erteknap$ = Vektor.MSFlexGrid1.TextMatrix(5, 1)
        If vs& = 8 Then
          bankszamla$ = pzbank$
          dnm$ = Vektor.MSFlexGrid1.TextMatrix(8, 1)
          dkod$ = bankszamla$ + dnm$
          devrec$ = dbxkey("PDEV", dkod$)
          If devrec$ = "" Then
            '--- nincs pdev rekord
          Else
            dkod$ = bankszamla$ + dnm$ + erteknap$
            arfrec$ = dbxkey("PDRF", dkod$)
            If arfrec$ = "" Then
              '--- nincs árfolyam
            Else
              arfkod$ = Mid$(irec$, 478, 1)
              Select Case arfkod$
                Case "V"
                  arf@ = xval(Mid$(arfrec$, 24, 10))
                Case "K"
                  arf@ = xval(Mid$(arfrec$, 34, 10))
                Case "E"
                  arf@ = xval(Mid$(arfrec$, 44, 10))
                Case Else
              End Select
              egyseg@ = xval(Mid$(arfrec$, 18, 6))
              If egyseg@ = 0 Then egyseg@ = 1
              arf1@ = arf@ / egyseg
              Vektor.MSFlexGrid1.TextMatrix(9, 1) = ertszam(Str$(arf1@), 12, 4)
            End If
          End If
        End If
      End If
      If objektum$ = "KSZL" Then
        If vo& = 5 Or vo& = 6 Then
          peng@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 6))
          liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 5))
          penft@ = (liar@ * peng@) / 100
          peft$ = ertszam(Str$(penft@), 12, 2)
          penft@ = xval(peft$)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = peft$
          Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = ertszam(Str$(liar@ + penft@), 12, 2)
        End If
        If vo& = 7 Then
          liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 5))
          penft@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 7))
          Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = ertszam(Str$(liar@ + penft@), 12, 2)
          If liar@ <> 0 Then peng@ = penft@ / (liar@ / 100) Else peng@ = 0
          Tabla.MSFlexGrid1.TextMatrix(vs&, 6) = ertszam(Str$(peng@), 6, 2)
        End If
        If vo& = 8 Then
          liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 5))
          elar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 8))
          penft@ = elar@ - liar@
          If liar@ <> 0 Then peng@ = penft@ / (liar@ / 100) Else peng@ = 0
          Tabla.MSFlexGrid1.TextMatrix(vs&, 6) = ertszam(Str$(peng@), 6, 2)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = ertszam(Str$(penft@), 12, 2)
        End If
        If vo& = 1 And mezohiba% <> -1 Then
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          minta$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(15), 8)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
              form1.List2.Clear
              form1.List2.AddItem Mid$(termrec$, 16, 60)
              form1.List2.AddItem Mid$(termrec$, 196, 60)
              minta$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
              If programnev$ = "AUW-CSZL" Then
                keszme@ = xval(Mid$(termrec$, 748, 14))
              Else
                Call rkeszlet(tkod$, minta$, Mid$(krakrec$, 1, 4), keszme@, foglme@)
              End If
              Select Case Mid$(termrec$, 442, 1)
                Case "R"
                  If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
                    form1.List2.AddItem langmodul(104) + " " + langmodul(102) + ". " + langmodul(96) + ": " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                    'form1.List2.AddItem "Áru készletkezeléssel. Aktuális készlet: " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                  Else
                    form1.List2.AddItem langmodul(104) + " " + langmodul(103) + "."
                    'form1.List2.AddItem "Áru készletkezelés nélkül."
                  End If
                Case "K", "F"
                  If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
                    form1.List2.AddItem langmodul(106) + " " + langmodul(102) + ". " + langmodul(96) + ": " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                    'form1.List2.AddItem "Saját termék készletkezeléssel. Aktuális készlet: " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                  Else
                    form1.List2.AddItem langmodul(106) + " " + langmodul(103) + "."
                    'form1.List2.AddItem "Saját termék készletkezelés nélkül."
                  End If
                Case "A"
                  If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
                    form1.List2.AddItem langmodul(105) + " " + langmodul(102) + ". " + langmodul(96) + ": " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                    'form1.List2.AddItem "Anyag készletkezeléssel. Aktuális készlet: " + Trim$(ertszam(Str(keszme@), 14, 3)) + " " + Trim(Mid$(termrec$, 484, 6))
                  Else
                    form1.List2.AddItem langmodul(105) + " " + langmodul(103) + "."
                    'form1.List2.AddItem "Anyag készletkezelés nélkül."
                  End If
                Case "S"
                  form1.List2.AddItem langmodul(100) + "."
                  'form1.List2.AddItem "Szolgáltatás."
                Case Else
              End Select
              Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = Mid$(termrec$, 484, 6)
              If nemarazni% = 0 Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = Trim(Mid$(termrec$, 580, 14))
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = Trim(Mid$(termrec$, 580, 14))
                Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = Mid$(termrec$, 706, 2)
              End If
              ' Eszi - mezoveg 3-ba attenni
              ' Ha belföldi a számla akkor 716, ha export akkor 1015
              ellszi$ = Mid$(termrec$, 716, 24)
              If Trim(ellszi$) = "" Then ellszi$ = Mid$(krakrec$, 151, 24)
              If programnev$ = "AUW-CSZL" Then
                Mid$(ellszi$, 9, 8) = Mid$(krakrec$, 159, 8)
              End If
              If Trim(ellszi$) <> "" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 10) = Mid$(ellszi$, 1, 8)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 11) = Mid$(ellszi$, 9, 8)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 12) = Mid$(ellszi$, 17, 8)
              End If
              If Trim(Mid$(partrec$, 783, 8)) <> "" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 10) = Mid$(partrec$, 783, 8)
              End If
              If nemarazni% = 0 Then
                elar@ = arazo(Mid$(partrec$, 1, 15), Mid$(termrec$, 1, 15), minta$, maidatum$, diszx$)
                'Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = ertszam(Str(elar@), 12, 2)
                liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 5))
                penft@ = elar@ - liar@
                peft$ = ertszam(Str$(penft@), 12, 2)
                penft@ = xval(peft$)
                If liar@ <> 0 Then peng@ = penft@ / (liar@ / 100) Else peng@ = 0
                Tabla.MSFlexGrid1.TextMatrix(vs&, 6) = ertszam(Str(peng@), 6, 2)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = peft$
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = ertszam(Str$(elar@), 12, 4)
              End If
            End If
          End If
        End If
        If vo& = 1 Or vo& = 3 Or vo& = 5 Or vo& = 6 Or vo& = 7 Or vo& = 8 Then
          '--- számlaéerék kiszámítása
          liert@ = 0: enge@ = 0: elert@ = 0: afa@ = 0
          gigsze% = 999
          If programnev$ = "AUW-RSZS" Then gigsze% = 199
          For i13% = 1 To gigsze%
            tkod$ = Tabla.MSFlexGrid1.TextMatrix(i13%, 1)
            If Trim$(tkod$) <> "" Then
              afakod$ = Tabla.MSFlexGrid1.TextMatrix(i13%, 9)
              If afakod$ = utafakod$ Then
                afakulcs@ = utafakulcs@
              Else
                afrec$ = dbxkey("PAFA", afakod$)
                afakulcs@ = xval(Mid$(afrec$, 33, 6))
                utafakod$ = afakod$
                utafakulcs@ = afakulcs@
              End If
              menny@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 3))
              liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 5))
              penft@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 7))
              elar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 8))
              If menny@ <> 0 Then
                alaposz@ = elar * menny@
                elert@ = elert@ + elar@ * menny@
                enge@ = enge@ + penft@ * menny@
              Else
                alaposz@ = elar@
                elert@ = elert@ + elar@
                enge@ = enge@ + penft@
              End If
              afaosz@ = (alaposz@ * afakulcs@) / 100
              afaker% = xval(Mid$(irec$, 345, 1))
              If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
              afaosz@ = xval(Format(afaosz@, fst$))
              afa@ = afa@ + afaosz@
            End If
          Next
          form1.List3.Clear
          form1.List3.AddItem langmodul(125) + ":" + ertszam(Str$(elert@ - enge@), 14, 2)
          'form1.List3.AddItem "Lista érték :" + ertszam(Str$(elert@ - enge@), 14, 2)
          form1.List3.AddItem langmodul(126) + ":" + ertszam(Str$(enge@), 14, 2)
          'form1.List3.AddItem "Engedmény   :" + ertszam(Str$(enge@), 14, 2)
          form1.List3.AddItem langmodul(127) + ":" + ertszam(Str$(elert@), 14, 2)
          'form1.List3.AddItem "Nettó érték :" + ertszam(Str$(elert@), 14, 2)
          form1.List3.AddItem langmodul(128) + ":" + ertszam(Str$(elert@ + afa@), 14, 2)
          'form1.List3.AddItem "Bruttó érték:" + ertszam(Str$(elert@ + afa@), 14, 2)
        End If
      End If
    Case "AUW-PSZL"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "PSZL" Or objektum$ = "PSZ2" Then
         iii% = 0
         If objektum$ = "PSZ2" Then
             iii% = 1
         End If

        If vo& = (4 + iii%) Or vo& = (5 + iii%) Then
          peng@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 5 + iii%))
          liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 4 + iii%))
          penft@ = (liar@ * peng@) / 100
          peft$ = ertszam(Str$(penft@), 12, 2)
          penft@ = xval(peft$)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 6 + iii%) = peft$
          Tabla.MSFlexGrid1.TextMatrix(vs&, 7 + iii%) = ertszam(Str$(liar@ + penft@), 12, 2)
        End If
        If vo& = (6 + iii%) Then
          liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 4 + iii%))
          penft@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 6 + iii%))
          Tabla.MSFlexGrid1.TextMatrix(vs&, 7 + iii%) = ertszam(Str$(liar@ + penft@), 12, 2)
          If liar@ <> 0 Then peng@ = penft@ / (liar@ / 100) Else peng@ = 0
          Tabla.MSFlexGrid1.TextMatrix(vs&, 5 + iii%) = ertszam(Str$(peng@), 6, 2)
        End If
        If vo& = (7 + iii%) Then
          liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 4 + iii%))
          elar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 7 + iii%))
          penft@ = elar@ - liar@
          If liar@ <> 0 Then peng@ = penft@ / (liar@ / 100) Else peng@ = 0
          Tabla.MSFlexGrid1.TextMatrix(vs&, 5 + iii%) = ertszam(Str$(peng@), 6, 2)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 6 + iii%) = ertszam(Str$(penft@), 12, 2)
        End If
        If vo& = 1 And mezohiba% <> -1 Then
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 6)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("PTRM", tkod$)
            If Trim(vtsztomb(vs&)) = "" Then vtsztomb(vs&) = Trim(Mid$(termrec$, 128, 12))
            If Trim(Mid$(termrec$, 7, 60)) = "" Then
              Do
                cikknev.Label3.Caption = Str(vs&) + "."
                cikknev.Text1.Text = Trim(meszovtomb(vs&, 1))
                cikknev.Text1.SelStart = Len(Trim(meszovtomb(vs&, 1)))
                cikknev.Text3.Text = Trim(meszovtomb(vs&, 2))
                cikknev.Text3.SelStart = Len(Trim(meszovtomb(vs&, 2)))
                cikknev.Text4.Text = Trim(meszovtomb(vs&, 3))
                cikknev.Text4.SelStart = Len(Trim(meszovtomb(vs&, 3)))
                cikknev.Text5.Text = Trim(meszovtomb(vs&, 4))
                cikknev.Text5.SelStart = Len(Trim(meszovtomb(vs&, 4)))
                cikknev.Text6.Text = Trim(meszovtomb(vs&, 5))
                cikknev.Text6.SelStart = Len(Trim(meszovtomb(vs&, 5)))
                cikknev.Text2.Text = Trim(vtsztomb(vs&))
                cikknev.Text2.SelStart = Len(Trim(vtsztomb(vs&)))
                cikknev.Show vbModal
                meszovtomb(vs&, 1) = Trim(cikknev.Text1.Text)
                meszovtomb(vs&, 2) = Trim(cikknev.Text3.Text)
                meszovtomb(vs&, 3) = Trim(cikknev.Text4.Text)
                meszovtomb(vs&, 4) = Trim(cikknev.Text5.Text)
                meszovtomb(vs&, 5) = Trim(cikknev.Text6.Text)
                vtsztomb(vs&) = Trim(cikknev.Text2.Text)
              Loop While Trim(meszovtomb$(vs&, 1)) = ""
            End If
            If termrec$ <> "" Then
              form1.List2.Clear
              If Trim(Mid$(termrec$, 7, 60)) = "" Then
                form1.List2.AddItem meszovtomb$(vs&, 1)
              Else
                form1.List2.AddItem Mid$(termrec$, 7, 60)
                form1.List2.AddItem Mid$(termrec$, 67, 60)
              End If
              If (Tabla.MSFlexGrid1.TextMatrix(vs&, 3 + iii%) = "") Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 3 + iii%) = Mid$(termrec$, 140, 6)
              End If
              If (Tabla.MSFlexGrid1.TextMatrix(vs&, 4 + iii%) = "") Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 4 + iii%) = Mid$(termrec$, 146, 12)
              End If
              If (Tabla.MSFlexGrid1.TextMatrix(vs&, 7 + iii%) = "") Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 7 + iii%) = Mid$(termrec$, 146, 12)
              End If
              Tabla.MSFlexGrid1.TextMatrix(vs&, 8 + iii%) = Mid$(termrec$, 161, 2)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 9 + iii%) = Mid$(termrec$, 163, 8)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 10 + iii%) = Mid$(termrec$, 171, 8)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 11 + iii%) = Mid$(termrec$, 179, 8)
              If Trim(Mid$(partrec$, 783, 8)) <> "" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 9 + iii%) = Mid$(partrec$, 783, 8)
              End If
              peng@ = xval(Mid$(partrec$, 333, 6))
              If peng@ <> 0 Then
                liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 4 + iii%))
                penft@ = (liar@ * peng@) / 100
                peft$ = ertszam(Str$(penft@), 12, 2)
                penft@ = xval(peft$)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 5 + iii%) = Mid$(partrec$, 333, 6)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 6 + iii%) = peft$
                Tabla.MSFlexGrid1.TextMatrix(vs&, 7 + iii%) = ertszam(Str$(liar@ + penft@), 12, 2)
              End If
            End If
          End If
        End If
        If vo& = 1 Or vo& = (2 + iii%) Or vo& = (4 + iii%) Or vo& = (5 + iii%) Or vo& = (6 + iii%) Or vo& = (7 + iii%) Or vo& = (8 + iii%) Then
          '--- számlaéerék kiszámítása
          liert@ = 0: enge@ = 0: elert@ = 0: afa@ = 0
          For i13% = 1 To 99
            tkod$ = Tabla.MSFlexGrid1.TextMatrix(i13%, 1)
            If Trim$(tkod$) <> "" Then
              afakod$ = Tabla.MSFlexGrid1.TextMatrix(i13%, 8 + iii%)
              If afakod$ = utafakod$ Then
                afakulcs@ = utafakulcs@
              Else
                afrec$ = dbxkey("PAFA", afakod$)
                afakulcs@ = xval(Mid$(afrec$, 33, 6))
                utafakod$ = afakod$
                utafakulcs@ = afakulcs@
              End If
' Eszi - kiírt számla érték számolás
              menny@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 2 + iii%))
              liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 4 + iii%))
              penft@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 6 + iii%))
              elar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 7 + iii%))
              If menny@ <> 0 Then
                alaposz@ = elar * menny@
                elert@ = elert@ + elar@ * menny@
                enge@ = enge@ + penft@ * menny@
              Else
                alaposz@ = elar@
                elert@ = elert@ + elar@
                enge@ = enge@ + penft@
              End If
              
              ertker% = xval(Mid$(irec$, 344, 1))
              If ertker% = 0 Then fste$ = "############0" Else fste$ = "#############0." + String(ertker%, "0")
              alaposz@ = xval(Format(alaposz@, fste$))
              elert@ = xval(Format(elert@, fste$))
              enge@ = xval(Format(enge@, fste$))
              
              afaosz@ = (alaposz@ * afakulcs@) / 100
              ' az ÁFÁt tételenként kerekíti
              afaker% = xval(Mid$(irec$, 345, 1))
              If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
              afaosz@ = xval(Format(afaosz@, fst$))
              afa@ = afa@ + afaosz@
            End If
          Next
          
          form1.List3.Clear
          form1.List3.AddItem langmodul(125) + ":" + ertszam(Str$(elert@ - enge@), 14, 2)
          'form1.List3.AddItem "Lista érték :" + ertszam(Str$(elert@ - enge@), 14, 2)
          form1.List3.AddItem langmodul(126) + ":" + ertszam(Str$(enge@), 14, 2)
          'form1.List3.AddItem "Engedmény   :" + ertszam(Str$(enge@), 14, 2)
          form1.List3.AddItem langmodul(127) + ":" + ertszam(Str$(elert@), 14, 2)
          'form1.List3.AddItem "Nettó érték :" + ertszam(Str$(elert@), 14, 2)
          form1.List3.AddItem langmodul(128) + ":" + ertszam(Str$(elert@ + afa@), 14, 2)
          'form1.List3.AddItem "Bruttó érték:" + ertszam(Str$(elert@ + afa@), 14, 2)
        End If
      End If
      If objektum$ = "PSZF" Then
        If vs& = 1 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            pjelleg$ = Mid$(partrec$, 700, 2)
            If pjelleg$ = "BK" Then Vektor.MSFlexGrid1.TextMatrix(2, 1) = "BS"
            If pjelleg$ = "KE" Then Vektor.MSFlexGrid1.TextMatrix(2, 1) = "ES"
            If pjelleg$ = "KG" Then Vektor.MSFlexGrid1.TextMatrix(2, 1) = "XS"
            Vektor.MSFlexGrid1.TextMatrix(6, 1) = Mid$(partrec$, 328, 2)
          End If
        End If
        If vs& = 3 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            napok% = xval(Mid$(partrec$, 330, 3))
            fidat$ = Vektor.MSFlexGrid1.TextMatrix(3, 1)
            If napok% > 0 Then
              For i13% = 1 To napok%
                xxxx$ = novdat(fidat$)
                fidat$ = xxxx$
              Next
            End If
            Vektor.MSFlexGrid1.TextMatrix(5, 1) = fidat$
          End If
        End If
        If vs& = 4 Then erteknap$ = Vektor.MSFlexGrid1.TextMatrix(4, 1)
        If vs& = 7 Then
          bankszamla$ = pzbank$
          dnm$ = Vektor.MSFlexGrid1.TextMatrix(7, 1)
          dkod$ = bankszamla$ + dnm$
          devrec$ = dbxkey("PDEV", dkod$)
          If devrec$ = "" Then
            '--- nincs pdev rekord
          Else
            dkod$ = bankszamla$ + dnm$ + erteknap$
            arfrec$ = dbxkey("PDRF", dkod$)
            If arfrec$ = "" Then
              '--- nincs árfolyam
            Else
              arfkod$ = Mid$(irec$, 478, 1)
              Select Case arfkod$
                Case "V"
                  arf@ = xval(Mid$(arfrec$, 24, 10))
                Case "K"
                  arf@ = xval(Mid$(arfrec$, 34, 10))
                Case "E"
                  arf@ = xval(Mid$(arfrec$, 44, 10))
                Case Else
              End Select
              egyseg@ = xval(Mid$(arfrec$, 18, 6))
              If egyseg@ = 0 Then egyseg@ = 1
              arf1@ = arf@ / egyseg
              Vektor.MSFlexGrid1.TextMatrix(8, 1) = ertszam(Str$(arf1@), 12, 4)
            End If
          End If
        End If
      End If
    Case "AUW-FVRG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "FTET" Then
        tetelui% = 0: osszegui@ = 0
        For i987% = 1 To Tabla.MSFlexGrid1.Rows - 1
          oao@ = xval(Tabla.MSFlexGrid1.TextMatrix(i987%, 6))
          If oao@ <> 0 Then tetelui% = tetelui% + 1: osszegui@ = osszegui@ + oao@
        Next
        form1.Text6.Text = tetelui%
        form1.Text7.Text = ertszam(Str(osszegui@), 14, 2)
      End If
      If objektum$ = "FTET" And vo& = 8 Then
        bankszamla$ = Mid$(irec$, 470, 8)
        Call forintosit(8, 7, 6, bankszamla$, erteknap$, "K", "T", vs&)
      End If
    Case "AUW-PBNK"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "BBIZ" Then
        If vs& = 3 Then
          tipkod$ = Trim(Vektor.MSFlexGrid1.TextMatrix(3, 1))
          Select Case tipkod$
            Case "V": Vektor.MSFlexGrid1.TextMatrix(10, 1) = "Vevõ számla kiegyenlítés"
            Case "S": Vektor.MSFlexGrid1.TextMatrix(10, 1) = "Szállító számla kiegyenlítés"
            Case "E": Vektor.MSFlexGrid1.TextMatrix(10, 1) = "Vevõ elõleg"
            Case "L": Vektor.MSFlexGrid1.TextMatrix(10, 1) = "Szállító elõleg"
            Case Else
          End Select
        End If
        If vs& = 1 Then erteknap$ = Vektor.MSFlexGrid1.TextMatrix(1, 1)
        If vs& = 9 Then
          If Vektor.MSFlexGrid1.TextMatrix(2, 1) = "J" Then
            '--- jóváírás
            If devizabank$ = "" Then
              '--- normál bankszámla
              Call forintosit(9, 8, 7, pzbank$, erteknap$, "K", "V", vs&)
            Else
              '--- deviza készletre
              arfkod$ = Mid$(irec$, 480, 1)
              Call forintosit(9, 8, 7, pzbank$, erteknap$, arfkod$, "V", vs&)
            End If
          Else
            '--- terhelés
            If devizabank$ = "" Then
              '--- normál bankszámla
              Call forintosit(9, 8, 7, pzbank$, erteknap$, "E", "V", vs&)
            Else
              '--- deviza készletrõl
              Call forintosit(9, 8, 7, pzbank$, erteknap$, "S", "V", vs&)
            End If
          End If
        End If
        If vs& = 5 Then
          If Trim(Vektor.MSFlexGrid1.TextMatrix(5, 1)) <> "" Then
            ekod$ = Right$("0000000" + Vektor.MSFlexGrid1.TextMatrix(5, 1), 7)
            erec$ = dbxkey("PELO", ekod$)
            Vektor.MSFlexGrid1.TextMatrix(6, 1) = Mid$(erec$, 23, 15)
          End If
        End If
      End If
      If objektum$ = "BBIZ" And vs& = 6 Then
        pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(6, 1) + Space$(15), 15)
        partrec$ = dbxkey("PART", pkod$)
      End If
      If objektum$ = "BTE3" And vo& = 2 Then
        Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Mid$(partrec$, 298, 8)
        If xdevbank@ <> 0 Then
          xxddvv@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 2))
          forit.Text1.Text = Str(xxddvv@)
          forit.Text2.Text = Str(zarfo#)
          If forintja@(vs&) = 0 Then
            forit.Text3.Text = ertszam(Str(xxddvv@ * zarfo#), 14, 2)
          Else
            forit.Text3.Text = ertszam(Str(forintja@(vs&)), 14, 2)
          End If
          forit.Show vbModal
          forintja@(vs&) = xval(forit.Text3.Text)
        End If
      End If
      If objektum$ = "BTE1" Then
        If abmod% = 1 Then
          '--- ablakban
          If vs& = 3 Then
            '--- összeg, árfolyamkülönbözetet számolni
            If xdevszamla@ <> 0 And xdevbank@ <> 0 Then
              '--- számla árfolyam
              osb@ = xval(Vektor.MSFlexGrid1.TextMatrix(3, 1))
              Call arfokul(osb@ * kxarf#, xforintszamla@, xdevszamla@, xforintbank@, xdevbank@ * kxarf#, szampmod$, arfkul@, ellenszamla$)
              If Trim(ellenszamla$) <> "" Then
                Vektor.MSFlexGrid1.TextMatrix(8, 1) = ertszam(Str$(arfkul@), 14, 2)
                Vektor.MSFlexGrid1.TextMatrix(9, 1) = ellenszamla$
              End If
            End If
          End If
          If vs& = 4 Then
            '--- diszkon eng.
            osb@ = xval(Vektor.MSFlexGrid1.TextMatrix(3, 1))
            disk@ = xval(Vektor.MSFlexGrid1.TextMatrix(4, 1))
            egyb@ = xval(Vektor.MSFlexGrid1.TextMatrix(6, 1))
            If osb@ > xaktkie@ Then
              osb@ = xaktkie@
            End If
            kieo@ = osb@ + disk@ + egyb@
            If kieo@ > xegyenleg@ Then
              osb@ = xegyenleg@ - disk@ - egyb@
            End If
            'osb@ = xaktkie@ - disk@ - egyb@
            If xdevszamla@ <> 0 And xdevbank@ <> 0 Then
              Call arfokul(osb@, xforintszamla@, xdevszamla@, xforintbank@, xdevbank@, szampmod$, arfkul@, ellenszamla$)
              If Trim(ellenszamla$) <> "" Then
                Vektor.MSFlexGrid1.TextMatrix(8, 1) = ertszam(Str$(arfkul@), 14, 2)
                Vektor.MSFlexGrid1.TextMatrix(9, 1) = ellenszamla$
              End If
            End If
            Vektor.MSFlexGrid1.TextMatrix(3, 1) = ertszam(Str$(osb@), 14, 2)
            If disk@ <> 0 Then
              If szampmod$ = "V" Then
                Vektor.MSFlexGrid1.TextMatrix(5, 1) = Mid$(irec$, 586, 8)
              Else
                Vektor.MSFlexGrid1.TextMatrix(5, 1) = Mid$(irec$, 594, 8)
              End If
            End If
          End If
          If vs& = 6 Then
            '--- egyéb korrekció
            osb@ = xval(Vektor.MSFlexGrid1.TextMatrix(3, 1))
            disk@ = xval(Vektor.MSFlexGrid1.TextMatrix(4, 1))
            egyb@ = xval(Vektor.MSFlexGrid1.TextMatrix(6, 1))
            If osb@ > xaktkie@ Then
              osb@ = xaktkie@
            End If
            kieo@ = osb@ + disk@ + egyb@
            If kieo@ > xegyenleg@ Then
              osb@ = xegyenleg@ - disk@ - egyb@
            End If
            'disk@ = Val(Vektor.MSFlexGrid1.TextMatrix(4, 1))
            'egyb@ = Val(Vektor.MSFlexGrid1.TextMatrix(6, 1))
            'osb@ = xaktkie@ - disk@ - egyb@
            If xdevszamla@ <> 0 And xdevbank@ <> 0 Then
              Call arfokul(osb@, xforintszamla@, xdevszamla@, xforintbank@, xdevbank@, szampmod$, arfkul@, ellenszamla$)
              If Trim(ellenszamla$) <> "" Then
                Vektor.MSFlexGrid1.TextMatrix(8, 1) = ertszam(Str$(arfkul@), 14, 2)
                Vektor.MSFlexGrid1.TextMatrix(9, 1) = ellenszamla$
              End If
            End If
            Vektor.MSFlexGrid1.TextMatrix(3, 1) = ertszam(Str$(osb@), 14, 2)
            If egyb@ <> 0 Then
              If szampmod$ = "V" Then
                Vektor.MSFlexGrid1.TextMatrix(7, 1) = Mid$(irec$, 570, 8)
              Else
                Vektor.MSFlexGrid1.TextMatrix(7, 1) = Mid$(irec$, 578, 8)
              End If
            End If
          End If
        Else
          '--- táblázatban
          If vo& = 1 Then
            xikta$ = Tabla.MSFlexGrid1.TextMatrix(vs&, 1)
            If szampmod$ = "V" Then
              szrec$ = dbxkey("PVSZ", xikta$)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 2) = Mid$(szrec$, 8, 15)
              Call szamlaegyenleg(szrec$, xossz@, xhelyb@, xkie@, xegyenleg@, "V", "391231", "391231", xforintegyenleg@)
            Else
              szrec$ = dbxkey("PSSZ", xikta$)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 2) = Mid$(szrec$, 8, 15)
              Call szamlaegyenleg(szrec$, xossz@, xhelyb@, xkie@, xegyenleg@, "S", "391231", "391231", xforintegyenleg@)
            End If
            xdevszamla@ = xval(Mid$(szrec$, 95, 14))
            xforintszamla@ = xval(Mid$(szrec$, 78, 14))
          End If
          If vo& = 3 Then
            '--- összeg, árfolyamkülönbözetet számolni
            If xdevszamla@ <> 0 And xdevbank@ <> 0 Then
              '--- számla árfolyam
              xxcx# = 1
              If xdevbank@ <> 0 Then
                xxddvv@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 3))
                forit.Label1.Caption = xdevnembank$
                forit.Text1.Text = Str(xxddvv@)
                forit.Text2.Text = Str(zarfo#)
                If forintja@(vs&) = 0 Then
                  forit.Text3.Text = ertszam(Str(xxddvv@ * zarfo#), 14, 2)
                Else
                  forit.Text3.Text = ertszam(Str(forintja@(vs&)), 14, 2)
                End If
                forit.Show vbModal
                forintja@(vs&) = xval(forit.Text3.Text)
              End If
              osb@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 3))
  xxcx# = keresztarfolyam#(vs&):
              If xxcx# = 0 Then xxcx# = 1
              If xxcx# = 1 Then
                xikta$ = Tabla.MSFlexGrid1.TextMatrix(vs&, 1)
                If szampmod$ = "V" Then
                  szrec$ = dbxkey("PVSZ", xikta$)
                Else
                  szrec$ = dbxkey("PSSZ", xikta$)
                End If
                sdev$ = Mid$(szrec$, 92, 3)  '--- számlázási devizanem
                If sdev$ <> xdevnembank Then
                  kxarf# = xarfolyam(sdev$, xdevnembank$)
                  xxcx# = kxarf#
                  keresztarfolyam#(vs&) = kxarf#
                End If
              End If
              Call arfokul(osb@ * xxcx#, xforintszamla@, xdevszamla@, xforintbank@, xdevbank@ * xxcx#, szampmod$, arfkul@, ellenszamla$)
              If Trim(ellenszamla$) <> "" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = ertszam(Str$(arfkul@), 14, 2)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = ellenszamla$
              End If
            End If
          End If
        End If
      End If
    Case "AUW-PTRG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "PBIZ" Then
        If vs& = 3 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(3, 1) + Space$(15), 15)
          If pkod$ <> Space(15) Then
            partrec$ = dbxkey("PART", pkod$)
            If partrec$ <> "" Then
              Vektor.MSFlexGrid1.TextMatrix(4, 1) = Trim(Mid$(partrec$, 16, 60))
            End If
          End If
        End If
      End If
      If objektum$ = "PTET" Then
        If vo& = 5 Then
          If penztarirany$ = "K" Then
            Call forintosit(5, 4, 3, pzbank$, erteknap$, "S", "T", vs&)
          Else
            Call forintosit(5, 4, 3, pzbank$, erteknap$, "V", "T", vs&)
          End If
        End If
      End If
      If voszl% = 3 And objektum$ = "PTET" Then
        marosz@ = 0
        For im1% = 1 To 50
          marosz@ = marosz@ + xval(Tabla.MSFlexGrid1.TextMatrix(im1%, 3))
        Next
        form1.Text5.Text = langmodul(130) + ":  " + Format(marosz@, "##########0.00")
        'form1.Text5.Text = "Bizonylat összege:  " + Format(marosz@, "##########0.00")
      End If
      If vo& = 8 And objektum$ = "PTET" Then
        xxxx$ = Tabla.MSFlexGrid1.TextMatrix(vs&, vo&)
        If Trim(xxxx$) <> "" Then
          elxrc$ = dbxkey("PELO", xxxx$)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = Mid$(elxrc$, 23, 15)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 14) = Mid$(elxrc$, 455, 8)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 15) = Mid$(elxrc$, 463, 8)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 16) = Mid$(elxrc$, 471, 8)
        End If
      End If
    Case "AUW-PKRG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "PKOR" Then
        If vs& = 1 Then erteknap$ = Vektor.MSFlexGrid1.TextMatrix(5, 1)
        If vs& = 9 Then
          If Vektor.MSFlexGrid1.TextMatrix(2, 1) = "S" Then
            arfkod$ = Mid$(irec$, 479, 1)
          Else
            arfkod$ = Mid$(irec$, 478, 1)
          End If
          Call forintosit(9, 8, 7, pzbank$, erteknap$, arfkod$, "V", vs&)
        End If
      End If
    Case "AUW-PVRG"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "PSHL" Then
        If vs& = 4 Then erteknap$ = Vektor.MSFlexGrid1.TextMatrix(4, 1)
        If vs& = 7 Then
          If Left$(Vektor.MSFlexGrid1.TextMatrix(1, 1), 1) = "V" Then
            arfkod$ = Mid$(irec$, 478, 1)
          Else
            arfkod$ = Mid$(irec$, 479, 1)
          End If
          Call forintosit(7, 6, 5, pzbank$, erteknap$, arfkod$, "V", vs&)
        End If
      End If
      If objektum$ = "PELO" Then
        If vs& = 5 Then erteknap$ = Vektor.MSFlexGrid1.TextMatrix(5, 1)
        If vs& = 8 Then
          If Vektor.MSFlexGrid1.TextMatrix(1, 1) = "V" Then
            arfkod$ = Mid$(irec$, 478, 1)
          Else
            arfkod$ = Mid$(irec$, 479, 1)
          End If
          Call forintosit(8, 7, 6, pzbank$, erteknap$, arfkod$, "V", vs&)
        End If
      End If
      If objektum$ = "PVSZ" Or objektum$ = "PSSZ" Or objektum$ = "XVSZ" Then
        If xvskod = "S" And (objektum$ = "PSSZ" Or objektum$ = "XVSZ") And vs& = 3 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(3, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            Skod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
            scim& = xval(Mid$(partrec$, 722, 10))
            If scim& > 0 Then
              fixa = FreeFile
              Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #fixa
              Do While scim& > 0
                psr$ = Space(1500)
                Get #fixa, scim& + 9, psr$
                If Mid$(psr$, 166, 1) <> "S" Then
                  If Mid$(psr$, 8, 15) = Skod$ Then
                    Call mess(langmodul(131), 5, 3, langmodul(132), valasz%)
                    'respons = MsgBox(langmodul(131), vbYesNo, langmodul(132))
                    'respons = MsgBox("Ennek a partnernek már van ilyen számlája! Mégis rögzíti?", vbYesNo, "Számlaszám hiba")
                    If valasz% = 0 Then
                    'If respons = vbNo Then
                      mezohiba% = 1
                      Close fixa
                      Exit Sub
                    End If
                  End If
                End If
                scim& = xval(Mid(psr$, 191, 10))
              Loop
              Close fixa
            End If
          End If
        End If
        If objektum$ = "XVSZ" And vs& = 14 Then
          csko$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space(4), 4)
          If Trim(csko$) <> "" Then
            cskrec$ = dbxkey("PCOP", csko$)
            If cskrec$ <> "" Then
              Vektor.MSFlexGrid1.TextMatrix(15, 1) = Mid$(cskrec$, 45, 2)
              Vektor.MSFlexGrid1.TextMatrix(18, 1) = Mid$(cskrec$, 47, 8)
              Vektor.MSFlexGrid1.TextMatrix(19, 1) = Mid$(cskrec$, 55, 8)
              Vektor.MSFlexGrid1.TextMatrix(20, 1) = Mid$(cskrec$, 63, 8)
              Vektor.MSFlexGrid1.TextMatrix(21, 1) = Mid$(cskrec$, 71, 8)
            End If
          End If
        End If
        If vs& = 4 Then
          Vektor.MSFlexGrid1.TextMatrix(5, 1) = Vektor.MSFlexGrid1.TextMatrix(4, 1)
          erteknap$ = Vektor.MSFlexGrid1.TextMatrix(4, 1)
        End If
        If objektum$ = "PSSZ" And vs& = 12 Then
          arfkod$ = Mid$(irec$, 479, 1)
           Call forintosit(12, 11, 10, pzbank$, erteknap$, arfkod$, "V", vs&)
        End If
        If (objektum$ = "PVSZ" Or objektum$ = "XVSZ") And vs& = 11 Then
          arfkod$ = Mid$(irec$, 478, 1)
          Call forintosit(11, 10, 9, pzbank$, erteknap$, arfkod$, "V", vs&)
        End If
        If objektum$ = "XVSZ" And vs& = 15 Then
          marosz@ = xval(Trim(Vektor.MSFlexGrid1.TextMatrix(9, 1)))
          afaker% = xval(Mid$(irec$, 345, 1))
          If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
          afik$ = Vektor.MSFlexGrid1.TextMatrix(vs&, 1)
          afrec$ = dbxkey("PAFA", afik$)
          afkod$ = Mid$(afrec$, 40, 2)
          afkulcs@ = xval(Mid$(afrec$, 33, 6))
          If afkod$ = "IU" Then
            alaposz@ = marosz@
            afaosz@ = alaposz@ * (afkulcs@ / 100)
            afaosz@ = xval(Format(afaosz@, fst$))
          Else
            alaposz@ = marosz@ / (1 + (afkulcs@ / 100))
            afaosz@ = xval(Format(marosz@ - alaposz@, fst$))
            alaposz@ = marosz@ - afaosz@
          End If
          Vektor.MSFlexGrid1.TextMatrix(vs& + 1, 1) = Format(alaposz@, fst$)
          Vektor.MSFlexGrid1.TextMatrix(vs& + 2, 1) = Format(afaosz@, fst$)
        End If
      End If
      If objektum$ = "PSAF" Then
        marosz@ = afosz@
        afaker% = xval(Mid$(irec$, 345, 1))
        If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
        For im1% = 1 To 5
          marosz@ = marosz@ - xval(Tabla.MSFlexGrid1.TextMatrix(im1%, 2)) - xval(Tabla.MSFlexGrid1.TextMatrix(im1%, 3))
        Next
        If voszl% = 1 Then
          '--- automatikus ÁFA bontás
          afik$ = Tabla.MSFlexGrid1.TextMatrix(vs&, vo&)
          afrec$ = dbxkey("PAFA", afik$)
          afkod$ = Mid$(afrec$, 40, 2)
          afkulcs@ = xval(Mid$(afrec$, 33, 6))
          If afkod$ = "IU" Then
            alaposz@ = marosz@
            afaosz@ = alaposz@ * (afkulcs@ / 100)
            afaosz@ = xval(Format(afaosz@, fst$))
          Else
            alaposz@ = marosz@ / (1 + (afkulcs@ / 100))
            afaosz@ = xval(Format(marosz@ - alaposz@, fst$))
            alaposz@ = marosz@ - afaosz@
          End If
          Tabla.MSFlexGrid1.TextMatrix(vs&, vo& + 1) = Format(alaposz@, fst$)
          Tabla.MSFlexGrid1.TextMatrix(vs&, vo& + 2) = Format(afaosz@, fst$)
        End If
        If voszl% = 2 Then
          afik$ = Tabla.MSFlexGrid1.TextMatrix(vs&, 1)
          afrec$ = dbxkey("PAFA", afik$)
          afkod$ = Mid$(afrec$, 40, 2)
          alaposz@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, vo&))
          afkulcs@ = xval(Mid$(afrec$, 33, 6))
          afaosz@ = (alaposz@ / 100) * afkulcs@
          Tabla.MSFlexGrid1.TextMatrix(vs&, vo& + 1) = Format(afaosz@, fst$)
        End If
      End If
    Case "AUW-SPDC"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "SPDC" Then
        If vs& = 3 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(3, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          form1.List1.AddItem "   Megbízó:"
          form1.List1.AddItem Mid$(partrec$, 16, 60)
        End If
        If vs& = 4 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(4, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          form1.List1.AddItem "   Számla fogadó:"
          form1.List1.AddItem Mid$(partrec$, 16, 60)
        End If
        If vs& = 5 Then
          form1.List1.AddItem "   Alvállakozó:"
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(5, 1) + Space$(15), 15)
          If Trim(pkod$) <> "" Then
            partrec$ = dbxkey("PART", pkod$)
            form1.List1.AddItem Mid$(partrec$, 16, 60)
          Else
            form1.List1.AddItem "Nincs"
          End If
        End If
        If vs& = 6 Then
          ptrm$ = Left$(Vektor.MSFlexGrid1.TextMatrix(6, 1) + Space$(6), 6)
          ptrmrec$ = dbxkey("PTRM", ptrm$)
          Vektor.MSFlexGrid1.TextMatrix(7, 1) = Trim(Mid$(ptrmrec, 7, 60))
          'Vektor.MSFlexGrid1.TextMatrix(7, 1) = Trim(Mid$(ptrmrec, 67, 60))
        End If
        If vs& = 16 Then
          pgk$ = Left$(Vektor.MSFlexGrid1.TextMatrix(15, 1) + Space$(10), 10)
          pgkrec$ = dbxkey("SPGJ", pgk$)
          Vektor.MSFlexGrid1.TextMatrix(17, 1) = Trim(Mid$(pgkrec, 112, 20))
          form1.List1.AddItem "   Gépjármû:"
          form1.List1.AddItem Mid$(pgkrec$, 11, 20)
          form1.List1.AddItem Mid$(pgkrec$, 31, 20)
          If Mid$(pgkrec$, 111, 1) = "I" Then form1.List1.AddItem "Idegen" Else form1.List1.AddItem "Saját"
        End If
      End If
      If objektum$ = "PSZF" Then
        If vs& = 1 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            pjelleg$ = Mid$(partrec$, 700, 2)
            If pjelleg$ = "BK" Then Vektor.MSFlexGrid1.TextMatrix(2, 1) = "BS"
            If pjelleg$ = "KE" Then Vektor.MSFlexGrid1.TextMatrix(2, 1) = "ES"
            If pjelleg$ = "KG" Then Vektor.MSFlexGrid1.TextMatrix(2, 1) = "XS"
            Vektor.MSFlexGrid1.TextMatrix(6, 1) = Mid$(partrec$, 328, 2)
          End If
        End If
        If vs& = 3 Then
          pkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
          partrec$ = dbxkey("PART", pkod$)
          If partrec$ <> "" Then
            napok% = xval(Mid$(partrec$, 330, 3))
            fidat$ = Vektor.MSFlexGrid1.TextMatrix(3, 1)
            If napok% > 0 Then
              For i13% = 1 To napok%
                xxxx$ = novdat(fidat$)
                fidat$ = xxxx$
              Next
            End If
            Vektor.MSFlexGrid1.TextMatrix(5, 1) = fidat$
          End If
        End If
        If vs& = 4 Then erteknap$ = Vektor.MSFlexGrid1.TextMatrix(4, 1)
      End If
      If objektum$ = "PSZL" Or objektum$ = "PSZ2" Then
        If vo& = 4 Or vo& = 5 Then
          peng@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 5))
          liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 4))
          penft@ = (liar@ * peng@) / 100
          peft$ = ertszam(Str$(penft@), 12, 2)
          penft@ = xval(peft$)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 6) = peft$
          Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = ertszam(Str$(liar@ + penft@), 12, 2)
        End If
        If vo& = 6 Then
          liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 4))
          penft@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 6))
          Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = ertszam(Str$(liar@ + penft@), 12, 2)
          If liar@ <> 0 Then peng@ = penft@ / (liar@ / 100) Else peng@ = 0
          Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = ertszam(Str$(peng@), 6, 2)
        End If
        If vo& = 7 Then
          liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 4))
          elar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 7))
          penft@ = elar@ - liar@
          If liar@ <> 0 Then peng@ = penft@ / (liar@ / 100) Else peng@ = 0
          Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = ertszam(Str$(peng@), 6, 2)
          Tabla.MSFlexGrid1.TextMatrix(vs&, 6) = ertszam(Str$(penft@), 12, 2)
        End If
        If vo& = 1 And mezohiba% <> -1 Then
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 6)
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("PTRM", tkod$)
            If Trim(vtsztomb(vs&)) = "" Then vtsztomb(vs&) = Trim(Mid$(termrec$, 128, 12))
            If Trim(Mid$(termrec$, 7, 60)) = "" Then
              Do
                cikknev.Label3.Caption = Str(vs&) + "."
                cikknev.Text1.Text = Trim(meszovtomb(vs&, 1))
                cikknev.Text1.SelStart = Len(Trim(meszovtomb(vs&, 1)))
                cikknev.Text3.Text = Trim(meszovtomb(vs&, 2))
                cikknev.Text3.SelStart = Len(Trim(meszovtomb(vs&, 2)))
                cikknev.Text4.Text = Trim(meszovtomb(vs&, 3))
                cikknev.Text4.SelStart = Len(Trim(meszovtomb(vs&, 3)))
                cikknev.Text5.Text = Trim(meszovtomb(vs&, 4))
                cikknev.Text5.SelStart = Len(Trim(meszovtomb(vs&, 4)))
                cikknev.Text6.Text = Trim(meszovtomb(vs&, 5))
                cikknev.Text6.SelStart = Len(Trim(meszovtomb(vs&, 5)))
                cikknev.Text2.Text = Trim(vtsztomb(vs&))
                cikknev.Text2.SelStart = Len(Trim(vtsztomb(vs&)))
                cikknev.Show vbModal
                meszovtomb(vs&, 1) = Trim(cikknev.Text1.Text)
                meszovtomb(vs&, 2) = Trim(cikknev.Text3.Text)
                meszovtomb(vs&, 3) = Trim(cikknev.Text4.Text)
                meszovtomb(vs&, 4) = Trim(cikknev.Text5.Text)
                meszovtomb(vs&, 5) = Trim(cikknev.Text6.Text)
                vtsztomb(vs&) = Trim(cikknev.Text2.Text)
              Loop While Trim(meszovtomb$(vs&, 1)) = ""
            End If
            If termrec$ <> "" Then
              form1.List2.Clear
              If Trim(Mid$(termrec$, 7, 60)) = "" Then
                form1.List2.AddItem meszovtomb$(vs&, 1)
              Else
                form1.List2.AddItem Mid$(termrec$, 7, 60)
                form1.List2.AddItem Mid$(termrec$, 67, 60)
              End If
              Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Mid$(termrec$, 140, 6)
              If xval(Mid$(termrec$, 146, 12)) <> 0 Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = Mid$(termrec$, 146, 12)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = Mid$(termrec$, 146, 12)
              End If
              Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = Mid$(termrec$, 161, 2)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = Mid$(termrec$, 163, 8)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 10) = Mid$(termrec$, 171, 8)
              Tabla.MSFlexGrid1.TextMatrix(vs&, 11) = Mid$(termrec$, 179, 8)
              If Trim(Mid$(partrec$, 783, 8)) <> "" Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = Mid$(partrec$, 783, 8)
              End If
              peng@ = xval(Mid$(partrec$, 333, 6))
              If peng@ <> 0 Then
                liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 4))
                penft@ = (liar@ * peng@) / 100
                peft$ = ertszam(Str$(penft@), 12, 2)
                penft@ = xval(peft$)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = Mid$(partrec$, 333, 6)
                Tabla.MSFlexGrid1.TextMatrix(vs&, 6) = peft$
                Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = ertszam(Str$(liar@ + penft@), 12, 2)
              End If
            End If
          End If
        End If
        If vo& = 1 Or vo& = 2 Or vo& = 4 Or vo& = 5 Or vo& = 6 Or vo& = 7 Then
          '--- számlaéerék kiszámítása
          liert@ = 0: enge@ = 0: elert@ = 0: afa@ = 0
          For i13% = 1 To 99
            tkod$ = Tabla.MSFlexGrid1.TextMatrix(i13%, 1)
            If Trim$(tkod$) <> "" Then
              afakod$ = Tabla.MSFlexGrid1.TextMatrix(i13%, 8)
              If afakod$ = utafakod$ Then
                afakulcs@ = utafakulcs@
              Else
                afrec$ = dbxkey("PAFA", afakod$)
                afakulcs@ = xval(Mid$(afrec$, 33, 6))
                utafakod$ = afakod$
                utafakulcs@ = afakulcs@
              End If
              menny@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 2))
              liar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 4))
              penft@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 6))
              elar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i13%, 7))
              If menny@ <> 0 Then
                alaposz@ = elar * menny@
                elert@ = elert@ + elar@ * menny@
                enge@ = enge@ + penft@ * menny@
              Else
                alaposz@ = elar@
                elert@ = elert@ + elar@
                enge@ = enge@ + penft@
              End If
              afaosz@ = (alaposz@ * afakulcs@) / 100
              afaker% = xval(Mid$(irec$, 345, 1))
              If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
              afaosz@ = xval(Format(afaosz@, fst$))
              afa@ = afa@ + afaosz@
            End If
          Next
          form1.List3.Clear
          form1.List3.AddItem langmodul(125) + ":" + ertszam(Str$(elert@ - enge@), 14, 2)
          form1.List3.AddItem langmodul(126) + ":" + ertszam(Str$(enge@), 14, 2)
          form1.List3.AddItem langmodul(127) + ":" + ertszam(Str$(elert@), 14, 2)
          form1.List3.AddItem langmodul(128) + ":" + ertszam(Str$(elert@ + afa@), 14, 2)
        End If
      End If
    Case Else
  End Select
  If mezohiba% = -1 Then mezohiba% = 1
End Sub

Public Sub kiszamolja(biznetto@, bizbrutto@, fogyert@, tip%)
  biznetto@ = 0: bizbrutto@ = 0: fogyert@ = 0
  For i121% = 1 To Tabla.MSFlexGrid1.Rows - 1
    If tip% = 1 Then
      biztkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(i121%, 1) + Space(15), 15)
    Else
      biztkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(i121%, 3) + Space(15), 15)
    End If
    If Trim(biztkod$) <> "" Then
      biztrec$ = dbxkey("KTRM", biztkod$)
      If biztrec$ <> "" Then
        afa@ = xval(Mid$(biztrec$, 706, 2))
        If tip% = 1 Then
          menny@ = xval(Tabla.MSFlexGrid1.TextMatrix(i121%, 2))
          bizar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i121%, 3))
          tpgar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i121%, 4))
        Else
          If tip% = 2 Then
            menny@ = xval(Tabla.MSFlexGrid1.TextMatrix(i121%, 5))
            bizar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i121%, 7))
            tpgar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i121%, 4))
          Else
            menny@ = xval(Tabla.MSFlexGrid1.TextMatrix(i121%, 4))
            bizar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i121%, 1))
            tpgar@ = xval(Tabla.MSFlexGrid1.TextMatrix(i121%, 2))
          End If
        End If
        fogar@ = xval(Mid$(biztrec$, 678, 14))
        biznetto@ = biznetto@ + menny@ * (tpgar@ + bizar@)
        If afa@ = 5 Then
          bizbrutto@ = bizbrutto@ + (menny@ * (tpgar@ + bizar@)) * 1.05
        Else
          If afa@ = 20 Then
            bizbrutto@ = bizbrutto@ + (menny@ * (tpgar@ + bizar@)) * 1.2
          Else
            bizbrutto@ = bizbrutto@ + (menny@ * (tpgar@ + bizar@))
          End If
        End If
        fogyert@ = fogyert@ + menny@ * fogar@
      End If
    End If
  Next
End Sub


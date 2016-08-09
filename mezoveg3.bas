Attribute VB_Name = "mezoveg3"
 '--- Eszes saját programjaihoz
Dim gyariszamok$(200, 200), hivatkozas$(200, 200), gysz$(200)
Public gyszelso As Boolean
Public M3erakt$, M3emozg$, M3Tipus$, szlatip$, M3erteknap$, bankvalasztas%

Public Function megnevbeall$(termrec$)
          If megnbeal.Option1.Value = True Then
            megnevbeall$ = Mid$(termrec$, 7, 50)
          Else
            If megnbeal.Option2.Value = True Then
              megnevbeall$ = Mid$(termrec$, 190, 50)
            Else
              If megnbeal.Option3.Value = True Then
                megnevbeall$ = Mid$(termrec$, 240, 50)
              Else
                megnevbeall$ = Mid$(termrec$, 67, 50)
              End If
            End If
          End If

End Function

Private Sub Gyariszambeker(megnev$, db@, sor&, iktato$, mozgtip, cikk$, raktar$)

             rogzites% = 0
             If db@ > 200 Then
                Call mess("Egy tételsorhoz csak 200 gyáriszámot rögzíthet! Több tételben vigye fel!", 2, 0, langmodul(99), valasz%)
                Exit Sub
             End If
             
             Gyariszam.MSFlexGrid1.Clear
             Gyariszam.MSFlexGrid1.Rows = 201
             Gyariszam.MSFlexGrid1.Cols = 3
             Gyariszam.MSFlexGrid1.FixedCols = 1
             Gyariszam.MSFlexGrid1.ColWidth(0) = 400
             Gyariszam.MSFlexGrid1.ColAlignment(0) = 1
             Gyariszam.MSFlexGrid1.TextMatrix(0, 0) = "Ssz"
             Gyariszam.MSFlexGrid1.ColWidth(1) = 4000
             Gyariszam.MSFlexGrid1.ColAlignment(1) = 1
             Gyariszam.biztip = mozgtip
             Gyariszam.Label2.Visible = True
             Gyariszam.Label3.Visible = True
             Gyariszam.Text2.Visible = True
             Gyariszam.Text3.Visible = True
             Gyariszam.MSFlexGrid1.TextMatrix(0, 1) = "S/N"
             If Val(iktato$) <> 0 Then
             ' módosítás
                For j10% = 1 To 199: gyariszamok$(sor&, j10%) = "": Next
                fil1 = FreeFile
                
                Open auditorutvonal$ + "auwker2.dbx" For Binary Shared As #fil1

                
                
                For i2% = 1 To 200
                  kulcs$ = iktato$ + Right("000" + Trim(Str(i2%)), 3)
                  kkfxrec$ = dbxkey("KKFX", kulcs$)
                  ' Betölt
                  gyariszamok$(sor&, i2%) = Mid$(kkfxrec$, 200, 40)
                  hivatkozas$(sor&, i2%) = Mid$(kkfxrec$, 48, 7)
                Next
             End If
             ' Feltöltés
             For j1% = 1 To 200
                Gyariszam.MSFlexGrid1.TextMatrix(j1%, 0) = Trim(Str(j1%))
                If Not Trim(gyariszamok$(sor&, j1%)) = "" Then
                   Gyariszam.MSFlexGrid1.TextMatrix(j1%, 1) = gyariszamok$(sor&, j1%)
                   Gyariszam.MSFlexGrid1.TextMatrix(j1%, 2) = hivatkozas$(sor&, j1%)
                End If
             Next
             Gyariszam.darab = db@
             
             Gyariszam.raktar = raktar$
             Gyariszam.cikk = cikk$
             
             Gyariszam.Label1.Caption = megnev$
             Gyariszam.Show vbModal
             If rogzites% = 1 Then
                For j1% = 1 To 200
                   If Not Trim(Gyariszam.MSFlexGrid1.TextMatrix(j1%, 1)) = "" Then
                      gyariszamok$(sor&, j1%) = Gyariszam.MSFlexGrid1.TextMatrix(j1%, 1)
                      hivatkozas$(sor&, j1%) = Gyariszam.MSFlexGrid1.TextMatrix(j1%, 2)
                   Else
                      gyariszamok$(sor&, j1%) = ""
                      hivatkozas$(sor&, j1%) = ""
                   End If
                Next
             End If
End Sub
Public Sub GyariszamAtadM3(sor%, gy$(), db%)
db% = 0
For i1% = 1 To 200
  If Not Trim(gyariszamok$(sor%, i1%)) = "" Then
    gy$(i1%, 1) = gyariszamok$(sor%, i1%)
    gy$(i1%, 2) = hivatkozas$(sor%, i1%)
    db% = db% + 1
  End If
Next

End Sub
Public Sub GyariszamTorolM3()

For j1% = 1 To 200
  For i1% = 1 To 200
    gyariszamok$(j1%, i1%) = ""
  Next
Next


End Sub



Public Sub mezovege3(vsor%, voszl%, mezohiba%, abmod%)
'--- ide írhatsz

  Select Case programnev$
      Case "AUW-QBWS"
         vs& = vsor%: vo& = voszl%
         If (vs& = 2 Or vs& = 3) Then
            Vektor.Text1.Font = "ER Kurier 1251"
         Else
            Vektor.Text1.Font = "MS Sans Serif"
         End If
      Case "AUW-QLIK"
      vs& = vsor%: vo& = voszl%
      If vo& = 1 Then
         mezohiba% = -1
      End If
      Case "AUW-QREAN"
      vs& = vsor%: vo& = voszl%
      If vs& = 2 Then
        termkod$ = Left$(Vektor.MSFlexGrid1.TextMatrix(2, 1) + Space$(15), 15)
        termrec$ = dbxkey("KTRM", termkod$)
        If termrec$ <> "" Then
           Vektor.MSFlexGrid1.TextMatrix(3, 1) = Mid$(termrec$, 16, 24)
           Vektor.MSFlexGrid1.TextMatrix(4, 1) = Mid$(termrec$, 484, 6)
        End If
      End If
      Case "AUW-QPTRG", "AUW-QSPTRG", "AUW-QSOPTRG"
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
        If vs& = 4 And programnev$ = "AUW-QPTRG" Then
           If Dir(auditorutvonal$ + terminal$ + "szov.txt") = terminal$ + "szov.txt" Then
              szovegval.Show vbModal
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
        If vo& = 12 Then
           tipus$ = Tabla.MSFlexGrid1.TextMatrix(vs&, 2)
           If tipus$ = "S" Then
              Call adoszamkot(vs&)
           End If
        End If
      End If
      ' Eszi - Surjány
      If vo& = 2 And programnev$ = "AUW-QSPTRG" Then
      tipu$ = Tabla.MSFlexGrid1.TextMatrix(vs&, vo&)
      If (tipu$ = "S" Or tipu$ = "V") And objektum$ = "PTET" And programnev$ = "AUW-QSPTRG" Then
        Tabla.MSFlexGrid1.TextMatrix(vs&, 9) = "KP"
        Tabla.MSFlexGrid1.TextMatrix(vs&, 11) = form1.Text12.Text
        
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

    Case "AUW-QRMEG"
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
      If objektum$ = "KXE1" Then
        tetelszam% = 0
        ertek@ = 0
        For i93% = 1 To 200
          
          If xval(Tabla3.MSFlexGrid1.TextMatrix(i93%, 2)) <> 0 Then tetelszam% = tetelszam% + 1
          ertek@ = ertek@ + xval(Tabla3.MSFlexGrid1.TextMatrix(i93%, 2)) * xval(Tabla3.MSFlexGrid1.TextMatrix(i93%, 3))
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
          termkod$ = Left(Tabla3.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
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
            'minta$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
            ' Eszi
            'diszt$ = Trim(Tabla.MSFlexGrid1.TextMatrix(vs&, 6))
            'egysar@ = arazo@(Mid$(partrec$, 1, 15), termkod$, minta$, erteknap$, diszt$)
            egysar@ = xval(Mid$(termrec$, 678, 14))
            Tabla3.MSFlexGrid1.TextMatrix(vs&, 3) = ertszam(Str(egysar), 12, 2)
          End If
        End If
        ' Eszi  - mintaszám
        'If vo& = 2 Then
        '  form1.MSFlexGrid4.Visible = False
        '  termkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
        '  minta$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
        '  If Trim(termkod$) <> "" And Trim(minta$) <> "" Then
        '    termrec$ = dbxkey("KTRM", termkod$)
        '    form1.List1.Clear
        '    form1.List1.AddItem Mid$(termrec$, 16, 60)
        '    form1.List1.AddItem Mid$(termrec$, 196, 60)
        '    form1.List1.AddItem langmodul(88) + " 1:" + Mid$(termrec$, 580, 14)
        '    'form1.List1.AddItem "Listaár 1:" + Mid$(termrec$, 580, 14)
        '    If Mid$(termrec$, 443, 1) = "C" Or Mid$(termrec$, 443, 1) = "M" Then
        '      Call rkeszlet(termkod$, minta$, Mid$(rakrec$, 1, 4), keszme@, foglme@)
        '      form1.List1.AddItem langmodul(89) + ":" + ertszam(Str(keszme@), 14, 2) + " " + langmodul(90) + ":" + ertszam(Str(foglme@), 14, 2) + " " + langmodul(91) + ":" + ertszam(Str(keszme@ - foglme@), 14, 2)
        '      'form1.List1.AddItem "Raktárkészlet:" + ertszam(Str(keszme@), 14, 2) + " Foglalt készlet:" + ertszam(Str(foglme@), 14, 2) + " Szabad készlet:" + ertszam(Str(keszme@ - foglme@), 14, 2)
        '    Else
        '     form1.List1.AddItem langmodul(95)
        '      'form1.List1.AddItem "Nincs készletnyilvántartás"
        '    End If
        '    '--- árazás
        '    minta$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
        '    'diszt$ = Trim(Tabla.MSFlexGrid1.TextMatrix(vs&, 6))
        '    'egysar@ = arazo@(Mid$(partrec$, 1, 15), termkod$, minta$, erteknap$, diszt$)
        '    egysar@ = xval(Mid$(termrec$, 678, 14))
        '    Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = ertszam(Str(egysar), 12, 2)
        '  End If
        'End If
      End If
     
 
   Case "AUW-QRTRM", "AUW-QNEW"
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
          'Vektor.MSFlexGrid1.TextMatrix(3, 1) = Mid$(termrec$, 1250, 24)
          Vektor.MSFlexGrid1.TextMatrix(3, 1) = Mid$(termrec$, 16, 24)
          Vektor.MSFlexGrid1.TextMatrix(4, 1) = Mid$(termrec$, 484, 6)
        End If
      End If
      If objektum$ = "KTRM" Then
        
         If vs& = 4 And ablaksorszam% = 3 Then
           If Not Vektor.MSFlexGrid1.TextMatrix(vs&, 1) = "27" Then
              Call mess("Hibás ÁFA kód!", 2, 0, langmodul(99), valasz%)
              Vektor.MSFlexGrid1.TextMatrix(vs&, 1) = ""
              mezohiba% = -1
           End If
         End If
         If vs& = 17 Then
           termkod$ = Left(Vektor.MSFlexGrid1.TextMatrix(1, 1) + Space$(15), 15)
           eankod$ = Left(Vektor.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(13), 13)
           If Not eankod$ = Space(13) Then
             eanrec$ = dbxkey("REAN", eankod$)
             If Not eanrec$ = "" Then
               If Not termkod$ = Mid$(eanrec$, 14, 15) Then
                 Call mess("Már van ilyen ean kód! /" + Mid$(eanrec$, 14, 15) + "/", 2, 0, langmodul(99), valasz%)
                 Vektor.MSFlexGrid1.TextMatrix(vs&, 1) = ""
                 mezohiba% = -1
               End If
             End If
           End If
         End If
      End If
    Case "AUW-QREGY"
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
            If Mid$(mozrec$, 47, 3) <> "   " Then
               Call mess("Kapcsolt mozgás tilos! /Raktárközit az U-új gombbal rögzítse!/", 2, 0, langmodul(99), valasz%)
               mezohiba% = -1
            End If
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
        If (vs& = 6) Or (vs& = 8) Then
            tip$ = Vektor.MSFlexGrid1.TextMatrix(3, 1)
            pkodszlasz$ = Vektor.MSFlexGrid1.TextMatrix(vs&, 1)
            If tip$ = "B" And Len(Trim$(pkodszlasz$)) = 0 Then
               Call mess("Partnerkódot/számlaszámot kötelezõ kitölteni", 2, 0, langmodul(99), valasz%)
               Vektor.MSFlexGrid1.Row = vs&
               mezohiba% = -1
            End If
        End If
        If (vs& = 7) Then
            tip$ = Vektor.MSFlexGrid1.TextMatrix(3, 1)
            pkodszlasz$ = Vektor.MSFlexGrid1.TextMatrix(vs&, 1)
            If tip$ = "F" And Len(Trim$(pkodszlasz$)) = 0 Then
               Call mess("Munkaszámot kötelezõ kitölteni", 2, 0, langmodul(99), valasz%)
               Vektor.MSFlexGrid1.Row = vs&
               mezohiba% = -1
            End If
        End If


      End If
      If objektum$ = "KXF1" Or objektum$ = "KXF3" Then
        If objektum$ = "KXF1" Then
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
            Mlaptetel1.raktar = rkod$
            Mlaptetel1.Show vbModal
            
          End If
          If vo& = 3 Then
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
            rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
            'If Not Mlaptetel1.mlellen(tkod$, rkod$) Then
            '   Call mess("Nincs ilyen termék a munkalapon!/Raktár jó?/", 2, 0, langmodul(99), valasz%)
            '   Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Space$(15)
            '   Tabla.Text1 = Space$(15)
            '   mezohiba% = -1
            'End If
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
                ' Eszi 2011.11.08
                If Mid$(termrec$, 846, 1) = "L" Then
                   Call mess("Termék letiltva!", 2, 0, langmodul(99), valasz%)
                    Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Space$(15)
                    Tabla.Text1 = Space$(15)
                    mezohiba% = -1
                End If
              End If
            End If
          End If
          If vo& = 4 Then
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
            rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
            meny@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, 4))
            If Not Mlaptetel1.mlellenMeny(tkod$, rkod$, meny@) Then
               Call mess("Hibás mennyiség!", 2, 0, langmodul(99), valasz%)
               mezohiba% = -1
            End If
          End If
          If vo& = 8 Then
            ujrkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 8) + "    ", 4)
            rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
            
            If M3Tipus$ = "R" And ujrkod$ = rkod$ Then
               Call mess("Új raktár egyenlõ a raktárral!", 2, 0, langmodul(99), valasz%)
               mezohiba% = -1
            End If
          
          End If
        End If
        If objektum$ = "KXF3" And vo& = 3 And Arszorzo1 = 1 Then
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
        If objektum$ = "KXF1" And vo& = 3 Or objektum$ = "KXF3" And (vo& = 1 Or vo& = 5) Then
          form1.List2.Clear
          If objektum$ = "KXF1" Then
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
'Ide kisker ár
              Else
                ' Eszi 2011.11.08
                If Mid$(termrec$, 846, 1) = "L" Then
                   Call mess("Termék letiltva!", 2, 0, langmodul(99), valasz%)
                   Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Space$(15)
                   Tabla.Text1 = Space$(15)
                   mezohiba% = -1
                Else

                form1.List2.Clear
                form1.List2.AddItem Mid$(termrec$, 16, 60)
                form1.List2.AddItem Mid$(termrec$, 196, 60)
                Select Case Mid$(termrec$, 442, 1)
                  Case "A", "R", "F", "K"
                    If Mid$(termrec$, 443, 1) <> "N" Then
                       If objektum$ = "KXF1" Then
                          rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
                       Else
                          rkod$ = M3erakt$
                       End If
                       rkszazon$ = rkod$ + tkod$
                       rkszrec$ = dbxkey("RKSZ", rkszazon$)
                       If rkszrec$ = "" Then
                          rkmenny@ = 0: rkfogl@ = 0: rkzarolt@ = 0: szabad@ = 0
                       Else
                          rkmenny@ = xval(Mid$(rkszrec$, 20, 12))
                          rkfogl@ = xval(Mid(rkszrec$, 32, 12))
                          rkzarolt@ = xval(Mid$(rkszrec$, 91, 12))
                          szabad@ = rkmenny@ - rkfogl@ - rkzarolt@
                       End If
                      
                      form1.List2.AddItem " Raktár készlet: " + Trim$(ertszam(Str$(szabad@), 14, 2)) + " " + Trim(Mid$(termrec$, 484, 6)) + " " + langmodul(97) + ":" + Trim$(ertszam(Mid$(termrec$, 762, 14), 14, 3)) + " " + langmodul(87) + ":" + Trim(ertszam(Mid$(termrec$, 552, 14), 14, 2)) + " " + langmodul(96) + ": " + Trim$(ertszam(Mid$(termrec$, 748, 14), 14, 3)) + "  Kisker ár: " + Trim(ertszam(Mid$(termrec$, 678, 14), 14, 2))
                      
                      
                      If vo& = 5 And Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = "I" Then
                        If Trim$(Mid$(termrec$, 456, 13)) = "" Then
                           Call mess("Nincs vonalkód a termékhez!", 2, 0, langmodul(99), valasz%)
                           Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = "  "
                        End If
                      End If
                    Else
                      Call mess(langmodul(98), 2, 0, langmodul(99), valasz%)
                      form1.List2.Clear
                    End If
                  Case "S"
                  ' 2009.09.18 - kivesz
                  '  Call mess(langmodul(100), 2, 0, langmodul(99), valasz%)
                  '  form1.List2.Clear
                  Case Else
                End Select
                End If
              End If
            Else
              form1.List2.Clear
            End If
          Else
            form1.List2.Clear
          End If
        End If
        If objektum$ = "KXF1" And vo& = 6 Or objektum$ = "KXF3" And vo& = 3 Then
' Eszi
'          form1.List2.Clear
          If objektum$ = "KXF1" Then
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
          Else
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          End If
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
               If Mid$(termrec$, 846, 1) = "L" Then
                  Call mess("Termék letiltva!", 2, 0, langmodul(99), valasz%)
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Space$(15)
                  Tabla.Text1 = Space$(15)
                  mezohiba% = -1
               Else

               If objektum$ = "KXF1" Then
                  menny@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, 4))
               Else
                  menny@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, 2))
               End If
               If objektum$ = "KXF1" Then
                 mkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + "   ", 3)
                 i12& = 1
                 Do While mkod$ = "   "
                   mkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs& - i12&, 1) + "   ", 3)
                   i12& = i12& + 1
                   If vs& = i12& Then
                     Exit Do
                   End If
                 Loop
               Else
                 mkod$ = M3emozg
               End If
               mozgrec$ = dbxkey("KMOX", mkod$)
               If mozgrec$ <> "" Then
                  mozgtip$ = Mid$(mozgrec$, 43, 2)
               End If
               If objektum$ = "KXF1" Then
                 rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
                 i12& = 1
                 Do While rkod$ = "    "
                   rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs& - i12&, 2) + "    ", 4)
                   i12& = i12& + 1
                   If vs& = i12& Then
                     Exit Do
                   End If
                 Loop
               
               Else
                 rkod$ = M3erakt$
               End If

               
               If Mid$(termrec$, 849, 1) = "I" Then
                 Call Gyariszambeker(Trim(Mid$(termrec$, 16, 60)), menny@, vs&, "0", mozgtip$, tkod$, rkod$)
               End If
               End If  ' Letiltva
            End If
          End If
        End If
      End If
      If objektum$ = "KKB1" Then
        If vs& = 3 Then
          tip$ = Vektor.MSFlexGrid1.TextMatrix(vs&, 1)
          If tip$ = "F" And Mid$(Vektor.MSFlexGrid1.TextMatrix(vs& - 1, 1), 1, 17) = "Kiadás munkalapra" Then
             ml$ = Mid$(Vektor.MSFlexGrid1.TextMatrix(vs& - 1, 1), 19, 7)
             Vektor.MSFlexGrid1.TextMatrix(vs& - 1, 1) = "Felh. munkalapról " + ml$
          End If
        End If
         If (vs& = 4) Or (vs& = 6) Then
            tip$ = Vektor.MSFlexGrid1.TextMatrix(3, 1)
            pkodszlasz$ = Vektor.MSFlexGrid1.TextMatrix(vs&, 1)
            If tip$ = "B" And Len(Trim$(pkodszlasz$)) = 0 Then
               Call mess("Partnerkódot/számlaszámot kötelezõ kitölteni", 2, 0, langmodul(99), valasz%)
               Vektor.MSFlexGrid1.Row = vs&
               mezohiba% = -1
            End If
         End If
         If (vs& = 5) Then
            tip$ = Vektor.MSFlexGrid1.TextMatrix(3, 1)
            pkodszlasz$ = Vektor.MSFlexGrid1.TextMatrix(vs&, 1)
            If tip$ = "F" And Len(Trim$(pkodszlasz$)) = 0 Then
               Call mess("Munkaszámot kötelezõ kitölteni", 2, 0, langmodul(99), valasz%)
               Vektor.MSFlexGrid1.Row = vs&
               mezohiba% = -1
            End If
         End If

      End If
      If objektum$ = "KKF2" Then
         If vo& = 3 Then
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(15), 15)
          iktato$ = Tabla.MSFlexGrid1.TextMatrix(vs&, 4)
          ktrec$ = dbxkey("KKFT", iktato$)
          If Not Trim(ktrec$) = "" Then
           emozg$ = Mid$(ktrec$, 21, 3)
           erakt$ = Mid$(ktrec$, 24, 4)
           mozgrec$ = dbxkey("KMOX", emozg$)
           If mozgrec$ <> "" Then
              mozgtip$ = Mid$(mozgrec$, 43, 2)
              mezohiba% = -1
           End If
      
          End If
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
               If Mid$(termrec$, 846, 1) = "L" Then
                  Call mess("Termék letiltva!", 2, 0, langmodul(99), valasz%)
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Space$(15)
                  Tabla.Text1 = Space$(15)
                  mezohiba% = -1
               Else

               menny@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, 3))
               iktato$ = Tabla.MSFlexGrid1.TextMatrix(vs&, 4)
               
               If Mid$(termrec$, 849, 1) = "I" Then
               
                 Call Gyariszambeker(Trim(Mid$(termrec$, 16, 60)), menny@, vs&, iktato$, mozgtip$, tkod$, erakt$)
               End If
               End If
            End If
          End If
         
         End If
      End If
    Case "AUW-QREGY", "AUW-QREGYF"
      vs& = vsor%: vo& = voszl%
      If objektum$ = "KXB1" Or objektum$ = "KXB3" Then
         If vs& = 4 Then
           szvszla$ = Trim(Vektor.MSFlexGrid1.TextMatrix(2, 1))
           fajta$ = Trim(Vektor.MSFlexGrid1.TextMatrix(4, 1))
           If fajta$ = "B" And Trim$(szvszla$) = "" Then
               Call mess("Szerzõdés/számla dátumot kötelezõ kitölteni", 2, 0, langmodul(99), valasz%)
               Vektor.MSFlexGrid1.Row = 2
               mezohiba% = -1
           End If
         End If
      End If
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
            If Mid$(mozrec$, 47, 3) <> "   " Then
               Call mess("Kapcsolt mozgás tilos! /Raktárközit az U-új gombbal rögzítse!/", 2, 0, langmodul(99), valasz%)
               mezohiba% = -1
            End If
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
      
      If objektum$ = "KXF2" Then
        If vo& = 1 Then
            devn$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(3), 3)
            devbear@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 1))
            If devn$ = Space$(3) And devbear@ <> 0 Then
             Call mess("Deviza nem üres!", 2, 0, langmodul(99), valasz%)
             Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = Space$(12)
             Tabla.Text1 = Space$(12)
             mezohiba% = -1
            Else
             Call forintosit(1, 2, 3, pzbank$, M3erteknap$, "V", "T", vs&)
            End If

        End If
      End If

      
      If objektum$ = "KXF1" Or objektum$ = "KXF3" Then
        If objektum$ = "KXF1" Then
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
            Mlaptetel1.raktar = rkod$
            Mlaptetel1.Show vbModal
            
          End If
          If vo& = 3 Then
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
            rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
            mkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + "   ", 3)

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
                ' Eszi 2011.11.08
                If Mid$(termrec$, 846, 1) = "L" Then
                   Call mess("Termék letiltva!", 2, 0, langmodul(99), valasz%)
                    Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Space$(15)
                    Tabla.Text1 = Space$(15)
                    mezohiba% = -1
                End If
              
                
                Call SzabkeszletKi(rkod$, tkod$, mkod$, 14)

              End If
            End If
          End If
          'If vo& = 4 Then
          '  tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
          '  rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
          '  meny@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, 4))
          '  If Not Mlaptetel1.mlellenMeny(tkod$, rkod$, meny@) Then
          '     Call mess("Hibás mennyiség!", 2, 0, langmodul(99), valasz%)
          '     mezohiba% = -1
          '  End If
          'End If
          If vo& = 6 Then
            devn$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 5) + Space$(3), 3)
            devbear@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 6))
            If devn$ = Space$(3) And devbear@ <> 0 Then
             Call mess("Deviza nem üres!", 2, 0, langmodul(99), valasz%)
             Tabla.MSFlexGrid1.TextMatrix(vs&, 7) = Space$(12)
             Tabla.Text1 = Space$(12)
             mezohiba% = -1
            Else
             Call forintosit(6, 5, 7, pzbank$, M3erteknap$, "V", "T", vs&)
            End If

          End If
          If vo& = 9 Then
            ujrkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 8) + "    ", 4)
            rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
            
            If M3Tipus$ = "R" And ujrkod$ = rkod$ Then
               Call mess("Új raktár egyenlõ a raktárral!", 2, 0, langmodul(99), valasz%)
               mezohiba% = -1
            End If
          
          End If
        End If
        If objektum$ = "KXF3" And vo& = 1 Then
           tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
           Call SzabkeszletKi(M3erakt$, tkod$, M3emozg, 6)
        End If
        ' Devizás beszerzési ár
        If objektum$ = "KXF3" And vo& = 4 Then
          devn$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(3), 3)
          devbear@ = xval(Tabla.MSFlexGrid1.TextMatrix(vs&, 4))
          If devn$ = Space$(3) And devbear@ <> 0 Then
             Call mess("Deviza nem üres!", 2, 0, langmodul(99), valasz%)
             Tabla.MSFlexGrid1.TextMatrix(vs&, 4) = Space$(14)
             Tabla.Text1 = Space$(14)
             mezohiba% = -1
          Else
             Call forintosit(4, 3, 5, pzbank$, M3erteknap$, "V", "T", vs&)
          End If
        
        End If
        If objektum$ = "KXF3" And vo& = 3 And Arszorzo1 = 1 Then
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
        If objektum$ = "KXF1" And vo& = 3 Or objektum$ = "KXF3" And (vo& = 1 Or vo& = 5) Then
          form1.List2.Clear
          If objektum$ = "KXF1" Then
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
'Ide kisker ár
              Else
                ' Eszi 2011.11.08
                If Mid$(termrec$, 846, 1) = "L" Then
                   Call mess("Termék letiltva!", 2, 0, langmodul(99), valasz%)
                   Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Space$(15)
                   Tabla.Text1 = Space$(15)
                   mezohiba% = -1
                Else

                form1.List2.Clear
                form1.List2.AddItem Mid$(termrec$, 16, 60)
                form1.List2.AddItem Mid$(termrec$, 196, 60)
                Select Case Mid$(termrec$, 442, 1)
                  Case "A", "R", "F", "K"
                    If Mid$(termrec$, 443, 1) <> "N" Then
                       If objektum$ = "KXF1" Then
                          rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
                       Else
                          rkod$ = M3erakt$
                       End If
                       rkszazon$ = rkod$ + tkod$
                       rkszrec$ = dbxkey("RKSZ", rkszazon$)
                       If rkszrec$ = "" Then
                          rkmenny@ = 0: rkfogl@ = 0: rkzarolt@ = 0: szabad@ = 0
                       Else
                          rkmenny@ = xval(Mid$(rkszrec$, 20, 12))
                          rkfogl@ = xval(Mid(rkszrec$, 32, 12))
                          rkzarolt@ = xval(Mid$(rkszrec$, 91, 12))
                          szabad@ = rkmenny@ - rkfogl@ - rkzarolt@
                       End If
                      
                      form1.List2.AddItem " Raktár készlet: " + Trim$(ertszam(Str$(szabad@), 14, 2)) + " " + Trim(Mid$(termrec$, 484, 6)) + " " + langmodul(97) + ":" + Trim$(ertszam(Mid$(termrec$, 762, 14), 14, 3)) + " " + langmodul(87) + ":" + Trim(ertszam(Mid$(termrec$, 552, 14), 14, 2)) + " " + langmodul(96) + ": " + Trim$(ertszam(Mid$(termrec$, 748, 14), 14, 3)) + "  Kisker ár: " + Trim(ertszam(Mid$(termrec$, 678, 14), 14, 2))
                      
                      
                      If vo& = 5 And Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = "I" Then
                        If Trim$(Mid$(termrec$, 456, 13)) = "" Then
                           Call mess("Nincs vonalkód a termékhez!", 2, 0, langmodul(99), valasz%)
                           Tabla.MSFlexGrid1.TextMatrix(vs&, 5) = "  "
                        End If
                      End If
                    Else
                      Call mess(langmodul(98), 2, 0, langmodul(99), valasz%)
                      form1.List2.Clear
                    End If
                  Case "S"
                  ' 2009.09.18 - kivesz
                  '  Call mess(langmodul(100), 2, 0, langmodul(99), valasz%)
                  '  form1.List2.Clear
                  Case Else
                End Select
                End If
              End If
            Else
              form1.List2.Clear
            End If
          Else
            form1.List2.Clear
          End If
        End If
        If objektum$ = "KXF1" And vo& = 6 Or objektum$ = "KXF3" And vo& = 3 Then
' Eszi
'          form1.List2.Clear
          If objektum$ = "KXF1" Then
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 3) + Space$(15), 15)
          Else
            tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
          End If
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
               If Mid$(termrec$, 846, 1) = "L" Then
                  Call mess("Termék letiltva!", 2, 0, langmodul(99), valasz%)
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Space$(15)
                  Tabla.Text1 = Space$(15)
                  mezohiba% = -1
               Else

               If objektum$ = "KXF1" Then
                  menny@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, 4))
               Else
                  menny@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, 2))
               End If
               If objektum$ = "KXF1" Then
                 mkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + "   ", 3)
                 i12& = 1
                 Do While mkod$ = "   "
                   mkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs& - i12&, 1) + "   ", 3)
                   i12& = i12& + 1
                   If vs& = i12& Then
                     Exit Do
                   End If
                 Loop
               Else
                 mkod$ = M3emozg
               End If
               mozgrec$ = dbxkey("KMOX", mkod$)
               If mozgrec$ <> "" Then
                  mozgtip$ = Mid$(mozgrec$, 43, 2)
               End If
               If objektum$ = "KXF1" Then
                 rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + "    ", 4)
                 i12& = 1
                 Do While rkod$ = "    "
                   rkod$ = Left(Tabla.MSFlexGrid1.TextMatrix(vs& - i12&, 2) + "    ", 4)
                   i12& = i12& + 1
                   If vs& = i12& Then
                     Exit Do
                   End If
                 Loop
               
               Else
                 rkod$ = M3erakt$
               End If

               
               If Mid$(termrec$, 849, 1) = "I" Then
                 Call Gyariszambeker(Trim(Mid$(termrec$, 16, 60)), menny@, vs&, "0", mozgtip$, tkod$, rkod$)
               End If
               End If  ' Letiltva
            End If
          End If
        End If
      End If
      If objektum$ = "KKB1" Then
        If vs& = 3 Then
          tip$ = Vektor.MSFlexGrid1.TextMatrix(vs&, 1)
          If tip$ = "F" And Mid$(Vektor.MSFlexGrid1.TextMatrix(vs& - 1, 1), 1, 17) = "Kiadás munkalapra" Then
             ml$ = Mid$(Vektor.MSFlexGrid1.TextMatrix(vs& - 1, 1), 19, 7)
             Vektor.MSFlexGrid1.TextMatrix(vs& - 1, 1) = "Felh. munkalapról " + ml$
          End If
        End If
      End If
      If objektum$ = "KKF2" Then
         If vo& = 3 Then
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(15), 15)
          iktato$ = Tabla.MSFlexGrid1.TextMatrix(vs&, 4)
          ktrec$ = dbxkey("KKFT", iktato$)
          If Not Trim(ktrec$) = "" Then
           emozg$ = Mid$(ktrec$, 21, 3)
           erakt$ = Mid$(ktrec$, 24, 4)
           mozgrec$ = dbxkey("KMOX", emozg$)
           If mozgrec$ <> "" Then
              mozgtip$ = Mid$(mozgrec$, 43, 2)
           End If
      
          End If
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
               If Mid$(termrec$, 846, 1) = "L" Then
                  Call mess("Termék letiltva!", 2, 0, langmodul(99), valasz%)
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 3) = Space$(15)
                  Tabla.Text1 = Space$(15)
                  mezohiba% = -1
               Else

               menny@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, 3))
               iktato$ = Tabla.MSFlexGrid1.TextMatrix(vs&, 4)
               
               If Mid$(termrec$, 849, 1) = "I" Then
               
                 Call Gyariszambeker(Trim(Mid$(termrec$, 16, 60)), menny@, vs&, iktato$, mozgtip$, tkod$, erakt$)
               End If
               End If
            End If
          End If
         
         End If
      End If
    
    
        
    Case "AUW-QRSZL"
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
      If objektum$ = "KSZQ" Then
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
        If vo& = 2 And mezohiba% <> -1 Then
          tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(15), 15)
          'minta$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(15), 8)
          rakt$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 4)
          krakrec$ = Space$(200)
          If Trim$(rakt$) <> "" Then
            krakrec$ = dbxkey("KRAK", rakt$)
          End If
          If Trim$(tkod$) <> "" Then
            termrec$ = dbxkey("KTRM", tkod$)
            If termrec$ <> "" Then
              form1.List2.Clear
              form1.List2.AddItem Mid$(termrec$, 16, 60)
              form1.List2.AddItem Mid$(termrec$, 196, 60)
              'minta$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 2) + Space$(8), 8)
              minta$ = Space$(8)
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
              ' Eszi
              ' Ha belföldi a számla akkor 716, ha export akkor 1015
              If szlatip$ = "ES" Then
                ellszi$ = Mid$(termrec$, 1015, 24)
              Else
                ellszi$ = Mid$(termrec$, 716, 24)
              End If
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
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = ertszam(Str$(elar@), 12, 2)
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
   
   Case "AUW-QRSZS"
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
        If vo& = 4 Then
           tkod$ = Left$(Tabla.MSFlexGrid1.TextMatrix(vs&, 1) + Space$(15), 15)
           menny@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, 3))
           rkod$ = M3erakt$
           
           termrec$ = dbxkey("KTRM", tkod$)
           If Mid$(termrec$, 849, 1) = "I" Then
             Call Gyariszambeker(Trim(Mid$(termrec$, 16, 60)), menny@, vs&, "0", "ER", tkod$, rkod$)
           End If

        End If
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
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = ertszam(Str$(elar@), 12, 2)
              End If
            End If
          End If
        End If
        If vo& = 1 Or vo& = 3 Or vo& = 5 Or vo& = 6 Or vo& = 7 Or vo& = 8 Then
          '--- számlaéerék kiszámítása
          liert@ = 0: enge@ = 0: elert@ = 0: afa@ = 0
          gigsze% = 999
          If programnev$ = "AUW-QRSZS" Then gigsze% = 199
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
    Case "AUW-QSSR"
    ' Szerzõdések karbantartása
      vs& = vsor%: vo& = voszl%
      If objektum$ = "SZRZ" Then
         If ablaksorszam% = 1 Then
         If vs& = 2 Then
          pkod$ = Trim(Vektor.MSFlexGrid1.TextMatrix(vs&, 1))
          parrec$ = dbxkey("PART", pkod$)
          Vektor.MSFlexGrid1.TextMatrix(3, 1) = Mid$(parrec$, 16, 60)
          Vektor.MSFlexGrid1.TextMatrix(4, 1) = Mid$(parrec$, 106, 8)
          'Vektor.MSFlexGrid1.TextMatrix(5, 1) = Trim$(Mid$(parrec$, 114, 30)) + " " + Trim$(Mid$(parrec$, 134, 30)) + " " + Trim$(Mid$(parrec$, 174, 10))
          Vektor.MSFlexGrid1.TextMatrix(5, 1) = postacim(parrec$, 114)
         End If
         End If
      End If
      If objektum$ = "SZSZ" Then
         If vo& = 3 Or vo& = 4 Or vo& = 5 Then
            szamla$ = Trim$(Tabla.MSFlexGrid1.TextMatrix(vs&, 6))
            If szamla$ = "" Then
              altalany@ = Val(form1.MSFlexGrid1.TextMatrix(vo& + 3, 1))
              osszeg@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, vo&))
              If Not altalany@ = osszeg@ Then
               MsgBox "Ha a jõvõre is ezt az összeget szeretné számlázni, módosítsa a díj összegét az elsõ lapon is!", 48, "Figyelmezteés"
              End If
            Else
               MsgBox "Nem módosíthatja ezt a sort. Már készült számla", 48, "Hiba"
            End If
         End If
      End If
    Case "AUW-QSSR4"
    ' Szerzõdések karbantartása VÁRVAG
      vs& = vsor%: vo& = voszl%
      If objektum$ = "SZRZ" Then
         If ablaksorszam% = 1 Then
           If vs& = 2 Then
             pkod$ = Trim(Vektor.MSFlexGrid1.TextMatrix(vs&, 1))
             parrec$ = dbxkey("PART", pkod$)
             Vektor.MSFlexGrid1.TextMatrix(3, 1) = Mid$(parrec$, 16, 60)
             Vektor.MSFlexGrid1.TextMatrix(4, 1) = Mid$(parrec$, 106, 8)
             'Vektor.MSFlexGrid1.TextMatrix(5, 1) = Trim$(Mid$(parrec$, 114, 30)) + " " + Trim$(Mid$(parrec$, 134, 30)) + " " + Trim$(Mid$(parrec$, 174, 10))
             Vektor.MSFlexGrid1.TextMatrix(5, 1) = postacim(parrec$, 114)
         End If
         If vs& = 7 Then
         If Trim$(Vektor.MSFlexGrid1.TextMatrix(7, 1)) = "" Then

           Bankval.Show vbModal
           If Not bankvalasztas% = 0 Then
            fkkod$ = Left(Bankval.MSFlexGrid1.TextMatrix(bankvalasztas% - 1, 0) + Space(8), 8)
            fkkrec$ = dbxkey("FKSZ", fkkod$)
            Vektor.MSFlexGrid1.TextMatrix(7, 1) = Mid$(fkkrec$, 419, 8) + "-" + Mid$(fkkrec$, 427, 8) + "-" + Mid$(fkkrec$, 435, 8)
           Else
            MsgBox "Nem választott banszámla számot!", 48, "Hiba"
            mezohiba% = 1
           End If

           
           End If
         End If
         End If
      End If
      If objektum$ = "SZSZ" Then
         If vo& = 3 Or vo& = 4 Or vo& = 5 Then
            szamla$ = Trim$(Tabla.MSFlexGrid1.TextMatrix(vs&, 6))
            If szamla$ = "" Then
              altalany@ = Val(form1.MSFlexGrid1.TextMatrix(vo& + 3, 1))
              osszeg@ = Val(Tabla.MSFlexGrid1.TextMatrix(vs&, vo&))
              If Not altalany@ = osszeg@ Then
               MsgBox "Ha a jõvõre is ezt az összeget szeretné számlázni, módosítsa a díj összegét az elsõ lapon is!", 48, "Figyelmezteés"
              End If
            Else
               MsgBox "Nem módosíthatja ezt a sort. Már készült számla", 48, "Hiba"
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
   
    Case "AUW-QPSZL3", "AUW-QPSZL4", "AUW-QPSZL", "AUW-QPSZL5"
    
      
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
            If (Trim(Mid$(termrec$, 7, 60)) = "" And programnev$ = "AUW-QPSZL") Or programnev$ = "AUW-QPSZL3" Or programnev$ = "AUW-QPSZL4" Then
              Do
                If Trim(meszovtomb(vs&, 1)) = "" Then
                   meszovtomb(vs&, 1) = Trim(Mid$(termrec$, 7, 60))
                   If programnev$ = "AUW-QPSZL3" Then
                      meszovtomb(vs&, 2) = "CMR:"
                      meszovtomb(vs&, 3) = "TARGA:"
                   Else
                     meszovtomb(vs&, 2) = megnevbeall$(termrec$)
                   End If
                End If
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
              If programnev$ = "AUW-QPSZL3" Then
                If (Tabla.MSFlexGrid1.TextMatrix(vs&, 13 + iii%) = "") Then
                  Tabla.MSFlexGrid1.TextMatrix(vs&, 13 + iii%) = Mid$(termrec$, 67, 4)
                End If
              End If
              If (Tabla.MSFlexGrid1.TextMatrix(vs&, 7 + iii%) = "") Then
                Tabla.MSFlexGrid1.TextMatrix(vs&, 7 + iii%) = Mid$(termrec$, 146, 12)
              End If

              If programnev$ = "AUW-QPSZL3" Then
                 ' Doza
                 If szlatip$ = "BS" Then
                   Tabla.MSFlexGrid1.TextMatrix(vs&, 8 + iii%) = Mid$(termrec$, 161, 2)
                   Tabla.MSFlexGrid1.TextMatrix(vs&, 9 + iii%) = Mid$(termrec$, 163, 8)
                 Else
                   Tabla.MSFlexGrid1.TextMatrix(vs&, 9 + iii%) = Mid$(termrec$, 352, 8)
                   Tabla.MSFlexGrid1.TextMatrix(vs&, 8 + iii%) = Mid$(termrec$, 360, 2)
                 End If
              Else
                   Tabla.MSFlexGrid1.TextMatrix(vs&, 8 + iii%) = Mid$(termrec$, 161, 2)
                   Tabla.MSFlexGrid1.TextMatrix(vs&, 9 + iii%) = Mid$(termrec$, 163, 8)
              End If
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
              
              afaker% = xval(Mid$(irec$, 345, 1))
              If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
              
              If form1.Check1.Value = 1 Then
                afaosz@ = alaposz@ - (alaposz@ / ((100 + afakulcs@) / 100))
                ' az ÁFÁt tételenként kerekíti
                fst$ = "#############0.00"
                afaosz@ = xval(Format(afaosz@, fst$))
                afa@ = afa@ + afaosz@
              Else
                afaosz@ = (alaposz@ * afakulcs@) / 100
                ' az ÁFÁt tételenként kerekíti
                afaosz@ = xval(Format(afaosz@, fst$))
                afa@ = afa@ + afaosz@
              End If
            End If
          Next
          
          form1.List3.Clear
          form1.List3.AddItem langmodul(125) + ":" + ertszam(Str$(elert@ - enge@), 14, 2)
          'form1.List3.AddItem "Lista érték :" + ertszam(Str$(elert@ - enge@), 14, 2)
          form1.List3.AddItem langmodul(126) + ":" + ertszam(Str$(enge@), 14, 2)
          'form1.List3.AddItem "Engedmény   :" + ertszam(Str$(enge@), 14, 2)
          If form1.Check1.Value = 1 Then
             form1.List3.AddItem langmodul(127) + ":" + ertszam(Str$(elert@ - afa@), 14, 2)
             'form1.List3.AddItem "Nettó érték :" + ertszam(Str$(elert@), 14, 2)
             form1.List3.AddItem langmodul(128) + ":" + ertszam(Str$(elert@), 14, 2)
             'form1.List3.AddItem "Bruttó érték:" + ertszam(Str$(elert@ + afa@), 14, 2)
          Else
             form1.List3.AddItem langmodul(127) + ":" + ertszam(Str$(elert@), 14, 2)
             'form1.List3.AddItem "Nettó érték :" + ertszam(Str$(elert@), 14, 2)
             form1.List3.AddItem langmodul(128) + ":" + ertszam(Str$(elert@ + afa@), 14, 2)
             'form1.List3.AddItem "Bruttó érték:" + ertszam(Str$(elert@ + afa@), 14, 2)
          End If
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
          If dnm$ = "EUR" Then
            dkod$ = bankszamla$ + "HUF"
          Else
            dkod$ = bankszamla$ + dnm$
          End If
          devrec$ = dbxkey("PDEV", dkod$)
          If devrec$ = "" Then
            '--- nincs pdev rekord
          Else
            If dnm$ = "EUR" Then
               dkod$ = bankszamla$ + "HUF" + erteknap$
            Else
               dkod$ = bankszamla$ + dnm$ + erteknap$
            End If
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
    Case "AUW-QRSZL3"
      vs& = vsor%: vo& = voszl%
      Call palmiv(vs&, vo&)
    Case Else
  End Select
  If mezohiba% = -1 Then mezohiba% = 1

End Sub
Sub palmie(vs&, vo&)
      
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
End Sub
Sub palmiv(vs&, vo&)
     
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
                Tabla.MSFlexGrid1.TextMatrix(vs&, 8) = ertszam(Str$(elar@), 12, 2)
              End If
              cikknev.Label3.Caption = Str(vs&) + "."
              cikknev.Text1.Text = Trim(meszovtomb(vs&, 1))
              cikknev.Text1.SelStart = Len(Trim(meszovtomb(vs&, 1)))
              cikknev.Text3.Text = Trim(meszovtomb(vs&, 2))
              cikknev.Show vbModal
              meszovtomb(vs&, 1) = Trim(cikknev.Text1.Text)
              meszovtomb(vs&, 2) = Trim(cikknev.Text3.Text)

              
            End If
          End If
        End If
        If vo& = 2 And mezohiba% <> -1 Then
              cikknev.Label3.Caption = Str(vs&) + "."
              cikknev.Text1.Text = Trim(meszovtomb(vs&, 1))
              cikknev.Text1.SelStart = Len(Trim(meszovtomb(vs&, 1)))
              cikknev.Text3.Text = Trim(meszovtomb(vs&, 2))
              cikknev.Show vbModal
              meszovtomb(vs&, 1) = Trim(cikknev.Text1.Text)
              meszovtomb(vs&, 2) = Trim(cikknev.Text3.Text)
        
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

End Sub

Sub SzabkeszletKi(rkod$, tkod$, mkod$, oszlop%)
             mozrec$ = dbxkey("KMOX", mkod$)
             irany$ = Mid$(mozrec$, 34, 1)
             If irany$ = "K" Then
                
                raktrec$ = dbxkey("KRAK", rkod$)
                termrec$ = dbxkey("KTRM", tkod$)
                Szabkeszl1.Label3 = Mid$(raktrec$, 5, 60)
                Szabkeszl1.Label4 = Mid$(termrec$, 16, 60)
                
                kszabrec$ = dbxkey("RKSX", rkod$ + tkod$)
              
                If kszabrec$ <> "" Then
                  Szabkeszl1.List1.Clear
                  For i% = 1 To 200
                     elem$ = Mid$(kszabrec$, (i% - 1) * 40 + 20, 40)
                     If Trim$(elem$) <> "" Then
                       Ikt$ = Mid$(elem$, 1, 7)
                       osszmenny@ = xval(Mid$(elem$, 8, 12))
                       fogymenny@ = xval(Mid$(elem$, 20, 12))
                       kiadhato@ = osszmenny@ - fogymenny@
                       kkftrec$ = dbxkey("KKFT", Ikt$)
                       bikt$ = Mid$(kkftrec$, 8, 7)
                       kkbzrec$ = dbxkey("KKBZ", bikt$)
                       pkod$ = Mid$(kkbzrec$, 55, 15)
                       devar@ = xval(Mid$(kkftrec$, 106, 10))
                       If Trim$(pkod$) = "" Then
                         partrec$ = Space(900)
                       Else
                         partrec$ = dbxkey("PART", pkod$)
                       End If
                       Szabkeszl1.List1.AddItem (Ikt$ + " " + datki(Mid$(kkftrec$, 15, 6)) + " " + Mid$(partrec$, 16, 40) + " " + Right$(Space(12) + Format(osszmenny@, "###########0.000"), 12) + " " + Right$(Space(12) + Format(kiadhato@, "###########0.000"), 12) + " " + Mid$(kkftrec$, 59, 12) + " " + Mid$(kkftrec$, 51, 3) + " " + Right$(Space(12) + Format(devar@, "###### ### ##0"), 12))
                     
                     End If
                  Next
                End If
                Szabkeszl1.oszl = oszlop%
                Szabkeszl1.Show vbModal
             End If
End Sub
Sub adoszamkot(vs&)
  afakod$ = Tabla.MSFlexGrid1.TextMatrix(vs&, 12)
  afrec$ = dbxkey("PAFA", afakod$)
  If afrec$ <> "" Then
     afkod$ = Mid$(afrec$, 40, 2)
     If afkod$ = "BR" Then
        patkod$ = Tabla.MSFlexGrid1.TextMatrix(vs&, 9)
        partrec$ = dbxkey("PART", patkod$)
        If partrec$ <> "" Then
           adoszam$ = Mid$(partrec$, 184, 15)
           If Trim$(adoszam$) = "" Then
              Call mess("Fordított ÁFA esetén vevõ adószáma kötelezõ!", 2, 0, langmodul(99), valasz%)
           End If
        End If
     End If
  End If
           
End Sub

Attribute VB_Name = "Infolap"
Public infolapbe%
Public Sub infomutat2(objneve$, rec$, hivashely$)
  '--- keresõtábla információ felmutatása
  Dim Tbknev As Object
  If hivashely$ = "Kereso" Then
    Set Tbknev = Kereso
  Else
    Set Tbknev = Hash
  End If
  Select Case objneve$
    Case "KKFT"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      tkod$ = Mid$(rec$, 36, 15)
      trec$ = dbxkey("KTRM", tkod$)
      If trec$ <> "" Then
        Tbknev.List1.AddItem Mid$(trec$, 16, 50)
      End If
    Case "JRAK"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      Select Case Mid$(rec$, 125, 1)
        Case "A": Tbknev.List1.AddItem "Egyszerûsített adóraktár"
        Case "J": Tbknev.List1.AddItem "Jövedéki engedélyes"
        Case "B": Tbknev.List1.AddItem "Bejegyzett fogadóhely"
        Case "M": Tbknev.List1.AddItem "Mûködési engedélyes"
        Case Else
      End Select
      Tbknev.List1.AddItem "Eng.szám: " + Trim(Mid$(rec$, 126, 13))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Borkisérõ okmány:"
      Tbknev.List1.AddItem "   " + Trim(Mid$(rec$, 260, 2)) + Trim(Mid$(rec$, 262, 10)) + "-" + Trim(Mid$(rec$, 260, 2)) + Trim(Mid$(rec$, 272, 10))
      Tbknev.List1.AddItem "Egyszerûsített kisérõ okmány:"
      Tbknev.List1.AddItem "   " + Trim(Mid$(rec$, 292, 2)) + Trim(Mid$(rec$, 294, 10)) + "-" + Trim(Mid$(rec$, 292, 2)) + Trim(Mid$(rec$, 304, 10))
      Tbknev.List1.AddItem "Adminisztratív kisérõ okmány:"
      Tbknev.List1.AddItem "   " + Trim(Mid$(rec$, 324, 2)) + Trim(Mid$(rec$, 326, 10)) + "-" + Trim(Mid$(rec$, 324, 2)) + Trim(Mid$(rec$, 336, 10))
    Case "JGYR"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      trec$ = torolvas("JRAK", Mid$(rec$, 8, 4), 1, 400)
      Tbknev.List1.AddItem "Raktár: "
      Tbknev.List1.AddItem Trim(Mid$(trec$, 5, 60))
      trec$ = torolvas("JTRM", Mid$(rec$, 18, 15), 1, 200)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Termék: "
      Tbknev.List1.AddItem Trim(Mid$(trec$, 90, 40))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Gyártás kezdete: " + datki(Mid$(rec$, 12, 6))
      Tbknev.List1.AddItem "Szárm: " + Trim(Mid$(rec$, 33, 20))
      Tbknev.List1.AddItem "OBI  : " + Trim(Mid$(rec$, 53, 20))
      If Mid$(rec$, 160, 1) = "B" Then
        Tbknev.List1.AddItem "B é r m u n k a "
      End If
      If Mid$(rec$, 73, 1) = "S" Then
        Tbknev.List1.AddItem "S z t o r n ó z v a!"
        Tbknev.List1.ForeColor = QBColor(12)
      End If
    Case "JTAR"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      trec$ = torolvas("JRAK", Mid$(rec$, 65, 4), 1, 400)
      Tbknev.List1.AddItem "Raktár: "
      Tbknev.List1.AddItem Trim(Mid$(trec$, 5, 60))
      Tbknev.List1.AddItem " "
      Select Case Mid$(rec$, 129, 1)
        Case "U"
          Tbknev.List1.AddItem "Ü r e s"
          Tbknev.List1.ForeColor = QBColor(2)
        Case "B"
          Tbknev.List1.AddItem "B o r"
          trec$ = torolvas("JTRM", Mid$(rec$, 130, 15), 1, 200)
          Tbknev.List1.AddItem Trim(Mid$(trec$, 90, 40))
        Case "M"
          Tbknev.List1.AddItem "M e l l é k t e r m é k"
          trec$ = torolvas("JMTR", Mid$(rec$, 145, 4), 1, 100)
          Tbknev.List1.AddItem Trim(Mid$(trec$, 5, 40))
        Case "A"
          Tbknev.List1.AddItem "A n y a g"
          trec$ = torolvas("JANY", Mid$(rec$, 149, 15), 1, 100)
          Tbknev.List1.AddItem Trim(Mid$(trec$, 16, 40))
        Case "S": Tbknev.List1.AddItem "S z a b a d t é r"
        Case "E": Tbknev.List1.AddItem "E g y é b  t e r m é k"
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Ûrtartalom :" + ertszam(Trim(Mid$(rec$, 41, 12)), 14, 2)
      Tbknev.List1.AddItem "Mennyiség  :" + ertszam(Trim(Mid$(rec$, 164, 14)), 14, 2)
      szh@ = xval(Mid$(rec$, 41, 12)) - xval(Mid$(rec$, 164, 14))
      Tbknev.List1.AddItem "Szabad hely:" + ertszam(Str$(szh@), 14, 2)
      If szh@ = 0 Then Tbknev.List1.ForeColor = QBColor(12)
    Case "JMUV"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      xx$ = " mûvelet"
      Select Case Mid$(rec$, 4, 1)
        Case "B": Tbknev.List1.AddItem "Bor" + xx$
        Case "S": Tbknev.List1.AddItem "Szõlõ" + xx$
        Case "A": Tbknev.List1.AddItem "Anyag" + xx$
        Case "M": Tbknev.List1.AddItem "Melléktermék" + xx$
        Case "G": Tbknev.List1.AddItem "Göngyöleg" + xx$
        Case "E": Tbknev.List1.AddItem "Egyéb jövedéki termék" + xx$
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      Select Case Mid$(rec$, 5, 1)
        Case "V": Tbknev.List1.AddItem "Vásárlás"
        Case "F": Tbknev.List1.AddItem "Felvásárlás"
        Case "E": Tbknev.List1.AddItem "Értékesítés"
        Case "G": Tbknev.List1.AddItem "Gyártás"
        Case "B": Tbknev.List1.AddItem "Belsõ mozgás"
        Case "H": Tbknev.List1.AddItem "Hiány"
        Case "T": Tbknev.List1.AddItem "Többlet"
        Case "K": Tbknev.List1.AddItem "Korrekció"
        Case "Y": Tbknev.List1.AddItem "Veszteség"
        Case "K": Tbknev.List1.AddItem "Korrekció"
        Case "S": Tbknev.List1.AddItem "Szüret"
        Case "N": Tbknev.List1.AddItem "Nyitó"
        Case "M": Tbknev.List1.AddItem "Bérmunka átadás-átvét"
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      xx$ = " tranzakció"
      Select Case Mid$(rec$, 6, 1)
        Case "B": Tbknev.List1.AddItem "Belföldi" + xx$
        Case "E": Tbknev.List1.AddItem "Közösségi" + xx$
        Case "X": Tbknev.List1.AddItem "Egyéb külföldi" + xx$
        Case "S": Tbknev.List1.AddItem "Saját kiskerhez kapcsolódó " + xx$
        Case "N": Tbknev.List1.AddItem "Saját nagykerhez kapcsolódó " + xx$
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      xx$ = " veszteség"
      Select Case Mid$(rec$, 7, 1)
        Case "T": Tbknev.List1.AddItem "Tárolási" + xx$
        Case "S": Tbknev.List1.AddItem "Szállítási" + xx$
        Case "M": Tbknev.List1.AddItem "Mûveleti" + xx$
        Case "K": Tbknev.List1.AddItem "Kiszerelési" + xx$
        Case Else
      End Select
    Case "JHZN"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      Select Case Mid$(rec$, 18, 1)
        Case "N": Tbknev.List1.AddItem "Nyitókészlet"
        Case "B": Tbknev.List1.AddItem "Beszerzés"
        Case "F": Tbknev.List1.AddItem "Felhasználás"
        Case "M": Tbknev.List1.AddItem "Megsemmisülés"
        Case "S": Tbknev.List1.AddItem "Selejtezés"
        Case "H": Tbknev.List1.AddItem "Hiány"
        Case "T": Tbknev.List1.AddItem "Többlet"
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Dátum :    " + datki(Mid$(rec$, 12, 6))
      Tbknev.List1.AddItem "Mennyiség :" + Mid$(rec$, 19, 8) + " db"
      Tbknev.List1.AddItem " "
      trec$ = torolvas("JHFJ", Mid$(rec$, 8, 4), 1, 100)
      Tbknev.List1.AddItem Trim(Mid$(trec$, 5, 30))
      If Mid$(rec$, 164, 1) = "K" Then
        Tbknev.List1.AddItem "Kannára"
      Else
        Tbknev.List1.AddItem "Hordóra"
      End If
      If Mid$(rec$, 57, 1) = "S" Then
        Tbknev.List1.AddItem "S z t o r n ó z v a!"
        Tbknev.List1.ForeColor = QBColor(12)
      End If
    Case "KPAR"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      Tbknev.List1.AddItem "Készpénzes partnerek"
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Név:"
      Tbknev.List1.AddItem Mid$(rec$, 1, 60)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Cím:"
      Tbknev.List1.AddItem Mid$(rec$, 61, 60)
    Case Else
      Tbknev.List1.Clear
      Tbknev.List1.AddItem "Nincs egyéb információ!"
  End Select
End Sub
Public Sub infomutat(objneve$, rec$, hivashely$)
  '--- keresõtábla információ felmutatása
  Dim Tbknev As Object
  If hivashely$ = "Kereso" Then
    Set Tbknev = Kereso
  Else
    Set Tbknev = Hash
  End If
  Tbknev.List1.ForeColor = QBColor(1)
  Select Case objneve$
    Case "PKOT"
      Tbknev.List1.ForeColor = QBColor(0)
      Tbknev.List1.Clear
      pkod$ = Mid$(rec$, 76, 15)
      partrec$ = dbxkey("PART", pkod$)
      Tbknev.List1.AddItem Trim(Mid$(partrec$, 16, 60))
      If Mid$(rec$, 91, 1) = "V" Then
        Tbknev.List1.AddItem "Követelés"
      Else
        Tbknev.List1.AddItem "Kötelezettség"
      End If
      Select Case Mid$(rec$, 92, 1)
        Case "A": Tbknev.List1.AddItem "Költségvetéssel szemben"
        Case "H": Tbknev.List1.AddItem "Hitel"
        Case "K": Tbknev.List1.AddItem "Kölcsön"
        Case "Z": Tbknev.List1.AddItem "Hozam"
        Case "E": Tbknev.List1.AddItem "Egyéb"
        Case Else
      End Select
      xx$ = "Gyakoriság: "
      Select Case Mid$(rec$, 93, 1)
        Case "E": Tbknev.List1.AddItem xx$ + "eseti"
        Case "H": Tbknev.List1.AddItem xx$ + "havi"
        Case "N": Tbknev.List1.AddItem xx$ + "negyedéves"
        Case "V": Tbknev.List1.AddItem xx$ + "éves"
        Case Else
      End Select
      Tbknev.List1.AddItem "Elsõ esedékesség  : " + datki(Mid$(rec$, 94, 6))
      Tbknev.List1.AddItem "Utolsó esedékesség: " + datki(Mid$(rec$, 100, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Teljes összeg:" + ertszamx(Mid$(rec$, 106, 14), 17, 0)
      Tbknev.List1.AddItem "Eseti összeg :" + ertszamx(Mid$(rec$, 120, 14), 17, 0)
      If Mid$(rec$, 134, 1) = "A" Then
        Tbknev.List1.AddItem "Aktív"
      Else
        Tbknev.List1.AddItem "Passzív"
        Tbknev.List1.ForeColor = QBColor(12)
      End If
    Case "PAFA"
      Tbknev.List1.ForeColor = QBColor(0)
      Tbknev.List1.Clear
      Tbknev.List1.AddItem Mid$(rec$, 3, 30)
      Tbknev.List1.AddItem "ÁFA kulcs :" + Trim(Mid$(rec$, 33, 6)) + " %"
      Tbknev.List1.AddItem " "
      Select Case Mid$(rec$, 40, 2)
        Case "BF"
          Tbknev.List1.AddItem "Belföldi forgalom"
        Case "XU"
          Tbknev.List1.AddItem "Közösségi értékesítés"
        Case "XE"
          Tbknev.List1.AddItem "Egyéb export"
        Case "IU"
          Tbknev.List1.AddItem "Közösségi beszerzés"
        Case "IE"
          Tbknev.List1.AddItem "Egyéb import"
        Case Else
      End Select
      Select Case Mid$(rec$, 39, 1)
        Case "A"
          Tbknev.List1.AddItem "Adóalap"
        Case "M"
          Tbknev.List1.AddItem "Adómentes"
        Case "N"
          Tbknev.List1.AddItem "Adóalapot nem képezõ"
        Case Else
      End Select
      Select Case Mid$(rec$, 42, 1)
        Case "T"
          Tbknev.List1.AddItem "Teljes egészében visszaigényelhetõ"
        Case "R"
          Tbknev.List1.AddItem "Részben visszaigénelhetõ"
        Case "N"
          Tbknev.List1.AddItem "Nem igényelhetõ vissza"
        Case Else
      End Select
      If Mid$(rec$, 42, 1) = "R" Then Tbknev.List1.AddItem "Visszaigényelhetõ az adó " + Trim(Mid$(rec$, 43, 6)) + " %-a"
    Case "GYUJ"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(0)
      Tbknev.List1.AddItem "Megrendelés iktató: " + Mid$(rec$, 8, 7)
      If Mid$(rec$, 129, 1) = "S" Then
        Tbknev.List1.AddItem "S z t o r n ó z o t t  gyûjtõ"
        Tbknev.List1.ForeColor = QBColor(12)
      End If
      Tbknev.List1.AddItem " "
      kszc$ = Mid$(rec$, 15, 15)
      pkod$ = Mid$(rec$, 30, 15)
      krak$ = Mid$(rec$, 45, 4)
      If Trim(pkod$) <> "" Then
        prec$ = dbxkey("PART", pkod$)
        Tbknev.List1.AddItem "Fizetõ:"
        Tbknev.List1.AddItem Trim(Mid$(prec$, 16, 60))
        Tbknev.List1.AddItem Trim(postacim(prec$, 106))
        Tbknev.List1.AddItem " "
      End If
      If Trim(kszc$) <> "" Then
        prec$ = dbxkey("KSZC", kszc$)
        Tbknev.List1.AddItem "Szállítási cím:"
        Tbknev.List1.AddItem Trim(Mid$(prec$, 31, 60))
        Tbknev.List1.AddItem Trim(postacim(prec$, 121))
        Tbknev.List1.AddItem " "
      End If
      If Trim(krak$) <> "" Then
        prec$ = dbxkey("KRAK", krak$)
        Tbknev.List1.AddItem "Raktár:"
        Tbknev.List1.AddItem Trim(Mid$(prec$, 5, 60))
        Tbknev.List1.AddItem Trim(Mid$(prec$, 68, 60))
        Tbknev.List1.AddItem " "
      End If
      tdb% = 0
      For i81% = 1 To 200
        elem$ = Mid$(rec$, (i81% - 1) * 25 + 200, 25)
        If Trim(elem$) <> "" Then
          tdb% = tdb% + 1
        End If
      Next
      Tbknev.List1.AddItem "Tételszám    :" + Trim(ertszam(Str(tdb), 10, 0)) + " db"
      megikt$ = Mid$(rec$, 8, 7)
      megrec$ = dbxkey("KMEG", megikt$)
      If megrec$ <> "" Then
        bte@ = 0
        For i81% = 1 To 200
          elem$ = Mid$(megrec$, (i81% - 1) * 59 + 500, 59)
          If Trim(elem$) <> "" Then
            bte@ = bte@ + xval(Mid$(elem$, 24, 12)) * xval(Mid$(elem$, 36, 12))
          End If
        Next
        Tbknev.List1.AddItem "Bruttó érték:" + ertszamx(Str(bte@), 15, 0)
      End If
    Case "FUVA"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(0)
      If Mid$(rec$, 212, 1) = "*" Then
        Tbknev.List1.AddItem "T ö r ö l t fuvarösszesítõ"
      Else
        szla% = 0: szlev% = 0: rkozi% = 0: nte@ = 0: bte@ = 0
        For i81% = 1 To 40
          megikt$ = Mid$(rec$, (i81% - 1) * 50 + 300, 7)
          If Trim(megikt$) <> "" Then
            megrec$ = dbxkey("ERTB", megikt$)
            If megrec$ <> "" Then
              If Mid$(megrec$, 201, 1) <> "S" Then
                If Mid$(megrec$, 26, 1) = "S" Then szla = szla + 1
                If Mid$(megrec$, 26, 1) = "L" Then szlev = szlev + 1
                If Mid$(megrec$, 26, 1) = "R" Then rkozi = rkozi + 1
                nte@ = nte@ + xval(Mid$(megrec$, 145, 10))
                bte@ = bte@ + xval(Mid$(megrec$, 165, 10))
              End If
            End If
          End If
        Next
        Tbknev.List1.AddItem "Bizonylatok száma:" + ertszam(Str(szla + szlev + rkozi), 10, 0)
        Tbknev.List1.AddItem "           Számla:" + ertszam(Str(szla), 10, 0)
        Tbknev.List1.AddItem "    Szállítólevél:" + ertszam(Str(szlev), 10, 0)
        Tbknev.List1.AddItem "       Raktárközi:" + ertszam(Str(rkozi), 10, 0)
        Tbknev.List1.AddItem "      Nettó érték:" + ertszam(Str(nte@), 10, 0)
        Tbknev.List1.AddItem "     Bruttó érték:" + ertszam(Str(bte@), 10, 0)
        Tbknev.List1.AddItem "  "
      End If
      If Trim(Mid$(rec$, 233, 8)) = "" Then
        Tbknev.List1.AddItem "Túra iktató: nincs jelölve"
        Tbknev.List1.AddItem Trim(Mid$(rec$, 14, 10)) + " " + Trim(Mid(rec$, 24, 30))
      Else
        Tbknev.List1.AddItem "Túra iktató: " + Mid$(rec$, 233, 7)
        komikt$ = Mid$(rec$, 233, 7)
        krec$ = dbxkey("KOMS", komikt$)
        If krec$ <> "" Then
          Tbknev.List1.AddItem Trim(Mid$(krec$, 14, 10)) + " " + Trim(Mid(krec$, 24, 30))
          If Mid$(krec$, 195, 1) = "S" Then
            Tbknev.List1.AddItem "T ö r ö l t túra"
            Tbknev.List1.ForeColor = QBColor(8)
            sat% = 0
          Else
            If Trim(Mid$(krec$, 162, 8)) = "" Then
              Tbknev.List1.AddItem "Túra nincs elindítva"
              Tbknev.List1.ForeColor = QBColor(0)
              sat% = 1
            Else
              If Trim(Mid$(krec$, 180, 8)) = "" Then
                Tbknev.List1.AddItem "Elindított túra"
                Tbknev.List1.ForeColor = QBColor(9)
                sat% = 2
              Else
                Tbknev.List1.AddItem "Elszámolt túra"
                Tbknev.List1.ForeColor = QBColor(2)
                sat% = 3
              End If
            End If
          End If
          If sat% > 0 Then
            Tbknev.List1.AddItem "      "
            Tbknev.List1.AddItem "Rögzítette :" + Mid$(krec$, 154, 8)
            If sat% > 1 Then
              Tbknev.List1.AddItem "Indította  :" + Mid$(krec$, 162, 8) + " "
              Tbknev.List1.AddItem "           :" + datki(Mid$(krec$, 170, 6)) + "  " + Mid$(krec$, 176, 2) + ":" + Mid$(krec$, 178, 2)
              If sat% > 2 Then
                Tbknev.List1.AddItem "Elszámolt  :" + Mid$(krec$, 180, 8) + " "
                Tbknev.List1.AddItem "           :" + datki(Mid$(krec$, 188, 6))
              End If
            End If
          End If
        End If
      End If
    Case "KOMS"
      Tbknev.List1.Clear
      If Mid$(rec$, 195, 1) = "S" Then
        Tbknev.List1.AddItem "T ö r ö l t túra"
        Tbknev.List1.AddItem "*** Nem módosítható ***"
        Tbknev.List1.ForeColor = QBColor(8)
        sat% = 0
      Else
        If Trim(Mid$(rec$, 162, 8)) = "" Then
          Tbknev.List1.AddItem "N i n c s elindítva"
          Tbknev.List1.AddItem "*** Módosítható ***"
          Tbknev.List1.ForeColor = QBColor(0)
          sat% = 1
        Else
          If Trim(Mid$(rec$, 180, 8)) = "" Then
            Tbknev.List1.AddItem "E l i n d í t o t t túra"
            Tbknev.List1.AddItem "*** Nem módosítható ***"
            Tbknev.List1.ForeColor = QBColor(9)
            sat% = 2
          Else
            If Mid$(rec$, 194, 1) = "V" Then
              Tbknev.List1.AddItem "E l s z á m o l t túra"
              Tbknev.List1.AddItem "*** Intézkedés szükséges ***"
              Tbknev.List1.AddItem "*** Nem módosítható ***"
              Tbknev.List1.ForeColor = QBColor(12)
            Else
              Tbknev.List1.AddItem "E l s z á m o l t túra"
              Tbknev.List1.AddItem "*** Rendben ***"
              Tbknev.List1.AddItem "*** Nem módosítható ***"
              Tbknev.List1.ForeColor = QBColor(2)
            End If
            sat% = 3
          End If
        End If
      End If
      If sat% > 0 Then
        Tbknev.List1.AddItem "      "
        Tbknev.List1.AddItem "Rögzítette :" + Mid$(rec$, 154, 8)
        If sat% > 1 Then
          Tbknev.List1.AddItem "Indította  :" + Mid$(rec$, 162, 8) + " "
          Tbknev.List1.AddItem "           :" + datki(Mid$(rec$, 170, 6)) + "  " + Mid$(rec$, 176, 2) + ":" + Mid$(rec$, 178, 2)
          If sat% > 2 Then
            Tbknev.List1.AddItem "Elszámolt  :" + Mid$(rec$, 180, 8) + " "
            Tbknev.List1.AddItem "           :" + datki(Mid$(rec$, 188, 6))
          End If
        End If
      End If
    Case "KAPC"
      Tbknev.List1.Clear
      kod$ = Mid$(rec$, 8, 1)
      qpart$ = Mid$(rec$, 9, 15)
      qszc$ = Mid$(rec$, 24, 15)
      qkod$ = Mid$(rec$, 39, 15)
      Select Case kod$
        Case "T"
          If Trim(qpart$) <> "" Then
            Tbknev.List1.AddItem "Tejdepo partner kód"
          End If
          If Trim(qszc$) <> "" Then
            Tbknev.List1.AddItem "Tejdepo száll.cím"
          End If
          If Trim(qkod$) <> "" Then
            Tbknev.List1.AddItem "Tejdepo termék kód"
          End If
        Case "B"
          If Trim(qpart$) <> "" Then
            Tbknev.List1.AddItem "Bónusz partner kód"
          End If
        Case "P"
          If Trim(qszc$) <> "" Then
            Tbknev.List1.AddItem "Pepsi vevõkód"
          End If
          If Trim(qkod$) <> "" Then
            Tbknev.List1.AddItem "Pepsi termék kód"
          End If
        Case Else
      End Select
      If Trim(qpart$) <> "" Then
        qrec$ = dbxkey("PART", qpart$)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem " "
          Tbknev.List1.AddItem "Partner:"
          Tbknev.List1.AddItem Trim(Mid$(qrec$, 16, 60))
        End If
      End If
      If Trim(qszc$) <> "" Then
        qrec$ = dbxkey("KSZC", qszc$)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem " "
          Tbknev.List1.AddItem "Száll.cím:"
          Tbknev.List1.AddItem Trim(Mid$(qrec$, 31, 60))
        End If
      End If
      If Trim(qkod$) <> "" Then
        qrec$ = dbxkey("KTRM", qkod$)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem " "
          Tbknev.List1.AddItem "Termék:"
          Tbknev.List1.AddItem Trim(Mid$(qrec$, 16, 60))
        End If
      End If
      If Mid$(rec$, 300, 1) = "S" Then
        Tbknev.List1.AddItem " "
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t kapcsolat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
    Case "KLEL"
      Tbknev.List1.Clear
      qkod$ = Mid$(rec$, 14, 4)
      qrec$ = dbxkey("KRAK", qkod$)
      If qrec$ <> "" Then
        Tbknev.List1.AddItem Mid$(qrec$, 5, 40)
      End If
      Tbknev.List1.AddItem " "
      qkod$ = Mid$(rec$, 18, 15)
      qrec$ = dbxkey("KTRM", qkod$)
      If qrec$ <> "" Then
        Tbknev.List1.AddItem Mid$(qrec$, 16, 60)
      End If
      Tbknev.List1.AddItem "Mintaszám : " + Mid$(rec$, 33, 8)
      Tbknev.List1.AddItem "Tárhely   : " + Mid$(rec$, 41, 8)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Fordulónap: " + datki(Mid$(rec$, 8, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Mennyiség : " + ertszam(Mid$(rec$, 49, 12), 12, 2)
      Tbknev.List1.AddItem "Egységár  : " + ertszam(Mid$(rec$, 85, 12), 12, 2)
      If Mid$(rec$, 111, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
        Tbknev.List1.AddItem "Sztornózva:" + datki(Mid$(rec$, 112, 6)) + " " + Mid$(rec$, 118, 8) + " "
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
    Case "JMEG"
      Tbknev.List1.Clear
      qpart$ = Mid$(rec$, 8, 15)
      qrec$ = dbxkey("PART", qpart$)
      If qrec$ <> "" Then
        Tbknev.List1.AddItem Mid$(qrec$, 16, 60)
      End If
      qjtrm$ = Mid$(rec$, 235, 15)
      qrec$ = dbxkey("JTRM", qjtrm$)
      If qrec$ <> "" Then
        Tbknev.List1.AddItem Mid$(qrec$, 90, 40)
      End If
      Tbknev.List1.AddItem "Kelte      : " + datki(Mid$(rec$, 43, 6))
      Tbknev.List1.AddItem "Határidõ   : " + datki(Mid$(rec$, 49, 6))
      Tbknev.List1.AddItem "Megrendelt : " + ertszam(Mid$(rec$, 138, 14), 14, 2)
      Tbknev.List1.AddItem "Teljesített: " + ertszam(Mid$(rec$, 152, 14), 14, 2)
      If Mid$(rec$, 226, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
    
    Case "JSZB"
      Tbknev.List1.Clear
      Select Case Mid$(rec$, 11, 2)
        Case "SL": Tbknev.List1.AddItem "Szállítólevél"
        Case "BS"
          If Trim(Mid$(rec$, 94, 3)) = "" Then
            Tbknev.List1.AddItem "Belföldi forint számla"
          Else
            Tbknev.List1.AddItem "Belföldi devizás számla"
          End If
        Case "BP": Tbknev.List1.AddItem "Belföldi proforma számla"
        Case "EU": Tbknev.List1.AddItem "Közösségi számla"
        Case "EP": Tbknev.List1.AddItem "Közösségi proforma számla"
        Case "XS": Tbknev.List1.AddItem "Export számla"
        Case "SP": Tbknev.List1.AddItem "Export proforma számla"
      End Select
      szc$ = Mid$(rec$, 235, 15)
      qpart$ = Mid$(rec$, 13, 15)
      If Trim(szc$) <> "" Then
        qrec$ = dbxkey("KSZC", szc$)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem Mid$(qrec$, 31, 60)
        End If
      Else
        qrec$ = dbxkey("PART", qpart)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem Mid$(qrec$, 16, 60)
        End If
      End If
      If Mid$(rec$, 187, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
        Tbknev.List1.AddItem "Sztornószámla: " + Mid$(rec$, 194, 10)
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Telj.kelte  : " + datki(Mid$(rec$, 74, 6))
      Tbknev.List1.AddItem "Számla kelte: " + datki(Mid$(rec$, 78, 6))
      Tbknev.List1.AddItem "Fiz.határidõ: " + datki(Mid$(rec$, 86, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Tételszám   : " + Str(xval(Mid$(rec$, 204, 3))) + " db"
      If Trim(Mid$(rec$, 94, 3)) <> "" Then
        Tbknev.List1.AddItem Mid$(rec$, 94, 3) + " árf.:" + Mid$(rec$, 97, 10)
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Kis.okmány : " + Mid$(rec$, 217, 12)
      Tbknev.List1.AddItem "Megrendelés: " + Mid$(rec$, 250, 15)
    Case "JFRG"
      Tbknev.List1.Clear
      muvk$ = Mid$(rec$, 377, 3)
      muvrec$ = dbxkey("JMUV", muvk$)
      If muvrec$ <> "" Then
        Tbknev.List1.AddItem Mid$(muvrec$, 9, 50)
      End If
      Select Case Mid$(rec$, 36, 1)
        Case "V": Tbknev.List1.AddItem "Vásárlás"
        Case "F": Tbknev.List1.AddItem "Felvásárlás"
        Case "E": Tbknev.List1.AddItem "Eladás"
        Case "G": Tbknev.List1.AddItem "Gyártás"
        Case "B": Tbknev.List1.AddItem "Belsõ mozgás"
        Case "H": Tbknev.List1.AddItem "Hiány"
        Case "T": Tbknev.List1.AddItem "Többlet"
        Case "K": Tbknev.List1.AddItem "Korrekció"
        Case "Y": Tbknev.List1.AddItem "Veszteség"
        Case "S": Tbknev.List1.AddItem "Szüret"
        Case "N": Tbknev.List1.AddItem "Nyitó készlet"
      End Select
      Tbknev.List1.AddItem " "
      Select Case Mid$(rec$, 35, 1)
        Case "B"
          Tbknev.List1.AddItem "Bor"
          tkod$ = Mid$(rec$, 40, 15)
          termrec$ = dbxkey("JTRM", tkod$)
          Tbknev.List1.AddItem Mid$(termrec$, 90, 40)
          Tbknev.List1.AddItem "Mennyiség:" + Mid$(rec$, 134, 14) + " L"
        Case "M"
          Tbknev.List1.AddItem "Melléktermék"
          tkod$ = Mid$(rec$, 81, 4)
          termrec$ = dbxkey("JMTR", tkod$)
          Tbknev.List1.AddItem Mid$(termrec$, 5, 40)
          Tbknev.List1.AddItem "Mennyiség:" + Mid$(rec$, 134, 14) + " " + Mid$(termrec$, 75, 6)
        Case "A"
          Tbknev.List1.AddItem "Anyag"
          tkod$ = Mid$(rec$, 66, 15)
          termrec$ = dbxkey("JANY", tkod$)
          Tbknev.List1.AddItem Mid$(termrec$, 16, 40)
          Tbknev.List1.AddItem "Mennyiség:" + Mid$(rec$, 134, 14) + " " + Mid$(termrec$, 86, 6)
        Case "S"
          Tbknev.List1.AddItem "Szõlõ"
          tkod$ = Mid$(rec$, 62, 4)
          termrec$ = dbxkey("JSZO", tkod$)
          Tbknev.List1.AddItem Mid$(termrec$, 5, 30)
          sm@ = xval(Mid$(rec$, 294, 12)) - xval(Mid$(rec$, 306, 12))
          Tbknev.List1.AddItem "Mennyiség:" + ertszam(Str(sm@), 14, 2) + " Kg"
        Case "G": Tbknev.List1.AddItem "Göngyöleg"
        Case "E"
          Tbknev.List1.AddItem "Egyéb jövedéki termék"
          tkod$ = Mid$(rec$, 40, 15)
          termrec$ = dbxkey("JTRM", tkod$)
          Tbknev.List1.AddItem Mid$(termrec$, 90, 40)
          Tbknev.List1.AddItem "Mennyiség:" + Mid$(rec$, 134, 14) + " " + Mid$(termrec$, 200, 6)
        Case Else
      End Select
      Select Case Mid$(rec$, 85, 1)
        Case "T"
          Tbknev.List1.AddItem "Tartály  :" + Mid$(rec$, 86, 20)
        Case "H"
          gkod$ = Mid$(rec$, 120, 4)
          grec$ = dbxkey("JGON", gkod$)
          Tbknev.List1.AddItem "Kanna    : " + Trim(Mid$(grec$, 38, 8)) + " L " + Trim(Mid$(rec$, 124, 10)) + " Db"
        Case "P"
          gkod$ = Mid$(rec$, 120, 4)
          grec$ = dbxkey("JGON", gkod$)
          Tbknev.List1.AddItem "Palack   : " + Trim(Mid$(grec$, 38, 8)) + " L " + Trim(Mid$(rec$, 124, 10)) + " Db"
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Belsõ bizonylat: " + Mid$(rec$, 370, 7)
      Tbknev.List1.AddItem "Számla száma   : " + Mid$(rec$, 162, 13)
      Tbknev.List1.AddItem "Kisérõ okmány  : " + Mid$(rec$, 205, 12)
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 217, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
        Tbknev.List1.AddItem "Sztornó kelte: " + datki(Mid$(rec$, 218, 6))
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
    Case "PSZB"
      Tbknev.List1.Clear
      Select Case Mid$(rec$, 76, 2)
        Case "SL": Tbknev.List1.AddItem "Szállítólevél"
        Case "BS"
          If Trim(Mid$(rec$, 98, 3)) = "" Then
            Tbknev.List1.AddItem "Belföldi forint számla"
          Else
            Tbknev.List1.AddItem "Belföldi devizás számla"
          End If
        Case "BP": Tbknev.List1.AddItem "Belföldi proforma számla"
        Case "EU": Tbknev.List1.AddItem "Közösségi számla"
        Case "EP": Tbknev.List1.AddItem "Közösségi proforma számla"
        Case "XS": Tbknev.List1.AddItem "Export számla"
        Case "SP": Tbknev.List1.AddItem "Export proforma számla"
      End Select
      qpart$ = Mid$(rec$, 61, 15)
      qrec$ = dbxkey("PART", qpart)
      If qrec$ <> "" Then
        Tbknev.List1.AddItem Mid$(qrec$, 16, 60)
      End If
      If Mid$(rec$, 35, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
        Tbknev.List1.AddItem "Sztornószámla: " + Mid$(rec$, 237, 15)
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Telj.kelte  : " + datki(Mid$(rec$, 84, 6))
      Tbknev.List1.AddItem "Számla kelte: " + datki(Mid$(rec$, 78, 6))
      Tbknev.List1.AddItem "Fiz.határidõ: " + datki(Mid$(rec$, 90, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Tételszám   : " + Str(xval(Mid$(rec$, 50, 3))) + " db"
      kod$ = Mid$(rec$, 60, 1)
      If Trim(Mid$(rec$, 98, 3)) <> "" Then
        Tbknev.List1.AddItem Mid$(rec$, 98, 3) + " árf.:" + Mid$(rec$, 101, 10)
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Rögzítette: " + Mid$(rec$, 21, 8) + " " + datki(Mid$(rec$, 29, 6))
      If Mid$(rec$, 35, 1) = "S" Then
        Tbknev.List1.AddItem "Sztornózta: " + Mid$(rec$, 42, 8) + " " + datki(Mid$(rec$, 36, 6))
      End If
    Case "KSZB"
      Tbknev.List1.Clear
      Select Case Mid$(rec$, 76, 2)
        Case "SL": Tbknev.List1.AddItem "Szállítólevél"
        Case "BS"
          If Trim(Mid$(rec$, 98, 3)) = "" Then
            Tbknev.List1.AddItem "Belföldi forint számla"
          Else
            Tbknev.List1.AddItem "Belföldi devizás számla"
          End If
        Case "BP": Tbknev.List1.AddItem "Belföldi proforma számla"
        Case "EU": Tbknev.List1.AddItem "Közösségi számla"
        Case "EP": Tbknev.List1.AddItem "Közösségi proforma számla"
        Case "XS": Tbknev.List1.AddItem "Export számla"
        Case "SP": Tbknev.List1.AddItem "Export proforma számla"
      End Select
      szc$ = Mid$(rec$, 260, 15)
      qpart$ = Mid$(rec$, 61, 15)
      qrec$ = dbxkey("PART", qpart)
      Tbknev.List1.AddItem Trim(Mid$(qrec$, 16, 60))
      If Trim(szc$) <> "" Then
        qrec$ = dbxkey("KSZC", szc$)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem Mid$(qrec$, 31, 60)
          kuzl$ = Mid$(qrec$, 200, 8)
          If Trim(kuzl$) <> "" Then
            kuzlrec$ = dbxkey("KUZL", kuzl$)
            If kuzlrec$ <> "" Then
              Tbknev.List1.AddItem "Üzl.:" + Mid$(kuzlrec$, 9, 60)
            End If
          End If
        End If
      Else
        qrec$ = dbxkey("PART", qpart)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem Mid$(qrec$, 16, 60)
          kuzl$ = Mid$(qrec$, 319, 8)
          If Trim(kuzl$) <> "" Then
            kuzlrec$ = dbxkey("KUZL", kuzl$)
            If kuzlrec$ <> "" Then
              Tbknev.List1.AddItem "Üzl.:" + Mid$(kuzlrec$, 9, 60)
            End If
          End If
        End If
      End If
      If Mid$(rec$, 35, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
        Tbknev.List1.AddItem "Sztornószámla: " + Mid$(rec$, 282, 15)
      Else
        Tbknev.List1.ForeColor = QBColor(3)
        If Mid$(rec$, 297, 1) = "I" Then
          Tbknev.List1.ForeColor = QBColor(9)
          Tbknev.List1.AddItem "K i a d a t l a n!"
        End If
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Telj.kelte  : " + datki(Mid$(rec$, 84, 6))
      Tbknev.List1.AddItem "Számla kelte: " + datki(Mid$(rec$, 78, 6))
      Tbknev.List1.AddItem "Fiz.határidõ: " + datki(Mid$(rec$, 90, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Tételszám   : " + Str(xval(Mid$(rec$, 50, 3))) + " db"
      kod$ = Mid$(rec$, 60, 1)
      If kod$ = "1" Or kod$ = "2" Or kod$ = "3" Or kod$ = "4" Then
        Tbknev.List1.AddItem "Bruttó érték: " + Mid$(rec$, 101, 10)
      Else
        If Trim(Mid$(rec$, 98, 3)) <> "" Then
          Tbknev.List1.AddItem Mid$(rec$, 98, 3) + " árf.:" + Mid$(rec$, 101, 10)
        End If
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Rögzítette: " + Mid$(rec$, 21, 8) + " " + datki(Mid$(rec$, 29, 6))
      If Mid$(rec$, 35, 1) = "S" Then
        Tbknev.List1.AddItem "Sztornózta: " + Mid$(rec$, 42, 8) + " " + datki(Mid$(rec$, 36, 6))
      End If
    Case "PMEG"
      Tbknev.List1.Clear
      If Mid$(rec$, 226, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem "Vevõ:"
      pkod$ = Mid$(rec$, 8, 15)
      If Trim(pkod) <> "" Then
        prec$ = dbxkey("PART", pkod$)
        Tbknev.List1.AddItem Trim(Mid$(prec$, 16, 60))
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Cím:"
        Tbknev.List1.AddItem Trim(postacim(prec$, 106))
      End If
      Tbknev.List1.AddItem "Hiv.szám:    " + Trim(Mid$(rec$, 23, 20))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Megr.kelte:  " + datki(Mid$(rec$, 43, 6))
      Tbknev.List1.AddItem "Sz.határidõ: " + datki(Mid$(rec$, 49, 6))
      Tbknev.List1.AddItem "Hiv.szám:    " + Trim(Mid$(rec$, 23, 20))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem Trim(Mid$(rec$, 61, 60))
      Tbknev.List1.AddItem "Egységár: " + Mid$(rec$, 121, 14) + " " + Mid$(rec$, 135, 3)
      Tbknev.List1.AddItem "Mennyiség:" + Mid$(rec$, 138, 14) + " " + Trim(Mid$(rec$, 226, 6))
      Tbknev.List1.AddItem " "
      m@ = xval(Mid$(rec$, 138, 14))
      t@ = xval(Mid$(rec$, 152, 14))
      If t@ >= m@ Then
        Tbknev.List1.AddItem "Teljesítés megtörtént!"
      Else
        If t@ < m@ And t@ > 0 Then
          Tbknev.List1.AddItem "Részteljesítés történt!"
        Else
          Tbknev.List1.AddItem "Nem volt teljesítés!"
        End If
      End If
    
    Case "KMEG"
      Tbknev.List1.Clear
      If Mid$(rec$, 406, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem "Vevõ:"
      Tbknev.List1.AddItem Trim(Mid$(rec$, 420, 60))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Szállítási cím:"
      Tbknev.List1.AddItem Trim(postacim(rec$, 88))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Megr.kelte:  " + datki(Mid$(rec$, 174, 6))
      Tbknev.List1.AddItem "Sz.határidõ: " + datki(Mid$(rec$, 180, 6))
      Tbknev.List1.AddItem "Ügyintézõ:   " + Trim(Mid$(rec$, 382, 8))
      Tbknev.List1.AddItem " "
      vidb% = 0: vitdb% = 0: viert@ = 0: vihdb% = 0
      For vi93% = 1 To 200
        elem$ = Mid$(rec$, (vi93% - 1) * 59 + 500, 59)
        tkod$ = Mid$(elem$, 1, 15)
        menny@ = xval(Mid$(elem$, 24, 12))
        If menny@ > 0 Then
          bear@ = xval(Mid$(elem$, 36, 12))
          telj@ = xval(Mid$(elem$, 48, 12))
          viert@ = viert@ + menny@ * bear@
          vidb% = vidb% + 1
          If telj@ > 0 Then vitdb% = vitdb% + 1
          If telj@ < menny@ Then vihdb% = vihdb% + 1
        End If
      Next
      Tbknev.List1.AddItem "Tételek: " + ertszam(Str(vidb%), 12, 0)
      Tbknev.List1.AddItem "N.érték: " + ertszam(Str(viert), 12, 2)
      Tbknev.List1.AddItem " "
      If vihdb = 0 Then
        Tbknev.List1.AddItem "Teljesítés megtörtént!"
      Else
        If vitdb > 0 Then
          Tbknev.List1.AddItem "Részteljesítés történt!"
        Else
          Tbknev.List1.AddItem "Nem volt teljesítés!"
        End If
      End If
    Case "KSMG"
      Tbknev.List1.Clear
      If Mid$(rec$, 6432, 1) = "O" Then
        Tbknev.List1.AddItem "5-ker részére"
      Else
        Tbknev.List1.AddItem "Merkatimpex részére"
      End If
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 398, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem "Szállító:"
      Tbknev.List1.AddItem Trim(Mid$(rec$, 23, 60))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Szállítási cím:"
      Tbknev.List1.AddItem Trim(Mid$(rec$, 420, 60))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Megr.kelte:  " + datki(Mid$(rec$, 192, 6))
      Tbknev.List1.AddItem "Sz.határidõ: " + datki(Mid$(rec$, 198, 6))
      Tbknev.List1.AddItem "Ügyintézõ:   " + Trim(Mid$(rec$, 6401, 30))
      Tbknev.List1.AddItem " "
      vidb% = 0: vitdb% = 0: viert@ = 0: vihdb% = 0
      For vi93% = 1 To 100
        elem$ = Mid$(rec$, (vi93% - 1) * 59 + 500, 59)
        tkod$ = Mid$(elem$, 1, 15)
        menny@ = xval(Mid$(elem$, 24, 12))
        If menny@ > 0 Then
          bear@ = xval(Mid$(elem$, 36, 12))
          telj@ = xval(Mid$(elem$, 48, 12))
          viert@ = viert@ + menny@ * bear@
          vidb% = vidb% + 1
          If telj@ > 0 Then vitdb% = vitdb% + 1
          If telj@ < menny@ Then vihdb% = vihdb% + 1
        End If
      Next
      Tbknev.List1.AddItem "Tételek: " + ertszam(Str(vidb%), 12, 0)
      Tbknev.List1.AddItem "N.érték: " + ertszam(Str(viert), 12, 2)
      Tbknev.List1.AddItem " "
      If vihdb = 0 Then
        Tbknev.List1.AddItem "Teljesítés megtörtént!"
      Else
        If vitdb > 0 Then
          Tbknev.List1.AddItem "Részteljesítés történt!"
        Else
          Tbknev.List1.AddItem "Nem volt teljesítés!"
        End If
      End If
    Case "KFRG"
      Tbknev.List1.Clear
      If Mid$(rec$, 101, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem "Hivatkozás: " + Mid$(rec$, 15, 15)
      Tbknev.List1.AddItem "Dátum:    : " + datki(Mid$(rec$, 30, 6))
      qkod$ = Mid$(rec$, 36, 3)
      If Trim(qkod$) <> "" Then
        qrec$ = dbxkey("KMOZ", qkod$)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem Trim(Mid$(qrec$, 4, 30))
        End If
      End If
      Tbknev.List1.AddItem " "
      qkod$ = Mid$(rec$, 53, 6)
      If Trim(qkod$) <> "" Then
        qrec$ = dbxkey("PTRM", qkod$)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem Trim(Mid$(qrec$, 7, 60))
        End If
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Besz.ár  :" + Mid$(rec$, 39, 14)
      Tbknev.List1.AddItem "Növekedés:" + Mid$(rec$, 59, 14)
      Tbknev.List1.AddItem "Csökkenés:" + Mid$(rec$, 73, 14)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Rögzítette: " + Mid$(rec$, 93, 8) + " " + datki(Mid$(rec$, 87, 6))
      If Mid$(rec$, 201, 1) = "S" Then
        Tbknev.List1.AddItem "Sztornózta: " + Mid$(rec$, 108, 8) + " " + datki(Mid$(rec$, 102, 6))
      End If
    
    Case "KKBZ"
      Tbknev.List1.Clear
      If Mid$(rec$, 92, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Select Case Mid$(rec$, 54, 1)
        Case "B": Tbknev.List1.AddItem "Beszerzés"
        Case "E": Tbknev.List1.AddItem "Értékesítés"
        Case "V": Tbknev.List1.AddItem "Visszáru szállítónak"
        Case "T": Tbknev.List1.AddItem "Termelés készletre vétel"
        Case "R": Tbknev.List1.AddItem "Raktárközi bizonylat"
        Case "F": Tbknev.List1.AddItem "Felhasználás"
        Case "S": Tbknev.List1.AddItem "Selejtezés"
        Case Else: Tbknev.List1.AddItem "Egyéb bizonylat"
      End Select
      qprt$ = Mid$(rec$, 55, 15)
      If Trim(qprt$) <> "" Then
        qpart$ = dbxkey("PART", qprt$)
        If qpart$ <> "" Then
          Tbknev.List1.AddItem Mid$(qpart$, 16, 60)
        End If
      End If
      Tbknev.List1.AddItem "Számlaszám  : " + Mid$(rec$, 107, 15)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Hivatkozás  : " + Mid$(rec$, 130, 10)
      Tbknev.List1.AddItem "Munkaszám   : " + Mid$(rec$, 70, 8)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Rögzítette: " + Mid$(rec$, 84, 8) + " " + datki(Mid$(rec$, 78, 6))
      If Mid$(rec$, 201, 1) = "S" Then
        Tbknev.List1.AddItem "Sztornózta: " + Mid$(rec$, 99, 8) + " " + datki(Mid$(rec$, 93, 6))
      End If
    Case "ERTB"
      Tbknev.List1.Clear
      szc$ = Mid$(rec$, 105, 15)
      If Trim(szc$) <> "" Then
        qrec$ = dbxkey("KSZC", szc$)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem Mid$(qrec$, 31, 60)
          kuzl$ = Mid$(qrec$, 200, 8)
          If Trim(kuzl$) <> "" Then
            kuzlrec$ = dbxkey("KUZL", kuzl$)
            If kuzlrec$ <> "" Then
              Tbknev.List1.AddItem "Üzl.:" + Mid$(kuzlrec$, 9, 60)
            End If
          End If
        Else
          Tbknev.List1.AddItem Mid$(rec$, 45, 60)
        End If
      Else
        Tbknev.List1.AddItem Mid$(rec$, 45, 60)
      End If
      If Mid$(rec$, 201, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n ó z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      If Mid$(rec$, 8, 2) = "M:" Then
        If Mid$(rec$, 201, 1) = "S" Then
          Tbknev.List1.ForeColor = QBColor(12)
        Else
          Tbknev.List1.ForeColor = RGB(0, 60, 120)
        End If
        Tbknev.List1.AddItem "M ó d o s í t ó bizonylat"
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Telj.kelte  : " + datki(Mid$(rec$, 39, 6))
      Tbknev.List1.AddItem "Száll.levél : " + Mid$(rec$, 155, 10)
      Tbknev.List1.AddItem "Számlaszám  : " + Mid$(rec$, 135, 10)
      Tbknev.List1.AddItem "Fiz.mód     : " + Mid$(rec$, 242, 2) + " --> " + datki(Mid$(rec$, 235, 6))
      'Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Tételszám   : " + Str(xval(Mid$(rec$, 189, 4))) + " db"
      Tbknev.List1.AddItem "Nettó érték : " + Mid$(rec$, 145, 10)
      Tbknev.List1.AddItem "Bruttó érték: " + Mid$(rec$, 165, 10)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Rögzítette: " + Mid$(rec$, 193, 8) + " " + datki(Mid$(rec$, 20, 6))
      If Mid$(rec$, 201, 1) = "S" Then
        Tbknev.List1.AddItem "Sztornózta: " + Mid$(rec$, 208, 8) + " " + datki(Mid$(rec$, 202, 6))
      End If
      If Len(Trim(Mid$(rec$, 216, 8))) = 7 Then
        Tbknev.List1.AddItem "Fuvar ikt.: " + Mid$(rec$, 216, 7)
        fuvrec$ = dbxkey("FUVA", Mid$(rec$, 216, 8))
        If fuvrec$ = "" Then
          Tbknev.List1.AddItem "Túra ikt. : nincs jelölve"
        Else
          If Trim(Mid$(fuvrec$, 233, 7)) = "" Then
            Tbknev.List1.AddItem "Túra ikt. : nincs jelölve"
          Else
            Tbknev.List1.AddItem "Túra ikt. : " + Mid$(fuvrec$, 233, 7)
          End If
        End If
      Else
        Tbknev.List1.AddItem "Fuvar ikt.: nincs jelölve"
      End If
      If Mid$(rec$, 155, 2) = "J:" Then
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Módosító számla: " + Mid$(rec$, 157, 7)
      Else
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Módosító számla: nincs"
      End If
    Case "FKTE"
      Tbknev.List1.Clear
      If Mid$(rec$, 61, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "Sztornózott tétel"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Select Case Mid$(rec$, 45, 4)
        Case "EIR ": Tbknev.List1.AddItem "Terv tétel"
        Case "TVE ": Tbknev.List1.AddItem "Vegyes könyvelési bizonylat"
        Case "TVNA": Tbknev.List1.AddItem "Automatikus nyitó tétel"
        Case "TVNK": Tbknev.List1.AddItem "Kézi nyitó tétel"
        Case "TPS ": Tbknev.List1.AddItem "Szállító számla"
        Case "TPSH": Tbknev.List1.AddItem "Szállító számla helyesbítõ"
        Case "TPV ": Tbknev.List1.AddItem "Vevõ számla"
        Case "TPVH": Tbknev.List1.AddItem "Vevõ számla helyesbítõ"
        Case "TPEV": Tbknev.List1.AddItem "Vevõ elõleg"
        Case "TPES": Tbknev.List1.AddItem "Szállító elõleg"
        Case "TPPP": Tbknev.List1.AddItem "Pénztári tétel"
        Case "TPPB": Tbknev.List1.AddItem "Banki tétel"
        Case "TPPK": Tbknev.List1.AddItem "Pénzügyi korrekciós tétel"
        Case "TFEA": Tbknev.List1.AddItem "Eszköz állományváltozás"
        Case "TFEC": Tbknev.List1.AddItem "Eszköz értékecsökkebés"
        Case "TFKE": Tbknev.List1.AddItem "Egyszerûsített készlet feladás"
        Case "TFKK": Tbknev.List1.AddItem "Készlet feladás"
        Case "TFB ": Tbknev.List1.AddItem "Munkabér feladás"
        Case "TFX ": Tbknev.List1.AddItem "Idegen szoftverbõl importált"
        Case Else
      End Select
      If Trim(Mid$(rec$, 192, 60)) <> "" Then Tbknev.List1.AddItem Mid$(rec$, 192, 60)
      If Trim(Mid$(rec$, 252, 25)) <> "" Then Tbknev.List1.AddItem Mid$(rec$, 252, 25)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Könyvelés kelte: " + datki(Mid$(rec$, 49, 6))
      Tbknev.List1.AddItem "Sztornó kelte:   " + datki(Mid$(rec$, 62, 6))
      Tbknev.List1.AddItem " "
      If Trim(Mid$(rec$, 108, 8)) <> "" Or Trim(Mid$(rec$, 124, 8)) <> "" Then
        Tbknev.List1.AddItem "T-" + Trim(Mid$(rec$, 100, 8)) + "(" + Trim(Mid$(rec$, 108, 8)) + ")" + "/K-" + Trim(Mid$(rec$, 116, 8)) + "(" + Trim(Mid$(rec$, 124, 8)) + ")"
      Else
        Tbknev.List1.AddItem "T-" + Trim(Mid$(rec$, 100, 8)) + "/K-" + Trim(Mid$(rec$, 116, 8))
      End If
      s$ = ertszam(Mid$(rec$, 132, 14), 14, 2)
      If Trim(Mid$(rec$, 146, 3)) <> "" Then
        s$ = s$ + "=" + Trim(ertszam(Mid$(rec$, 149, 14), 14, 2)) + " " + Mid$(rec$, 146, 3)
      End If
      Tbknev.List1.AddItem s$
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Szerv.egység:   " + Mid$(rec$, 84, 8)
      Tbknev.List1.AddItem "Munkaszám:      " + Mid$(rec$, 92, 8)
      Tbknev.List1.AddItem "Partner kód:    " + Mid$(rec$, 163, 15)
      If Trim(Mid$(rec$, 163, 15)) <> "" Then
        qpart$ = Mid$(rec$, 163, 15)
        qrec$ = dbxkey("PART", qpart$)
        If qrec$ <> "" Then Tbknev.List1.AddItem Mid$(qrec$, 16, 60)
      End If
      Tbknev.List1.AddItem "Számla iktató:  " + Mid$(rec$, 178, 7)
      Tbknev.List1.AddItem "P.forg.iktató:  " + Mid$(rec$, 185, 7)
    Case "KRAK"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      Tbknev.List1.AddItem Mid$(rec$, 68, 60)
      Select Case Mid$(rec$, 65, 1)
        Case "S": Tbknev.List1.AddItem "Saját központi raktár"
        Case "L": Tbknev.List1.AddItem "Lerakat"
        Case "M": Tbknev.List1.AddItem "Mozgó raktár"
        Case "K": Tbknev.List1.AddItem "Kiskereskedelmi egység"
        Case "J": Tbknev.List1.AddItem "Jövedéki raktár"
        Case "B": Tbknev.List1.AddItem "Bizományos"
        Case "Z": Tbknev.List1.AddItem "Beszállítói raktár"
        Case "I": Tbknev.List1.AddItem "ISO elõminõsítõ raktár"
      End Select
      pkod$ = Mid$(rec$, 128, 15)
      qpart$ = dbxkey("PART", pkod$)
      If qpart$ <> "" Then
        Tbknev.List1.AddItem "Partner: " + qpart$
        Tbknev.List1.AddItem Mid$(qpart, 16, 60)
        Tbknev.List1.AddItem postacim(qpart, 106)
      End If
      If Mid$(rec$, 66, 1) = "I" Then Tbknev.List1.AddItem "Tárhelyes raktár"
      If Mid$(rec$, 67, 1) = "I" Then Tbknev.List1.AddItem "Mintaszámos raktár "
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Készl.számla: " + Mid$(rec$, 143, 8)
      Tbknev.List1.AddItem "Árbev.számla: " + Mid$(rec$, 151, 8)
      Tbknev.List1.AddItem "ELÁBÉ.számla: " + Mid$(rec$, 185, 8)
      Tbknev.List1.AddItem "Szerv.egység: " + Mid$(rec$, 159, 8)
      Tbknev.List1.AddItem "Munkaszám:    " + Mid$(rec$, 167, 8)
    Case "KMOX"
      Tbknev.List1.Clear
      If Mid$(rec$, 34, 1) = "B" Then
        Tbknev.List1.ForeColor = QBColor(3)
        Tbknev.List1.AddItem "Készlet növelõ (bevét)"
      Else
        Tbknev.List1.ForeColor = QBColor(4)
        Tbknev.List1.AddItem "Készlet csökkentõ (kiadás)"
      End If
      Select Case Mid$(rec$, 43, 2)
        Case "BS": Tbknev.List1.AddItem "Beszerzés"
        Case "TM": Tbknev.List1.AddItem "Termelés"
        Case "TB": Tbknev.List1.AddItem "Leltár többlet"
        Case "BE": Tbknev.List1.AddItem "Egyéb bevét"
        Case "ER": Tbknev.List1.AddItem "Értékesítés"
        Case "VI": Tbknev.List1.AddItem "Visszáru"
        Case "SE": Tbknev.List1.AddItem "Selejtezés"
        Case "FH": Tbknev.List1.AddItem "Saját felhasználás"
        Case "HI": Tbknev.List1.AddItem "Leltár hiány"
        Case "KE": Tbknev.List1.AddItem "Egyéb kiadás"
        Case "RK": Tbknev.List1.AddItem "Raktárközi mozgás"
      End Select
      Tbknev.List1.AddItem "Kapcsolt mozgás: " + Mid$(rec$, 47, 3)
      If Mid$(rec$, 45, 1) = "I" Then
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "A mozgáshoz munkaszám kötelezõ"
      End If
      If Mid$(rec$, 46, 1) = "K" Then
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "A mozgáshoz besz.ár kitöltése:"
        Tbknev.List1.AddItem "   k ö t e l e z õ! "
        Tbknev.List1.AddItem "Az átlagár módosul."
      Else
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "A mozgáshoz besz.ár kitöltése:"
        Tbknev.List1.AddItem "   t i l o s! "
        Tbknev.List1.AddItem "Az átlagár nem módosul."
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Ellenszámla :" + Mid$(rec$, 35, 8)
      Tbknev.List1.AddItem "Költséghely :" + Mid$(rec$, 67, 8)
      Tbknev.List1.AddItem "Szerv.egység:" + Mid$(rec$, 51, 8)
      Tbknev.List1.AddItem "Munkaszám:   " + Mid$(rec$, 59, 8)
    Case "PTRM"
      Tbknev.List1.Clear
      kk@ = xval(Mid$(rec$, 315, 14))
      If kk@ > 0 Then
        Tbknev.List1.ForeColor = QBColor(3)
      Else
        If kk@ = 0 Then
          Tbknev.List1.ForeColor = QBColor(4)
        Else
          Tbknev.List1.ForeColor = QBColor(12)
        End If
      End If
      Select Case Mid$(rec$, 127, 1)
        Case "A": Tbknev.List1.AddItem "Anyag, áru"
        Case "S": Tbknev.List1.AddItem "Szolgáltatás"
        Case "F": Tbknev.List1.AddItem "Folyamatos szolgáltatása"
      End Select
      Select Case Mid$(rec$, 187, 1)
        Case "I": Tbknev.List1.AddItem "Készlet kezeléssel"
        Case Else: Tbknev.List1.AddItem "Készlet kezelés nélkül"
      End Select
      Tbknev.List1.AddItem "M.egység: " + Mid$(rec$, 140, 6)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Egységár:  " + Mid$(rec$, 146, 12) + " " + Mid$(rec$, 158, 3)
      Tbknev.List1.AddItem "Nyilv.ár:" + Mid$(rec$, 301, 14) + " " + Mid$(rec$, 158, 3)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Készlet: " + Mid$(rec$, 315, 14) + " " + Mid$(rec$, 158, 3)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Készl.számla:" + Mid$(rec$, 330, 8)
      Tbknev.List1.AddItem "Árbev.számla:" + Mid$(rec$, 163, 8)
      Tbknev.List1.AddItem "Szerv.egység:" + Mid$(rec$, 171, 8)
      Tbknev.List1.AddItem "Munkaszám:   " + Mid$(rec$, 179, 8)
    Case "ESZK"
      Tbknev.List1.Clear
      If Trim(Mid$(rec$, 736, 6)) = "" Then
        Tbknev.List1.ForeColor = QBColor(3)
      Else
        Tbknev.List1.ForeColor = QBColor(4)
      End If
      enrm$ = Mid$(rec$, 606, 8)
      enrmrec$ = dbxkey("ENRM", enrm$)
      If enrmrec$ <> "" Then
        Tbknev.List1.AddItem "Norma kód:" + Trim(Mid$(rec$, 606, 8))
        Tbknev.List1.AddItem Mid$(enrmrec$, 9, 30)
      End If
      Tbknev.List1.AddItem "ÉCS kulcs:" + Mid$(rec$, 615, 6) + " %"
      Tbknev.List1.AddItem "Tárolás:  " + Mid$(rec$, 542, 40)
      Tbknev.List1.AddItem "Létesítés:" + datki(Mid$(rec$, 583, 6))
      Tbknev.List1.AddItem "Aktiválás:" + datki(Mid$(rec$, 730, 6))
      Tbknev.List1.AddItem "Kivezetés:" + datki(Mid$(rec$, 736, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Bruttó érték:    " + ertszamx(Mid$(rec$, 661, 14), 15, 0)
      Tbknev.List1.AddItem "Maradvány érték: " + ertszamx(Mid$(rec$, 675, 14), 15, 0)
      Tbknev.List1.AddItem "Fejl.tartalékból:" + ertszamx(Mid$(rec$, 750, 14), 15, 0)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Nyilv.számla:" + Mid$(rec$, 621, 8)
      Tbknev.List1.AddItem "ÉCS.számla:  " + Mid$(rec$, 629, 8)
      Tbknev.List1.AddItem "Költséghely: " + Mid$(rec$, 637, 8)
      Tbknev.List1.AddItem "Szerv.egység:" + Mid$(rec$, 645, 8)
      Tbknev.List1.AddItem "Munkaszám:   " + Mid$(rec$, 653, 8)
    Case "FKSZ"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      Select Case Mid$(rec$, 69, 1)
        Case "T": Tbknev.List1.AddItem "Technikai"
        Case "B": Tbknev.List1.AddItem "Bank"
        Case "P": Tbknev.List1.AddItem "Pénztár"
        Case "K": Tbknev.List1.AddItem "Költséghely, költségviselõ"
        Case "V": Tbknev.List1.AddItem "Nyilvántartási"
        Case "N": Tbknev.List1.AddItem "Normál"
      End Select
      Select Case Mid$(rec$, 70, 1)
        Case "E": Tbknev.List1.AddItem "Eszköz"
        Case "F": Tbknev.List1.AddItem "Forrás"
        Case "K": Tbknev.List1.AddItem "Költség, ráfordítás"
        Case "B": Tbknev.List1.AddItem "Bevétel, árbevétel"
        Case "N": Tbknev.List1.AddItem "Nyitómérleg"
        Case "V": Tbknev.List1.AddItem "Nyilvántartási"
        Case "Z": Tbknev.List1.AddItem "Zárómérleg"
        Case "R": Tbknev.List1.AddItem "Eredmény"
      End Select
      Select Case Mid$(rec$, 71, 1)
        Case "I": Tbknev.List1.AddItem "Könyvelés mindenhol"
        Case "V": Tbknev.List1.AddItem "Csak vegyes bizonylaton"
        Case "N": Tbknev.List1.AddItem "Vegyes bizonylaton tilos"
        Case "L": Tbknev.List1.AddItem "Könyvelés letiltva"
      End Select
      If Mid$(r$, 76, 1) = "*" Then Tbknev.List1.AddItem "Gyûjtõ számla" Else Tbknev.List1.AddItem "Analitikus számla"
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "          Tartozik     Követel"
      Tbknev.List1.AddItem "Nyitó:" + ertszam(Mid$(rec$, 290, 14), 12, 0) + ertszam(Mid$(rec$, 304, 14), 12, 0)
      Tbknev.List1.AddItem "Forg.:" + ertszam(Mid$(rec$, 318, 14), 12, 0) + ertszam(Mid$(rec$, 332, 14), 12, 0)
      zegya@ = xval(Mid$(rec$, 290, 14)) - xval(Mid$(rec$, 304, 14))
      zegya@ = zegya@ + xval(Mid$(rec$, 318, 14)) - xval(Mid$(rec$, 332, 14))
      If zegya@ > 0 Then tegy@ = zegya@: kegy@ = 0 Else kegy@ = Abs(zegya@): tegy@ = 0
      Tbknev.List1.AddItem "Záró :" + ertszam(Str(tegy@), 12, 0) + ertszam(Str(kegy@), 12, 0)
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 69, 1) = "B" Then
        Tbknev.List1.ForeColor = QBColor(4)
        Tbknev.List1.AddItem "Bank:" + Mid$(rec$, 389, 30)
        Tbknev.List1.AddItem "Szla:" + Mid$(rec$, 419, 24)
        Tbknev.List1.AddItem "IBAN:" + Mid$(rec$, 443, 28)
        Tbknev.List1.AddItem "Dev.:" + Mid$(rec$, 471, 3)
      End If
      If Mid$(rec$, 69, 1) = "P" Then
        Tbknev.List1.ForeColor = QBColor(4)
        Tbknev.List1.AddItem "Zár.:" + datki(Mid$(rec$, 443, 6))
        Tbknev.List1.AddItem "Dev.:" + Mid$(rec$, 471, 3)
      End If
    Case "AKCF"
      Tbknev.List1.Clear
      If Mid$(rec$, 38, 1) = "I" Then
        Tbknev.List1.AddItem "Szállítói akció"
        Tbknev.List1.AddItem "   " + datki(Mid$(rec$, 41, 6)) + "-" + datki(Mid$(rec$, 47, 6))
      End If
      If Mid$(rec$, 39, 1) = "I" Then
        Tbknev.List1.AddItem "Nagyker akció"
        Tbknev.List1.AddItem "   " + datki(Mid$(rec$, 53, 6)) + "-" + datki(Mid$(rec$, 59, 6))
      End If
      If Mid$(rec$, 40, 1) = "I" Then
        Tbknev.List1.AddItem "Fogyasztói akció"
        Tbknev.List1.AddItem "   " + datki(Mid$(rec$, 65, 6)) + "-" + datki(Mid$(rec$, 71, 6))
      End If
      gykod$ = Mid$(rec$, 77, 15)
      Tbknev.List1.AddItem " "
      If Trim(gykod$) <> "" Then
        prec$ = dbxkey("PART", gykod)
        Tbknev.List1.AddItem "Gyártó:"
        Tbknev.List1.AddItem (Mid$(prec$, 16, 60))
        Tbknev.List1.AddItem " "
      End If
      szkod$ = Mid$(rec$, 92, 15)
      If Trim(szkod$) <> "" Then
        prec$ = dbxkey("PART", szkod)
        Tbknev.List1.AddItem "Szállító:"
        Tbknev.List1.AddItem (Mid$(prec$, 16, 60))
      End If
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 141, 1) = "S" Then
        Tbknev.List1.AddItem "Sztornózott akció"
        Tbknev.List1.ForeColor = QBColor(12)
      Else
        Tbknev.List1.ForeColor = QBColor(0)
      End If
    Case "CRMG"
      Tbknev.List1.Clear
      If Mid$(rec$, 31, 1) = "V" Then szt$ = " vevõi megállapodás" Else szt$ = " szállítói megállapodás"
      If Mid$(rec$, 62, 1) = "I" Then
        Tbknev.List1.ForeColor = QBColor(3)
        Tbknev.List1.AddItem "Élõ" + szt$
      Else
        Tbknev.List1.ForeColor = QBColor(4)
        Tbknev.List1.AddItem "Passzív" + szt$
      End If
      Tbknev.List1.AddItem " "
      cikkszamka$ = Mid$(rec$, 16, 15)
      qrec$ = dbxkey("KTRM", cikkszamka$)
      If qrec$ <> "" Then
        Tbknev.List1.AddItem Mid(qrec$, 16, 60)
      End If
      qpart$ = Mid$(rec$, 1, 15)
      qrec$ = dbxkey("PART", qpart$)
      If qrec$ <> "" Then
        Tbknev.List1.AddItem Mid(qrec$, 16, 60)
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Spec.ár:  " + Mid(rec$, 50, 12)
      Tbknev.List1.AddItem "Spec.engedmény: " + Mid(rec$, 44, 6) + " %"
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Hatályba lépés: " + datki(Mid$(rec$, 32, 6))
      Tbknev.List1.AddItem "Lejárat:        " + datki(Mid$(rec$, 32, 6))
    Case "ARMG"
      Tbknev.List1.Clear
      If Mid$(rec$, 46, 1) = "A" Then
        Tbknev.List1.ForeColor = QBColor(3)
        Tbknev.List1.AddItem "Aktív"
      Else
        Tbknev.List1.ForeColor = QBColor(4)
        Tbknev.List1.AddItem "Passzív"
      End If
      Tbknev.List1.AddItem " "
      cikkszamka$ = Mid$(rec$, 1, 15)
      qrec$ = dbxkey("KTRM", cikkszamka$)
      If qrec$ <> "" Then
        Tbknev.List1.AddItem Mid(qrec$, 16, 60)
      End If
      qpart$ = Mid$(rec$, 16, 15)
      qrec$ = dbxkey("PART", qpart$)
      If qrec$ <> "" Then
        Tbknev.List1.AddItem Mid(qrec$, 16, 60)
      End If
      Tbknev.List1.AddItem " "
      For i711% = 1 To 20
        elem$ = Mid$(rec$, (i711% - 1) * 40 + 200, 40)
        If dtm(Mid$(elem$, 1, 6)) <= dtm(maidatum) And dtm(Mid$(elem$, 7, 6)) >= dtm(maidatum) Then
          Tbknev.List1.AddItem "Spec.ár:  " + Mid(elem$, 15, 12)
          Tbknev.List1.AddItem "Spec.engedmény: " + Mid(elem$, 27, 6) + " %"
          Tbknev.List1.AddItem "Hatályba lépés: " + datki(Mid$(elem$, 1, 6))
          Tbknev.List1.AddItem "Lejárat:        " + datki(Mid$(elem$, 7, 6))
          Exit For
        End If
      Next
    Case "REAN"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      cikkszamka$ = Mid$(rec$, 14, 15)
      If Trim(cikkszamka$) <> "" Then
        qrec$ = dbxkey("KTRM", cikkszamka$)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem Mid(qrec$, 16, 60)
          Tbknev.List1.AddItem " "
          Tbknev.List1.AddItem "Fogyasztói ár:" + Mid$(qrec$, 680, 12)
          If Trim(Mid$(qrec$, 1097, 6)) <> "" Then
            Tbknev.List1.AddItem "Árváltozás kelte: " + datki(Mid$(qrec$, 1097, 6))
          End If
          If Mid$(qrec$, 1248, 1) = "I" Then
            Tbknev.List1.AddItem "Helyi ár"
          Else
            Tbknev.List1.AddItem "Központi ár"
          End If
          Tbknev.List1.AddItem " "
        End If
      End If
    Case "KSZC"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      Tbknev.List1.AddItem Trim(Mid$(rec$, 31, 60))
      Tbknev.List1.AddItem postacim(rec$, 121)
      Tbknev.List1.AddItem " "
      pkod$ = Mid$(rec$, 16, 15)
      partrec$ = torolvas("PART", pkod$, 0, 0)
      rkod$ = Mid$(rec$, 273, 4)
      rakrec$ = torolvas("KRAK", rkod$, 0, 0)
      If rakrec$ <> "" Then
        Tbknev.List1.AddItem "Belsõ :" + Trim(Mid$(rakrec$, 5, 60))
      Else
        Tbknev.List1.AddItem "Fizetõ:" + Trim(Mid$(partrec$, 15, 60))
      End If
      Select Case Mid$(rec$, 199, 1)
        Case "L": Tbknev.List1.AddItem "Biz: szállítólevél"
        Case "S": Tbknev.List1.AddItem "Biz: számla"
        Case "E": Tbknev.List1.AddItem "Biz: export számla"
        Case "R": Tbknev.List1.AddItem "Biz: raktárközi szállítólevél"
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      If partrec$ <> "" Then
        Tbknev.List1.AddItem "Adószám     :" + Mid$(partrec$, 184, 15)
      End If
      Tbknev.List1.AddItem "Eng.száma   :" + Mid$(rec$, 277, 1) + "-" + Trim(Mid$(rec$, 278, 20))
      If Trim(Mid$(rec$, 208, 4)) <> "" Then
        Tbknev.List1.AddItem "Régió     :" + Trim(torolvas("RGIO", Mid$(rec$, 208, 4), 5, 60))
      Else
        If partrec$ <> "" Then
          Tbknev.List1.AddItem "Régió     :" + Trim(torolvas("RGIO", Mid$(partrec$, 315, 4), 5, 60))
        End If
      End If
      If Trim(Mid$(rec$, 200, 8)) <> "" Then
        Tbknev.List1.AddItem "Üzletkötõ :" + Trim(torolvas("KUZL", Mid$(rec$, 200, 8), 8, 60))
      Else
        If partrec$ <> "" Then
          Tbknev.List1.AddItem "Üzletkötõ :" + Trim(torolvas("KUZL", Mid$(partrec$, 319, 8), 8, 60))
        End If
      End If
      If partrec$ <> "" Then
        Tbknev.List1.AddItem "Fizm.mód  :" + Trim(torolvas("PFIZ", Mid$(partrec$, 328, 2), 3, 30)) + " " + Mid$(partrec$, 330, 3) + " nap"
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Tartozás  :" + ertszamx(Mid$(partrec$, 659, 14), 16, 2)
        If Mid$(partrec$, 327, 1) = "T" Then
          Tbknev.List1.AddItem "Partner letiltva!!!"
          Tbknev.List1.ForeColor = QBColor(12)
        Else
          If Mid$(partrec$, 327, 1) = "K" Then
            Tbknev.List1.AddItem "Csak készpénre!!!"
            Tbknev.List1.ForeColor = QBColor(4)
          Else
          End If
        End If
      End If
    Case "PART", "JPAR"
      If Left(programnev$, 5) = "AUW-J" Then
        jpk$ = Mid$(rec$, 1, 15)
        prec$ = dbxkey("PART", jpk)
        prec$ = Left(prec$ + Space(900), 900)
        Mid$(prec$, 802, 21) = Mid$(rec$, 26, 21)
        bbrec$ = rec$: rec$ = prec$
      End If
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      Tbknev.List1.AddItem Trim(Mid$(rec$, 16, 60))
      Tbknev.List1.AddItem postacim(rec$, 106)
      Select Case Mid$(rec$, 700, 2)
        Case "BK": Tbknev.List1.AddItem "Belföldi kereskedelmi partner"
        Case "KE": Tbknev.List1.AddItem "Közösségi partner"
        Case "KG": Tbknev.List1.AddItem "Egyéb külföldi partner"
        Case "PH": Tbknev.List1.AddItem "Pénzügyi hatóság"
        Case "PG": Tbknev.List1.AddItem "Egyéb pénzügyi partner"
        Case Else
      End Select
      Tbknev.List1.AddItem "Adószám     :" + Mid$(rec$, 184, 15)
      Tbknev.List1.AddItem "Köz.adószám :" + Mid$(rec$, 199, 15)
      Tbknev.List1.AddItem "Eng.száma   :" + Mid$(rec$, 802, 1) + "-" + Trim(Mid$(rec$, 803, 20))
      Tbknev.List1.AddItem banktagol(Mid$(rec$, 244, 24))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Vevõ kat. :" + Trim(torolvas("VKAT", Mid$(rec$, 314, 1), 2, 60))
      Tbknev.List1.AddItem "Árkateg.  :" + Mid$(rec$, 782, 1)
      Tbknev.List1.AddItem "Régió     :" + Trim(torolvas("RGIO", Mid$(rec$, 315, 4), 5, 60))
      Tbknev.List1.AddItem "Üzletkötõ :" + Trim(torolvas("KUZL", Mid$(rec$, 319, 8), 9, 60))
      Tbknev.List1.AddItem "Fizm.mód  :" + Trim(torolvas("PFIZ", Mid$(rec$, 328, 2), 3, 30)) + " " + Mid$(rec$, 330, 3) + " nap"
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Tartozás  :" + ertszamx(Mid$(rec$, 659, 14), 16, 2)
      If Mid$(rec$, 327, 1) = "T" Then
        Tbknev.List1.AddItem "Letiltva!!!"
        Tbknev.List1.ForeColor = QBColor(12)
      Else
        If Mid$(rec$, 327, 1) = "K" Then
          Tbknev.List1.AddItem "Csak készpénzre!!!"
          Tbknev.List1.ForeColor = QBColor(4)
        Else
        End If
      End If
      If Left(programnev$, 5) = "AUW-J" Then
        rec$ = bbrec$
      End If
    Case "PELO"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      pkod$ = Mid$(rec$, 23, 15): If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod) Else partrec$ = ""
      If partrec$ <> "" Then Tbknev.List1.AddItem Trim(Mid$(partrec$, 16, 60))
      Tbknev.List1.AddItem " "
      Select Case Mid$(rec$, 225, 1)
        Case "V": Tbknev.List1.AddItem "Vevõtõl kapott elõleg"
        Case "S": Tbknev.List1.AddItem "Szállítónak adott elõleg"
      End Select
      Select Case Mid$(rec$, 224, 1)
        Case "E": Tbknev.List1.AddItem "Elõleg fizetés"
        Case "V": Tbknev.List1.AddItem "Elõleg visszafizetés"
        Case "B": Tbknev.List1.AddItem "Elõleg beszámítás"
      End Select
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Számlaszám:" + Mid$(rec$, 8, 15)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Teljesítés kelte  :" + datki(Mid$(rec$, 38, 6))
      Tbknev.List1.AddItem "Pénztári iktató   :" + datki(Mid$(rec$, 38, 6))
      Tbknev.List1.AddItem "Banki iktató      :" + datki(Mid$(rec$, 38, 6))
      Tbknev.List1.AddItem "Könyvelési iktató :" + datki(Mid$(rec$, 38, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Elõleg összege  :" + ertszamx(Mid$(rec$, 44, 14), 14, 2)
      Tbknev.List1.AddItem "Beszám.összeg   :" + ertszamx(Mid$(rec$, 134, 14), 14, 2)
      Tbknev.List1.AddItem "Visszafiz.összeg:" + ertszamx(Mid$(rec$, 148, 14), 14, 2)
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 90, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Sztornózva:" + datki(Mid$(rec$, 91, 6)) + " " + Mid$(rec$, 97, 8) + " "
      End If
    Case "PVSZ", "PSSZ"
      If objneve$ = "PVSZ" Then vsmod$ = "V" Else vsmod$ = "S"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      pkod$ = Mid$(rec$, 38, 15): If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod) Else partrec$ = ""
      If partrec$ <> "" Then Tbknev.List1.AddItem Trim(Mid$(partrec$, 16, 60))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Számlaszám:" + Mid$(rec$, 8, 15)
      Tbknev.List1.AddItem "Hiv.szám  :" + Mid$(rec$, 23, 15)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Teljesítés kelte :" + datki(Mid$(rec$, 211, 6))
      Tbknev.List1.AddItem "Könyvelés kelte  :" + datki(Mid$(rec$, 58, 6))
      Tbknev.List1.AddItem "Számla kelte     :" + datki(Mid$(rec$, 64, 6))
      Tbknev.List1.AddItem "Fizetési határidõ:" + datki(Mid$(rec$, 70, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Összeg  :" + ertszamx(Mid$(rec$, 78, 14), 14, 2)
      egyenleg@ = szegyen(rec$, maidatum$, maidatum$)
      If egyenleg@ <> 0 Then Call szamlaegyenleg(rec$, ossz@, helybit@, kiegy@, egyenleg@, vsmod$, maidatum, maidatum, forintegyenleg@)
      Tbknev.List1.AddItem "Egyenleg:" + ertszamx(Str(egyenleg), 14, 2)
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 166, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Sztornózva:" + datki(Mid$(rec$, 167, 6)) + " " + Mid$(rec$, 181, 8) + " "
        Tbknev.List1.AddItem "Szt.számla:" + Mid$(rec$, 220, 15) + " "
      End If
    Case "PBNK"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      If Mid$(rec$, 8, 1) = "T" Then vr$ = "Terhelés" Else vr$ = "Jóváírás"
      Tbknev.List1.AddItem vr$
      fksz$ = Mid$(rec$, 22, 8): fkszrec$ = dbxkey("FKSZ", fksz)
      Tbknev.List1.AddItem Mid$(rec$, 240, 30)
      Tbknev.List1.AddItem datki(Mid$(rec$, 16, 6)) + " " + Trim(Mid$(fkszrec$, 9, 60))
      Select Case Mid$(rec$, 55, 1)
        Case "V": Tbknev.List1.AddItem "Vevõ számla kiegy. :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "S": Tbknev.List1.AddItem "Száll.számla kiegy.:" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "E": Tbknev.List1.AddItem "Elõleg vevõtõl     :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "L": Tbknev.List1.AddItem "Elõleg szállítónak :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "M": Tbknev.List1.AddItem "Elszámolási elõleg :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "X": Tbknev.List1.AddItem "Egyéb              :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
      End Select
      Tbknev.List1.AddItem " "
      pkod$ = Mid$(rec$, 108, 15): If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod) Else partrec$ = ""
      If partrec$ <> "" Then Tbknev.List1.AddItem Trim(Mid$(partrec$, 16, 60))
      Tbknev.List1.AddItem Mid$(rec$, 30, 25)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Számlaszám  :" + Mid$(rec$, 123, 15)
      Tbknev.List1.AddItem "Ellenszámla :" + Trim(Mid$(rec$, 146, 8)) + " Kt.:" + Mid$(rec$, 207, 7)
      Tbknev.List1.AddItem "Szerv.egység:" + Mid$(rec$, 162, 8)
      Tbknev.List1.AddItem "Munkaszám   :" + Mid$(rec$, 170, 8)
      If Mid$(rec$, 192, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Sztornózva:" + datki(Mid$(rec$, 193, 6)) + " " + Mid$(rec$, 199, 8) + " "
      End If
    Case "PKTE"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      If Mid$(rec$, 8, 1) = "B" Then vr$ = Mid$(rec$, 9, 7) + ".sz. bevételi bizonylat" Else vr$ = Mid$(rec$, 9, 7) + ".sz. kiadási bizonylat"
      Tbknev.List1.AddItem vr$
      fksz$ = Mid$(rec$, 22, 8): fkszrec$ = dbxkey("FKSZ", fksz)
      Tbknev.List1.AddItem Mid$(rec$, 240, 30)
      Tbknev.List1.AddItem datki(Mid$(rec$, 16, 6)) + " " + Trim(Mid$(fkszrec$, 9, 60))
      Select Case Mid$(rec$, 55, 1)
        Case "V": Tbknev.List1.AddItem "Vevõ számla kiegy. :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "S": Tbknev.List1.AddItem "Száll.számla kiegy.:" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "E": Tbknev.List1.AddItem "Elõleg vevõtõl     :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "L": Tbknev.List1.AddItem "Elõleg szállítónak :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "M": Tbknev.List1.AddItem "Elszámolási elõleg :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "X": Tbknev.List1.AddItem "Egyéb              :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
      End Select
      Tbknev.List1.AddItem " "
      pkod$ = Mid$(rec$, 108, 15): If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod) Else partrec$ = ""
      If partrec$ <> "" Then Tbknev.List1.AddItem Trim(Mid$(partrec$, 16, 60))
      Tbknev.List1.AddItem Mid$(rec$, 30, 25)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Ellenszámla :" + Trim(Mid$(rec$, 146, 8)) + " Kt.:" + Mid$(rec$, 207, 7)
      Tbknev.List1.AddItem "Szerv.egység:" + Mid$(rec$, 162, 8)
      Tbknev.List1.AddItem "Munkaszám   :" + Mid$(rec$, 170, 8)
      If Mid$(rec$, 192, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Sztornózva:" + datki(Mid$(rec$, 193, 6)) + " " + Mid$(rec$, 199, 8) + " "
      End If
    Case "PKOR"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      Select Case Mid$(rec$, 14, 1)
        Case "V": Tbknev.List1.AddItem "Vevõ korrekció:" + ertszam(Mid$(rec$, 45, 14), 12, 2): Tbknev.List1.ForeColor = RGB(0, 150, 0)
        Case "S": Tbknev.List1.AddItem "Szállító korrekció:" + ertszam(Mid$(rec$, 45, 14), 12, 2): Tbknev.List1.ForeColor = RGB(0, 0, 150)
        Case "K": Tbknev.List1.AddItem "Kompenzáció:" + ertszam(Mid$(rec$, 45, 14), 12, 2): Tbknev.List1.ForeColor = RGB(50, 50, 50)
      End Select
      If Mid$(rec$, 15, 1) = "I" Then
        Tbknev.List1.AddItem "ÁFA alapot módosít."
      Else
        Tbknev.List1.AddItem "ÁFA alapot nem módosít."
      End If
      Tbknev.List1.AddItem " "
      pkod$ = Mid$(rec$, 16, 15): If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod) Else partrec$ = ""
      If partrec$ <> "" Then Tbknev.List1.AddItem Trim(Mid$(partrec$, 16, 60))
      Tbknev.List1.AddItem " "
      vikt$ = Trim(Mid$(rec$, 31, 7)): sikt$ = Trim(Mid$(rec$, 38, 7))
      tikt$ = Mid$(rec$, 356, 7)
      If vikt$ <> "" Then
        virec$ = dbxkey("PVSZ", vikt$)
        Tbknev.List1.AddItem "Vevõ számla :" + vikt$ + "  " + Trim(Mid$(virec$, 8, 15))
        Tbknev.List1.AddItem "Könyv.tétel :" + tikt$ + " "
      End If
      If sikt$ <> "" Then
        sirec$ = dbxkey("PSSZ", sikt$)
        Tbknev.List1.AddItem "Száll.számla:" + sikt$ + "  " + Trim(Mid$(sirec$, 8, 15))
        If Mid$(rec$, 14, 1) = "S" Then
          Tbknev.List1.AddItem "Könyv.tétel :" + tikt$ + " "
        Else
          tik1& = xval(tikt$)
          tikt1$ = Right("0000000" + Trim(Str(tik1&)), 7)
          Tbknev.List1.AddItem "Könyv.tétel :" + tikt1$ + " "
        End If
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem Mid$(rec$, 76, 30)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Ellenszámla :" + Trim(Mid$(rec$, 324, 8))
      Tbknev.List1.AddItem "Szerv.egység:" + Mid$(rec$, 340, 8)
      Tbknev.List1.AddItem "Munkaszám   :" + Mid$(rec$, 348, 8)
      If Mid$(rec$, 120, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Sztornózva:" + datki(Mid$(rec$, 121, 6)) + " " + Mid$(rec$, 127, 8) + " "
      End If
    Case "JTRM"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      Tbknev.List1.AddItem Mid$(rec$, 90, 40)
      Tbknev.List1.AddItem "Évj: " + Mid$(rec$, 76, 4) + " Fok: " + Trim(Mid$(rec$, 5960, 10))
      Select Case Mid$(rec$, 286, 1)
        Case "B": Tbknev.List1.AddItem "Szõlõbor"
        Case "A": Tbknev.List1.AddItem "Aszú"
        Case "S": Tbknev.List1.AddItem "Sör"
        Case "R": Tbknev.List1.AddItem "Egyéb bor"
        Case "E": Tbknev.List1.AddItem "Egyéb alkoholtermék"
        Case Else: Tbknev.List1.AddItem "Jelleg hiányzik"
      End Select
      Tbknev.List1.AddItem "Palackozott Vtsz: " + Mid$(rec$, 142, 12)
      Tbknev.List1.AddItem "Lédig Vtsz      : " + Mid$(rec$, 142, 12)
    Case "KTRM"
      If programnev$ = "AUW-RKER" Then
        rkind$ = "0001" + Mid$(rec$, 1, 15)
        rkszrec$ = dbxkey("RKSZ", rkind$)
        If rkszrec$ = "" Then
          keszle@ = 0
        Else
          keszle@ = xval(Mid$(rkszrec$, 20, 12))
        End If
      Else
        aktkeszle@ = xval(Mid$(rec$, 748, 14))
        kozkeszle@ = xval(Mid$(rec$, 927, 14))
        vegkeszle@ = xval(Mid$(rec$, 941, 14))
        cpckeszle@ = xval(Mid$(rec$, 955, 14))
        gonkeszle@ = xval(Mid$(rec$, 969, 14))
        komkeszle@ = xval(Mid$(rec$, 983, 14))
        fogkeszle@ = xval(Mid$(rec$, 762, 14))
        zolkeszle@ = xval(Mid$(rec$, 1001, 12))
        Select Case raktartipus
          Case 1
            '--- komissiózó
            keszle@ = komkeszle@ - fogkeszle@
          Case 2
            '--- c+c
            keszle@ = cpckeszle@ + vegkeszle@
          Case 3
            '--- vegyi
            keszle@ = vegkeszle@
          Case 4
            '--- zöldség
            keszle@ = zolkeszle@
          Case 5
            '--- göngyöleg
            keszle@ = gonkeszle@
          Case Else
            '--- nincs beállítva
            keszle@ = aktkeszle@ - fogkeszle@
        End Select
      End If
      Tbknev.List1.Clear
      Tbknev.List1.AddItem Mid(rec$, 16, 60)
      If xval(Mid$(rec$, 1226, 7)) <> 0 Then
        Tbknev.List1.AddItem Trim(Mid$(rec$, 1226, 7)) + " db/karton"
      Else
        Tbknev.List1.AddItem " "
      End If
      If keszle@ > 0 Then
        ve$ = Trim(Mid$(rec$, 1219, 7))
        vk$ = Trim(Mid$(rec$, 1226, 7))
        vs$ = Trim(Mid$(rec$, 1233, 7))
        vr$ = Trim(Mid$(rec$, 1240, 7))
        Tbknev.List1.AddItem "Készlet:" + ertszam(Str(keszle@), 14, 2) + " " + Mid$(rec$, 484, 6)
        Tbknev.List1.ForeColor = QBColor(3)
        Call kiszvalt(keszle@, egyscsom&, qkarton&, qsor&, raklap&, toredek@, torede@, toredk@, toreds@, rec$)
        If ve$ <> "" And egyscsom& <> 0 And egyscsom& <> keszle@ Then
          If toredek@ <> 0 Then xa$ = " +" + Trim(ertszam(Str(toredek@), 8, 2))
          Tbknev.List1.AddItem "Egys.csom.:" + Right(Space(12) + Str(egyscsom&), 8) + "(x" + ve$ + ")" + xa$
        End If
        If qkarton& <> 0 And qkarton& <> egyscsom& Then
          If torede@ <> 0 Then xa$ = " +" + Trim(ertszam(Str(torede@), 5, 0))
          Tbknev.List1.AddItem "Karton:   :" + Right(Space(12) + Str(qkarton&), 8) + "(x" + vk$ + ")" + xa$
        End If
        If qsor& <> 0 And qsor& <> qkarton& Then
          If toredk@ <> 0 Then xa$ = " +" + Trim(ertszam(Str(toredk@), 5, 0))
          Tbknev.List1.AddItem "Sor:      :" + Right(Space(12) + Str(qsor&), 8) + "(x" + vs$ + ")" + xa$
        End If
        If raklap& <> 0 And raklap& <> qsor& Then
          If toreds@ <> 0 Then xa$ = " +" + Trim(ertszam(Str(toreds@), 5, 0))
          Tbknev.List1.AddItem "Raklap    :" + Right(Space(12) + Str(raklap&), 8) + "(x" + vr$ + ")" + xa$
        End If
        Tbknev.List1.AddItem " "
      Else
        If keszle@ < 0 Then
          Tbknev.List1.AddItem "Készlet:" + ertszam(Str(keszle@), 14, 2) + " " + Mid$(rec$, 484, 6)
          Tbknev.List1.AddItem "N e g a t í v  k é s z l e t"
          Tbknev.List1.ForeColor = QBColor(12)
          Tbknev.List1.AddItem " "
        Else
          Tbknev.List1.AddItem "Készlet: nincs eladható készlet"
          Tbknev.List1.ForeColor = QBColor(4)
          Tbknev.List1.AddItem " "
        End If
      End If
      Tbknev.List1.AddItem "Ut.besz.ár:" + Mid$(rec$, 568, 12)
      Tbknev.List1.AddItem "Ref.ár    :" + Mid$(rec$, 1276, 12)
      Tbknev.List1.AddItem "Ár1:" + Mid$(rec$, 582, 12) + " Ár5:" + Mid$(rec$, 638, 12)
      Tbknev.List1.AddItem "Ár2:" + Mid$(rec$, 596, 12) + " Ár6:" + Mid$(rec$, 652, 12)
      Tbknev.List1.AddItem "Ár3:" + Mid$(rec$, 610, 12) + " Ár7:" + Mid$(rec$, 666, 12)
      Tbknev.List1.AddItem "Ár4:" + Mid$(rec$, 624, 12)
      Tbknev.List1.AddItem "Disztr.ár:" + Mid$(rec$, 895, 12)
      Tbknev.List1.AddItem "Kisker.ár:" + Mid$(rec$, 680, 12)
      Tbknev.List1.AddItem " "
      If kozkeszle <> 0 Then Tbknev.List1.AddItem "Közp.rakt:" + ertszam(Str(kozkeszle@), 12, 2) + " " + Mid$(rec$, 484, 6)
      If cpckeszle <> 0 Then Tbknev.List1.AddItem "C+C      :" + ertszam(Str(cpckeszle@), 12, 2) + " " + Mid$(rec$, 484, 6)
      If vegkeszle <> 0 Then Tbknev.List1.AddItem "Vegyi rak:" + ertszam(Str(vegkeszle@), 12, 2) + " " + Mid$(rec$, 484, 6)
      If gonkeszle <> 0 Then Tbknev.List1.AddItem "Göngy.rak:" + ertszam(Str(gonkeszle@), 12, 2) + " " + Mid$(rec$, 484, 6)
      If zolkeszle <> 0 Then Tbknev.List1.AddItem "Zölds.rak:" + ertszam(Str(zolkeszle@), 12, 2) + " " + Mid$(rec$, 484, 6)
    Case Else
      Call infomutat2(objneve$, rec$, hivashely$)
  End Select
End Sub

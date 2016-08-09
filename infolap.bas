Attribute VB_Name = "Infolap"
Public infolapbe%
Public Sub infomutat2(objneve$, rec$, hivashely$)
  '--- keres�t�bla inform�ci� felmutat�sa
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
        Case "A": Tbknev.List1.AddItem "Egyszer�s�tett ad�rakt�r"
        Case "J": Tbknev.List1.AddItem "J�ved�ki enged�lyes"
        Case "B": Tbknev.List1.AddItem "Bejegyzett fogad�hely"
        Case "M": Tbknev.List1.AddItem "M�k�d�si enged�lyes"
        Case Else
      End Select
      Tbknev.List1.AddItem "Eng.sz�m: " + Trim(Mid$(rec$, 126, 13))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Borkis�r� okm�ny:"
      Tbknev.List1.AddItem "   " + Trim(Mid$(rec$, 260, 2)) + Trim(Mid$(rec$, 262, 10)) + "-" + Trim(Mid$(rec$, 260, 2)) + Trim(Mid$(rec$, 272, 10))
      Tbknev.List1.AddItem "Egyszer�s�tett kis�r� okm�ny:"
      Tbknev.List1.AddItem "   " + Trim(Mid$(rec$, 292, 2)) + Trim(Mid$(rec$, 294, 10)) + "-" + Trim(Mid$(rec$, 292, 2)) + Trim(Mid$(rec$, 304, 10))
      Tbknev.List1.AddItem "Adminisztrat�v kis�r� okm�ny:"
      Tbknev.List1.AddItem "   " + Trim(Mid$(rec$, 324, 2)) + Trim(Mid$(rec$, 326, 10)) + "-" + Trim(Mid$(rec$, 324, 2)) + Trim(Mid$(rec$, 336, 10))
    Case "JGYR"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      trec$ = torolvas("JRAK", Mid$(rec$, 8, 4), 1, 400)
      Tbknev.List1.AddItem "Rakt�r: "
      Tbknev.List1.AddItem Trim(Mid$(trec$, 5, 60))
      trec$ = torolvas("JTRM", Mid$(rec$, 18, 15), 1, 200)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Term�k: "
      Tbknev.List1.AddItem Trim(Mid$(trec$, 90, 40))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Gy�rt�s kezdete: " + datki(Mid$(rec$, 12, 6))
      Tbknev.List1.AddItem "Sz�rm: " + Trim(Mid$(rec$, 33, 20))
      Tbknev.List1.AddItem "OBI  : " + Trim(Mid$(rec$, 53, 20))
      If Mid$(rec$, 160, 1) = "B" Then
        Tbknev.List1.AddItem "B � r m u n k a "
      End If
      If Mid$(rec$, 73, 1) = "S" Then
        Tbknev.List1.AddItem "S z t o r n � z v a!"
        Tbknev.List1.ForeColor = QBColor(12)
      End If
    Case "JTAR"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      trec$ = torolvas("JRAK", Mid$(rec$, 65, 4), 1, 400)
      Tbknev.List1.AddItem "Rakt�r: "
      Tbknev.List1.AddItem Trim(Mid$(trec$, 5, 60))
      Tbknev.List1.AddItem " "
      Select Case Mid$(rec$, 129, 1)
        Case "U"
          Tbknev.List1.AddItem "� r e s"
          Tbknev.List1.ForeColor = QBColor(2)
        Case "B"
          Tbknev.List1.AddItem "B o r"
          trec$ = torolvas("JTRM", Mid$(rec$, 130, 15), 1, 200)
          Tbknev.List1.AddItem Trim(Mid$(trec$, 90, 40))
        Case "M"
          Tbknev.List1.AddItem "M e l l � k t e r m � k"
          trec$ = torolvas("JMTR", Mid$(rec$, 145, 4), 1, 100)
          Tbknev.List1.AddItem Trim(Mid$(trec$, 5, 40))
        Case "A"
          Tbknev.List1.AddItem "A n y a g"
          trec$ = torolvas("JANY", Mid$(rec$, 149, 15), 1, 100)
          Tbknev.List1.AddItem Trim(Mid$(trec$, 16, 40))
        Case "S": Tbknev.List1.AddItem "S z a b a d t � r"
        Case "E": Tbknev.List1.AddItem "E g y � b  t e r m � k"
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "�rtartalom :" + ertszam(Trim(Mid$(rec$, 41, 12)), 14, 2)
      Tbknev.List1.AddItem "Mennyis�g  :" + ertszam(Trim(Mid$(rec$, 164, 14)), 14, 2)
      szh@ = xval(Mid$(rec$, 41, 12)) - xval(Mid$(rec$, 164, 14))
      Tbknev.List1.AddItem "Szabad hely:" + ertszam(Str$(szh@), 14, 2)
      If szh@ = 0 Then Tbknev.List1.ForeColor = QBColor(12)
    Case "JMUV"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      xx$ = " m�velet"
      Select Case Mid$(rec$, 4, 1)
        Case "B": Tbknev.List1.AddItem "Bor" + xx$
        Case "S": Tbknev.List1.AddItem "Sz�l�" + xx$
        Case "A": Tbknev.List1.AddItem "Anyag" + xx$
        Case "M": Tbknev.List1.AddItem "Mell�kterm�k" + xx$
        Case "G": Tbknev.List1.AddItem "G�ngy�leg" + xx$
        Case "E": Tbknev.List1.AddItem "Egy�b j�ved�ki term�k" + xx$
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      Select Case Mid$(rec$, 5, 1)
        Case "V": Tbknev.List1.AddItem "V�s�rl�s"
        Case "F": Tbknev.List1.AddItem "Felv�s�rl�s"
        Case "E": Tbknev.List1.AddItem "�rt�kes�t�s"
        Case "G": Tbknev.List1.AddItem "Gy�rt�s"
        Case "B": Tbknev.List1.AddItem "Bels� mozg�s"
        Case "H": Tbknev.List1.AddItem "Hi�ny"
        Case "T": Tbknev.List1.AddItem "T�bblet"
        Case "K": Tbknev.List1.AddItem "Korrekci�"
        Case "Y": Tbknev.List1.AddItem "Vesztes�g"
        Case "K": Tbknev.List1.AddItem "Korrekci�"
        Case "S": Tbknev.List1.AddItem "Sz�ret"
        Case "N": Tbknev.List1.AddItem "Nyit�"
        Case "M": Tbknev.List1.AddItem "B�rmunka �tad�s-�tv�t"
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      xx$ = " tranzakci�"
      Select Case Mid$(rec$, 6, 1)
        Case "B": Tbknev.List1.AddItem "Belf�ldi" + xx$
        Case "E": Tbknev.List1.AddItem "K�z�ss�gi" + xx$
        Case "X": Tbknev.List1.AddItem "Egy�b k�lf�ldi" + xx$
        Case "S": Tbknev.List1.AddItem "Saj�t kiskerhez kapcsol�d� " + xx$
        Case "N": Tbknev.List1.AddItem "Saj�t nagykerhez kapcsol�d� " + xx$
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      xx$ = " vesztes�g"
      Select Case Mid$(rec$, 7, 1)
        Case "T": Tbknev.List1.AddItem "T�rol�si" + xx$
        Case "S": Tbknev.List1.AddItem "Sz�ll�t�si" + xx$
        Case "M": Tbknev.List1.AddItem "M�veleti" + xx$
        Case "K": Tbknev.List1.AddItem "Kiszerel�si" + xx$
        Case Else
      End Select
    Case "JHZN"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      Select Case Mid$(rec$, 18, 1)
        Case "N": Tbknev.List1.AddItem "Nyit�k�szlet"
        Case "B": Tbknev.List1.AddItem "Beszerz�s"
        Case "F": Tbknev.List1.AddItem "Felhaszn�l�s"
        Case "M": Tbknev.List1.AddItem "Megsemmis�l�s"
        Case "S": Tbknev.List1.AddItem "Selejtez�s"
        Case "H": Tbknev.List1.AddItem "Hi�ny"
        Case "T": Tbknev.List1.AddItem "T�bblet"
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "D�tum :    " + datki(Mid$(rec$, 12, 6))
      Tbknev.List1.AddItem "Mennyis�g :" + Mid$(rec$, 19, 8) + " db"
      Tbknev.List1.AddItem " "
      trec$ = torolvas("JHFJ", Mid$(rec$, 8, 4), 1, 100)
      Tbknev.List1.AddItem Trim(Mid$(trec$, 5, 30))
      If Mid$(rec$, 164, 1) = "K" Then
        Tbknev.List1.AddItem "Kann�ra"
      Else
        Tbknev.List1.AddItem "Hord�ra"
      End If
      If Mid$(rec$, 57, 1) = "S" Then
        Tbknev.List1.AddItem "S z t o r n � z v a!"
        Tbknev.List1.ForeColor = QBColor(12)
      End If
    Case "KPAR"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(1)
      Tbknev.List1.AddItem "K�szp�nzes partnerek"
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "N�v:"
      Tbknev.List1.AddItem Mid$(rec$, 1, 60)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "C�m:"
      Tbknev.List1.AddItem Mid$(rec$, 61, 60)
    Case Else
      Tbknev.List1.Clear
      Tbknev.List1.AddItem "Nincs egy�b inform�ci�!"
  End Select
End Sub
Public Sub infomutat(objneve$, rec$, hivashely$)
  '--- keres�t�bla inform�ci� felmutat�sa
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
        Tbknev.List1.AddItem "K�vetel�s"
      Else
        Tbknev.List1.AddItem "K�telezetts�g"
      End If
      Select Case Mid$(rec$, 92, 1)
        Case "A": Tbknev.List1.AddItem "K�lts�gvet�ssel szemben"
        Case "H": Tbknev.List1.AddItem "Hitel"
        Case "K": Tbknev.List1.AddItem "K�lcs�n"
        Case "Z": Tbknev.List1.AddItem "Hozam"
        Case "E": Tbknev.List1.AddItem "Egy�b"
        Case Else
      End Select
      xx$ = "Gyakoris�g: "
      Select Case Mid$(rec$, 93, 1)
        Case "E": Tbknev.List1.AddItem xx$ + "eseti"
        Case "H": Tbknev.List1.AddItem xx$ + "havi"
        Case "N": Tbknev.List1.AddItem xx$ + "negyed�ves"
        Case "V": Tbknev.List1.AddItem xx$ + "�ves"
        Case Else
      End Select
      Tbknev.List1.AddItem "Els� esed�kess�g  : " + datki(Mid$(rec$, 94, 6))
      Tbknev.List1.AddItem "Utols� esed�kess�g: " + datki(Mid$(rec$, 100, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Teljes �sszeg:" + ertszamx(Mid$(rec$, 106, 14), 17, 0)
      Tbknev.List1.AddItem "Eseti �sszeg :" + ertszamx(Mid$(rec$, 120, 14), 17, 0)
      If Mid$(rec$, 134, 1) = "A" Then
        Tbknev.List1.AddItem "Akt�v"
      Else
        Tbknev.List1.AddItem "Passz�v"
        Tbknev.List1.ForeColor = QBColor(12)
      End If
    Case "PAFA"
      Tbknev.List1.ForeColor = QBColor(0)
      Tbknev.List1.Clear
      Tbknev.List1.AddItem Mid$(rec$, 3, 30)
      Tbknev.List1.AddItem "�FA kulcs :" + Trim(Mid$(rec$, 33, 6)) + " %"
      Tbknev.List1.AddItem " "
      Select Case Mid$(rec$, 40, 2)
        Case "BF"
          Tbknev.List1.AddItem "Belf�ldi forgalom"
        Case "XU"
          Tbknev.List1.AddItem "K�z�ss�gi �rt�kes�t�s"
        Case "XE"
          Tbknev.List1.AddItem "Egy�b export"
        Case "IU"
          Tbknev.List1.AddItem "K�z�ss�gi beszerz�s"
        Case "IE"
          Tbknev.List1.AddItem "Egy�b import"
        Case Else
      End Select
      Select Case Mid$(rec$, 39, 1)
        Case "A"
          Tbknev.List1.AddItem "Ad�alap"
        Case "M"
          Tbknev.List1.AddItem "Ad�mentes"
        Case "N"
          Tbknev.List1.AddItem "Ad�alapot nem k�pez�"
        Case Else
      End Select
      Select Case Mid$(rec$, 42, 1)
        Case "T"
          Tbknev.List1.AddItem "Teljes eg�sz�ben visszaig�nyelhet�"
        Case "R"
          Tbknev.List1.AddItem "R�szben visszaig�nelhet�"
        Case "N"
          Tbknev.List1.AddItem "Nem ig�nyelhet� vissza"
        Case Else
      End Select
      If Mid$(rec$, 42, 1) = "R" Then Tbknev.List1.AddItem "Visszaig�nyelhet� az ad� " + Trim(Mid$(rec$, 43, 6)) + " %-a"
    Case "GYUJ"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(0)
      Tbknev.List1.AddItem "Megrendel�s iktat�: " + Mid$(rec$, 8, 7)
      If Mid$(rec$, 129, 1) = "S" Then
        Tbknev.List1.AddItem "S z t o r n � z o t t  gy�jt�"
        Tbknev.List1.ForeColor = QBColor(12)
      End If
      Tbknev.List1.AddItem " "
      kszc$ = Mid$(rec$, 15, 15)
      pkod$ = Mid$(rec$, 30, 15)
      krak$ = Mid$(rec$, 45, 4)
      If Trim(pkod$) <> "" Then
        prec$ = dbxkey("PART", pkod$)
        Tbknev.List1.AddItem "Fizet�:"
        Tbknev.List1.AddItem Trim(Mid$(prec$, 16, 60))
        Tbknev.List1.AddItem Trim(postacim(prec$, 106))
        Tbknev.List1.AddItem " "
      End If
      If Trim(kszc$) <> "" Then
        prec$ = dbxkey("KSZC", kszc$)
        Tbknev.List1.AddItem "Sz�ll�t�si c�m:"
        Tbknev.List1.AddItem Trim(Mid$(prec$, 31, 60))
        Tbknev.List1.AddItem Trim(postacim(prec$, 121))
        Tbknev.List1.AddItem " "
      End If
      If Trim(krak$) <> "" Then
        prec$ = dbxkey("KRAK", krak$)
        Tbknev.List1.AddItem "Rakt�r:"
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
      Tbknev.List1.AddItem "T�telsz�m    :" + Trim(ertszam(Str(tdb), 10, 0)) + " db"
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
        Tbknev.List1.AddItem "Brutt� �rt�k:" + ertszamx(Str(bte@), 15, 0)
      End If
    Case "FUVA"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(0)
      If Mid$(rec$, 212, 1) = "*" Then
        Tbknev.List1.AddItem "T � r � l t fuvar�sszes�t�"
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
        Tbknev.List1.AddItem "Bizonylatok sz�ma:" + ertszam(Str(szla + szlev + rkozi), 10, 0)
        Tbknev.List1.AddItem "           Sz�mla:" + ertszam(Str(szla), 10, 0)
        Tbknev.List1.AddItem "    Sz�ll�t�lev�l:" + ertszam(Str(szlev), 10, 0)
        Tbknev.List1.AddItem "       Rakt�rk�zi:" + ertszam(Str(rkozi), 10, 0)
        Tbknev.List1.AddItem "      Nett� �rt�k:" + ertszam(Str(nte@), 10, 0)
        Tbknev.List1.AddItem "     Brutt� �rt�k:" + ertszam(Str(bte@), 10, 0)
        Tbknev.List1.AddItem "  "
      End If
      If Trim(Mid$(rec$, 233, 8)) = "" Then
        Tbknev.List1.AddItem "T�ra iktat�: nincs jel�lve"
        Tbknev.List1.AddItem Trim(Mid$(rec$, 14, 10)) + " " + Trim(Mid(rec$, 24, 30))
      Else
        Tbknev.List1.AddItem "T�ra iktat�: " + Mid$(rec$, 233, 7)
        komikt$ = Mid$(rec$, 233, 7)
        krec$ = dbxkey("KOMS", komikt$)
        If krec$ <> "" Then
          Tbknev.List1.AddItem Trim(Mid$(krec$, 14, 10)) + " " + Trim(Mid(krec$, 24, 30))
          If Mid$(krec$, 195, 1) = "S" Then
            Tbknev.List1.AddItem "T � r � l t t�ra"
            Tbknev.List1.ForeColor = QBColor(8)
            sat% = 0
          Else
            If Trim(Mid$(krec$, 162, 8)) = "" Then
              Tbknev.List1.AddItem "T�ra nincs elind�tva"
              Tbknev.List1.ForeColor = QBColor(0)
              sat% = 1
            Else
              If Trim(Mid$(krec$, 180, 8)) = "" Then
                Tbknev.List1.AddItem "Elind�tott t�ra"
                Tbknev.List1.ForeColor = QBColor(9)
                sat% = 2
              Else
                Tbknev.List1.AddItem "Elsz�molt t�ra"
                Tbknev.List1.ForeColor = QBColor(2)
                sat% = 3
              End If
            End If
          End If
          If sat% > 0 Then
            Tbknev.List1.AddItem "      "
            Tbknev.List1.AddItem "R�gz�tette :" + Mid$(krec$, 154, 8)
            If sat% > 1 Then
              Tbknev.List1.AddItem "Ind�totta  :" + Mid$(krec$, 162, 8) + " "
              Tbknev.List1.AddItem "           :" + datki(Mid$(krec$, 170, 6)) + "  " + Mid$(krec$, 176, 2) + ":" + Mid$(krec$, 178, 2)
              If sat% > 2 Then
                Tbknev.List1.AddItem "Elsz�molt  :" + Mid$(krec$, 180, 8) + " "
                Tbknev.List1.AddItem "           :" + datki(Mid$(krec$, 188, 6))
              End If
            End If
          End If
        End If
      End If
    Case "KOMS"
      Tbknev.List1.Clear
      If Mid$(rec$, 195, 1) = "S" Then
        Tbknev.List1.AddItem "T � r � l t t�ra"
        Tbknev.List1.AddItem "*** Nem m�dos�that� ***"
        Tbknev.List1.ForeColor = QBColor(8)
        sat% = 0
      Else
        If Trim(Mid$(rec$, 162, 8)) = "" Then
          Tbknev.List1.AddItem "N i n c s elind�tva"
          Tbknev.List1.AddItem "*** M�dos�that� ***"
          Tbknev.List1.ForeColor = QBColor(0)
          sat% = 1
        Else
          If Trim(Mid$(rec$, 180, 8)) = "" Then
            Tbknev.List1.AddItem "E l i n d � t o t t t�ra"
            Tbknev.List1.AddItem "*** Nem m�dos�that� ***"
            Tbknev.List1.ForeColor = QBColor(9)
            sat% = 2
          Else
            If Mid$(rec$, 194, 1) = "V" Then
              Tbknev.List1.AddItem "E l s z � m o l t t�ra"
              Tbknev.List1.AddItem "*** Int�zked�s sz�ks�ges ***"
              Tbknev.List1.AddItem "*** Nem m�dos�that� ***"
              Tbknev.List1.ForeColor = QBColor(12)
            Else
              Tbknev.List1.AddItem "E l s z � m o l t t�ra"
              Tbknev.List1.AddItem "*** Rendben ***"
              Tbknev.List1.AddItem "*** Nem m�dos�that� ***"
              Tbknev.List1.ForeColor = QBColor(2)
            End If
            sat% = 3
          End If
        End If
      End If
      If sat% > 0 Then
        Tbknev.List1.AddItem "      "
        Tbknev.List1.AddItem "R�gz�tette :" + Mid$(rec$, 154, 8)
        If sat% > 1 Then
          Tbknev.List1.AddItem "Ind�totta  :" + Mid$(rec$, 162, 8) + " "
          Tbknev.List1.AddItem "           :" + datki(Mid$(rec$, 170, 6)) + "  " + Mid$(rec$, 176, 2) + ":" + Mid$(rec$, 178, 2)
          If sat% > 2 Then
            Tbknev.List1.AddItem "Elsz�molt  :" + Mid$(rec$, 180, 8) + " "
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
            Tbknev.List1.AddItem "Tejdepo partner k�d"
          End If
          If Trim(qszc$) <> "" Then
            Tbknev.List1.AddItem "Tejdepo sz�ll.c�m"
          End If
          If Trim(qkod$) <> "" Then
            Tbknev.List1.AddItem "Tejdepo term�k k�d"
          End If
        Case "B"
          If Trim(qpart$) <> "" Then
            Tbknev.List1.AddItem "B�nusz partner k�d"
          End If
        Case "P"
          If Trim(qszc$) <> "" Then
            Tbknev.List1.AddItem "Pepsi vev�k�d"
          End If
          If Trim(qkod$) <> "" Then
            Tbknev.List1.AddItem "Pepsi term�k k�d"
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
          Tbknev.List1.AddItem "Sz�ll.c�m:"
          Tbknev.List1.AddItem Trim(Mid$(qrec$, 31, 60))
        End If
      End If
      If Trim(qkod$) <> "" Then
        qrec$ = dbxkey("KTRM", qkod$)
        If qrec$ <> "" Then
          Tbknev.List1.AddItem " "
          Tbknev.List1.AddItem "Term�k:"
          Tbknev.List1.AddItem Trim(Mid$(qrec$, 16, 60))
        End If
      End If
      If Mid$(rec$, 300, 1) = "S" Then
        Tbknev.List1.AddItem " "
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t kapcsolat"
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
      Tbknev.List1.AddItem "Mintasz�m : " + Mid$(rec$, 33, 8)
      Tbknev.List1.AddItem "T�rhely   : " + Mid$(rec$, 41, 8)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Fordul�nap: " + datki(Mid$(rec$, 8, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Mennyis�g : " + ertszam(Mid$(rec$, 49, 12), 12, 2)
      Tbknev.List1.AddItem "Egys�g�r  : " + ertszam(Mid$(rec$, 85, 12), 12, 2)
      If Mid$(rec$, 111, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
        Tbknev.List1.AddItem "Sztorn�zva:" + datki(Mid$(rec$, 112, 6)) + " " + Mid$(rec$, 118, 8) + " "
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
      Tbknev.List1.AddItem "Hat�rid�   : " + datki(Mid$(rec$, 49, 6))
      Tbknev.List1.AddItem "Megrendelt : " + ertszam(Mid$(rec$, 138, 14), 14, 2)
      Tbknev.List1.AddItem "Teljes�tett: " + ertszam(Mid$(rec$, 152, 14), 14, 2)
      If Mid$(rec$, 226, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
    
    Case "JSZB"
      Tbknev.List1.Clear
      Select Case Mid$(rec$, 11, 2)
        Case "SL": Tbknev.List1.AddItem "Sz�ll�t�lev�l"
        Case "BS"
          If Trim(Mid$(rec$, 94, 3)) = "" Then
            Tbknev.List1.AddItem "Belf�ldi forint sz�mla"
          Else
            Tbknev.List1.AddItem "Belf�ldi deviz�s sz�mla"
          End If
        Case "BP": Tbknev.List1.AddItem "Belf�ldi proforma sz�mla"
        Case "EU": Tbknev.List1.AddItem "K�z�ss�gi sz�mla"
        Case "EP": Tbknev.List1.AddItem "K�z�ss�gi proforma sz�mla"
        Case "XS": Tbknev.List1.AddItem "Export sz�mla"
        Case "SP": Tbknev.List1.AddItem "Export proforma sz�mla"
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
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
        Tbknev.List1.AddItem "Sztorn�sz�mla: " + Mid$(rec$, 194, 10)
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Telj.kelte  : " + datki(Mid$(rec$, 74, 6))
      Tbknev.List1.AddItem "Sz�mla kelte: " + datki(Mid$(rec$, 78, 6))
      Tbknev.List1.AddItem "Fiz.hat�rid�: " + datki(Mid$(rec$, 86, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "T�telsz�m   : " + Str(xval(Mid$(rec$, 204, 3))) + " db"
      If Trim(Mid$(rec$, 94, 3)) <> "" Then
        Tbknev.List1.AddItem Mid$(rec$, 94, 3) + " �rf.:" + Mid$(rec$, 97, 10)
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Kis.okm�ny : " + Mid$(rec$, 217, 12)
      Tbknev.List1.AddItem "Megrendel�s: " + Mid$(rec$, 250, 15)
    Case "JFRG"
      Tbknev.List1.Clear
      muvk$ = Mid$(rec$, 377, 3)
      muvrec$ = dbxkey("JMUV", muvk$)
      If muvrec$ <> "" Then
        Tbknev.List1.AddItem Mid$(muvrec$, 9, 50)
      End If
      Select Case Mid$(rec$, 36, 1)
        Case "V": Tbknev.List1.AddItem "V�s�rl�s"
        Case "F": Tbknev.List1.AddItem "Felv�s�rl�s"
        Case "E": Tbknev.List1.AddItem "Elad�s"
        Case "G": Tbknev.List1.AddItem "Gy�rt�s"
        Case "B": Tbknev.List1.AddItem "Bels� mozg�s"
        Case "H": Tbknev.List1.AddItem "Hi�ny"
        Case "T": Tbknev.List1.AddItem "T�bblet"
        Case "K": Tbknev.List1.AddItem "Korrekci�"
        Case "Y": Tbknev.List1.AddItem "Vesztes�g"
        Case "S": Tbknev.List1.AddItem "Sz�ret"
        Case "N": Tbknev.List1.AddItem "Nyit� k�szlet"
      End Select
      Tbknev.List1.AddItem " "
      Select Case Mid$(rec$, 35, 1)
        Case "B"
          Tbknev.List1.AddItem "Bor"
          tkod$ = Mid$(rec$, 40, 15)
          termrec$ = dbxkey("JTRM", tkod$)
          Tbknev.List1.AddItem Mid$(termrec$, 90, 40)
          Tbknev.List1.AddItem "Mennyis�g:" + Mid$(rec$, 134, 14) + " L"
        Case "M"
          Tbknev.List1.AddItem "Mell�kterm�k"
          tkod$ = Mid$(rec$, 81, 4)
          termrec$ = dbxkey("JMTR", tkod$)
          Tbknev.List1.AddItem Mid$(termrec$, 5, 40)
          Tbknev.List1.AddItem "Mennyis�g:" + Mid$(rec$, 134, 14) + " " + Mid$(termrec$, 75, 6)
        Case "A"
          Tbknev.List1.AddItem "Anyag"
          tkod$ = Mid$(rec$, 66, 15)
          termrec$ = dbxkey("JANY", tkod$)
          Tbknev.List1.AddItem Mid$(termrec$, 16, 40)
          Tbknev.List1.AddItem "Mennyis�g:" + Mid$(rec$, 134, 14) + " " + Mid$(termrec$, 86, 6)
        Case "S"
          Tbknev.List1.AddItem "Sz�l�"
          tkod$ = Mid$(rec$, 62, 4)
          termrec$ = dbxkey("JSZO", tkod$)
          Tbknev.List1.AddItem Mid$(termrec$, 5, 30)
          sm@ = xval(Mid$(rec$, 294, 12)) - xval(Mid$(rec$, 306, 12))
          Tbknev.List1.AddItem "Mennyis�g:" + ertszam(Str(sm@), 14, 2) + " Kg"
        Case "G": Tbknev.List1.AddItem "G�ngy�leg"
        Case "E"
          Tbknev.List1.AddItem "Egy�b j�ved�ki term�k"
          tkod$ = Mid$(rec$, 40, 15)
          termrec$ = dbxkey("JTRM", tkod$)
          Tbknev.List1.AddItem Mid$(termrec$, 90, 40)
          Tbknev.List1.AddItem "Mennyis�g:" + Mid$(rec$, 134, 14) + " " + Mid$(termrec$, 200, 6)
        Case Else
      End Select
      Select Case Mid$(rec$, 85, 1)
        Case "T"
          Tbknev.List1.AddItem "Tart�ly  :" + Mid$(rec$, 86, 20)
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
      Tbknev.List1.AddItem "Bels� bizonylat: " + Mid$(rec$, 370, 7)
      Tbknev.List1.AddItem "Sz�mla sz�ma   : " + Mid$(rec$, 162, 13)
      Tbknev.List1.AddItem "Kis�r� okm�ny  : " + Mid$(rec$, 205, 12)
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 217, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
        Tbknev.List1.AddItem "Sztorn� kelte: " + datki(Mid$(rec$, 218, 6))
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
    Case "PSZB"
      Tbknev.List1.Clear
      Select Case Mid$(rec$, 76, 2)
        Case "SL": Tbknev.List1.AddItem "Sz�ll�t�lev�l"
        Case "BS"
          If Trim(Mid$(rec$, 98, 3)) = "" Then
            Tbknev.List1.AddItem "Belf�ldi forint sz�mla"
          Else
            Tbknev.List1.AddItem "Belf�ldi deviz�s sz�mla"
          End If
        Case "BP": Tbknev.List1.AddItem "Belf�ldi proforma sz�mla"
        Case "EU": Tbknev.List1.AddItem "K�z�ss�gi sz�mla"
        Case "EP": Tbknev.List1.AddItem "K�z�ss�gi proforma sz�mla"
        Case "XS": Tbknev.List1.AddItem "Export sz�mla"
        Case "SP": Tbknev.List1.AddItem "Export proforma sz�mla"
      End Select
      qpart$ = Mid$(rec$, 61, 15)
      qrec$ = dbxkey("PART", qpart)
      If qrec$ <> "" Then
        Tbknev.List1.AddItem Mid$(qrec$, 16, 60)
      End If
      If Mid$(rec$, 35, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
        Tbknev.List1.AddItem "Sztorn�sz�mla: " + Mid$(rec$, 237, 15)
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Telj.kelte  : " + datki(Mid$(rec$, 84, 6))
      Tbknev.List1.AddItem "Sz�mla kelte: " + datki(Mid$(rec$, 78, 6))
      Tbknev.List1.AddItem "Fiz.hat�rid�: " + datki(Mid$(rec$, 90, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "T�telsz�m   : " + Str(xval(Mid$(rec$, 50, 3))) + " db"
      kod$ = Mid$(rec$, 60, 1)
      If Trim(Mid$(rec$, 98, 3)) <> "" Then
        Tbknev.List1.AddItem Mid$(rec$, 98, 3) + " �rf.:" + Mid$(rec$, 101, 10)
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "R�gz�tette: " + Mid$(rec$, 21, 8) + " " + datki(Mid$(rec$, 29, 6))
      If Mid$(rec$, 35, 1) = "S" Then
        Tbknev.List1.AddItem "Sztorn�zta: " + Mid$(rec$, 42, 8) + " " + datki(Mid$(rec$, 36, 6))
      End If
    Case "KSZB"
      Tbknev.List1.Clear
      Select Case Mid$(rec$, 76, 2)
        Case "SL": Tbknev.List1.AddItem "Sz�ll�t�lev�l"
        Case "BS"
          If Trim(Mid$(rec$, 98, 3)) = "" Then
            Tbknev.List1.AddItem "Belf�ldi forint sz�mla"
          Else
            Tbknev.List1.AddItem "Belf�ldi deviz�s sz�mla"
          End If
        Case "BP": Tbknev.List1.AddItem "Belf�ldi proforma sz�mla"
        Case "EU": Tbknev.List1.AddItem "K�z�ss�gi sz�mla"
        Case "EP": Tbknev.List1.AddItem "K�z�ss�gi proforma sz�mla"
        Case "XS": Tbknev.List1.AddItem "Export sz�mla"
        Case "SP": Tbknev.List1.AddItem "Export proforma sz�mla"
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
              Tbknev.List1.AddItem "�zl.:" + Mid$(kuzlrec$, 9, 60)
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
              Tbknev.List1.AddItem "�zl.:" + Mid$(kuzlrec$, 9, 60)
            End If
          End If
        End If
      End If
      If Mid$(rec$, 35, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
        Tbknev.List1.AddItem "Sztorn�sz�mla: " + Mid$(rec$, 282, 15)
      Else
        Tbknev.List1.ForeColor = QBColor(3)
        If Mid$(rec$, 297, 1) = "I" Then
          Tbknev.List1.ForeColor = QBColor(9)
          Tbknev.List1.AddItem "K i a d a t l a n!"
        End If
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Telj.kelte  : " + datki(Mid$(rec$, 84, 6))
      Tbknev.List1.AddItem "Sz�mla kelte: " + datki(Mid$(rec$, 78, 6))
      Tbknev.List1.AddItem "Fiz.hat�rid�: " + datki(Mid$(rec$, 90, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "T�telsz�m   : " + Str(xval(Mid$(rec$, 50, 3))) + " db"
      kod$ = Mid$(rec$, 60, 1)
      If kod$ = "1" Or kod$ = "2" Or kod$ = "3" Or kod$ = "4" Then
        Tbknev.List1.AddItem "Brutt� �rt�k: " + Mid$(rec$, 101, 10)
      Else
        If Trim(Mid$(rec$, 98, 3)) <> "" Then
          Tbknev.List1.AddItem Mid$(rec$, 98, 3) + " �rf.:" + Mid$(rec$, 101, 10)
        End If
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "R�gz�tette: " + Mid$(rec$, 21, 8) + " " + datki(Mid$(rec$, 29, 6))
      If Mid$(rec$, 35, 1) = "S" Then
        Tbknev.List1.AddItem "Sztorn�zta: " + Mid$(rec$, 42, 8) + " " + datki(Mid$(rec$, 36, 6))
      End If
    Case "PMEG"
      Tbknev.List1.Clear
      If Mid$(rec$, 226, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem "Vev�:"
      pkod$ = Mid$(rec$, 8, 15)
      If Trim(pkod) <> "" Then
        prec$ = dbxkey("PART", pkod$)
        Tbknev.List1.AddItem Trim(Mid$(prec$, 16, 60))
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "C�m:"
        Tbknev.List1.AddItem Trim(postacim(prec$, 106))
      End If
      Tbknev.List1.AddItem "Hiv.sz�m:    " + Trim(Mid$(rec$, 23, 20))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Megr.kelte:  " + datki(Mid$(rec$, 43, 6))
      Tbknev.List1.AddItem "Sz.hat�rid�: " + datki(Mid$(rec$, 49, 6))
      Tbknev.List1.AddItem "Hiv.sz�m:    " + Trim(Mid$(rec$, 23, 20))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem Trim(Mid$(rec$, 61, 60))
      Tbknev.List1.AddItem "Egys�g�r: " + Mid$(rec$, 121, 14) + " " + Mid$(rec$, 135, 3)
      Tbknev.List1.AddItem "Mennyis�g:" + Mid$(rec$, 138, 14) + " " + Trim(Mid$(rec$, 226, 6))
      Tbknev.List1.AddItem " "
      m@ = xval(Mid$(rec$, 138, 14))
      t@ = xval(Mid$(rec$, 152, 14))
      If t@ >= m@ Then
        Tbknev.List1.AddItem "Teljes�t�s megt�rt�nt!"
      Else
        If t@ < m@ And t@ > 0 Then
          Tbknev.List1.AddItem "R�szteljes�t�s t�rt�nt!"
        Else
          Tbknev.List1.AddItem "Nem volt teljes�t�s!"
        End If
      End If
    
    Case "KMEG"
      Tbknev.List1.Clear
      If Mid$(rec$, 406, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem "Vev�:"
      Tbknev.List1.AddItem Trim(Mid$(rec$, 420, 60))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Sz�ll�t�si c�m:"
      Tbknev.List1.AddItem Trim(postacim(rec$, 88))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Megr.kelte:  " + datki(Mid$(rec$, 174, 6))
      Tbknev.List1.AddItem "Sz.hat�rid�: " + datki(Mid$(rec$, 180, 6))
      Tbknev.List1.AddItem "�gyint�z�:   " + Trim(Mid$(rec$, 382, 8))
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
      Tbknev.List1.AddItem "T�telek: " + ertszam(Str(vidb%), 12, 0)
      Tbknev.List1.AddItem "N.�rt�k: " + ertszam(Str(viert), 12, 2)
      Tbknev.List1.AddItem " "
      If vihdb = 0 Then
        Tbknev.List1.AddItem "Teljes�t�s megt�rt�nt!"
      Else
        If vitdb > 0 Then
          Tbknev.List1.AddItem "R�szteljes�t�s t�rt�nt!"
        Else
          Tbknev.List1.AddItem "Nem volt teljes�t�s!"
        End If
      End If
    Case "KSMG"
      Tbknev.List1.Clear
      If Mid$(rec$, 6432, 1) = "O" Then
        Tbknev.List1.AddItem "5-ker r�sz�re"
      Else
        Tbknev.List1.AddItem "Merkatimpex r�sz�re"
      End If
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 398, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem "Sz�ll�t�:"
      Tbknev.List1.AddItem Trim(Mid$(rec$, 23, 60))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Sz�ll�t�si c�m:"
      Tbknev.List1.AddItem Trim(Mid$(rec$, 420, 60))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Megr.kelte:  " + datki(Mid$(rec$, 192, 6))
      Tbknev.List1.AddItem "Sz.hat�rid�: " + datki(Mid$(rec$, 198, 6))
      Tbknev.List1.AddItem "�gyint�z�:   " + Trim(Mid$(rec$, 6401, 30))
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
      Tbknev.List1.AddItem "T�telek: " + ertszam(Str(vidb%), 12, 0)
      Tbknev.List1.AddItem "N.�rt�k: " + ertszam(Str(viert), 12, 2)
      Tbknev.List1.AddItem " "
      If vihdb = 0 Then
        Tbknev.List1.AddItem "Teljes�t�s megt�rt�nt!"
      Else
        If vitdb > 0 Then
          Tbknev.List1.AddItem "R�szteljes�t�s t�rt�nt!"
        Else
          Tbknev.List1.AddItem "Nem volt teljes�t�s!"
        End If
      End If
    Case "KFRG"
      Tbknev.List1.Clear
      If Mid$(rec$, 101, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Tbknev.List1.AddItem "Hivatkoz�s: " + Mid$(rec$, 15, 15)
      Tbknev.List1.AddItem "D�tum:    : " + datki(Mid$(rec$, 30, 6))
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
      Tbknev.List1.AddItem "Besz.�r  :" + Mid$(rec$, 39, 14)
      Tbknev.List1.AddItem "N�veked�s:" + Mid$(rec$, 59, 14)
      Tbknev.List1.AddItem "Cs�kken�s:" + Mid$(rec$, 73, 14)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "R�gz�tette: " + Mid$(rec$, 93, 8) + " " + datki(Mid$(rec$, 87, 6))
      If Mid$(rec$, 201, 1) = "S" Then
        Tbknev.List1.AddItem "Sztorn�zta: " + Mid$(rec$, 108, 8) + " " + datki(Mid$(rec$, 102, 6))
      End If
    
    Case "KKBZ"
      Tbknev.List1.Clear
      If Mid$(rec$, 92, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Select Case Mid$(rec$, 54, 1)
        Case "B": Tbknev.List1.AddItem "Beszerz�s"
        Case "E": Tbknev.List1.AddItem "�rt�kes�t�s"
        Case "V": Tbknev.List1.AddItem "Vissz�ru sz�ll�t�nak"
        Case "T": Tbknev.List1.AddItem "Termel�s k�szletre v�tel"
        Case "R": Tbknev.List1.AddItem "Rakt�rk�zi bizonylat"
        Case "F": Tbknev.List1.AddItem "Felhaszn�l�s"
        Case "S": Tbknev.List1.AddItem "Selejtez�s"
        Case Else: Tbknev.List1.AddItem "Egy�b bizonylat"
      End Select
      qprt$ = Mid$(rec$, 55, 15)
      If Trim(qprt$) <> "" Then
        qpart$ = dbxkey("PART", qprt$)
        If qpart$ <> "" Then
          Tbknev.List1.AddItem Mid$(qpart$, 16, 60)
        End If
      End If
      Tbknev.List1.AddItem "Sz�mlasz�m  : " + Mid$(rec$, 107, 15)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Hivatkoz�s  : " + Mid$(rec$, 130, 10)
      Tbknev.List1.AddItem "Munkasz�m   : " + Mid$(rec$, 70, 8)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "R�gz�tette: " + Mid$(rec$, 84, 8) + " " + datki(Mid$(rec$, 78, 6))
      If Mid$(rec$, 201, 1) = "S" Then
        Tbknev.List1.AddItem "Sztorn�zta: " + Mid$(rec$, 99, 8) + " " + datki(Mid$(rec$, 93, 6))
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
              Tbknev.List1.AddItem "�zl.:" + Mid$(kuzlrec$, 9, 60)
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
        Tbknev.List1.AddItem "S z t o r n � z o t t bizonylat"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      If Mid$(rec$, 8, 2) = "M:" Then
        If Mid$(rec$, 201, 1) = "S" Then
          Tbknev.List1.ForeColor = QBColor(12)
        Else
          Tbknev.List1.ForeColor = RGB(0, 60, 120)
        End If
        Tbknev.List1.AddItem "M � d o s � t � bizonylat"
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Telj.kelte  : " + datki(Mid$(rec$, 39, 6))
      Tbknev.List1.AddItem "Sz�ll.lev�l : " + Mid$(rec$, 155, 10)
      Tbknev.List1.AddItem "Sz�mlasz�m  : " + Mid$(rec$, 135, 10)
      Tbknev.List1.AddItem "Fiz.m�d     : " + Mid$(rec$, 242, 2) + " --> " + datki(Mid$(rec$, 235, 6))
      'Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "T�telsz�m   : " + Str(xval(Mid$(rec$, 189, 4))) + " db"
      Tbknev.List1.AddItem "Nett� �rt�k : " + Mid$(rec$, 145, 10)
      Tbknev.List1.AddItem "Brutt� �rt�k: " + Mid$(rec$, 165, 10)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "R�gz�tette: " + Mid$(rec$, 193, 8) + " " + datki(Mid$(rec$, 20, 6))
      If Mid$(rec$, 201, 1) = "S" Then
        Tbknev.List1.AddItem "Sztorn�zta: " + Mid$(rec$, 208, 8) + " " + datki(Mid$(rec$, 202, 6))
      End If
      If Len(Trim(Mid$(rec$, 216, 8))) = 7 Then
        Tbknev.List1.AddItem "Fuvar ikt.: " + Mid$(rec$, 216, 7)
        fuvrec$ = dbxkey("FUVA", Mid$(rec$, 216, 8))
        If fuvrec$ = "" Then
          Tbknev.List1.AddItem "T�ra ikt. : nincs jel�lve"
        Else
          If Trim(Mid$(fuvrec$, 233, 7)) = "" Then
            Tbknev.List1.AddItem "T�ra ikt. : nincs jel�lve"
          Else
            Tbknev.List1.AddItem "T�ra ikt. : " + Mid$(fuvrec$, 233, 7)
          End If
        End If
      Else
        Tbknev.List1.AddItem "Fuvar ikt.: nincs jel�lve"
      End If
      If Mid$(rec$, 155, 2) = "J:" Then
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "M�dos�t� sz�mla: " + Mid$(rec$, 157, 7)
      Else
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "M�dos�t� sz�mla: nincs"
      End If
    Case "FKTE"
      Tbknev.List1.Clear
      If Mid$(rec$, 61, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem "Sztorn�zott t�tel"
      Else
        Tbknev.List1.ForeColor = QBColor(3)
      End If
      Select Case Mid$(rec$, 45, 4)
        Case "EIR ": Tbknev.List1.AddItem "Terv t�tel"
        Case "TVE ": Tbknev.List1.AddItem "Vegyes k�nyvel�si bizonylat"
        Case "TVNA": Tbknev.List1.AddItem "Automatikus nyit� t�tel"
        Case "TVNK": Tbknev.List1.AddItem "K�zi nyit� t�tel"
        Case "TPS ": Tbknev.List1.AddItem "Sz�ll�t� sz�mla"
        Case "TPSH": Tbknev.List1.AddItem "Sz�ll�t� sz�mla helyesb�t�"
        Case "TPV ": Tbknev.List1.AddItem "Vev� sz�mla"
        Case "TPVH": Tbknev.List1.AddItem "Vev� sz�mla helyesb�t�"
        Case "TPEV": Tbknev.List1.AddItem "Vev� el�leg"
        Case "TPES": Tbknev.List1.AddItem "Sz�ll�t� el�leg"
        Case "TPPP": Tbknev.List1.AddItem "P�nzt�ri t�tel"
        Case "TPPB": Tbknev.List1.AddItem "Banki t�tel"
        Case "TPPK": Tbknev.List1.AddItem "P�nz�gyi korrekci�s t�tel"
        Case "TFEA": Tbknev.List1.AddItem "Eszk�z �llom�nyv�ltoz�s"
        Case "TFEC": Tbknev.List1.AddItem "Eszk�z �rt�kecs�kkeb�s"
        Case "TFKE": Tbknev.List1.AddItem "Egyszer�s�tett k�szlet felad�s"
        Case "TFKK": Tbknev.List1.AddItem "K�szlet felad�s"
        Case "TFB ": Tbknev.List1.AddItem "Munkab�r felad�s"
        Case "TFX ": Tbknev.List1.AddItem "Idegen szoftverb�l import�lt"
        Case Else
      End Select
      If Trim(Mid$(rec$, 192, 60)) <> "" Then Tbknev.List1.AddItem Mid$(rec$, 192, 60)
      If Trim(Mid$(rec$, 252, 25)) <> "" Then Tbknev.List1.AddItem Mid$(rec$, 252, 25)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "K�nyvel�s kelte: " + datki(Mid$(rec$, 49, 6))
      Tbknev.List1.AddItem "Sztorn� kelte:   " + datki(Mid$(rec$, 62, 6))
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
      Tbknev.List1.AddItem "Szerv.egys�g:   " + Mid$(rec$, 84, 8)
      Tbknev.List1.AddItem "Munkasz�m:      " + Mid$(rec$, 92, 8)
      Tbknev.List1.AddItem "Partner k�d:    " + Mid$(rec$, 163, 15)
      If Trim(Mid$(rec$, 163, 15)) <> "" Then
        qpart$ = Mid$(rec$, 163, 15)
        qrec$ = dbxkey("PART", qpart$)
        If qrec$ <> "" Then Tbknev.List1.AddItem Mid$(qrec$, 16, 60)
      End If
      Tbknev.List1.AddItem "Sz�mla iktat�:  " + Mid$(rec$, 178, 7)
      Tbknev.List1.AddItem "P.forg.iktat�:  " + Mid$(rec$, 185, 7)
    Case "KRAK"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      Tbknev.List1.AddItem Mid$(rec$, 68, 60)
      Select Case Mid$(rec$, 65, 1)
        Case "S": Tbknev.List1.AddItem "Saj�t k�zponti rakt�r"
        Case "L": Tbknev.List1.AddItem "Lerakat"
        Case "M": Tbknev.List1.AddItem "Mozg� rakt�r"
        Case "K": Tbknev.List1.AddItem "Kiskereskedelmi egys�g"
        Case "J": Tbknev.List1.AddItem "J�ved�ki rakt�r"
        Case "B": Tbknev.List1.AddItem "Bizom�nyos"
        Case "Z": Tbknev.List1.AddItem "Besz�ll�t�i rakt�r"
        Case "I": Tbknev.List1.AddItem "ISO el�min�s�t� rakt�r"
      End Select
      pkod$ = Mid$(rec$, 128, 15)
      qpart$ = dbxkey("PART", pkod$)
      If qpart$ <> "" Then
        Tbknev.List1.AddItem "Partner: " + qpart$
        Tbknev.List1.AddItem Mid$(qpart, 16, 60)
        Tbknev.List1.AddItem postacim(qpart, 106)
      End If
      If Mid$(rec$, 66, 1) = "I" Then Tbknev.List1.AddItem "T�rhelyes rakt�r"
      If Mid$(rec$, 67, 1) = "I" Then Tbknev.List1.AddItem "Mintasz�mos rakt�r "
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "K�szl.sz�mla: " + Mid$(rec$, 143, 8)
      Tbknev.List1.AddItem "�rbev.sz�mla: " + Mid$(rec$, 151, 8)
      Tbknev.List1.AddItem "EL�B�.sz�mla: " + Mid$(rec$, 185, 8)
      Tbknev.List1.AddItem "Szerv.egys�g: " + Mid$(rec$, 159, 8)
      Tbknev.List1.AddItem "Munkasz�m:    " + Mid$(rec$, 167, 8)
    Case "KMOX"
      Tbknev.List1.Clear
      If Mid$(rec$, 34, 1) = "B" Then
        Tbknev.List1.ForeColor = QBColor(3)
        Tbknev.List1.AddItem "K�szlet n�vel� (bev�t)"
      Else
        Tbknev.List1.ForeColor = QBColor(4)
        Tbknev.List1.AddItem "K�szlet cs�kkent� (kiad�s)"
      End If
      Select Case Mid$(rec$, 43, 2)
        Case "BS": Tbknev.List1.AddItem "Beszerz�s"
        Case "TM": Tbknev.List1.AddItem "Termel�s"
        Case "TB": Tbknev.List1.AddItem "Lelt�r t�bblet"
        Case "BE": Tbknev.List1.AddItem "Egy�b bev�t"
        Case "ER": Tbknev.List1.AddItem "�rt�kes�t�s"
        Case "VI": Tbknev.List1.AddItem "Vissz�ru"
        Case "SE": Tbknev.List1.AddItem "Selejtez�s"
        Case "FH": Tbknev.List1.AddItem "Saj�t felhaszn�l�s"
        Case "HI": Tbknev.List1.AddItem "Lelt�r hi�ny"
        Case "KE": Tbknev.List1.AddItem "Egy�b kiad�s"
        Case "RK": Tbknev.List1.AddItem "Rakt�rk�zi mozg�s"
      End Select
      Tbknev.List1.AddItem "Kapcsolt mozg�s: " + Mid$(rec$, 47, 3)
      If Mid$(rec$, 45, 1) = "I" Then
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "A mozg�shoz munkasz�m k�telez�"
      End If
      If Mid$(rec$, 46, 1) = "K" Then
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "A mozg�shoz besz.�r kit�lt�se:"
        Tbknev.List1.AddItem "   k � t e l e z �! "
        Tbknev.List1.AddItem "Az �tlag�r m�dosul."
      Else
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "A mozg�shoz besz.�r kit�lt�se:"
        Tbknev.List1.AddItem "   t i l o s! "
        Tbknev.List1.AddItem "Az �tlag�r nem m�dosul."
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Ellensz�mla :" + Mid$(rec$, 35, 8)
      Tbknev.List1.AddItem "K�lts�ghely :" + Mid$(rec$, 67, 8)
      Tbknev.List1.AddItem "Szerv.egys�g:" + Mid$(rec$, 51, 8)
      Tbknev.List1.AddItem "Munkasz�m:   " + Mid$(rec$, 59, 8)
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
        Case "A": Tbknev.List1.AddItem "Anyag, �ru"
        Case "S": Tbknev.List1.AddItem "Szolg�ltat�s"
        Case "F": Tbknev.List1.AddItem "Folyamatos szolg�ltat�sa"
      End Select
      Select Case Mid$(rec$, 187, 1)
        Case "I": Tbknev.List1.AddItem "K�szlet kezel�ssel"
        Case Else: Tbknev.List1.AddItem "K�szlet kezel�s n�lk�l"
      End Select
      Tbknev.List1.AddItem "M.egys�g: " + Mid$(rec$, 140, 6)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Egys�g�r:  " + Mid$(rec$, 146, 12) + " " + Mid$(rec$, 158, 3)
      Tbknev.List1.AddItem "Nyilv.�r:" + Mid$(rec$, 301, 14) + " " + Mid$(rec$, 158, 3)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "K�szlet: " + Mid$(rec$, 315, 14) + " " + Mid$(rec$, 158, 3)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "K�szl.sz�mla:" + Mid$(rec$, 330, 8)
      Tbknev.List1.AddItem "�rbev.sz�mla:" + Mid$(rec$, 163, 8)
      Tbknev.List1.AddItem "Szerv.egys�g:" + Mid$(rec$, 171, 8)
      Tbknev.List1.AddItem "Munkasz�m:   " + Mid$(rec$, 179, 8)
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
        Tbknev.List1.AddItem "Norma k�d:" + Trim(Mid$(rec$, 606, 8))
        Tbknev.List1.AddItem Mid$(enrmrec$, 9, 30)
      End If
      Tbknev.List1.AddItem "�CS kulcs:" + Mid$(rec$, 615, 6) + " %"
      Tbknev.List1.AddItem "T�rol�s:  " + Mid$(rec$, 542, 40)
      Tbknev.List1.AddItem "L�tes�t�s:" + datki(Mid$(rec$, 583, 6))
      Tbknev.List1.AddItem "Aktiv�l�s:" + datki(Mid$(rec$, 730, 6))
      Tbknev.List1.AddItem "Kivezet�s:" + datki(Mid$(rec$, 736, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Brutt� �rt�k:    " + ertszamx(Mid$(rec$, 661, 14), 15, 0)
      Tbknev.List1.AddItem "Maradv�ny �rt�k: " + ertszamx(Mid$(rec$, 675, 14), 15, 0)
      Tbknev.List1.AddItem "Fejl.tartal�kb�l:" + ertszamx(Mid$(rec$, 750, 14), 15, 0)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Nyilv.sz�mla:" + Mid$(rec$, 621, 8)
      Tbknev.List1.AddItem "�CS.sz�mla:  " + Mid$(rec$, 629, 8)
      Tbknev.List1.AddItem "K�lts�ghely: " + Mid$(rec$, 637, 8)
      Tbknev.List1.AddItem "Szerv.egys�g:" + Mid$(rec$, 645, 8)
      Tbknev.List1.AddItem "Munkasz�m:   " + Mid$(rec$, 653, 8)
    Case "FKSZ"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      Select Case Mid$(rec$, 69, 1)
        Case "T": Tbknev.List1.AddItem "Technikai"
        Case "B": Tbknev.List1.AddItem "Bank"
        Case "P": Tbknev.List1.AddItem "P�nzt�r"
        Case "K": Tbknev.List1.AddItem "K�lts�ghely, k�lts�gvisel�"
        Case "V": Tbknev.List1.AddItem "Nyilv�ntart�si"
        Case "N": Tbknev.List1.AddItem "Norm�l"
      End Select
      Select Case Mid$(rec$, 70, 1)
        Case "E": Tbknev.List1.AddItem "Eszk�z"
        Case "F": Tbknev.List1.AddItem "Forr�s"
        Case "K": Tbknev.List1.AddItem "K�lts�g, r�ford�t�s"
        Case "B": Tbknev.List1.AddItem "Bev�tel, �rbev�tel"
        Case "N": Tbknev.List1.AddItem "Nyit�m�rleg"
        Case "V": Tbknev.List1.AddItem "Nyilv�ntart�si"
        Case "Z": Tbknev.List1.AddItem "Z�r�m�rleg"
        Case "R": Tbknev.List1.AddItem "Eredm�ny"
      End Select
      Select Case Mid$(rec$, 71, 1)
        Case "I": Tbknev.List1.AddItem "K�nyvel�s mindenhol"
        Case "V": Tbknev.List1.AddItem "Csak vegyes bizonylaton"
        Case "N": Tbknev.List1.AddItem "Vegyes bizonylaton tilos"
        Case "L": Tbknev.List1.AddItem "K�nyvel�s letiltva"
      End Select
      If Mid$(r$, 76, 1) = "*" Then Tbknev.List1.AddItem "Gy�jt� sz�mla" Else Tbknev.List1.AddItem "Analitikus sz�mla"
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "          Tartozik     K�vetel"
      Tbknev.List1.AddItem "Nyit�:" + ertszam(Mid$(rec$, 290, 14), 12, 0) + ertszam(Mid$(rec$, 304, 14), 12, 0)
      Tbknev.List1.AddItem "Forg.:" + ertszam(Mid$(rec$, 318, 14), 12, 0) + ertszam(Mid$(rec$, 332, 14), 12, 0)
      zegya@ = xval(Mid$(rec$, 290, 14)) - xval(Mid$(rec$, 304, 14))
      zegya@ = zegya@ + xval(Mid$(rec$, 318, 14)) - xval(Mid$(rec$, 332, 14))
      If zegya@ > 0 Then tegy@ = zegya@: kegy@ = 0 Else kegy@ = Abs(zegya@): tegy@ = 0
      Tbknev.List1.AddItem "Z�r� :" + ertszam(Str(tegy@), 12, 0) + ertszam(Str(kegy@), 12, 0)
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
        Tbknev.List1.AddItem "Z�r.:" + datki(Mid$(rec$, 443, 6))
        Tbknev.List1.AddItem "Dev.:" + Mid$(rec$, 471, 3)
      End If
    Case "AKCF"
      Tbknev.List1.Clear
      If Mid$(rec$, 38, 1) = "I" Then
        Tbknev.List1.AddItem "Sz�ll�t�i akci�"
        Tbknev.List1.AddItem "   " + datki(Mid$(rec$, 41, 6)) + "-" + datki(Mid$(rec$, 47, 6))
      End If
      If Mid$(rec$, 39, 1) = "I" Then
        Tbknev.List1.AddItem "Nagyker akci�"
        Tbknev.List1.AddItem "   " + datki(Mid$(rec$, 53, 6)) + "-" + datki(Mid$(rec$, 59, 6))
      End If
      If Mid$(rec$, 40, 1) = "I" Then
        Tbknev.List1.AddItem "Fogyaszt�i akci�"
        Tbknev.List1.AddItem "   " + datki(Mid$(rec$, 65, 6)) + "-" + datki(Mid$(rec$, 71, 6))
      End If
      gykod$ = Mid$(rec$, 77, 15)
      Tbknev.List1.AddItem " "
      If Trim(gykod$) <> "" Then
        prec$ = dbxkey("PART", gykod)
        Tbknev.List1.AddItem "Gy�rt�:"
        Tbknev.List1.AddItem (Mid$(prec$, 16, 60))
        Tbknev.List1.AddItem " "
      End If
      szkod$ = Mid$(rec$, 92, 15)
      If Trim(szkod$) <> "" Then
        prec$ = dbxkey("PART", szkod)
        Tbknev.List1.AddItem "Sz�ll�t�:"
        Tbknev.List1.AddItem (Mid$(prec$, 16, 60))
      End If
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 141, 1) = "S" Then
        Tbknev.List1.AddItem "Sztorn�zott akci�"
        Tbknev.List1.ForeColor = QBColor(12)
      Else
        Tbknev.List1.ForeColor = QBColor(0)
      End If
    Case "CRMG"
      Tbknev.List1.Clear
      If Mid$(rec$, 31, 1) = "V" Then szt$ = " vev�i meg�llapod�s" Else szt$ = " sz�ll�t�i meg�llapod�s"
      If Mid$(rec$, 62, 1) = "I" Then
        Tbknev.List1.ForeColor = QBColor(3)
        Tbknev.List1.AddItem "�l�" + szt$
      Else
        Tbknev.List1.ForeColor = QBColor(4)
        Tbknev.List1.AddItem "Passz�v" + szt$
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
      Tbknev.List1.AddItem "Spec.�r:  " + Mid(rec$, 50, 12)
      Tbknev.List1.AddItem "Spec.engedm�ny: " + Mid(rec$, 44, 6) + " %"
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Hat�lyba l�p�s: " + datki(Mid$(rec$, 32, 6))
      Tbknev.List1.AddItem "Lej�rat:        " + datki(Mid$(rec$, 32, 6))
    Case "ARMG"
      Tbknev.List1.Clear
      If Mid$(rec$, 46, 1) = "A" Then
        Tbknev.List1.ForeColor = QBColor(3)
        Tbknev.List1.AddItem "Akt�v"
      Else
        Tbknev.List1.ForeColor = QBColor(4)
        Tbknev.List1.AddItem "Passz�v"
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
          Tbknev.List1.AddItem "Spec.�r:  " + Mid(elem$, 15, 12)
          Tbknev.List1.AddItem "Spec.engedm�ny: " + Mid(elem$, 27, 6) + " %"
          Tbknev.List1.AddItem "Hat�lyba l�p�s: " + datki(Mid$(elem$, 1, 6))
          Tbknev.List1.AddItem "Lej�rat:        " + datki(Mid$(elem$, 7, 6))
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
          Tbknev.List1.AddItem "Fogyaszt�i �r:" + Mid$(qrec$, 680, 12)
          If Trim(Mid$(qrec$, 1097, 6)) <> "" Then
            Tbknev.List1.AddItem "�rv�ltoz�s kelte: " + datki(Mid$(qrec$, 1097, 6))
          End If
          If Mid$(qrec$, 1248, 1) = "I" Then
            Tbknev.List1.AddItem "Helyi �r"
          Else
            Tbknev.List1.AddItem "K�zponti �r"
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
        Tbknev.List1.AddItem "Bels� :" + Trim(Mid$(rakrec$, 5, 60))
      Else
        Tbknev.List1.AddItem "Fizet�:" + Trim(Mid$(partrec$, 15, 60))
      End If
      Select Case Mid$(rec$, 199, 1)
        Case "L": Tbknev.List1.AddItem "Biz: sz�ll�t�lev�l"
        Case "S": Tbknev.List1.AddItem "Biz: sz�mla"
        Case "E": Tbknev.List1.AddItem "Biz: export sz�mla"
        Case "R": Tbknev.List1.AddItem "Biz: rakt�rk�zi sz�ll�t�lev�l"
        Case Else
      End Select
      Tbknev.List1.AddItem " "
      If partrec$ <> "" Then
        Tbknev.List1.AddItem "Ad�sz�m     :" + Mid$(partrec$, 184, 15)
      End If
      Tbknev.List1.AddItem "Eng.sz�ma   :" + Mid$(rec$, 277, 1) + "-" + Trim(Mid$(rec$, 278, 20))
      If Trim(Mid$(rec$, 208, 4)) <> "" Then
        Tbknev.List1.AddItem "R�gi�     :" + Trim(torolvas("RGIO", Mid$(rec$, 208, 4), 5, 60))
      Else
        If partrec$ <> "" Then
          Tbknev.List1.AddItem "R�gi�     :" + Trim(torolvas("RGIO", Mid$(partrec$, 315, 4), 5, 60))
        End If
      End If
      If Trim(Mid$(rec$, 200, 8)) <> "" Then
        Tbknev.List1.AddItem "�zletk�t� :" + Trim(torolvas("KUZL", Mid$(rec$, 200, 8), 8, 60))
      Else
        If partrec$ <> "" Then
          Tbknev.List1.AddItem "�zletk�t� :" + Trim(torolvas("KUZL", Mid$(partrec$, 319, 8), 8, 60))
        End If
      End If
      If partrec$ <> "" Then
        Tbknev.List1.AddItem "Fizm.m�d  :" + Trim(torolvas("PFIZ", Mid$(partrec$, 328, 2), 3, 30)) + " " + Mid$(partrec$, 330, 3) + " nap"
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Tartoz�s  :" + ertszamx(Mid$(partrec$, 659, 14), 16, 2)
        If Mid$(partrec$, 327, 1) = "T" Then
          Tbknev.List1.AddItem "Partner letiltva!!!"
          Tbknev.List1.ForeColor = QBColor(12)
        Else
          If Mid$(partrec$, 327, 1) = "K" Then
            Tbknev.List1.AddItem "Csak k�szp�nre!!!"
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
        Case "BK": Tbknev.List1.AddItem "Belf�ldi kereskedelmi partner"
        Case "KE": Tbknev.List1.AddItem "K�z�ss�gi partner"
        Case "KG": Tbknev.List1.AddItem "Egy�b k�lf�ldi partner"
        Case "PH": Tbknev.List1.AddItem "P�nz�gyi hat�s�g"
        Case "PG": Tbknev.List1.AddItem "Egy�b p�nz�gyi partner"
        Case Else
      End Select
      Tbknev.List1.AddItem "Ad�sz�m     :" + Mid$(rec$, 184, 15)
      Tbknev.List1.AddItem "K�z.ad�sz�m :" + Mid$(rec$, 199, 15)
      Tbknev.List1.AddItem "Eng.sz�ma   :" + Mid$(rec$, 802, 1) + "-" + Trim(Mid$(rec$, 803, 20))
      Tbknev.List1.AddItem banktagol(Mid$(rec$, 244, 24))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Vev� kat. :" + Trim(torolvas("VKAT", Mid$(rec$, 314, 1), 2, 60))
      Tbknev.List1.AddItem "�rkateg.  :" + Mid$(rec$, 782, 1)
      Tbknev.List1.AddItem "R�gi�     :" + Trim(torolvas("RGIO", Mid$(rec$, 315, 4), 5, 60))
      Tbknev.List1.AddItem "�zletk�t� :" + Trim(torolvas("KUZL", Mid$(rec$, 319, 8), 9, 60))
      Tbknev.List1.AddItem "Fizm.m�d  :" + Trim(torolvas("PFIZ", Mid$(rec$, 328, 2), 3, 30)) + " " + Mid$(rec$, 330, 3) + " nap"
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Tartoz�s  :" + ertszamx(Mid$(rec$, 659, 14), 16, 2)
      If Mid$(rec$, 327, 1) = "T" Then
        Tbknev.List1.AddItem "Letiltva!!!"
        Tbknev.List1.ForeColor = QBColor(12)
      Else
        If Mid$(rec$, 327, 1) = "K" Then
          Tbknev.List1.AddItem "Csak k�szp�nzre!!!"
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
        Case "V": Tbknev.List1.AddItem "Vev�t�l kapott el�leg"
        Case "S": Tbknev.List1.AddItem "Sz�ll�t�nak adott el�leg"
      End Select
      Select Case Mid$(rec$, 224, 1)
        Case "E": Tbknev.List1.AddItem "El�leg fizet�s"
        Case "V": Tbknev.List1.AddItem "El�leg visszafizet�s"
        Case "B": Tbknev.List1.AddItem "El�leg besz�m�t�s"
      End Select
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Sz�mlasz�m:" + Mid$(rec$, 8, 15)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Teljes�t�s kelte  :" + datki(Mid$(rec$, 38, 6))
      Tbknev.List1.AddItem "P�nzt�ri iktat�   :" + datki(Mid$(rec$, 38, 6))
      Tbknev.List1.AddItem "Banki iktat�      :" + datki(Mid$(rec$, 38, 6))
      Tbknev.List1.AddItem "K�nyvel�si iktat� :" + datki(Mid$(rec$, 38, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "El�leg �sszege  :" + ertszamx(Mid$(rec$, 44, 14), 14, 2)
      Tbknev.List1.AddItem "Besz�m.�sszeg   :" + ertszamx(Mid$(rec$, 134, 14), 14, 2)
      Tbknev.List1.AddItem "Visszafiz.�sszeg:" + ertszamx(Mid$(rec$, 148, 14), 14, 2)
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 90, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Sztorn�zva:" + datki(Mid$(rec$, 91, 6)) + " " + Mid$(rec$, 97, 8) + " "
      End If
    Case "PVSZ", "PSSZ"
      If objneve$ = "PVSZ" Then vsmod$ = "V" Else vsmod$ = "S"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      pkod$ = Mid$(rec$, 38, 15): If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod) Else partrec$ = ""
      If partrec$ <> "" Then Tbknev.List1.AddItem Trim(Mid$(partrec$, 16, 60))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Sz�mlasz�m:" + Mid$(rec$, 8, 15)
      Tbknev.List1.AddItem "Hiv.sz�m  :" + Mid$(rec$, 23, 15)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Teljes�t�s kelte :" + datki(Mid$(rec$, 211, 6))
      Tbknev.List1.AddItem "K�nyvel�s kelte  :" + datki(Mid$(rec$, 58, 6))
      Tbknev.List1.AddItem "Sz�mla kelte     :" + datki(Mid$(rec$, 64, 6))
      Tbknev.List1.AddItem "Fizet�si hat�rid�:" + datki(Mid$(rec$, 70, 6))
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "�sszeg  :" + ertszamx(Mid$(rec$, 78, 14), 14, 2)
      egyenleg@ = szegyen(rec$, maidatum$, maidatum$)
      If egyenleg@ <> 0 Then Call szamlaegyenleg(rec$, ossz@, helybit@, kiegy@, egyenleg@, vsmod$, maidatum, maidatum, forintegyenleg@)
      Tbknev.List1.AddItem "Egyenleg:" + ertszamx(Str(egyenleg), 14, 2)
      Tbknev.List1.AddItem " "
      If Mid$(rec$, 166, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Sztorn�zva:" + datki(Mid$(rec$, 167, 6)) + " " + Mid$(rec$, 181, 8) + " "
        Tbknev.List1.AddItem "Szt.sz�mla:" + Mid$(rec$, 220, 15) + " "
      End If
    Case "PBNK"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      If Mid$(rec$, 8, 1) = "T" Then vr$ = "Terhel�s" Else vr$ = "J�v��r�s"
      Tbknev.List1.AddItem vr$
      fksz$ = Mid$(rec$, 22, 8): fkszrec$ = dbxkey("FKSZ", fksz)
      Tbknev.List1.AddItem Mid$(rec$, 240, 30)
      Tbknev.List1.AddItem datki(Mid$(rec$, 16, 6)) + " " + Trim(Mid$(fkszrec$, 9, 60))
      Select Case Mid$(rec$, 55, 1)
        Case "V": Tbknev.List1.AddItem "Vev� sz�mla kiegy. :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "S": Tbknev.List1.AddItem "Sz�ll.sz�mla kiegy.:" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "E": Tbknev.List1.AddItem "El�leg vev�t�l     :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "L": Tbknev.List1.AddItem "El�leg sz�ll�t�nak :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "M": Tbknev.List1.AddItem "Elsz�mol�si el�leg :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "X": Tbknev.List1.AddItem "Egy�b              :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
      End Select
      Tbknev.List1.AddItem " "
      pkod$ = Mid$(rec$, 108, 15): If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod) Else partrec$ = ""
      If partrec$ <> "" Then Tbknev.List1.AddItem Trim(Mid$(partrec$, 16, 60))
      Tbknev.List1.AddItem Mid$(rec$, 30, 25)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Sz�mlasz�m  :" + Mid$(rec$, 123, 15)
      Tbknev.List1.AddItem "Ellensz�mla :" + Trim(Mid$(rec$, 146, 8)) + " Kt.:" + Mid$(rec$, 207, 7)
      Tbknev.List1.AddItem "Szerv.egys�g:" + Mid$(rec$, 162, 8)
      Tbknev.List1.AddItem "Munkasz�m   :" + Mid$(rec$, 170, 8)
      If Mid$(rec$, 192, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Sztorn�zva:" + datki(Mid$(rec$, 193, 6)) + " " + Mid$(rec$, 199, 8) + " "
      End If
    Case "PKTE"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      If Mid$(rec$, 8, 1) = "B" Then vr$ = Mid$(rec$, 9, 7) + ".sz. bev�teli bizonylat" Else vr$ = Mid$(rec$, 9, 7) + ".sz. kiad�si bizonylat"
      Tbknev.List1.AddItem vr$
      fksz$ = Mid$(rec$, 22, 8): fkszrec$ = dbxkey("FKSZ", fksz)
      Tbknev.List1.AddItem Mid$(rec$, 240, 30)
      Tbknev.List1.AddItem datki(Mid$(rec$, 16, 6)) + " " + Trim(Mid$(fkszrec$, 9, 60))
      Select Case Mid$(rec$, 55, 1)
        Case "V": Tbknev.List1.AddItem "Vev� sz�mla kiegy. :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "S": Tbknev.List1.AddItem "Sz�ll.sz�mla kiegy.:" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "E": Tbknev.List1.AddItem "El�leg vev�t�l     :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "L": Tbknev.List1.AddItem "El�leg sz�ll�t�nak :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "M": Tbknev.List1.AddItem "Elsz�mol�si el�leg :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
        Case "X": Tbknev.List1.AddItem "Egy�b              :" + ertszam(Mid$(rec$, 56, 14), 12, 2)
      End Select
      Tbknev.List1.AddItem " "
      pkod$ = Mid$(rec$, 108, 15): If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod) Else partrec$ = ""
      If partrec$ <> "" Then Tbknev.List1.AddItem Trim(Mid$(partrec$, 16, 60))
      Tbknev.List1.AddItem Mid$(rec$, 30, 25)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Ellensz�mla :" + Trim(Mid$(rec$, 146, 8)) + " Kt.:" + Mid$(rec$, 207, 7)
      Tbknev.List1.AddItem "Szerv.egys�g:" + Mid$(rec$, 162, 8)
      Tbknev.List1.AddItem "Munkasz�m   :" + Mid$(rec$, 170, 8)
      If Mid$(rec$, 192, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Sztorn�zva:" + datki(Mid$(rec$, 193, 6)) + " " + Mid$(rec$, 199, 8) + " "
      End If
    Case "PKOR"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      Select Case Mid$(rec$, 14, 1)
        Case "V": Tbknev.List1.AddItem "Vev� korrekci�:" + ertszam(Mid$(rec$, 45, 14), 12, 2): Tbknev.List1.ForeColor = RGB(0, 150, 0)
        Case "S": Tbknev.List1.AddItem "Sz�ll�t� korrekci�:" + ertszam(Mid$(rec$, 45, 14), 12, 2): Tbknev.List1.ForeColor = RGB(0, 0, 150)
        Case "K": Tbknev.List1.AddItem "Kompenz�ci�:" + ertszam(Mid$(rec$, 45, 14), 12, 2): Tbknev.List1.ForeColor = RGB(50, 50, 50)
      End Select
      If Mid$(rec$, 15, 1) = "I" Then
        Tbknev.List1.AddItem "�FA alapot m�dos�t."
      Else
        Tbknev.List1.AddItem "�FA alapot nem m�dos�t."
      End If
      Tbknev.List1.AddItem " "
      pkod$ = Mid$(rec$, 16, 15): If Trim(pkod$) <> "" Then partrec$ = dbxkey("PART", pkod) Else partrec$ = ""
      If partrec$ <> "" Then Tbknev.List1.AddItem Trim(Mid$(partrec$, 16, 60))
      Tbknev.List1.AddItem " "
      vikt$ = Trim(Mid$(rec$, 31, 7)): sikt$ = Trim(Mid$(rec$, 38, 7))
      tikt$ = Mid$(rec$, 356, 7)
      If vikt$ <> "" Then
        virec$ = dbxkey("PVSZ", vikt$)
        Tbknev.List1.AddItem "Vev� sz�mla :" + vikt$ + "  " + Trim(Mid$(virec$, 8, 15))
        Tbknev.List1.AddItem "K�nyv.t�tel :" + tikt$ + " "
      End If
      If sikt$ <> "" Then
        sirec$ = dbxkey("PSSZ", sikt$)
        Tbknev.List1.AddItem "Sz�ll.sz�mla:" + sikt$ + "  " + Trim(Mid$(sirec$, 8, 15))
        If Mid$(rec$, 14, 1) = "S" Then
          Tbknev.List1.AddItem "K�nyv.t�tel :" + tikt$ + " "
        Else
          tik1& = xval(tikt$)
          tikt1$ = Right("0000000" + Trim(Str(tik1&)), 7)
          Tbknev.List1.AddItem "K�nyv.t�tel :" + tikt1$ + " "
        End If
      End If
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem Mid$(rec$, 76, 30)
      Tbknev.List1.AddItem " "
      Tbknev.List1.AddItem "Ellensz�mla :" + Trim(Mid$(rec$, 324, 8))
      Tbknev.List1.AddItem "Szerv.egys�g:" + Mid$(rec$, 340, 8)
      Tbknev.List1.AddItem "Munkasz�m   :" + Mid$(rec$, 348, 8)
      If Mid$(rec$, 120, 1) = "S" Then
        Tbknev.List1.ForeColor = QBColor(12)
        Tbknev.List1.AddItem " "
        Tbknev.List1.AddItem "Sztorn�zva:" + datki(Mid$(rec$, 121, 6)) + " " + Mid$(rec$, 127, 8) + " "
      End If
    Case "JTRM"
      Tbknev.List1.Clear
      Tbknev.List1.ForeColor = QBColor(3)
      Tbknev.List1.AddItem Mid$(rec$, 90, 40)
      Tbknev.List1.AddItem "�vj: " + Mid$(rec$, 76, 4) + " Fok: " + Trim(Mid$(rec$, 5960, 10))
      Select Case Mid$(rec$, 286, 1)
        Case "B": Tbknev.List1.AddItem "Sz�l�bor"
        Case "A": Tbknev.List1.AddItem "Asz�"
        Case "S": Tbknev.List1.AddItem "S�r"
        Case "R": Tbknev.List1.AddItem "Egy�b bor"
        Case "E": Tbknev.List1.AddItem "Egy�b alkoholterm�k"
        Case Else: Tbknev.List1.AddItem "Jelleg hi�nyzik"
      End Select
      Tbknev.List1.AddItem "Palackozott Vtsz: " + Mid$(rec$, 142, 12)
      Tbknev.List1.AddItem "L�dig Vtsz      : " + Mid$(rec$, 142, 12)
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
            '--- komissi�z�
            keszle@ = komkeszle@ - fogkeszle@
          Case 2
            '--- c+c
            keszle@ = cpckeszle@ + vegkeszle@
          Case 3
            '--- vegyi
            keszle@ = vegkeszle@
          Case 4
            '--- z�lds�g
            keszle@ = zolkeszle@
          Case 5
            '--- g�ngy�leg
            keszle@ = gonkeszle@
          Case Else
            '--- nincs be�ll�tva
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
        Tbknev.List1.AddItem "K�szlet:" + ertszam(Str(keszle@), 14, 2) + " " + Mid$(rec$, 484, 6)
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
          Tbknev.List1.AddItem "K�szlet:" + ertszam(Str(keszle@), 14, 2) + " " + Mid$(rec$, 484, 6)
          Tbknev.List1.AddItem "N e g a t � v  k � s z l e t"
          Tbknev.List1.ForeColor = QBColor(12)
          Tbknev.List1.AddItem " "
        Else
          Tbknev.List1.AddItem "K�szlet: nincs eladhat� k�szlet"
          Tbknev.List1.ForeColor = QBColor(4)
          Tbknev.List1.AddItem " "
        End If
      End If
      Tbknev.List1.AddItem "Ut.besz.�r:" + Mid$(rec$, 568, 12)
      Tbknev.List1.AddItem "Ref.�r    :" + Mid$(rec$, 1276, 12)
      Tbknev.List1.AddItem "�r1:" + Mid$(rec$, 582, 12) + " �r5:" + Mid$(rec$, 638, 12)
      Tbknev.List1.AddItem "�r2:" + Mid$(rec$, 596, 12) + " �r6:" + Mid$(rec$, 652, 12)
      Tbknev.List1.AddItem "�r3:" + Mid$(rec$, 610, 12) + " �r7:" + Mid$(rec$, 666, 12)
      Tbknev.List1.AddItem "�r4:" + Mid$(rec$, 624, 12)
      Tbknev.List1.AddItem "Disztr.�r:" + Mid$(rec$, 895, 12)
      Tbknev.List1.AddItem "Kisker.�r:" + Mid$(rec$, 680, 12)
      Tbknev.List1.AddItem " "
      If kozkeszle <> 0 Then Tbknev.List1.AddItem "K�zp.rakt:" + ertszam(Str(kozkeszle@), 12, 2) + " " + Mid$(rec$, 484, 6)
      If cpckeszle <> 0 Then Tbknev.List1.AddItem "C+C      :" + ertszam(Str(cpckeszle@), 12, 2) + " " + Mid$(rec$, 484, 6)
      If vegkeszle <> 0 Then Tbknev.List1.AddItem "Vegyi rak:" + ertszam(Str(vegkeszle@), 12, 2) + " " + Mid$(rec$, 484, 6)
      If gonkeszle <> 0 Then Tbknev.List1.AddItem "G�ngy.rak:" + ertszam(Str(gonkeszle@), 12, 2) + " " + Mid$(rec$, 484, 6)
      If zolkeszle <> 0 Then Tbknev.List1.AddItem "Z�lds.rak:" + ertszam(Str(zolkeszle@), 12, 2) + " " + Mid$(rec$, 484, 6)
    Case Else
      Call infomutat2(objneve$, rec$, hivashely$)
  End Select
End Sub

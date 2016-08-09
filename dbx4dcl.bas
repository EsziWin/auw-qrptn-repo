Attribute VB_Name = "dbx4dcl"
'--- DBX4 daklarációk és közös szubrutinok
'--- Adatbázis leírások
'--- Adatbázis táblázat
Public DBXTAB(20) As String        'DBX fájlok nevei
Public dbxdb As Integer            'DBXTAB elemszám
'--- Objektum táblázat
Type OBTAB                         'objektum elem
  obaz As String * 4               'objektum azonosító
  obnev As String * 40             'objektum fejléc
  dbxindex As Integer              'DBXTAB-ra mutató sorszám
  rekhossz As Integer              'logikai rekordhossz
  pointdb As Integer               'pointerek száma
  obi(10) As Integer               'indexek INDTAB-ra mutató sorszám, obi(0)=darabszám
  oba(10) As Integer               'ablakok ABLTAB-ra mutató sorszám, oba(0)=darabszám
  reszam As Long                   'rekordok száma
  renszam As Long                  'rendezett rekordok száma (elsõdleges terület)
  rekfirst As Long                 'elsõ rekord címe
  reklast As Long                  'utolsó rekord címe
  iktlast As Long                  'utolsó kiadott iktató
  obcim As Long                    'aktuális rekord címe
  obind As Long                    'aktuális rekord fõindex beli sorszáma
  obnex As Long                    'következõ rekord fõpointer szerinti címe, ha nulla EOF
  icccim As Long                   'icc elem címe az icc fájlban
  hashcod As Integer               'hash code adatab-ra mutatója
  hashmez(10) As Integer           'a hash keresõben megmutatandó mezok kódja
  keresokoddarab As Integer        'az adott objektumra hány keresõkódot kell építeni
End Type
Public OBJTAB(120) As OBTAB         'objektum tábla
Public objdb As Integer            'OBJTAB elemszám
'--- Ablak táblázat
Type ABTAB                         'ablak elem
  fejlec As String * 40            'ablak fejléce
  abmod As String * 1              '0-rejtett 1-adatbeviteli
  adatsorsz(110) As Integer         'ADATAB-ra mutató sorszámok, adatsorsz(0)=darabszám
End Type
Public ABLTAB(200) As ABTAB         'ablak tábla
Public abldb As Integer            'ABLTAB elemszám
'--- Index táblázat
Type IXTAB                         'index elem
  indnev As String * 8             'indexfájl neve
  obsorsz As Integer               'OBJTAB-ra mutató sorszám
  adatsorsz As Integer             'ADATAB-ra mutató sorszám
End Type
Public INDTAB(100) As IXTAB         'index táblázat
Public inddb As Integer            'INDTAB elemszám
Type STTAB
  setnev As String * 8             'set azonositó kód
  obsorsz As Integer               'member objektum sorszáma
  robsorsz As Integer              'owner objektum sorszáma
  adatsorsz As Integer             'kapcsolatot definiáló adat kódja (memberben)
  rootpoz As Integer               'startpointer kezdõpoz.ownerben
  nextpoz As Integer               'nextpointer kezdõpoz.memberben
  priorpoz As Integer              'priorpointer kezdõpoz.memberben
  owneraz As String * 4            'az owner objektum azonosítója
  memberaz As String * 4           'a member objektum azonosítója
End Type
Public SETTAB(80) As STTAB         'owner-member setek táblázata
Public setdb As Integer            'SETTAB elemszám
'--- Adatmezõ táblázat
Type ADTAB                         'adatmezõ elem
  adkod As String * 8              'adatkód
  obsorsz As Integer               'OBJTAB-ra mutat
  absorsz As Integer               'ABLTAB-ra mutat
  adatnev As String * 25           'adatmezõ megnevezése
  adatkp As Integer                'kezdõ pozíció a rekordban
  adatho As Integer                'adatmezõ hossza
  attr As String * 10              'rögzítési atribútumok
  kapcsobnev As String * 4         'kapcsolódó objektum neve
  kapcsob As Integer               'kapcsolódó OBJTAB-ra mutat, 0=nincs kapcsolat
  kapcsdbx As Integer              'kapcsolódó DBXTAB-ra mutat
  keresokodindex As Integer        'az objektumon belül hányadik keresõ kód
  keresokodhossz As Integer        'az keresokod hossza
End Type
Public MAGYARAZAT$(1500)
'--- Magyarázó szövegek az adatokhoz
Public ELLENORZO$(1500)
'--- Ellenõrzõ feltételek az adatokhoz
Public ADATAB(1500) As ADTAB        'adatmezõ táblázat
Public adadb As Integer            'ADATAB elemszám
Public utrec$(100)                  'utolsó beolvasott rekord objektumonként
Public keysorszam&
Public halprin%
'--- PRX szerkezete
'---   DDF=adatkód&adatnév&attribútum&magyarázat (nem kötelezõ)
'---   OBJ=objazon&megnevezés&rekhossz&pointerszám
'---     ABL=fejléc&mód(rejtett vagy látható)
'---       MEZ=adatkód&név&kezdet&hossz&attrib&kapcsdbx&kapcsobj
'---     IND=indexfájlnév&adatkód
'--- Fizikai struktúra
'--- DBX  1,4 objektum prefix
'---      5,1 törlõjel ;=élõ *=törölt
'---      6,4 fõpointer
'---     10,  a logikai rekord
'--- NDX  1,4 adatbázis cím
'---      5,n index elem
'---      5+n,1 törlõjel ;=élõ *=törölt
'--- ICC  1,4 objektum prefix
'---      5,4 rekordok száma
'---      9,4 rendezett rekordok száma
'---     13,4 elsõ rekord pointere
'---     17,4 utolsó rekord pointere
'---     21,4 utoljára kiadott iktató
'--- Adatbázis struktúra vége
Type komtab
  komkod As String * 1      '--- 1=adat 2=intervallum 3=menu
  komszov As String * 500   '--- adat neve (1,2) menuszoveg (3) esetén
  komatr As String * 10     '--- attributum (1,2) esetén
  komtol As String * 20     '--- kezdo ertek
  komig  As String * 20     '--- zaro ertek
  kommnv As Integer         '--- menuválasztás
  komdbx As String * 8      '--- kapcsolódó adatbázis
  komobj As String * 4      '--- kapcsolódó objektum
End Type
Public komt(30) As komtab   '--- kommunikációs paraméterek
Public komdb As Integer     '--- kom.par darabszáma
Public kommenudb%, komadatdb%
Public mt$(1001), mtb$(1001), mho%(1001), mesor%(1001), mkp%(1001)
Public nt$(1001)
Public kmt$(1001), kmtb$(1001), kmho%(1001), kmkp%(1001), kmesor%(1001), kmehsx%(1001), kmehsxh%(1001), kmnumer%(1001)
Public lckfi, lckulcsfi, rogzites%, rendpar$
Public sortorol%, torstart&, foindex%
Public tol@, ig@, keress%, gombs%, sztart%
Public wwtoprow&, wwaktrow&, wwkezdpo&, wwzarpo&, wwaktpo&, wwtabstatusz%
Public toprow&, aktrow&, keresomod%, kezdpo&, zarpo&, aktpo&, tabstatusz%, aktdarab&, aktext$
Public toprox&, aktrox&, tabstatusx%, resx&(2000), resxdb&
Public resm&(500), resmdb%, vsorszam&, maidatum$
Public keresobj$, gombsorszam%, talalat%, rekord$, kommegsem%
Public mn%, men$, bo%, bs%, termfelold%, utafakod$, utafakulcs@
Public terminal$, task$, ugyintezo$, regszam$, auditorutvonal$, listautvonal$, programutvonal$
Public programnev$, objektum$, ablaksorszam%, ugyintneve$, jogok$, penztarirany$
Public rootutvonal$, licensztulaj$, cegneve$, munkautvonal$
Public afosz@, marosz@, irec$, ikonfrec$, lezardat$, targyev$, pzbank$, partrec$, erteknap$, devizabank$
Public abizo$, abizrec$, betukod$, abizptmod$, tablamaszk$
Public kulcstomb%(14, 2), epoztomb%(14, 2), form852%
Public rakrec$, rendezohiba%, igen123%, szamlaeset%, althobjektum$, altxobjektum$, cimvektorvan%, megmodos%
Public autoinfo%, ertulcs&, ugyintrec$, kerlezardat$, zkermod%
Public raktartipus% '--- 0-nincs 1-komissiózó 2-c+c 3-vegyi 4-zöldség 5-göngyöleg
Public reszjog$(300, 2), reszjogdarab&, jparam$(200), jparamdb%
Public unijog%, dbx3mod%

Public Sub joginit()
  '--- részfunkció jogok beolvasása
  fi55 = FreeFile
  Open Left(auditorutvonal$, 1) + ":\auwin\jogok.par" For Binary Shared As #fi55
  fim55& = LOF(fi55)
  Close fi55
  If fim55& < 5 Then
    reszjogdarab& = 0
  Else
    reszjogdarab& = 0
    fi55 = FreeFile
    Open Left(auditorutvonal$, 1) + ":\auwin\jogok.par" For Input Shared As #fi55
    Do
      Line Input #fi55, jax$
      pzj% = InStr(jax$, "=")
      If pzj% > 1 Then
        bal$ = UCase(Trim(Left(jax$, pzj% - 1)))
        jobb$ = UCase(Trim(Mid$(jax$, pzj% + 1)))
        If bal$ <> "" And jobb$ <> "" Then
          Call linpar(jobb$, jparam(), ",", jparamdb%)
          If jparamdb% > 0 Then
            For i882% = 1 To jparamdb%
              job$ = Trim(jparam(i882%))
              If job$ <> "" Then
                reszjogdarab& = reszjogdarab& + 1
                reszjog$(reszjogdarab, 1) = bal$
                reszjog$(reszjogdarab, 2) = job$
              End If
            Next
          End If
        End If
      End If
    Loop While Not EOF(fi55)
    Close fi55
  End If
End Sub

Public Function jogellen%(kulcs$)
  '--- részfunkciók ellenõrzése
  If reszjogdarab& = 0 Then jogellen = 1: Exit Function
  For i881% = 1 To reszjogdarab&
    If UCase(kulcs$) = reszjog$(i881%, 1) And UCase(Trim(ugyintezo$)) = reszjog$(i881%, 2) Then jogellen = 1: Exit Function
  Next
  jogellen = 0
End Function

Public Sub dbxki(objazon, rec$, tjel$, umod$, gener$, hiba%)
  '--- adatbázisrekord írása
  '--- tjel$=";" élõ, tjel$="*" törlés
  '--- umod$="U" új rekord, egyébként visszaírás
  '--- gener$="G" generált iktatóval, egyébként sima rekord
  '--- be: hiba%=0 nincs ellenõzés, hiba%=1 ellenõrzés
  '--- ki: hiba%=0 az írás sikeres volt, hiba%=1 nem történt írás
  '--- 2002.11.23
  If rec$ = "" Then
    hiba% = 1: Exit Sub
  End If
  iras% = 1: Hashkod% = 0
  ob% = obsorszama(objazon)
  If OBJTAB(ob%).hashcod <> 0 Then
    hsp% = ADATAB(OBJTAB(ob%).hashcod).adatkp
    Hashkod% = 1
    cod1$ = xkonver(Mid$(rec$, hsp%, 1))
    cod2$ = xkonver(Mid$(rec$, hsp% + 1, 1))
    cod3$ = xkonver(Mid$(rec$, hsp% + 2, 1))
  End If
  dbxn$ = dbxneve(objazon)
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  lofilmeret& = LOF(dbfi)
  Close dbfi
  dbxfilmeret& = FileLen(auditorutvonal$ + dbxn$ + ".dbx")
  If lofilmeret& > dbxfilmeret& Then dbxfilmeret& = lofilmeret&
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  icfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".icc" For Binary Shared As #icfi
  icim& = OBJTAB(ob%).icccim
  If gener$ = "G" Then
    '--- új iktató megállapítása
    '--- ideiglenesen kikapcsolva tesztelés céljából
    w1% = OBJTAB(ob%).obi(1)
    w2% = INDTAB(w1%).adatsorsz
    kp% = ADATAB(w2%).adatkp
    ho% = ADATAB(w2%).adatho
    Get #icfi, icim& + 20, Ikt&
    Ikt& = Ikt& + 1
    ik$ = Right$("0000000" + LTrim$(Str$(Ikt&)), 7)
    Mid$(rec$, kp%, ho%) = ik$
  End If
  If umod$ = "U" Then
    '--- új rekordcim
    '--- tesztelés 2004.01.11
    'rcim& = LOF(dbfi) + 1
    cmvfi = FreeFile
    Open auditorutvonal$ + objazon + ".cmv" For Append As #cmvfi
    rcim& = dbxfilmeret& + 1
  Else
    '--- utoljára beolvasott rekord címe
    rcim& = OBJTAB(ob%).obcim
    If rcim& > 0 Then
      If hiba% = 1 Then
        '--- commit ellenõrzés
        Seek #dbfi, rcim& + 9
        rh& = OBJTAB(ob%).rekhossz
        rrec$ = Space(rh&): Get #dbfi, , rrec$
        If rrec$ <> utrec$(ob%) Then
          iras% = 0
        End If
        If rrec$ = rec$ Then
          Close dbfi: Close icfi
          hiba% = 0
          Exit Sub
        End If
      Else
        If tjel$ <> "*" Then
          Seek #dbfi, rcim& + 9
          rh& = OBJTAB(ob%).rekhossz
          rrec$ = Space(rh&): Get #dbfi, , rrec$
          If rrec$ = rec$ Then
            Close dbfi: Close icfi
            Exit Sub
          End If
        End If
      End If
    Else
      Close dbfi: Close icfi
      Exit Sub
    End If
  End If
  If iras% = 1 Then
    hiba% = 0
    '--- rekord felírása
    '--- írás az eseménynaplóba
    If objazon = "JFRG" Or igen123% = 1 Then
      lgfi = FreeFile
      Open auditorutvonal$ + dbxn$ + ".123" For Append As #lgfi
      z$ = objazon + tjel$ + terminal$ + task$ + ugyintezo$ + Date$ + Time$
      Print #lgfi, z$;
      Print #lgfi, rec$
      Close lgfi
    End If
    If umod$ = "U" Then
      '--- új írás
      Get #icfi, icim& + 12, elso&
      z$ = objazon + ";"
      Put #dbfi, rcim&, z$
      Put #dbfi, rcim& + 5, elso&
      Put #dbfi, rcim& + 9, rec$
      '--- icc módosítása
      Put #icfi, icim& + 12, rcim&
      cicike$ = Trim(Str(rcim&))
      Print #cmvfi, cicike$
      If umod$ = "U" Then
        Get #icfi, icim& + 4, rc&
        rc& = rc& + 1
        Put #icfi, icim& + 4, rc&
      End If
      If gener$ = "G" Then
        Put #icfi, icim& + 20, Ikt&
      End If
    Else
      '--- visszaírás
      If tjel$ = "*" Then
        Put #dbfi, rcim& + 4, tjel$
      Else
        z$ = objazon + ";"
        Put #dbfi, rcim&, z$
        rx& = OBJTAB(ob%).obnex
        Put #dbfi, rcim& + 5, rx&
        Put #dbfi, rcim& + 9, rec$
      End If
    End If
    '--- indexek felírása
    For i5% = 1 To OBJTAB(ob%).obi(0)
      w9% = OBJTAB(ob%).obi(i5%)
      inev$ = Trim$(INDTAB(w9%).indnev)
      ndfi = FreeFile
      Open auditorutvonal$ + inev$ + ".ndx" For Binary Shared As #ndfi
      Close ndfi
      ndxfilmeret& = FileLen(auditorutvonal$ + inev$ + ".ndx")
      ndfi = FreeFile
      Open auditorutvonal$ + inev$ + ".ndx" For Binary Shared As #ndfi
      w1% = OBJTAB(ob%).obi(i5%)
      w2% = INDTAB(w1%).adatsorsz
      kp% = ADATAB(w2%).adatkp
      ho% = ADATAB(w2%).adatho
      indbej$ = Mid$(rec$, kp%, ho%)
      If umod$ = "U" Then
        '--- 2004.01.11 tesztelés
        'incim& = LOF(ndfi) + 1
        'insor& = Int(LOF(ndfi) / (ho% + 5)) + 1
        incim& = ndxfilmeret& + 1
        insor& = Int(ndxfilmeret& / (ho% + 5)) + 1
      Else
        poz& = OBJTAB(ob%).obind
        incim& = (poz& - 1) * (ho% + 5) + 1
      End If
      Put #ndfi, incim&, rcim&
      xx$ = indbej$ + tjel$
      Put #ndfi, incim& + 4, xx$
      Close ndfi
    Next
    '--- hash kódok felírása
    If Hashkod% = 1 Then
      hs1fi = FreeFile
      Open auditorutvonal$ + "auw-" + objazon + ".hs1" For Binary Shared As #hs1fi
      Close hs1fi
      hs1fi = FreeFile
      hs1filmeret& = FileLen(auditorutvonal$ + "auw-" + objazon + ".hs1")
      hs1fi = FreeFile
      Open auditorutvonal$ + "auw-" + objazon + ".hs1" For Binary Shared As #hs1fi
      hs2fi = FreeFile
      Open auditorutvonal$ + "auw-" + objazon + ".hs2" For Binary Shared As #hs2fi
      hs3fi = FreeFile
      Open auditorutvonal$ + "auw-" + objazon + ".hs3" For Binary Shared As #hs3fi
      If umod$ = "U" Then
        '--- tesztelés 2004.01.11
        'poz& = LOF(hs1fi) + 1
        poz& = hs1filmeret& + 1
      Else
        poz& = OBJTAB(ob%).obind
      End If
      Put #hs1fi, poz&, cod1$
      Put #hs2fi, poz&, cod2$
      Put #hs3fi, poz&, cod3$
      Close hs1fi: Close hs2fi: Close hs3fi
    End If
    '--- objtab módosítása
    utrec$(ob%) = rec$
    If umod$ = "U" Then
      OBJTAB(ob%).obcim = rcim&
      OBJTAB(ob%).obind = insor&
      OBJTAB(ob%).obnex = elso&
      OBJTAB(ob%).reszam = OBJTAB(ob%).reszam + 1
      If gener$ = "G" Then OBJTAB(ob%).iktlast = Ikt&
      Close cmvfi
    End If
  Else
    hiba% = 1
  End If
  Close dbfi: Close icfi
End Sub
Public Sub dbxvir(objazon, dbfi, rec$, recim&, rehos&)
  '--- tranzakciós rekord visszaírása
  If recim& <> 0 Then
    Seek #dbfi, recim&
    rrec$ = Space(rehos&): Get #dbfi, , rrec$
    If rec$ <> rrec$ Then
      If igen123% = 1 Then
        lgfi = FreeFile
        Open auditorutvonal$ + dbxn$ + ".123" For Append As #lgfi
        z$ = objazon + ";" + terminal$ + task$ + ugyintezo$ + Date$ + Time$
        Print #lgfi, z$;
        Print #lgfi, rec$;
        Close lgfi
      End If
      Put #dbfi, recim&, rec$
    End If
  End If
End Sub

Public Sub dbxtrkulcs(kulcs$, hiba%)
  '--- kulcsszavas lockolás
  On Error GoTo letiltva
  lckulcsfi = FreeFile
  DoEvents
  Open auditorutvonal$ + kulcs$ + ".lck" For Binary Shared As lckulcsfi
  DoEvents
  Put #lckulcsfi, 1, kulcs$
  Lock #lckulcsfi
  hiba% = 0
  Exit Sub
letiltva:
  DoEvents
  Close lckulcsfi
  hiba% = 1
End Sub

Public Sub dbxtrkezd5(utvo$, dbxn$)
  '--- tranzakció kezdet lck beállítása otker
  On Error GoTo letiltva
  'dbxn$ = "axneve(objazon)
  lckfi = FreeFile
ismlok:
  Open utvo$ + dbxn$ + ".lck" For Binary Shared As lckfi
  Lock #lckfi
  Exit Sub
letiltva:
  DoEvents
  Close lckfi
  Resume ismlok
End Sub
Public Sub dbxtrkezd(objazon)
  '--- tranzakció kezdet lck beállítása
  On Error GoTo letiltva
  dbxn$ = dbxneve(objazon)
  lckfi = FreeFile
ismlok:
  Open auditorutvonal$ + dbxn$ + ".lck" For Binary Shared As lckfi
  Lock #lckfi
  Exit Sub
letiltva:
  DoEvents
  Close lckfi
  Resume ismlok
End Sub

Public Sub dbxtrkulcsvege()
  '--- kulcsszavas lockolás feloldása
  Unlock #lckulcsfi
  Close lckulcsfi
End Sub

Public Sub dbxtrvege()
  '--- tranzakció vége, lck felszabadítása
  Unlock #lckfi
  Close lckfi
End Sub
Public Sub dbxtulcs(objazon)
  '--- túlcsordulási vektor
  ob% = obsorszama(objazon)
  dbxn$ = dbxneve(objazon)
  w1% = OBJTAB(ob%).obi(1)
  indn$ = RTrim$(INDTAB(w1%).indnev)
  ih& = ADATAB(INDTAB(w1%).adatsorsz).adatho + 5
  kulcs$ = Left(kulcsa$ + Space(ih& - 5), ih& - 5)
  'ih& = Len(kulcs$) + 5
  ic& = OBJTAB(ob%).renszam
  ic1& = OBJTAB(ob%).reszam
  iccim& = OBJTAB(ob%).icccim
  rh& = OBJTAB(ob%).rekhossz
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  rc& = Int(LOF(ndfi) / ih&)
  If rc& > ic& Then
    For i789& = ic& + 1 To rc&
    
    Next
  End If
  Close dbfi: Close ndfi
End Sub

 Public Function dbxkey$(objazon, kulcsa$)
  '--- rekord beolvasasa index szerint
  '--- visszaadja a rekordot és beállítja az objtab-ban
  On Error GoTo hibakez
  If Trim(kulcsa$) = "" Then dbxkey = "": Exit Function
  ob% = obsorszama(objazon)
  dbxn$ = dbxneve(objazon)
  w1% = OBJTAB(ob%).obi(1)
  indn$ = RTrim$(INDTAB(w1%).indnev)
  ih& = ADATAB(INDTAB(w1%).adatsorsz).adatho + 5
  kulcs$ = Left(kulcsa$ + Space(ih& - 5), ih& - 5)
  'ih& = Len(kulcs$) + 5
  ic& = OBJTAB(ob%).renszam
  ic1& = OBJTAB(ob%).reszam
  iccim& = OBJTAB(ob%).icccim
  rh& = OBJTAB(ob%).rekhossz
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  rc& = Int(LOF(ndfi) / ih&)
  If rc& < 1 Then
    OBJTAB(ob%).obcim = 0
    OBJTAB(ob%).obind = 0
    OBJTAB(ob%).obnex = 0
    dbxkey$ = ""
    Close dbfi: Close ndfi
    Exit Function
  End If
  If OBJTAB(ob%).iktlast = 0 Then
    also& = 1: felso& = ic&
    v& = felso& + 1
  Else
    also& = 1: felso& = rc&
    v& = felso& + 1
  End If
  'v& = felso& + 1
  If also& <= felso& Then
    Get #ndfi, (felso& - 1) * ih& + 1, rcim&
    Seek #ndfi, (felso& - 1) * ih& + 5
    xx$ = Space(ih& - 4): Get #ndfi, , xx$
    bj$ = Left$(xx$, ih& - 5)
    tjel$ = Right$(xx$, 1)
    If kulcs$ = bj$ Then
      If tjel$ = "*" Then
        GoTo tulcsord
      Else
        OBJTAB(ob%).obind = felso&
        GoTo olvas
      End If
    End If
    Get #ndfi, (also& - 1) * ih& + 1, rcim&
    Seek #ndfi, (also& - 1) * ih& + 5
    xx$ = Space(ih& - 4): Get #ndfi, , xx$
    bj$ = Left$(xx$, ih& - 5)
    tjel$ = Right$(xx$, 1)
    If kulcs$ = bj$ Then
      If tjel$ = "*" Then
        GoTo tulcsord
      Else
        OBJTAB(ob%).obind = also&
        GoTo olvas
      End If
    End If
    Do
      kozep& = Int((felso& + also&) / 2)
      Get #ndfi, (kozep& - 1) * ih& + 1, rcim&
      Seek #ndfi, (kozep& - 1) * ih& + 5
      xx$ = Space(ih& - 4): Get #ndfi, , xx$
      bj$ = Left$(xx$, ih& - 5)
      tjel$ = Right$(xx$, 1)
      If kulcs$ = bj$ Then
        If tjel$ = "*" Then
          Exit Do
        Else
          OBJTAB(ob%).obind = kozep&
          GoTo olvas
        End If
      Else
        If kulcs$ < bj$ Then
          felso& = kozep&
        Else
          also& = kozep&
        End If
      End If
    Loop While also& + 1 < felso&
  End If
tulcsord:
  If objazon = "ERTT" Then
    'If ertulcs <> 0 Then v& = ertulcs Else ertulcs = v&
    If v& > 0& And v& <= rc& Then
      'beho& = (rc& - v& + 1) * ih&
      'ihevy$ = Space(beho&)
      'Get #ndfi, (v& - 1) * ih& + 1, ihevy$
      'For i4& = 1 To rc& - v& + 1
      '  iopzo& = (i4& - 1) * ih& + 5
      '  bj$ = Mid$(ihv$, iopzo&, ih& - 5)
      '  If bj$ = kulcs$ Then
      '    Get #ndfi, (i4& + v& - 1 - 1) * ih& + 1, rcim&
      '    OBJTAB(ob%).obind = i4& + v& - 1
      '    GoTo olvas
      '  End If
      'Next
      For i4& = v& To rc&
        Seek #ndfi, (i4& - 1) * ih& + 5
        bj$ = Space(ih& - 5): Get #ndfi, , bj$
        If bj$ = kulcs$ Then
          Get #ndfi, (i4& - 1) * ih& + 1, rcim&
          OBJTAB(ob%).obind = i4&
          GoTo olvas
vissza1:
        End If
      Next
    End If
  Else
    If v& > 0& And v& <= rc& Then
      For i4& = v& To rc&
        Seek #ndfi, (i4& - 1) * ih& + 5
        bj$ = Space(ih& - 5): Get #ndfi, , bj$
        If bj$ = kulcs$ Then
          Get #ndfi, (i4& - 1) * ih& + 1, rcim&
          OBJTAB(ob%).obind = i4&
          GoTo olvas
vissza:
        End If
      Next
    End If
  End If
  '--- nincs találat
  Close dbfi: Close ndfi
  keysorszam = 0
  dbxkey$ = ""
  Exit Function
olvas:
  '--- kilépés találattal
  If rcim& > 0 Then
    OBJTAB(ob%).obcim = rcim&
    torlo$ = " "
    Get #dbfi, rcim& + 4, torlo$
    Get #dbfi, rcim& + 5, nex&
    OBJTAB(ob%).obnex = nex&
    Seek #dbfi, rcim& + 9
    rax$ = Space(rh&): Get #dbfi, , rax$
    utrec$(ob%) = rax$
    ' Ha töröltet talál folytatja a keresést
    If torlo$ = "*" Then
      If objazon = "ERTT" Then
        GoTo vissza1
      Else
        GoTo vissza
      End If
    End If
    
    Close dbfi: Close ndfi
    keysorszam = OBJTAB(ob%).obind
    dbxkey$ = utrec$(ob%)
  Else
    OBJTAB(ob%).obcim = 0
    OBJTAB(ob%).obind = 0
    OBJTAB(ob%).obnex = 0
    keysorszam = 0
    dbxkey$ = ""
    Close dbfi: Close ndfi
    Exit Function
  End If
Exit Function
hibakez:
  keysorszam = 0
  dbxkey$ = ""
  Call mess("Dbxkey - Hiba kód: " + Str(Err.Number) + " " + Err.Description, 3, 0, "Hiba", valasz%)
  Resume Next
End Function

Public Function dbxfirst$(objazon)
  '--- a fõpointer szerinti elsõ rekord beolvasása
  ob% = obsorszama(objazon)
  dbxn$ = dbxneve(objazon)
  icfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".icc" For Binary Shared As #icfi
  icim& = OBJTAB(ob%).icccim
  Get #icfi, icim& + 12, rekci&
  Close icfi
  OBJTAB(ob%).rekfirst = rekci&
  cim& = OBJTAB(ob%).rekfirst
  If cim& = 0 Then dbxfirst$ = "": Exit Function
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  Get #dbfi, cim& + 5, nexcim&
  Seek #dbfi, cim& + 9
  X1$ = Space$(rh&): Get #dbfi, , X1$
  Close dbfi
  '--- aktualis beállítások
  OBJTAB(ob%).obcim = cim&
  OBJTAB(ob%).obind = 0
  OBJTAB(ob%).obnex = nexcim&
  utrec$(ob%) = X1$
  dbxfirst$ = X1$
End Function

Public Function dbxlast$(objazon)
  '--- a fõpointer szerinti utolsó rekord beolvasása
  ob% = obsorszama(objazon)
  dbxn$ = dbxneve(objazon)
  icfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".icc" For Binary Shared As #icfi
  icim& = OBJTAB(ob%).icccim
  Get #icfi, icim& + 16, rekci&
  Close icfi
  If rekci& = 0 Then dbxlast$ = "": Exit Function
  OBJTAB(ob%).reklast = rekci&
  cim& = rekci&
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  Seek #dbfi, cim& + 9
  X1$ = Space(rh&): Get #dbfi, , X1$
  Close dbfi
  '--- aktualis beállítások
  OBJTAB(ob%).obcim = cim&
  OBJTAB(ob%).obind = 0
  OBJTAB(ob%).obnex = 0
  utrec$(ob%) = X1$
  dbxlast$ = X1$
End Function

Public Function dbxnext$(objazon)
  '--- a fõpointer szerinti következõ rekord beolvasása
  ob% = obsorszama(objazon)
  dbxn$ = dbxneve(objazon)
  If OBJTAB(ob%).obnex = 0 Then dbxnext$ = "": Exit Function
  rekci& = OBJTAB(ob%).obnex
  cim& = rekci&
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  Get #dbfi, cim& + 5, nexcim&
  Seek #dbfi, cim& + 9
  X1$ = Space(rh&): Get #dbfi, , X1$
  Close dbfi
  '--- aktualis beállítások
  OBJTAB(ob%).obcim = cim&
  OBJTAB(ob%).obind = 0
  OBJTAB(ob%).obnex = nexcim&
  utrec$(ob%) = X1$
  dbxnext$ = X1$
End Function


Public Sub dbxopen(dbxazon, objazon, tipus, runtimhiba%)
  '--- objektum megnyitása
  '--- dbxazon=adatbázis fájl neve
  '--- objazon=objektum azonosító
  '--- tipus=0 rejtett mezõk nem tipus=1 rejtett mezõk is
  '--- tipus=2 egy ablakban minden mezõ
  Dim param$(10)       'prx line paraméterek értelmezése
  Dim kapcsdbxek(30)   'kapcsolódó dbxek neve
  Dim kapcsobjok(30)   'kapcsolódó objektumok azonosítója
  runtimhiba% = 0
'  On Error GoTo hibakez
  kapcsdb% = 0
  objmegvan% = 0
  If obsorszama(objazon) <> 0 Then
    '--- már meg van nyitva
    Exit Sub
  End If
  '--- dbx4 forma
  '--- icc értelmezése
  iccfil = FreeFile
  Open auditorutvonal$ + dbxazon + ".icc" For Binary Shared As #iccfil
  iclen = LOF(iccfil)
  Seek #iccfil, 1
  icrek$ = Space(iclen): Get #iccfil, , icrek$
  iccim& = InStr(icrek$, objazon)
  If (iccim& - 1) Mod 24 <> 0 Then
    i999& = iccim&
    iccim& = InStr(i999& + 1, icrek$, objazon)
  End If
  icrek$ = ""
  prxfil = FreeFile
  '--- prx értelezése
  If Trim(langutvonal$) <> "" Then
    Open langutvonal$ + dbxazon + ".pr4" For Input Shared As #prxfil
  Else
    Open programutvonal$ + dbxazon + ".pr4" For Input Shared As #prxfil
  End If
  Do
    Line Input #prxfil, x$
    kulcs$ = Left$(x$, 3)
    sor$ = Mid$(x$, 5)
    Select Case kulcs$
      Case "OBJ"
        paramdb% = 10
        Call linpar(sor$, param$(), "\", paramdb%)
        If param$(1) = objazon Then
          '--- objektum definició kezdete
          objmegvan% = 1
          If dbxsorszam(dbxazon) = 0 Then
            dbxdb = dbxdb + 1
            DBXTAB(dbxdb) = dbxazon
            dbxindexe = dbxdb
          Else
            dbxindexe = dbxsorszam(dbxazon)
          End If
          objdb = objdb + 1
          OBJTAB(objdb).obaz = param$(1)
          OBJTAB(objdb).obnev = param$(2)
          OBJTAB(objdb).dbxindex = dbxindexe
          OBJTAB(objdb).rekhossz = Val(param$(3))
          OBJTAB(objdb).pointdb = Val(param$(4))
          OBJTAB(objdb).obi(0) = 0
          OBJTAB(objdb).oba(0) = 0
          OBJTAB(objdb).keresokoddarab = 0
          Get #iccfil, iccim& + 4, X1&
          OBJTAB(objdb).reszam = X1&
          Get #iccfil, iccim& + 8, X1&
          OBJTAB(objdb).renszam = X1&
          Get #iccfil, iccim& + 12, X1&
          OBJTAB(objdb).rekfirst = X1&
          Get #iccfil, iccim& + 16, X1&
          OBJTAB(objdb).reklast = X1&
          Get #iccfil, iccim& + 20, X1&
          OBJTAB(objdb).iktlast = X1&
          OBJTAB(objdb).obcim = 0
          OBJTAB(objdb).obind = 0
          OBJTAB(objdb).obnex = 0
          OBJTAB(objdb).icccim = iccim&
          '--- kapcsolódó objektum korábban megnyitott objektumhoz?
          For i2% = 1 To adadb
            If ADATAB(i2%).kapcsobnev = objazon Then
              ADATAB(i2%).kapcsob = objdb
              ADATAB(i2%).kapcsdbx = dbxindexe
            End If
          Next
        Else
          If objmegvan% = 1 Then Exit Do
        End If
      Case "ABL"
        '--- ablak definició kezdete
        If objmegvan% = 1 Then
          If OBJTAB(objdb).oba(0) = 0 Or tipus <> 2 Then
            paramdb% = 10
            Call linpar(sor$, param$(), "\", paramdb%)
            abldb = abldb + 1
            ABLTAB(abldb).fejlec = param$(1)
            ABLTAB(abldb).abmod = param$(2)
            For j% = 1 To 50: ABLTAB(abldb).adatsorsz(j%) = 0: Next
            ii% = OBJTAB(objdb).oba(0) + 1
            OBJTAB(objdb).oba(0) = ii%
            OBJTAB(objdb).oba(ii%) = abldb
          End If
        End If
      Case "MEZ"
        '--- mezõ definició kezdete
        If objmegvan% = 1 Then
          paramdb% = 10
          Call linpar(sor$, param$(), "\", paramdb%)
          If InStr(param$(5), "G") > 0 Or InStr(param$(5), "R") = 0 Or tipus <> 0 Then
            adadb = adadb + 1
            ADATAB(adadb).adkod = param$(1)
            ADATAB(adadb).obsorsz = objdb
            ADATAB(adadb).adatnev = param$(2)
            ADATAB(adadb).adatkp = Val(param$(3))
            ADATAB(adadb).adatho = Val(param$(4))
            ADATAB(adadb).attr = param$(5)
            ADATAB(adadb).kapcsob = 0
            ADATAB(adadb).kapcsdbx = 0
            If InStr(param$(5), "G") = 0 Or tipus = 2 Then
              ii% = ABLTAB(abldb).adatsorsz(0) + 1
              ABLTAB(abldb).adatsorsz(0) = ii%
              ABLTAB(abldb).adatsorsz(ii%) = adadb
            End If
            '--- kapcsolodo objektum kezelése
            If paramdb% > 5 Then
              If obsorszama(param$(7)) > 0 Then
                ADATAB(adadb).kapcsobnev = param$(7)
                ADATAB(adadb).kapcsob = obsorszama(param$(7))
                ADATAB(adadb).kapcsdbx = dbxindexe
              Else
                ADATAB(adadb).kapcsobnev = param$(7)
                kapcsdb% = kapcsdb% + 1
                kapcsdbxek(kapcsdb%) = param$(6)
                kapcsobjok(kapcsdb%) = param$(7)
              End If
            End If
            '--- keresõ kód megállapítása
            If InStr(param$(5), "+[") Then
              OBJTAB(objdb).keresokoddarab = OBJTAB(objdb).keresokoddarab + 1
              ADATAB(adadb).keresokodindex = OBJTAB(objdb).keresokoddarab
              pq1% = InStr(param$(5), "+[")
              pq2% = InStr(param$(5), "]")
              pqh% = Val(Mid$(param$(5), pq1% + 2, pq2% - pq1% - 2))
              ADATAB(adadb).keresokodhossz = pqh
            End If
            Line Input #prxfil, xmagy$
            If Left$(xmagy$, 4) = "ELL=" Then
              ELLENORZO$(adadb) = Mid$(xmagy$, 5)
              Line Input #prxfil, xmagy$
            End If
            MAGYARAZAT$(adadb) = Mid$(xmagy$, 3)
          End If
        End If
      Case "IND"
        '--- index definició kezdete
        If objmegvan% = 1 Then
          paramdb% = 10
          Call linpar(sor$, param$(), "\", paramdb%)
          inddb = inddb + 1
          INDTAB(inddb).indnev = param$(1)
          INDTAB(inddb).obsorsz = objdb
          adatazon$ = param$(2)
          INDTAB(inddb).adatsorsz = adatsorszama(adatazon$, objazon)
          ii% = OBJTAB(objdb).obi(0) + 1
          OBJTAB(objdb).obi(0) = ii%
          OBJTAB(objdb).obi(ii%) = inddb
        End If
      Case "SET"
        If objmegvan% = 1 Then
          paramdb% = 10
          Call linpar(sor$, param$(), "\", paramdb%)
          setdb = setdb + 1
          SETTAB(setdb).setnev = param$(1)
          SETTAB(setdb).robsorsz = obsorszama(param$(2))
          SETTAB(setdb).obsorsz = objdb
          SETTAB(setdb).rootpoz = Val(param$(3))
          SETTAB(setdb).nextpoz = Val(param$(4))
          SETTAB(setdb).priorpoz = Val(param$(5))
          SETTAB(setdb%).owneraz = param$(2)
          SETTAB(setdb%).memberaz = objazon
          adatazon$ = param$(1)
          SETTAB(setdb).adatsorsz = adatsorszama(adatazon$, objazon)
        End If
      Case "HAS"
        If objmegvan% = 1 Then
          paramdb% = 10
          Call linpar(sor$, param$(), "\", paramdb%)
          OBJTAB(obsorszama(param$(1))).hashcod = adatsorszama(param$(2), param$(1))
          OBJTAB(obsorszama(param$(1))).hashmez(0) = 0
          i3% = 0
          For i2% = 3 To paramdb%
            i3% = i3% + 1
            OBJTAB(obsorszama(param$(1))).hashmez(0) = i3%
            w11% = adatsorszama(param$(i2%), param$(1))
            OBJTAB(obsorszama(param$(1))).hashmez(i3%) = w11%
          Next
        End If
      Case Else
    End Select
  Loop While Not EOF(prxfil)
  Close prxfil
  Close iccfil
  If kapcsdb% > 0 Then
    '--- kapcsolódó objektumok megnyitása
    For i3% = 1 To kapcsdb%
      db1$ = kapcsdbxek(i3%)
      ob1$ = kapcsobjok(i3%)
      pap = tipus
      Call dbxopen(db1$, ob1$, pap, runtimhiba%)
      If runtimhiba% = 1 Then Exit Sub
    Next
  End If
  If objmegvan% = 0 Then GoTo hibakez
  Exit Sub
hibakez:
  Call mess(langmodul(161), 1, 0, langmodul(159), valasz%)
  runtimhiba% = 1
End Sub

Public Function adatsorszama(adatazon$, objazon)
  '--- adatab sorszám adatkód alapján
  For i% = 1 To adadb
    If ADATAB(i%).obsorsz = obsorszama(objazon) And Trim$(ADATAB(i%).adkod) = adatazon$ Then
      adatsorszama = i%: Exit Function
    End If
  Next
  adatsorszama = 0
End Function

Public Function dbxsorszam(dbxazon)
  '--- dbxtab sorszám azonosító alapján
  For i% = 1 To dbxdb
    If UCase$(DBXTAB(i%)) = UCase$(dbxazon) Then dbxsorszam = i%: Exit Function
  Next
  dbxsorszam = 0
End Function

Public Sub linpar(sor$, param$(), hatarolo$, paramdb%)
  '--- sor paraméterek szétválasztása
  For p% = 1 To paramdb%: param$(p%) = " ": Next
  paramdb% = 0
  x$ = sor$
  Do
    p% = InStr(x$, hatarolo)
    If p% > 0 Then
      paramdb% = paramdb% + 1
      param$(paramdb%) = Left$(x$, p% - 1)
      x$ = Mid$(x$, p% + 1)
    Else
      If Len(x$) > 0 Then
        paramdb% = paramdb% + 1
        param$(paramdb%) = x$
        x$ = ""
      End If
    End If
  Loop While Len(x$) > 0
End Sub

Public Function obsorszama(objazon)
  '--- objtab sorszám objektum azonosító alapján
  For i% = 1 To objdb
    If OBJTAB(i%).obaz = objazon Then
      obsorszama = i%: Exit Function
    End If
  Next
  obsorszama = 0
End Function

Public Function dbxneve(objazon)
  '--- objektumot tartalmazó dbx neve
  sor% = obsorszama(objazon)
  If sor% <> 0 Then
    dbxneve = DBXTAB(OBJTAB(sor%).dbxindex)
  Else
    dbxneve = ""
  End If
End Function

Public Sub tablazat(objazon, ablaksz%, darab%, mt$(), ttop&, lleft&, sszel&, mmag&)
  '--- tablazatos adatok bevitele
  '--- ablakszám az objektumon belül ablakszám=0 esetén teljes objektum
  '--- mt$-ban a rekordok
  '--- tablazat mérete és poziciója
  '--- ttop,lleft,mmag,sszel twip-ben
  form1.Label2.Caption = ""
  form1.Label3.Caption = ""
ujra:
  Tabla.MSFlexGrid1.Rows = darab% + 2
  Tabla.MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  Tabla.MSFlexGrid1.Font.Size = 8
  Tabla.Text1.Font.Name = "Microsoft Sans Serif"
  Tabla.Text1.Font.Size = 8
  w1& = obsorszama(objazon)
  w2& = OBJTAB(w1&).oba(ablaksz%)
  odarab& = ABLTAB(w2&).adatsorsz(0)
  Tabla.MSFlexGrid1.Cols = odarab& + 1
  For i1& = 1 To darab%
    Tabla.MSFlexGrid1.TextMatrix(i1&, 0) = Str$(i1&) + "." + langmodul(63)
  Next
  gw& = 0
  '--- adatok feltöltése
  For i1& = 1 To darab%
    For i2& = 1 To odarab&
      mh% = ADATAB(ABLTAB(w2&).adatsorsz(i2&)).adatho
      kp% = ADATAB(ABLTAB(w2&).adatsorsz(i2&)).adatkp
      amezo$ = Mid$(mt$(i1&), kp%, mh%)
      Tabla.MSFlexGrid1.TextMatrix(i1&, i2&) = Trim$(amezo$)
    Next
  Next
  '--- fejlec
  For i1& = 1 To odarab&
    ne$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatnev)
    ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).attr)
    mh% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatho
    kp% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatkp
    Tabla.MSFlexGrid1.TextMatrix(0&, i1&) = ne$
    hxh% = Tabla.TextWidth(ne$)
    w3& = Len(ne$)
    If w3& > mh% Then
      h% = hxh% + 100
      'h% = w3& * 110
    Else
      h% = mh% * 110
    End If
    Tabla.MSFlexGrid1.ColWidth(i1&) = h%
    gw& = gw& + h%
    Tabla.Caption = ABLTAB(w2&).fejlec
    If InStr(ar$, "J") > 0 Then
      Tabla.MSFlexGrid1.ColAlignment(i1&) = 6
    Else
      Tabla.MSFlexGrid1.ColAlignment(i1&) = 1
    End If
    mtb$(i1&) = ar$: mho%(i1&) = mh%
    mesor%(i1&) = ABLTAB(w2&).adatsorsz(i1&)
  Next
  '--- méretek beállítása
  If ttop& <> 0 Then Tabla.Top = ttop& Else Tabla.Top = 2000
  If lleft& <> 0 Then Tabla.Left = lleft& Else Tabla.Left = 500
  If sszel& <> 0 Then szel& = sszel& Else szel& = gw& + 1000
  If mmag& <> 0 Then mag& = mmag& Else mag& = darab% * 240 + 1810
  balf& = Tabla.Left
  If balf& + szel& > 12300 Then szel& = 12300 - balf&
  tetf& = Tabla.Top
  If tetf& + mag& > 8700 Then mag& = 8700 - tetf&
  Tabla.Height = mag&
  Tabla.MSFlexGrid1.Height = mag& - 1500
  Tabla.Width = szel&
  Tabla.MSFlexGrid1.Width = szel& - 300
  Tabla.Text1.Top = mag& - 1300
  Tabla.Text1.Width = szel& - 2500
  Tabla.Text2.Top = mag& - 1300
  Tabla.Text2.Width = szel& - 2500
  Tabla.Text3.Text = Str$(Tabla.MSFlexGrid1.Rows - 2) + "." + langmodul(63)
  Tabla.Command1.Top = mag& - 900
  Tabla.Command1.Left = szel& - 2700
  Tabla.Command2.Left = szel& - 1400
  Tabla.Text3.Top = Tabla.Text2.Top
  Tabla.Command4.Top = mag& - 1300
  Tabla.Command4.Left = szel& - 1400
  Tabla.Text3.Left = szel& - 2200
  Tabla.Command2.Top = mag& - 900
  Tabla.Command3.Top = mag& - 900
  Tabla.Command3.Width = szel& - Tabla.Command2.Width - Tabla.Command1.Width - 500
  Tabla.com3wid = Tabla.Command3.Width
  '--- megjelenítés
  objektum$ = objazon: ablaksorszam% = ablaksz%
ismetel:
  Tabla.Show vbModal
   If rogzites% = 1 Then
    '--- adatok ellenõrzése
    kihiba% = 0
    For i1& = 1 To darab%
      svr$ = ""
      For i2& = 1 To odarab&
        amez$ = Tabla.MSFlexGrid1.TextMatrix(i1&, i2&)
        svr$ = svr$ + Trim(amez$)
      Next
      If Trim(svr$) <> "" Then
        If sortorol% = 0 Then
          For i2& = 1 To odarab&
            ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i2&)).attr)
            amez$ = Tabla.MSFlexGrid1.TextMatrix(i1&, i2&)
            anvv$ = Trim(ADATAB(ABLTAB(w2&).adatsorsz(i2&)).adatnev)
            If InStr(ar$, "*") > 0 Then
              If Trim(amez$) = "" Then
                Call mess(anvv$ + "kitöltése kötelezõ!" + Chr$(13) + "Hiba a(z) " + Str(i1&) + ".sorban", 2, 0, "Hiba", valasz%)
                kihiba% = 1: Exit For
              End If
            End If
          Next
        End If
      End If
      If kihiba% = 1 Then Exit For
    Next
    If kihiba% = 1 Then GoTo ismetel
    '--- adatok visszaírása a rekordba
    For i1& = 1 To darab%
      For i2& = 1 To odarab&
        ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i2&)).attr)
        mh% = ADATAB(ABLTAB(w2&).adatsorsz(i2&)).adatho
        kp% = ADATAB(ABLTAB(w2&).adatsorsz(i2&)).adatkp
        amez$ = Tabla.MSFlexGrid1.TextMatrix(i1&, i2&)
        If InStr(ar$, "J") > 0 Then
          amez$ = Right$(Space$(mh%) + amez$, mh%)
        Else
          amez$ = Left$(amez$ + Space$(mh%), mh%)
        End If
        Mid$(mt$(i1&), kp%, mh%) = amez$
      Next
    Next
  End If
  If sortorol% = 1 Then
    If programnev$ = "AUW-RSZL" And szamlaeset% = 7 Then
      Mid$(mt$(torstart&), 7, 12) = Space$(12)
      Mid$(mt$(torstart&), 25, 42) = Space$(42)
    Else
      For i7& = torstart& To darab%
        mt$(i7&) = mt$(i7& + 1)
      Next
      mt$(darab%) = Space$(Len(mt$(1)))
    End If
    sortorol% = 0
    Unload Tabla
    GoTo ujra
  Else
     
    Unload Tabla
  End If
End Sub

Public Sub vektabl(objazon, ablaksz%, rec$, ttop&, lleft&, sszel&, mmag&)
  '--- vektoros adatok bevitele
  '--- ablakszám az objektumon belül ablakszám=0 esetén teljes objektum
  '--- rec$-ban a rekord
  '--- tablazat mérete és poziciója
  '--- ttop,lleft,mmag,sszel twip-ben
  charpertwipa% = 120
  charpertwip% = 120
  form1.Label2.Caption = ""
  form1.Label3.Caption = ""
  Vektor.MSFlexGrid1.Clear
  Vektor.MSFlexGrid1.Cols = 2
  w1& = obsorszama(objazon)
  w2& = OBJTAB(w1&).oba(ablaksz%)
  wodarab& = ABLTAB(w2&).adatsorsz(0)
  foindex% = INDTAB(OBJTAB(w1&).obi(1)).adatsorsz
  Vektor.MSFlexGrid1.TextMatrix(0, 0) = langmodul(75)
  Vektor.MSFlexGrid1.TextMatrix(0, 1) = langmodul(76)
  mmax% = 12: hxmax% = 0
  mhmax% = 10
  odarab& = 0
  Vektor.Font.Name = "Microsoft Sans Serif"
  Vektor.Font.Size = 8
  Vektor.MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  Vektor.Text1.Font.Name = "Microsoft Sans Serif"
  'Vektor.MSFlexGrid1.Font.Size = 9
  For i1& = 1 To wodarab&
    ne$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatnev)
    ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).attr)
    If InStr(ar$, "R") = 0 Then
      odarab& = odarab& + 1
      Vektor.MSFlexGrid1.Rows = odarab& + 2
      ar$ = RTrim$(ADATAB(ABLTAB(w2&).adatsorsz(i1&)).attr)
      mh% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatho
      kp% = ADATAB(ABLTAB(w2&).adatsorsz(i1&)).adatkp
      Vektor.MSFlexGrid1.TextMatrix(odarab&, 0) = ne$
      w3& = Len(ne$)
      hxh% = Vektor.TextWidth(ne$) + 100
      If hxmax% < hxh% Then hxmax% = hxh%
      'If w3& > mmax% Then h% = w3& * charpertwipa%: mmax% = w3& Else h% = mmax% * charpertwipa%
      Vektor.MSFlexGrid1.ColWidth(0) = hxmax%
      'Vektor.MSFlexGrid1.ColWidth(0) = h%
      If mh% > mhmax% Then h1% = mh% * charpertwip%: mhmax% = mh% Else h1% = mhmax% * charpertwip%
      Vektor.MSFlexGrid1.ColWidth(1) = h1%
      Vektor.MSFlexGrid1.Width = hxmax% + mhmax% * charpertwip% + 70
      Vektor.Width = Vektor.MSFlexGrid1.Width + 300
      'Vektor.MSFlexGrid1.Width = mmax% * charpertwip% + mhmax% * charpertwip% + 70
      'Vektor.Width = mmax% * charpertwipa% + mhmax% * charpertwip% + 370
      Vektor.Caption = ABLTAB(w2&).fejlec
      Vektor.MSFlexGrid1.ColSel = 1
      Vektor.MSFlexGrid1.RowSel = odarab&
      If InStr(ar$, "J") > 0 And InStr(ar$, "NZJ") = 0 Then
        Vektor.MSFlexGrid1.CellAlignment = flexAlignLeftCenter
      Else
        Vektor.MSFlexGrid1.CellAlignment = flexAlignLeftCenter
      End If
      amezo$ = Trim$(Mid$(rec$, kp%, mh%))
      Vektor.MSFlexGrid1.TextMatrix(odarab&, 1) = amezo$
      mtb$(odarab&) = ar$: mho%(odarab&) = mh%: mkp(odarab&) = kp%
      mesor(odarab&) = ABLTAB(w2&).adatsorsz(i1&)
    End If
  Next
  If sszel& <> 0 Then Vektor.Width = sszel&: Vektor.MSFlexGrid1.Width = sszel& - 300
  If ttop& <> 0 Then
    Vektor.Top = ttop& '+ form1.Top
  Else
    Vektor.Top = 1500 '+ form1.Top
  End If
  If lleft& <> 0 Then
    Vektor.Left = lleft& '+ form1.Left
  Else
    Vektor.Left = 100 '+ form1.Left
  End If
  If Vektor.Left + Vektor.Width > form1.Width Then Vektor.Left = form1.Width - Vektor.Width
  If mmag& <> 0 Then mag& = mmag& Else mag& = odarab& * 240 + 2200
  tetf& = Vektor.Top
  If tetf& + mag& > 8700 Then mag& = 8700 - tetf&
  'If sszel& = 0 And Vektor.Width < 5000 Then
  If sszel& = 0 And Vektor.Width < 3750 Then
    Vektor.Command1.Width = Vektor.Width - 250
    Vektor.Command2.Width = (Vektor.Width - 300) / 2
    Vektor.Command3.Width = (Vektor.Width - 300) / 2
    Vektor.Command3.Left = Vektor.Command2.Left + Vektor.Command2.Width + 50
  End If
  Vektor.Height = mag& - 200
  Vektor.MSFlexGrid1.Height = mag& - 1870
  Vektor.Text1.Top = ttop& + mag& - 1300
  Vektor.Text1.Width = (mmax% + mhmax%) * charpertwip%
  Vektor.Text2.Top = mag& - 1750
  Vektor.Text2.Width = Vektor.MSFlexGrid1.Width
  Vektor.com1wid = Vektor.Command1.Width
  Vektor.Command1.Top = mag& - 1400
  Vektor.Command2.Top = mag& - 975
  Vektor.Command3.Top = mag& - 975
  objektum$ = objazon: ablaksorszam% = ablaksz%
  Vektor.Show vbModal
  If rogzites% = 1 Then
    '--- adatok visszaírása a rekordba
    For i1& = 1 To odarab&
      ar$ = mtb(i1&)
      mh% = mho(i1&)
      kp% = mkp(i1&)
      amez$ = Vektor.MSFlexGrid1.TextMatrix(i1&, 1)
      If InStr(ar$, "J") > 0 Then
        amez$ = Right$(Space$(mh%) + amez$, mh%)
      Else
        amez$ = Left$(amez$ + Space$(mh%), mh%)
      End If
      Mid$(rec$, kp%, mh%) = amez$
    Next
  End If
  Unload Vektor
End Sub

Public Sub alth(objaz, azonosito$)
  '--- ALt+H val rekord választás
  '--- azonosito$-ban a választott kulcs, vagy azonosito$="", ha nincs választás
  Call dbxker("", objaz, 1, talalat%, rec$)
  If talalat% = 0 Then
    azonosito$ = ""
  Else
    wq1& = obsorszama(objaz)
    wq2& = OBJTAB(wq1&).obi(1)
    wq3& = INDTAB(wq2&).adatsorsz
    kp% = ADATAB(wq3&).adatkp
    ho% = ADATAB(wq3&).adatho
    azonosito$ = Mid$(rec$, kp%, ho%)
  End If
End Sub

Public Sub altx(objaz, azonosito$)
  '--- ALt+X hash rekord választás
  '--- azonosito$-ban a választott kulcs, vagy azonosito$="", ha nincs választás
  If OBJTAB(obsorszama(objaz)).hashcod = 0 Then Exit Sub
  If altxobjektum$ <> objaz Then
    aktrox& = 1: toprox& = 1: aktdarab& = 0: tabstatusx% = 0
    altxobjektum$ = objaz
  Else
  End If
  Call dbxhash(objaz, 0, talalat%, rec$)
  If talalat% = 0 Then
    azonosito$ = ""
  Else
    wq1& = obsorszama(objaz)
    wq2& = OBJTAB(wq1&).obi(1)
    wq3& = INDTAB(wq2&).adatsorsz
    kp% = ADATAB(wq3&).adatkp
    ho% = ADATAB(wq3&).adatho
    azonosito$ = Mid$(rec$, kp%, ho%)
  End If
End Sub

Public Sub rendez(rsor$)
  '--- rendezõprogram hívása
  rendpar$ = rsor$
  rendezo.Timer1.Interval = 1000
  rendezo.Show vbModal
  Unload rendezo
End Sub
Public Sub informa()
  '--- üres rutin
End Sub

Public Sub dbxhash(objazon, umod%, talalat%, rec$)
  '--- keresés egy objektumba hashkód alapján
  '--- gombsor csak kezdolap, umod%=0 esetén
  '--- umod%=1 keresõtábla, umod%=0 kezdõlap
  '--- umod%=1 esetén talalat%=1 van kiválasztott, talalat%=0 nincs
  '--- rec a választott rekord (objtab-ban is beállítva)
  Dim param$(10)
  keresomod% = umod%
  Hash.MSFlexGrid1.Cols = 0
  Hash.MSFlexGrid1.Rows = 15
  w1& = obsorszama(objazon)
  w2& = OBJTAB(w1&).hashmez(0)
  For i1& = 1 To w2&
    w3& = OBJTAB(w1&).hashmez(i1&)
    Hash.MSFlexGrid1.Cols = Hash.MSFlexGrid1.Cols + 1
    oslo% = Hash.MSFlexGrid1.Cols - 1
    ne$ = RTrim$(ADATAB(w3&).adatnev)
    ar$ = RTrim$(ADATAB(w3&).attr)
    mh% = ADATAB(w3&).adatho
    kp% = ADATAB(w3&).adatkp
    Hash.MSFlexGrid1.TextMatrix(0&, oslo%) = ne$
    w3& = Len(ne$)
    If w3& > mh% Then h% = w3& * 100 Else h% = mh% * 100
    Hash.MSFlexGrid1.ColWidth(oslo%) = h%
    gw& = gw& + h%
    Hash.Caption = OBJTAB(w1&).obnev
    If InStr(ar$, "J") > 0 Then
      Hash.MSFlexGrid1.ColAlignment(oslo%) = 6
    Else
      Hash.MSFlexGrid1.ColAlignment(oslo%) = 1
    End If
    kmtb$(oslo% + 1) = ar$: kmho%(oslo% + 1) = mh%: kmkp%(oslo% + 1) = kp%
    kmesor%(oslo% + 1) = i1&: kmesor%(0) = oslo% + 1
  Next
  keresobj$ = objazon
  Hash.Show vbModal
  rec$ = rekord$
  utrec$(w1&) = rekord$
  Unload Hash
End Sub
Public Sub dbxreszh(objazon)
  '--- keresés és részhalmaz kezelése
  Dim param$(20)
  Reszhalm.MSFlexGrid1.Cols = 0
  Reszhalm.MSFlexGrid1.Rows = 10
  w1& = obsorszama(objazon)
  For i1& = 1 To adadb%
    If ADATAB(i1&).obsorsz = w1& Then
      Reszhalm.MSFlexGrid1.Cols = Reszhalm.MSFlexGrid1.Cols + 1
      oslo% = Reszhalm.MSFlexGrid1.Cols - 1
      ne$ = RTrim$(ADATAB(i1&).adatnev)
      ar$ = RTrim$(ADATAB(i1&).attr)
      mh% = ADATAB(i1&).adatho
      kp% = ADATAB(i1&).adatkp
      hsxsor% = ADATAB(i1&).keresokodindex
      hsxhos% = ADATAB(i1&).keresokodhossz
      Reszhalm.MSFlexGrid1.TextMatrix(0&, oslo%) = ne$
      w3& = Len(ne$)
      hxh% = Reszhalm.TextWidth(ne$)
      hxk% = Reszhalm.TextWidth(String(mh%, "O"))
      If hxh% > hxk% Then
        h% = hxh% + 100
      Else
        h% = hxk% + 100
      End If
      Reszhalm.MSFlexGrid1.ColWidth(oslo%) = h%
      gw& = gw& + h%
      If InStr(ar$, "J") > 0 Then
        Reszhalm.MSFlexGrid1.ColAlignment(oslo%) = 6
      Else
        Reszhalm.MSFlexGrid1.ColAlignment(oslo%) = 1
      End If
      kmtb$(oslo% + 1) = ar$: kmho%(oslo% + 1) = mh%: kmkp%(oslo% + 1) = kp%
      kmesor%(oslo% + 1) = i1&: kmesor%(0) = oslo% + 1
      kmehsx%(oslo% + 1) = hsxsor%: kmehsxh%(oslo% + 1) = hsxhos%
    End If
  Next
  Reszhalm.Caption = OBJTAB(w1&).obnev
  keresobj$ = objazon
  Reszhalm.Show 'vbModal
End Sub

Public Sub dbxker(gombsor$, objazon, umod%, talalat%, rec$)
  '--- keresés egy objektumba
  '--- gombsor csak kezdolap, umod%=0 esetén
  '--- umod%=1 keresõtábla, umod%=0 kezdõlap
  '--- umod%=1 esetén talalat%=1 van kiválasztott, talalat%=0 nincs
  '--- rec a választott rekord (objtab-ban is beállítva)
  '--- gombsorszam%-ban a lenyomott gomb sorszáma
  Dim param$(20)
  On Error GoTo hibakez
  keresomod% = umod%
  If umod% = 2 Then umod% = 1: keresomod% = 0
  If umod% = 0 Then
    '--- kezdõlap
    zkermod% = 0
    paramdb% = 11
    Call linpar(gombsor$, param$(), "&", paramdb%)
    If paramdb% > 0 Then
      Kereso.Command1.Visible = True: Kereso.Command1.Caption = param$(1)
    Else
      Kereso.Command1.Visible = False
    End If
    If paramdb% > 1 Then
      Kereso.Command2.Visible = True: Kereso.Command2.Caption = param$(2)
    Else
      Kereso.Command2.Visible = False
    End If
    If paramdb% > 2 Then
      Kereso.Command3.Visible = True: Kereso.Command3.Caption = param$(3)
    Else
      Kereso.Command3.Visible = False
    End If
    If paramdb% > 3 Then
      Kereso.Command4.Visible = True: Kereso.Command4.Caption = param$(4)
    Else
      Kereso.Command4.Visible = False
    End If
    If paramdb% > 4 Then
      Kereso.Command5.Visible = True: Kereso.Command5.Caption = param$(5)
    Else
      Kereso.Command5.Visible = False
    End If
    If paramdb% > 5 Then Kereso.Command12.Visible = True: Kereso.Command12.Caption = param$(6) Else Kereso.Command12.Visible = False
    If paramdb% > 6 Then
      Kereso.Command13.Visible = True: Kereso.Command13.Caption = param$(7)
      If paramdb% > 7 Then
        Kereso.Frame1.Height = 1212
        Kereso.Height = 6456
        Kereso.Command6.Top = 720
      Else
        Kereso.Command13.Top = 240
        Kereso.Command13.Left = 6120
      End If
    Else
      Kereso.Command13.Visible = False
    End If
    If paramdb% > 7 Then Kereso.Command14.Visible = True: Kereso.Command14.Caption = param$(8) Else Kereso.Command14.Visible = False
    If paramdb% > 8 Then Kereso.Command15.Visible = True: Kereso.Command15.Caption = param$(9) Else Kereso.Command15.Visible = False
    If paramdb% > 9 Then Kereso.Command16.Visible = True: Kereso.Command16.Caption = param$(10) Else Kereso.Command16.Visible = False
    If paramdb% > 10 Then Kereso.Command17.Visible = True: Kereso.Command17.Caption = param$(11) Else Kereso.Command17.Visible = False
    If paramdb% > 11 Then Kereso.Command27.Visible = True: Kereso.Command27.Caption = param$(12) Else Kereso.Command27.Visible = False
    If paramdb% > 12 Then Kereso.Command28.Visible = True: Kereso.Command28.Caption = param$(13) Else Kereso.Command28.Visible = False
    If paramdb% > 13 Then Kereso.Command29.Visible = True: Kereso.Command29.Caption = param$(14) Else Kereso.Command29.Visible = False
  Else
    '--- keresõtábla
    zkermod% = 1
    Kereso.Command1.Visible = False
    Kereso.Command2.Visible = False
    Kereso.Command3.Visible = False
    Kereso.Command4.Visible = False
    Kereso.Command5.Visible = False
    Kereso.Command13.Visible = False
    Kereso.Command12.Visible = True
    Kereso.Command12.Caption = langmodul(77)
    Kereso.Command12.BackColor = RGB(255, 255, 255)
    Kereso.Command6.Visible = True
    Kereso.Command6.Caption = langmodul(78)
    Kereso.Command6.Cancel = True
  End If
  Kereso.MSFlexGrid1.Cols = 0
  Kereso.MSFlexGrid1.Rows = 18
  w1& = obsorszama(objazon)
  For i1& = 1 To adadb%
    If ADATAB(i1&).obsorsz = w1& Then
      Kereso.MSFlexGrid1.Cols = Kereso.MSFlexGrid1.Cols + 1
      oslo% = Kereso.MSFlexGrid1.Cols - 1
      ne$ = RTrim$(ADATAB(i1&).adatnev)
      ar$ = RTrim$(ADATAB(i1&).attr)
      mh% = ADATAB(i1&).adatho
      kp% = ADATAB(i1&).adatkp
      hsxsor% = ADATAB(i1&).keresokodindex
      hsxhos% = ADATAB(i1&).keresokodhossz
      Kereso.MSFlexGrid1.TextMatrix(0&, oslo%) = ne$
      w3& = Len(ne$)
      hxh% = Kereso.TextWidth(ne$)
      hxk% = Kereso.TextWidth(String(mh%, "O"))
      If hxh% > hxk% Then
        h% = hxh% + 100
      Else
        h% = hxk% + 100
      End If
      Kereso.MSFlexGrid1.ColWidth(oslo%) = h%
      gw& = gw& + h%
      'If Kereso.Caption = langmodul(79) Then Kereso.Caption = OBJTAB(w1&).obnev
      If InStr(ar$, "J") > 0 Then
        Kereso.MSFlexGrid1.ColAlignment(oslo%) = 6
      Else
        Kereso.MSFlexGrid1.ColAlignment(oslo%) = 1
      End If
      kmtb$(oslo% + 1) = ar$: kmho%(oslo% + 1) = mh%: kmkp%(oslo% + 1) = kp%
      kmesor%(oslo% + 1) = i1&: kmesor%(0) = oslo% + 1
      kmehsx%(oslo% + 1) = hsxsor%: kmehsxh%(oslo% + 1) = hsxhos%
    End If
  Next
  'If umod% > 0 Then Kereso.Caption = OBJTAB(w1&).obnev
  Kereso.Caption = OBJTAB(w1&).obnev
  keresobj$ = objazon
  Kereso.Show vbModal
  rec$ = rekord$
  utrec$(w1&) = rekord$
  Unload Kereso
  Exit Sub
hibakez:
  rec$ = ""
  talalat% = 0
  Call mess("Dbxker - Hiba kód: " + Str(Err.Number) + " " + Err.Description, 3, 0, "Hiba", valasz%)

End Sub

Public Function konver$(k$)
  '--- string konvertálása 852,437 ---> Windows
  z$ = k$
  q2$ = Chr$(160) + Chr$(130) + Chr$(161) + Chr$(162) + Chr$(148) + Chr$(139) + Chr$(163) + Chr$(129) + Chr$(251)
  q2$ = q2$ + Chr$(143) + Chr$(181) + Chr$(144) + Chr$(214) + Chr$(224) + Chr$(153) + Chr$(138) + Chr$(233) + Chr$(154) + Chr$(235)
  q1$ = "áéíóöõúüûÁÁÉÍÓÖÕÚÜÛ"
  i1% = Len(k$)
  For i2% = 1 To i1%
    For i3% = 1 To Len(q1$)
      If Mid$(z$, i2%, 1) = Mid$(q2$, i3%, 1) Then
        Mid$(z$, i2%, 1) = Mid$(q1$, i3%, 1): Exit For
      End If
    Next
  Next
  konver$ = z$
End Function
Public Function konver852$(k$)
  z$ = k$
  q1$ = Chr$(160) + Chr$(130) + Chr$(161) + Chr$(162) + Chr$(148) + Chr$(139) + Chr$(163) + Chr$(129) + Chr$(251)
  q1$ = q1$ + Chr$(181) + Chr$(143) + Chr$(144) + Chr$(214) + Chr$(224) + Chr$(153) + Chr$(138) + Chr$(233) + Chr$(154) + Chr$(235)
  q2$ = "áéíóöõúüûÁÁÉÍÓÖÕÚÜÛ"
  i1% = Len(k$)
  For i2% = 1 To i1%
    For i3% = 1 To Len(q1$)
      If Mid$(z$, i2%, 1) = Mid$(q2$, i3%, 1) Then
        Mid$(z$, i2%, 1) = Mid$(q1$, i3%, 1): Exit For
      End If
    Next
  Next
  konver852$ = z$
End Function

Public Sub rekinfo(objazon$, azonosito$)
  '--- Alt+I
  '--- rekord tartalom megjelenítése listbox-ban
  rex$ = dbxkey(objazon$, azonosito$)
  objekts% = obsorszama(objazon$)
 
  If rex$ <> "" Then
    form1.Text1.Visible = True
    form1.Info.Visible = True
    form1.Info.Clear
    form1.Info.Font.Name = "Courier New"
    form1.Info.Font.Size = 9
    form1.Info.Font.Bold = True
    form1.Info.Left = 7040
    form1.Info.Width = 5092
    form1.Text1.Top = form1.Info.Top - 350
    form1.Text1.Left = form1.Info.Left
    form1.Text1.Width = form1.Info.Width
    form1.Info.ForeColor = RGB(35, 120, 120)
    form1.Text1.BackColor = RGB(35, 120, 120)
    form1.Text1.ForeColor = RGB(255, 255, 255)
    'form1.Text1.BackColor = RGB(255, 255, 255)
    'form1.Text1.ForeColor = RGB(35, 100, 100)
    form1.Text1.Font.Name = "Ariel"
    form1.Text1.Font.Size = 10
    form1.Text1.Font.Bold = True
    ssd% = -1
    For i6& = 1 To adadb
      If ADATAB(i6&).obsorsz = objekts% Then
        kp% = ADATAB(i6&).adatkp
        ho% = ADATAB(i6&).adatho
        nev$ = ADATAB(i6&).adatnev
        mezo$ = Mid$(rex$, kp%, ho%)
        'If Len(Trim(nev$)) < 10 Then nell% = 10 Else nell% = 15
        If objazon$ <> "KSZC" And ssd% = 0 Or objazon$ = "KSZC" And ssd% = 1 Then
          form1.Text1.Text = Trim$(nev$) + ": " + Trim$(mezo$)
        Else
          If ho% > 30 Then
            form1.Info.AddItem Trim$(nev$) + ": " + Trim$(mezo$)
          Else
            form1.Info.AddItem Left(Trim$(nev$) + Space(15), 15) + ": " + Trim$(mezo$)
          End If
        End If
        ssd% = ssd% + 1
        If ssd% = 29 Then Exit For
      End If
    Next
    If ssd% * 210 < 6000 Then
      form1.Info.Height = ssd% * 210
    Else
      form1.Info.Height = 6000
    End If
    form1.Info.Visible = True
  End If
End Sub
Public Sub komin1(komnev$, komsorsz%, komfeltetel%(), komxsor%, komxro%, fejlec$, komment$, runtimhiba%)
  '--- Kérdezõ tábla kom file alapján táblázatos
  runtimhiba% = 0
  Dim param$(30)
  komdb% = 0: komadatdb% = 0: kommenudb% = 0
  opt1db% = 0: opt2db% = 0: chekdb% = 0
  komfi = FreeFile
  If langutvonal$ <> "" Then
    Open langutvonal$ + komnev$ + ".kom" For Input Shared As #komfi
  Else
    Open programutvonal$ + komnev$ + ".kom" For Input Shared As #komfi
  End If
  komdb% = 0
  Do
    Line Input #komfi, a$
    If Mid(a$, 1, 2) = "--" Then kcimsor$ = Trim(Mid$(a$, 5)): Komind1.Caption = kcimsor
    If a$ = Right$("0" + Trim$(Str$(komsorsz%)), 2) Then Exit Do
  Loop While Not EOF(komfi)
  Do
    Line Input #komfi, a$
    If a$ = "*" Then Exit Do
    komdb% = komdb% + 1
    For i9& = 1 To 10: param$(i9&) = "": Next
    paramdb% = 10
    Call linpar(a$, param$(), "\", paramdb%)
    If param(1) = "E" Then
      komt(komdb%).komkod = "1"
      komt(komdb%).komatr = param$(2)
      komt(komdb%).komszov = param$(3)
      komadatdb% = komadatdb% + 1
    End If
    If param(1) = "I" Then
      komt(komdb%).komkod = "2"
      komt(komdb%).komatr = param$(2)
      komt(komdb%).komszov = param$(3)
      komadatdb% = komadatdb% + 1
    End If
    If param(1) = "M" Then
      menusor% = komdb%
      komt(komdb%).komkod = "3"
      komt(komdb%).komszov = param$(2)
      komt(komdb%).kommnv = 1
      kommenudb% = kommenudb% + 1
    End If
    If paramdb% = 5 Then
      komt(komdb%).komdbx = UCase$(param$(4))
      komt(komdb%).komobj = UCase$(param$(5))
      Call dbxopen(UCase(param$(4)), UCase(param$(5)), 0, runtimhiba%)
      If runtimhiba% = 1 Then Exit Sub
    Else
      komt(komdb%).komobj = Space$(4)
    End If
  Loop While Not EOF(komfi)
  Close komfi
  If komadatdb% > 1 Then
    For w1% = 2 To komadatdb%
      Komind1.Option1(w1% - 2).Caption = Trim(komt(w1%).komszov)
      Komind1.Option1(w1% - 2).Visible = True
      Komind1.Option2(w1% - 2).Caption = Trim(komt(w1%).komszov)
      Komind1.Option2(w1% - 2).Visible = True
    Next
  End If
  If kommenudb% > 0 Then
    Call linpar(komt(menusor%).komszov, param(), "&", w2%)
    For w1% = 1 To w2% - 1
      Komind1.Check1(w1% - 1).Caption = Trim(param(w1%))
      Komind1.Check1(w1% - 1).Visible = True
    Next
  End If
  Komind1.MSFlexGrid1.Rows = komadatdb% + 1
  Komind1.MSFlexGrid1.Cols = 3
  Komind1.MSFlexGrid1.ColWidth(0) = 1500
  Komind1.MSFlexGrid1.ColWidth(1) = 1200
  Komind1.MSFlexGrid1.ColWidth(2) = 1200
  Komind1.MSFlexGrid1.TextMatrix(0, 1) = langmodul(80)
  Komind1.MSFlexGrid1.TextMatrix(0, 2) = langmodul(81)
  For w1% = 1 To komadatdb%
    Komind1.MSFlexGrid1.TextMatrix(w1%, 0) = Trim$(komt(w1%).komszov)
    Komind1.MSFlexGrid1.TextMatrix(w1%, 1) = ""
    Komind1.MSFlexGrid1.TextMatrix(w1%, 2) = ""
  Next
  Komind1.Frame2.Visible = False
  Komind1.Frame3.Visible = False
  Komind1.Frame4.Visible = False
  Komind1.Command2.Visible = False
  Komind1.Show vbModal
  'Unload Komind1
  komment$ = ""
  For w2% = 1 To komadatdb%
    If Trim(komt(w2%).komtol) <> "" Or Trim(komt(w2%).komig) <> "" Then
      komfeltetel(w2%) = 1
    Else
      komfeltetel(w2%) = 0
    End If
  Next
  komxsor = 0
  For w2% = 1 To komadatdb%
    If Komind1.Option1(w2% - 1).Value = True Then komxsor = w2%: Exit For
  Next
  komxro = 0
  For w2% = 1 To komadatdb%
    If Komind1.Option2(w2% - 1).Value = True Then komxro = w2%: Exit For
  Next
  For w2% = 1 To komdb%
    Select Case komt(w2%).komkod
      Case "2"
        If Trim$(komt(w2%).komtol) <> "" And Trim$(komt(w2%).komig) <> "" Then
          komment$ = komment$ + Trim(komt(w2%).komszov) + ":" + Trim(komt(w2%).komtol) + "-" + Trim(komt(w2%).komig) + ", "
        End If
      Case "1"
        If Trim$(komt(w2%).komtol) <> "" Then
          komment$ = komment$ + Trim(komt(w2%).komszov) + ":" + Trim(komt(w2%).komtol) + ", "
        End If
      Case Else
    End Select
  Next
  Exit Sub
hibakez:
  Call mess(langmodul(158), 1, 0, langmodul(159), valasz%)
  runtimhiba% = 1
End Sub
Public Sub komin(komnev$, komsorsz%, fejlec$, komment$, runtimhiba%)
  '--- Kérdezõ tábla kom file alapján
  '--- komnev$ a .kom file neve
  '--- komsorsz% a szakasz sorszáma a kom file-ban
  '--- fejléc a komind.caption számára
  runtimhiba% = 0
  Dim param$(20)
  Dim lis As ListBox
  On Error GoTo hibakez
  komdb% = 0: komadatdb% = 0: kommenudb% = 0
  komfi = FreeFile
  If langutvonal$ <> "" Then
    Open langutvonal$ + komnev$ + ".kom" For Input Shared As #komfi
  Else
    Open programutvonal$ + komnev$ + ".kom" For Input Shared As #komfi
  End If
  komdb% = 0
  Do
    Line Input #komfi, a$
    If Mid(a$, 1, 2) = "--" Then kcimsor$ = Trim(Mid$(a$, 5))
    If a$ = Right$("0" + Trim$(Str$(komsorsz%)), 2) Then Exit Do
  Loop While Not EOF(komfi)
  Do
    Line Input #komfi, a$
    If a$ = "*" Then Exit Do
    komdb% = komdb% + 1
    For i9& = 1 To 10: param$(i9&) = "": Next
    paramdb% = 10
    Call linpar(a$, param$(), "\", paramdb%)
    If param(1) = "E" Then
      komt(komdb%).komkod = "1"
      komt(komdb%).komatr = param$(2)
      komt(komdb%).komszov = param$(3)
      komadatdb% = komadatdb% + 1
    End If
    If param(1) = "I" Then
      komt(komdb%).komkod = "2"
      komt(komdb%).komatr = param$(2)
      komt(komdb%).komszov = param$(3)
      komadatdb% = komadatdb% + 1
    End If
    If param(1) = "M" Then
      komt(komdb%).komkod = "3"
      komt(komdb%).komszov = param$(2)
      komt(komdb%).kommnv = 1
      kommenudb% = kommenudb% + 1
    End If
    If paramdb% = 5 Then
      komt(komdb%).komdbx = UCase$(param$(4))
      komt(komdb%).komobj = UCase$(param$(5))
      Call dbxopen(UCase(param$(4)), UCase(param$(5)), 0, runtimhiba%)
      If runtimhiba% = 1 Then Exit Sub
    Else
      komt(komdb%).komobj = Space$(4)
    End If
  Loop While Not EOF(komfi)
  Close komfi
  Komind.MSFlexGrid1.Rows = komadatdb% + 1
  Komind.MSFlexGrid1.Cols = 3
  Komind.MSFlexGrid1.ColWidth(0) = 2000
  Komind.MSFlexGrid1.ColWidth(1) = 1050
  Komind.MSFlexGrid1.ColWidth(2) = 1050
  Komind.MSFlexGrid1.TextMatrix(0, 1) = langmodul(80)
  Komind.MSFlexGrid1.TextMatrix(0, 2) = langmodul(81)
  xac% = Komind.MSFlexGrid1.RowHeight(0)
  For w1% = 1 To komadatdb%
    Komind.MSFlexGrid1.TextMatrix(w1%, 0) = Trim$(komt(w1%).komszov)
    Komind.MSFlexGrid1.TextMatrix(w1%, 1) = ""
    Komind.MSFlexGrid1.TextMatrix(w1%, 2) = ""
    xac% = xac% + Komind.MSFlexGrid1.RowHeight(w1%)
  Next
  Komind.MSFlexGrid1.Height = xac% + 20
  Komind.Label1.Top = Komind.MSFlexGrid1.Top + xac% + 20
  Komind.Label1.BackColor = RGB(120, 0, 0)
  Komind.Label1.ForeColor = RGB(255, 255, 255)
  Komind.Label1.Width = 3070
  Komind.Text1.Top = Komind.Label1.Top + 220
  Komind.Text1.Height = 250
  Komind.Text1.Width = 3070
  Komind.Command3.Left = Komind.Label1.Left + Komind.Label1.Width
  Komind.Command3.Width = Komind.MSFlexGrid1.Width - Komind.Label1.Width
  Komind.Command3.Top = Komind.Label1.Top
  Komind.Command3.Height = 240
  For w1% = 1 To kommenudb%
    If w1% = 1 Then Set lis = Komind.List1
    If w1% = 2 Then Set lis = Komind.List2
    If w1% = 3 Then Set lis = Komind.List3
    If w1% = 4 Then Set lis = Komind.List4
    If w1% = 5 Then Set lis = Komind.List5
    If w1% = 6 Then Set lis = Komind.List6
    For i9& = 1 To 10: param$(i9&) = "": Next
    b$ = Trim$(komt(w1% + komadatdb%).komszov)
    paramdb% = 15
    Call linpar(b$, param$(), "&", paramdb%)
    For w2% = 1 To paramdb%
      lis.AddItem param$(w2%)
    Next
    lis.ListIndex = 0
    'lis.Height = paramdb% * 220
    If Not w1% = 6 Then
       lis.Height = paramdb% * 200 + 20
    Else
       lis.Height = paramdb% * 200 + 200
    End If
  Next
  Komind.Caption = kcimsor$
  Komind.Show vbModal
  Unload Komind
  komment$ = ""
  For w2% = 1 To komdb%
    Select Case komt(w2%).komkod
      Case "2"
        If Trim$(komt(w2%).komtol) <> "" And Trim$(komt(w2%).komig) <> "" Then
          If InStr(komt(w2%).komatr, "D") <> 0 Then
            dzu1$ = datki(Trim(komt(w2%).komtol))
            dzu2$ = datki(Trim(komt(w2%).komig))
            komment$ = komment$ + Trim(komt(w2%).komszov) + ":" + dzu1$ + "-" + dzu2$ + ", "
          Else
            komment$ = komment$ + Trim(komt(w2%).komszov) + ":" + Trim(komt(w2%).komtol) + "-" + Trim(komt(w2%).komig) + ", "
          End If
        End If
      Case "1"
        If Trim$(komt(w2%).komtol) <> "" Then
          If InStr(komt(w2%).komatr, "D") <> 0 Then
            dzu1$ = datki(Trim(komt(w2%).komtol))
            komment$ = komment$ + Trim(komt(w2%).komszov) + ":" + dzu1$ + ", "
          Else
            komment$ = komment$ + Trim(komt(w2%).komszov) + ":" + Trim(komt(w2%).komtol) + ", "
          End If
        End If
      Case "3"
        paramdb% = 10
        menuk$ = komt(w2%).komszov
        Call linpar(menuk$, param$(), "&", paramdb%)
        komment$ = komment$ + Trim$(param$(komt(w2%).kommnv)) + ", "
      Case Else
    End Select
  Next
  Exit Sub
hibakez:
  Call mess(langmodul(158), 1, 0, langmodul(159), valasz%)
  runtimhiba% = 1
End Sub

Public Sub gomb(gombsor$, gombsi%, ttop&, lleft&, ir$)
  '--- gombsor megjelenítése
  '--- gombsor$-ban a gombok szövege & jellel elválasztva
  '--- gombs% a lenyomott gomb sorszáma
  '--- ir$=V vizszintes, ir$=F fügõleges
  '--- ttop,lleft a bal felsõ sarok twip-ben
  Dim param$(10)
  Dim gom As CommandButton
  paramdb% = 10
  Call linpar(gombsor$, param$(), "&", paramdb%)
  Gombok.Left = form1.Left + lleft&
  Gombok.Top = form1.Top + ttop&
  If ir$ = "V" Then
    bal% = 30
    For w1% = 1 To paramdb%
      Set gom = Gombok.Command1(w1% - 1)
      ho% = Len(Trim(param$(w1%)))
      gom.Caption = param$(w1%)
      gom.Left = bal%
      gom.Top = 30
      gom.Visible = True
      'gom.Width = ho% * 120
      gom.Width = Gombok.TextWidth("  " + param$(w1%))
      If w1% = paramdb% Then gom.Cancel = True
      'bal% = bal% + ho% * 120 + 50
      bal% = bal% + gom.Width + 50
    Next
    Gombok.Width = bal% + 10
    Gombok.Height = gom.Height + 80
  Else
    tet% = 30
    ho% = 1
    For w1% = 1 To paramdb%
      If ho% < Len(param$(w1%)) Then ho% = Len(param$(w1%))
    Next
    ho% = ho% * 120
    For w1% = 1 To paramdb%
      Set gom = Gombok.Command1(w1% - 1)
      gom.Caption = param$(w1%)
      gom.Left = 30
      gom.Top = tet%
      gom.Visible = True
      gom.Width = ho%
      'gom.FontBold = True
      tet% = tet% + gom.Height
    Next
    Gombok.Height = tet% + 30
    Gombok.Width = ho% + 60
  End If
  For w1% = paramdb% + 1 To 10
    Set gom = Gombok.Command1(w1% - 1)
    gom.Visible = False
  Next
  Gombok.Show vbModal
  gombsi% = gombs%
End Sub
Public Sub eankodbe(mezo$)
  Eanbe.Text1.Text = ""
  Eanbe.Show vbModal
  mezo$ = Left(Trim(Eanbe.Text1.Text) + Space(13), 13)
End Sub

Public Sub inrut1(iadatnev$, iattri$, imagyarazat$, mezo$)
  Inrutin.Caption = iadatnev$
  Inrutin.Label1.Caption = iadatnev$
  Inrutin.Label2.Caption = imagyarazat$
  atri$ = Mid$(iattri$, 3)
  mh% = xval(Mid$(iattri$, 1, 2))
  Do
    Inrutin.Text1.MaxLength = mh%
    Inrutin.Text1.Text = mezo$
    Inrutin.Text1.SelStart = Len(mezo$) + 1
    Inrutin.Show vbModal
    mezo$ = Inrutin.Text1.Text
    If mezo$ = "" Then Exit Sub
    Call kodvizsg(mezo$, atri$, khiba%, mh%)
  Loop While khiba = 1
  If InStr(atri$, "J") > 0 Then
    mezo$ = Right$(Space$(mh%) + mezo$, mh%)
  Else
    mezo$ = Left$(mezo$ + Space$(mh%), mh%)
  End If
End Sub

Public Sub inrut(aso%, mezo$)
  Inrutin.Label1.Caption = langmodul(82) + " " + LCase(Trim$(ADATAB(aso%).adatnev)) + ":"
  Do
    mh% = ADATAB(aso%).adatho
    Inrutin.Label1.Caption = langmodul(82) + " " + LCase(Trim$(ADATAB(aso%).adatnev)) + ":"
    Inrutin.Label2.Caption = Trim(MAGYARAZAT$(aso%)) + " (max.:" + Trim(Str(mh%)) + " karakter)"
    Inrutin.Text1.MaxLength = mh%
    Inrutin.Text1.Text = mezo$
    Inrutin.Text1.SelStart = Len(mezo$) + 1
    Inrutin.Show vbModal
    mezo$ = Inrutin.Text1.Text
    If mezo$ = "" Then Exit Sub
    atri$ = ADATAB(aso%).attr
    mh% = ADATAB(aso%).adatho
    Call kodvizsg(mezo$, atri$, khiba%, mh%)
  Loop While khiba% = 1
  mh% = ADATAB(aso%).adatho
  If InStr(atri$, "J") > 0 Then
    mezo$ = Right$(Space$(mh%) + mezo$, mh%)
  Else
    mezo$ = Left$(mezo$ + Space$(mh%), mh%)
  End If
End Sub

Public Sub kodvizsg(mezo$, atri$, khiba%, mh%)
  khiba% = 0
  mezo$ = Trim$(mezo$)
  ah% = Len(mezo$)
'--- jobbra illesztés
  If InStr(atri$, "NZJ") > 0 Then
    mezo$ = Right$("000000000" + mezo$, mh%)
  End If
'--- nagybetû
  If InStr(atri$, "U") > 0 Then
    mezo$ = UCase(mezo$)
  End If
'--- kötelezõ hossz
  If InStr(atri$, "K") > 0 Then
    '--- kötelezõ hossz
    If mh% <> ah% Then khiba% = 1: Exit Sub
  End If
'--- numerikus
  If InStr(atri$, "N") > 0 Then
    mas$ = "1234567890"
    If InStr(atri$, "T") > 0 Then mas$ = mas$ + ",."
    If InStr(atri$, "-") > 0 Then mas$ = mas$ + "-"
    If ah% > 0 Then
      tdb% = 0: mdb% = 0
      For w8% = 1 To ah%
        If Mid$(mezo$, w8%, 1) = "." Or Mid$(mezo$, w8%, 1) = "," Then tdb% = tdb% + 1
        If Mid$(mezo$, w8%, 1) = "-" Then mdb% = mdb% + 1
        If InStr(mas$, Mid$(mezo$, w8%, 1)) = 0 Then khiba% = 1: Exit Sub
      Next
      If tdb% > 1 Or mdb% > 1 Then khiba% = 1: Exit Sub
      If mdb% = 1 And Mid$(mezo$, 1, 1) <> "-" Then khiba% = 1: Exit Sub
    End If
  End If
'--- dátum
  If InStr(atri$, "D") > 0 Then
    '--- dátum adat vizsgálata
    If ah% <> 6 Then khiba% = 1: Exit Sub
    hon% = xval(Mid$(mezo$, 3, 2))
    If hon% = 4 Or hon% = 6 Or hon% = 9 Or hon% = 11 Then nap% = 30 Else nap% = 31
    If hon% = 2 Then nap% = 29
    If hon% > 12 Or hon% = 0 Then khiba% = 1: Exit Sub
    If xval(Mid$(mezo$, 5, 2)) > nap% Or xval(Mid$(mezo$, 5, 2)) = 0 Then khiba% = 1
  End If
End Sub

Public Function xkonver(k$)
  '--- magyar szöveg konvertálása az angol abc nagybetûire
  q1$ = "áÁéÉíÍóÓöÖõÕúÚüÛ"
  q2$ = "AAEEIIOOOOOOUUUU"
  z$ = UCase$(k$)
  i1% = Len(k$)
  For i3% = 1 To 16
    kk$ = Mid$(q1$, i3%, 1)
    Do
      i2% = InStr(z$, kk$)
      If i2% <> 0 Then
        Mid$(z$, i2%, 1) = Mid$(q2$, i3%, 1)
      End If
    Loop While i2% <> 0
  Next
  xkonver = z$
End Function

Public Function mkonver(k$)
  '--- magyar szöveg konvertálása ascii kódos rendezéshez
  q1$ = "AÁBCDEÉFGHIÍJKLMNOÓÖÕPQRSTUÚÜÛVWXYZaábcdeéfghiíjklmnoóöõpqrstuúüûvwxyz"
  q2$ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abc"
  z$ = k$
  i1% = Len(k$)
  For i2% = 1 To i1%
    For i3% = 1 To Len(q1$)
      If Mid$(z$, i2%, 1) = Mid$(q1$, i3%, 1) Then
        Mid$(z$, i2%, 1) = Mid$(q2$, i3%, 1): Exit For
      End If
    Next
  Next
  mkonver = z$
End Function
Public Sub hozzad(k$, kep%, hop%, oo@, tizedh%)
  '--- adott mezõ növelése egy adott értékkel
  If kep% <> 0 And hop% <> 0 And oo@ <> 0 Then
    oszw@ = xval(Mid$(k$, kep%, hop%)) + oo@
    Select Case tizedh%
      Case 0: oszy$ = Right$(Space$(hop%) + Format$(oszw@, "#############0"), hop%)
      Case 1: oszy$ = Right$(Space$(hop%) + Format$(oszw@, "###########0.0"), hop%)
      Case 2: oszy$ = Right$(Space$(hop%) + Format$(oszw@, "##########0.00"), hop%)
      Case 3: oszy$ = Right$(Space$(hop%) + Format$(oszw@, "#########0.000"), hop%)
    End Select
    Mid$(k$, kep%, hop%) = oszy$
  End If
End Sub
Public Sub kivvon(k$, kep%, hop%, oo@, tizedh%)
  '--- adott mezõ csökkentése egy adott értékkel
  If kep% <> 0 And hop% <> 0 And oo@ <> 0 Then
    oszw@ = xval(Mid$(k$, kep%, hop%)) - oo@
    Select Case tizedh%
      Case 0: oszy$ = Right$(Space$(hop%) + Format$(oszw@, "#############0"), hop%)
      Case 1: oszy$ = Right$(Space$(hop%) + Format$(oszw@, "###########0.0"), hop%)
      Case 2: oszy$ = Right$(Space$(hop%) + Format$(oszw@, "##########0.00"), hop%)
      Case 3: oszy$ = Right$(Space$(hop%) + Format$(oszw@, "#########0.000"), hop%)
    End Select
    Mid$(k$, kep%, hop%) = oszy$
  End If
End Sub

Public Sub auwini(inihiba%)
  On Error GoTo inihib
  inihiba% = 0
  cmfi = FreeFile
  Open "c:\auwin\auw-cmv.par" For Binary Shared As #cmfi
  cmfm& = LOF(cmfi)
  If cmfm& > 0 Then
    cick$ = " "
    Get #cmfi, 1, cick$
    If cick$ = "1" Then cimvektorvan = 1 Else cimvektorvan = 0
  Else
    cimvektorvan = 0
  End If
  Close cmfi
  exfi = FreeFile
  Open "c:\auwset\auw" + terminal$ + task$ + ".auw" For Binary As #exfi
  exfm& = LOF(exfi)
  Close exfi
  If exfm& = 0 Then
    Call mess(langmodul(83), 1, 0, langmodul(84), valasz%)
    'MsgBox langmodul(83), 48, langmodul(84)
    inihiba% = 1
  Else
    exfi = FreeFile
    Open "c:\auwset\auw" + terminal$ + task$ + ".auw" For Input As #exfi
    Line Input #exfi, ugyintezo$
    Line Input #exfi, ugyintneve$
    Line Input #exfi, jogok$
    Line Input #exfi, regszam$
    Line Input #exfi, auditorutvonal$
    Line Input #exfi, listautvonal$
    Line Input #exfi, programutvonal$
    Line Input #exfi, rootutvonal$
    Line Input #exfi, licensztulaj$
    Line Input #exfi, cegneve$
    Close exfi
    inihiba% = 0
    Exit Sub
  End If
inihib:
  inihiba% = 1
End Sub
Public Sub lancra(dbxazon, setneve$, ownerrec$, aktucim&, akturec$)
  '--- objektum felfûzése next és prior láncra és owner visszaírása
  '--- dbxazon   az adatbázis neve
  '--- setnev    a set neve
  '--- ownerrec  az owner rekord
  '--- aktucim   az aktuális memberrekord címe
  '--- akturec   az aktuális member rekord
  If ownerrec$ = "" Or akturec$ = "" Or aktucim& = 0 Then Exit Sub
  seti% = 0
  For i% = 1 To setdb
    If Trim$(setneve$) = Trim$(SETTAB(i%).setnev) Then
      seti% = i%
      Exit For
    End If
  Next
  If seti% = 0 Then Exit Sub
  '--- next pointer
  mupo$ = Mid$(ownerrec$, SETTAB(seti%).rootpoz, 10)
  '--- owner next pointere törölve
  Mid$(ownerrec$, SETTAB(seti%).rootpoz, 10) = Right$(Space$(10) + Str$(aktucim&), 10)
  '--- owner visszaírás
  ownaz$ = OBJTAB(SETTAB(seti).robsorsz).obaz
  Call dbxki(ownaz$, ownerrec$, ";", "", "", hiba%)
  '--- aktuális next pointere arra mutat, amire az owner mutatott
  Mid$(akturec$, SETTAB(seti%).nextpoz, 10) = mupo$
  memaz$ = OBJTAB(SETTAB(seti%).obsorsz).obaz
  Call dbxki(memaz$, akturec$, ";", "", "", hiba%)
  '--- prior pointer
  'If setneve$ = "TSZAMLA" Or setneve$ = "KSZAMLA" Or setneve$ = "KFRKFRG" Then
  'If setneve$ <> "PSZLKFRG" Then
  '  dbfi = FreeFile
  '  Open auditorutvonal$ + dbxazon + ".dbx" For Binary Shared As #dbfi
  '  '--- a next rekord beolvasása és az aktcim beírása a prior pointerbe
  '  rcim& = Val(mupo$)
  '  If rcim& > 0 Then
  '    Seek #dbfi, rcim& + 9
  '    w1% = OBJTAB(SETTAB(seti%).obsorsz).rekhossz
  '    t1rec$ = Space(w1%): Get #dbfi, , t1rec$
  '    Mid$(t1rec$, SETTAB(seti%).priorpoz, 10) = Right$(Space$(10) + Str$(aktucim&), 10)
  '    Put #dbfi, rcim& + 9, t1rec$
  '  End If
  '  Close dbfi
  'End If
End Sub

Public Sub lancrol(dbxazon, setneve$, ownerrec$, akturec$)
  '--- objektum leválasztása next és prior láncról
  '--- objektum felfûzése next és prior láncra és owner visszaírása
  '--- dbxazon   az adatbázis neve
  '--- setnev    a set neve
  '--- ownerrec  az owner rekord
  '--- akturec   az aktuális member rekord
  If ownerrec$ = "" Or akturec$ = "" Then Exit Sub
  For i% = 1 To setdb
    If Trim$(setneve$) = Trim$(SETTAB(i%).setnev) Then
      seti% = i%
      Exit For
    End If
  Next
  '--- next és prior pointerek
  nepo$ = Mid$(akturec$, SETTAB(seti%).nextpoz, 10): nepoi& = xval(nepo$)
  'prpo$ = Mid$(akturec$, SETTAB(seti%).priorpoz, 10): prpoi& = xval(prpo$)
  '--- lefûzés next pointer
  ownaz$ = OBJTAB(SETTAB(seti).robsorsz).obaz
  Call dbxki(ownaz$, ownerrec$, ";", "", "", hiba%)
  If prpoi& = 0 Then
    '--- nincs prior (elõzõ rekord)
    '--- az ownerben lévõ next helyébe az aktuális next tartalma
    Mid$(ownerrec$, SETTAB(seti%).rootpoz, 10) = nepo$
    ownaz$ = OBJTAB(SETTAB(seti).robsorsz).obaz
    Call dbxki(ownaz$, ownerrec$, ";", "", "", hiba%)
  Else
    '--- van prior
    dbfi = FreeFile
    Open auditorutvonal$ + dbxazon + ".dbx" For Binary Shared As #dbfi
    '--- a megelõzõ rekord beolvasása
    Seek #dbfi, prpoi& + 9
    w1% = OBJTAB(SETTAB(seti%).obsorsz).rekhossz
    t1rec$ = Space(w1%): Get #dbfi, , t1rec$
    '--- a megelõzõ next pointerébe az aktuális next tartalma
    Mid$(t1rec$, SETTAB(seti%).nextpoz, 10) = nepo$
    Put #dbfi, prpoi& + 9, t1rec$
    Close dbfi
  End If
  '--- lefûzés prior pointer
  If nepoi& = 0 Then
    '--- nincs következõ
  Else
    dbfi = FreeFile
    Open auditorutvonal$ + dbxazon + ".dbx" For Binary Shared As #dbfi
    '--- következõ beolvasása
    Seek #dbfi, nepoi& + 9
    w1% = OBJTAB(SETTAB(seti%).obsorsz).rekhossz
    '--- következõ priorba az aktuális prior
    t1rec$ = Space(w1%): Get #dbfi, , t1rec$
    Mid$(t1rec$, SETTAB(seti%).priorpoz, 10) = prpo$
    Put #dbfi, nepoi& + 9, t1rec$
    Close dbfi
  End If
  '--- az aktuális pointereinek törlése
  Mid$(akturec$, SETTAB(seti%).nextpoz, 10) = Space$(10)
  Mid$(akturec$, SETTAB(seti%).priorpoz, 10) = Space$(10)
  memaz$ = OBJTAB(SETTAB(seti%).obsorsz).obaz
  '--- aktuális visszaírás
  Call dbxki(memaz$, akturec$, ";", "", "", hiba%)
End Sub

Public Function dtm(dat$)
  ev% = Val(Left$(dat$, 2)) + 60
  dtm = Right$("00" + Trim$(Str$(ev%)), 2) + Mid$(dat$, 3, 4)
End Function

Public Function novdat(dat$)
  ev% = Val(Mid$(dat$, 1, 2))
  ho% = Val(Mid$(dat$, 3, 2))
  na% = Val(Mid$(dat$, 5, 2))
  Select Case ho%
    Case 2
      If ev% Mod 4 = 0 Then maxna% = 29 Else maxna% = 28
    Case 1, 3, 5, 7, 8, 10, 12
      maxna% = 31
    Case 4, 6, 9, 11
      maxna% = 30
    Case Else
  End Select
  na% = na% + 1
  If na% > maxna% Then
    na% = 1
    ho% = ho% + 1
    If ho% > 12 Then
      ev% = ev% + 1: ho% = 1
    End If
  End If
  novdat = Right$("00" + Trim$(Str$(ev%)), 2) + Right$("00" + Trim$(Str$(ho%)), 2) + Right$("00" + Trim$(Str$(na%)), 2)
End Function
Public Function csokdat$(dat$)
  ev% = Val(Mid$(dat$, 1, 2))
  ho% = Val(Mid$(dat$, 3, 2))
  na% = Val(Mid$(dat$, 5, 2))
  Select Case ho%
    Case 3
      If ev% Mod 4 = 0 Then maxna% = 29 Else maxna% = 28
    Case 1, 2, 4, 6, 8, 9, 11
      maxna% = 31
    Case 5, 7, 10, 12
      maxna% = 30
    Case Else
  End Select
  na% = na% - 1
  If na% > 0 Then
  Else
    na% = maxna%
    ho% = ho% - 1
    If ho% > 0 Then
    Else
      ev% = ev% - 1: ho% = 12
    End If
  End If
  csokdat$ = Right$("00" + Trim$(Str$(ev%)), 2) + Right$("00" + Trim$(Str$(ho%)), 2) + Right$("00" + Trim$(Str$(na%)), 2)
End Function

Public Function csokho$(dat$)
  ev% = Val(Mid$(dat$, 1, 2))
  ho% = Val(Mid$(dat$, 3, 2))
  na% = Val(Mid$(dat$, 5, 2))
  Select Case ho%
    Case 3
      If ev% Mod 4 = 0 Then maxna% = 29 Else maxna% = 28
    Case 1, 2, 4, 6, 8, 9, 11
      maxna% = 31
    Case 5, 7, 10, 12
      maxna% = 30
    Case Else
  End Select
  ho% = ho% - 1
  If ho% > 0 Then
  Else
    ev% = ev% - 1: ho% = 12
    If na% > maxna% Then na% = maxna%
  End If
  csokho$ = Right$("00" + Trim$(Str$(ev%)), 2) + Right$("00" + Trim$(Str$(ho%)), 2) + Right$("00" + Trim$(Str$(na%)), 2)
End Function

Public Function csokev$(dat$)
  ev% = Val(Mid$(dat$, 1, 2)) - 1
  csokev$ = Right$("00" + Trim$(Str$(ev%)), 2) + Mid$(dat$, 3, 4)
End Function

Public Function datki(dat$)
  If Trim(dat$) = "" Then datki = Space(10): Exit Function
  If Mid$(dat$, 1, 2) > "39" Then evv$ = "19" Else evv$ = "20"
  If langhun% > 1 Then
    datki = Mid$(dat$, 5, 2) + "." + Mid$(dat$, 3, 2) + "." + evv$ + Mid$(dat$, 1, 2)
  Else
    datki = evv$ + Mid$(dat$, 1, 2) + "." + Mid$(dat$, 3, 2) + "." + Mid$(dat$, 5, 2)
  End If
End Function

Public Function ertszamx(mez$, hossz%, tizedes)
  If tizedes = 0 Then
    fst$ = "##########,##0"
  Else
    fst$ = "########,##0." + String(tizedes, "0")
  End If
  ertszamx = Right$(Space$(hossz%) + Format(xval(mez$), fst$), hossz%)
End Function

Public Function ertszamy(mez$, hossz%, tizedes)
  If xval(mez$) = 0 Then ertszamy = Space(hossz%): Exit Function
  If tizedes = 0 Then
    fst$ = "##########,##0"
  Else
    fst$ = "########,##0." + String(tizedes, "0")
  End If
  ertszamy = Right$(Space$(hossz%) + Format(xval(mez$), fst$), hossz%)
End Function

Public Function ertszamz(mez$, hossz%, tizedes)
  If xval(mez$) = 0 Then ertszamz = Space(hossz%): Exit Function
  If tizedes = 0 Then
    fst$ = "##############0"
  Else
    fst$ = "############0." + String(tizedes, "0")
  End If
  ertszamz = Right$(Space$(hossz%) + Format(xval(mez$), fst$), hossz%)
End Function

Public Function ertszamw(mez$, hossz%, tizedes)
  If tizedes = 0 Then
    fst$ = "#############0"
  Else
    fst$ = "###########0." + String(tizedes, "0")
  End If
  zzzz$ = Right$(Space$(hossz%) + Format(xval(mez$), fst$), hossz%)
  For i19% = 1 To hossz%
    If Mid$(zzzz$, i19%, 1) = "," Then Mid$(zzzz$, i19%, 1) = "."
  Next
  ertszamw = zzzz$
End Function

Public Function ertszam(mez$, hossz%, tizedes)
  If tizedes = 0 Then
    fst$ = "#############0"
  Else
    fst$ = "###########0." + String(tizedes, "0")
  End If
  ertszam = Right$(Space$(hossz%) + Format(xval(mez$), fst$), hossz%)
End Function

Public Function jobbra(mez$, hossz%)
  jobbra = Right$(Space$(hossz%) + mez$, hossz%)
End Function

Public Function novel(irec$, k%, h%)
  s& = xval(Mid$(irec$, k%, h%)) + 1
  novel = Right$("000000000000" + Trim$(Str$(s&)), h%)
End Function

Public Function ini123%()
  lgfi = FreeFile
  Open "c:\auwin\log123.par" For Binary As #lgfi
  lgmer& = LOF(lgfi)
  Close lgfi
  If lgmer& > 0 Then ini123% = 1 Else ini123% = 0
End Function

Public Function bankir$(bsz$)
  Select Case langhun%
    Case 4
      '--- szerb
    Case Else
      bankir$ = Mid$(bsz$, 1, 8) + "-" + Mid$(bsz$, 9, 8) + "-" + Mid$(bsz$, 17, 8)
  End Select
End Function

Public Function jogell%(jogkod%)
  '1-törzskarbantartás
  '2-számvitel hozzáférés
  '3-eszköz hozzáférés
  '4-készlet hozzáférés
  '5-bér hozzáférés
  '6-vegyes bizonylat  (SKN)
  '7-vevõ       (SKN)
  '8-szállító   (SKN)
  '9-pénztár    (SKN)
  '10-bank      (SKN)
  '11-számlázás (SKN)
  '12-feladás
  '13-zárás
  '14-ujrafeldolgozás
  '15-újraszervezés
  'kimenõ=0 nincs jog, 1 bizonylat jog, 2 sztornó jog
  If (jogkod% >= 1 And jogkod% <= 5) Or jogkod% >= 12 Then
    If Mid$(jogok$, jogkod%, 1) = " " Then Mid$(jogok$, jogkod%, 1) = "I"
    If Mid$(jogok$, jogkod%, 1) = "N" Then jogell% = 0: Exit Function
    If Mid$(jogok$, jogkod%, 1) = "I" Then jogell% = 1: Exit Function
  Else
    If jogkod% >= 6 And jogkod% <= 11 Then
      If Mid$(jogok$, jogkod%, 1) = " " Then Mid$(jogok$, jogkod%, 1) = "S"
      If Mid$(jogok$, jogkod%, 1) = "N" Then jogell% = 0: Exit Function
      If Mid$(jogok$, jogkod%, 1) = "K" Then jogell% = 1: Exit Function
      If Mid$(jogok$, jogkod%, 1) = "S" Then jogell% = 2: Exit Function
    End If
  End If
  jogell% = 0
End Function

Public Sub hetinap(dat$, sorsz%, napnev$)
  '--- hét napjának meghatározása
  sorsz% = WeekDay(datki(dat$), vbMonday)
  Select Case sorsz%
    Case 1: napnev$ = "Hétfõ"
    Case 2: napnev$ = "Kedd"
    Case 3: napnev$ = "Szerda"
    Case 4: napnev$ = "Csütörtök"
    Case 5: napnev$ = "Péntek"
    Case 6: napnev$ = "Szombat"
    Case 7: napnev$ = "Vasárnap"
    Case Else: napnev$ = ""
  End Select
End Sub

Public Function seqolv$(qvart$, filszi, rhossz&)
  '--- index szerinti olvasás
  ziz1& = Asc(Mid(qvart$, 1, 1))
  ziz2& = Asc(Mid(qvart$, 2, 1)) * 256&
  ziz3& = Asc(Mid(qvart$, 3, 1)) * 65536
  ziz4& = Asc(Mid(qvart$, 4, 1)) * 16777216
  rcim& = ziz1& + ziz2& + ziz3& + ziz4&
  rekko$ = Space(rhossz&)
  Get #filszi, rcim& + 9, rekko$
  seqolv = rekko$
End Function

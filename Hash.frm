VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Hash 
   BackColor       =   &H8000000A&
   Caption         =   "Keresés"
   ClientHeight    =   4560
   ClientLeft      =   2508
   ClientTop       =   2472
   ClientWidth     =   7908
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   7908
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H0055D7F7&
      Caption         =   "Információk"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3170
      Left            =   3360
      TabIndex        =   12
      Top             =   420
      Visible         =   0   'False
      Width           =   3840
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   7.8
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2772
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3612
      End
   End
   Begin VB.CommandButton Command26 
      Height          =   400
      Left            =   7320
      Picture         =   "Hash.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   444
   End
   Begin VB.CommandButton Command23 
      Height          =   400
      Left            =   7320
      Picture         =   "Hash.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   444
   End
   Begin VB.CommandButton Command22 
      Height          =   400
      Left            =   7320
      Picture         =   "Hash.frx":0EC4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   444
   End
   Begin VB.CommandButton Command25 
      Height          =   400
      Left            =   7320
      Picture         =   "Hash.frx":1B06
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   444
   End
   Begin VB.CommandButton Command24 
      Height          =   400
      Left            =   7320
      Picture         =   "Hash.frx":2748
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   444
   End
   Begin VB.CommandButton Command21 
      Height          =   400
      Left            =   7320
      Picture         =   "Hash.frx":338A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   444
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   4170
      Width           =   1332
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   4170
      Width           =   1332
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Mégsem"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4150
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Választ"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4150
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.2
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   7692
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7092
      _ExtentX        =   12510
      _ExtentY        =   6160
      _Version        =   327680
      FixedCols       =   0
      BackColorSel    =   12648384
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      ScrollBars      =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "Hash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rh&, ih&, dbxn$, indn$, elsosor&, aktsor&, darab&, rc&, sorbak&, hc&
Dim oobcim&, oobind&, oobnext&, ob%
Dim statusz% '--- =0 teljes halmaz, =1 szûkített halmaz
Dim teldarab&, colwi%(200), colwimod%, fontsi%
Dim resa$(30000), resm&(30000), res&(30000), resdb&  '--- rekordok fizikai sorszáma
Dim resaf(30000)
Dim hst1$, hst2$, hst3$, hck%, belep%, lapmeret%, hsmezohossz%

Private Sub Command21_Click()
  If colwimod% = 0 Then
    For izu% = 1 To MSFlexGrid1.Cols
      hxh% = Hash.TextWidth(Trim(MSFlexGrid1.TextMatrix(0, izu% - 1)))
      hxk% = Hash.TextWidth(String(kmho%(izu%), " "))
      If hxh% > hxk% Then MSFlexGrid1.ColWidth(izu% - 1) = hxh% + 100 Else MSFlexGrid1.ColWidth(izu% - 1) = hxk% + 100
    Next
    colwimod% = 1
  Else
    For izu% = 1 To MSFlexGrid1.Cols
      MSFlexGrid1.ColWidth(izu% - 1) = colwi(izu%)
      colwimod% = 0
    Next
  End If
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Command22_Click()
  If fontsi% = 0 Then fontsi% = 8
  If fontsi% < 10 Then fontsi% = fontsi% + 1: MSFlexGrid1.Font.Size = fontsi%
  Select Case fontsi%
    Case 6: lapmeret% = 17
    Case 7: lapmeret% = 15
    Case 8: lapmeret% = 14
    Case 9: lapmeret% = 12
    Case 10: lapmeret% = 11
    Case Else
  End Select
  MSFlexGrid1.Rows = lapmeret% + 1
  Call MSFlexGrid1_Click
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Command23_Click()
  If fontsi% = 0 Then fontsi% = 8
  If fontsi% > 6 Then fontsi% = fontsi% - 1: MSFlexGrid1.Font.Size = fontsi%
  Select Case fontsi%
    Case 6: lapmeret% = 17
    Case 7: lapmeret% = 15
    Case 8: lapmeret% = 14
    Case 9: lapmeret% = 12
    Case 10: lapmeret% = 11
    Case Else
  End Select
  MSFlexGrid1.Rows = lapmeret% + 1
  Call MSFlexGrid1_Click
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Command24_Click()
  Hash.Left = 1
  Hash.Width = 12250
  Frame3.Left = 7690
End Sub

Private Sub Command25_Click()
  Hash.Left = 1000
  Hash.Width = 10440
  Frame3.Left = 5880
End Sub

Private Sub Command26_Click()
  If infolapbe% = 0 Then infolapbe% = 1: Frame3.Visible = True Else infolapbe% = 0: Frame3.Visible = False
  If infolapbe% = 1 Then Call infomutat(keresobj$, r$, "Hash")
  DoEvents
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Form_Resize()
  If Hash.Width < 5000 Then Hash.Width = 5000
  MSFlexGrid1.Width = Hash.Width - 825
  Command1.Left = Hash.Width - 2960
  Command2.Left = Hash.Width - 1520
  Command21.Left = Hash.Width - 600
  Command22.Left = Hash.Width - 600
  Command23.Left = Hash.Width - 600
  Command24.Left = Hash.Width - 600
  Command25.Left = Hash.Width - 600
  Command26.Left = Hash.Width - 600
  Text1.Width = Hash.Width - 300
End Sub


Private Sub Command1_Click()
  If darab& > 0 Then
    talalat% = 1
    gombsorszam% = 6
    OBJTAB(ob%).obcim = oobcim&
    OBJTAB(ob%).obind = oobind&
    OBJTAB(ob%).obnex = oobnext&
    toprox& = elsosor&: aktrox& = aktsor&: aktdarab& = darab&: tabstatusx% = statusz%
    aktext$ = Text1.Text
    Call langclos
    Hash.Hide
  Else
    Call Command2_Click
  End If
End Sub

Private Sub Command2_Click()
  talalat% = 0
  toprox& = elsosor&: aktrox& = aktsor&: aktdarab& = darab&: tabstatusx% = statusz%
  aktext$ = Text1.Text
  Call langclos
  Hash.Hide
End Sub

Private Sub Form_Activate()
  '--- form aktíválása
  '--- kezdeti értékek beállítása
  '--- nyitó táblázat megmutatása
  Text1.Top = 3680
  ob% = obsorszama(keresobj$)
  hck% = ADATAB(OBJTAB(ob%).hashcod).adatkp
  hsmezohossz% = ADATAB(OBJTAB(ob%).hashcod).adatho
  dbxn$ = dbxneve(keresobj$)
  w1% = OBJTAB(ob%).obi(1)
  indn$ = RTrim$(INDTAB(w1%).indnev)
  ih& = ADATAB(INDTAB(w1%).adatsorsz).adatho + 5
  rh& = OBJTAB(ob%).rekhossz
  lapmeret% = 14
  '--- hash kódok beolvasása
  For izu% = 1 To MSFlexGrid1.Cols
    colwi(izu%) = MSFlexGrid1.ColWidth(izu% - 1)
  Next
  hsfi = FreeFile
  Open auditorutvonal$ + "auw-" + keresobj$ + ".hs1" For Binary Shared As #hsfi
  hc& = LOF(hsfi)
  hst1$ = Space(hc&): Get #hsfi, , hst1$
  Close hsfi
  hsfi = FreeFile
  Open auditorutvonal$ + "auw-" + keresobj$ + ".hs2" For Binary Shared As #hsfi
  hst2$ = Space(hc&): Get #hsfi, , hst2$
  Close hsfi
  hsfi = FreeFile
  Open auditorutvonal$ + "auw-" + keresobj$ + ".hs3" For Binary Shared As #hsfi
  hst3$ = Space(hc&): Get #hsfi, , hst3$
  Close hsfi
  statusz% = tabstatusx%
  If aktdarab& = 0 Then
    If hc& > 0 Then
      darab& = hc&
      elsosor& = 1: aktsor& = 1
      MSFlexGrid1.Row = 1
      statusz% = 0
      Call megmutat(dbxn$, indn$)
    End If
  Else
    darab& = aktdarab&
    elsosor& = toprox&
    aktsor& = aktrox&
    aktsor& = aktsor& + 1: If aktsor& > darab& Then aktsor& = darab&
    If aktsor& > elsosor& + lapmeret% - 1 Then elsosor& = elsosor& + 1
    belep% = 1
    Text1.Text = aktext$
    belep% = 0
    Call megmutat(dbxn$, indn$)
  End If
  Command25_Click
  If autoinfo = 1 Then infolapbe = 0: Call Command26_Click
  If Text1.Visible = True Then Text1.SetFocus
End Sub
Private Sub megmutat(dbxn$, indn$)
  '--- aktuális táblázat megmutatas kereso.msflexgrid1-ben
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  If darab& - elsosor& + 1 > 30000 Then
    sdb% = 30000
  Else
    sdb% = darab& - elsosor& + 1
  End If
  If sdb% > lapmeret% Then sdb% = lapmeret%
  For i& = 1 To sdb%
    inex& = elsosor& + i& - 1
    If statusz% = 0 Then spz& = inex& Else spz& = res&(inex&)
    If spz& > 0 Then
      Get #ndfi, (spz& - 1) * ih& + 1, rcim&
      '--- rekordok beallitasa es kitoltes
      Seek #dbfi, rcim& + 9
      r$ = Space(rh&): Get #dbfi, , r$
      For i2& = 1 To kmesor%(0)
        adamez$ = Mid$(r$, kmkp%(i2&), kmho%(i2&))
        MSFlexGrid1.TextMatrix(i&, i2& - 1) = adamez$
      Next
    End If
  Next
  If sdb% < lapmeret% Then
    For i& = sdb% + 1 To lapmeret%
      For i2& = 1 To kmesor%(0)
        adamez$ = Space$(kmho%(i2&))
        MSFlexGrid1.TextMatrix(i&, i2& - 1) = adamez$
      Next
    Next
  End If
  MSFlexGrid1.Row = aktsor& - elsosor& + 1
  MSFlexGrid1.BackColorSel = &HC0FFC0
  MSFlexGrid1.RowSel = aktsor& - elsosor& + 1
  MSFlexGrid1.ColSel = kmesor%(0) - 1
  '--- rekord beolvasasa es objektum beallitasa
  inex& = aktsor&
  If statusz% = 0 Then spz& = inex& Else spz& = res&(inex&)
  If spz& > 0 Then
    Get #ndfi, (spz& - 1) * ih& + 1, rcim&
    '--- rekordok beallitasa es kitoltes
    Seek #dbfi, rcim& + 4
    torlojel$ = " ": Get #dbfi, , torlojel$
    Get #dbfi, rcim& + 5, oobnext&
    Seek #dbfi, rcim& + 9
    r$ = Space(rh&): Get #dbfi, , r$
    oobcim& = rcim&
    oobind& = spz&
    rekord$ = r$
    If infolapbe = 1 Then Call infomutat(keresobj$, r$, "Hash")
  End If
  Close dbfi
  Close ndfi
  Text2.Text = langform(4) + Str$(darab&)
  If torljel$ = "*" Then
    rekord$ = ""
    Text3.ForeColor = QBColor(4)
    Text3.FontBold = True
    Text3.Text = langform(5)
  Else
    Text3.ForeColor = QBColor(0)
    Text3.Text = langform(6) + Str$(aktsor&)
  End If
End Sub

Private Sub Form_Load()
  MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  MSFlexGrid1.Font.Size = 8
  
  Call langinit("hash", 2)
  Call szkriptel("hash")
End Sub

Private Sub MSFlexGrid1_Click()
  akt& = MSFlexGrid1.Row + elsosor& - 1
  If akt& > darab& Then
    MSFlexGrid1.Row = aktsor& - elsosor& + 1
  Else
    aktsor& = akt&
  End If
  MSFlexGrid1.Col = 0
  Call megmutat(dbxn$, indn$)
  Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_DblClick()
  Text1.SetFocus
End Sub

Private Sub Text1_Change()
  If belep% = 1 Then Exit Sub
  hk$ = xkonver(Text1.Text)
  hkh% = Len(hk$)
  If hkh% > 0 Then
    Select Case hkh%
      Case 1
        resdb& = 0
        For i97& = 1 To Len(hst1$)
          If hk$ = Mid$(hst1$, i97&, 1) Then
            If resdb& < 30000 Then
              resdb& = resdb& + 1: res&(resdb&) = i97&
            Else
              Exit For
            End If
          End If
        Next
        'If resdb& > 0 Then
          '--- ABC rendezés
          Call abcsort
          statusz% = 1
          elsosor& = 1: aktsor& = 1: darab& = resdb&
          Call megmutat(dbxn$, indn$)
        
        'End If
      Case 2
        resdb& = 0
        For i97& = 1 To Len(hst1$)
          If hk$ = Mid$(hst1$, i97&, 1) + Mid$(hst2$, i97&, 1) Then
            If resdb& < 30000 Then
              resdb& = resdb& + 1: res&(resdb&) = i97&
            Else
              Exit For
            End If
          End If
        Next
        'If resdb& > 0 Then
          Call abcsort
          statusz% = 1
          elsosor& = 1: aktsor& = 1: darab& = resdb&
          Call megmutat(dbxn$, indn$)
        'End If
      Case 3
        resdb& = 0
        For i97& = 1 To Len(hst1$)
          If hk$ = Mid$(hst1$, i97&, 1) + Mid$(hst2$, i97&, 1) + Mid$(hst3$, i97&, 1) Then
            If resdb& < 30000 Then
              resdb& = resdb& + 1: res&(resdb&) = i97&
            Else
              Exit For
            End If
          End If
        Next
        'If resdb& > 0 Then
          Call abcsort
          statusz% = 1
          elsosor& = 1: aktsor& = 1: darab& = resdb&
          Call megmutat(dbxn$, indn$)
        'End If
      Case Else
        '--- teljeskörû keresés
        If resdb& > 1 Then
          dbfi = FreeFile
          Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
          ndfi = FreeFile
          Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
          resmdb% = 0
          For i97& = 1 To resdb&
            Get #ndfi, (res&(i97&) - 1) * ih& + 1, rcim&
            Seek #dbfi, rcim& + 9
            r$ = Space$(rh&): Get #dbfi, , r$
            If xkonver(Mid$(r$, hck%, hkh%)) = hk$ Then
              If resmdb% < 30000 Then
                resmdb% = resmdb% + 1
                resm&(resmdb%) = res&(i97&)
              End If
            End If
          Next
          Close dbfi: Close ndfi
          'If resmdb% > 0 Then
            For i97& = 1 To resmdb%: res&(i97&) = resm&(i97&): Next
            resdb& = resmdb%
            Call abcsort
            statusz% = 1
            elsosor& = 1: aktsor& = 1: darab& = resdb&
            Call megmutat(dbxn$, indn$)
          'End If
        End If
    End Select
  Else
    statusz% = 0: elsosor& = 1: aktsor& = 1: darab& = hc&
    Call megmutat(dbxn$, indn$)
  End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- billentyû lenyomása text1-ben
  If darab& = 0 Then
    Select Case KeyCode
      Case vbKeyDelete
       KeyCode = 0: Text1.Text = ""
      Case vbKeyEscape, vbKeyReturn
         KeyCode = 0
         Call Command2_Click
      Case Else
    End Select
    Exit Sub
  End If
  Select Case KeyCode
    Case vbKeyDelete
      KeyCode = 0: Text1.Text = ""
    Case vbKeyEscape
      KeyCode = 0
      Call Command2_Click
    Case vbKeyReturn
      '--- választás (Enter)
      talalat% = 1
      OBJTAB(ob%).obcim = oobcim&
      OBJTAB(ob%).obind = oobind&
      OBJTAB(ob%).obnex = oobnext&
      toprox& = elsosor&: aktrox& = aktsor&: aktdarab& = darab&: tabstatusx% = statusz%
      aktext$ = Text1.Text
      Call langclos
      Hash.Hide
    Case vbKeyUp
      '--- elõzõ sor
      aktsor& = aktsor& - 1: If aktsor& < 1 Then aktsor& = 1
      If aktsor& < elsosor& Then elsosor& = elsosor& - 1
      KeyCode = 0
    Case vbKeyDown
      '--- következõ sor
      aktsor& = aktsor& + 1: If aktsor& > darab& Then aktsor& = darab&
      If aktsor& > elsosor& + lapmeret% - 1 Then elsosor& = elsosor& + 1
      KeyCode = 0
    Case vbKeyPageUp
      '--- lapozás vissza
      aktsor& = aktsor& - lapmeret%
      If aktsor& < 1 Then
        aktsor& = 1: elsosor& = 1
      Else
        elsosor& = elsosor& - lapmeret%
        If elsosor& < 1 Then elsosor& = 1
      End If
      KeyCode = 0
    Case vbKeyPageDown
      '--- lapozás elõre
      aktsor& = aktsor& + lapmeret%
      If aktsor& > darab& Then
        aktsor& = darab&
        elsosor& = darab& - lapmeret% + 1
        If elsosor& < 1 Then elsosor& = 1
      Else
        elsosor& = elsosor& + lapmeret%
        If elsosor& > darab& Then elsosor = darab&
      End If
      KeyCode = 0
    Case vbKeyHome
      '--- elsõ sor elsõ mezõ
      aktsor& = 1: elsosor& = 1
      KeyCode = 0
    Case vbKeyEnd
      '--- utolsó sor utolsó mezõ
      aktsor& = darab&
      If darab& < lapmeret% Then
        elsosor& = 1
      Else
        elsosor& = darab& - lapmeret% + 1
      End If
      KeyCode = 0
    Case Else
      '--- hash code kezelése
  End Select
  Call megmutat(dbxn$, indn$)
End Sub
Private Sub abcsort()
 Exit Sub
 dbfi = FreeFile
 Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
 ndfi = FreeFile
 Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
 If hsmezohossz% > 30 Then hsmezohossz% = 30
 For i97& = 1 To resdb&
   Get #ndfi, (res&(i97&) - 1) * ih& + 1, rcim&
   Seek #dbfi, rcim& + 9
   r$ = Space$(rh&): Get #dbfi, , r$
   resa$(i97&) = xkonver(Mid$(r$, hck%, hsmezohossz%))
 Next
 Close dbfi: Close ndfi
 Call qsort&(resa$(), res&(), resdb&, "N")
End Sub


VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form rendezo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00B4F0FF&
   Caption         =   "Rendezés"
   ClientHeight    =   1272
   ClientLeft      =   2928
   ClientTop       =   6756
   ClientWidth     =   4752
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   1272
   ScaleWidth      =   4752
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Megszakítás"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1212
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1212
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4572
      _ExtentX        =   8065
      _ExtentY        =   445
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar ProgressBar4 
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4572
      _ExtentX        =   8065
      _ExtentY        =   445
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "rendezo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  '--- a rendezo program
  Dim n$(6), e$(6), k(10), h(10), rmo(10), sorp%(100), sorv%(100)
  Dim msorp&(4000), msorv&(4000), recc$(4000), ar$(4000), arn@(4000)
  Dim az(1 To 4000) As String * 80
  Dim azn@(4000)
  Dim axu&(1 To 4000)
  Dim q1$, q2$, q3$, q4$, q5$, q6$, karmod%, utv$, csokkeno%
  rendezohiba% = 0
  'On Error GoTo hibakez
  a$ = rendpar$
  If Mid$(rendpar$, 5, 1) = "D" Then
    csokkeno% = 1
    a$ = Left(rendpar$, 4) + Mid$(rendpar$, 6)
  Else
    csokkeno% = 0
  End If
  rendezohiba% = 0
  n$(1) = "FIE"
  n$(2) = "INP"
  n$(3) = "OUT"
  n$(4) = "RLE"
  n$(5) = "MOD"
  n$(6) = "BMX"
  For i = 1 To 6
    poz = InStr(a$, "/")
    If poz <> 0 Then
      b$ = Left$(a$, poz - 1)
      a$ = Mid$(a$, poz + 1)
    Else
      b$ = a$
    End If
    poz1 = InStr(b$, "=")
    nev$ = Left$(b$, poz1 - 1)
    ert$ = Mid$(b$, poz1 + 1)
    For j& = 1 To 6
      If Left$(nev$, 3) = n$(j&) Then e$(j&) = ert$
    Next
    If i = 5 And poz = 0 Then Exit For
  Next
  kdb% = 0
  a$ = e$(1)
  Do
    If Len(a$) <> 0 Then
      poz = InStr(a$, ",")
      k1$ = Left$(a$, poz - 1)
      a$ = Mid$(a$, poz + 1)
      poz = InStr(a$, ",")
      If poz <> 0 Then
        h1$ = Left$(a$, poz - 1)
        a$ = Mid$(a$, poz + 1)
      Else
        h1$ = a$
      End If
      kdb% = kdb% + 1
      k(kdb%) = Val(k1$)
      h(kdb%) = Val(h1$)
    End If
  Loop While poz <> 0
  inp$ = e$(2)
  ou$ = e$(3)
  modr$ = e$(5)
  '--- kalkulacio
  rh& = Val(e$(4))
  pufdb& = Int(15000& / rh&)
  pufho& = pufdb& * rh&
  sorhossz% = 50
  fi1 = FreeFile
  Open inp$ For Binary As #fi1
  rc& = Int(LOF(fi1) / rh&)
  If rc& = 0 Then
    Close fi1
    'rendezo.Visible = False
    Call langclos
    rendezo.Hide
    Exit Sub
  End If
  sormer& = Int(rc& / 13)
  If sormer& < 500 Then sormer& = 500
  If sormer& > 1000 Then sormer& = 1000
  works& = Int(rc& / sormer&)
  Last& = rc& Mod sormer&
  If Last& > 0 Then
    works& = works& + 1
  End If
  '--- work terulet
  fi2 = FreeFile
  Open "C:\auwin\" + terminal$ + task$ + "x.swk" For Output As #fi2
  pufp& = 0
  ProgressBar1.Min = 1
  ProgressBar1.Max = 100
  For q& = 1 To works&
    DoEvents
    ProgressBar1.Value = pscale(q&, works&)
    rcs& = sormer&
    If q& = works& And Last& <> 0 Then rcs& = Last&
    '--- input fazis
    wi& = (q& - 1) * sormer&
    For j& = 1 To rcs&
      rcim& = wi& + j&
      Seek #fi1, (rcim& - 1) * rh& + 1
      rec$ = Space(rh&)
      Get #fi1, , rec$
      af$ = ""
      For k11& = 1 To kdb%
        af$ = af$ + Mid$(rec$, k(k11&), h(k11&))
      Next
      If modr$ = "M" Then
        az(j&) = mkonver(af$)
      Else
        If modr$ = "N" Then
          azn@(j&) = xval(af$)
        Else
          az(j&) = af$
        End If
      End If
      axu&(j&) = wi& + j&
    Next
    '--- sort fazis
    sorok% = Int(rcs& / sorhossz%)
    uti& = rcs& Mod sorhossz%
    For gu& = 1 To sorok% + 1
      inx% = (gu& - 1) * sorhossz% + 1
      iny% = inx% + sorhossz% - 1
      sorp%(gu&) = inx%
      If gu& = sorok% + 1 Then
        If uti& = 0 Then Exit For
        iny% = inx% + uti& - 1
      End If
      sorv%(gu&) = iny%
      '--- 50-as sor rendezese
      If inx% <= iny% - 1 Then
        For l% = inx% To iny% - 1
          mimu% = l%
          For k12% = l% + 1 To iny%
            If modr$ = "N" Then
              If csokkeno% = 0 Then
                If azn@(k12%) < azn@(mimu%) Then mimu% = k12%
              Else
                If azn@(k12%) > azn@(mimu%) Then mimu% = k12%
              End If
            Else
              If csokkeno% = 0 Then
                If az(k12%) < az(mimu%) Then mimu% = k12%
              Else
                If az(k12%) > az(mimu%) Then mimu% = k12%
              End If
            End If
          Next
          If mimu% <> l% Then
            xxx$ = az(mimu%): az(mimu%) = az(l%): az(l%) = xxx$
            yyy@ = azn@(mimu%): azn@(mimu%) = azn@(l%): azn@(l%) = yyy@
            zzz& = axu&(mimu%): axu&(mimu%) = axu&(l%): axu&(l%) = zzz&
          End If
        Next
      End If
    Next
    '--- merge-output fazis
    If uti& = 0 Then
      sssx% = sorok%
    Else
      sssx% = sorok% + 1
    End If
    For j& = 1 To rcs&
      '--- min meghat
      volt% = 0
      For h13% = 1 To sssx%
        wi9% = sorp%(h13%)
        If wi9% <= sorv%(h13%) Then
          If volt% = 0 Then mixi% = h13%: volt% = 1: wj% = sorp%(mixi%)
          If modr$ = "N" Then
            If csokkeno% = 0 Then
              If azn@(wi9%) < azn@(wj%) Then
                mixi% = h13%
                wj% = sorp%(mixi%)
              End If
            Else
              If azn@(wi9%) > azn@(wj%) Then
                mixi% = h13%
                wj% = sorp%(mixi%)
              End If
            End If
          Else
            If csokkeno% = 0 Then
              If az(wi9%) < az(wj%) Then
                mixi% = h13%
                wj% = sorp%(mixi%)
              End If
            Else
              If az(wi9%) > az(wj%) Then
                mixi% = h13%
                wj% = sorp%(mixi%)
              End If
            End If
          End If
        End If
      Next
      rcim& = axu&(sorp%(mixi%))
      Seek #fi1, (rcim& - 1) * rh& + 1
      rec$ = Space(rh&)
      Get #fi1, , rec$
      filsz = fi2
      GoSub outr
      sorp%(mixi%) = sorp%(mixi%) + 1
    Next
  Next
  If pufp& > 0 Then Print #fi2, Left$(puf$, pufp& * rh&);
  Close fi1: Close fi2
  '--- final merge
  For i1& = 1 To works&
    msorp&(i1&) = 1
    msorv&(i1&) = sormer&
    If i1& = works& And Last& <> 0 Then msorv&(i1&) = Last&
  Next
  fi2 = FreeFile
  Open "C:\auwin\" + terminal$ + task$ + "x.SWK" For Binary As fi2
  For q& = 1 To works&
    w& = (q& - 1) * sormer& + 1&
    Seek #fi2, (w& - 1) * rh& + 1
    rax$ = Space(rh&)
    Get #fi2, , rax$
    recc$(q&) = rax$
    af$ = ""
    For k14% = 1 To kdb%
      af$ = af$ + Mid$(recc$(q&), k(k14%), h(k14%))
    Next
    If modr$ = "M" Then
      ar$(q&) = mkonver(af$)
    Else
      If modr$ = "N" Then
        arn@(q&) = xval(af$)
      Else
        ar$(q&) = af$
      End If
    End If
  Next
  ox& = 0
  fi20 = FreeFile
  Open ou$ For Output As #fi20
  pufp& = 0
  rsz& = rc& / 100&: If rsz& < 1 Then rsz& = 1
  ProgressBar4.Min = 1
  ProgressBar4.Max = 100
  Do
    '--- min meghat
    DoEvents
    volt% = 0
    For h18% = 1 To works&
      If msorp&(h18%) <= msorv&(h18%) Then
        If volt% = 0 Then mixi% = h18%: volt% = 1
        If modr$ = "N" Then
          If csokkeno% = 0 Then
            If arn@(h18%) < arn@(mixi%) Then
              mixi% = h18%
            End If
          Else
            If arn@(h18%) > arn@(mixi%) Then
              mixi% = h18%
            End If
          End If
        Else
          If csokkeno% = 0 Then
            If ar$(h18%) < ar$(mixi%) Then
              mixi% = h18%
            End If
          Else
            If ar$(h18%) > ar$(mixi%) Then
              mixi% = h18%
            End If
          End If
        End If
      End If
    Next
    ox& = ox& + 1
    ProgressBar4.Value = pscale(ox&, rc&)
    rec$ = recc$(mixi%)
    filsz = fi20
    GoSub outr
    msorp&(mixi%) = msorp&(mixi%) + 1
    If msorp&(mixi%) <= msorv&(mixi%) Then
      w& = (mixi% - 1) * sormer& + msorp&(mixi%)
      Seek #fi2, (w& - 1) * rh& + 1
      rax$ = Space(rh&)
      Get #fi2, , rax$
      recc$(mixi%) = rax$
      af$ = ""
      For k16% = 1 To kdb%
        af$ = af$ + Mid$(recc$(mixi%), k(k16%), h(k16%))
      Next
      If modr$ = "M" Then
        ar$(mixi%) = mkonver(af$)
      Else
        If modr$ = "N" Then
          arn@(mixi%) = xval(af$)
        Else
          ar$(mixi%) = af$
        End If
      End If
    End If
    If ox& Mod rsz& = 0 Then
      w3& = Int(ox& / rsz&)
    End If
  Loop While ox& < rc&
  If pufp& > 0 Then Print #fi20, Left$(puf$, pufp& * rh&);
  Close fi2: Close fi20
  Kill "C:\auwin\" + terminal$ + task$ + "x.swk"
  rendezo.Visible = False
  Call langclos
  rendezo.Hide
  'Call Command2_Click
  Exit Sub
outr:
 If pufp& = pufdb& Then pufp& = 0: Print #filsz, puf$;
 If pufp& = 0 Then puf$ = Space$(pufho&)
 pufp& = pufp& + 1&
 Mid$(puf$, (pufp& - 1&) * rh& + 1&) = rec$
Return
Exit Sub
hibakez:
  rendezohiba% = 1
  Call mess(langmodul$(164), 2, 0, langmodul$(165), valasz%)
  Call Command2_Click
End Sub


Private Sub Command2_Click()
  rendezohiba% = 1
  rendezo.Visible = False
  Call langclos
  rendezo.Hide
End Sub

Private Sub Form_Load()
  Call langinit("rendezo", 2)
  Call szkriptel("rendezo")
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Call Command1_Click
End Sub

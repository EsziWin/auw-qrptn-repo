VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Kereso 
   BackColor       =   &H8000000A&
   Caption         =   "Keresés"
   ClientHeight    =   5388
   ClientLeft      =   2508
   ClientTop       =   2472
   ClientWidth     =   10272
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5388
   ScaleWidth      =   10272
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command31 
      BackColor       =   &H00FFFFFF&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4200
      Width           =   444
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H00E0FDFE&
      Caption         =   "Rejt"
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4920
      Width           =   444
   End
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
      Left            =   5880
      TabIndex        =   34
      Top             =   170
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
         TabIndex        =   35
         Top             =   240
         Width           =   3612
      End
   End
   Begin VB.CommandButton Command26 
      Height          =   400
      Left            =   9840
      Picture         =   "Kereso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3480
      Width           =   444
   End
   Begin VB.CommandButton Command25 
      Height          =   400
      Left            =   9840
      Picture         =   "Kereso.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   600
      Width           =   444
   End
   Begin VB.CommandButton Command24 
      Height          =   400
      Left            =   9840
      Picture         =   "Kereso.frx":0EC4
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   120
      Width           =   444
   End
   Begin VB.CommandButton Command23 
      Height          =   400
      Left            =   9840
      Picture         =   "Kereso.frx":1B06
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2040
      Width           =   444
   End
   Begin VB.CommandButton Command22 
      Height          =   400
      Left            =   9840
      Picture         =   "Kereso.frx":2748
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1560
      Width           =   444
   End
   Begin VB.CommandButton Command20 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.2
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9840
      Picture         =   "Kereso.frx":338A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Rendezés csökkenõ sorrenbev"
      Top             =   3000
      Width           =   444
   End
   Begin VB.CommandButton Command19 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.2
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9840
      MaskColor       =   &H00404040&
      Picture         =   "Kereso.frx":360C
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Rendezés növekvõ sorrendben"
      Top             =   2520
      Width           =   444
   End
   Begin VB.CommandButton Command21 
      Height          =   400
      Left            =   9840
      Picture         =   "Kereso.frx":388E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1080
      Width           =   444
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   3720
      TabIndex        =   17
      Top             =   3720
      Width           =   6012
      _ExtentX        =   10605
      _ExtentY        =   445
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   3720
      TabIndex        =   18
      Top             =   3680
      Visible         =   0   'False
      Width           =   2292
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
      Left            =   1800
      TabIndex        =   16
      Text            =   "Pozíció:"
      ToolTipText     =   "Az aktuális rekord sorszáma"
      Top             =   3680
      Width           =   1812
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
      TabIndex        =   15
      Text            =   "Sorok:"
      ToolTipText     =   "A táblázat sorainak a száma"
      Top             =   3680
      Width           =   1572
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3492
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   9612
      _ExtentX        =   16955
      _ExtentY        =   6160
      _Version        =   327680
      Rows            =   18
      FixedCols       =   0
      BackColorFixed  =   5270
      ForeColorFixed  =   16777215
      BackColorSel    =   14737632
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      Enabled         =   -1  'True
      FillStyle       =   1
      ScrollBars      =   1
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Keresés"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   9650
      Begin VB.CommandButton Command10 
         BackColor       =   &H0025B5EB&
         Caption         =   "Mind"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Visszatér a teljes táblázathoz (F1)"
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H0091E9FB&
         Caption         =   "Index"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Direkt beállás elsõdleges azonosító szerint (F6)"
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0091E9FB&
         Caption         =   "Pozíció"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "A következõ szöveg elõfordulásra lép (F4)"
         Top             =   240
         Width           =   1092
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   7.8
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "A keresett érték"
         Top             =   240
         Width           =   2292
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0055D7F7&
         Caption         =   "Intervallum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Tól-ig keresés a kijelölt oszlopban (F5)"
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0091E9FB&
         Caption         =   "Mezõ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Szöveg keresés a kijelölt oszlop értéke szerint (F3)"
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0091E9FB&
         Caption         =   "Szöveg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Szöveg keresés a teljes rekordban (F2)"
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Mûvelet"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   9650
      Begin VB.CommandButton Command29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command29"
         Height          =   372
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.CommandButton Command28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command28"
         Height          =   372
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton Command27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command27"
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command17"
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command16"
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command15"
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
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command14"
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command12"
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
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "A megjelölt rekord kiválasztása (Enter)"
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kilépés"
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
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Kilépés választás nélkül (Esc)"
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command5"
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command4"
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command3"
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
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command2"
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command1"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H001428C8&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   6372
      Left            =   9780
      Top             =   0
      Width           =   576
   End
End
Attribute VB_Name = "Kereso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- KERESO form (Keresõtábla) kódja
Dim rh&, ih&, dbxn$, indn$, elsosor&, aktsor&, darab&, rc&, sorbak&
Dim oobcim&, oobind&, oobnext&, ob%, fontsi%, tizenharom%, tizenketto%
Dim statusz% '--- =0 teljes halmaz, =1 szûkített halmaz
Dim teldarab&, indexdarabszam&, mindnyomott%, colwi%(200), colwimod%
Dim resa$(100000), resm&(100000), res&(100000), resdb&, resmdb& '--- rekordok fizikai sorszáma
Dim keresokodvektor$, bejegyzeshossz&, keresokodelemszam&, kereskoddb%, keresokodtabla%(200, 3) 's,1 kezdopoz s,2 hossz s,3 elõzõ hossz
Dim ebbenvansor&, ebbenvoltsor&, r$, yktop%, ykleft%, ykheight%, ykwidth%, ycleft%, yctop%, ycwidth%, ycheight%, keresrejt%

Private Sub Command1_Click()
  '--- Gombsor 1.gomb
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  gombsorszam% = 1
  Kereso.Command1.ToolTipText = ""
  Kereso.Command2.ToolTipText = ""
  Call langclos
  Kereso.Hide
End Sub

Private Sub Command10_Click()
  '--- alaphalmaz visszaállítása (F1)
  mindnyomott% = 1
  statusz% = 0
  elsosor& = 1: darab& = teldarab&: aktsor& = 1
  kezd& = 1: akt& = 1: If darab& > tizenharom Then zar& = tizenharom Else zar& = darab&
  Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
  DoEvents
  Text1.SetFocus
End Sub

Private Sub Command11_Click()
  '--- pozícióra beállás (F4)
  Dim szopar$(20), szopardb%
  If darab& < 1 Then Exit Sub
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  ProgressBar1.Visible = True
  ProgressBar1.Min = 1 'pscale(aktsor&, darab&)
  ProgressBar1.Max = 100
  If aktsor& < darab& Then
    szopardb% = 0
    kte$ = xkonver(Trim$(Text1.Text))
    If InStr(kte$, " ") > 0 Then
      tobbszavas% = 1
      Call linpar(kte$, szopar$, " ", szopardb%)
      kte$ = szopar$(1)
    End If
    For rcc& = aktsor& + 1 To darab&
      ProgressBar1.Value = pscale(rcc&, darab&)
      If statusz% = 0 Then spz& = rcc& Else spz& = res&(rcc&)
      Get #ndfi, (spz& - 1) * ih& + 1, rcim&
      '--- rekordok beallitasa es kitoltes
      Get #dbfi, rcim& + 5, oobnext&
      Seek #dbfi, rcim& + 9
      rww$ = Space(rh&): Get #dbfi, , rww$: r$ = xkonver(rww$)
      If InStr(r$, kte$) > 0 Then
        If szopardb% > 0 Then
          jujo% = 1
          For juj% = 1 To szopardb%
            If InStr(r$, szopar(juj%)) = 0 Then jujo% = 0: Exit For
          Next
        Else
          jujo% = 1
        End If
        If jujo% = 1 Then
          aktsor& = rcc&
          If aktsor& + tizenketto > darab& Then
            elsosor& = darab& - tizenharom + 1
            If elsosor& < 1 Then elsosor& = 1
            kezd& = elsosor&: akt& = aktsor& - elsosor& + 1
            zar& = elsosor& + tizenketto
            If zar& > darab& Then zar& = darab&
          Else
            elsosor& = aktsor&
            kezd& = elsosor&: akt& = 1: zar& = kezd& + tizenketto
          End If
          ProgressBar1.Visible = False
          Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
          Exit For
        End If
      End If
    Next
    If ProgressBar1.Visible = True Then
      ProgressBar1.Visible = False
      Call mess(langform(17), 3, 0, langform(18), valasz%)
      'MsgBox langform(17), 48, langform(18)
    End If
  End If
  Close ndfi: Close dbfi
  DoEvents
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Command12_Click()
  '--- gombsor 6.gomb
  '--- keresõtábla esetén választás Enter
  althobjektum$ = keresobj$
  talalat% = 1
  gombsorszam% = 6
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  Kereso.Command1.ToolTipText = ""
  Kereso.Command2.ToolTipText = ""

  Call langclos
  Kereso.Hide
End Sub

Private Sub Command13_Click()
  talalat% = 1
  gombsorszam% = 7
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  Kereso.Command1.ToolTipText = ""
  Kereso.Command2.ToolTipText = ""

  Call langclos
  Kereso.Hide
End Sub

Private Sub Command14_Click()
  talalat% = 1
  gombsorszam% = 8
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  Kereso.Command1.ToolTipText = ""
  Kereso.Command2.ToolTipText = ""

  Call langclos
  Kereso.Hide
End Sub

Private Sub Command15_Click()
  talalat% = 1
  gombsorszam% = 9
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  Kereso.Command1.ToolTipText = ""
  Kereso.Command2.ToolTipText = ""

  Call langclos
  Kereso.Hide
End Sub

Private Sub Command16_Click()
  talalat% = 1
  gombsorszam% = 10
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  Kereso.Command1.ToolTipText = ""
  Kereso.Command2.ToolTipText = ""

  Call langclos
  Kereso.Hide
End Sub

Private Sub Command17_Click()
  talalat% = 1
  gombsorszam% = 11
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  Kereso.Command1.ToolTipText = ""
  Kereso.Command2.ToolTipText = ""

  Call langclos
  Kereso.Hide
End Sub

Private Sub Command27_Click()
  talalat% = 1
  gombsorszam% = 12
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  Kereso.Command1.ToolTipText = ""
  Kereso.Command2.ToolTipText = ""
  
  Call langclos
  Kereso.Hide
End Sub


Private Sub Command18_Click()
  Call indexre
End Sub

Private Sub Command19_Click()
  '--- növekvõ rendezés
  If resdb& > 1 And mindnyomott% = 0 Then
    dbfi = FreeFile
    Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
    ndfi = FreeFile
    Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
    ProgressBar1.Visible = True
    ProgressBar1.Min = 1
    ProgressBar1.Max = 100
    i2& = MSFlexGrid1.Col + 1
    If kmho%(i2&) > 10 Then ooko% = 10 Else ooko% = kmho%(i2&)
    If InStr(kmtb(i2&), "N") > 0 Then num% = 1 Else num% = 0
    wpz& = kmkp%(i2&) + 8
    For rcc& = 1 To darab&
      ProgressBar1.Value = pscale(rcc&, darab&)
      spz& = res&(rcc&)
      Get #ndfi, (spz& - 1) * ih& + 1, rcim&
      '--- rekordok beallitasa es kitoltes
      adaba$ = Space(ooko%)
      Get #dbfi, rcim& + wpz&, adaba$
      If num% = 0 Then
        resa(rcc&) = mkonver(adaba$)
      Else
        resa(rcc&) = adaba$
      End If
    Next
    Close dbfi: Close ndfi
    Call qsort(resa(), res(), darab&, "N")
    ProgressBar1.Value = 1
    statusz% = 1
    darab& = resdb&
    elsosor& = 1: kezd& = 1: aktsor& = 1: akt& = 1
    If darab& > tizenharom Then zar& = tizenharom Else zar& = darab&
    Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
  End If
  DoEvents
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Command2_Click()
  '--- gombsor 2.gomb
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  gombsorszam% = 2
  If keresomod% = 0 Then toprow& = elsosor&: aktrow& = aktsor&
  Call langclos
  Kereso.Hide
End Sub

Private Sub Command20_Click()
  '--- csökkenõ rendezés
  If resdb& > 1 And mindnyomott% = 0 Then
    dbfi = FreeFile
    Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
    ndfi = FreeFile
    Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
    ProgressBar1.Visible = True
    ProgressBar1.Min = 1
    ProgressBar1.Max = 100
    i2& = MSFlexGrid1.Col + 1
    If kmho%(i2&) > 10 Then ooko% = 10 Else ooko% = kmho%(i2&)
    If InStr(kmtb(i2&), "N") > 0 Then num% = 1 Else num% = 0
    wpz& = kmkp%(i2&) + 8
    For rcc& = 1 To darab&
      ProgressBar1.Value = pscale(rcc&, darab&)
      spz& = res&(rcc&)
      Get #ndfi, (spz& - 1) * ih& + 1, rcim&
      '--- rekordok beallitasa es kitoltes
      adaba$ = Space(ooko%)
      Get #dbfi, rcim& + wpz&, adaba$
      If num% = 0 Then
        resa(rcc&) = mkonver(adaba$)
      Else
        resa(rcc&) = adaba$
      End If
    Next
    Close dbfi: Close ndfi
    Call qsort(resa(), res(), darab&, "C")
    ProgressBar1.Value = 1
    statusz% = 1
    darab& = resdb&
    elsosor& = 1: kezd& = 1: aktsor& = 1: akt& = 1
    If darab& > tizenharom Then zar& = tizenharom Else zar& = darab&
    Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
  End If
  DoEvents
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Command21_Click()
  If colwimod% = 0 Then
    For izu% = 1 To MSFlexGrid1.Cols
      hxh% = Kereso.TextWidth(Trim(MSFlexGrid1.TextMatrix(0, izu% - 1)))
      hxk% = Kereso.TextWidth(String(kmho%(izu%), " "))
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
    Case 6: tizenharom = 15
    Case 7: tizenharom = 13
    Case 8: tizenharom = 13
    Case 9: tizenharom = 11
    Case 10: tizenharom = 10
    Case Else
  End Select
  tizenketto = tizenharom - 1
  MSFlexGrid1.Rows = tizenharom + 1
  Call MSFlexGrid1_Click
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Command23_Click()
  If fontsi% = 0 Then fontsi% = 8
  If fontsi% > 6 Then fontsi% = fontsi% - 1: MSFlexGrid1.Font.Size = fontsi%
  Select Case fontsi%
    Case 6: tizenharom = 15
    Case 7: tizenharom = 13
    Case 8: tizenharom = 13
    Case 9: tizenharom = 11
    Case 10: tizenharom = 10
    Case Else
  End Select
  tizenketto = tizenharom - 1
  MSFlexGrid1.Rows = tizenharom + 1
  Call MSFlexGrid1_Click
  If Text1.Visible = True Then Text1.SetFocus
End Sub


Private Sub Command24_Click()
  Kereso.Left = 1
  Kereso.Width = 12250
  Frame3.Left = 7690
  Kereso.Shape1.Left = 11590
End Sub

Private Sub Command25_Click()
  Kereso.Left = 1000
  Kereso.Width = 10440
  Frame3.Left = 5880
  Kereso.Shape1.Left = 9780
End Sub

Private Sub Command26_Click()
  'If MSFlexGrid1.Font.Name = "Microsoft Sans Serif" Then
  '  MSFlexGrid1.Font.Name = "Arial Narrow"
  'Else
  '  MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  'End If
  If infolapbe% = 0 Then infolapbe% = 1: Frame3.Visible = True Else infolapbe% = 0: Frame3.Visible = False
  If infolapbe% = 1 Then Call infomutat(keresobj$, r$, "Kereso")
  DoEvents
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Command28_Click()
  '--- gombsor 13.gomb
  talalat% = 1
  gombsorszam% = 13
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  Call langclos
  Kereso.Hide
End Sub

Private Sub Command29_Click()
  '--- gombsor 14.gomb
  talalat% = 1
  gombsorszam% = 14
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  Call langclos
  Kereso.Hide
End Sub


Private Sub Command3_Click()
  '--- gombsor 3.gomb
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  gombsorszam% = 3
  Call langclos
  Kereso.Hide
End Sub

Private Sub Command30_Click()
  If keresrejt% = 0 Then
    keresrejt% = 1
    yktop = Kereso.Top
    ykleft = Kereso.Left
    ykwidth = Kereso.Width
    ykheight = Kereso.Height
    ycleft = Command30.Left: yctop = Command30.Top
    ycheight = Command30.Height: ycwidth = Command30.Width
    Kereso.Width = 200: Kereso.Height = 800
    Kereso.Top = 7900: Kereso.Left = 100
    Command30.Left = 1: Command30.Top = 1
    Command30.Width = Kereso.Width - 100: Command30.Height = Kereso.Height - 460
    Command30.Caption = "Keresõ tábla kinyitása"
  Else
    keresrejt% = 0
    Kereso.Top = yktop
    Kereso.Left = ykleft
    Kereso.Width = ykwidth
    Kereso.Height = ykheight
    Command30.Left = ycleft: Command30.Top = yctop
    Command30.Width = ycwidth: Command30.Height = ycheight
    Command30.Caption = "Rejt"
    If Text1.Visible = True Then Text1.SetFocus
  End If
End Sub

Private Sub Command31_Click()
  '--- leválogatott halmaz nyomtatása
  Dim qanev$(200), qattr$(200), qakp%(200), qamh%(200), qakell%(200), qoho%(200)
  qw1& = obsorszama(keresobj)
  Halprint.List1.Clear
  qadb& = 0
  For i1& = 1 To adadb%
    If ADATAB(i1&).obsorsz = qw1& Then
      ne$ = RTrim$(ADATAB(i1&).adatnev)
      ar$ = RTrim$(ADATAB(i1&).attr)
      mh% = ADATAB(i1&).adatho
      kp% = ADATAB(i1&).adatkp
      Halprint.List1.AddItem ne$
      qadb = qadb + 1
      qanev(qadb) = ne$
      qattr(qadb) = ar$
      qakp%(qadb) = kp%
      qamh(qadb) = mh%
      qoho%(qadb) = mh%
      If Len(ne$) > mh% Then qoho%(qadb) = Len(ne$)
    End If
  Next
  Halprint.Label3.Caption = "Sorok száma: " + Trim(Str(darab))
  Halprint.Show vbModal
  If halprin = 1 Then
    '--- nyomtatás
    If darab& < 1 Then Exit Sub
    For i44& = 1 To qadb
      If Halprint.List1.Selected(i44& - 1) = True Then
        qakell(i44&) = 1
      Else
        qakell(i44&) = 0
      End If
    Next
    '--- cím, fejlec nyomtatása
    prfi = FreeFile
    Open listautvonal + terminal$ + task$ + "rhl.lst" For Output As #prfi
    qs$ = "CM" + Trim(OBJTAB(qw1&).obnev) + " (részhalmaz). Készült:" + datki(maidatum): Print #prfi, qs$
    sh% = 0
    qs$ = "FL"
    For i44 = 1 To qadb
      If qakell(i44) = 1 Then
        If InStr(qattr(i44), "J") > 0 Then
          qs$ = qs$ + Right(Space(qoho(i44)) + qanev(i44), qoho(i44)) + " "
          sh% = sh% + qoho(i44) + 1
        Else
          qs$ = qs$ + Left(qanev(i44) + Space(qoho(i44)), qoho(i44)) + " "
          sh% = sh% + qoho(i44) + 1
        End If
      End If
    Next
    Print #prfi, "FL" + String(sh%, "=")
    Print #prfi, qs$
    Print #prfi, "FL" + String(sh%, "-")
    dbfi = FreeFile
    Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
    ndfi = FreeFile
    Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
    rh& = OBJTAB(qw1&).rekhossz
    If aktsor& < darab& - 1 Then
      For rcc& = 1 To darab&
        DoEvents
        If statusz% = 0 Then spz& = rcc& Else spz& = res&(rcc&)
        Get #ndfi, (spz& - 1) * ih& + 1, rcim&
        '--- rekordok beallitasa es kitoltes
        rww$ = Space(rh&)
        Get #dbfi, rcim& + 9, rww$
        qs$ = "TS"
        For i44 = 1 To qadb
          If qakell(i44) = 1 Then
            qwadat$ = Mid$(rww$, qakp(i44), qamh(i44))
            If InStr(qattr(i44), "J") > 0 Then
              qs$ = qs$ + Right(Space(qoho(i44)) + qwadat, qoho(i44)) + " "
            Else
              qs$ = qs$ + Left(qwadat + Space(qoho(i44)), qoho(i44)) + " "
            End If
          End If
        Next
        Print #prfi, qs$
      Next
    End If
    Close ndfi: Close dbfi: Close prfi
    Shell programutvonal$ + "dbx4-sho " + terminal$ + task$ + "rhl/" + listautvonal$, vbNormalFocus
  End If
End Sub

Private Sub Command4_Click()
  '--- gombsor 4.gomb
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  gombsorszam% = 4
  Call langclos
  Kereso.Hide
End Sub

Private Sub Command5_Click()
  '--- gombsor 5.gomb
  OBJTAB(ob%).obcim = oobcim&
  OBJTAB(ob%).obind = oobind&
  OBJTAB(ob%).obnex = oobnext&
  gombsorszam% = 5
  Call langclos
  Kereso.Hide
End Sub

Private Sub Command6_Click()
  '--- kezdõlap esetén kilépés  (Esc)
  '--- keresõtábla esetén mégsem
  althobjektum$ = ""
  talalat% = 0
  gombsorszam% = 0
  Call langclos
  Kereso.Hide
End Sub

Private Sub Command7_Click()
  '--- szöveg keresése (F2)
  Dim szopar$(20), szopardb%
  If darab& < 1 Then Exit Sub
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  ProgressBar1.Visible = True
  ProgressBar1.Min = 1
  ProgressBar1.Max = 100
  resdb& = 0
  '---  statusz% = 0
  rcc& = 0
  If aktsor& < darab& - 1 Then
    szopardb% = 0
    kte$ = xkonver(Trim$(Text1.Text))
    If InStr(kte$, " ") > 0 Then
      tobbszavas% = 1
      Call linpar(kte$, szopar$, " ", szopardb%)
      kte$ = szopar$(1)
    End If
    For rcc& = 1 To darab&
      ProgressBar1.Value = pscale(rcc&, darab&)
      If statusz% = 0 Then spz& = rcc& Else spz& = res&(rcc&)
      Get #ndfi, (spz& - 1) * ih& + 1, rcim&
      '--- rekordok beallitasa es kitoltes
      Get #dbfi, rcim& + 5, oobnext&
      Seek #dbfi, rcim& + 9
      rww$ = Space(rh&): Get #dbfi, , rww$: r$ = xkonver(rww$)
      If InStr(r$, kte$) > 0 Then
        If szopardb% > 0 Then
          jujo% = 1
          For juj% = 1 To szopardb%
            If InStr(r$, szopar(juj%)) = 0 Then jujo% = 0: Exit For
          Next
          If jujo% = 1 Then
            resdb& = resdb& + 1
            res&(resdb&) = spz&
            If resdb& = 100000 Then Exit For
          End If
        Else
          resdb& = resdb& + 1
          res&(resdb&) = spz&
          If resdb& = 100000 Then Exit For
        End If
      End If
    Next
  End If
  Close ndfi: Close dbfi
  ProgressBar1.Visible = False
  If resdb& = 0 Then
    Call mess(langform(17), 3, 0, langform(18), valasz%)
    'MsgBox langform(17), 48, langform(18)
    If statusz% = 1 Then resdb& = darab&
  Else
    mindnyomott% = 0
    statusz% = 1
    darab& = resdb&
    elsosor& = 1: kezd& = 1: aktsor& = 1: akt& = 1
    If darab& > tizenharom Then zar& = tizenharom Else zar& = darab&
    Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
  End If
  DoEvents
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub indexre()
  '--- direkt beállás indexre
  'resdb& = 0
  If darab& < 1 Then
    Exit Sub
  Else
    rww$ = dbxkey(keresobj$, UCase(Trim(Text1.Text)))
    If rww$ = "" Then
      Exit Sub
      'Call Command8_Click
    Else
      mindnyomott% = 0
      resdb& = 1
      res&(resdb&) = OBJTAB(obsorszama(keresobj)).obind
      statusz% = 1
      darab& = resdb&
      elsosor& = 1: kezd& = 1: aktsor& = 1: akt& = 1
      If darab& > tizenharom Then zar& = tizenharom Else zar& = darab&
      Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
      DoEvents
      If Text1.Visible = True Then Text1.SetFocus
    End If
  End If
End Sub

Private Sub Command8_Click()
  '--- mezõ keresése (F3)
  Dim szopar$(20), szopardb%
  If darab& < 1 Then Exit Sub
  If indexdarabszam& = 0 Then Exit Sub
  If keresobj = "REAN" Then
    cis$ = Trim(Text1.Text)
    If Len(cis$) = 8 Or Len(cis$) = 13 Then
      voltcis% = 0
      For jj81% = 1 To Len(cis$)
        If Mid$(cis$, jj81%, 1) = "ö" Then Mid$(cis$, jj81%, 1) = "0": voltcis% = 1
      Next
      If voltcis% = 1 Then Text1.Text = cis$: DoEvents
    End If
  End If
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  ProgressBar1.Visible = True
  ProgressBar1.Min = 1
  ProgressBar1.Max = 100
  resdb& = 0
  If aktsor& <= darab& Then
    ob% = obsorszama(keresobj$)
    i2& = MSFlexGrid1.Col + 1
    If kmehsx%(i2&) > 0 And keresokodelemszam& > 0 And indexdarabszam& = darab& Then
      '--- kereses a kereso kód táblában
      szopardb% = 0
      i2& = MSFlexGrid1.Col + 1
      startcim& = keresokodelemszam& * keresokodtabla(kmehsx%(i2&), 3)
      tobbszavas% = 0
      If Trim(Text1.Text) = langform(19) Then
        kte$ = Space(kmho%(i2&))
      Else
        kte$ = xkonver(Trim$(Text1.Text))
        If InStr(kte$, " ") > 0 Then
          tobbszavas% = 1
          Call linpar(kte$, szopar$, " ", szopardb%)
          kte$ = szopar$(1)
        End If
      End If
      ktlength% = Len(kte$): kmehsxh1% = kmehsxh%(i2&)
      vegcsillag$ = ""
      If Right(kte$, 1) = "*" Then
        ktlength% = ktlength% - 1
        kte$ = Left(kte$, ktlength%)
        vegcsillag$ = "*"
      End If
      iqwi% = 1
      If iqwi% = 1 Then
        '--- instringes keresés
        ktepozi& = startcim& + 1
        Do
          DoEvents
          talpozi& = InStr(ktepozi&, keresokodvektor$, kte$)
          If talpozi& > startcim& + keresokodelemszam& * kmehsxh1% Then Exit Do
          If talpozi& > 0 Then
            hanyadikbejegyzes& = (talpozi& - startcim&) \ kmehsxh1% + 1
            If hanyadikbejegyzes& > keresokodelemszam& Then Exit Do
            azonbelul& = (talpozi& - startcim&) Mod kmehsxh1%
            pozcim& = startcim& + (hanyadikbejegyzes& - 1) * kmehsxh1% + 1
            If azonbelul& + ktlength% <= kmehsxh1% + 1 Then
              If vegcsillag$ = "" Or azonbelul& = 1 Then
                If szopardb% > 0 Then
                  jujo% = 1
                  xxvvxx$ = Mid$(keresokodvektor$, (hanyadikbejegyzes& - 1) * kmehsxh1% + 1, kmehsxh1%)
                  For juj% = 1 To szopardb%
                    If InStr(xxvvxx$, szopar(juj%)) = 0 Then jujo% = 0: Exit For
                  Next
                  If jujo% = 1 Then
                    resdb& = resdb& + 1
                    res&(resdb&) = hanyadikbejegyzes&
                    If resdb& = 100000 Then Exit Do
                  End If
                Else
                  resdb& = resdb& + 1
                  res&(resdb&) = hanyadikbejegyzes&
                  If resdb& = 100000 Then Exit Do
                End If
              End If
            End If
            ktepozi& = startcim& + hanyadikbejegyzes& * kmehsxh1% + 1
            If ktepozi& > startcim& + keresokodelemszam& * kmehsxh1% Then Exit Do
          End If
        Loop While talpozi& > 0
      Else
        '--- ciklikus keresés
        For rcc& = 1 To keresokodelemszam&
          ProgressBar1.Value = pscale(rcc&, darab&)
          tall% = 0
          spz& = rcc&
          pozcim& = startcim& + (rcc& - 1) * kmehsxh1% + 1
          adamez$ = xkonver(Mid$(keresokodvektor$, pozcim&, kmehsxh1%))
          If Right$(kte$, 1) = "*" Then
            If InStr(adamez$, Left$(kte$, ktlength% - 1)) = 1 Then tall% = 1
          Else
            If InStr(adamez$, kte$) > 0 Then tall% = 1
          End If
          If tall% = 1 Then
            resdb& = resdb& + 1
            res&(resdb&) = spz&
            If resdb& = 100000 Then Exit For
          End If
        Next
      End If
      '--- keresõtábla vége
      If darab& > keresokodelemszam& And resdb& < 100000 Then
        For rcc& = keresokodelemszam& + 1 To darab&
          spz& = rcc&
          Get #ndfi, (rcc& - 1) * ih& + 1, rcim&
          If i2& = 1 Then
            Seek #ndfi, (spz& - 1) * ih& + 5
            rww$ = Space(kmho%(i2&)): Get #ndfi, , rww$: adamez$ = xkonver(rww$)
          Else
            '--- rekordok beallitasa es kitoltes
            adamez$ = Space(kmho%(i2&))
            Get #dbfi, rcim& + kmkp%(i2&) + 8, adamez$
            adamez$ = xkonver(adamez$)
          End If
          If Trim(Text1.Text) = langform(20) Then
            If Trim(adamez$) <> "" Then
              resdb& = resdb& + 1
              res&(resdb&) = spz&
              If resdb& = 100000 Then Exit For
            End If
          Else
            tall% = 0
            If Right$(kte$, 1) = "*" Then
              If InStr(adamez$, Left$(kte$, Len(kte$) - 1)) = 1 Then tall% = 1
            Else
              If szopardb% > 0 Then
                jujo% = 1
                For juj% = 1 To szopardb%
                  If InStr(adamez$, szopar(juj%)) = 0 Then jujo% = 0: Exit For
                Next
                If jujo% = 1 Then tall% = 1
              Else
                If InStr(adamez$, kte$) > 0 Then tall% = 1
              End If
            End If
            If tall% = 1 Then
              resdb& = resdb& + 1
              res&(resdb&) = spz&
              If resdb& = 100000 Then Exit For
            End If
          End If
        Next
      End If
    Else
      szopardb% = 0
      If Trim(Text1.Text) = langform(19) Then
        kte$ = Space(kmho%(i2&))
      Else
        kte$ = xkonver(Trim$(Text1.Text))
        If InStr(kte$, " ") > 0 Then
          tobbszavas% = 1
          Call linpar(kte$, szopar$, " ", szopardb%)
          kte$ = szopar$(1)
        End If
      End If
      For rcc& = 1 To darab&
        'DoEvents
        ProgressBar1.Value = pscale(rcc&, darab&)
        If statusz% = 0 Then spz& = rcc& Else spz& = res&(rcc&)
        i2& = MSFlexGrid1.Col + 1
        If indexdarabszam& = darab& Then
          If cimvektorvan = 1 Then
            rcim& = cimvekt(rcc&)
          Else
            Get #ndfi, (spz& - 1) * ih& + 1, rcim&
          End If
        Else
          Get #ndfi, (spz& - 1) * ih& + 1, rcim&
        End If
        If i2& = 1 Then
          Seek #ndfi, (spz& - 1) * ih& + 5
          rww$ = Space(kmho%(i2&)): Get #ndfi, , rww$: adamez$ = xkonver(rww$)
          adamez$ = xkonver(adamez$)
        Else
          '--- rekordok beallitasa es kitoltes
          adamez$ = Space(kmho%(i2&))
          Get #dbfi, rcim& + kmkp%(i2&) + 8, adamez$
          adamez$ = xkonver(adamez$)
        End If
        If Trim(Text1.Text) = langform(20) Then
          If Trim(adamez$) <> "" Then
            resdb& = resdb& + 1
            res&(resdb&) = spz&
            If resdb& = 100000 Then Exit For
          End If
        Else
          tall% = 0
          If Right$(kte$, 1) = "*" Then
            If InStr(adamez$, Left$(kte$, Len(kte$) - 1)) = 1 Then tall% = 1
          Else
            If szopardb% > 0 Then
              jujo% = 1
              For juj% = 1 To szopardb%
                If InStr(adamez$, szopar(juj%)) = 0 Then jujo% = 0: Exit For
              Next
              If jujo% = 1 Then tall% = 1
            Else
              If InStr(adamez$, kte$) > 0 Then tall% = 1
            End If
          End If
          If tall% = 1 Then
            resdb& = resdb& + 1
            res&(resdb&) = spz&
            If resdb& = 100000 Then Exit For
          End If
        End If
      Next
    End If
  End If
  Close ndfi: Close dbfi
  ProgressBar1.Visible = False
  If resdb& = 0 Then
    Call mess(langform(17), 3, 0, langform(18), valasz%)
    If statusz% = 1 Then resdb& = darab&
  Else
    mindnyomott% = 0
    statusz% = 1
    darab& = resdb&
    elsosor& = 1: kezd& = 1: aktsor& = 1: akt& = 1
    If darab& > tizenharom Then zar& = tizenharom Else zar& = darab&
    Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
  End If
  DoEvents
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Command9_Click()
  '--- intervallum keresés (F5)
  If darab& < 1 Then Exit Sub
  minus% = InStr(Text1.Text, "-")
  If minus% > 1 Then
    tol@ = xval(Left$(Text1.Text, minus% - 1))
    ig@ = xval(Mid$(Text1.Text, minus% + 1))
    katol$ = RTrim(Left$(Text1.Text, minus% - 1))
    tolh% = Len(katol$)
    katig$ = RTrim(Mid$(Text1.Text, minus% + 1))
    igh% = Len(katig$)
    dbfi = FreeFile
    Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
    ndfi = FreeFile
    Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
    ProgressBar1.Visible = True
    ProgressBar1.Min = 1
    ProgressBar1.Max = 100
    resdb& = 0
    If aktsor& < darab& - 1 Then
      For rcc& = 1 To darab&
        ProgressBar1.Value = pscale(rcc&, darab&)
        If statusz% = 0 Then spz& = rcc& Else spz& = res&(rcc&)
        i2& = MSFlexGrid1.Col + 1
        Get #ndfi, (spz& - 1) * ih& + 1, rcim&
        If i2& = 1 Then
          Seek #ndfi, (spz& - 1) * ih& + 5
          rww$ = Space(kmho%(i2&)): Get #ndfi, , rww$: adamez$ = rww$
        Else
          '--- rekordok beallitasa es kitoltes
          Get #dbfi, rcim& + 5, oobnext&
          Seek #dbfi, rcim& + 9
          rww$ = Space(rh&): Get #dbfi, , rww$: r$ = UCase$(rww$)
          adamez$ = Mid$(r$, kmkp%(i2&), kmho%(i2&))
        End If
        katri$ = kmtb(i2&)
        ert@ = xval(adamez$)
        'ert@ = Val(Mid$(r$, kmkp%(i2&), kmho%(i2&)))
        If InStr(katri$, "NJ") > 0 Then
          If ert@ >= tol@ And ert@ <= ig@ Then
            resdb& = resdb& + 1
            res&(resdb&) = spz&
            If resdb& = 100000 Then Exit For
          End If
        Else
          If igh% > tolh% Then ssho% = igh% Else ssho% = tolh%
          If xkonver(katol$) <= xkonver(Left(adamez$ + Space(ssho%), tolh%)) Then
            If xkonver(katig$) >= xkonver(Left(adamez$ + Space(ssho%), igh%)) Then
              resdb& = resdb& + 1
              res&(resdb&) = spz&
              If resdb& = 100000 Then Exit For
            End If
          End If
        End If
      Next
    End If
    Close ndfi: Close dbfi
    ProgressBar1.Visible = False
    If resdb& = 0 Then
      Call mess(langform(17), 2, 0, langform(18), valasz%)
      'MsgBox langform(17), 48, langform(18)
      If statusz% = 1 Then resdb& = darab&
    Else
      mindnyomott% = 0
      statusz% = 1
      darab& = resdb&
      elsosor& = 1: kezd& = 1: aktsor& = 1: akt& = 1
      If darab& > tizenharom Then zar& = tizenharom Else zar& = darab&
      Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
    End If
  End If
  DoEvents
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Form_Activate()
  '--- form aktíválása
  '--- kezdeti értékek beállítása
  '--- nyitó táblázat megmutatása
  For izu% = 1 To MSFlexGrid1.Cols
    colwi(izu%) = MSFlexGrid1.ColWidth(izu% - 1)
  Next
  ob% = obsorszama(keresobj$)
  dbxn$ = dbxneve(keresobj$)
  w1% = OBJTAB(ob%).obi(1)
  indn$ = RTrim$(INDTAB(w1%).indnev)
  ih& = ADATAB(INDTAB(w1%).adatsorsz).adatho + 5
  rh& = OBJTAB(ob%).rekhossz
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  rc& = Int(LOF(ndfi) / ih&)
  indexdarabszam& = rc&
  Close ndfi
  If althobjektum$ = keresobj$ Then
    If resxdb& > 0 Then
      statusz% = wwtabstatusz%
      If statusz% <> 0 Then
        For i11& = 1 To resxdb&
          res&(i11&) = resx&(i11&)
        Next
        resdb& = resxdb&
        darab& = resxdb&
        teldarab& = rc&
        elsosor& = wwtoprow&
        aktsor& = wwaktrow&
        kezd& = wwkezdpo&
        zar& = wwzarpo&
        akt& = wwaktpo&
        Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
        GoTo xveg
      End If
    End If
  End If
  If rc& > 0 Then
    If keresomod% = 0 And toprow& <> 0 Then
      darab& = rc&
      statusz% = tabstatusz%
      If statusz% <> 0 Then
        If resmdb& > 0 Then
          For i11& = 1 To resmdb&
            res&(i11&) = resm&(i11&)
          Next
          resdb& = resmdb&
          darab& = resdb&
        End If
      End If
      teldarab& = rc&
      MSFlexGrid1.CellBackColor = &HC0FFC0
      If aktrow& = -1 Then
        '--- végére állunk
        statusz% = 0
        elsosor& = 1: darab& = teldarab&: aktsor& = 1
        kezd& = 1: akt& = 1: If darab& > tizenharom Then zar& = tizenharom Else zar& = darab&
        aktsor& = rc&
        If rc& < tizenharom Then
          elsosor& = 1: kezd& = 1: zar& = rc&: akt& = rc&: aktsor& = rc&
        Else
          zar& = rc&: kezd& = zar& - tizenketto: akt& = tizenharom
          aktsor& = rc&: elsosor& = zar& - tizenketto
        End If
        toprow& = elsosor&: aktrow& = aktsor&: kezdpo& = kezd&
        zarpo& = zar&: aktpo& = akt&
        Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
      Else
        '--- fotabla masodik
        elsosor& = toprow&
        aktsor& = aktrow&
        kezd& = kezdpo&
        zar& = zarpo&
        akt& = aktpo&
        Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
      End If
    Else
      '--- fõtábla, altábla elsõ
      statusz% = 0
      teldarab& = rc&
      MSFlexGrid1.Col = 0: MSFlexGrid1.Row = 1
      MSFlexGrid1.CellBackColor = &HC0FFC0
      elsosor& = 1: aktsor& = 1: darab& = rc&
      If darab& < tizenharom Then kezd& = 1: zar& = darab&: akt& = 1 Else kezd& = 1: zar& = tizenharom: akt& = 1
      Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
      If OBJTAB(ob%).keresokoddarab > 0 Then
        '--- keresõkód táblák beolvasása
        hsxfi = FreeFile
        Open auditorutvonal$ + "auw-" + keresobj$ + ".hsx" For Binary Shared As #hsxfi
        fim& = LOF(hsxfi)
        bejegyzeshossz& = 0
        For i2& = 1 To adadb%
          If ADATAB(i2&).obsorsz = ob% Then
            w1% = i2&
            If ADATAB(w1%).keresokodindex > 0 Then
              keresokodtabla%(ADATAB(w1%).keresokodindex, 1) = ADATAB(w1%).adatkp
              keresokodtabla%(ADATAB(w1%).keresokodindex, 2) = ADATAB(w1%).keresokodhossz
              keresokodtabla%(ADATAB(w1%).keresokodindex, 3) = bejegyzeshossz&
              bejegyzeshossz& = bejegyzeshossz& + ADATAB(w1%).keresokodhossz
            End If
          End If
        Next
        If bejegyzeshossz& <> 0 Then
          keresokodelemszam& = fim& / bejegyzeshossz&
          keresokodvektor$ = Space(fim&)
          Get #hsxfi, 1, keresokodvektor$
        Else
          keresokodelemszam& = 0
        End If
        Close hsxfi
      End If
    End If
    If cimvektorvan = 1 Then
      If darab& = teldarab& Then Call cimvektor(keresobj, indexdarabszam&)
    End If
  End If
xveg:
  DoEvents
  If autoinfo = 1 Then infolapbe = 0: Call Command26_Click
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub megmutat(dbxn$, indn$, kezd&, akt&, zar&)
  '--- aktuális táblázat megmutatas kereso.msflexgrid1-ben
  dbfi = FreeFile
  Open auditorutvonal$ + dbxn$ + ".dbx" For Binary Shared As #dbfi
  ndfi = FreeFile
  Open auditorutvonal$ + indn$ + ".ndx" For Binary Shared As #ndfi
  sdb% = zar& - kezd& + 1
  For i& = 1 To sdb%
    inex& = kezd& + i& - 1
    If statusz% = 0 Then spz& = inex& Else spz& = res&(inex&)
    If spz& > 0 Then
      Get #ndfi, (spz& - 1) * ih& + 1, rcim&
      '--- rekordok beallitasa es kitoltes
      tjele$ = " ": Get #dbfi, rcim& + 4, tjele$
      Seek #dbfi, rcim& + 9
      r$ = Space(rh&): Get #dbfi, , r$
      For i2& = 1 To kmesor%(0)
        If tjele$ = "*" And i2& = 1 Then
          adamez$ = "* " + Mid$(r$, kmkp%(i2&), kmho%(i2&))
        Else
          adamez$ = Mid$(r$, kmkp%(i2&), kmho%(i2&))
        End If
        MSFlexGrid1.TextMatrix(i&, i2& - 1) = adamez$
      Next
    End If
  Next
  If sdb% < tizenharom Then
    For i& = sdb% + 1 To tizenharom
      For i2& = 1 To kmesor%(0)
        adamez$ = Space$(kmho%(i2&))
        MSFlexGrid1.TextMatrix(i&, i2& - 1) = adamez$
      Next
    Next
  End If
  MSFlexGrid1.Row = akt&
  MSFlexGrid1.RowSel = MSFlexGrid1.Row
  MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
  '--- rekord beolvasasa es objektum beallitasa
  inex& = kezd& + akt& - 1
  If statusz% = 0 Then spz& = inex& Else spz& = res&(inex&)
  If spz& > 0 Then
    Get #ndfi, (spz& - 1) * ih& + 1, rcim&
    '--- rekordok beallitasa es kitoltes
    If Text4.Visible = True Then Text4.Text = Str$(rcim&)
    Seek #dbfi, rcim& + 4
    torlojel$ = " ": Get #dbfi, , torlojel$
    Get #dbfi, rcim& + 5, oobnext&
    Seek #dbfi, rcim& + 9
    r$ = Space(rh&): Get #dbfi, , r$
    If spz& <> ebbenvoltsor& Then
      ebbenvoltsor& = spz&
      If infolapbe = 1 Then Call infomutat(keresobj$, r$, "Kereso")
    End If
    oobcim& = rcim&
    oobind& = spz&
    rekord$ = r$
  End If
  Close dbfi
  Close ndfi
  Text2.Text = langform(21) + Str$(darab&)
  If torlojel$ = "*" Then
    rekord$ = ""
    Text3.ForeColor = QBColor(4)
    Text3.FontBold = True
    Text3.Text = langform(23)
  Else
    Text3.ForeColor = QBColor(0)
    Text3.Text = langform(22) + Str$(aktsor&)
  End If
  sorbak& = akt&
  If keresomod% = 0 Then
    toprow& = elsosor&: aktrow& = aktsor&
    kezdpo& = kezd&: zarpo& = zar&: aktpo& = akt&
    tabstatusz% = statusz%
    If statusz% <> 0 Then
      resmdb& = resdb&
      For i11& = 1 To resdb&
        resm&(i11&) = res&(i11&)
        If i11& = 100000 Then resmdb& = 100000: Exit For
      Next
    End If
  Else
    wwtoprow& = elsosor&: wwaktrow& = aktsor&
    wwkezdpo& = kezd&: wwzarpo& = zar&: wwaktpo& = akt&
    wwtabstatusz% = statusz%
    If statusz% <> 0 Then
      If resdb& < 2000 Then
        resxdb& = resdb&
        For ru& = 1 To resdb&: resx&(ru&) = res&(ru&): Next
      Else
        althobjektum$ = ""
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  MSFlexGrid1.Font.Size = 8
  Call langinit("kereso", 2)
  Call szkriptel("kereso")
  tizenharom = 13: tizenketto = 12: keresrejt = 0
End Sub

Private Sub Form_Resize()
  If Kereso.Width < 5000 Then Kereso.Width = 5000
  MSFlexGrid1.Width = Kereso.Width - 825
  Command19.Left = Kereso.Width - 600
  Command20.Left = Kereso.Width - 600
  Command21.Left = Kereso.Width - 600
  Command22.Left = Kereso.Width - 600
  Command23.Left = Kereso.Width - 600
  Command24.Left = Kereso.Width - 600
  Command25.Left = Kereso.Width - 600
  Command26.Left = Kereso.Width - 600
  Command30.Left = Kereso.Width - 600
End Sub

Private Sub MSFlexGrid1_Click()
  '--- aktuális állapot megmutattása
  '--- focus átirányítása text1-re
  If elsosor& + MSFlexGrid1.Row - 1 > darab& Then
    MSFlexGrid1.Row = darab& - elsosor& + 1
  End If
  aktsor& = elsosor& + MSFlexGrid1.Row - 1
  MSFlexGrid1.RowSel = MSFlexGrid1.Row
  MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
  kezd& = elsosor&
  akt& = aktsor& - elsosor& + 1
  zar& = elsosor& + tizenketto
  If zar& > darab& Then zar& = darab&
  Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
  DoEvents
  If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_DblClick()
  '--- választás dupla kattintással
  akibaki& = aktsor&
  aktsor& = elsosor& + MSFlexGrid1.Row - 1
  If aktsor& > darab& Then aktsor& = akibaki&
  MSFlexGrid1.RowSel = MSFlexGrid1.Row
  MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
  kezd& = elsosor&
  akt& = aktsor& - elsosor& + 1
  zar& = elsosor& + tizenketto
  If zar& > darab& Then zar& = darab&
  Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
  talalat% = 1
  Call Command12_Click
End Sub

Private Sub MSFlexGrid1_entercell()
  '--- kiválasztott cella színezése
  MSFlexGrid1.CellBackColor = &HC0FFC0
End Sub

Private Sub MSFlexGrid1_LeaveCell()
  '--- elhagyott cella színezés törlése
  MSFlexGrid1.CellBackColor = QBColor(15)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- billentyû lenyomása text1-ben
  If darab& = 0 And KeyCode <> vbKeyInsert Then Exit Sub
  Select Case KeyCode
    Case vbKeyF1
      '--- teljes táblázat visszaálítása (F1)
      Call Command10_Click
      KeyCode = 0
    Case vbKeyF2
      '--- szöveg keresés teljes rekordban (F2)
      Call Command7_Click
      KeyCode = 0
    Case vbKeyF6
      '--- fõindex keresése
      Call indexre
    Case vbKeyF5
      '--- intervallum keresés (F5)
      Call Command9_Click
      KeyCode = 0
    Case vbKeyF3
      '--- rovat keresés (F3)
      Call Command8_Click
      KeyCode = 0
    Case vbKeyF4
      '--- pozicionálás (F4)
      Call Command11_Click
      KeyCode = 0
    Case vbKeyEscape
      KeyCode = 0
      Call Command6_Click
    Case vbKeyReturn
      '--- választás (Enter)
      If zkermod% = 0 Then
        '--- kezdõlap
        If keresobj = "KTRM" And programnev$ = "AUW-RTRM" Then
          eankodja$ = Left(Trim(Text1.Text) + Space(13), 13)
          atex$ = eankodja$
          For iaa% = 1 To 13
            If Mid$(atex$, iaa%, 1) = "ö" Then Mid$(atex$, iaa%, 1) = "0"
          Next
          eankodja$ = atex$
          eanrecja$ = dbxkey("REAN", eankodja$)
          If eanrecja$ <> "" Then
            Text1.Text = Mid$(eanrecja$, 14, 15)
            Call indexre
          End If
        Else
          Call Command8_Click
          KeyCode = 0
        End If
      Else
        '--- keresõ tábla
        talalat% = 1
        Call Command12_Click
      End If
    Case vbKeyLeft
      '--- balra egy oszloppal
      If MSFlexGrid1.Col > 0 Then MSFlexGrid1.Col = MSFlexGrid1.Col - 1
      If MSFlexGrid1.LeftCol > MSFlexGrid1.Col Then
        Do While MSFlexGrid1.ColIsVisible(MSFlexGrid1.Col) = False
          MSFlexGrid1.LeftCol = MSFlexGrid1.LeftCol - 1
        Loop
      Else
        If MSFlexGrid1.ColIsVisible(MSFlexGrid1.Col) = False Then MSFlexGrid1.LeftCol = MSFlexGrid1.Col
      End If
      KeyCode = 0
    Case vbKeyRight
      '--- jobbra egy oszloppal
      If MSFlexGrid1.Col < MSFlexGrid1.Cols - 1 Then MSFlexGrid1.Col = MSFlexGrid1.Col + 1
      If MSFlexGrid1.LeftCol > MSFlexGrid1.Col Then MSFlexGrid1.LeftCol = MSFlexGrid1.Col
      Do While MSFlexGrid1.ColIsVisible(MSFlexGrid1.Col) = False
        MSFlexGrid1.LeftCol = MSFlexGrid1.LeftCol + 1
      Loop
      KeyCode = 0
    Case vbKeyUp
      '--- elõzõ sor
      aktsor& = aktsor& - 1: If aktsor& < 1 Then aktsor& = 1
      If aktsor& < elsosor& Then elsosor& = elsosor& - 1
      KeyCode = 0
    Case vbKeyDown
      '--- következõ sor
      aktsor& = aktsor& + 1: If aktsor& > darab& Then aktsor& = darab&
      If aktsor& > elsosor& + tizenketto Then elsosor& = elsosor& + 1
      KeyCode = 0
    Case vbKeyPageUp
      '--- lapozás vissza
      aktsor& = aktsor& - tizenharom
      If aktsor& < 1 Then
        aktsor& = 1: elsosor& = 1
      Else
        elsosor& = elsosor& - tizenharom
        If elsosor& < 1 Then elsosor& = 1
      End If
      KeyCode = 0
    Case vbKeyPageDown
      '--- lapozás elõre
      aktsor& = aktsor& + tizenharom
      If aktsor& > darab& Then
        aktsor& = darab&
        elsosor& = darab& - tizenharom + 1
        If elsosor& < 1 Then elsosor& = 1
      Else
        elsosor& = elsosor& + tizenharom
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
      If darab& < tizenharom Then
        elsosor& = 1
      Else
        elsosor& = darab& - tizenharom + 1
      End If
      KeyCode = 0
    Case vbKeyInsert
      If Command1.Visible = True Then Call Command1_Click
    Case vbKeyF12
      If Shift Then
        'unijog = 1
      End If
    Case Else
      If Shift And vbAltMask Then
        carr$ = Chr$(KeyCode)
        If Command1.Visible = True And UCase(Left$(Command1.Caption, 1)) = carr$ Then Call Command1_Click
        If Command2.Visible = True And UCase(Left$(Command2.Caption, 1)) = carr$ Then Call Command2_Click
        If Command3.Visible = True And UCase(Left$(Command3.Caption, 1)) = carr$ Then Call Command3_Click
        If Command4.Visible = True And UCase(Left$(Command4.Caption, 1)) = carr$ Then Call Command4_Click
        If Command5.Visible = True And UCase(Left$(Command5.Caption, 1)) = carr$ Then Call Command5_Click
      End If
  End Select
  kezd& = elsosor&
  akt& = aktsor& - elsosor& + 1
  zar& = elsosor& + tizenketto
  If zar& > darab& Then zar& = darab&
  Call megmutat(dbxn$, indn$, kezd&, akt&, zar&)
End Sub

Private Sub Text5_Change()
  '--- üres
End Sub


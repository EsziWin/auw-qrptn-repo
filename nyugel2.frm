VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Nyugel1 
   BackColor       =   &H0055D7F7&
   Caption         =   "Új vevõ"
   ClientHeight    =   8208
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   12252
   HelpContextID   =   100
   LinkTopic       =   "Form2"
   ScaleHeight     =   8208
   ScaleWidth      =   12252
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mégsem"
      Height          =   372
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7850
      Width           =   972
   End
   Begin VB.CheckBox CheckTámop 
      BackColor       =   &H0055D7F7&
      Caption         =   "Támop más."
      Height          =   252
      Left            =   5280
      TabIndex        =   45
      Top             =   7560
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0FFFF&
      Caption         =   "TÁMOP szla"
      Height          =   372
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7440
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0055D7F7&
      Caption         =   "Szállító lev."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   6.6
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6600
      TabIndex        =   5
      Top             =   7560
      Width           =   972
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Napi összesítõ"
      Height          =   372
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7850
      Width           =   1092
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   156
      Left            =   120
      TabIndex        =   37
      Top             =   7188
      Visible         =   0   'False
      Width           =   12012
      _ExtentX        =   21188
      _ExtentY        =   275
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kedvezmény"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Számla kedvezmény"
      Top             =   7440
      Width           =   972
   End
   Begin VB.ListBox eloleg 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3516
      Left            =   450
      TabIndex        =   36
      Top             =   3240
      Visible         =   0   'False
      Width           =   7452
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zárás feladása"
      Height          =   372
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7850
      Width           =   1212
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kassza összesítõ"
      Height          =   372
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Pénztárgépes napi forgalom összesítõje"
      Top             =   7850
      Width           =   1092
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Elõleg"
      Height          =   372
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Elõleg beszámítás"
      Top             =   7440
      Width           =   852
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sztornó, számla másolat"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Számla, nyugta sztornózása, újranyomtatása, számla szállító levelrõl, származtatás"
      Top             =   7440
      Width           =   1212
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bevételezés"
      Height          =   372
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Egyéb készlet mozgás"
      Top             =   7440
      Width           =   1092
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Munkalap"
      Height          =   372
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Munkalap"
      Top             =   7440
      Width           =   1092
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Megrendelés"
      Height          =   372
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Vevõ megrendelés rögzítése"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Garancia"
      Height          =   372
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Kiadás garancia jegyre"
      Top             =   7850
      Width           =   972
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pénztár"
      Height          =   372
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Ellátmány, lefölözés rögzítése, pénztár"
      Top             =   7850
      Width           =   972
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hitel számla"
      Height          =   372
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Átutalások ill. bankkártyás számla"
      Top             =   7850
      Width           =   1212
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0055D7F7&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   732
      Left            =   1920
      TabIndex        =   31
      Top             =   7560
      Width           =   1212
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Szállítólevél"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8280
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0091E9FB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   288
      Left            =   840
      TabIndex        =   28
      Top             =   5760
      Width           =   1332
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C5FCC8&
      Caption         =   "Sor törlés"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7440
      Width           =   1452
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2292
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   12252
      Begin VB.TextBox Text13 
         Height          =   288
         Left            =   6120
         TabIndex        =   49
         Text            =   "Text13"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C B"
         Height          =   612
         Left            =   7390
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Cím felbontása"
         Top             =   480
         Width           =   252
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cím felbontása"
         Height          =   252
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Cím felbontása"
         Top             =   120
         Width           =   1212
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H0091E9FB&
         Height          =   288
         Left            =   4440
         TabIndex        =   42
         Top             =   1200
         Width           =   1452
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H0091E9FB&
         Height          =   288
         Left            =   4440
         MaxLength       =   58
         TabIndex        =   41
         Top             =   1920
         Width           =   7692
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Kp partner feltöltés"
         Height          =   252
         Left            =   10680
         TabIndex        =   40
         Top             =   1200
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.TextBox Text10 
         Height          =   372
         Left            =   3480
         TabIndex        =   39
         Text            =   "Text10"
         Top             =   1920
         Visible         =   0   'False
         Width           =   2172
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H0091E9FB&
         Height          =   288
         Left            =   1200
         TabIndex        =   21
         Text            =   "Text9"
         Top             =   1920
         Width           =   1212
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H0091E9FB&
         Height          =   288
         Left            =   4440
         MaxLength       =   60
         TabIndex        =   22
         Top             =   1560
         Width           =   7692
      End
      Begin VB.TextBox Text7 
         Height          =   288
         Left            =   3480
         TabIndex        =   34
         Text            =   "01"
         Top             =   1200
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H0091E9FB&
         Height          =   288
         Left            =   1200
         TabIndex        =   20
         Top             =   1560
         Width           =   1212
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H0091E9FB&
         Height          =   288
         Left            =   1200
         TabIndex        =   19
         Top             =   1200
         Width           =   2172
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H0091E9FB&
         Height          =   288
         Left            =   1200
         TabIndex        =   18
         Top             =   840
         Width           =   6132
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0091E9FB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1200
         TabIndex        =   17
         Top             =   480
         Width           =   6132
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H0091E9FB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Top             =   120
         Width           =   2052
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Adószám:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   3480
         TabIndex        =   43
         Top             =   1200
         Width           =   852
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Teljesítés kelte:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   0
         TabIndex        =   38
         Top             =   1920
         Width           =   1692
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Megjegyzés:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   3360
         TabIndex        =   35
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fiz. hat.idõ:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   1332
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fiz. mód:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   28.8
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   732
         Left            =   7800
         TabIndex        =   29
         Top             =   240
         Width           =   4212
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0067670C&
         FillStyle       =   0  'Solid
         Height          =   972
         Left            =   7680
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   4452
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cím:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   852
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Név:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   852
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vevõkód:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   852
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kp-s számla"
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Készpénz fizetési számla"
      Top             =   7850
      Width           =   1212
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4692
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   12252
      _ExtentX        =   21611
      _ExtentY        =   8276
      _Version        =   327680
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nyugta"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nyugta"
      Top             =   7850
      Width           =   852
   End
End
Attribute VB_Name = "Nyugel1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' A munkalap számlázása esetén a mennyiséget ne lehessen töbre venni
' mint ami ki lett a munkalapra adva és temékkódot sem lehet módosítani.
Dim utermkod$, billscr%, uText1$, teljesossz@
Dim gyariszamok$(200, 200), hivatkozas$(200, 200), fizetesimod$()
Dim nt$(100), mt$(1000)
Dim afakodok$(5), afaalapok@(5), afaosszegek@(5), afaszamlak$(5)
Dim ellszla$(50), ellosszegek@(50), devellosszegek@(50), elto$(50), krak$, kertmoz$
Dim szlaszok$(90000), tbizikt$(90000)



Public betoltve%, ipartkod$, parteng@, fizmod&, vaneloleg

Const BANKPARTN As String = "BANK           "
Const KASSAPARTN As String = "KAS            "
Const KASSAFIZM As String = "01"
Private Function Keszletellenorzes()

If form1.megrendelesbol Then
          ' Készlet ellenõrzés, foglalt  készlet feloldása
            
            kmegrec$ = dbxkey("KMEG", form1.megrend$)
            krak$ = Mid$(kmegrec$, 480, 4)
            stat$ = Mid$(kmegrec$, 192, 1)

            hibas = False
            For i1% = 1 To 200
              tkod$ = Left(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) + Space(15), 15)
              ktrmrec$ = dbxkey("KTRM", tkod$)
              ' Készlet kezelés
              If Not Mid$(ktrmrec$, 443, 1) = "N" Then
              
                If Mid$(termrec$, 846, 1) = "L" Then
                   Call mess("Termék letiltva!", 2, 0, langmodul(99), valasz%)
                   hibas = True
                Else
                If Trim(tkod$) <> "" Then
                   rkszkod$ = krak$ + tkod$
                   rkszrec$ = dbxkey("RKSZ", rkszkod$)
                   keszle@ = 0
                   If rkszrec$ <> "" Then
                      If stat$ = "D" Then
                        keszle@ = xval(Mid$(rkszrec$, 20, 12))
                      Else
                        keszle@ = xval(Mid$(rkszrec$, 20, 12)) - xval(Mid$(rkszrec$, 32, 12))
                      End If
                   End If
                   menny@ = Val(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
                   If menny@ > keszle@ Then
                     Call mess(Str(i1%) + ". sorban nincs elég készlet!/" + Str(keszle@) + "/", 2, 0, "Hiba", valasz%)
                     hibas = True
                   End If
                End If
                End If
              End If
           Next
           Keszletellenorzes = True
           If hibas Then
              Keszletellenorzes = False
           End If
          
      Else
        
        ' 2015.03.17
        'krak$ = form1.Text5.Text
        krak$ = form1.ttkrak$
 
        hibas = False
        krak$ = form1.ttkrak$
        For i1% = 1 To 200
              tkod$ = Left(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 1) + Space(15), 15)
              ktrmrec$ = dbxkey("KTRM", tkod$)
              ' Készlet kezelés
              If Not Mid$(ktrmrec$, 443, 1) = "N" Then
                kkftikt$ = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 6)
                If Len(kkftikt$) = 7 Then
                   kkftrec$ = dbxkey("KKFT", kkftikt$)
                   If Not Trim$(kkftrec$) = "" Then
                      krak$ = Mid$(kkftrec$, 24, 4)
                   End If
                Else
                   ' 2015.03.17
                   'krak$ = Left(Trim$(form1.Text5.Text) + Space$(4), 4)
                   'krak$ = form1.ttkrak$
                   ' 2015.07.29
                   krak$ = Nyugel1.MSFlexGrid1.TextMatrix(i1%, 8)
                End If
                If Trim(tkod$) <> "" Then
                   rkszkod$ = krak$ + tkod$
                   rkszrec$ = dbxkey("RKSZ", rkszkod$)
                   keszle@ = 0
                   If rkszrec$ <> "" Then
                      keszle@ = xval(Mid$(rkszrec$, 20, 12))
                      If form1.szallitobol Then
                        keszle@ = keszle@ + Val(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
                      End If
                   End If
                   menny@ = Val(Nyugel1.MSFlexGrid1.TextMatrix(i1%, 3))
                   If menny@ > keszle@ Then
                 
                     Call mess(Str(i1%) + ". sorban nincs elég készlet! /" + Str(keszle@) + "/", 2, 0, "Hiba", valasz%)
                     hibas = True
                   End If
                End If
              End If
        Next
        Keszletellenorzes = True
        If hibas Then
           Keszletellenorzes = False
        End If

       
      End If

End Function
Public Sub sorttorol()
  If MSFlexGrid1.Row < 200 Then
    For i77% = MSFlexGrid1.Row To 199
      For i78% = 1 To 8
        MSFlexGrid1.TextMatrix(i77%, i78%) = MSFlexGrid1.TextMatrix(i77% + 1, i78%)
      Next
      
      For i78% = 1 To 200
        gyariszamok$(i77%, i78%) = gyariszamok$(i77% + 1, i78%)
        hivatkozas$(i77%, i78%) = hivatkozas$(i77% + 1, i78%)
      Next

    Next
  End If
  For i78% = 1 To 8
    MSFlexGrid1.TextMatrix(200, i78%) = ""
  Next
  For i78% = 1 To 200
    gyariszamok$(200, i78%) = ""
    hivatkozas$(200, i78%) = ""
  Next
End Sub
Public Function NtAtad(ind%)
 NtAtad = nt$(ind%)
End Function
Public Sub GyariszamTorol()
For j2% = 1 To 200
 For i2% = 1 To 200
  gyariszamok$(j2%, i2%) = ""
 Next
Next
For j2% = 1 To 200
 For i2% = 1 To 200
  hivatkozas$(j2%, i2%) = ""
 Next
Next

End Sub

Public Sub GyariszamFeltolt(sor%, gysz$, hiv$)
For i2% = 1 To 200
  If gyariszamok$(sor%, i2%) = "" Then
     gyariszamok$(sor%, i2%) = gysz$
     hivatkozas$(sor%, i2%) = hiv$
     Exit For
  End If
Next
End Sub

Public Sub GyariszamBetolt(sor%, Ikt$)
For i2% = 1 To 200
  ktikt$ = Ikt$ + Right("000" + Trim(Str(i2%)), 3)
  kszxrec$ = dbxkey("KSZX", ktikt$)
  If kszxrec$ = "" Then
     gyariszamok$(sor%, i2%) = ""
  Else
     gyariszamok$(sor%, i2%) = Mid$(kszxrec$, 200, 40)
     
     kkfxikt1$ = Mid$(kszxrec$, 330, 10)
     kkfxrec1$ = dbxkey("KKFX", kkfxikt1$)
     If Not kkfxrec1$ = "" Then
       Mid$(kkfxrec1$, 57, 11) = Space(11)
       Call dbxki("KKFX", kkfxrec1$, ";", "", "", hiba%)
     End If
     
  End If
Next
End Sub
Public Function GyariszamAtad(sor%, gy$(), h$()) As String
a = ""
db% = 0
For i1% = 1 To 200
  If Not Trim(gyariszamok$(sor, i1%)) = "" Then
    a = a + gyariszamok$(sor, i1%) + ";"
    gy$(i1%) = gyariszamok$(sor, i1%)
    h$(i1%) = hivatkozas$(sor, i1%)
    db% = db% + 1
  End If
Next
dbs = Str(db%)
a = dbs + ":" + a
GyariszamAtad = a

End Function
Private Sub Gyariszambeker(megnev$, db%)

             rogzites% = 0
             If db% > 200 Then
                Call mess("Egy tételsorhoz csak 200 gyáriszámot rögzíthet! Több tételben vigye fel!", 3, 0, "Hiba", valasz%)
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
             Gyariszam.MSFlexGrid1.TextMatrix(0, 1) = "S/N"
             Gyariszam.Label2.Visible = False
             Gyariszam.Label3.Visible = False
             Gyariszam.Text2.Visible = False
             Gyariszam.Text3.Visible = False
             Gyariszam.MSFlexGrid1.Height = 4292

             ' Feltöltés
             For j1% = 1 To 200
                Gyariszam.MSFlexGrid1.TextMatrix(j1%, 0) = Trim(Str(j1%))
                If Not Trim(gyariszamok$(MSFlexGrid1.Row, j1%)) = "" Then
                   Gyariszam.MSFlexGrid1.TextMatrix(j1%, 1) = gyariszamok$(MSFlexGrid1.Row, j1%)
                   Gyariszam.MSFlexGrid1.TextMatrix(j1%, 2) = hivatkozas$(MSFlexGrid1.Row, j1%)
                End If
             Next
             Gyariszam.darab = db%
             Gyariszam.biztip = "ER"
            
             Gyariszam.cikk = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
             Gyariszam.raktar = form1.Text5.Text
             
             
             Gyariszam.Label1.Caption = megnev$
             Gyariszam.betoltvegy = 0
             Gyariszam.Show vbModal
             If rogzites% = 1 Then
                For j1% = 1 To 200
                   If Not Trim(Gyariszam.MSFlexGrid1.TextMatrix(j1%, 1)) = "" Then
                      gyariszamok$(MSFlexGrid1.Row, j1%) = Gyariszam.MSFlexGrid1.TextMatrix(j1%, 1)
                      hivatkozas$(MSFlexGrid1.Row, j1%) = Gyariszam.MSFlexGrid1.TextMatrix(j1%, 2)
                   Else
                      gyariszamok$(MSFlexGrid1.Row, j1%) = ""
                      hivatkozas$(MSFlexGrid1.Row, j1%) = ""
                   End If
                Next
             End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
   Command8.Caption = "Garancia"
ElseIf Check1.Value = 1 Then
   Command8.Caption = "Szállító lev."
End If
End Sub

Private Sub Command1_Click()
  '--- kisker nyugta
  If vegosszege(1) < 0 Then
    Call mess("Végösszeg nem lehet negatív!", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If
  '--- nyugta
  If xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Üres nyugta!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  
  fizetnem@ = Val(Label1.Caption)
  Call kerekit510(fizetnem@, fizetni@, kerek@, "K")
  Label1.Caption = ertszam(Str(fizetni@), 12, 2)
  
  If Not (Trim(Text5.Text) = "Készpénz" Or Trim(Text5.Text) = "") Then
    Call mess("Csak készpénz fizetési mód lehet", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If
 
  
  If vaneloleg Then
    Call mess("Elõleget számlába számítson be!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  ' Gyáriszám ellenõrzés
  For i1% = 1 To 200
    termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(i1%, 1)) + Space(15), 15)
    If Not Trim(termkod$) = "" Then
        ktrmrec$ = dbxkey("KTRM", termkod$)
        If Mid$(ktrmrec$, 849, 1) = "I" Then
           db% = 0
           For j1% = 1 To 200
             If Not Trim(gyariszamok$(i1%, j1%)) = "" Then
                db% = db% + 1
             End If
           Next
           If Not db% = xval(MSFlexGrid1.TextMatrix(i1%, 3)) Then
              Call mess(Str(i1%) + ".sorban nem " + MSFlexGrid1.TextMatrix(i1%, 3) + " gyári szám van rögzítve!", 3, 0, "Figyelmeztetés", valasz%)
              Exit Sub
           End If
        End If
    End If
  Next
  form1.nyugtavolt = 1
  
  If Not Keszletellenorzes Then
     Exit Sub
  End If

  
  ' Eszi - Visszajáró számítás
  Visszajaro.Label5 = "A végösszeget üsse be a pénztárgépbe!"
  Visszajaro.Label6 = ""
  Visszajaro.Label7 = ""
  
  Visszajaro.Label2(1) = Right$(Space$(12) + Format(Val(Nyugel1.Label1), "# ### ### ##0"), 12)
  
  Visszajaro.Show vbModal
  
  'Call mess("A végösszeget üsse be a pénztárgépbe!", 4, 0, "Nyugta befejezése", valasz%)
  
  Nyugel1.Hide
End Sub

Private Sub Command10_Click()
  If Not xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Nem üres a nyugta!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
 
  form1.nyugtavolt = 5
  Nyugel1.Hide

End Sub

Private Sub Command11_Click()
  If Not xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Nem üres a nyugta!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  form1.nyugtavolt = 4
  Nyugel1.Hide

End Sub

Private Sub Command12_Click()
  If Not xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Nem üres számla!", 3, 0, "Figyelmeztetés", valasz%)
    Command2.SetFocus
    Exit Sub
  End If
  
  
  form1.nyugtavolt = 10
  Nyugel1.Hide

End Sub
Private Sub engedvissza()
  '--- nettóból illetve bruttóól adott engedmény visszaosztása
    
  busz@ = xval(Szenged.Text2.Text)
  fst$ = "######0.00"
  busz@ = Format(busz@, fst$)
  buszt$ = Str$(busz@)

  For i13% = 1 To 200
    elem$ = Trim(Nyugel1.MSFlexGrid1.TextMatrix(i13%, 1))
    If elem$ <> "" Then
       'bar$ = Nyugel1.MSFlexGrid1.TextMatrix(i13%, 5)
       'bar@ = xval(bar$)
       Nyugel1.MSFlexGrid1.TextMatrix(i13%, 5) = buszt$
    End If
  Next
End Sub


Private Sub Command13_Click()
   'Call mess("Az elõleg kezelés még nincs kész!", 3, 0, "Figyelmeztetés", valasz%)
   'Text1.SetFocus
   'Exit Sub
  
  If xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Üres számla!", 3, 0, "Figyelmeztetés", valasz%)
    Text1.SetFocus
    Exit Sub
  End If
  
    If ellhiba% = 0 Then
                    
            wujeloleg = 1
            
            Telobeszam.Show vbModal
            If welobeikt <> "" Then
              wqwel$ = torolvas("PELV", welobeikt, 1, 650)
              ' üressor keresés
              For q2q% = 1 To 200
                If Trim$(MSFlexGrid1.TextMatrix(q2q%, 1)) = "" Then
                  Exit For
                End If
              Next
              If Not q2q% = 200 Then
               q1q% = q2q%
               MSFlexGrid1.TextMatrix(q1q%, 1) = Trim(Mid$(wqwel$, 266, 15))
               MSFlexGrid1.TextMatrix(q1q%, 2) = Trim(Mid$(wqwel$, 315, 60))
               MSFlexGrid1.TextMatrix(q1q%, 7) = Mid$(wqwel$, 1, 7)
               tkod$ = Mid$(wqwel$, 266, 15)
               termrec$ = dbxkey("KTRM", tkod$)
               MSFlexGrid1.TextMatrix(q1q%, 6) = "Elõleg besz."
               
               If Not Trim$(termrec$) = "" Then

                 If welobemenny <> 0 Then
                   afakulcs@ = Val(Mid$(wqwel$, 246, 6))
                   nettoar@ = Val(Mid$(wqwel$, 281, 14))
                   bruttoar@ = nettoar@ * (1 + afakulcs@ / 100)
                   MSFlexGrid1.TextMatrix(q1q%, 3) = "-" + Trim(ertszam(Str(welobemenny), 14, 2))
                   'MSFlexGrid1.TextMatrix(q1q%, 7) = Trim(Mid$(wqwel$, 309, 6))
                   'MSFlexGrid1.TextMatrix(q1q%, 4) = Trim(Mid$(wqwel$, 281, 14))
                   MSFlexGrid1.TextMatrix(q1q%, 4) = Str(bruttoar@)
                  
               
                   tetnevmt(q1q%) = Mid$(wqwel$, 315, 120)
                   Mid$(tetnevmt(q1q%), 61, 60) = "Elõlegbeszámítás. Eredeti számla:" + Mid$(wqwel$, 8, 15)
                   MSFlexGrid1.TextMatrix(q1q%, 2) = "Elõlegbeszámítás. Eredeti számla:" + Mid$(wqwel$, 8, 15)
                   'MSFlexGrid1.TextMatrix(q1q%, 13) = "B"
               wbeszeloikt(q1q%) = welobeikt
               Call ujraszamol
               End If
               Else
                  Call mess("Nincs elõleg a termék törzsben!", 3, 0, "Hiba", valasz%)
               End If
              Else
               Call mess("Nincs üres sor a számlán!", 3, 0, "Hiba", valasz%)
              End If
            End If
     End If
                    

  
  
  
  'If ellhiba% = 0 Then
                    
  '   eloleg.Clear
  '   eloleg.AddItem "Beszámítható elõlegek"
  '   eloleg.AddItem "Név                            Számlaszám      Dátum  Felhasználható"
  '   f1$ = "--------------------------------------------------------------------"
  '   eloleg.AddItem f1$
  '   pkod$ = Text4(1).Text
  '   partrec$ = dbxkey("PART", pkod$)
  '   nxptr& = Val(Mid$(partrec$, 702, 10))
  '   elolegdb% = 0
  '   vaneloleg = False
  '   dxfi = FreeFile
  '   Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #dxfi
  '   fim& = LOF(dxfi)
  '   Do While nxptr& > 0
  '      Seek #dxfi, nxptr& + 9
  '      elrec$ = Space(650): Get #dxfi, , elrec$
  '      nxptr& = Val(Mid$(elrec$, 204, 10))
  '      elokod$ = "EV"
  '      If Mid$(elrec$, 90, 1) <> "S" And Mid$(elrec$, 224, 2) = elokod$ Then
  '         szabossz@ = xval(Mid$(elrec$, 44, 14)) + xval(Mid$(elrec$, 134, 14)) + xval(Mid$(elrec$, 148, 14))
  '         If szabossz@ > 0 Then
  '            ptariktato$ = Mid$(elrec$, 105, 7)
  '            ptarrec$ = dbxkey("PKTE", ptariktato$)
  '            elem1$ = Mid$(ptarrec$, 240, 30) + " " + Mid$(elrec$, 8, 15) + " " + Mid$(elrec$, 38, 6) + " " + Right$(Space$(14) + Format(szabossz@, "##########0.00"), 14) + " " + Mid(elrec$, 1, 7)
  '            elolegdb% = elolegdb% + 1
  '            eloleg.AddItem elem1$
  '            elto$(elolegdb%) = Mid$(elrec$, 1, 22) + Right$(Space$(14) + Format(szabossz@, "##########0.00"), 14)
  '         End If
  '      End If
  '   Loop
  '   Close dxfi
  '   If elolegdb% <> 0 Then
'elolegujra:
  '      eloleg.Visible = True
  '      Call tablazat("PSEL", 1, 5, nt$(), 3640, 7900, 4200, 3200)
  '      If rogzites% = 0 Then
  '         For i5% = 1 To 5: nt$(i5%) = Space$(43): Next
  '         rogzites% = 1
  '      Else
           '--- van elõleg
           
  '         For i5% = 1 To 5
  '           If Trim$(nt$(i5%)) <> "" Then
  '              etal% = 0
  '              For i6% = 1 To elolegdb%
  '                 If Mid$(nt$(i5%), 8, 15) = Mid$(elto$(i6%), 8, 15) Then
  '                    If (Mid$(nt$(i5%), 24, 14)) > Val(Mid$(elto$(i6%), 22, 14)) Then
  '                      Call mess("Túl sok az elõleg! / Maximum " + Trim$(Mid$(elto$(i6%), 22, 14)) + " lehet./", 3, 0, "Hiba", valasz%)
  '                      GoTo elolegujra
  '                    End If
  '                    etal% = i6%
  '                    Exit For
  '                 End If
  '              Next
  '              If etal% > 0 Then
  '                 Mid$(nt$(i5%), 1, 22) = elto$(etal%)
  '                 vaneloleg = True
  '              Else
  '                 Call mess("Nincs ilyen elõleg számla ! / " + Mid$(nt$(i5%), 8, 15) + "/", 3, 0, "Hiba", valasz%)
  '                 nt$(i5%) = Space$(43)
  '                 GoTo elolegujra
  '              End If
  '           End If
  '         Next
  '      End If
  '      eloleg.Visible = False
  '      Call ujraszamol
  '      If xval(Trim(Label1.Caption)) < 0 Then
  '       Call mess("Végösszeg nem lehet negatív!", 3, 0, "Hiba", valasz%)
  '       GoTo elolegujra
  '     End If
  '     End If
  'End If
  
                
  
  'form1.nyugtavolt = 11
  'Nyugel1.Hide

End Sub

Private Sub Command14_Click()
  form1.nyugtavolt = 0
  If Not xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Nem üres a nyugta!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  
  Call komin("AUWKER2", 1, "Kassza összesítõ", szoveg$, runtimhiba%)
  If runtimhiba% = 0 Then
   If kommegsem% = 1 Then
      Exit Sub
   End If
   ProgressBar1.Visible = True
   biziktdb& = 0
   fi1 = FreeFile
   Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #fi1
   fi2 = FreeFile
   Open auditorutvonal$ + "auw-kszt.ndx" For Binary Shared As #fi2
   fi3 = FreeFile
   Open listautvonal$ + terminal$ + task$ + "NP0.TMP" For Output As #fi3
   
   fi4 = FreeFile
   Open listautvonal$ + terminal$ + task$ + "NPS.TMP" For Output As #fi4
   
   
   rc& = Int(LOF(fi2) / 18)
   If rc& > 0 Then
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    '--- válogatás
    For il& = 1 To rc&
       Get #fi2, (il& - 1) * 18 + 1, rcim&
       Seek #fi1, rcim& + 9
       ksztrec$ = Space(170): Get #fi1, , ksztrec$
       DoEvents
       ProgressBar1.Value = pscale(il&, rc&)
       teljkelt$ = Mid$(ksztrec$, 130, 6)
       If Trim(komt(1).komtol) <= teljkelt$ And teljkelt$ <= Trim(komt(1).komig) Then
         tip$ = Mid$(ksztrec$, 126, 1)
         ' Eszi 2009.02.18 - készpénzes számla lista
         If (((komt(3).kommnv = 1 And (tip$ = "N" Or tip$ = "K")) Or komt(3).kommnv = 2) And Not komt(2).kommnv = 4) Or (komt(2).kommnv = 4 And ((komt(3).kommnv = 1 And tip$ = "K") Or (komt(3).kommnv = 2 And (tip$ = "K" Or tip$ = "H")))) Then
          blar@ = Val(Mid$(ksztrec$, 57, 12))
          bar@ = Val(Mid$(ksztrec$, 57, 12))
          enge@ = Val(Mid$(ksztrec$, 51, 6))
          bar@ = bar@ + bar@ * enge@ / 100
          blar@ = blar@ + blar@ * enge@ / 100
          ert@ = Val(Mid$(ksztrec$, 21, 12)) * blar@
          Mid$(ksztrec$, 83, 12) = Right$(Space$(12) + Format(ert@, "############0"), 12)
          Mid$(ksztrec$, 127, 1) = " "
          Mid$(ksztrec$, 57, 12) = Right$(Space$(12) + Format(bar@, "############0"), 12)
          Print #fi3, ksztrec$;
          
          ktetikt$ = Mid$(ksztrec$, 158, 7)
          kfttrec$ = dbxkey("KKFT", ktetikt$)
          If kfttrec$ <> "" Then
             bizikt$ = Mid(kfttrec$, 8, 7)
             
             nincs = True
             For j& = 1 To biziktdb&
               If bizikt$ = tbizikt$(j&) Then
                 nincs = False
                End If
             Next
             If nincs Then
               biziktdb& = biziktdb& + 1
               tbizikt$(biziktdb&) = bizikt$
             
               ksybrec$ = dbxkey("KSYB", bizikt$)
               For i9% = 1 To 4
                  elem$ = Mid$(ksybrec$, (i9% - 1) * 36 + 143, 36)
                  If Not Trim$(elem$) = "" And Mid$(elem$, 8, 4) = form1.elolegprefix$ Then
                    ' 11.02.23 - 10,10,18 winshop
                    wksztrec$ = ksztrec$
                    ert@ = Val(Mid$(elem$, 23, 14))
                    Mid$(wksztrec$, 36, 15) = "Elõleg"
                    Mid$(wksztrec$, 83, 12) = Right$(Space$(12) + Format(-ert@, "############0"), 12)
                    Mid$(wksztrec$, 126, 1) = "E"
                    Mid$(wksztrec$, 107, 15) = Mid$(elem$, 8, 10) + Space$(5)
                    Mid$(wksztrec$, 57, 12) = Right$(Space$(12) + Format(ert@, "############0"), 12)
                    'Mid$(wksztrec$, 122, 3) = Mid$(kszbrec$, 42, 3)
                  
                    Print #fi3, wksztrec$;
                  End If
               Next
             End If
          End If

         End If
       End If

       stornojel$ = Mid$(ksztrec$, 14, 1)
       If stornojel$ = "S" Then
             szlaszam$ = Mid$(ksztrec$, 1, 10)
             kszbrec$ = dbxkey("KSZB", szlaszam$)
             stkelt$ = Mid$(kszbrec$, 36, 6)
             szlkelt$ = Mid$(kszbrec$, 29, 6)
             stszamla$ = Mid$(kszbrec$, 282, 15)
             If Trim(komt(1).komtol) <= stkelt$ And stkelt$ <= Trim(komt(1).komig) And Trim(komt(1).komtol) <= szlkelt$ And szlkelt$ <= Trim(komt(1).komig) Then
               tip$ = Mid$(ksztrec$, 126, 1)
               ' ESzi 2009.02.18 - nem kell a sztornó
               If (((komt(3).kommnv = 1 And (tip$ = "N" Or tip$ = "K")) Or komt(3).kommnv = 2) And Not komt(2).kommnv = 4) Or (komt(2).kommnv = 4 And ((komt(3).kommnv = 1 And tip$ = "K") Or (komt(3).kommnv = 2 And (tip$ = "K" Or tip$ = "H")))) Then
                blar@ = Val(Mid$(ksztrec$, 57, 12))
                bar@ = Val(Mid$(ksztrec$, 57, 12))
                'enge@ = Val(Mid$(ksztrec$, 51, 6))
                'bar@ = bar@ + bar@ * enge@ / 100
                'blar@ = blar@ + blar@ * enge@ / 100
                ert@ = Val(Mid$(ksztrec$, 21, 12)) * blar@
                Mid$(ksztrec$, 83, 12) = Right$(Space$(12) + Format(-ert@, "############0"), 12)
                Mid$(ksztrec$, 127, 1) = "S"
                Mid$(ksztrec$, 130, 6) = stkelt$
                Mid$(ksztrec$, 57, 12) = Right$(Space$(12) + Format(bar@, "############0"), 12)
                Mid$(ksztrec$, 122, 3) = Mid$(kszbrec$, 42, 3)
                Mid$(ksztrec$, 1, 10) = stszamla$
                Print #fi3, ksztrec$;
                ' Elõleg beszámítás
                ktetikt$ = Mid$(ksztrec$, 158, 7)
                kfttrec$ = dbxkey("KKFT", ktetikt$)
                If kfttrec$ <> "" Then
                  bizikt$ = Mid(kfttrec$, 8, 7)
                  
                  nincs = True
                  For j& = 1 To biziktdb&
                    If bizikt$ = tbizikt$(j&) Then
                      nincs = False
                    End If
                  Next
                  If nincs Then
                    biziktdb& = biziktdb& + 1
                    tbizikt$(biziktdb&) = bizikt$

                  
                    ksybrec$ = dbxkey("KSYB", bizikt$)
                    For i9% = 1 To 4
                      elem$ = Mid$(ksybrec$, (i9% - 1) * 36 + 143, 36)
                      If Not Trim$(elem$) = "" Then
                        '  11.02.23 - 10,10,18 winshop
                        wksztrec$ = ksztrec$
                        ert@ = Val(Mid$(elem$, 23, 14))
                        Mid$(wksztrec$, 36, 15) = "Elõleg"
                        Mid$(wksztrec$, 83, 12) = Right$(Space$(12) + Format(ert@, "############0"), 12)
                        Mid$(wksztrec$, 126, 1) = "E"
                        Mid$(wksztrec$, 107, 15) = Mid$(elem$, 8, 10) + Space$(5)
                        Mid$(wksztrec$, 57, 12) = Right$(Space$(12) + Format(-ert@, "############0"), 12)
                        'Mid$(wksztrec$, 122, 3) = Mid$(kszbrec$, 42, 3)
                  
                        Print #fi3, wksztrec$;
                      End If
                    Next
                  End If
                End If
               End If
             End If
             ' Elõzõ napi sztornó
             'If Trim(komt(1).komtol) <= stkelt$ And stkelt$ <= Trim(komt(1).komig) And Not (Trim(komt(1).komtol) <= szlkelt$ And szlkelt$ <= Trim(komt(1).komig)) Then
             If ((Trim(komt(1).komtol) <= stkelt$ And stkelt$ <= Trim(komt(1).komig)) Or (Trim(komt(1).komtol) > stkelt$ And stkelt$ > Trim(komt(1).komig))) And Not (Trim(komt(1).komtol) <= szlkelt$ And szlkelt$ <= Trim(komt(1).komig)) Then
               tip$ = Mid$(ksztrec$, 126, 1)
               
               If (((komt(3).kommnv = 1 And (tip$ = "N" Or tip$ = "K")) Or komt(3).kommnv = 2) And Not komt(2).kommnv = 4) Then
                blar@ = Val(Mid$(ksztrec$, 57, 12))
                bar@ = Val(Mid$(ksztrec$, 57, 12))
                enge@ = Val(Mid$(ksztrec$, 51, 6))
                bar@ = bar@ + bar@ * enge@ / 100
                blar@ = blar@ + blar@ * enge@ / 100
                ert@ = Val(Mid$(ksztrec$, 21, 12)) * blar@
                Mid$(ksztrec$, 83, 12) = Right$(Space$(12) + Format(-ert@, "############0"), 12)
                Mid$(ksztrec$, 127, 1) = "S"
                Mid$(ksztrec$, 57, 12) = Right$(Space$(12) + Format(bar@, "############0"), 12)
                Mid$(ksztrec$, 122, 3) = Mid$(kszbrec$, 42, 3)
                Print #fi4, ksztrec$;
               End If
             End If
       
       End If
       
    Next
    Close fi3, fi2, fi1, fi4
    ProgressBar1.Visible = False
    
    Me.Hide
    ' rendezés
    If UCase(Dir$(listautvonal$ + terminal$ + task$ + "npt.tmp")) = UCase(terminal$ + task$ + "npt.tmp") Then
      Kill (listautvonal$ + terminal$ + task$ + "npt.tmp")
    End If
    If komt(2).kommnv = 4 Then
      rsor$ = "FIE=1,10/INP=" + listautvonal$ + terminal$ + task$ + "np0.tmp/OUT=" + listautvonal$ + terminal$ + task$ + "npt.tmp/RLE=170/MOD=A"
    Else
      rsor$ = "FIE=130,6,122,3,126,2,1,10/INP=" + listautvonal$ + terminal$ + task$ + "np0.tmp/OUT=" + listautvonal$ + terminal$ + task$ + "npt.tmp/RLE=170/MOD=A"
    End If
    Call rendez(rsor$)
    
    
    
    If rendezohiba% = 0 Then
       If komt(2).kommnv = 1 Or komt(2).kommnv = 2 Or komt(2).kommnv = 4 Then
        For i14% = 1 To 14: kulcstomb%(i14%, 1) = 0: kulcstomb%(i14%, 2) = 0: Next
       
        kulcstomb%(1, 1) = 130: kulcstomb%(1, 2) = 6
        kulcstomb%(2, 1) = 122: kulcstomb%(2, 2) = 3
        kulcstomb%(3, 1) = 126: kulcstomb%(3, 2) = 2
        kulcstomb%(4, 1) = 1: kulcstomb%(4, 2) = 10
        epoztomb%(1, 1) = 83: epoztomb%(1, 2) = 12
       
        If UCase(Dir$(listautvonal$ + terminal$ + task$ + "np1.tmp")) = UCase(terminal$ + task$ + "np1.tmp") Then
           Kill (listautvonal$ + terminal$ + task$ + "np1.tmp")
        End If
       
        Call osszevon("npt.tmp", "np1.tmp", 170, 0, renhiba%, 0, 0, 0, 0, 0, 0, 0, 0)
        form1.Label6.Visible = False
        form1.ProgressBar3.Visible = False
        form1.Refresh
        
        fi1 = FreeFile
        Open listautvonal$ + terminal$ + task$ + "NP1.TMP" For Binary Shared As #fi1
        fi2 = FreeFile
        Open listautvonal$ + terminal$ + task$ + "NP2.TMP" For Output Shared As #fi2

        rrc& = Int(LOF(fi1) / 170)
        
        If rrc& > 0 Then
           form1.ProgressBar3.Max = 100
           For i32& = 1 To rrc&
             DoEvents
             form1.ProgressBar3.Value = pscale(i32&, rrc&)
             Seek #fi1, (i32& - 1) * 170 + 1
             rec$ = Space(170): Get #fi1, , rec$
             ertek@ = Val(Mid$(rec$, 83, 12))
             tip$ = Mid$(rec$, 126, 1)
             If tip$ = "N" Or tip$ = "K" Then
               Call kerekit510(ertek@, kerekossz@, kerek@, "K")
               If ertek@ <> kerekossz@ Then
                  mezo$ = Right$(Space$(12) + Format(kerekossz@, "############"), 12)
                  Mid$(rec$, 83, 12) = mezo$
               End If
             End If
             Print #fi2, rec$;
           Next
        
        End If
    
        For i14% = 1 To 14: kulcstomb%(i14%, 1) = 0: kulcstomb%(i14%, 2) = 0: Next
        
        Close fi1, fi2
        
        If komt(2).kommnv = 4 Or komt(2).kommnv = 2 Then
         ' Eszi - nevet, címet kiírni - 158,7 =kkft iktató -> kkft 8,7  kkbz iktató, ksyb iktató
         fi1 = FreeFile
         Open listautvonal$ + terminal$ + task$ + "NP2.TMP" For Binary Shared As #fi1
         fi2 = FreeFile
         Open listautvonal$ + terminal$ + task$ + "NP3.TMP" For Output Shared As #fi2
         rrc& = Int(LOF(fi1) / 170)
        
         If rrc& > 1 Then
           form1.ProgressBar3.Max = 100
           For i32& = 1 To rrc&
             DoEvents
             form1.ProgressBar3.Value = pscale(i32&, rrc&)
             Seek #fi1, (i32& - 1) * 170 + 1
             rec$ = Space(170): Get #fi1, , rec$
             kkftikt$ = Mid$(rec$, 158, 7)
             kktfrec$ = dbxkey("KKFT", kkftikt$)
             kkBZikt$ = Mid$(kktfrec$, 8, 7)
             ksybrec$ = dbxkey("KSYB", kkBZikt$)
             nevcim$ = Space(120)
             If Not ksybrec$ = "" Then
               nevcim$ = Mid$(ksybrec$, 23, 120)
             End If
             rec$ = rec$ + nevcim$
             Print #fi2, rec$;
           Next
         End If
         Close fi1, fi2
         If komt(2).kommnv = 2 Then
           Call Nylistazo("auw-nps", listanev$, szoveg$, listhiba%)
         Else
            Call Nylistazo("auw-npk", listanev$, szoveg$, listhiba%)
         End If
         If listhiba% = 0 Then
          '--- lista megmutatása
           Shell programutvonal$ + "dbx4-sho.exe " + terminal$ + task$ + "NPK/" + listautvonal$, vbNormalFocus
           Call gomb(" Tovább &", gg%, 8320, 100, "V")
           Exit Sub
         End If

        Else
          kulcstomb%(1, 1) = 130: kulcstomb%(1, 2) = 6
          kulcstomb%(2, 1) = 122: kulcstomb%(2, 2) = 3
          kulcstomb%(3, 1) = 126: kulcstomb%(3, 2) = 2
          epoztomb%(1, 1) = 83: epoztomb%(1, 2) = 12
          
          If UCase(Dir$(listautvonal$ + terminal$ + task$ + "npo.tmp")) = UCase(terminal$ + task$ + "npo.tmp") Then
            Kill (listautvonal$ + terminal$ + task$ + "npo.tmp")
          End If
       
          Call osszevon("np2.tmp", "npo.tmp", 170, 0, renhiba%, 0, 0, 0, 0, 0, 0, 0, 0)
          form1.Label6.Visible = False
          form1.ProgressBar3.Visible = False
          form1.Refresh
        
        
        
          If renhiba% = 0 Then
            Call Nylistazo("auw-npo", listanev$, szoveg$, listhiba%)
            ' Elõzõ napi sztornók
            
            fi4 = FreeFile
            Open listautvonal$ + terminal$ + task$ + "NPS.TMP" For Binary As #fi4
            
            
            rrc& = Int(LOF(fi4) / 170)
        
            If rrc& > 0 Then
               
               fi5 = FreeFile
               Open listautvonal$ + terminal$ + task$ + "NPO.LST" For Append As #fi5
               
               Print #fi5, "TS"
               Print #fi5, "TS Elõzõ idõszaki /jövõbeli/, de most sztornózott tételek "
               Print #fi5, "TS============================================="
               Print #fi5, "TS"
               Print #fi5, "TSTelj   Kassza T        Bruttó"
               Print #fi5, "TSkelte         p         érték"
               Print #fi5, "TS............................."

               form1.ProgressBar3.Max = 100
               osszesen@ = 0
               For i32& = 1 To rrc&
                 DoEvents
                 form1.ProgressBar3.Value = pscale(i32&, rrc&)
                 Seek #fi1, (i32& - 1) * 170 + 1
                 rec$ = Space(170): Get #fi1, , rec$
                 osszesen@ = osszesen@ + Val(Mid$(rec$, 83, 12))
                 Print #fi5, "TS" + Mid$(rec$, 130, 6) + " " + Mid$(rec$, 122, 3) + "    " + Mid$(rec$, 126, 2) + " " + Mid$(rec$, 83, 12)
               Next
               Print #fi5, "TS............................."
               mezo$ = Right$(Space$(12) + Format(osszesen@, "############"), 12)
               Print #fi5, "TS        Összesen:" + mezo$
               Close fi5
            End If
            Close fi4
          End If
        End If
       Else
        ' tételes
        Call Nylistazo("auw-npt", listanev$, szoveg$, listhiba%)
       End If
       form1.ProgressBar3.Visible = False
       form1.ProgressBar2.Visible = False
       form1.ProgressBar1.Visible = False
       If listhiba% = 0 Then
          '--- lista megmutatása
          Shell programutvonal$ + "dbx4-sho.exe " + terminal$ + task$ + "npo/" + listautvonal$, vbNormalFocus
          Call gomb(" Tovább &", gg%, 8320, 100, "V")
       End If
       Me.Show
    End If
   
   End If
   
  End If
End Sub


Private Sub Command15_Click()
  form1.nyugtavolt = 0
  If Not xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Nem üres a nyugta!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  
  Call komin("AUWKER2", 2, "Napi pénztárgépes forgalom feladása", szoveg$, runtimhiba%)
  If runtimhiba% = 0 Then
   If kommegsem% = 1 Then
      Exit Sub
   End If
   biziktdb& = 0
   fil2 = FreeFile
   Open auditorutvonal$ + "auw-pvsz.ndx" For Binary Shared As #fil2
   fil1 = FreeFile
   Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #fil1
   rc& = Int(LOF(fil2) / 12)
   szladb& = 0
   If rc& > 0 Then
      For ia& = 1 To rc&
         Get #fil2, (ia& - 1) * 12 + 1, rcim&
         szrec$ = Space(1500)
         Get #fil1, rcim& + 9, szrec$
         If Not Mid$(szrec$, 166, 1) = "S" Then
           szladb& = szladb& + 1
           szlaszok$(szladb&) = Mid$(szrec$, 8, 15)
           If szladb& > 90000 Then
              Call mess("Több mint 90 000 számla! Hívja a fejlesztõt!", 3, 0, "Hiba", valasz%)
              Exit Sub
           End If
         End If
      Next
   End If
   Close fil1: Close fil2



   ProgressBar1.Visible = True
   fi1 = FreeFile
   Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #fi1
   fi2 = FreeFile
   Open auditorutvonal$ + "auw-kszt.ndx" For Binary Shared As #fi2
   fi3 = FreeFile
   Open listautvonal$ + terminal$ + task$ + "NP1.TMP" For Output As #fi3
   
   rc& = Int(LOF(fi2) / 18)
   If rc& > 0 Then
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    '--- válogatás
    For il& = 1 To rc&
       Get #fi2, (il& - 1) * 18 + 1, rcim&
       Seek #fi1, rcim& + 9
       ksztrec$ = Space(170): Get #fi1, , ksztrec$
       DoEvents
       ProgressBar1.Value = pscale(il&, rc&)
       teljkelt$ = Mid$(ksztrec$, 130, 6)
       If Trim(komt(1).komtol) <= teljkelt$ And teljkelt$ <= Trim(komt(1).komig) Then
         tip$ = Mid$(ksztrec$, 126, 1)
         If (tip$ = "N" Or tip$ = "K") Then
          blar@ = Val(Mid$(ksztrec$, 57, 12))
          enge@ = Val(Mid$(ksztrec$, 51, 6))
          blar@ = blar@ + blar@ * enge@ / 100
          Mid$(ksztrec$, 57, 12) = Right$(Space$(12) + Format(blar@, "############0"), 12)
          ert@ = Val(Mid$(ksztrec$, 21, 12)) * blar@
          Mid$(ksztrec$, 83, 12) = Right$(Space$(12) + Format(ert@, "############0"), 12)
          Mid$(ksztrec$, 127, 1) = " "
          Print #fi3, ksztrec$;
         End If
       End If
       ' A sztornót nem adja fel
       stornojel$ = Mid$(ksztrec$, 14, 1)
       If stornojel$ = "S" Then
             szlaszam$ = Mid$(ksztrec$, 1, 10)
             kszbrec$ = dbxkey("KSZB", szlaszam$)
             stkelt$ = Mid$(kszbrec$, 36, 6)
             szlkelt$ = Mid$(kszbrec$, 29, 6)
             If Trim(komt(1).komtol) <= stkelt$ And stkelt$ <= Trim(komt(1).komig) And Trim(komt(1).komtol) <= szlkelt$ And szlkelt$ <= Trim(komt(1).komig) Then
               tip$ = Mid$(ksztrec$, 126, 1)
               If (tip$ = "N" Or tip$ = "K") Then
                blar@ = Val(Mid$(ksztrec$, 57, 12))
                enge@ = Val(Mid$(psztrec$, 51, 6))
                blar@ = blar@ + blar@ * enge@ / 100
                ert@ = Val(Mid$(ksztrec$, 21, 12)) * blar@
                Mid$(ksztrec$, 83, 12) = Right$(Space$(12) + Format(-ert@, "############0"), 12)
                Mid$(ksztrec$, 127, 1) = "S"
                Mid$(ksztrec$, 130, 6) = stkelt$
                Print #fi3, ksztrec$;
               End If
             End If
       End If
       
    Next
    Close fi3, fi2, fi1
    Me.Hide
    ' rendezés
    rsor$ = "FIE=130,6,122,3,126,2,1,10/INP=" + listautvonal$ + terminal$ + task$ + "np1.tmp/OUT=" + listautvonal$ + terminal$ + task$ + "npt.tmp/RLE=170/MOD=A"
    Call rendez(rsor$)
    
    
    
    If rendezohiba% = 0 Then
        
        For i14% = 1 To 14: kulcstomb%(i14%, 1) = 0: kulcstomb%(i14%, 2) = 0: Next

        ' Számlánként felgyûjt
        kulcstomb%(1, 1) = 130: kulcstomb%(1, 2) = 6
        kulcstomb%(2, 1) = 122: kulcstomb%(2, 2) = 3
        kulcstomb%(3, 1) = 126: kulcstomb%(3, 2) = 2
        kulcstomb%(4, 1) = 1: kulcstomb%(4, 2) = 10
        epoztomb%(1, 1) = 83: epoztomb%(1, 2) = 12
       
        Call osszevon("npt.tmp", "np1.tmp", 170, 0, renhiba%, 0, 0, 0, 0, 0, 0, 0, 0)
        form1.Label6.Visible = False
        form1.ProgressBar3.Visible = False
        form1.Refresh
        
        ' 5 forintra kerekít
        fi1 = FreeFile
        Open listautvonal$ + terminal$ + task$ + "NP1.TMP" For Binary Shared As #fi1
        fi2 = FreeFile
        Open listautvonal$ + terminal$ + task$ + "NP2.TMP" For Output Shared As #fi2

        rrc& = Int(LOF(fi1) / 170)
        
        If rrc& > 1 Then
           form1.ProgressBar3.Max = 100
           For i32& = 1 To rrc&
             DoEvents
             form1.ProgressBar3.Value = pscale(i32&, rrc&)
             Seek #fi1, (i32& - 1) * 170 + 1
             rec$ = Space(170): Get #fi1, , rec$
             ertek@ = Val(Mid$(rec$, 83, 12))
             Call kerekit510(ertek@, kerekossz@, kerek@, "K")
             If ertek@ <> kerekossz@ Then
                mezo$ = Right$(Space$(12) + Format(kerekossz@, "############"), 12)
                Mid$(rec$, 83, 12) = mezo$
             End If
             Print #fi2, rec$;
           Next
        End If
    
       For i14% = 1 To 14: kulcstomb%(i14%, 1) = 0: kulcstomb%(i14%, 2) = 0: Next
        
       Close fi1, fi2
        
       kulcstomb%(1, 1) = 130: kulcstomb%(1, 2) = 6
       kulcstomb%(2, 1) = 122: kulcstomb%(2, 2) = 3
       epoztomb%(1, 1) = 83: epoztomb%(1, 2) = 12
       
       Call osszevon("np2.tmp", "npo.tmp", 170, 0, renhiba%, 0, 0, 0, 0, 0, 0, 0, 0)
       form1.Label6.Visible = False
       form1.ProgressBar3.Visible = False
       form1.Refresh
       
       If renhiba% = 0 Then
         Call listazo("auw-npz", listanev$, szoveg$, listhiba%)
       End If
       form1.ProgressBar1.Visible = False
       form1.ProgressBar2.Visible = False
       form1.ProgressBar3.Visible = False
       
       If listhiba% = 0 Then
          '--- lista megmutatása
          Shell programutvonal$ + "dbx4-sho.exe " + terminal$ + task$ + "npz/" + listautvonal$, vbNormalFocus
          Call gomb(" Feladás & Mégsem &", gg%, 8320, 100, "V")
          If gg% = 1 Then
            ' pvsz képzés - Ha már van nem adja fel
            '--- számla tartalmának könyvelése pvsz-be
            kkerek@ = 0
            rterminal$ = terminal$
            fi1 = FreeFile
            Open listautvonal$ + terminal$ + task$ + "NPT.TMP" For Binary As #fi1
            rrc& = Int(LOF(fi1) / 170)
            If rrc& > 0 Then
               Seek #fi1, 1
               ProgressBar1.Max = 100
               i32& = 1
               i34& = 1
               Seek #fi1, (i32& - 1) * 170 + 1
               hrec$ = Space(170): Get #fi1, , hrec$
               Do While i34& <= rrc&
                nap$ = Mid$(hrec$, 130, 6)
                Do While (i34& <= rrc&) And nap$ = Mid$(hrec$, 130, 6)
                 terminal$ = Mid$(hrec$, 122, 3)
                 
                 devizaertek@ = 0
                 forintertek@ = 0
                 For i99% = 1 To 50
                  If i99% < 6 Then
                       afakodok$(i99%) = "": afaalapok@(i99%) = 0: afaosszegek@(i99%) = 0: afaszamlak$(i99%) = ""
                  End If
                  ellszla$(i99%) = "": ellosszegek@(i99%) = 0: devellosszegek@(i99%) = 0
                 Next

                 Do While (i34& <= rrc&) And nap$ = Mid$(hrec$, 130, 6) And terminal$ = Mid$(hrec$, 122, 3)

               
                  For i33& = 1 To 1000: mt$(i33&) = "": Next
                  szlasz$ = Mid$(hrec$, 1, 10)
                  st$ = Mid$(hrec$, 127, 1)
                  j32& = 0
                  Do While (i34& <= rrc&) And szlasz$ = Mid$(hrec$, 1, 10) And nap$ = Mid$(hrec$, 130, 6) And terminal$ = Mid$(hrec$, 122, 3) And szlasz$ = Mid$(hrec$, 1, 10) And st$ = Mid$(hrec$, 127, 1)
                  
                   j32& = j32& + 1
                   mt$(j32&) = hrec$
                   If i34& < rrc& Then
                    Seek #fi1, (i34&) * 170 + 1
                    hrec$ = Space(170): Get #fi1, , hrec$
                    i34& = i34& + 1
                   Else
                     i34& = i34& + 1
                     Exit Do
                   End If
                   DoEvents
                   ProgressBar1.Value = pscale(i32&, rrc&)
                  
                  Loop
                  GoSub szamlaszamol2
                 Loop
                 GoSub folyokonyvel2
                Loop
               Loop
            End If
            Close fi1
            terminal$ = rterminal$
          End If
       End If
       Me.Show
    End If
   
   End If
   ProgressBar1.Visible = False
  End If
Exit Sub
szamlaszamol2:
  '--- számla kiszámítása áfa és kontírozás tombok feltoltese
  '--- devizakonverzió
  '--- elõleg iktatók beírása nt$-ba
  '--- feltölteni afakodok$(5), afaalapok@(5), afaosszegek@(5),afaszamlak$(5)
  '---            ellszla$(10), ellosszegek@(10), devellosszegek@(10)
  '--- 080229 kerekítés
  kpsszamla = 0
  If seset% = 1 Then
    '--- belföldi számla
    If Mid$(fejr$, 38, 3) = "   " Then
      fizm$ = Mid$(fejr$, 36, 2)
      fizmrec$ = dbxkey("PFIZ", fizm$)
      If Mid$(fizmrec$, 33, 1) = "K" Then kpsszamla = 1
    End If
  End If
  'devizaertek@ = 0
  'forintertek@ = 0
  ertker% = Val(Mid$(irec$, 344, 1))
  afaker% = Val(Mid$(irec$, 345, 1))
  If ertker% = 0 Then fste$ = "############0" Else fste$ = "#############0." + String(ertker%, "0")
  If afaker% = 0 Then fst$ = "############0" Else fst$ = "#############0." + String(afaker%, "0")
  If arf@ <> 0 Then fste$ = "#############0.00": fst$ = "#############0.00"
  onert@ = 0: obert@ = 0
  fforintertek@ = 0
  For i13% = 1 To 200
    elem$ = mt$(i13%)
    If Trim$(elem$) <> "" Then
      tkod$ = Mid$(elem$, 107, 15)
      termrec$ = dbxkey("KTRM", tkod$)
      afakod$ = Mid$(termrec$, 706, 2)
      afrec$ = dbxkey("PAFA", afakod$)
      afakulcs@ = xval(Mid$(afrec$, 33, 6))
      afajel$ = Mid$(afrec$, 39, 1)
      afszla$ = Mid$(afrec$, 49, 8)
      menny@ = xval(Mid$(elem$, 21, 12))
      liar@ = xval(Mid$(elem$, 57, 12))
      pensz@ = xval(Mid$(elem$, 51, 6))
'      penft@ = xval(Mid$(elem$, 43, 12))
      elar@ = xval(Mid$(elem$, 57, 12))
      ' kedvezmény kezelése
'      If pensz@ <> 0 Then
'          elar@ = elar@ / 100 * (100 - pensz@)
'      End If


      If menny@ <> 0 Then
        
        If Mid$(elem$, 127, 1) = "S" Then menny@ = -menny@
        
        bert@ = elar * menny@
        If arf@ <> 0 Then bert@ = bert@ * arf@
        bert@ = xval(Format(bert@, fste$))
      'Else
      '  bert@ = elar@
      '  If arf@ <> 0 Then bert@ = bert@ * arf@
      End If

      nert@ = bert@ / (1 + afakulcs@ / 100)
      nert@ = xval(Format(nert@, fste$))
      afaosz@ = bert@ - nert@
      afaosz@ = xval(Format(afaosz@, fst$))
          
      If menny@ <> 0 Then
         elar@ = nert@ / menny@
         elar@ = xval(Format(elar@, fste$))
      Else
         elar@ = nert@
      End If

      
      fforintertek@ = fforintertek@ + bert@
      If arf@ <> 0 Then devizaertek@ = devizaertek@ + bert@ / arf@


      For j99% = 1 To 5
        If afakodok$(j99%) = afakod$ Or afakodok$(j99%) = "" Then
          If afakodok$(j99%) = "" Then
            afakodok$(j99%) = afakod$
            afaalapok@(j99%) = nert@
            afaosszegek@(j99%) = afaosz@
            afaszamlak$(j99%) = afszla$
          Else
            afaalapok@(j99%) = afaalapok@(j99%) + nert@
            afaosszegek@(j99%) = afaosszegek@(j99%) + afaosz@
          End If
          Exit For
        End If
      Next
      krakrec$ = dbxkey("KRAK", krak$)
      'ellszi$ = Mid$(elem$, 69, 8) + Space$(8) + Mid$(elem$, 77, 16)
      'If Trim(ellszi$) = "" Then
        ellszi$ = Mid$(termrec$, 716, 8) + Space$(8) + Mid$(termrec$, 724, 16)
      'End If
      If Trim(ellszi$) = "" Then
        ellszi$ = Mid$(krakrec$, 151, 8) + Space$(8) + Mid$(krakrec$, 159, 16)
      End If
      For j99% = 1 To 50
        If ellszla$(j99%) = ellszi$ Or ellszla$(j99%) = "" Then
          If ellszla$(j99%) = "" Then
            ellszla$(j99%) = ellszi$
            ellosszegek@(j99%) = nert@
            If arf@ <> 0 Then devellosszegek@(j99%) = nert@ / arf@
          Else
            ellosszegek@(j99%) = ellosszegek@(j99%) + nert@
            If arf@ <> 0 Then devellosszegek@(j99%) = devellosszegek@(j99%) + nert@ / arf@
          End If
'          forintertek@ = forintertek@ + nert@
'          If arf@ <> 0 Then devizaertek@ = devizaertek@ + nert@ / arf@
          Exit For
        End If
      Next
    End If
  Next
  '--- afa kontirozasT  - CSAK egyszer kell megcsinálni!
  
'  For i99 = 1 To 5
'    If afakodok$(i99%) <> "" Then
'      ellszi$ = afaszamlak$(i99%) + Space$(24)
'      nert@ = afaosszegek@(i99%)
'      For j99% = 1 To 50
'        If ellszla$(j99%) = ellszi$ Or ellszla$(j99%) = "" Then
'          If ellszla$(j99%) = "" Then
'            ellszla$(j99%) = ellszi$
'            ellosszegek@(j99%) = nert@
'            If arf@ <> 0 Then devellosszegek@(j99%) = nert@ / arf@
'          Else
'            ellosszegek@(j99%) = ellosszegek@(j99%) + nert@
'            If arf@ <> 0 Then devellosszegek@(j99%) = devellosszegek@(j99%) + nert@ / arf@
'          End If
''          forintertek@ = forintertek@ + nert@
''          If arf@ <> 0 Then devizaertek@ = devizaertek@ + nert@ / arf@
'          Exit For
'        End If
'      Next
'    End If
'  Next
  '--- 080229 kerekítés
' If kpsszamla = 1 Then
    Call kerekit510(fforintertek@, kerekossz@, kerek@, "K")
    kkerek@ = kkerek@ + kerek@
    forintertek@ = forintertek@ + kerekossz@
'  End If
Return
folyokonyvel2:
  '--- számla tartalmának könyvelése pvsz-be
  szamlaszam$ = nap$ + " " + terminal$ + Space$(5)
  ' Van már ilyen számla?
  tal% = 0
  If szladb& > 0 Then
     For ia& = 1 To szladb&
         If szamlaszam$ = szlaszok$(ia&) Then tal% = 1: Exit For
     Next
  End If

  If tal% = 0 Then
  
  
  
   '--- afa kontirozasT  - CSAK egyszer kell megcsinálni!
  
  For i99 = 1 To 5
    If afakodok$(i99%) <> "" Then
      ellszi$ = afaszamlak$(i99%) + Space$(24)
      nert@ = afaosszegek@(i99%)
      For j99% = 1 To 50
        If ellszla$(j99%) = ellszi$ Or ellszla$(j99%) = "" Then
          If ellszla$(j99%) = "" Then
            ellszla$(j99%) = ellszi$
            ellosszegek@(j99%) = nert@
            If arf@ <> 0 Then devellosszegek@(j99%) = nert@ / arf@
          Else
            ellosszegek@(j99%) = ellosszegek@(j99%) + nert@
            If arf@ <> 0 Then devellosszegek@(j99%) = devellosszegek@(j99%) + nert@ / arf@
          End If
'          forintertek@ = forintertek@ + nert@
'          If arf@ <> 0 Then devizaertek@ = devizaertek@ + nert@ / arf@
          Exit For
        End If
      Next
   End If
  Next
      If kkerek@ <> 0 Then
      
      '--- áfa nem adóalap
      afakod$ = form1.kerafakod
      afrec$ = dbxkey("PAFA", afakod$)
      afakulcs@ = xval(Mid$(afrec$, 33, 6))
      afajel$ = Mid$(afrec$, 39, 1)
      afszla$ = Mid$(afrec$, 49, 8)
      For j99% = 1 To 5
        If afakodok$(j99%) = afakod$ Or afakodok$(j99%) = "" Then
          If afakodok$(j99%) = "" Then
            afakodok$(j99%) = afakod$
            afaalapok@(j99%) = kkerek@
            afaosszegek@(j99%) = 0
            afaszamlak$(j99%) = afszla$
          Else
            afaalapok@(j99%) = afaalapok@(j99%) + kkerek@
          End If
          Exit For
        End If
      Next
      '--- kerekítés kontírozása
      If kkerek@ > 0 Then
        ellszi$ = form1.kerbev
      Else
        ellszi$ = form1.kerraf
      End If
      For j99% = 1 To 50
        If ellszla$(j99%) = ellszi$ Or ellszla$(j99%) = "" Then
          If ellszla$(j99%) = "" Then
            ellszla$(j99%) = ellszi$
            ellosszegek@(j99%) = kkerek@
          Else
            ellosszegek@(j99%) = ellosszegek@(j99%) + kkerek@
          End If
          Exit For
        End If
      Next
    End If

  
  pvszrec$ = Space$(1500)
  Mid$(pvszrec$, 8, 10) = szamlaszam$
  Mid$(pvszrec$, 38, 15) = KASSAPARTN
  Mid$(pvszrec$, 211, 6) = nap$
  Mid$(pvszrec$, 58, 6) = nap$
  Mid$(pvszrec$, 64, 6) = nap$
  Mid$(pvszrec$, 70, 6) = nap$
  Mid$(pvszrec$, 76, 2) = KASSAFIZM
  'Mid$(pvszrec$, 201, 8) = Mid$(fejr$, 171, 8)
  'Mid$(pvszrec$, 1495, 4) = krak$
  'arf@ = xval(Mid$(fejr$, 41, 10))
  
  
  'If arf@ <> 0 Then
  '  Mid$(pvszrec$, 92, 3) = Mid$(fejr$, 38, 3)
  '  Mid$(pvszrec$, 95, 14) = ertszam(Str$(devizaertek@), 14, 2)
  'End If
  Mid$(pvszrec$, 78, 14) = ertszam(Str$(forintertek@), 14, 2)
  Mid$(pvszrec$, 109, 1) = "N"
  Mid$(pvszrec$, 110, 3) = "Napi bevét"
  Mid$(pvszrec$, 160, 6) = maidatum$
  Mid$(pvszrec$, 173, 8) = "Rendszer"
  '--- áfa bontás
  For i9% = 1 To 5
    If afaalapok@(i9%) <> 0 Then
      elem9$ = Space$(30)
      Mid$(elem9$, 1, 2) = afakodok$(i9%)
      Mid$(elem9$, 3, 14) = ertszam(Str$(afaalapok(i9%)), 14, 2)
      Mid$(elem9$, 17, 14) = ertszam(Str$(afaosszegek@(i9%)), 14, 2)
      Mid$(pvszrec$, (i9% - 1) * 30 + 250, 30) = elem9$
    End If
  Next
  '--- kontírozás
  kiegrec$ = ""
  For i9% = 1 To 50
    If ellosszegek@(i9%) <> 0 Then
      If i9% > 10 Then
        If kiegrec$ = "" Then
          kiegrec$ = Space(3000)
          Mid$(kiegrec$, 1, 7) = Mid$(pvszrec$, 1, 7)
          Mid$(kiegrec$, 8, 2) = "UJ"
        End If
        elem9$ = Space$(53)
        Mid$(elem9$, 1, 14) = ertszam(Str$(ellosszegek@(i9%)), 14, 2)
        Mid$(elem9$, 15, 32) = ellszla$(i9%)
        Mid$(kiegrec$, (i9% - 11) * 53 + 10, 53) = elem9$
      Else
        elem9$ = Space$(53)
        Mid$(elem9$, 1, 14) = ertszam(Str$(ellosszegek@(i9%)), 14, 2)
        Mid$(elem9$, 15, 32) = ellszla$(i9%)
        Mid$(pvszrec$, (i9% - 1) * 53 + 400, 53) = elem9$
      End If
    End If
  Next
  '-- elõleg beszámítások beírása
  'For i9% = 1 To 5
  '  If Mid$(nt$(i9%), 8, 15) <> Space$(15) Then
  '    Mid$(pvszrec$, (i9% - 1) * 43 + 1280, 43) = nt$(i9%)
  '  End If
  'Next
  '--- pvsz kiírás
  Call pszkonyvel(pvszrec$, kiegrec$, "V")
  vikt$ = Mid$(pvszrec$, 1, 7)
  Else
    Call mess("Nap:" + datki(nap$) + " Terminál:" + terminal$ + " már fel van adva.", 3, 0, "Hiba", valasz%)
  End If
Return

End Sub

Private Sub Command16_Click()
   Szenged.Text1.Visible = False
   Szenged.Label1.Visible = False
   Szenged.Label2.Visible = False
   
   Szenged.Text2.Text = "0"
   Szenged.Show vbModal
   If xval(Szenged.Text2.Text) Then
      Call engedvissza
      Call ujraszamol
   End If

End Sub

Private Sub Command17_Click()
  Dim ptar$(5), pfksz$(5), ossz@(5)
  
  For i1% = 1 To 5: ptar$(i1%) = "": pfksz$(i1%) = "": ossz@(i1%) = 0: Next
  If form1.ttkrak = "001 " Then
     ptar$(1) = "M01"
     pfksz$(1) = "3815    "
     ptar$(2) = "N01"
     pfksz$(2) = "3814    "
  ElseIf form1.ttkrak = "002 " Then
     ptar$(1) = "B01"
     pfksz$(1) = "3812    "
  ElseIf form1.ttkrak = "004 " Then
     ptar$(1) = "A01"
     pfksz$(1) = "38162   "
  End If
  
  form1.nyugtavolt = 0
  If Not xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Nem üres a nyugta!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  
  Call komin("AUWKER2", 3, "Napi összesítõ", szoveg$, runtimhiba%)
  If runtimhiba% = 0 Then
   If kommegsem% = 1 Then
      Exit Sub
   End If
   ProgressBar1.Visible = True
   biziktdb& = 0
   fi1 = FreeFile
   Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #fi1
   fi2 = FreeFile
   Open auditorutvonal$ + "auw-kszt.ndx" For Binary Shared As #fi2
   fi3 = FreeFile
   Open listautvonal$ + terminal$ + task$ + "NP0.TMP" For Output As #fi3
   
   rc& = Int(LOF(fi2) / 18)
   If rc& > 0 Then
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    '--- válogatás
    For il& = 1 To rc&
       Get #fi2, (il& - 1) * 18 + 1, rcim&
       Seek #fi1, rcim& + 9
       ksztrec$ = Space(170): Get #fi1, , ksztrec$
       DoEvents
       ProgressBar1.Value = pscale(il&, rc&)
       teljkelt$ = Mid$(ksztrec$, 130, 6)
       If Trim(komt(1).komtol) <= teljkelt$ And teljkelt$ <= Trim(komt(1).komig) Then
         tip$ = Mid$(ksztrec$, 126, 1)
         kell = True
         If tip$ = "N" Or tip$ = "K" Then
           Mid$(ksztrec$, 126, 1) = "1"
         ElseIf tip$ = "H" Then
           szlaszam$ = Mid$(ksztrec$, 1, 10)
           kszbrec$ = dbxkey("KSZB", szlaszam$)
           If Mid$(kszbrec$, 61, 4) = "BANK" Then
             Mid$(ksztrec$, 126, 1) = "2"
           Else
             Mid$(ksztrec$, 126, 1) = "3"
           End If
         ElseIf tip$ = "G" Then
            Mid$(ksztrec$, 126, 1) = "4"
            szla1$ = Mid$(ksztrec$, 1, 1)
            If szla1$ >= "0" And szla1$ <= "9" Then
               kell = False
            End If
         End If
          blar@ = Val(Mid$(ksztrec$, 57, 12))
          bar@ = Val(Mid$(ksztrec$, 57, 12))
          enge@ = Val(Mid$(ksztrec$, 51, 6))
          bar@ = bar@ + bar@ * enge@ / 100
          blar@ = blar@ + blar@ * enge@ / 100
          ert@ = Val(Mid$(ksztrec$, 21, 12)) * blar@
          Mid$(ksztrec$, 83, 12) = Right$(Space$(12) + Format(ert@, "############0"), 12)
          Mid$(ksztrec$, 127, 1) = " "
          Mid$(ksztrec$, 57, 12) = Right$(Space$(12) + Format(bar@, "############0"), 12)
          
          If kell Then
             Print #fi3, ksztrec$;
          End If
          
          ktetikt$ = Mid$(ksztrec$, 158, 7)
          kfttrec$ = dbxkey("KKFT", ktetikt$)
          If kfttrec$ <> "" Then
             bizikt$ = Mid(kfttrec$, 8, 7)
                          
             nincs = True
             For j& = 1 To biziktdb&
               If bizikt$ = tbizikt$(j&) Then
                 nincs = False
                End If
             Next
             If nincs Then
               biziktdb& = biziktdb& + 1
               tbizikt$(biziktdb&) = bizikt$

               ksybrec$ = dbxkey("KSYB", bizikt$)
               For i9% = 1 To 4
                  elem$ = Mid$(ksybrec$, (i9% - 1) * 36 + 143, 36)
                  '
                  If Not Trim$(elem$) = "" And Mid$(elem$, 8, 4) = form1.elolegprefix$ Then
                    ' Eszi - komment van a mezõben - elõlegnek nézi
                    ert@ = Val(Mid$(elem$, 23, 14))
                    Mid$(ksztrec$, 36, 15) = "Elõleg"
                    Mid$(ksztrec$, 83, 12) = Right$(Space$(12) + Format(-ert@, "############0"), 12)
                    'Mid$(ksztrec$, 126, 1) = "E" - már fel van töltve
                    Mid$(ksztrec$, 107, 15) = Mid$(elem$, 8, 10) + Space$(5)
                    Mid$(ksztrec$, 57, 12) = Right$(Space$(12) + Format(ert@, "############0"), 12)
                    'Mid$(ksztrec$, 122, 3) = Mid$(kszbrec$, 42, 3)
                  
                    Print #fi3, ksztrec$;
                  End If
               Next
             End If
          End If
         
       End If
       stornojel$ = Mid$(ksztrec$, 14, 1)
       If stornojel$ = "S" Then
             szlaszam$ = Mid$(ksztrec$, 1, 10)
             kszbrec$ = dbxkey("KSZB", szlaszam$)
             stkelt$ = Mid$(kszbrec$, 36, 6)
             If Trim(komt(1).komtol) <= stkelt$ And stkelt$ <= Trim(komt(1).komig) Then
               tip$ = Mid$(ksztrec$, 126, 1)
               kell = True
               If tip$ = "N" Or tip$ = "K" Then
                  Mid$(ksztrec$, 126, 1) = "1"
               ElseIf tip$ = "H" Then
                  szlaszam$ = Mid$(ksztrec$, 1, 10)
                  kszbrec$ = dbxkey("KSZB", szlaszam$)
                  If Mid$(kszbrec$, 61, 4) = "BANK" Then
                     Mid$(ksztrec$, 126, 1) = "2"
                  Else
                     Mid$(ksztrec$, 126, 1) = "3"
                  End If
                ElseIf tip$ = "G" Then
                   Mid$(ksztrec$, 126, 1) = "4"
                   ' Eszi
                   szla1$ = Mid$(ksztrec$, 1, 1)
                   If szla1$ >= "0" And szla1$ <= "9" Then
                      kell = False
                   End If

                End If
               
                blar@ = Val(Mid$(ksztrec$, 57, 12))
                bar@ = Val(Mid$(ksztrec$, 57, 12))
                enge@ = Val(Mid$(ksztrec$, 51, 6))
                bar@ = bar@ + bar@ * enge@ / 100
                blar@ = blar@ + blar@ * enge@ / 100
                ert@ = Val(Mid$(ksztrec$, 21, 12)) * blar@
                Mid$(ksztrec$, 83, 12) = Right$(Space$(12) + Format(-ert@, "############0"), 12)
                Mid$(ksztrec$, 127, 1) = "S"
                Mid$(ksztrec$, 130, 6) = stkelt$
                Mid$(ksztrec$, 57, 12) = Right$(Space$(12) + Format(bar@, "############0"), 12)
                Mid$(ksztrec$, 122, 3) = Mid$(kszbrec$, 42, 3)
                If kell Then
                  Print #fi3, ksztrec$;
                End If
             End If
       End If
       
    Next
    Close fi3, fi2, fi1
    ProgressBar1.Visible = False
    
    Me.Hide
    ' rendezés
    If UCase(Dir$(listautvonal$ + terminal$ + task$ + "npt.tmp")) = UCase(terminal$ + task$ + "npt.tmp") Then
      Kill (listautvonal$ + terminal$ + task$ + "npt.tmp")
    End If
    rsor$ = "FIE=126,1,122,3,127,1,1,10/INP=" + listautvonal$ + terminal$ + task$ + "np0.tmp/OUT=" + listautvonal$ + terminal$ + task$ + "npt.tmp/RLE=170/MOD=A"

    Call rendez(rsor$)
    
    
    
    If rendezohiba% = 0 Then
       
        For i14% = 1 To 14: kulcstomb%(i14%, 1) = 0: kulcstomb%(i14%, 2) = 0: Next
       
        kulcstomb%(1, 1) = 126: kulcstomb%(1, 2) = 1
        kulcstomb%(2, 1) = 122: kulcstomb%(2, 2) = 3
        kulcstomb%(3, 1) = 127: kulcstomb%(3, 2) = 1
        kulcstomb%(4, 1) = 1: kulcstomb%(4, 2) = 10
        
        epoztomb%(1, 1) = 83: epoztomb%(1, 2) = 12
       
        If UCase(Dir$(listautvonal$ + terminal$ + task$ + "np1.tmp")) = UCase(terminal$ + task$ + "np1.tmp") Then
           Kill (listautvonal$ + terminal$ + task$ + "np1.tmp")
        End If
       
        Call osszevon("npt.tmp", "np1.tmp", 170, 0, renhiba%, 0, 0, 0, 0, 0, 0, 0, 0)
        form1.Label6.Visible = False
        form1.ProgressBar3.Visible = False
        form1.Refresh
        
        fi1 = FreeFile
        Open listautvonal$ + terminal$ + task$ + "NP1.TMP" For Binary Shared As #fi1
        fi2 = FreeFile
        Open listautvonal$ + terminal$ + task$ + "NP2.TMP" For Output Shared As #fi2

        rrc& = Int(LOF(fi1) / 170)
        
        If rrc& > 1 Then
           form1.ProgressBar3.Max = 100
           For i32& = 1 To rrc&
             DoEvents
             form1.ProgressBar3.Value = pscale(i32&, rrc&)
             Seek #fi1, (i32& - 1) * 170 + 1
             rec$ = Space(170): Get #fi1, , rec$
             ertek@ = Val(Mid$(rec$, 83, 12))
             tip$ = Mid$(rec$, 126, 1)
             If tip$ = "1" Then
               Call kerekit510(ertek@, kerekossz@, kerek@, "K")
               If ertek@ <> kerekossz@ Then
                  mezo$ = Right$(Space$(12) + Format(kerekossz@, "############"), 12)
                  Mid$(rec$, 83, 12) = mezo$
               End If
             End If
             Print #fi2, rec$;
           Next
        End If
    
        For i14% = 1 To 14: kulcstomb%(i14%, 1) = 0: kulcstomb%(i14%, 2) = 0: Next
        
        Close fi1, fi2
        
        
          kulcstomb%(1, 1) = 126: kulcstomb%(1, 2) = 1
          kulcstomb%(2, 1) = 122: kulcstomb%(2, 2) = 3
          kulcstomb%(3, 1) = 127: kulcstomb%(3, 2) = 1
          epoztomb%(1, 1) = 83: epoztomb%(1, 2) = 12
        
          If UCase(Dir$(listautvonal$ + terminal$ + task$ + "npo.tmp")) = UCase(terminal$ + task$ + "npo.tmp") Then
            Kill (listautvonal$ + terminal$ + task$ + "npo.tmp")
          End If
       
          Call osszevon("np2.tmp", "npo.tmp", 170, 0, renhiba%, 0, 0, 0, 0, 0, 0, 0, 0)
          form1.Label6.Visible = False
          form1.ProgressBar3.Visible = False
          form1.Refresh
        
        
        
          If renhiba% = 0 Then
            'Call Nylistazo("auw-npb", listanev$, szoveg$, listhiba%)
             lfi = FreeFile
             Open listautvonal$ + terminal$ + task$ + "npb.lst" For Output As #lfi
                         
             sr$ = "CM" + "Napi összesítõ " + Trim(cegneve$)
             Print #lfi, sr$
             sr$ = "FL" + String(60, "=")
             Print #lfi, sr$
             sr$ = "FL" + "Készült: " + datki(maidatum$) + "   " + ugyintezo$
             Print #lfi, sr$
             sr$ = "FL" + "Idõszak: " + datki(Trim(komt(1).komtol)) + "  -  " + datki(Trim(komt(1).komig))
             Print #lfi, sr$
             Print #lfi, "TS" + String(60, ".")
             
             Print #lfi, "TS" + String(60, " ")
             Print #lfi, "TS" + String(60, " ")
             
             sr$ = "TS" + "Pénztárgép"
             Print #lfi, sr$
                
             sr$ = "TS" + String(60, "=")
             Print #lfi, sr$

             sr$ = "12345678901 12345678901234  12345678901234  12345678901234"
             sr$ = "TS" + "Ptg.             Készpénz          Sztornó     Összesen"
             Print #lfi, sr$
             Print #lfi, "TS" + String(60, ".")

             fi1 = FreeFile
             Open listautvonal$ + terminal$ + task$ + "NPO.TMP" For Binary Shared As #fi1
 
             rrc& = Int(LOF(fi1) / 170)
             
             sr$ = ""
             ertek1@ = 0
        
             If rrc& > 1 Then
               form1.ProgressBar3.Max = 100
               rtip$ = ""
               rptg$ = ""
               Seek #fi1, 1
               rec$ = Space(170): Get #fi1, , rec$
               i32& = 1
               mossz1@ = 0
               mossz2@ = 0
               Do While i32& <= rrc&
                 tip$ = Mid$(rec$, 126, 1)
                 rtip$ = tip$
                 sr$ = "TS"
                 If tip$ = "2" Then
                     sr$ = sr$ + "Bankkártya:  "
                 ElseIf tip$ = "3" Then
                     sr$ = sr$ + "Átutalás:    "
                 ElseIf tip$ = "4" Then
                    Exit Do
                 End If
                 tossz1@ = 0
                 tossz2@ = 0
                 wertek@ = 0
                 Do While rtip$ = tip$ And i32& <= rrc&
                    ptg$ = Mid$(rec$, 122, 3)
                    rptg$ = ptg$
                    If tip$ = "1" Then
                       sr$ = "TS"
                       sr$ = sr$ + ptg$ + "          "
                       wertek@ = 0
                    End If
                    
                    Do While rptg$ = ptg$ And rtip$ = tip$ And i32& <= rrc&
                       stip$ = Mid$(rec$, 127, 1)
                       rstip$ = stip$
                       
                       Do While rstip$ = stip$ And rptg$ = ptg$ And rtip$ = tip$ And i32& <= rrc&
                         ertek@ = Val(Mid$(rec$, 83, 12))
                         If rtip$ = "1" Then
                            sr$ = sr$ + ertszam(Str$(ertek@), 14, 0)
                         End If
                         If rstip$ = " " Then
                           tossz1 = tossz1 + ertek@
                         Else
                           tossz2 = tossz2 + ertek@
                         End If
                         wertek@ = wertek@ + ertek@
                         
                         DoEvents
                         form1.ProgressBar3.Value = pscale(i32&, rrc&)
                         i32& = i32& + 1
                         Seek #fi1, (i32& - 1) * 170 + 1
                         rec$ = Space(170): Get #fi1, , rec$
                         tip$ = Mid$(rec$, 126, 1)
                         ptg$ = Mid$(rec$, 122, 3)
                         stip$ = Mid$(rec$, 127, 1)
                      Loop
                      If Not (rptg$ = ptg$ And rtip$ = tip$ And i32& <= rrc&) And rtip$ = "1" Then
                         ' nincs sztornó
                          If rstip$ = " " Then
                             sr$ = sr$ + ertszam(Str$(0@), 14, 0) + ertszam(Str$(wertek@), 14, 0)
                          Else
                             sr$ = sr$ + ertszam(Str$(wertek@), 14, 0)
                          End If
                      End If
                   Loop
                   If rtip = "1" Then
                      Print #lfi, sr$
                      For i1% = 1 To 5:
                        If ptar$(i1%) = rptg$ Then
                           ossz@(i1%) = wertek@
                        End If
                      Next
                   End If
                 Loop
                 mossz1@ = mossz1@ + tossz1@
                 mossz2@ = mossz2@ + tossz2@
                 If rtip = "1" Then
                      Print #lfi, "TS" + String(60, "-")
                      sr$ = "TSKp. összesen:" + ertszam(Str$(tossz1@), 14, 0) + ertszam(Str$(tossz2@), 14, 0) + ertszam(Str$(tossz1@ + tossz2@), 14, 0)
                      Print #lfi, sr$
                      Print #lfi, "TS" + String(60, "=")
                 End If
                 
                 If rtip = "2" Or rtip = "3" Then
                      sr$ = sr$ + ertszam(Str$(tossz1@), 14, 0) + ertszam(Str$(tossz2@), 14, 0) + ertszam(Str$(tossz1@ + tossz2@), 14, 0)
                      Print #lfi, sr$
                 End If
                 
               Loop
               Print #lfi, "TS" + String(60, "=")
               sr$ = "TSMind össz.:  " + ertszam(Str$(mossz1@), 14, 0) + ertszam(Str$(mossz2@), 14, 0) + ertszam(Str$(mossz1@ + mossz2@), 14, 0)
               Print #lfi, sr$
               
             End If
             Close fi1
             
             ' pénztár
             
              For i2% = 1 To 5
                If Not ptar$(i2%) = "" Then
                   utpzar$ = csokdat(Trim(komt(1).komtol))
                   pzardat$ = Trim(komt(1).komig)
                   ptfok$ = pfksz$(i2%)
                   
                   fi1 = FreeFile
                   Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #fi1
                   fi2 = FreeFile
                   Open auditorutvonal$ + "auw-pkte.ndx" For Binary Shared As #fi2
                   rc& = Int(LOF(fi2) / 12)
                
                   elso% = 1
                   nyito@ = 0: zaro@ = 0
                   For iq& = 1 To rc&
                     Get #fi2, (iq& - 1) * 12 + 1, rcim&
                     Seek #fi1, rcim& + 9
                     trec$ = Space(270): Get #fi1, , trec$
                     If Mid$(trec$, 192, 1) <> "S" Then
                       fkk$ = Mid$(trec$, 22, 8)
                       If fkk$ = ptfok$ Then
                         dat$ = Mid$(trec$, 16, 6)
                         If dat$ <= utpzar$ And Trim(utpzar$) <> "" Then
                           '--- nyitó egyenlegbe
                           If Mid$(trec$, 8, 1) = "B" Then
                              nyito@ = nyito@ + xval(Mid$(trec$, 56, 14))
                           Else
                              nyito@ = nyito@ - xval(Mid$(trec$, 56, 14))
                           End If
                         End If
                       End If
                     End If
                   Next
               
                Print #lfi, "TS" + String(60, " ")
                Print #lfi, "TS" + String(60, " ")
                Print #lfi, "TS" + String(60, " ")
                
                sr$ = "TS" + "Telepi pénztár " + ptar$(i2%) + " " + ptfok$
                Print #lfi, sr$
                sr$ = "TS" + String(60, "=")
                Print #lfi, sr$
                
                sr$ = "TSPtg J        Összeg Megjegyzés"
                Print #lfi, sr$
                
                sr$ = "TS" + String(60, "-")
                Print #lfi, sr$
                
                sr$ = "TS" + ptar$(i2%) + " N" + ertszam(Str$(nyito@), 14, 0) + " Nyitó"
                Print #lfi, sr$
                
                sr$ = "TS" + String(60, "-")
                Print #lfi, sr$
                
                
                zaro@ = nyito@
                lefol@ = 0
                For iq& = 1 To rc&
                  Get #fi2, (iq& - 1) * 12 + 1, rcim&
                  Seek #fi1, rcim& + 9
                  trec$ = Space(270): Get #fi1, , trec$
                  If Mid$(trec$, 192, 1) <> "S" Then
                    fkk$ = Mid$(trec$, 22, 8)
                    If fkk$ = ptfok$ Then
                      dat$ = Mid$(trec$, 16, 6)
                      If dat$ <= utpzar$ And Trim(utpzar$) <> "" Then
                        '--- nyitó egyenlegbe
                      Else
                        If dat$ > utpzar$ And dat$ <= pzardat$ Then
                          '--- napi tétel
                          '--- tétel írása
                          sr$ = "TS" + ptar$(i2%) + " " + Mid$(trec$, 8, 1) + ertszam(Mid$(trec$, 56, 14), 14, 0) + " " + Mid$(trec$, 30, 25)
                          Print #lfi, sr$
                          If Mid$(trec$, 8, 1) = "B" Then
                            zaro@ = zaro@ + xval(Mid$(trec$, 56, 14))
                            If UCase(Trim$(Mid$(trec$, 30, 25))) = UCase("Kassza lefölözés") Then
                               lefol@ = lefol@ + Val(Mid$(trec$, 56, 14))
                            End If
                          Else
                            zaro@ = zaro@ - xval(Mid$(trec$, 56, 14))
                          End If
                        End If
                      End If
                    End If
                  End If
                Next
                '--- láblécek nyomtatása
                sr$ = "TS" + String(60, "=")
                Print #lfi, sr$
                sr$ = "TS" + ptar$(i2%) + " Z" + ertszam(Str$(zaro@), 14, 0) + " Záró"
                Print #lfi, sr$
                If nyito@ + ossz@(i2%) - lefol@ = zaro@ Then
                   Print #lfi, "TSOK - Nyitó /" + Str(nyito@) + "/ + kp. forgalom /" + Str(ossz@(i2%)) + "/ - lefölözés " + Str(lefol@) + " = Záró /" + Str(zaro@)
                Else
                   Print #lfi, "TSHibás - Nyitó /" + Str(nyito@) + "/ + kp. forgalom /" + Str(ossz@(i2%)) + "/ - lefölözés " + Str(lefol@) + " = Záró /" + Str(zaro@)
                   
                End If

             
             
                End If
              Next
             
              Close fi1, fi2
             ' sztornó
             fi2 = FreeFile
             Open listautvonal$ + terminal$ + task$ + "NP2.TMP" For Binary Shared As #fi2
             rrc& = Int(LOF(fi2) / 170)
             
        
             If rrc& > 1 Then
               Print #lfi, "TS" + String(60, " ")
               Print #lfi, "TS" + String(60, " ")
               Print #lfi, "TS" + String(60, " ")
               
               sr$ = "TS" + "Sztornó számlák:"
               Print #lfi, sr$
                
               sr$ = "TS" + String(60, "=")
               Print #lfi, sr$

             
               sr$ = "TS" + "Ptg.Számlaszám        Összeg"
               Print #lfi, sr$
               Print #lfi, "TS" + String(60, ".")


               form1.ProgressBar3.Max = 100
               Seek #fi1, 1
               rec$ = Space(170): Get #fi1, , rec$
               i32& = 1
               mossz1@ = 0
               Do While i32& <= rrc&
                 tip$ = Mid$(rec$, 126, 1)
                 tossz1@ = 0
                 Do While tip$ = Mid$(rec$, 126, 1) And i32& <= rrc&
                   stip$ = Mid$(rec$, 127, 1)
                   If stip$ = "S" Then
                     tip$ = Mid$(rec$, 126, 1)
                     If tip$ <> "4" Then
                       ertek@ = Val(Mid$(rec$, 83, 12))
                       tossz1@ = tossz1@ + ertek@
                       sr$ = "TS" + Mid$(rec$, 122, 3) + " " + Mid$(rec$, 1, 10) + ertszam(Str$(ertek@), 14, 2)
                       Print #lfi, sr$
                     End If
                   End If
                   i32& = i32& + 1
                   Seek #fi1, (i32& - 1) * 170 + 1
                   rec$ = Space(170): Get #fi1, , rec$
                 Loop
                 If tip$ = "1" Then
                   sr$ = "TSKp-és:        "
                 ElseIf tip$ = "2" Or tip$ = "3" Then
                   sr$ = "Hiteles:        "
                 Else
                   Exit Do
                 End If
                 mossz1@ = mossz1@ + tossz1@
                 If Not tossz1@ = 0 Then
                   Print #lfi, "TS" + String(60, "-")
                   sr$ = sr$ + ertszam(Str$(tossz1@), 14, 2)
                   Print #lfi, sr$
                 End If
               Loop
               Print #lfi, "TS" + String(60, "=")
               sr$ = "TSÖsszesen:       " + ertszam(Str$(mossz1@), 14, 2)
               Print #lfi, sr$
            End If
          End If
          Close fi2
       
          fi2 = FreeFile
          Open listautvonal$ + terminal$ + task$ + "NP2.TMP" For Binary Shared As #fi2
          rrc& = Int(LOF(fi2) / 170)
             
        
          If rrc& > 1 Then
               Print #lfi, "TS" + String(60, " ")
               Print #lfi, "TS" + String(60, " ")
               Print #lfi, "TS" + String(60, " ")
               
               sr$ = "TS" + "Garancia jegyek:"
               Print #lfi, sr$
                
               sr$ = "TS" + String(60, "=")
               Print #lfi, sr$

             
               sr$ = "TS" + "Ptg.Garjegy szám      Összeg Számlaszám"
               Print #lfi, sr$
               Print #lfi, "TS" + String(60, ".")


               form1.ProgressBar3.Max = 100
               Seek #fi1, 1
               rec$ = Space(170): Get #fi1, , rec$
               i32& = 1
               mossz1@ = 0
               Do While i32& <= rrc&
                 tip$ = Mid$(rec$, 126, 1)
                 stip$ = Mid$(rec$, 127, 1)
                 G$ = Mid$(rec$, 3, 1)
                 If tip$ = "4" And G$ = "G" Then
                     tip$ = Mid$(rec$, 126, 1)
                     ertek@ = Val(Mid$(rec$, 83, 12))
                     mossz1@ = mossz1@ + ertek@
                     sr$ = "TS" + Mid$(rec$, 122, 3) + " " + Mid$(rec$, 1, 10) + ertszam(Str$(ertek@), 14, 0) + " ........................."
                     Print #lfi, sr$
                  End If
                  i32& = i32& + 1
                  Seek #fi1, (i32& - 1) * 170 + 1
                  rec$ = Space(170): Get #fi1, , rec$
               Loop
               Print #lfi, "TS" + String(60, "=")
               sr$ = "TSÖsszesen:       " + ertszam(Str$(mossz1@), 14, 0)
               Print #lfi, sr$
            
          End If
          Close fi2
          
          
         Print #lfi, "TS" + String(60, " ")
         Print #lfi, "TS" + String(60, " ")
         Print #lfi, "TS" + String(60, " ")
               
         sr$ = "TS" + "'F1-es' készpénzes számlák"
         Print #lfi, sr$
                
         sr$ = "TS" + String(60, "-")
         Print #lfi, sr$
          
         sr$ = "TS" + "Számlaszám Sz.kelte           Összeg"
         Print #lfi, sr$
         Print #lfi, "TS" + String(60, ".")

         
         dbfi = FreeFile
         Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #dbfi
         ndfi = FreeFile
         Open auditorutvonal$ + "auw-pszb.ndx" For Binary Shared As #ndfi
         rc& = Int(LOF(ndfi) / 15)
         mossz1 = 0
         For i1d& = 1 To rc&
           Get #ndfi, (i1d& - 1) * 15 + 1, rcim&
           trec$ = Space(300): Get #dbfi, rcim& + 9, trec$
           storno$ = Mid$(trec$, 35, 1)
           If Not storno$ = "S" Then
             stkelt$ = Mid$(trec$, 78, 6)
             fizmd$ = Mid$(trec$, 92, 2)
             If Trim(komt(1).komtol) <= stkelt$ And stkelt$ <= Trim(komt(1).komig) And fizmd$ = "01" Then
                szlasz$ = Mid$(trec$, 1, 10)
                vikt$ = Mid$(trec$, 53, 7)
                pvszrec$ = dbxkey("PVSZ", vikt$)
                bossz@ = Val(Mid$(pvszrec$, 78, 14))
                mossz1 = mossz1 + bossz@
                sr$ = "TS" + szlasz$ + " " + datki(stkelt$) + " " + ertszam(Str$(bossz@), 14, 0)
                Print #lfi, sr$

             End If
           End If
         Next
         Print #lfi, "TS" + String(60, " ")
         sr$ = "TS" + String(60, "=")
         Print #lfi, sr$
         
         sr$ = "TS Öszesen:" + String(13, " ") + ertszam(Str$(mossz1@), 14, 0)
         Print #lfi, sr$
         sr$ = "TS" + String(60, "=")
         Print #lfi, sr$

         Close dbfi, ndfi
          
          
          
         If komt(2).kommnv = 2 Then
         
         Print #lfi, "TS" + String(60, " ")
         Print #lfi, "TS" + String(60, " ")
         Print #lfi, "TS" + String(60, " ")
               
         sr$ = "TS" + "Bevételezések"
         Print #lfi, sr$
                
         sr$ = "TS" + String(99, "-")
         Print #lfi, sr$

         If komt(3).kommnv = 2 Then
          sr$ = "TS" + String(99, "-")
          Print #lfi, sr$
          
          sr$ = "TS" + "Termékkód       N é v                                       Mennyiség Nettó besz.ár  Bruttó érték"
          Print #lfi, sr$
          Print #lfi, "TS" + String(99, ".")
         End If

         
         dbfi = FreeFile
         Open auditorutvonal$ + "auwker.dbx" For Binary Shared As #dbfi
         ndfi = FreeFile
         Open auditorutvonal$ + "auw-kkft.ndx" For Binary Shared As #ndfi
         rc& = Int(LOF(ndfi) / 12)
         mossz1 = 0
         For i1d& = 1 To rc&
           Get #ndfi, (i1d& - 1) * 12 + 1, rcim&
           trec$ = Space(130): Get #dbfi, rcim& + 9, trec$
           storno$ = Mid$(trec$, 95, 1)
           If Not storno$ = "S" Then
             stkelt$ = Mid$(trec$, 15, 6)
             If Trim(komt(1).komtol) <= stkelt$ And stkelt$ <= Trim(komt(1).komig) Then
                bizikt$ = Mid$(trec$, 8, 7)
                bizrec$ = dbxkey("KKBZ", bizikt$)
                fajta$ = Mid$(bizrec$, 54, 1)
                If fajta$ = "B" Then
                  termek$ = Mid$(trec$, 36, 15)
                  termrec$ = dbxkey("KTRM", termek$)
                  afakod$ = Mid$(termrec$, 706, 2)
                  afarec$ = dbxkey("PAFA", afakod$)
                  afa@ = xval(Mid$(afarec$, 33, 6))
                  beszar@ = xval(Mid$(trec$, 59, 12))
                  menny@ = xval(Mid$(trec$, 71, 12))
                  bert@ = (menny@ * beszar@) * (1 + afa@ / 100)
                  mossz1 = mossz1 + bert@
                  If komt(3).kommnv = 2 Then
                    sr$ = "TS" + termek$ + " " + Mid$(termrec$, 16, 40) + " " + ertszam(Str$(menny@), 12, 0) + ertszam(Str$(beszar@), 14, 0) + ertszam(Str$(bert@), 14, 0)
                    Print #lfi, sr$
                  End If
                End If
             End If
           End If
         Next
         Print #lfi, "TS" + String(60, " ")
         sr$ = "TS" + String(99, "=")
         Print #lfi, sr$
         
         sr$ = "TS Beszerzési érték öszesen:" + Space(60) + ertszam(Str$(mossz1@), 14, 0)
         Print #lfi, sr$

         Close dbfi, ndfi
         End If
         
       Close lfi
       form1.ProgressBar3.Visible = False
       form1.ProgressBar2.Visible = False
       form1.ProgressBar1.Visible = False
       If listhiba% = 0 Then
          '--- lista megmutatása
          Shell programutvonal$ + "dbx4-sho.exe " + terminal$ + task$ + "npb/" + listautvonal$, vbNormalFocus
          Call gomb(" Tovább &", gg%, 8320, 100, "V")
       End If
       Me.Show
    End If
   
   End If

  End If

End Sub

Private Sub Command18_Click()
     Dim nevcim$(15000)
     ndb = 0
     dbfi = FreeFile
     Open auditorutvonal$ + "auwker2.dbx" For Binary Shared As #dbfi
     ndfi = FreeFile
     Open auditorutvonal$ + "AUW-KPAR.ndx" For Binary Shared As #ndfi
     rc& = Int(LOF(ndfi) / 12)
     For i1% = 1 To rc&
         Get #ndfi, (i1% - 1) * 12& + 1, rcim&
         Seek #dbfi, rcim& + 9
         kparrec$ = Space(200): Get #dbfi, , kparrec$
         ndb = ndb + 1
         nevcim$(ndb) = Mid$(kparrec$, 1, 200)
     Next
     Close dbfi, ndfi
     
     dbfi = FreeFile
     Open auditorutvonal$ + "auwker2.dbx" For Binary Shared As #dbfi
     ndfi = FreeFile
     Open auditorutvonal$ + "AUW-KSYB.ndx" For Binary Shared As #ndfi
     rc& = Int(LOF(ndfi) / 12)
     For i1% = 1 To rc&
         Get #ndfi, (i1% - 1) * 12& + 1, rcim&
         Seek #dbfi, rcim& + 9
         ksybrec$ = Space(200): Get #dbfi, , ksybrec$
         If (Trim$(Mid$(ksybrec$, 8, 15)) = "" Or Mid$(ksybrec$, 8, 6) = "BANK  ") And Not Trim$(Mid$(ksybrec$, 23, 60)) = "" Then
          van = False
          For j1% = 1 To ndb
            If Mid$(nevcim$(j1%), 1, 60) = Mid$(ksybrec$, 23, 60) Then
              van = True
              Exit For
            End If
          Next
          If van Then
            'ndb = ndb + 1
            'nevcim$(ndb) = Mid$(ksybrec$, 23, 120)
            
            azonosito$ = Mid$(nevcim$(j1%), 190, 7)
            kparrec$ = Space$(120)
            kparrec$ = dbxkey("KPAR", azonosito$)
            If Mid$(kparrec$, 1, 60) = Mid$(kparrec$, 61, 60) Then
               Mid$(kparrec$, 61, 60) = Mid$(ksybrec$, 83, 60)
               Call dbxki("KPAR", kparrec$, ";", " ", " ", hiba%)
             End If
          End If
         End If

     Next
     Close dbfi, ndfi
End Sub

Private Sub Command19_Click()

If CheckTámop.Value = 1 Then
  AzonositoBeker.Show vbModal
  
Else
  TBrutto.Show vbModal
' Támop számla képzés
form1.nyugtavolt = 12
teljesossz@ = Val(TBrutto.Text1.Text)

If Text4(1).Text = "" Then
    Call mess("Partner kódot tölse ki! /T..../", 3, 0, "Hiba", valasz%)
    Exit Sub
Else
  If Mid$(Text4(1).Text, 1, 1) <> "T" Then
    Call mess("A partnerkód T-vel kezdõdik!", 3, 0, "Hiba", valasz%)
    Exit Sub
  Else
    partkod$ = Left$(Text4(1).Text + Space$(15), 15)
    partrec$ = dbxkey("PART", partkod$)
    ' Már van számlája
    
    If partrec$ <> "" Then
       scim& = xval(Mid$(partrec$, 712, 10))
            If scim& > 0 Then
              fixa = FreeFile
              Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #fixa
              Do While scim& > 0
                psr$ = Space(1500)
                Get #fixa, scim& + 9, psr$
                If Mid$(psr$, 166, 1) <> "S" Then
                  
                    Call mess("Ennek a partnernek már van számlája", 5, 3, "Figyelmeztetés", valasz%)
                    If valasz% = 0 Then
                    
                      Exit Do
                      
                  
                    End If
                  
                End If
                scim& = xval(Mid(psr$, 191, 10))
              Loop
              Close fixa
            End If
    End If
    
    
    
    
    
    kod1$ = Trim$(Mid$(partrec$, 363, 60))
    kod2$ = Trim$(Mid$(partrec$, 423, 60))
    pos% = InStr(kod2, "Bev")
    If pos% > 0 Then
      kod2$ = Mid$(kod2, 1, pos% - 1)
    End If
    If kod1$ = "" Or kod2$ = "" Or Trim$(Text12.Text) = "" Then
      If Trim$(Text12.Text) = "" Then
         Call mess("Adószám nincs kitöltve!", 3, 0, "Hiba", valasz%)
      Else
         Call mess("A tanfolyam adatai nincsenek kitöltve!", 3, 0, "Hiba", valasz%)
      End If
      Exit Sub
    Else
      Text8.Text = kod1$
      Text11.Text = kod2$
      Text7.Text = "02"
      Text5.Text = "Átutalás"
      fizhatido% = 60
      fidat$ = maidatum$
      For i13% = 1 To fizhatido%
         xxxx$ = novdat(fidat$)
         fidat$ = xxxx$
      Next
      Text6.Text = fidat$

      
      
      Text1.Text = "TAMOPDIJ"
      MSFlexGrid1.TextMatrix(1, 1) = "TAMOPDIJ"
      MSFlexGrid1.TextMatrix(2, 1) = "TAMOPONRESZ"
      ' Elõleg
      pkod$ = Text4(1).Text
      partrec$ = dbxkey("PART", pkod$)
      nxptr& = Val(Mid$(partrec$, 702, 10))
      elolegdb% = 0
      vaneloleg = False
      dxfi = FreeFile
      Open auditorutvonal$ + "auwszamv.dbx" For Binary Shared As #dxfi
      fim& = LOF(dxfi)
      Do While nxptr& > 0
        Seek #dxfi, nxptr& + 9
        elrec$ = Space(650): Get #dxfi, , elrec$
        nxptr& = Val(Mid$(elrec$, 204, 10))
        elokod$ = "EV"
        If Mid$(elrec$, 90, 1) <> "S" And Mid$(elrec$, 224, 2) = elokod$ Then
           szabossz@ = xval(Mid$(elrec$, 44, 14)) + xval(Mid$(elrec$, 134, 14)) + xval(Mid$(elrec$, 148, 14))
           If szabossz@ > 0 Then
              ptariktato$ = Mid$(elrec$, 105, 7)
              ptarrec$ = dbxkey("PKTE", ptariktato$)
              elem1$ = Mid$(ptarrec$, 240, 30) + " " + Mid$(elrec$, 8, 15) + " " + Mid$(elrec$, 38, 6) + " " + Right$(Space$(14) + Format(szabossz@, "##########0.00"), 14) + " " + Mid(elrec$, 1, 7)
              elolegdb% = elolegdb% + 1
              eloleg.AddItem elem1$
              elto$(elolegdb%) = Mid$(elrec$, 1, 22) + Right$(Space$(14) + Format(szabossz@, "##########0.00"), 14)
              Exit Do
           End If
        End If
      Loop
      Close dxfi
      MSFlexGrid1.TextMatrix(2, 4) = Right$(Space$(14) + Format(szabossz@, "##########0.00"), 14)
      MSFlexGrid1.TextMatrix(1, 4) = Right$(Space$(14) + Format(teljesossz@ - szabossz@, "##########0.00"), 14)
      
      For i5% = 1 To 5: nt$(i5%) = Space$(43): Next
      ' Több elõleg
      Mid$(nt$(1), 1, 22) = elto$(1)
      Mid$(nt$(1), 24, 14) = Right$(Space$(14) + Format(szabossz@, "##########0.00"), 14)
      
      ktrmkod$ = Left(Trim(MSFlexGrid1.TextMatrix(1, 1)) + Space$(15), 15)
      ktrmrec$ = dbxkey("KTRM", ktrmkod$)
      
      MSFlexGrid1.TextMatrix(1, 2) = Trim(Mid$(ktrmrec$, 16, 60))
      If MSFlexGrid1.TextMatrix(1, 3) = "" Then MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "1"
      MSFlexGrid1.TextMatrix(1, 5) = ""
      MSFlexGrid1.TextMatrix(1, 6) = "Szolgáltatás"
      MSFlexGrid1.TextMatrix(1, 7) = Trim(Mid$(ktrmrec$, 484, 6))
                
      ktrmkod$ = Left(Trim(MSFlexGrid1.TextMatrix(2, 1)) + Space$(15), 15)
      ktrmrec$ = dbxkey("KTRM", ktrmkod$)
      
      MSFlexGrid1.TextMatrix(2, 2) = Trim(Mid$(ktrmrec$, 16, 60))
      If MSFlexGrid1.TextMatrix(2, 3) = "" Then MSFlexGrid1.TextMatrix(2, 3) = "1"
      MSFlexGrid1.TextMatrix(2, 5) = ""
      MSFlexGrid1.TextMatrix(2, 6) = "Szolgáltatás"
      MSFlexGrid1.TextMatrix(2, 7) = Trim(Mid$(ktrmrec$, 484, 6))
      tk1o@ = 0
      For j9% = 1 To 2
        tk1o@ = tk1o@ + Val(MSFlexGrid1.TextMatrix(j9%, 4))
      Next
  
      For j9% = 1 To 5
         tk1o@ = tk1o@ - Val(Mid$(nt$(j9%), 24, 14))
      Next
      
      
      
      Label1.Caption = ertszam(Str(tk1o@), 12, 2)
 
      
      
      
    End If
  End If
End If
End If
End Sub

Private Sub Command2_Click()
  '--- mégsem
  form1.nyugtavolt = 0
  form1.szallitobol = False
  Nyugel1.Hide
  'form1.Command3.SetFocus
     
End Sub

Private Sub Command21_Click()
If Not RTrim$(Text4(1).Text) = "" And Not RTrim$(Text4(1).Text) = "BANK" Then
  partkod$ = Text4(1).Text
  partkod$ = Left$(Text4(1).Text + Space$(15), 15)
  partrec$ = dbxkey("PART", partkod$)
  ttop& = 1420
  lleft& = 100
  sszel& = 0
  mmag& = 4720
  kxcbrec$ = Space$(140)
  Mid$(kxcbrec$, 1, 8) = Mid$(partrec$, 106, 8)
  ' Irszám
  Mid$(kxcbrec$, 10, 30) = Mid$(partrec$, 114, 30)
  ' Település
  Mid$(kxcbrec$, 40, 10) = Mid$(partrec$, 543, 10)
  ' Kerület
  Mid$(kxcbrec$, 50, 30) = Mid$(partrec$, 144, 30)
  ' Közterület (utca, tér stb)
  Mid$(kxcbrec$, 80, 10) = Mid$(partrec$, 553, 10)
  ' Közterület jellege
  Mid$(kxcbrec$, 90, 10) = Mid$(partrec$, 174, 10)
  ' Házszám vagy helyrajzi szám
  Mid$(kxcbrec$, 100, 10) = Mid$(partrec$, 563, 10)
  ' Épület
  Mid$(kxcbrec$, 110, 10) = Mid$(partrec$, 573, 10)
  ' Lépcsõház
  Mid$(kxcbrec$, 120, 10) = Mid$(partrec$, 583, 10)
  ' Szint
  Mid$(kxcbrec$, 130, 10) = Mid$(partrec$, 593, 10)
  ' Ajtó
vissza:
  Call vektabl("KXCB", 1, kxcbrec$, ttop&, lleft&, sszel&, mmag&)
  If rogzites% <> 0 Then
   If RTrim$(Mid$(kxcbrec$, 80, 10)) = "" Then
       Call mess("Bontsa fel a címet! 'Közterület jellege'", 3, 0, "Hiba", valasz%)
       GoTo vissza
    End If

     Mid$(partrec$, 106, 8) = Mid$(kxcbrec$, 1, 8)
  ' Irszám
     Mid$(partrec$, 114, 30) = Mid$(kxcbrec$, 10, 30)
  ' Település
     Mid$(partrec$, 543, 10) = Mid$(kxcbrec$, 40, 10)
  ' Kerület
     Mid$(partrec$, 144, 30) = Mid$(kxcbrec$, 50, 30)
  ' Közterület (utca, tér stb)
     Mid$(partrec$, 553, 10) = Mid$(kxcbrec$, 80, 10)
  ' Közterület jellege
     Mid$(partrec$, 174, 10) = Mid$(kxcbrec$, 90, 10)
  ' Házszám vagy helyrajzi szám
    Mid$(partrec$, 563, 10) = Mid$(kxcbrec$, 100, 10)
  ' Épület
    Mid$(partrec$, 573, 10) = Mid$(kxcbrec$, 110, 10)
  ' Lépcsõház
    Mid$(partrec$, 583, 10) = Mid$(kxcbrec$, 120, 10)
  ' Szint
    Mid$(partrec$, 593, 10) = Mid$(kxcbrec$, 130, 10)
    Call dbxki("PART", partrec$, ";", "", "", hiba%)
    
    Text3.Text = postacim(partrec$, 106)
    
    
  End If
End If
End Sub

Private Sub Command22_Click()
If RTrim$(Text4(1).Text) = "" Or RTrim$(Text4(1).Text) = "BANK" Then
If Not RTrim$(Text13.Text) = "" Then
  
  If Not Text13.Text = "Új" Then
    partkod$ = Left$(Text13.Text + Space$(7), 7)
    cimbrec$ = dbxkey("KCIM", partkod$)
  Else
    cimbrec$ = ""
    partkod$ = ""
  End If
  If cimbrec$ = "" Then
     If Not partkod$ = "" Then
       'partrec$ = dbxkey("KPAR", partkod$)
       kpar$ = dbxkey("KPAR", partkod$)
       kxcbrec$ = Space$(140)
       Mid$(kxcbrec$, 1, 8) = Mid$(kpar$, 61, 4)
       ' Irszám
       Mid$(kxcbrec$, 10, 30) = Mid$(kpar$, 65, 55)
       ' Település
     Else
       kxcbrec$ = Space$(140)
     End If
  Else
     kxcbrec$ = cimbrec$
  End If
  ttop& = 1500
  lleft& = 100
  sszel& = 0
  mmag& = 4720
vissza:
  Call vektabl("KXCB", 1, kxcbrec$, ttop&, lleft&, sszel&, mmag&)
  If RTrim$(Mid$(kxcbrec$, 40, 100)) = "" Then
     Call mess("Bontsa fel a címet! 'Csak az irányítószam és a település van kitöltve'", 3, 0, "Hiba", valasz%)
     GoTo vissza
  End If
  
  If rogzites% <> 0 Then
    
    
    Text3.Text = RTrim$(Mid$(kxcbrec$, 1, 8))
    If Not RTrim$(Mid$(kxcbrec$, 10, 30)) = "" Then
       Text3.Text = Text3.Text + " " + RTrim$(Mid$(kxcbrec$, 10, 30))
    End If
    If Not RTrim$(Mid$(kxcbrec$, 40, 10)) = "" Then
        szep = ","
       If Right(RTrim$(Mid$(kxcbrec$, 40, 10)), 1) = "," Then
          szep = ""
       End If
       Text3.Text = Text3.Text + " " + RTrim$(Mid$(kxcbrec$, 40, 10)) + szep
    Else
      If Not Right(RTrim$(Text3.Text), 1) = "," Then
         Text3.Text = Text3.Text + ","
      End If
    End If
    If Not RTrim$(Mid$(kxcbrec$, 50, 30)) = "" Then
       Text3.Text = Text3.Text + " " + RTrim$(Mid$(kxcbrec$, 50, 30))
    End If
    If Not RTrim$(Mid$(kxcbrec$, 80, 10)) = "" Then
       Text3.Text = Text3.Text + " " + RTrim$(Mid$(kxcbrec$, 80, 10))
    End If
    If Not RTrim$(Mid$(kxcbrec$, 90, 10)) = "" Then
       szep = "."
       If Right(RTrim$(Mid$(kxcbrec$, 90, 10)), 1) = "." Then
          szep = ""
       End If
       Text3.Text = Text3.Text + " " + RTrim$(Mid$(kxcbrec$, 90, 10)) + szep
    End If
    If Not RTrim$(Mid$(kxcbrec$, 100, 10)) = "" Then
       szep = ".ép"
       If Right(RTrim$(Mid$(kxcbrec$, 100, 10)), 3) = ".ép" Then
          szep = ""
       End If
       Text3.Text = Text3.Text + " " + RTrim$(Mid$(kxcbrec$, 100, 10)) + szep
    End If
    If Not RTrim$(Mid$(kxcbrec$, 110, 10)) = "" Then
       szep = ".lh"
       If Right(RTrim$(Mid$(kxcbrec$, 110, 10)), 3) = ".lh" Then
          szep = ""
       End If
       Text3.Text = Text3.Text + " " + RTrim$(Mid$(kxcbrec$, 110, 10)) + szep
    End If
    If Not RTrim$(Mid$(kxcbrec$, 120, 10)) = "" Then
       szep = ".em"
       If Right(RTrim$(Mid$(kxcbrec$, 120, 10)), 3) = ".em" Or UCase(RTrim$(Mid$(kxcbrec$, 120, 10))) = "FSZ" Or UCase(RTrim$(Mid$(kxcbrec$, 120, 10))) = "FSZ." Or UCase(RTrim$(Mid$(kxcbrec$, 120, 10))) = "FÖLDSZINT" Then
          szep = ""
       End If
       Text3.Text = Text3.Text + " " + RTrim$(Mid$(kxcbrec$, 120, 10)) + szep
    End If
    If Not RTrim$(Mid$(kxcbrec$, 130, 10)) = "" Then
       szep = ".ajtó"
       If Right(RTrim$(Mid$(kxcbrec$, 130, 10)), 5) = ".ajtó" Then
          szep = ""
       End If
       If Right(RTrim$(Mid$(kxcbrec$, 130, 10)), 1) = "." Then
          szep = "ajtó"
       End If
      
       Text3.Text = Text3.Text + " " + RTrim$(Mid$(kxcbrec$, 130, 10)) + szep
    End If
    
    If cimbrec$ = "" Then
      If partkod$ = "" Then
         kpar$ = Space(200)
         Mid$(kpar$, 1, 60) = Nyugel1.Text2
         Mid$(kpar$, 61, 60) = Nyugel1.Text3
         Call dbxki("KPAR", kpar$, ";", "U", "G", hiba%)
         partkod$ = Mid$(kpar$, 190, 7)
         If Text13.Text = "Új" Then
           Text13.Text = partkod$
         End If
      End If
      ncv$ = Mid$(Nyugel1.Text2 + Space(60), 1, 60) + Mid$(Nyugel1.Text3 + Space(60), 1, 60)
      Call form1.nevcimtolt(ncv$, Mid$(kpar$, 190, 7))
      kxcbrec$ = kxcbrec$ + Space$(10)
      Mid$(kxcbrec$, 140, 7) = partkod$
      Call dbxki("KCIM", kxcbrec$, ";", "U", "", hiba%)

    Else
      Call dbxki("KCIM", kxcbrec$, ";", "", "", hiba%)
      If Not partkod$ = "" Then
         kpar$ = dbxkey("KPAR", partkod$)
         If RTrim$(Mid$(kpar$, 1, 60)) = RTrim$(Nyugel1.Text2) Then
               Mid$(kpar$, 61, 60) = Space$(60)
               Mid$(kpar$, 61, 60) = Nyugel1.Text3
               Call dbxki("KPAR", kpar$, ";", " ", " ", hiba%)
               ncv$ = Mid$(Nyugel1.Text2 + Space(60), 1, 60) + Mid$(Nyugel1.Text3 + Space(60), 1, 60)
               Call form1.nevcimmod$(ncv$, partkod$)
         End If
      End If
    End If

    
  End If
End If
End If
End Sub

Private Sub Command3_Click()
  '--- Áfás számla, kisker, nagyker, átutalás
  If vegosszege(1) < 0 Then
    Call mess("Végösszeg nem lehet negatív!", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If
  If xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Üres számla!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  
  fizetnem@ = Val(Label1.Caption)
  Call kerekit510(fizetnem@, fizetni@, kerek@, "K")
  Label1.Caption = ertszam(Str(fizetni@), 12, 2)

  
  If Trim(Text2.Text) = "" Or Trim(Text3.Text) = "" Then
    Call mess("Áfás számla esetén a vevõ neve és címe kötelezõ", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If
  If Not (Trim(Text5.Text) = "Készpénz" Or Trim(Text5.Text) = "") Then
    Call mess("Csak készpénz fizetési mód lehet", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If
    ' Gyáriszám ellenõrzés
  For i1% = 1 To 200
    termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(i1%, 1)) + Space(15), 15)
    If Not Trim(termkod$) = "" Then
        ktrmrec$ = dbxkey("KTRM", termkod$)
        If Mid$(ktrmrec$, 849, 1) = "I" Then
           db% = 0
           For j1% = 1 To 200
             If Not Trim(gyariszamok$(i1%, j1%)) = "" Then
                db% = db% + 1
             End If
           Next
           If Not db% = MSFlexGrid1.TextMatrix(i1%, 3) Then
              Call mess(Str(i1%) + ".sorban nem " + MSFlexGrid1.TextMatrix(i1%, 3) + " gyári szám van rögzítve!", 3, 0, "Figyelmeztetés", valasz%)
              Exit Sub
           End If
        End If
    End If
  Next
  
  If Not Keszletellenorzes Then
     Exit Sub
  End If


  ' Eszi - Vissza járó számítás
  Visszajaro.Label5 = "A végösszeget üsse be a pénztárgépbe!"
  Visszajaro.Label6 = "Az Áfás számla 1.példánya a vevõé."
  Visszajaro.Label7 = "A 2.példányt a pénztárgépes nyugtával együtt irattározni kell!"
  Visszajaro.Label2(1) = Right$(Space$(12) + Format(Val(Nyugel1.Label1), "# ### ### ##0"), 12)
  
  Visszajaro.Show vbModal
  
'  Call mess("A végösszeget üsse be a pénztárgépbe!" + Chr$(13) + "Az Áfás számla 1.példánya a vevõé." + Chr$(13) + "A 2.példányt a pénztárgépes nyugtával együtt irattározni kell!", 4, 0, "Figylmeztetés", valasz%)
  form1.nyugtavolt = 2
  Nyugel1.Hide
End Sub

Private Sub Command4_Click()
  If MSFlexGrid1.Row < 200 Then
    For i77% = MSFlexGrid1.Row To 199
      For i78% = 1 To 7
        MSFlexGrid1.TextMatrix(i77%, i78%) = MSFlexGrid1.TextMatrix(i77% + 1, i78%)
      Next
      
      For i78% = 1 To 200
        gyariszamok$(i77%, i78%) = gyariszamok$(i77% + 1, i78%)
        hivatkozas$(i77%, i78%) = hivatkozas$(i77% + 1, i78%)
      Next

    Next
  End If
  For i78% = 1 To 7
    MSFlexGrid1.TextMatrix(200, i78%) = ""
  Next
  For i78% = 1 To 200
    gyariszamok$(200, i78%) = ""
    hivatkozas$(200, i78%) = ""
  Next

  Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
  Text1.SelStart = Len(Trim(Text1.Text))
  Text1.SetFocus
  Call ujraszamol
End Sub

Private Sub Command5_Click()
  '--- Áfás számla,  nagyker, átutalás
  If vegosszege(1) < 0 Then
    Call mess("Végösszeg nem lehet negatív!", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If
  If xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Üres számla!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  If Trim(Text2.Text) = "" Or Trim(Text3.Text) = "" Then
    Call mess("Áfás számla esetén a vevõ neve és címe kötelezõ", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If
  If Trim(Text5.Text) = "Készpénz" Then
    Call mess("Készpénz fizatési mód nem lehet", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If
  If (Trim(Text4(1).Text) = "BANK" Or Trim(Text4(1).Text) = "") And Trim(Text5.Text) = "Átutalás" Then
    Call mess("BANK vagy üres partnerkódnak nem lehet átutalás a fizetési mód", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If
  If Not (Trim(Text4(1).Text) = "BANK" Or Trim(Text4(1).Text) = "" Or Trim(Text4(1).Text) = "KP") And UCase(Trim(Text5.Text)) = "BANKKÁRTYA" Then
    Call mess("Hibás partnerkód átutalásos fizetési módnál", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If

  ' Gyáriszám ellenõrzés
  For i1% = 1 To 200
    termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(i1%, 1)) + Space(15), 15)
    If Not Trim(termkod$) = "" Then
        ktrmrec$ = dbxkey("KTRM", termkod$)
        If Mid$(ktrmrec$, 849, 1) = "I" Then
           db% = 0
           For j1% = 1 To 200
             If Not Trim(gyariszamok$(i1%, j1%)) = "" Then
                db% = db% + 1
             End If
           Next
           If Not db% = MSFlexGrid1.TextMatrix(i1%, 3) Then
              Call mess(Str(i1%) + ".sorban nem " + MSFlexGrid1.TextMatrix(i1%, 3) + " gyári szám van rögzítve!", 3, 0, "Figyelmeztetés", valasz%)
              Exit Sub
           End If
        End If
    End If
  Next
  
  If Not Keszletellenorzes Then
     Exit Sub
  End If
  
  If Not form1.nyugtavolt = 12 Then form1.nyugtavolt = 3
  Nyugel1.Hide

End Sub

Private Sub Command6_Click()
  If Not xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Nem üres a nyugta!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  
  form1.nyugtavolt = 8
  Nyugel1.Hide

End Sub

Private Sub Command8_Click()
  If xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Üres garancia jegy!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  If Trim(Text2.Text) = "" Or Trim(Text3.Text) = "" Then
    Call mess("Garancia jegyre a vevõ neve és címe kötelezõ", 3, 0, "Hiba", valasz%)
    Exit Sub
  End If
  If vaneloleg Then
    Call mess("Elõleget számlába számítson be!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If

    ' Gyáriszám ellenõrzés
  For i1% = 1 To 200
    termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(i1%, 1)) + Space(15), 15)
    If Not Trim(termkod$) = "" Then
        ktrmrec$ = dbxkey("KTRM", termkod$)
        If Mid$(ktrmrec$, 849, 1) = "I" Then
           db% = 0
           For j1% = 1 To 200
             If Not Trim(gyariszamok$(i1%, j1%)) = "" Then
                db% = db% + 1
             End If
           Next
           If Not db% = MSFlexGrid1.TextMatrix(i1%, 3) Then
              Call mess(Str(i1%) + ".sorban nem " + MSFlexGrid1.TextMatrix(i1%, 3) + " gyári szám van rögzítve!", 3, 0, "Figyelmeztetés", valasz%)
              Exit Sub
           End If
        End If
    End If
  Next

  If Not Keszletellenorzes Then
     Exit Sub
  End If
  
  form1.nyugtavolt = 7
  Nyugel1.Hide

End Sub

Private Sub Command9_Click()
  If Not xval(Trim(Label1.Caption)) = 0 Then
    Call mess("Nem üres a nyugta!", 3, 0, "Figyelmeztetés", valasz%)
    Exit Sub
  End If
  
  form1.nyugtavolt = 6
  Nyugel1.Hide

End Sub

Private Sub Form_Activate()
  If betoltve% = 0 Then
    parteng@ = 0
    billscr% = 0
    betoltve% = 1
    For j1% = 1 To 200: Nyugel1.MSFlexGrid1.TextMatrix(j1%, 0) = Trim(Str(j1%)): Next
    For j1% = 1 To 200
      For k1% = 1 To 200
        gyariszamok$(j1%, k1%) = Space$(40)
      Next
    Next
    For i1% = 1 To 100: nt$(i1%) = Space(50): Next
    Call torzsbe("PFIZ", fizetesimod$(), fizmod&)

    MSFlexGrid1.Row = 1: MSFlexGrid1.Col = 1
    'MSFlexGrid1.FixedCols = 1
    MSFlexGrid1.SetFocus
    Text1.SetFocus
    
    autoinfo = 1
    vaneloleg = False
    If Trim$(ugyintezo$) = "ESZES" Then
       Command18.Visible = True
    Else
       Command18.Visible = False
    End If
    Text13.Text = ""
  Else
  ' MSFlexGrid1.SetFocus
  End If
  
End Sub


Private Sub MSFlexGrid1_Click()
      termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
      If Not Trim(termkod$) = "" Then
        ktrmrec$ = dbxkey("KTRM", termkod$)
        If Mid$(ktrmrec$, 849, 1) = "I" Then
           Call Gyariszambeker(Trim(Mid$(ktrmrec$, 16, 60)), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
        End If
        Text1.SetFocus
      End If
End Sub

Private Sub MSFlexGrid1_gotfocus()
  If Trim(Nyugel1.Text6.Text) = "" Then
    Nyugel1.Text6.Text = maidatum$
  End If
  If Trim(Nyugel1.Text9.Text) = "" Then
    Nyugel1.Text9.Text = maidatum$
  End If
  
  Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
  Text1.Height = MSFlexGrid1.CellHeight
  Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
  Text1.Width = MSFlexGrid1.CellWidth
  Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
  Text1.SelStart = Len(Trim(Text1.Text))
  termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
  utermkod$ = termkod$
  'If MSFlexGrid1.Col = 2 Or MSFlexGrid1.Col = 6 Or MSFlexGrid1.Col = 7 Then
  '     umText1$ = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
  'End If

  Text1.SetFocus
  
End Sub

Private Sub MSFlexGrid1_Scroll()
  If billscr% = 0 Then
    MSFlexGrid1.Row = MSFlexGrid1.toprow
    MSFlexGrid1.Col = MSFlexGrid1.LeftCol
  End If
  Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
  Text1.Height = MSFlexGrid1.CellHeight
  MSFlexGrid1.Col = 1
  Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
  Text1.Width = MSFlexGrid1.CellWidth
  Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
  Text1.SelStart = Len(Trim(Text1.Text))
  billscr% = 0
End Sub

Private Sub Text1_Click()
  Text1.SelStart = Len(Trim(Text1.Text))
End Sub

Private Sub Text1_GotFocus()
  termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
  If Trim(termkod$) <> "" Then
    utermkod$ = termkod$
  End If
  If MSFlexGrid1.Col = 1 Or MSFlexGrid1.Col = 2 Or MSFlexGrid1.Col = 3 Or MSFlexGrid1.Col = 6 Or MSFlexGrid1.Col = 7 Then
       uText1$ = Text1
  End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
10      On Error GoTo hiba
        '--- adat táblázat
20      billscr% = 0
30      If KeyCode = vbKeyReturn Then
          '--- mezõ ellenõrzés
          If form1.nyugtavolt = 12 Then Exit Sub
40        If MSFlexGrid1.Col = 2 Or MSFlexGrid1.Col = 3 Or MSFlexGrid1.Col = 6 Or MSFlexGrid1.Col = 7 Then
50         If MSFlexGrid1.Col = 3 Then
           
60           If Len(Trim$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))) = 7 Then
70             If Val(Text1) > (MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)) Then
80               Call mess("Többet nem adhat ki mint, amennyi a munkalapon van! (" + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)) + ")", 2, 0, "Hiba", valasz%)
90               Text1 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
100            Else
110             If Val(Text1) < (MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)) Then
120                bm = (MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)) - Val(Text1)
130                Call mess("Kevesebbet ad ki mint, amennyi a munkalapon van!" + Str(bm) + " db bent marad a szervizes raktárában!" + Chr(10) + Chr(13) + "( A munka lapon " + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)) + " volt.)", 2, 0, "Figyelmeztetés", valasz%)
140             End If
150            End If
160          Else
170          End If
             
180        Else
190             Text1 = uText1
200        End If
             
210       End If

220       khiba% = 0
230       If MSFlexGrid1.Col = 1 Then
            '--- termék kód, vagy vonalkód
240         ktrmkod$ = Left(Trim(Text1.Text) + Space(15), 15)
      ' Eszi - vonalkód
250         atex$ = ktrmkod$
260         For iaa% = 1 To Len(atex$)
270           If Mid$(atex$, iaa%, 1) = "ö" Then Mid$(atex$, iaa%, 1) = "0"
280         Next
290         btex$ = Left(Trim(atex$) + Space(15), 13)
300         reanrec$ = dbxkey("REAN", btex$)
310         If reanrec$ <> "" Then
320            Text1.Text = Mid$(reanrec$, 14, 15)
330         Else
340            If Mid$(btex$, 1, 1) = "0" Then
350              btex$ = Mid$(btex$, 2, 12) + " "
360              reanrec$ = dbxkey("REAN", btex$)
370              If reanrec$ <> "" Then
380                  Text1.Text = Mid$(reanrec$, 14, 15)
390              Else
400                Text1.Text = atex$
410              End If
420            End If
430         End If
440         Text1.Text = UCase(Text1.Text)
450         ktrmkod$ = Text1.Text
460         ktrmrec$ = dbxkey("KTRM", ktrmkod$)
470         jelleg$ = Mid$(ktrmrec$, 442, 1)
480         If ktrmrec$ = "" Then
490           Call mess("Hibás termék kód!", 2, 0, "Hiba", valasz%)
500           khiba% = 1
510         Else
              '--- termék adatok feltöltése
              '--- termék ára
512           If Mid$(ktrmrec$, 846, 1) = "L" Then
514                Call mess("Termék letiltva!", 2, 0, langmodul(99), valasz%)
516                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = Space$(15)
517                Text1.Text = Space$(15)
518                khiba% = 1
519           Else

520           If form1.munkalap = "" Or Not Len(Trim$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))) = 7 Then
              
              If Mid$(ktrmrec$, 716, 4) = "9133" Then
                 If Trim$(Text8.Text) = "" Or Trim$(Text8.Text) <> "A számla közvetített szolgáltatást tartalmaz." Then
                    If Trim$(Text8.Text) = "" Then
                       Text8.Text = "A számla közvetített szolgáltatást tartalmaz."
                    Else
                       If Trim$(Text11.Text) = "" Or Trim$(Text11.Text) <> "A számla közvetített szolgáltatást tartalmaz." Then
                          If Trim$(Text11.Text) = "" Then
                             Text11.Text = "A számla közvetített szolgáltatást tartalmaz."
                          Else
                             Call mess("A számára rá kell írni:" + Chr(10) + Chr(13) + "A számla közvetített szolgáltatást tartalmaz.", 2, 0, "Figyelmeztetés", valasz%)
                          End If
                       End If
                    End If
                 End If
              End If
530           krak$ = Left(Trim(form1.Text5.Text) + "    ", 4)
540           rkszkod$ = krak$ + ktrmkod$
550           rkszrec$ = dbxkey("RKSZ", rkszkod$)
560           keszle@ = 0
570           If rkszrec$ <> "" Then
580             stat$ = " "
590             If form1.megrendelesbol Then
                  
600               kmegrec$ = dbxkey("KMEG", form1.megrend$)
      '            krak$ = Mid$(kmegrec$, 480, 4)
610               stat$ = Mid$(kmegrec$, 192, 1)
620             End If
630             If stat$ = "D" Then
640               keszle@ = xval(Mid$(rkszrec$, 20, 12))
650             Else
660               keszle@ = xval(Mid$(rkszrec$, 20, 12)) - xval(Mid$(rkszrec$, 32, 12))
670             End If
680             If keszle@ <= 0 Then rkszrec$ = ""
690           End If
700           If rkszrec$ = "" And Not jelleg$ = "S" Then

        
710             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Trim(Mid$(ktrmrec$, 16, 60))
720             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = Trim(Mid$(ktrmrec$, 678, 14))
730             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = Trim(Mid$(ktrmrec$, 484, 6))
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = form1.ttkrak$
      '          Call mess("Nincs készlet!", 2, 0, "Hiba", valasz%)
      '          MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = ""
      '          MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = ""
      '          MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
      '          MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = ""
      '          Text1.Text = ""

                
      '          khiba% = 1
740           Else
750             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Trim(Mid$(ktrmrec$, 16, 60))
760             If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "" Then MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "1"
770             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = Trim(Mid$(ktrmrec$, 678, 14))
780             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = ""
790             If jelleg$ = "S" Then
800               MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = "Szolgáltatás"
810             Else
820               MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = Str(keszle@)
830             End If
840             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = Trim(Mid$(ktrmrec$, 484, 6))
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = form1.ttkrak$
                ' ide
850           End If
860           Else
      '          If MSFlexGrid1.Col = 1 Then
      '            Text1 = uText1
      '          End If
870             Call mess("Nem módosítható! Munkalapról származik.", 2, 0, "Hiba", valasz%)
880           End If
882         End If
890         End If
900       Else
            '--- egyéb mezõk
910         ktrmkod$ = Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) + Space(15), 15)
920         ktrmrec$ = dbxkey("KTRM", ktrmkod$)
930         jelleg$ = Mid$(ktrmrec$, 442, 1)

940         mezo$ = Trim(Text1.Text): khiba% = 0
950         If MSFlexGrid1.Col > 3 And MSFlexGrid1.Col < 6 Then
960           If MSFlexGrid1.Col = 5 Then
970             Call kodvizsg(mezo$, "NJT-", khiba%, 12)
980           Else
990             Call kodvizsg(mezo$, "NJT", khiba%, 12)
1000          End If
1010          If khiba% = 1 Then
1020            Call mess("Hibás adat!", 2, 0, "Hiba", valasz%)
1030          Else
1040              If (Not MSFlexGrid1.Col = 5 And Len(mezo$) > 7) Or (MSFlexGrid1.Col = 5 And ((xval(mezo$) >= 0 And Len(mezo$) > 2)) Or (xval(mezo$) < 0 And Len(mezo$) > 3)) Then
1050                Text1.Text = ""
1060                khiba% = 1
1070                Call mess("Túl nagy szám!", 3, 0, "Hiba", valasz%)
1080                If MSFlexGrid1.Col = 4 Then
1090                  Text1.Text = Trim(Mid$(ktrmrec$, 678, 14))
1100                End If
1110                Exit Sub
1120              End If
1130          End If
1140        Else
1150          If MSFlexGrid1.Col = 3 Then
1160            Call kodvizsg(mezo$, "NJT-", khiba%, 12)
1170            If khiba% = 0 Then
1180              If Len(mezo$) > 7 Then
1190                Text1.Text = ""
1200                khiba% = 1
1210                Call mess("Mennyiség túl sok!", 3, 0, "Hiba", valasz%)
1220                Text1.Text = "1"
1230                Exit Sub
1240              Else
1250                darab& = xval(mezo$)
1260              End If
1270              ktrmkod$ = Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) + Space(15), 15)
1280              ktrmrec$ = dbxkey("KTRM", ktrmkod$)
1290              If ktrmrec$ = "" Then
1300                khiba% = 1
1310                Call mess("Nincs ilyen termék!", 3, 0, "Hiba", valasz%)
1320              Else
1330                krak$ = Left(Trim(form1.Text5.Text) + "    ", 4)
1340                rkszkod$ = krak$ + ktrmkod$
1350                rkszrec$ = dbxkey("RKSZ", rkszkod$)
1360                keszle@ = 0
1370                If rkszrec$ <> "" Then
1380                  stat$ = " "
1390                  If form1.megrendelesbol Then
                        
1400                    kmegrec$ = dbxkey("KMEG", form1.megrend$)
          '             krak$ = Mid$(kmegrec$, 480, 4)
1410                    stat$ = Mid$(kmegrec$, 192, 1)
1420                  End If
1430                  If stat$ = "D" Then
1440                     keszle@ = xval(Mid$(rkszrec$, 20, 12))
1450                  Else
1460                     keszle@ = xval(Mid$(rkszrec$, 20, 12)) - xval(Mid$(rkszrec$, 32, 12))
1470                  End If
1480                End If
                    ' Már van ilyen cikkszám
1490                kiadossz@ = 0
1500                szov$ = ""
1510                For i3% = 1 To 200
1520                  If Not (MSFlexGrid1.Row = i3%) And MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = MSFlexGrid1.TextMatrix(i3%, 1) Then
1530                     kiadossz@ = kiadossz@ + xval(MSFlexGrid1.TextMatrix(i3%, 3))
1540                     szov$ = "/ Több tételben van ez az áru/"
1550                  End If
1560                Next
                    
1570                kiadossz@ = kiadossz@ + xval(mezo$)
                    
1580                If Not jelleg$ = "S" And keszle@ < kiadossz@ And xval(mezo$) > 0 Then
                      ' Ha nem munkalapról jött
1590                  If Not Len(Trim$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))) = 7 Then
1600                    Call mess("Készlet kevés!" + szov$, 2, 0, "Hiba", valasz%)
1610                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = 0
1620                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = 0
1630                    Text1.Text = 0
1640                    Call ujraszamol
1650                    khiba% = 1
1660                  End If
1670                End If
1680              End If
1690            Else
1700              khiba% = 1
1710              Call mess("Hibás adat!", 2, 0, "Hiba", valasz%)
1720            End If
1730          End If
1740        End If
1750        If khiba% = 0 Then
1760          If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) = "" Then
1770            Call mess("Termék kód kötelezõ!", 2, 0, "Hiba", valasz%)
1780            khiba% = 1

1790          End If
1800        End If
1810      End If
1820      If khiba% = 0 Then
1830        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col) = Trim(Text1.Text)
            ' Eszi ide
1840        If MSFlexGrid1.Col = 1 Then
1850          MSFlexGrid1.Col = MSFlexGrid1.Col + 2
1860        Else
1870          If MSFlexGrid1.Col < 6 Then
1880            MSFlexGrid1.Col = MSFlexGrid1.Col + 1
1890          End If
1900        End If
1910        billscr% = 1
1920        If MSFlexGrid1.Col < 6 Then
1930        Else
1940          If MSFlexGrid1.Row < 200 Then
1950            termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
1960            ktrmrec$ = dbxkey("KTRM", termkod$)
1970            If Mid$(ktrmrec$, 849, 1) = "I" Then
1980               db% = xval(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
1990               Call Gyariszambeker(Trim(Mid$(ktrmrec$, 16, 60)), db%)
2000            End If
                      
2010            billscr% = 1
2020            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
2030            MSFlexGrid1.Col = 1
2040          End If
2050        End If
2060        KeyCode = 0
2070        Call ujraszamol
2080        MSFlexGrid1.SetFocus
2090      End If
2100    Else
2110      Select Case KeyCode
            Case vbKeyDelete
2120          Text1.Text = ""
2130        Case vbKeyF1
        '      Shell programutvonal$ + "Winszamla.hlp"
              
2140        Case vbKeyF5
2150          Call Command2_Click
2160        Case vbKeyF3
2170          Call Command1_Click
2180        Case vbKeyAdd
2190          Call Command3_Click
2200        Case vbKeyA And (Shift And vbAltMask) > 0
2210          Call Command5_Click
2220        Case vbKeyG And (Shift And vbAltMask) > 0
2230        ' Call Command8_Click
2240        Case vbKeyX, vbKeyH
2250          If Shift And vbAltMask Then
2260            oszlop% = MSFlexGrid1.Col
2270            If oszlop% = 1 Then
2280              If KeyCode = vbKeyX Then
2290                Call altx("KTRM", azonosito$)
2300              Else
2310                Call alth("KTRM", azonosito$)
2320              End If
2330              If azonosito$ <> "" Then
2340                If Not Len(Trim$(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))) = 7 Then
2350                   Text1.Text = azonosito$
2360                   MySendKeys "{ENTER}"
2370                Else
2380                   Call mess("Nem módosítható! Munkalapról származik.", 2, 0, "Hiba", valasz%)
2390                End If
2400              End If
2410            End If
2420          End If
2430        Case vbKeyHome
2440          billscr% = 1
2450          MSFlexGrid1.Row = 1: MSFlexGrid1.Col = 1: KeyCode = 0: MSFlexGrid1.SetFocus
2460        Case vbKeyEnd
2470          billscr% = 1
2480          MSFlexGrid1.Row = 200: MSFlexGrid1.Col = 1: KeyCode = 0: MSFlexGrid1.SetFocus
2490        Case vbKeyPageDown
2500          billscr% = 1
              
2510          termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
2520          ktrmrec$ = dbxkey("KTRM", termkod$)
2530          If Mid$(ktrmrec$, 849, 1) = "I" Then
2540               Call Gyariszambeker(Trim(Mid$(ktrmrec$, 16, 60)), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
2550          End If

2560          If MSFlexGrid1.Row + 10 <= 200 Then
2570            MSFlexGrid1.Row = MSFlexGrid1.Row + 10: KeyCode = 0: MSFlexGrid1.SetFocus
2580          Else
2590            MSFlexGrid1.Row = 200: KeyCode = 0: MSFlexGrid1.SetFocus
2600          End If
2610        Case vbKeyPageUp
2620          termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
2630          ktrmrec$ = dbxkey("KTRM", termkod$)
2640          If Mid$(ktrmrec$, 849, 1) = "I" Then
2650               Call Gyariszambeker(Trim(Mid$(ktrmrec$, 16, 60)), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
2660          End If
              
2670          billscr% = 1
2680          If MSFlexGrid1.Row - 10 >= 1 Then
2690            MSFlexGrid1.Row = MSFlexGrid1.Row - 10: KeyCode = 0: MSFlexGrid1.SetFocus
2700          Else
2710            MSFlexGrid1.Row = 1: KeyCode = 0: MSFlexGrid1.SetFocus
2720          End If
2730        Case vbKeyUp
2740          termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
2750          ktrmrec$ = dbxkey("KTRM", termkod$)
2760          If Mid$(ktrmrec$, 849, 1) = "I" Then
2770               Call Gyariszambeker(Trim(Mid$(ktrmrec$, 16, 60)), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
2780          End If
            
2790          billscr% = 1
2800          If MSFlexGrid1.Row > 1 Then MSFlexGrid1.Row = MSFlexGrid1.Row - 1: KeyCode = 0: MSFlexGrid1.SetFocus
2810        Case vbKeyDown
2820          termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
2830          ktrmrec$ = dbxkey("KTRM", termkod$)
2840          If Mid$(ktrmrec$, 849, 1) = "I" Then
2850               Call Gyariszambeker(Trim(Mid$(ktrmrec$, 16, 60)), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
2860          End If
       
2870          billscr% = 1
2880          If MSFlexGrid1.Row < 200 Then MSFlexGrid1.Row = MSFlexGrid1.Row + 1: KeyCode = 0: MSFlexGrid1.SetFocus
2890        Case vbKeyLeft
2900          billscr% = 1
2910          If MSFlexGrid1.Col = 3 Then
2920            MSFlexGrid1.Col = MSFlexGrid1.Col - 2: KeyCode = 0: MSFlexGrid1.SetFocus
2930          Else
2940            If MSFlexGrid1.Col > 1 Then MSFlexGrid1.Col = MSFlexGrid1.Col - 1: KeyCode = 0: MSFlexGrid1.SetFocus
2950          End If
2960        Case vbKeyRight
2970          billscr% = 1
2980          If MSFlexGrid1.Col = 1 Then
2990             MSFlexGrid1.Col = MSFlexGrid1.Col + 2: KeyCode = 0: MSFlexGrid1.SetFocus
3000          Else
3010             If MSFlexGrid1.Col < 5 Then MSFlexGrid1.Col = MSFlexGrid1.Col + 1: KeyCode = 0: MSFlexGrid1.SetFocus
3020          End If
3030        Case Else
3040      End Select
3050    End If
3060  Exit Sub
hiba:
        
3070    hbuz$ = "Hiba: AUW-QRPTN (Text1.KeyDown) " + "Sor:" + Str$(Erl) + " " + Str$(Err.Number) + " " + Err.Description
3080    Call mess(hbuz$, 1, 0, "Hiba", valasz%)
             
3090    Call hibakiir(2, hbuz$)
      'Resume Next
End Sub
Private Sub hibakiir(hiv%, hbuz$)
  hbdatum$ = "* " + Right(Date$, 2) + "." + Left(Date$, 2) + "." + Mid$(Date$, 4, 2) + " " + Str(Time) + " " + terminal$ + task$ + " " + ugyintezo$
  
  
  fhb = FreeFile
  Open auditorutvonal$ + "hiba.txt" For Append As #fhb
  Print #fhb, hbdatum$
  Print #fhb, hbuz$
  
End Sub
Public Sub ujraszamol()
  On Error GoTo hiba
 
  tk1o@ = 0
  For j9% = 1 To 200
    tk1$ = Trim(MSFlexGrid1.TextMatrix(j9%, 1))
    If tk1$ <> "" Then
      tkmenny@ = xval(Trim(MSFlexGrid1.TextMatrix(j9%, 3)))
      tkear@ = xval(Trim(MSFlexGrid1.TextMatrix(j9%, 4)))
      enge@ = xval(Trim(MSFlexGrid1.TextMatrix(j9%, 5)))
      If enge@ <> 0 Then
        engft1@ = tkear@ * enge@ / 100
        engft2@ = xval(Trim(ertszam(Str(engft1@), 12, 0)))
        tk1o@ = tk1o@ + tkmenny@ * (tkear@ - engft2@)
      Else
        tk1o@ = tk1o@ + tkmenny@ * tkear@
      End If
    End If
  Next
  For j9% = 1 To 5
     tk1o@ = tk1o@ - Val(Mid$(nt$(j9%), 24, 14))
  Next
  ' Elõleg beszámítása ide
  Label1.Caption = ertszam(Str(tk1o@), 12, 2)
  Exit Sub
hiba:
  Exit Sub
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
       
     MSFlexGrid1.SetFocus
     KeyCode = 0
  End If
End Sub

Private Sub Text3_GotFocus()
    ' Ha nem választ kp partnert és nincs partnerkód
    If (Trim(Text4(1).Text) = "" Or Trim(Text4(1).Text) = "BANK") And Trim(Text13.Text) = "" Then
       Text13.Text = "Új"
       Call Command22_Click
       Text3.Enabled = False
    End If

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Text5.SelStart = Len(Trim(Text5.Text)) + 1
    Text5.SetFocus
    
    KeyCode = 0
 End If
End Sub


Private Sub Text4_GotFocus(Index As Integer)
form1.rpartner = Text4(1).Text
End Sub

Private Sub Text4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyX, vbKeyH, vbKeyInsert
      If Shift And vbAltMask Then
        If KeyCode = vbKeyX Then
          Call altx("PART", azonosito$)
        End If
        If KeyCode = vbKeyH Then
          Call alth("PART", azonosito$)
        End If
        If KeyCode = vbKeyInsert Then
          rrepar$ = "AUWSZAMV/PART" + "/" + terminal$ + task$ + "/" + auditorutvonal$
          rre = Shell(programutvonal$ + "DBX4-NEW.EXE " + rrepar$, vbNormalFocus)
          Call mess(langmodul(162), 4, 0, langmodul(163), valasz%)
          pxfx = FreeFile
          Open programutvonal$ + terminal$ + task$ + "new.txt" For Binary Shared As pxfx
          csx$ = Space(3)
          Get #pxfx, 1, csx$
          If csx$ <> "NIX" Then
            csx$ = Space(15)
            Get #pxfx, 1, csx$
            azonosito$ = csx$
          End If
          Close pxfx
        End If
        If azonosito$ <> "" Then
          Text4(1).Text = azonosito$
        End If
      End If
    Case vbKeyReturn
      pkod$ = UCase(Left(Trim(Text4(1).Text + Space(15)), 15))
      Text4(1).Text = pkod$
      If Trim(pkod$) <> "" Then
        prec$ = dbxkey("PART", pkod$)
        welobepartrec = prec$
        If Not Trim$(pkod$) = "BANK" Then
           Text13.Text = ""
        End If
        If prec$ = "" Then
          Call mess("Hibás vevõ kód!", 3, 0, "Hiba", valasz%)
          KeyCode = 0
        Else
          If Trim$(pkod$) = "BANK" Then
            If Trim$(Text2.Text) = "" Then
              'Text2.Text = Trim(Mid$(prec$, 16, 60))
              'Text3.Text = postacim(prec$, 106)
              'Text3.Enabled = False

            End If
          Else
            Text2.Text = Trim(Mid$(prec$, 16, 60))
            Text3.Text = postacim(prec$, 106)
            Text12.Text = Trim(Mid$(prec$, 184, 15))
            If Not Mid$(Text12.Text, 9, 1) = "-" And Len(Trim(Text12.Text)) = 11 Then
              Text12.Text = Mid$(Text12.Text, 1, 8) + "-" + Mid$(Text12.Text, 9, 1) + "-" + Mid$(Text12.Text, 10, 2)
            End If
            If RTrim$(Mid$(prec$, 553, 10)) = "" Then
               Call Command21_Click
            End If
          End If
          fmkod$ = Mid$(prec$, 328, 2)
          If Not Trim$(fmkod$) = "" Then
            fmrec$ = dbxkey("PFIZ", fmkod$)
            Text5.Text = Mid$(fmrec$, 3, 30)
            Text7.Text = fmkod$
          End If
          fizhatido% = Val(Mid$(prec$, 330, 3))
          fidat$ = maidatum$
          For i13% = 1 To fizhatido%
            xxxx$ = novdat(fidat$)
            fidat$ = xxxx$
          Next
          Text6.Text = fidat$
          Text2.SelStart = Len(Trim(Text2.Text)) + 1
          Text2.SetFocus
          
          'MSFlexGrid1.SetFocus
        End If
      End If
    Case Else
  End Select
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
  If Trim$(Text4(1).Text) = "" Or Mid$(Text4(1).Text, 1, 4) = "BANK" Then
  Select Case KeyCode
    Case vbKeyX, vbKeyH
      If Shift And vbAltMask Then
        Text3.Enabled = True
        If KeyCode = vbKeyX Then
          Call altx("KPAR", azonosito$)
        End If
        If KeyCode = vbKeyH Then
          Call alth("KPAR", azonosito$)
        End If
        If azonosito$ <> "" Then
          kparrec$ = dbxkey("KPAR", azonosito$)
          If kparrec$ <> "" Then
             Text2.Text = Mid$(kparrec$, 1, 60)
             Text3.Text = Mid$(kparrec$, 61, 60)
             Text13.Text = azonosito$
             Text3.Enabled = False
             kparrec$ = dbxkey("KCIM", azonosito$)
             If kparrec$ = "" Then
               kparrec$ = form1.nevcimkeres$(Text2.Text + Text3.Text)
               
               If kparrec$ = "" Then
                Call Command22_Click
               End If
             End If
             
          End If
        End If
      End If
    
    Case Else
  End Select
  End If
  
  If KeyCode = vbKeyReturn Then
    ' Ha nem választ kp partnert és nincs partnerkód
    If (Trim(Text4(1).Text) = "" Or Trim(Text4(1).Text) = "BANK") And (Trim(Text13.Text) = "" Or Trim(Text13.Text) = "Új") Then
       Text13.Text = "Új"
       Call Command22_Click
    End If
    Text5.SelStart = Len(Trim(Text5.Text)) + 1
    Text5.SetFocus
    Text3.Enabled = False
    KeyCode = 0
  End If
  
  
End Sub


Private Function kellen(cikk$)
  '--- kiadás készleteinek ellenõrzése
  khiba% = 1: om@ = 0
  termrec$ = dbxkey("KTRM", cikk$)
  If Mid$(termrec$, 443, 1) = "N" Then kellen = 0: Exit Function
  For j9% = 1 To 200
    tk1$ = Left(Trim(MSFlexGrid1.TextMatrix(j9%, 1)) + Space(15), 15)
    If tk1$ = cikk$ Then
      tkmenny@ = xval(Trim(MSFlexGrid1.TextMatrix(j9%, 3)))
      om@ = om@ + tkmenny@
    End If
  Next
  Rkindex$ = form1.ttkrak + cikk$
  rkszrec$ = dbxkey("RKSZ", Rkindex$)
  If termrec$ <> "" Then
    rkmenny@ = xval(Trim(Mid$(termrec$, 955, 14)))
    If rkmenny@ >= om@ Then khiba% = 0
  End If
  kellen = khiba%
End Function

Private Function vegosszege@(a%)
  tk2o@ = 0
  For j9% = 1 To 200
    tk1$ = Trim(MSFlexGrid1.TextMatrix(j9%, 1))
    If tk1$ <> "" Then
      tkmenny@ = xval(Trim(MSFlexGrid1.TextMatrix(j9%, 3)))
      tkear@ = xval(Trim(MSFlexGrid1.TextMatrix(j9%, 4)))
      enge@ = xval(Trim(MSFlexGrid1.TextMatrix(j9%, 5)))
      If enge@ <> 0 Then
        engft1@ = tkear@ * enge@ / 100
        engft2@ = xval(Trim(ertszam(Str(engft1@), 12, 0)))
        tk2o@ = tk2o@ + tkmenny@ * (tkear@ + engft2@)
      Else
        tk2o@ = tk2o@ + tkmenny@ * tkear@
      End If
    End If
  Next
  vegosszege@ = tk2o@
End Function

Private Function tetelertek@(sor%, hmod%)
  If hmod% = 0 Then
    tkmenny@ = xval(Trim(Text1.Text))
  Else
    tkmenny@ = xval(Trim(MSFlexGrid1.TextMatrix(sor%, 3)))
  End If
  trekdb@ = xval(Trim(MSFlexGrid1.TextMatrix(sor%, 7)))
  tkear@ = xval(Trim(MSFlexGrid1.TextMatrix(sor%, 5)))
  tpgar@ = xval(Trim(MSFlexGrid1.TextMatrix(sor%, 6)))
  tggar@ = xval(Trim(MSFlexGrid1.TextMatrix(sor%, 8)))
  tetelertek = tkmenny@ * tkear@ + tkmenny@ * tpgar@ + trekdb@ * tggar@
End Function

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyX, vbKeyH, vbKeyInsert
      If Shift And vbAltMask Then
        If KeyCode = vbKeyX Then
          Call altx("PFIZ", azonosito$)
        End If
        If KeyCode = vbKeyH Then
          Call alth("PFIZ", azonosito$)
        End If
        If azonosito$ <> "" Then
          fmrec$ = dbxkey("PFIZ", azonosito$)
          Text5.Text = Mid$(fmrec$, 3, 30)
          Text7.Text = azonosito$
        End If
        
      End If
    Case vbKeyReturn
       ' Eszi
       ' leellenõrizni a beírt fizetési mód jó-e?
       van = 0
       For i& = 1 To fizmod&
         If Trim(Mid$(UCase(fizetesimod$(i&)), 3, 30)) = Trim(UCase(Text5.Text)) Then
           Text5.Text = (Mid$(fizetesimod$(i&), 3, 30))
           Text7.Text = Mid$(fizetesimod$(i&), 1, 2)
           van = 1
           Exit For
         End If
       Next
       If van = 0 Then
         Call mess("Nincs ilyen fizetési mód", 3, 0, "Hiba", valasz%)
         KeyCode = 0
         Exit Sub
       End If
       If InStr(UCase(Text5.Text), "BANK") > 0 Then
          If Trim(Text4(1).Text) = "" Then
             Text4(1).Text = BANKPARTN
          End If
          Text6.Text = maidatum$
       End If
       If InStr(UCase(Text5.Text), "KÉSZPÉNZ") > 0 Then
          Text9.Text = maidatum$
          Text6.Text = maidatum$
       End If
       Text6.SelStart = Len(Trim(Text6.Text)) + 1
       Text6.SetFocus
       KeyCode = 0
    Case Else
  End Select

End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Text9.SelStart = Len(Trim(Text5.Text)) + 1
    Text9.SetFocus
    
    KeyCode = 0
 End If
  
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Text11.SelStart = Len(Trim(Text8.Text)) + 1
    Text11.SetFocus
    
    KeyCode = 0
 End If

End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
       If InStr(UCase(Text5.Text), "KÉSZPÉNZ") > 0 And Not (Text9.Text = maidatum$) Then
          Text9.Text = maidatum$
          Call mess("Készpénzes fizetési mód esetén a teljesítés kelte csak a mai nap lehet!", 3, 0, "Hiba", valasz%)
       End If
     
     MSFlexGrid1.SetFocus
  End If

End Sub

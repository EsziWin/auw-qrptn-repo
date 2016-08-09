VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Gyariszam 
   BackColor       =   &H0091E9FB&
   Caption         =   "Gy�rt�si sz�mok"
   ClientHeight    =   5268
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5376
   HelpContextID   =   300
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5268
   ScaleWidth      =   5376
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C5FCC8&
      Caption         =   "Minden sor t�rl�s"
      Height          =   252
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Miden gy�risz�m t�rl�se"
      Top             =   5050
      Width           =   1452
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   9
      ToolTipText     =   "Intervallum v�ge vagy db + Enter r�gz�ti"
      Top             =   4330
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Intervallum kezdete - null�val is kezd�dhet"
      Top             =   4330
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C5FCC8&
      Caption         =   "Sor t�rl�s"
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
      TabIndex        =   5
      Top             =   4800
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "R�gz�t"
      Height          =   372
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Kil�p�s a tartalom r�gz�t�s�vel (Esc)"
      Top             =   4920
      Width           =   1692
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "M�gsem"
      Height          =   372
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Kil�p�s a tartalom r�gz�t�se n�lk�l (Ctrl+T)"
      Top             =   4920
      Width           =   1692
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0091E9FB&
      BorderStyle     =   0  'None
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   960
      TabIndex        =   1
      Top             =   2160
      Width           =   2772
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3492
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4812
      _ExtentX        =   8488
      _ExtentY        =   6160
      _Version        =   327680
      BackColorBkg    =   -2147483643
   End
   Begin VB.Label Label3 
      BackColor       =   &H0091E9FB&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2640
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label Label2 
      BackColor       =   &H0091E9FB&
      Caption         =   "Gy�rt�si sz�m intervallum:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   4572
   End
End
Attribute VB_Name = "Gyariszam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim utermkod$, billscr%, hiba
Dim gysz$(200, 2), gyszt$(200), hivt$(200)
Public betoltvegy%, darab%, biztip$, raktar$, cikk$

Private Sub Command1_Click()
    For i77% = 1 To 200
       MSFlexGrid1.TextMatrix(i77%, 1) = ""
       MSFlexGrid1.TextMatrix(i77%, 2) = ""
    Next
   Text1.Text = ""
   MSFlexGrid1.Row = 1

End Sub

Private Sub Command2_Click()
  db% = 0
  For i78% = 1 To 200
    If Not Trim(MSFlexGrid1.TextMatrix(i78%, 1)) = "" Then
       db% = db% + 1
    End If
  Next
  If db% = darab% Then
    rogzites% = 1
    gyszelso = False
    Gyariszam.MSFlexGrid1.Row = 1
    Gyariszam.Text1.SetFocus
    
    Me.Hide
  Else
     If hiba Then
       Call mess("Hib�sak a gy�risz�mok!", 3, 0, "Hiba", valasz%)
     Else
       Call mess(Str(darab%) + " gy�rt�si sz�mot kell r�gz�teni!", 3, 0, "Hiba", valasz%)
       hiba = True
     End If
  End If
End Sub

Private Sub Command3_Click()
  '--- Kil�p�s r�gz�t�s n�lk�l
  rogzites% = 0
  gyszelso = False
  Gyariszam.MSFlexGrid1.Row = 1
  Gyariszam.Text1.SetFocus

  Me.Hide

End Sub

Private Sub Command4_Click()
  If MSFlexGrid1.Row < 200 Then
    For i77% = MSFlexGrid1.Row To 199
      For i78% = 1 To 2
        MSFlexGrid1.TextMatrix(i77%, i78%) = MSFlexGrid1.TextMatrix(i77% + 1, i78%)
      Next
    Next
  End If
  For i78% = 1 To 2
    MSFlexGrid1.TextMatrix(200, i78%) = ""
  Next
  Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
  Text1.SelStart = Len(Trim(Text1.Text))
  Text1.SetFocus

End Sub

Private Sub Form_Activate()
 
  If betoltvegy% = 0 Then
    parteng@ = 0
    billscr% = 0
    betoltvegy% = 1
    MSFlexGrid1.Row = 1: MSFlexGrid1.Col = 1
    ' 5-�s hib�val itt elsz�ll, ha hiba �zenet van
    Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    Text1.Height = MSFlexGrid1.CellHeight
    Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    Text1.Width = MSFlexGrid1.CellWidth
    Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
    Text1.SelStart = Len(Trim(Text1.Text))
    termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
    Text1.SetFocus
    hiba = False
    autoinfo = 1
  Else
'    MSFlexGrid1.Row = 1
    Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    Text1.Height = MSFlexGrid1.CellHeight
    Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    Text1.Width = MSFlexGrid1.CellWidth

    Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
    Text1.SelStart = Len(Trim(Text1.Text))

  End If

End Sub

Private Sub Form_Load()
Me.Top = 2500
Me.Left = 6700
hiba = False

End Sub

Private Sub MSFlexGrid1_gotfocus()
  Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
  Text1.Height = MSFlexGrid1.CellHeight
  Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
  Text1.Width = MSFlexGrid1.CellWidth
  Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
  Text1.SelStart = Len(Trim(Text1.Text))
  termkod$ = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) + Space(15), 15)
'  utermkod$ = termkod$
  Text1.SetFocus

End Sub
Private Sub MSFlexGrid1_Scroll()
  If billscr% = 0 Then
    MSFlexGrid1.Row = MSFlexGrid1.toprow
    MSFlexGrid1.Col = MSFlexGrid1.LeftCol
  End If
  Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
  Text1.Height = MSFlexGrid1.CellHeight
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
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- adat t�bl�zat
  billscr% = 0
  If KeyCode = vbKeyReturn Then
    '--- mez� ellen�rz�s
    khiba% = 0
    If MSFlexGrid1.Col = 1 Then
      '--- gy�risz�m
      ktrmkod$ = Trim(Text1.Text)

      atex$ = ktrmkod$
      For iaa% = 1 To Len(atex$)
        If Mid$(atex$, iaa%, 1) = "�" Then Mid$(atex$, iaa%, 1) = "0"
      Next
      Text1.Text = atex$
      
      hiba = False
      
      
      ' Ha kiad�s
      If Not biztip$ = "BS" Then
        If atex$ = "" Then
           'Ha �res, akkor felhozza a bentl�v� gy�risz�mokat
           
            Hash2.cikk = cikk$
            Hash2.raktar = raktar$
            
           aktrox& = 1: toprox& = 1: aktdarab& = 0: tabstatusx% = 0
            Hash2.HelpContextID = 400
            
            Call dbxhash2("KKFX", 1, talalat%, kkfxrec$)
            If talalat% Then
              Text1.Text = Mid$(kkfxrec$, 200, 40)
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Mid$(kkfxrec$, 310, 10)
              atex$ = Mid$(kkfxrec$, 200, 40)
              
            End If
        Else
           'Ha nem �res ellen�rzi
             
            kkfxrec$ = Hash2.GyariszamKer(raktar$, cikk$, atex$, biztip$)
            If kkfxrec$ = "" Then
              Text1.Text = ""
              Call mess("Nincs ilyen gy�rt�si sz�m: " + atex$, 3, 0, "Hiba", valasz%)
              hiba = True
            Else
              MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Mid$(kkfxrec$, 310, 10)
            End If
         End If
      Else
       ' FElvitel vagy m�dos�t�s
        kkfxrec$ = Hash2.GyariszamKer(raktar$, cikk$, Text1.Text, biztip$)
        If Not kkfxrec$ = "" Then
           If Not Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) = "" Then
             If Mid$(kkfxrec$, 48, 7) <> MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) Then
               Call mess("M�r van ilyen gy�rt�si sz�m: " + Mid$(kkfxrec$, 48, 7) + ". bizonylat iktat�n", 3, 0, "Hiba", valasz%)
               hiba = True
             End If
           Else
             Call mess("M�r van ilyen gy�rt�si sz�m: " + Mid$(kkfxrec$, 48, 7) + ". bizonylat iktat�n", 3, 0, "Hiba", valasz%)
             hiba = True
           End If
           
        End If

      End If
      ' Ha el�z�leg ugyanazt a cikket r�gz�tik - van-e m�r ez a gy�risz�m
      If Not hiba Then
       If biztip$ = "ER" Then
        For iaa% = 1 To Nyugel1.MSFlexGrid1.Row - 1
         ecikk$ = Nyugel1.MSFlexGrid1.TextMatrix(iaa%, 1)
         
         If Trim(ecikk$) = Trim(cikk$) Then
            gyariszamok$ = Nyugel1.GyariszamAtad(iaa%, gyszt$(), hivt$())
            pzx% = InStr(gyariszamok$, ":")
            db% = Val(Mid$(gyariszamok$, 1, pzx% - 1))

            For iab% = 1 To db%
              If Trim(Text1.Text) = gyszt$(iab%) Then
                 Call mess("M�r van ilyen gy�rt�si sz�m: " + Str(iaa%) + ". forgalmi t�tel" + Str(iab%) + ". sor�ban!", 3, 0, "Hiba", valasz%)
                 Text1.Text = ""
                 hiba = True
                 Exit For
              End If
            Next
         End If
        Next
      
       Else
        For iaa% = 1 To Tabla.MSFlexGrid1.Row - 1
         If Tabla.MSFlexGrid1.Cols = 4 Then
            ecikk$ = Tabla.MSFlexGrid1.TextMatrix(iaa%, 1)
         Else
            ecikk$ = Tabla.MSFlexGrid1.TextMatrix(iaa%, 3)
         End If
         If Trim(ecikk$) = Trim(cikk$) Then
            Call GyariszamAtadM3(iaa%, gysz$(), db%)
            For iab% = 1 To db%
              If Trim(Text1.Text) = gysz$(iab%, 1) Then
                 Call mess("M�r van ilyen gy�rt�si sz�m: " + Str(iaa%) + ". forgalmi t�tel" + Str(iab%) + ". sor�ban!", 3, 0, "Hiba", valasz%)
                 Text1.Text = ""
                 hiba = True
                 Exit For
              End If
            Next
         End If
        Next
       End If
      End If
      ' Gyariszam.MsFlexgridben van-e
      If MSFlexGrid1.Row <= darab% Then
        If Not hiba Then
         MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col) = Trim(Text1.Text)
         For iaa% = 1 To MSFlexGrid1.Row - 1
          If MSFlexGrid1.TextMatrix(iaa%, 1) = Trim(atex$) Then
             Call mess("M�r van ilyen gy�rt�si sz�m: " + Str(iaa%) + ". sorban", 3, 0, "Hiba", valasz%)
             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = ""
             Text1.Text = ""
             hiba = True
             Exit For
          End If
         Next
        End If
      Else
        Call mess(Str(darab%) + " gy�rt�si sz�mot kell r�gz�teni!", 3, 0, "Hiba", valasz%)
        hiba = True
      End If
     
      billscr% = 1
      If MSFlexGrid1.Row < 200 Then
          billscr% = 1
          If Not hiba Then
            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
          Else
          
          End If
          MSFlexGrid1.Col = 1
      End If
      KeyCode = 0
      MSFlexGrid1.SetFocus
    End If
  ElseIf KeyCode = vbKeyEscape Then
    KeyCode = 0:
    Call Command2_Click
  ElseIf KeyCode = vbKeyUp Then
      If MSFlexGrid1.Row > 1 Then
          billscr% = 1
          MSFlexGrid1.Row = MSFlexGrid1.Row - 1
          MSFlexGrid1.Col = 1
      End If
      KeyCode = 0
      MSFlexGrid1.SetFocus
  
  ElseIf KeyCode = vbKeyDown Then
        If MSFlexGrid1.Row < 200 Then
          billscr% = 1
          MSFlexGrid1.Row = MSFlexGrid1.Row + 1
          MSFlexGrid1.Col = 1
      End If
      KeyCode = 0
      MSFlexGrid1.SetFocus
  ElseIf KeyCode = vbPageUp Then
      billscr% = 1
      MSFlexGrid1.Row = MSFlexGrid1.Row - 14
      MSFlexGrid1.Col = 1

      If MSFlexGrid1.Row < 1 Then
         MSFlexGrid1.Row = 1
      End If
      KeyCode = 0
      MSFlexGrid1.SetFocus
  
  ElseIf KeyCode = vbPageDown Then
      billscr% = 1
      MSFlexGrid1.Row = MSFlexGrid1.Row + 14
      MSFlexGrid1.Col = 1

      If MSFlexGrid1.Row > 200 Then
         MSFlexGrid1.Row = 200
      End If
      KeyCode = 0
      MSFlexGrid1.SetFocus
  
  End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 20 Then
  Call Command3_Click
End If
End Sub

Private Sub Text2_Change()
 Text2.SelStart = Len(Trim(Text2.Text))
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Text3.SelStart = Len(Trim(Text3.Text)) + 1
    Text3.SetFocus
    KeyCode = 0
  End If

End Sub

Private Sub Text3_Change()
Text3.SelStart = Len(Trim(Text3.Text))
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
   If xval(Text2.Text) <= xval(Text3.Text) Then
      kezd% = 1
      For iaa% = 1 To 200
        If Trim(MSFlexGrid1.TextMatrix(iaa%, 1)) = "" Then
           kezd% = iaa%
           Exit For
        End If
      Next
      szam% = 0
      For iaa% = Len(Text2.Text) To 1 Step -1
         If Not (Mid$(Text2.Text, iaa%, 1) >= "0" And Mid$(Text2.Text, iaa%, 1) <= "9") Then
            szam% = iaa%
            Exit For
         End If
      Next
      'szam% = szam% - 1
      prefix$ = Mid$(Text2.Text, 1, szam%)
      Text2.Text = Mid$(Text2.Text, szam% + 1, Len(Text2.Text) - szam%)
      If Mid$(Text3.Text, 1, szam%) = prefix$ Then
        Text3.Text = Mid$(Text3.Text, szam% + 1, Len(Text3.Text) - szam%)
        db% = xval(Text3.Text) - xval(Text2.Text)
      Else
        db% = xval(Text3.Text) - 1
      End If
      kezdert@ = xval(Text2.Text)
      hossz% = Len(Trim(Text2.Text))
      If Mid$(Trim(Text2.Text), 1, 1) = "0" Then
        kieg$ = String(hossz%, "0")
      Else
        kieg$ = Space(hossz%)
      End If
      For iaa% = 0 To db%
        MSFlexGrid1.TextMatrix(kezd% + iaa%, 1) = prefix$ + Right(kieg$ + Trim(Str(kezdert@ + iaa%)), hossz%)
        kkfxrec$ = Hash2.GyariszamKer(raktar$, cikk$, MSFlexGrid1.TextMatrix(kezd% + iaa%, 1), "BS")
        If Not kkfxrec$ = "" Then
             If Not Mid$(kkfxrec$, 320, 1) = "S" Then
               Call mess("M�r van ilyen gy�rt�si sz�m: " + Mid$(kkfxrec$, 48, 7) + ". bizonylat iktat�n", 3, 0, "Hiba", valasz%)
               hiba = True
               Exit For
             End If
        End If

      Next
      MSFlexGrid1.Row = kezd% + iaa%
      MSFlexGrid1.SetFocus
      Text2.Text = ""
      Text3.Text = ""
  Else
     Call mess("Hib�s intervallum!", 3, 0, "Hiba", valasz%)
     
     Text2.SetFocus
  End If

   KeyCode = 0
End If
End Sub

Private Sub dbxhash2(objazon, umod%, talalat%, rec$)
  '--- keres�s egy objektumba hashk�d alapj�n
  '--- gombsor csak kezdolap, umod%=0 eset�n
  '--- umod%=1 keres�t�bla, umod%=0 kezd�lap
  '--- umod%=1 eset�n talalat%=1 van kiv�lasztott, talalat%=0 nincs
  '--- rec a v�lasztott rekord (objtab-ban is be�ll�tva)
  Dim param$(10)
  keresomod% = umod%
  Hash2.MSFlexGrid1.Cols = 0
  Hash2.MSFlexGrid1.Rows = 15
  w1& = obsorszama(objazon)
  w2& = OBJTAB(w1&).hashmez(0)
  For i1& = 1 To w2&
    w3& = OBJTAB(w1&).hashmez(i1&)
    Hash2.MSFlexGrid1.Cols = Hash2.MSFlexGrid1.Cols + 1
    oslo% = Hash2.MSFlexGrid1.Cols - 1
    ne$ = RTrim$(ADATAB(w3&).adatnev)
    ar$ = RTrim$(ADATAB(w3&).attr)
    mh% = ADATAB(w3&).adatho
    kp% = ADATAB(w3&).adatkp
    Hash2.MSFlexGrid1.TextMatrix(0&, oslo%) = ne$
    w3& = Len(ne$)
    If w3& > mh% Then h% = w3& * 100 Else h% = mh% * 100
    Hash2.MSFlexGrid1.ColWidth(oslo%) = h%
    gw& = gw& + h%
    Hash2.Caption = OBJTAB(w1&).obnev
    If InStr(ar$, "J") > 0 Then
      Hash2.MSFlexGrid1.ColAlignment(oslo%) = 6
    Else
      Hash2.MSFlexGrid1.ColAlignment(oslo%) = 1
    End If
    kmtb$(oslo% + 1) = ar$: kmho%(oslo% + 1) = mh%: kmkp%(oslo% + 1) = kp%
    kmesor%(oslo% + 1) = i1&: kmesor%(0) = oslo% + 1
  Next
  keresobj$ = objazon
  Hash2.Show vbModal
  rec$ = rekord$
  'utrec$(w1&) = rekord$
  Unload Hash2
End Sub


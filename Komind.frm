VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Komind 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Kommunikáció"
   ClientHeight    =   4032
   ClientLeft      =   1740
   ClientTop       =   1476
   ClientWidth     =   9924
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4032
   ScaleWidth      =   9924
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List6 
      BackColor       =   &H0032A8E7&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1392
      Left            =   7200
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Tovább (Esc)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2880
      TabIndex        =   10
      Top             =   3240
      Width           =   1385
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.2
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   4140
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mehet"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "A lista indítása"
      Top             =   3360
      Width           =   1332
   End
   Begin VB.ListBox List5 
      Appearance      =   0  'Flat
      BackColor       =   &H001CC9F4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1368
      Left            =   7440
      TabIndex        =   5
      Top             =   1080
      Width           =   2412
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   492
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Kilépés a lista futtatása nélkül"
      Top             =   3360
      Width           =   1212
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BackColor       =   &H0055D7F7&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1944
      Left            =   6600
      TabIndex        =   4
      Top             =   840
      Width           =   2892
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H0091E9FB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1944
      Left            =   5880
      TabIndex        =   3
      Top             =   600
      Width           =   2892
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1752
      Left            =   5160
      TabIndex        =   2
      Top             =   360
      Width           =   2892
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   792
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   2892
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3132
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   5525
      _Version        =   327680
      Rows            =   13
      Cols            =   3
      BackColor       =   16777215
      BackColorBkg    =   4210752
      ScrollBars      =   0
      Appearance      =   0
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H0091E9FB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   5412
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H0091E9FB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   5412
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H0091E9FB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   5412
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H0091E9FB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   5412
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0091E9FB&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H0055D7F7&
      Height          =   252
      Left            =   4440
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   5412
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   336
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   3012
   End
End
Attribute VB_Name = "Komind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- KOMIND form (Lista kérdezõ tábla) kódja
Dim oszlop%, sor%, tovabb%, ok%

Private Sub Command1_Click()
  '--- kilépés mégsem
  kommegsem% = 1
  Call langclos
  Komind.Hide
End Sub
Private Sub Command2_Click()
  '--- kilépés kitöltött adatokkal (Enter)
  kommegsem% = 0
  Call langclos
  Komind.Hide
End Sub

Private Sub Command3_Click()
  tovabb% = 1
  Call Text1_KeyDown(vbKeyEscape, 0)
End Sub

Private Sub Form_Activate()
  '--- aktíválás focus az elsõ adatmezõre, ha van
  If tovabb% = 0 Then
    kommegsem% = 0
    If komadatdb% > 0 Then
      If komt(sor%).komkod = 2 Then
        Label1.Caption = Trim$(MSFlexGrid1.TextMatrix(sor%, 0)) + " " + langform(5)
      Else
        Label1.Caption = Trim$(MSFlexGrid1.TextMatrix(sor%, 0)) + ":"
      End If
      Text1.SetFocus
    End If
  End If
End Sub

Private Sub Form_Load()
  '--- komind form betöltése
  '--- kezdeti megjelenés beálítása
  MSFlexGrid1.Font.Name = "Microsoft Sans Serif"
  MSFlexGrid1.Font.Size = 8
  Call langinit("komind", 2)
  Call szkriptel("komind")
  Komind.Picture = LoadPicture(programutvonal$ + "auwlis1.jpg")
  List1.Visible = False
  List2.Visible = False
  List3.Visible = False
  List4.Visible = False
  List5.Visible = False
  List6.Visible = False
  Command1.Visible = True
  Command2.Visible = False
  kommegsem% = 0
  If komadatdb% > 0 Then
    sor% = 1: oszlop% = 1
    MSFlexGrid1.CellBackColor = QBColor(14)
    Text1.Text = MSFlexGrid1.TextMatrix(1, 1)
    If komt(sor%).komkod = 2 Then
      Label1.Caption = Trim$(MSFlexGrid1.TextMatrix(sor%, 0)) + " " + langform(5)
    Else
      Label1.Caption = Trim$(MSFlexGrid1.TextMatrix(sor%, 0)) + ":"
    End If
  End If
End Sub

Private Sub List1_dblClick()
  '--- elsõ menü választás
  komt(komadatdb% + 1).kommnv = List1.ListIndex + 1
  If kommenudb% > 1 Then
    List2.Visible = True
    List2.SetFocus
  Else
    Command1.Visible = True
    Command2.Visible = True
    Command1.Cancel = True
    Command2.SetFocus
  End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- elsõ menü enter vagy esc
  If KeyCode = vbKeyReturn Then KeyCode = 0: Call List1_dblClick
  If KeyCode = vbKeyEscape Then
    KeyCode = 0
    Call Command1_Click
  End If
End Sub
Private Sub list2_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- második menü enter vagy esc
  If KeyCode = vbKeyReturn Then KeyCode = 0: Call List2_dblClick
  If KeyCode = vbKeyEscape Then
    KeyCode = 0
    Call Command1_Click
  End If
End Sub

Private Sub list3_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- harmadik menü enter vagy esc
  If KeyCode = vbKeyReturn Then KeyCode = 0: Call list3_dblclick
  If KeyCode = vbKeyEscape Then
    KeyCode = 0
    Call Command1_Click
  End If
End Sub

Private Sub list4_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- negyedik menü enter vagy esc
  If KeyCode = vbKeyReturn Then KeyCode = 0: Call list4_dblclick
  If KeyCode = vbKeyEscape Then
    KeyCode = 0
    Call Command1_Click
  End If
End Sub

Private Sub list5_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- ötödik menü enter vagy esc
  If KeyCode = vbKeyReturn Then KeyCode = 0: Call list5_dblclick
  If KeyCode = vbKeyEscape Then
    KeyCode = 0
    Call Command1_Click
  End If
End Sub
Private Sub list6_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- ötödik menü enter vagy esc
  If KeyCode = vbKeyReturn Then KeyCode = 0: Call list6_dblclick
  If KeyCode = vbKeyEscape Then
    KeyCode = 0
    Call Command1_Click
  End If
End Sub

Private Sub List2_dblClick()
  '--- második menü választás
  komt(komadatdb% + 2).kommnv = List2.ListIndex + 1
  If kommenudb% > 2 Then
    List3.Visible = True
    List3.SetFocus
  Else
    Command1.Visible = True
    Command2.Visible = True
    Command1.Cancel = True
    Command2.SetFocus
  End If
End Sub
Private Sub list3_dblclick()
  '--- harmadik menü választás
  komt(komadatdb% + 3).kommnv = List3.ListIndex + 1
  If kommenudb% > 3 Then
    List4.Visible = True
    List4.SetFocus
  Else
    Command1.Visible = True
    Command2.Visible = True
    Command1.Cancel = True
    Command2.SetFocus
  End If
End Sub
Private Sub list4_dblclick()
  '--- negyedik menü választás
  komt(komadatdb% + 4).kommnv = List4.ListIndex + 1
  If kommenudb% > 4 Then
    List5.Visible = True
    List5.SetFocus
  Else
    Command1.Visible = True
    Command2.Visible = True
    Command1.Cancel = True
    Command2.SetFocus
  End If
End Sub
Private Sub list5_dblclick()
  '--- ötödik menü választás
  komt(komadatdb% + 5).kommnv = List4.ListIndex + 1
  If kommenudb% > 5 Then
    List6.Visible = True
    List6.SetFocus
  Else
    Command1.Visible = True
    Command2.Visible = True
    Command1.Cancel = True
    Command2.SetFocus
  End If
  
End Sub
Private Sub list6_dblclick()
  '--- ötödik menü választás
  komt(komadatdb% + 6).kommnv = List6.ListIndex + 1
  Command1.Visible = True
  Command2.Visible = True
  Command1.Cancel = True
  Command2.SetFocus
End Sub


Private Sub MSFlexGrid1_Click()
  Command2.Visible = False
  Command3.Visible = True
  List1.Visible = False
  List2.Visible = False
  List3.Visible = False
  List4.Visible = False
  List5.Visible = False
  List6.Visible = False
  sx% = MSFlexGrid1.Row
  ox% = MSFlexGrid1.Col
  MSFlexGrid1.Col = oszlop%
  MSFlexGrid1.Row = sor%
  MSFlexGrid1.CellBackColor = QBColor(15)
  sor% = sx%: oszlop% = ox%
  If komt(sor%).komkod <> 2 Then oszlop% = 1
  MSFlexGrid1.Col = oszlop%
  MSFlexGrid1.Row = sor%
  If vezcar% <> 2 Then MSFlexGrid1.CellBackColor = QBColor(14)
  If komt(sor%).komkod = 2 Then
    If oszlop% = 1 Then
      Label1.Caption = Trim$(MSFlexGrid1.TextMatrix(sor%, 0)) + " " + langform(5)
    Else
      Label1.Caption = Trim$(MSFlexGrid1.TextMatrix(sor%, 0)) + " " + langform(6)
    End If
  Else
    Label1.Caption = Trim$(MSFlexGrid1.TextMatrix(sor%, 0)) + ":"
  End If
  Text1.Text = MSFlexGrid1.TextMatrix(sor%, oszlop%)
End Sub

Private Sub MSFlexGrid1_gotfocus()
  '--- táblázatra click esetén focus átirányítás text1-re
  If tovabb% = 0 Then
    Command3.Visible = True
    Command1.Cancel = False
    Text1.SetFocus
  Else
    Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
    Text1.SetFocus
  End If
End Sub

Private Sub Text1_Change()
  '--- text változás (gépelés) esetén táblázat szinkronizálása
  MSFlexGrid1.Text = Text1.Text
  Text1.SelStart = Len(Text1.Text) + 1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  '--- billentyû a text1 mezõben
  '--- enter esetén kódvizsgálat
  vezcar% = 1
  Select Case KeyCode
    Case vbKeyX
      '--- Alt+X kapcsolat
      vezcar% = 0
      If Shift And vbAltMask Then
        '--- Alt+X
        KeyCode = 0
        If komt(sor%).komobj <> Space$(4) Then
          objazon1$ = komt(sor%).komobj
          azonosito$ = ""
          Call altx(objazon1$, azonosito$)
          If azonosito$ <> "" Then
            Text1.Text = Trim$(azonosito$)
          End If
        End If
      End If
    Case vbKeyH
      '--- Alt+H kapcsolat
      vezcar% = 0
      If Shift And vbAltMask Then
        '--- Alt+H
        KeyCode = 0
        If komt(sor%).komobj <> Space$(4) Then
          objazon1$ = komt(sor%).komobj
          azonosito$ = ""
          Call alth(objazon1$, azonosito$)
          If azonosito$ <> "" Then
            Text1.Text = Trim$(azonosito$)
          End If
        End If
      End If
    Case vbKeyDown
      '--- következõ sor
      sor% = sor% + 1: If sor% > komadatdb% Then sor% = komadatdb%
      KeyCode = 0
    Case vbKeyUp
      '--- elõzõ sor
      sor% = sor% - 1: If sor% < 1 Then sor% = 1
      KeyCode = 0
    Case vbKeyF12
      KeyCode = 0
      If oszlop% = 2 Then
        Text1.Text = Trim(MSFlexGrid1.TextMatrix(sor%, 1))
      End If
    Case vbKeyRight, vbKeyLeft
      '--- oszlop váltás
      KeyCode = 0
      If oszlop% = 1 Then oszlop% = 2 Else oszlop% = 1
    Case vbKeyReturn
      '--- enter
      '--- kódvizsgálat
      If InStr(UCase$(Trim(komt(sor%).komszov)), UCase$("Megr.iktató")) > 0 Then
        If oszlop% = 1 Then Text1.Text = 1 Else Text1.Text = "9999999"
      End If
      If InStr(komt(sor%).komatr, "D") > 0 Then
        If Len(Trim$(Text1.Text)) = 0 Then
          If programnev = "AUW-FEZR" Then
            If InStr(UCase$(Trim(komt(sor%).komszov)), UCase$(langform(7))) > 0 Then
              If oszlop% = 1 Then
                irec$ = dbxkey("INST", "INST")
                lev& = xval(Mid$(irec$, 404, 4))
                ujle& = lev& + 1
                ujlev$ = Trim(Str(ujle&))
                llev$ = Trim(Str(lev&))
                If langhun% > 1 Then
                  fapp$ = Mid$(irec$, 662, 4)
                  If Trim(fapp$) = "" Then fapp$ = "3112"
                  If fapp$ = "3112" Then
                    Text1.Text = "0101" + Mid$(ujlev$, 3, 2)
                  Else
                    xxzz$ = novdat(Mid$(ujlev$, 3, 2) + Right(fapp$, 2) + Left(fapp$, 2))
                    Text1.Text = datfor(xxzz$)
                  End If
                Else
                  fapp$ = Mid$(irec$, 662, 4): If Trim(fapp$) = "" Then fapp$ = "1231"
                  If fapp$ = "1231" Then
                    Text1.Text = Mid$(ujlev$, 3, 2) + "0101"
                  Else
                    Text1.Text = novdat(Mid$(ujlev$, 3, 2) + fapp$)
                  End If
                End If
              End If
              If oszlop% = 2 Then
                irec$ = dbxkey("INST", "INST")
                lev& = xval(Mid$(irec$, 404, 4))
                ujle& = lev& + 1
                ujlev$ = Trim(Str(ujle&))
                llev$ = Trim(Str(lev& + 2))
                If langhun% > 1 Then
                  fapp$ = Mid$(irec$, 662, 4): If Trim(fapp$) = "" Then fapp$ = "3112"
                  If fapp$ = "3112" Then
                    Text1.Text = fapp$ + Mid$(ujlev$, 3, 2)
                  Else
                    Text1.Text = fapp$ + Mid$(llev$, 3, 2)
                  End If
                Else
                  fapp$ = Mid$(irec$, 662, 4): If Trim(fapp$) = "" Then fapp$ = "1231"
                  If fapp$ = "1231" Then
                    Text1.Text = Mid$(ujlev$, 3, 2) + fapp$
                  Else
                    Text1.Text = Mid$(llev$, 3, 2) + fapp$
                  End If
                End If
              End If
            End If
          Else
            If InStr(UCase$(Trim(komt(sor%).komszov)), UCase$(langform(8))) > 0 Then
              If langhun% > 1 Then
                Text1.Text = "3112" + Mid$(maidatum$, 1, 2)
              Else
                Text1.Text = Mid$(maidatum$, 1, 2) + "1231"
              End If
            End If
            If InStr(UCase$(Trim(komt(sor%).komszov)), UCase$(langform(7))) > 0 Then
              If oszlop% = 1 Then
                If langhun% > 1 Then
                  Text1.Text = "0101" + Mid$(maidatum$, 1, 2)
                Else
                  Text1.Text = Mid$(maidatum$, 1, 2) + "0101"
                End If
              End If
              If oszlop% = 2 Then
                If langhun% > 1 Then
                  Text1.Text = "3112" + Mid$(maidatum$, 1, 2)
                Else
                  Text1.Text = Mid$(maidatum$, 1, 2) + "1231"
                End If
              End If
            End If
          End If
        End If
      End If
      KeyCode = 0
      '--- kodvizsgálat
      '--- teljes mezo ellenorzése
      hiba% = 0
      attr$ = komt(sor%).komatr
      ho% = xval(Left$(attr$, 2))
      If Len(Trim$(Text1.Text)) > ho% Then
        hiba% = 1
        hsz$ = langform(9)
        txx$ = Left$(Text1.Text, ho%)
        Text1.Text = txx$
      End If
      If InStr(attr$, "NZJ") > 0 Then
        txx$ = Right$("000000000000000" + LTrim$(Text1.Text), ho%)
        Text1.Text = txx$
      End If
      If InStr(attr$, "N") > 0 Then
        ell$ = "0123456789"
        If InStr(attr$, "T") > 0 Then ell$ = ell$ + "."
        If InStr(attr$, "-") > 0 Then ell$ = ell$ + "-"
        mit$ = Trim$(Text1.Text)
        If Len(mit$) > 0 Then
          For w1% = 1 To Len(mit$)
            If InStr(ell$, Mid$(mit$, w1%, 1)) = 0 Then hiba% = 1: hsz$ = langform(10): Exit For
          Next
        End If
      End If
      If InStr(attr$, "K") > 0 Then
        If Len(Trim$(Text1.Text)) <> ho% Then hsz$ = langform(11) + Str$(ho%): hiba% = 1
      End If
      If hiba% = 0 Then
        MSFlexGrid1.TextMatrix(sor%, oszlop%) = Text1.Text
        If oszlop% = 1 Then
          If komt(sor%).komkod = 2 Then
             oszlop% = 2
          Else
            sor% = sor% + 1: If sor% > komadatdb% Then sor% = komadatdb%
          End If
        Else
          oszlop% = 1: sor% = sor% + 1
          If sor% > komadatdb% Then sor% = 1
        End If
      Else
        Call mess(hsz$, 2, 0, langform(12), valasz%)
        'MsgBox hsz$, 48, langform(12)
      End If
    Case vbKeyEscape
      '--- Esc
      '--- kitöltöttségvizsgálat
      vezcar% = 2
      KeyCode = 0
      ok% = 1
      For w3% = 1 To komadatdb%
        If InStr(komt(w3%).komatr, "*") > 0 Then
          If Trim$(MSFlexGrid1.TextMatrix(w3%, 1)) = "" Then
            Call mess(Trim$(komt(w3%).komszov) + " " + langform(13), 2, 0, langform(15), valasz%)
            'MsgBox Trim$(komt(w3%).komszov) + " " + langform(13), 48, langform(15)
            ok% = 0: Exit For
          Else
            If komt(w3%).komkod = 2 Then
              If Trim$(MSFlexGrid1.TextMatrix(w3%, 2)) = "" Then
                ok% = 0
                Call mess(Trim$(komt(w3%).komszov) + " " + langform(14), 2, 0, langform(15), valasz%)
                'MsgBox Trim$(komt(w3%).komszov) + " " + langform(14), 48, langform(15)
                Exit For
              End If
            End If
          End If
        End If
      Next
      If ok% = 1 Then
        For w3% = 1 To komadatdb%
          If InStr(komt(w3%).komatr, "D") > 0 And langhun% > 1 Then
            komt(w3%).komtol = datfor(MSFlexGrid1.TextMatrix(w3%, 1))
          Else
            komt(w3%).komtol = MSFlexGrid1.TextMatrix(w3%, 1)
          End If
          If komt(w3%).komkod = 2 Then
            If InStr(komt(w3%).komatr, "D") > 0 And langhun% > 1 Then
              komt(w3%).komig = datfor(MSFlexGrid1.TextMatrix(w3%, 2))
            Else
              komt(w3%).komig = MSFlexGrid1.TextMatrix(w3%, 2)
            End If
          End If
        Next
        If kommenudb% > 0 Then
          Command3.Visible = False
          List1.Visible = True
          List1.SetFocus
        Else
          Command1.Visible = True
          Command2.Visible = True
          Command1.Cancel = True
          Command2.Default = True
        End If
      Else
        'MSFlexGrid1.SetFocus
        Text1.SetFocus
        Exit Sub
      End If
    Case Else
      '--- normál begépelt karakter
      vezcar% = 0
  End Select
  If vezcar% <> 0 Then
    '--- vezérlõ karakter volt
    If komt(sor%).komkod <> 2 Then oszlop% = 1
    MSFlexGrid1.CellBackColor = QBColor(15)
    MSFlexGrid1.Col = oszlop%
    MSFlexGrid1.Row = sor%
    If vezcar% <> 2 Then MSFlexGrid1.CellBackColor = QBColor(14)
    If komt(sor%).komkod = 2 Then
      If oszlop% = 1 Then
        Label1.Caption = Trim$(MSFlexGrid1.TextMatrix(sor%, 0)) + " " + langform(5)
      Else
        Label1.Caption = Trim$(MSFlexGrid1.TextMatrix(sor%, 0)) + " " + langform(6)
      End If
    Else
      Label1.Caption = Trim$(MSFlexGrid1.TextMatrix(sor%, 0)) + ":"
    End If
    Text1.Text = MSFlexGrid1.TextMatrix(sor%, oszlop%)
  End If
End Sub

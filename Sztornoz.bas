Attribute VB_Name = "Sztornoz"
'--- sztornó kezelése
Public sztornovegrehajt%

Public Sub sztornokelt(tev$, eredatum$, bizkelt$, irec$, uzemmod$, sztordat$, sztornoszamlaszam$)
  '--- tev-tárgyév
  '--- eredatum-eredeti biz.kelte
  '--- irec-alap paraméter
  '--- sztordat-sztornozás kelte vagy hiba esetén=üres string
  '--- uzemmod-mûködési mód
  '---   S-számla
  '---   H-helyesbítõ, elõleg
  '---   F-folyószámla
  '---   E-egyéb
  '--- sztornó beállítás
  If Trim(irec$) = "" Then
    irec$ = dbxkey("INST", "INST")
  End If
  maidatum$ = Right(Date$, 2) + Left(Date$, 2) + Mid(Date$, 4, 2)
  tevx$ = Mid$(irec$, 400, 4)
  If langhun% > 1 Then
    Sztorno.Text3.Text = datfor(eredatum$)
    Sztorno.Text3.Locked = True
    Sztorno.Text4.Text = datfor(maidatum$)
  Else
    Sztorno.Text3.Text = eredatum$
    Sztorno.Text3.Locked = True
    Sztorno.Text4.Text = maidatum$
  End If
  Select Case uzemmod$
    Case "F"
      If langhun% > 1 Then
        Sztorno.Text1.Text = datfor(eredatum$)
      Else
        Sztorno.Text1.Text = eredatum$
      End If
      Sztorno.Option1.Value = True
      Sztorno.Option2.Visible = True
      Sztorno.Label2.Visible = True
      Sztorno.Text2.Visible = True
    Case "S", "H"
      If langhun% > 1 Then
        Sztorno.Text1.Text = datfor(maidatum$)
      Else
        Sztorno.Text1.Text = maidatum$
      End If
      Sztorno.Option1.Value = True
      Sztorno.Option2.Visible = True
      Sztorno.Label2.Visible = False
      Sztorno.Text2.Visible = False
    Case "E"
      If langhun% > 1 Then
        Sztorno.Text1.Text = datfor(eredatum$)
      Else
        Sztorno.Text1.Text = eredatum$
      End If
      Sztorno.Option2.Visible = False
      Sztorno.Label2.Visible = False
      Sztorno.Text2.Visible = False
    Case Else
  End Select
  Do
    ellhiba% = 0
    Sztorno.Show vbModal
    If sztornovegrehajt = 1 Then
      If langhun% > 1 Then
        sztordat$ = datfor(Trim(Sztorno.Text1.Text))
      Else
        sztordat$ = Trim(Sztorno.Text1.Text)
      End If
      sztornoszamlaszam$ = Trim(Sztorno.Text2.Text)
      If jodatum(sztordat$) = 0 Then
        Call mess(langmodul(137) + " " + langmodul(133) + "!", 1, 0, langmodul(138), valasz%)
        'MsgBox langmodul(137) + " " + langmodul(133) + "!", 48, langmodul(138)
        'MsgBox "Hibás sztornó dátum", 48, "Sztornó hiba"
        ellhiba% = 1
      Else
        ' Eszi - 2009.02.13
        'If dtm(sztordat$) < dtm(eredatum$) Then
        '  ellhiba% = 1
        '  Call mess(langmodul(133) + " < " + langmodul(134) + "!", 1, 0, langmodul(138), valasz%)
        '  'MsgBox langmodul(133) + " < " + langmodul(134), 48, langmodul(138)
        '  '"Sztornó kelte < könyvelés kelte", 48, langmodul(138)
        'Else
          If dtm(sztordat$) > dtm(maidatum$) Then
            Call mess(langmodul(133) + " > " + langmodul(135) + "!", 1, 0, langmodul(138), valasz%)
            'MsgBox langmodul(133) + " > " + langmodul(135) + "!", 48, langmodul(138)
            'MsgBox "Sztornó kelte > mai dátum!", 48, "Sztornó hiba"
            ellhiba% = 1
          Else
            If Sztorno.Option2.Value = True And uzemmod$ = "F" Then
              If Trim(sztornoszamlaszam$) = "" Then
                ellhiba% = 1
                Call mess(langmodul(139), 1, 0, langmodul(138), valasz%)
                'MsgBox langmodul(139), 48, langmodul(138)
                'MsgBox "Hibás teljesítés esetén sztornószámla száma kötelezõ", 48, "Sztornó hiba"
              End If
            End If
          End If
        'End If
      End If
    End If
  Loop While ellhiba% = 1 And sztornovegrehajt = 1
  If sztornovegrehajt = 1 Then
    If uzemmod$ <> "E" And Sztorno.Option2.Value = True Then
      If sztordat$ <> maidatum$ Then
        Call mess(langmodul(140) + " " + langmodul(142), 5, 3, langmodul(136), valasz%)
        If valasz% = 0 Then sztornovegrehajt = 0
        'respons = MsgBox(langmodul(140) + " " + langmodul(142), vbYesNo, langmodul(136))
        'If respons <> vbYes Then sztornovegrehajt = 0
      End If
    End If
    If sztornovegrehajt = 1 And Left$(sztordat$, 2) <> Mid$(tevx$, 3, 2) Then
      Call mess(langmodul(141) + " " + langmodul(142), 5, 3, langmodul(136), valasz%)
      If valasz% = 0 Then sztornovegrehajt = 0
      'respons = MsgBox(langmodul(141) + " " + langmodul(142), vbYesNo, langmodul(136))
      'If respons <> vbYes Then sztornovegrehajt = 0
    End If
  End If
  If sztornovegrehajt = 0 Then
    sztordat$ = ""
  End If
  '--- sztornó beállítás ellenõrzésének vége
End Sub

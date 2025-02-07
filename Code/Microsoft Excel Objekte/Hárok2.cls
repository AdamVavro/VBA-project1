VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hárok2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''Procedura provedna pøed kliknutím na hypertextový odkaz.
'Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
'
''    MsgBox ("Otvaram hyperlink, skryjem riadok vzorcov")
'
'    'Ak je pás s nástrojmi pripnutý, tak ho zbalí
'        If Application.CommandBars("Ribbon").Height > 100 Then 'ActiveWindow.ToggleRibbon
''           MsgBox ("Pás je pripnutý")
'        Application.CommandBars.ExecuteMso "MinimizeRibbon"
'        End If
'
''       MsgBox ("Otvaram hyperlink, skryjem riadok vzorcov, pás nastrojov by už mal by skrytý")
'    'Skryje riadok vzorcov
'        Application.DisplayFormulaBar = False
'
''        Application.Wait Now + TimeValue("00:00:02")
'
'End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

'ActiveSheet.Unprotect Password:="Lis.0123"


''''''''''
'FUNGUJE
'Ked zmenená bunka patri do stlpca 7 a bunka nad òou ma farbu RGB(210, 245, 45) _
do bunky nad zmenenou bunkou sa zapíše hodnota zmenenej bunky-1 _
a do komentára sa uloží prvé štvorèíslie nástroja

    If Target.Column = 7 Then
        BunkaNadTarget = (Target.Offset(-1).Address)
        If Range(BunkaNadTarget).Interior.Color = RGB(210, 245, 45) Then
    '       ActiveSheet.Unprotect Password:="Lis.0123"
           Application.EnableEvents = False
    '        MsgBox (Target.Address)
    '        MsgBox (Range(BunkaNadTarget).Value - 1)
            Range(BunkaNadTarget).Value = Target.Value - 1
            Application.EnableEvents = True
            
    '       Komentár zaciatok
                Dim HodnotaMID As String
                HodnotaMID = Mid(Target.Value, 1, 4)
    
                If Range(BunkaNadTarget).Comment Is Nothing Then
                    Range(BunkaNadTarget).AddComment Text:=HodnotaMID
                    Range(BunkaNadTarget).Comment.Shape.TextFrame.AutoSize = True
                Else
                'MsgBox (Target.Address)
                    Range(BunkaNadTarget).Comment.Text Text:=HodnotaMID
                End If
    '       Komentár koniec
    '            ActiveSheet.Protect Password:="Lis.0123"
                
        End If
    End If
''''''''''
'FUNGUJE
'Po vložení údaja do bunky v 10. ståpci sa údaj uloží do komentára danej bunky, _
po dalšej zmene tejto bunky sa do komentára doplní nový údaj, prièom predchádzajuci tam zostane tiež _
(Komentár je tak v podstate história parametrov)

'DATUM POSLEDNEJ AKTUALIZACIE
    If Target.Column = 6 Then
        If Target.Comment Is Nothing Then
            Target.AddComment Text:=Now() & " / " & Target.Text & vbLf
            Target.Comment.Shape.TextFrame.AutoSize = True
            Target.Cells.Interior.Color = RGB(64, 64, 64)
            
        Else
            'MsgBox (Target.Address)
            Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & vbLf
        End If
    End If


'TlakBaran = Columns(10)

'BARAN PRACOVNE TLAKY
    If Target.Column = 15 Then
        If Target.Comment Is Nothing Then
            Target.AddComment Text:=Now() & " / " & Target.Text & "kN" & vbLf
            Target.Comment.Shape.TextFrame.AutoSize = True
        Else
            'MsgBox (Target.Address)
            Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "kN" & vbLf
        End If
    End If
    
'HORNY PRIDRZIAVAC PRACOVNE TLAKY
    If Target.Column = 16 Then
        If Target.Comment Is Nothing Then
            Target.AddComment Text:=Now() & " / " & Target.Text & "kN" & vbLf
            Target.Comment.Shape.TextFrame.AutoSize = True
        Else
            'MsgBox (Target.Address)
            Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "kN" & vbLf
        End If
    End If
    
'DOLNY PRIDRZIAVAC PRACOVNE TLAKY
    If Target.Column = 17 Then
        If Target.Comment Is Nothing Then
            Target.AddComment Text:=Now() & " / " & Target.Text & "kN" & vbLf
            Target.Comment.Shape.TextFrame.AutoSize = True
        Else
            'MsgBox (Target.Address)
            Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "kN" & vbLf
        End If
    End If
    
'BARAN TUSIROVACIE TLAKY
    If Target.Column = 20 Then
        If Target.Comment Is Nothing Then
            Target.AddComment Text:=Now() & " / " & Target.Text & "kN" & vbLf
            Target.Comment.Shape.TextFrame.AutoSize = True
        Else
            'MsgBox (Target.Address)
            Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "kN" & vbLf
        End If
    End If
    
'HORNY PRIDRZIAVAC TUSIROVACIE TLAKY
    If Target.Column = 21 Then
        If Target.Comment Is Nothing Then
            Target.AddComment Text:=Now() & " / " & Target.Text & "kN" & vbLf
            Target.Comment.Shape.TextFrame.AutoSize = True
        Else
            'MsgBox (Target.Address)
            Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "kN" & vbLf
        End If
    End If
    
'DOLNY PRIDRZIAVAC TUSIROVACIE TLAKY
    If Target.Column = 22 Then
        If Target.Comment Is Nothing Then
            Target.AddComment Text:=Now() & " / " & Target.Text & "kN" & vbLf
            Target.Comment.Shape.TextFrame.AutoSize = True
        Else
            'MsgBox (Target.Address)
            Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "kN" & vbLf
        End If
    End If
    
'BRZDA
    If Target.Column = 23 Then
        If Target.Comment Is Nothing Then
            Target.AddComment Text:=Now() & " / " & Target.Text & "mm" & vbLf
            Target.Comment.Shape.TextFrame.AutoSize = True
        Else
            'MsgBox (Target.Address)
            Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "mm" & vbLf
        End If
    End If
    
'ÈAPY NAD STOL
    If Target.Column = 24 Then
        If Target.Comment Is Nothing Then
            Target.AddComment Text:=Now() & " / " & Target.Text & "mm" & vbLf
            Target.Comment.Shape.TextFrame.AutoSize = True
        Else
            'MsgBox (Target.Address)
            Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "mm" & vbLf
        End If
    End If

'ActiveSheet.Protect Password:="Lis.0123"

End Sub

'
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'
'    If Target.Column = 7 Then
'
'          MsgBox ("Otvaram hyperlink, skryjem riadok vzorcov")
'
'    'Ak je pás s nástrojmi pripnutý, tak ho zbalí
'        If Application.CommandBars("Ribbon").Height > 100 Then 'ActiveWindow.ToggleRibbon
''           MsgBox ("Pás je pripnutý")
'        Application.CommandBars.ExecuteMso "MinimizeRibbon"
'        End If
'
'    'Skryje riadok vzorcov
'        Application.DisplayFormulaBar = False
''
''        Application.Wait Now + TimeValue("00:00:02")
'    End If
'
'End Sub

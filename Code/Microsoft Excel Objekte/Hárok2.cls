VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "H�rok2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''Procedura provedna p�ed kliknut�m na hypertextov� odkaz.
'Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
'
''    MsgBox ("Otvaram hyperlink, skryjem riadok vzorcov")
'
'    'Ak je p�s s n�strojmi pripnut�, tak ho zbal�
'        If Application.CommandBars("Ribbon").Height > 100 Then 'ActiveWindow.ToggleRibbon
''           MsgBox ("P�s je pripnut�")
'        Application.CommandBars.ExecuteMso "MinimizeRibbon"
'        End If
'
''       MsgBox ("Otvaram hyperlink, skryjem riadok vzorcov, p�s nastrojov by u� mal by� skryt�")
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
'Ked zmenen� bunka patri do stlpca 7 a bunka nad �ou ma farbu RGB(210, 245, 45) _
do bunky nad zmenenou bunkou sa zap�e hodnota zmenenej bunky-1 _
a do koment�ra sa ulo�� prv� �tvor��slie n�stroja

    If Target.Column = 7 Then
        BunkaNadTarget = (Target.Offset(-1).Address)
        If Range(BunkaNadTarget).Interior.Color = RGB(210, 245, 45) Then
    '       ActiveSheet.Unprotect Password:="Lis.0123"
           Application.EnableEvents = False
    '        MsgBox (Target.Address)
    '        MsgBox (Range(BunkaNadTarget).Value - 1)
            Range(BunkaNadTarget).Value = Target.Value - 1
            Application.EnableEvents = True
            
    '       Koment�r zaciatok
                Dim HodnotaMID As String
                HodnotaMID = Mid(Target.Value, 1, 4)
    
                If Range(BunkaNadTarget).Comment Is Nothing Then
                    Range(BunkaNadTarget).AddComment Text:=HodnotaMID
                    Range(BunkaNadTarget).Comment.Shape.TextFrame.AutoSize = True
                Else
                'MsgBox (Target.Address)
                    Range(BunkaNadTarget).Comment.Text Text:=HodnotaMID
                End If
    '       Koment�r koniec
    '            ActiveSheet.Protect Password:="Lis.0123"
                
        End If
    End If
''''''''''
'FUNGUJE
'Po vlo�en� �daja do bunky v 10. st�pci sa �daj ulo�� do koment�ra danej bunky, _
po dal�ej zmene tejto bunky sa do koment�ra dopln� nov� �daj, pri�om predch�dzajuci tam zostane tie� _
(Koment�r je tak v podstate hist�ria parametrov)

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
    
'�APY NAD STOL
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
'    'Ak je p�s s n�strojmi pripnut�, tak ho zbal�
'        If Application.CommandBars("Ribbon").Height > 100 Then 'ActiveWindow.ToggleRibbon
''           MsgBox ("P�s je pripnut�")
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

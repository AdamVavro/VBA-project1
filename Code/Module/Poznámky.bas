Attribute VB_Name = "Pozn·mky"
Sub VyplnLis1Zlta() 'Vypln vybranej oblasti zmeni na ûlt˙
    
   Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
   
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'Zlta
    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
        .Color = 65535
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub VyplnLis2Cervena() 'Vypln vybranej oblasti zmeni na Ëerven˙
    
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'Cervena
    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
        .Color = 255
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub VyplnLis3Modra() 'Vypln vybranej oblasti zmeni na modr˙

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'Modr·
    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
        .Color = 15773696
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"

End Sub

Sub VyplnLis4Fialova() 'Vypln vybranej oblasti zmeni na fialov˙

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'Fialov·
    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
        .Color = 16751052
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub BezVyplne() 'Vypln vybranej oblasti zmeni na "bez v˝plne"
      
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'Bez podfarbenia
    With Selection.Interior
        .Pattern = xlNone
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub ViditelnosùTlacidielVyplneVPoznamkach()
Attribute ViditelnosùTlacidielVyplneVPoznamkach.VB_ProcData.VB_Invoke_Func = " \n14"
'TlaËÌdl· pre zmenu farby v˝plne v pozn·mkach bud˙ viditeænÈ

    Worksheets("AIO_Plan").CommandButton8.Visible = True
    Worksheets("AIO_Plan").CommandButton9.Visible = True
    Worksheets("AIO_Plan").CommandButton10.Visible = True
    Worksheets("AIO_Plan").CommandButton11.Visible = True
    Worksheets("AIO_Plan").CommandButton12.Visible = True
    Worksheets("AIO_Plan").CommandButton13.Visible = True
    Worksheets("AIO_Plan").CommandButton14.Visible = True
    Worksheets("AIO_Plan").CommandButton15.Visible = True
    Worksheets("AIO_Plan").CommandButton16.Visible = True
    Worksheets("AIO_Plan").CommandButton17.Visible = True
    Worksheets("AIO_Plan").CommandButton18.Visible = True

End Sub
Sub NeviditelnostTlacidielVyplneVPozn·mkach()
Attribute NeviditelnostTlacidielVyplneVPozn·mkach.VB_ProcData.VB_Invoke_Func = " \n14"
'TlaËÌdl· pre zmenu farby v˝plne v pozn·mkach bud˙ neviditeænÈ

    Worksheets("AIO_Plan").CommandButton8.Visible = False
    Worksheets("AIO_Plan").CommandButton9.Visible = False
    Worksheets("AIO_Plan").CommandButton10.Visible = False
    Worksheets("AIO_Plan").CommandButton11.Visible = False
    Worksheets("AIO_Plan").CommandButton12.Visible = False
    Worksheets("AIO_Plan").CommandButton13.Visible = False
    Worksheets("AIO_Plan").CommandButton14.Visible = False
    Worksheets("AIO_Plan").CommandButton15.Visible = False
    Worksheets("AIO_Plan").CommandButton16.Visible = False
    Worksheets("AIO_Plan").CommandButton17.Visible = False
    Worksheets("AIO_Plan").CommandButton18.Visible = False

End Sub

Sub CervenePÌsmo() 'PÌsmo vybranej oblasti zmeni na ËervenÈ
      
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'»ervenÈ
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub
Sub CiernePÌsmo() 'PÌsmo vybranej oblasti zmeni na Ëierne
      
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'»ierne
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub CerveneAleboCierne() 'PÌsmo vybranej oblasti zmeni na Ëervene alebo Ëierne podla podmienky
      
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'»ervene alebo Ëierne
'    On Error Resume Next
    If Selection.Font.Color = 0 Then
'        MsgBox ("Black")
        With Selection.Font
        .Color = -16776961 '255
        .TintAndShade = 0
        End With
    Else: 'MsgBox ("Red")
        With Selection.Font
        .ThemeColor = xlThemeColorLight1 '0
        .TintAndShade = 0
        End With
    End If
'    On Error GoTo 0

    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub VyplnCierna() 'Vypln vybranej oblasti zmeni na Ëiernu

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'»ierna
    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
        .Color = 0
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub VyplnZltoCervena() 'Vypln vybranej oblasti zmeni na ûltoËerven˙

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'ZltoCervena
    With Selection.Interior
        .Pattern = xlPatternLinearGradient '4000 ËÌselnÈ vyjadrenie
'        .Gradient.Degree = 0
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .Color = 65535
'        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .Color = 255
'        .TintAndShade = 0
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub ZarovnanieTextuNaStred() 'Text vo vybranej oblasti zarovn· na stred

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'Zarovnanie textu na stred
    With Selection
        .HorizontalAlignment = xlCenter '-4108 ËÌselnÈ vyjadrenie
        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub ZarovnanieTextuVlavo() 'Text vo vybranej oblasti zarovn· vlavo

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'Zarovnanie textu vlavo
    With Selection
        .HorizontalAlignment = xlLeft '-4131 ËÌselnÈ vyjadrenie
        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub ZarovnanieTextuVpravo() 'Text vo vybranej oblasti zarovn· vpravo

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    Dim PznmkVStrd As Range
    Set PznmkVStrd = Range("$B$15:$AN$25,$I$14")
    
    Intersect(Cll, PznmkVStrd).Select
    
'Zarovnanie textu vpravo
    With Selection
        .HorizontalAlignment = xlRight '-4152 ËÌselnÈ vyjadrenie
        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
    End With
    
    Intersect(Cll, PznmkVStrd).Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub PoznamkyVStredeBezPodfarbenia()

    Range("$B$15:$AN$25,$I$14").Select
    Call BezVyplne

End Sub
'SKOPIRUJE FARBU VYPLENE, FONT,ZAROVNANIE
Sub CopyInteriorColorFontHorizontalAlignmentOfCellInNotes()

    IC = Range("B15").Interior.Color
    FC = Range("B15").Font.Color
    HA = Range("B15").HorizontalAlignment
    IP = Range("B15").Interior.Pattern


'    MsgBox ("Interior.Color: " & Range("B15").Interior.Color) 'OK
'    MsgBox ("Font.Color: " & Range("B15").Font.Color) 'OK
'    MsgBox ("HorizontalAlignment: " & Range("B15").HorizontalAlignment) 'OK
'    MsgBox ("Iterior.Pattern: " & Range("B15").Interior.Pattern) 'OK ' _

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    If IP <> 4000 Then
       MsgBox ("BeûÌ If")
        ActiveCell.Interior.Color = IC
        ActiveCell.Font.Color = FC
        ActiveCell.HorizontalAlignment = HA
    Else:
        MsgBox ("BeûÌ Else")
        ActiveCell.Font.Color = FC
        ActiveCell.HorizontalAlignment = HA
        Call VyplnZltoCervena
    End If
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"

End Sub
'SKOPIRUJE FARBU VYPLENE, FONT,ZAROVNANIE
Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach()

    IC = KopirovanaPoznamka.Interior.Color
    FC = KopirovanaPoznamka.Font.Color
    HA = KopirovanaPoznamka.HorizontalAlignment
    IP = KopirovanaPoznamka.Interior.Pattern


'    MsgBox ("Interior.Color: " & Range("B15").Interior.Color) 'OK
'    MsgBox ("Font.Color: " & Range("B15").Font.Color) 'OK
'    MsgBox ("HorizontalAlignment: " & Range("B15").HorizontalAlignment) 'OK
'    MsgBox ("Iterior.Pattern: " & Range("B15").Interior.Pattern) 'OK ' _

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    If IP <> 4000 Then
       MsgBox ("BeûÌ If")
        CielovaPoznamka.Interior.Color = IC
        CielovaPoznamka.Font.Color = FC
        CielovaPoznamka.HorizontalAlignment = HA
    Else:
        MsgBox ("BeûÌ Else")
        CielovaPoznamka.Font.Color = FC
        CielovaPoznamka.HorizontalAlignment = HA
        
        With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.ColorStops.Clear
        End With
        With Selection.Interior.Gradient.ColorStops.Add(0)
            .Color = 65535
        End With
        With Selection.Interior.Gradient.ColorStops.Add(1)
            .Color = 255
        End With
    End If
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"

End Sub


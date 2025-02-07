Attribute VB_Name = "TlËn»p_Cntr»p_VlnMst"
Sub CentrovaciCap()
'
' X Makro
' CentrovacÌ Ëap
'
' Kl·vesov· skratka: Ctrl+Shift+X
'
'K je hodnota aktÌvnej bunky rovna hodnete bunky "U48" alebo "AH48"
    'skopÌruje bunku "F48" a prilepÌ na aktÌvnu bunku
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti CelyStol
    Dim Cll As Range
    Set Cll = Selection

    Dim RasterStola As Range
    Set RasterStola = Range("$E$34:$AK$48")

    Intersect(Cll, RasterStola).Select

    If (ActiveCell.Value) = Range("B30") Or (ActiveCell.Value) = Range("B31") Then
        Range("B29").Copy
'        Worksheets("AIO_Plan").Paste
        Selection.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If
        
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub
Sub TlacnyCap()
'
' O Makro
' TlaËn˝ Ëap
'
' Kl·vesov· skratka: Ctrl+Shift+O
'
'K je hodnota aktÌvnej bunky rovna hodnete bunky "F48" alebo "AH48"
    'skopÌruje bunku "F48" a prilepÌ na aktÌvnu bunku
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti CelyStol
    Dim Cll As Range
    Set Cll = Selection

    Dim RasterStola As Range
    Set RasterStola = Range("$E$34:$AK$48")

    Intersect(Cll, RasterStola).Select
        
    If (ActiveCell.Value) = Range("B29") Or (ActiveCell.Value) = Range("B31") Then
        Range("B30").Copy
'        Worksheets("AIO_Plan").Paste
        Selection.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If
        
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub
Sub VolneMiesto()
'
' PlusSPodmienkou Makro
'
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
    Application.ScreenUpdating = False   'vypne prekreslovanie obrazovky, t˝m sa makro zr˝chli
           
'Ak sa hodnota bunky/buniek vo v˝bere rovn· hodnote buky "U48" alebo "F48"
      'prilepÌ skopÌrovanu bunku
      'a preverÌ bunky v rozsahu "StredStola"
      'ak je v bunke "+" a nem· nastavenÈ TuËnÈ pÌsmo,
      'tak nastavÌ bunke parametre.
      
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti CelyStol
    Dim Cll As Range
    Set Cll = Selection
      
    Dim RasterStola As Range
    Set RasterStola = Range("$E$34:$AK$48")

    Intersect(Cll, RasterStola).Select
    
    For Each cell In Cll
        If cell.Value = Range("B29") Or cell.Value = Range("B30") Then
            Range("B31").Copy
'            Worksheets("AIO_Plan").Paste
            Selection.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
        End If
    Next cell
    '
    
    Dim rng As Range
    Set rng = Range("StredStola")
    For Each cell In rng
        If (cell.Value) = "+" And cell.Font.Bold = False Then
            cell.Font.Bold = True
            cell.Font.Name = "PorscheNextTT"
            cell.Font.Size = 14
            cell.Font.Color = RGB(0, 0, 0)
            cell.HorizontalAlignment = xlCenter
            cell.VerticalAlignment = xlCenter
        End If
    Next cell
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
            
    Application.ScreenUpdating = True

'Funguje
End Sub

Sub ViditelnosùTlacidielPreRasterStola()
'TlaËÌdl· pre zmenu rastra stola bud˙ viditeænÈ

    Worksheets("AIO_Plan").CommandButton3.Visible = True
    Worksheets("AIO_Plan").CommandButton4.Visible = True
    Worksheets("AIO_Plan").CommandButton5.Visible = True
    Worksheets("AIO_Plan").CommandButton6.Visible = True

End Sub

Sub NeviditelnostTlacidielPreRasterStola()
'TlaËÌdl· pre zmenu rastra stola bud˙ neviditeænÈ

   On Error Resume Next
        Worksheets("AIO_Plan").CommandButton3.Visible = False
        Worksheets("AIO_Plan").CommandButton4.Visible = False
        Worksheets("AIO_Plan").CommandButton5.Visible = False
        Worksheets("AIO_Plan").CommandButton6.Visible = False
    On Error GoTo 0

End Sub
Sub CervenAleboCierneCentrovanie() 'Symbol centrovania zmeni na Ëervene alebo Ëierne podla podmienky
      
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'Vyberie priesecnÌk Oblasti v˝beru a Oblasti Pozn·mkyVStrede
    Dim Cll As Range
    Set Cll = Selection
    
    If Cll.Value = Range("B29").Value Then
'        MsgBox ("CentrovacÌ Ëap")
    
        Dim RasterStola As Range
        Set RasterStola = Range("$E$34:$AK$48")
    
        Intersect(Cll, RasterStola).Select
    
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
    End If
    
        Intersect(Cll, RasterStola).Select
        
        Call SpocitaCerveneCentrovacieCapy
        
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub



'SPOCITA CERVENE CENTROVACIE CAPY
Sub SpocitaCerveneCentrovacieCapy()

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"

    'Declare variables
    
    Dim TableRng, LookupFillColor As Range
    On Error Resume Next
    'Set range to variables
    Set TableRng = Range("E34:AK48")
    Set LookupFillColor = Range("AM29")
    'Set a input box for result output
    Set OutputRng = Range("AN29") 'Application.InputBox("select a cell:", "ExcelDemy", Selection.Address, , , , , 8)
    If OutputRng Is Nothing Then Exit Sub
    
    Dim rng As Range
    
        For Each rng In TableRng
            If rng.Font.Color = LookupFillColor.Font.Color And rng.Value = Range("AM29").Value Then
                CountColour = CountColour + 1
            End If
        Next
    
    Range("AN29").Value = CountColour
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

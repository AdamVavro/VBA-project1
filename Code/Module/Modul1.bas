Attribute VB_Name = "Modul1"
Sub DalsieTlacitka()
Attribute DalsieTlacitka.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DalsieTlacitka Makro
'

'

    Range("B16:AN16").Select
    'Èierna výplò
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    Range("B17:AN17").Select
    'Zarovnanie na stred
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    'Zarovnanie dolava
    Range("B17:AN17").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    'Zarovnanie doprava
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    'žltoèervená výplò
    Range("B18:AN18").Select
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 0
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .Color = 65535
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .Color = 255
        .TintAndShade = 0
    End With
End Sub

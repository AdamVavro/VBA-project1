Attribute VB_Name = "Orámovanie"
Sub Drážky()
Attribute Drážky.VB_ProcData.VB_Invoke_Func = " \n14"
' ORAMUJE DRAZKY STOLA
' Drážky Makro
'

'
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
   Range("H34:I48,L34:M48,P34:Q48,T34:V48,Y34:Z48,AC34:AD48,AG34:AH48").Select
    Range("AG34").Activate
    
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
'    Selection.Borders(xlEdgeTop).LineStyle = xlNone
'    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
'    Selection.Borders(xlInsideVertical).LineStyle = xlNone
'    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("S1").Select
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"

End Sub

Attribute VB_Name = "NaVysku_NaSirku"
Sub FormatNaVysku()
'
' FormatNaVysku Makro(poznámky k nástroju dole)
'

'
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
    Application.ScreenUpdating = False
    
    Columns("AP:BO").Select
'    Range("AP5").Activate
    Selection.EntireColumn.Hidden = True
    Rows("51:63").Select
    Selection.EntireRow.Hidden = False
'    ActiveWindow.SmallScroll Down:=12
    Range("A1:AO63").Select
'    Range("AO63").Activate
    Worksheets("AIO_Plan").PageSetup.PrintArea = "$A$1:$AO$63"
        
    Application.ScreenUpdating = True
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub
Sub FormatNaSirku()
'
' FormatNaSirku Makro(poznámky k nástroju dole)
'

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
    Application.ScreenUpdating = False
    
    Rows("51:63").Select
    Selection.EntireRow.Hidden = True
'    ActiveWindow.SmallScroll Down:=-42
    Columns("AP:BO").Select
'    Range("AO5").Activate
    Selection.EntireColumn.Hidden = False
    Range("A1:BO50").Select
'    Range("BO50").Activate
    Worksheets("AIO_Plan").PageSetup.PrintArea = "$A$1:$BO$50"
    
    Application.ScreenUpdating = True
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"

End Sub

Sub FormatNaVyskuPoznamkyVStrede()
'
' FormatNaVysku Makro(poznámky k nástroju dole)
'

'
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
    Application.ScreenUpdating = False
    
    Columns("AP:BO").Select
'    Range("AP5").Activate
    Selection.EntireColumn.Hidden = True
    Rows("14:26").Select
    Selection.EntireRow.Hidden = False
'    ActiveWindow.SmallScroll Down:=12
    Range("A1:AO50").Select
'    Range("AO63").Activate
    Worksheets("AIO_Plan").PageSetup.PrintArea = "$A$1:$AO$50"
        
    Application.ScreenUpdating = True
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub
Sub FormatNaSirkuPoznamkyVStrede()
'
' FormatNaSirku Makro(poznámky k nástroju dole)
'

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
    Application.ScreenUpdating = False
    
    Rows("14:26").Select
    Selection.EntireRow.Hidden = True
'    ActiveWindow.SmallScroll Down:=-42
    Columns("AP:BO").Select
'    Range("AO5").Activate
    Selection.EntireColumn.Hidden = False
    Range("A1:BO50").Select
'    Range("BO50").Activate
    Worksheets("AIO_Plan").PageSetup.PrintArea = "$A$1:$BO$50"
    
    Application.ScreenUpdating = True
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"

End Sub

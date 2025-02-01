Attribute VB_Name = "Pomocný"
 Sub VerziaPlanuUpinania()
    
    Dim HodnotaMID As String
    HodnotaMID = Mid(Range("A64").Value, 1, 3)
    
    MsgBox HodnotaMID
    
    If HodnotaMID = "F77" Then
        MsgBox "Nový plán upínania"
    Else
        MsgBox "Aktualizovaný plán upínania"
    End If
 
 End Sub
 
Sub HorizontalAlignment()

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    ActiveCell.HorizontalAlignment = -4108
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub
Sub OnlyInteriorColor()

'    Range("B15").Interior.Color
    MsgBox ("Výplò: " & Range("B15").Interior.Color) 'OK
    MsgBox ("Font: " & Range("B15").Font.Color) 'OK
    MsgBox ("Zarovnanie: " & Range("B15").HorizontalAlignment) 'OK
    MsgBox ("Iterior.Pattern: " & Range("B15").Interior.Pattern) 'OK ' _
     & "," & Range("B15").Interior.Gradient.ColorStops.Clear _
     & "," & Range("B15").Interior.Gradient.ColorStops.Add(0).Color _
     & "," & Range("B15").Interior.Gradient.ColorStops.Add(1).Color)
'    MsgBox ("Interior.Gradient.ColorStops.Clear: " & Range("B15").Interior.Gradient.ColorStops.Clear)
'    MsgBox ("Interior.Gradient.ColorStops.Add(0).Color: " & Range("B15").Interior.Gradient.ColorStops.Add(0).Color)
'    MsgBox ("Interior.Gradient.ColorStops.Add(1).Color: " & Range("B15").Interior.Gradient.ColorStops.Add(1).Color)
'    With Selection.Interior
'        .Pattern = xlPatternLinearGradient
''        .Gradient.Degree = 0
'        .Gradient.ColorStops.Clear
'    End With
'    With Selection.Interior.Gradient.ColorStops.Add(0)
'        .Color = 65535
''        .TintAndShade = 0
'    End With
'    With Selection.Interior.Gradient.ColorStops.Add(1)
'        .Color = 255
''        .TintAndShade = 0
'    End With
    
End Sub


Sub riadok8H()

    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster8HoreZaèiatok & ":" & AdresaOblastNajdiRaster8HoreKoniec).Select

End Sub
Sub KopirovavnieRastra()
'
' KopirovavnieRastra Makro
'

'
    
    
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    Windows("AIO_Data").Activate
    Range("BS492:CY492").Select
    Selection.Copy
    Windows( _
        "F..._Plán upínania do lisov PWS_KT16_Import udajov z parametre nastrojov.xlsm" _
        ).Activate
    Range("E34:AK34").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub

Sub PocetCapov()

    If Range("AN28").Value = "" Or Range("AN28").Value = "0" Then
        MsgBox ("Nekopirujem raster")
    Else
        MsgBox ("Kopirujem raster")
    End If
    
End Sub
Sub ImportKomentarov()
Attribute ImportKomentarov.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ImportKomentarov Makro
'

'
   
    Range("O492").Select
    Selection.Copy
    
    Windows( _
        "F..._Plán upínania do lisov PWS_KT16_Import udajov z parametre nastrojov.xlsm" _
        ).Activate
    Range("S10:AM9").Select
    ActiveSheet.Unprotect
    
    
    Windows( _
        "F..._Plán upínania do lisov PWS_KT16_Import udajov z parametre nastrojov.xlsm" _
        ).Activate
    Selection.PasteSpecial Paste:=xlPasteComments, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

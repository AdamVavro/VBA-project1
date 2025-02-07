VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "H·rok1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public VyËistitPlanUpinania_Running As Boolean 'Premenn· v Sub VyËistitPlanUpinania_Running(Private Sub CommandButton2_Click)
'Public VerziaPU As Variant 'Premenn· v Sub ULOZIT PLAN UPINANIA(Private Sub CommandButton7_Click)



'TLACITKO "Otvoriù formul·r"
Private Sub CommandButton1_Click()
'''''
'''''    Application.WindowState = xlNormal
'''''        Application.Left = 226
'''''        Application.Top = 1
'''''        Application.Width = 686 '(976)
'''''        Application.Height = 870
'''''
'''''    Application.WindowState = xlMinimized
    
    UserForm1.Show
    
'    If UserForm1.TextBox2.Text = "" Then
'        UserForm1.TextBox1.Text = ""
'        UserForm1.TextBox1.TabIndex = 1
'        UserForm1.TextBox2.TabIndex = 0
'    End If
'
'    If UserForm1.TextBox2.Text > "" Then
'        UserForm1.TextBox1.Text = Range("S1").Value
'        UserForm1.TextBox1.TabIndex = 0
'        UserForm1.TextBox2.TabIndex = 1
'    End If

End Sub

'TLACIDLO MODRA VYPLN V POZNAMKACH
Private Sub CommandButton10_Click()

    Call VyplnLis3Modra

End Sub

'TLACIDLO FIALOVA VYPLN V POZNAMKACH
Private Sub CommandButton11_Click()

    Call VyplnLis4Fialova

End Sub

'TLACIDLO BEZ VYPLNE V POZNAMKACH
Private Sub CommandButton12_Click()

    Call BezVyplne

End Sub

'TLACITKO CERVENE PISMO V POZNAMKACH
Private Sub CommandButton13_Click()

    Call CerveneAleboCierne

End Sub
'TLACITKO CIERNA VYPLN V POZNAMKACH
Private Sub CommandButton14_Click()

    Call VyplnCierna

End Sub
'TLACITKO ZLTOCERVENA VYPLN V POZNAMKACH
Private Sub CommandButton15_Click()

    Call VyplnZltoCervena

End Sub
'TLACITKO ZAROVNANIE TEXTU VLAVO
Private Sub CommandButton16_Click()

    Call ZarovnanieTextuVlavo

End Sub
'TLACITKO ZAROVNANIE TEXTU NA STRED
Private Sub CommandButton17_Click()

    Call ZarovnanieTextuNaStred

End Sub
'TLACITKO ZAROVNANIE TEXTU VPRAVO
Private Sub CommandButton18_Click()

    Call ZarovnanieTextuVpravo

End Sub
'TLACITKO CERVENY SYMBOL CENTROVACIEHO CAPU
Private Sub CommandButton19_Click()

    Call CervenAleboCierneCentrovanie

End Sub

'TLACITKO "VyËistiù PLAN UPINANIA"
Private Sub CommandButton2_Click()

    I = MsgBox("Naozaj chcete vymazaù ˙daje vo formul·ri?", vbYesNo + vbExclamation, "Vymazaù ˙daje")
    
    Select Case I
        Case vbNo
    '        MsgBox ("Nie")
        Case vbYes
        
            'PoËas priebehu proced˙ry "VyËistitPlanUpinania_Running" m· premenn· hodnotu "TRUE" po skonenÌ proced˙ri _
            sa zmenÌ hodnota premennÈj na "FALSE"
            VyËistitPlanUpinania_Running = True
            
            Application.ScreenUpdating = False
                        
            Call PoznamkyVStredeBezPodfarbenia

            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            
            Range("VymazUdajeBunky").ClearContents
            Range("$B$15:$AN$25,$I$14:$AN$14").ClearContents 'Vymaûe pozn·mky v strede
            
        'Hlavicky
            Worksheets("AIO_Plan").PageSetup.RightHeader = ""
            Worksheets("AIO_Plan").PageSetup.LeftHeader = ""
                        
        'CheckBoxy.value Zdvih obvodov˝ch valcov/odstavovacÌch blokov
            Range("R5").Value = False
            Range("S5").Value = False
            Range("L6").Value = "Zdvih obvodov˝ch GDF"
            
        'CheckBoxy PridrûivaËe
            Range("J13").Value = False
            Range("J12").Value = False
            Range("N13").Value = False
            Range("N12").Value = False
            Range("R13").Value = False
            Range("R12").Value = False
            
        'CheckBoxy moûnosù upnutia do lisu
            Range("Y7").Value = False
            Range("AA7").Value = False
            Range("AC7").Value = False
            Range("AE7").Value = False
            
        'CheckBoxy smer lisovania
            Range("U32").Value = False
            Range("U50").Value = False
            Range("C41").Value = False
            Range("AM41").Value = False
            
        'Vzdialenosù medzi upÌnacÌmi dr·ûkami OT
            Range("AI5").Value = ""
            
        'Vymaze koment·re Range("S9, S13")
            Range("S10").ClearComments
            Range("S9").ClearComments
            Range("S13").ClearComments
            Range("S12").ClearComments
            Range("S11").ClearComments
            
            
            Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 3", _
                    "Straight Arrow Connector 13")).Visible = msoFalse
            Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 23", _
                    "Straight Arrow Connector 24")).Visible = msoFalse
            Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 19", _
                    "Straight Arrow Connector 20")).Visible = msoFalse
            Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 22", _
                    "Straight Arrow Connector 21")).Visible = msoFalse
            
            Range("CelyStol").Select
            
            Dim Cll As Range
            Set Cll = Selection
                For Each cell In Cll
                    If cell.Value = Range("B29") Or cell.Value = Range("B30") Then
                        Range("B31").Copy
            '            Worksheets("AIO_Plan").Paste
                        Selection.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                        Application.CutCopyMode = False
                    End If
                Next cell
                        
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
            
            Call Dr·ûky
            
            Call SpocitaCerveneCentrovacieCapy
            
            Application.ScreenUpdating = True
            
            
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
            
            UserForm1.TextBox1.Text = ""
            
            VyËistitPlanUpinania_Running = False
    End Select

End Sub
'TLACITKO "Vloûiù centrovanie"
Private Sub CommandButton3_Click()

    Call CentrovaciCap

End Sub
'TLACITKO "Vloûiù tlaËn˝ Ëap"
Private Sub CommandButton4_Click()

    Call TlacnyCap

End Sub
'TLACITKO "Vloûiù voænÈ miesto"
Private Sub CommandButton5_Click()

    Call VolneMiesto

End Sub
'TLACITKO "OK"
Private Sub CommandButton6_Click()
'UPRAVI FORMAT OKNA, AKY JE PO OTVORENI FORMULARA

    Call Dr·ûky

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'    Application.ScreenUpdating = False
            
    Worksheets("AIO_Plan").Shapes.Range(Array("Group 1")).Select
        With Selection.ShapeRange
             .Left = 527
             .Top = 590 '363Hodnota,ked su pozn·mky dole
        End With
        
    Range("AN28").Select
        
'    Application.Wait Now + TimeValue("00:00:01") 'zdrûanie 1 sekundu
        
        
''Odkryje riadok vzorcov
'    Application.DisplayFormulaBar = True
'            '
''Odkryje z·hlavia
'    ActiveWindow.DisplayHeadings = True
          
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
        
    Application.WindowState = xlNormal
        Application.Left = 226
        Application.Top = 1
        Application.Width = 668 '686 '(976)
        Application.Height = 870
        
    Application.Wait Now + TimeValue("00:00:01") 'zdrûanie 1 sekundu
        
'''''    Application.WindowState = xlMinimized
'    '""""""""
'    Application.Wait Now + TimeValue("00:00:01") 'zdrûanie 1 sekundu
'    '""""""""
        
    Application.ScreenUpdating = True
    
'    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
    UserForm1.Show

End Sub
'TLACITKO "ULOZIT PLAN UPINANIA"
Private Sub CommandButton7_Click()

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
''''''''    Fcesta = "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…"
    FcestaJPG = ThisWorkbook.Path
''''''''    "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…\PU_JPG"
    FcestaPDF = ThisWorkbook.Path
''''''''    "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…\PU_PDF"
''''''''    FcestaPdfNaSiet = "T:\430_F\10_Verejne\10_planovanieLisov\Pl·ny upÌnania n·strojov"
''''''''    FcestaPdfTablet = "C:\Users\lisy\Desktop\Pl·ny upÌnania tablet"
    FCisloNastroja = Sheets("AIO_Plan").Range("S1").Text
    FOperacia = Sheets("AIO_Plan").Range("AM1").Text
    FKrok = Sheets("AIO_Plan").Range("AM3").Text
    FCisloDielu = Sheets("AIO_Plan").Range("S3").Text
    
    If Range("AM3").Text = "" Then
        NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_Pl·n upÌnania"
        Range("AJ3:AL3").NumberFormat = ";;;"
    Else
        NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_S" & FKrok & "_Pl·n upÌnania"
        Range("AJ3:AL3").NumberFormat = "@"
    End If
    
    I = MsgBox("Uloûiù ako:  " & NazovPlanuUpinania & vbCrLf & _
            "Pl·n upÌnania sa uloûÌ v PDF form·te na nasledovnÈ miesta:" & vbCrLf & _
            "" & vbCrLf & _
            "C:\Users\lisy\Desktop\Pl·ny upÌnania tablet" & vbCrLf & _
            "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…\PU_PDF" & vbCrLf & _
            "T:\430_F\10_Verejne\10_planovanieLisov\Pl·ny upÌnania n·strojov", vbYesNo + vbQuestion, "Uloûiù ako")
    
    Select Case I
        Case vbNo
        '   Unload Me
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    '       MsgBox ("Nie")
    
        Case vbYes
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'VyznaËÌ verziu pl·nu upÌnania v pl·ne upÌnania a v zosite "AIO_Data" v koment·ri bunky "D·tum poslednej aktualiz·cie"

            Dim HodnotaMID As String
            HodnotaMID = Mid(ThisWorkbook.Name, 1, 3)
            
'            MsgBox HodnotaMID
            
            If HodnotaMID = "F77" Then
            
'                MsgBox "Nov˝ pl·n upÌnania"
                
                Range("A64") = ThisWorkbook.Name
                VerziaPU = Range("A64").Value
            Else
'                MsgBox "Aktualiz·cia"
            End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        'Ak je lava hlavicka pr·zdna vloûÌ do nej datum vytvorenia, _
        upravÌ velkosù a font. V podstate doplni datum a cas vytvorenia _
        pri prvom uloûenÌ pl·nu upÌnania a potom ho uû neprepisuje. _
        Do pravej hlavicky vloûÌ vûdy pri uloûenÌ cez tlaËÌtko "ULOZIT PLAN UPINANIA" _
        aktu·lny d·tum a Ëas.
        
            
            date_test = Now()
            Teraz = Format(date_test, "dd.mm.yyyy hh:mm") 'NastavÌ form·t Ëasu
        '    MsgBox (Teraz)
        
            DatumVytvorenia = "D·tum vytvorenia: " & Teraz
            DatumPoslednejAktualizacie = "D·tum poslednej aktualiz·cie: " & Teraz
    
            Worksheets("AIO_Plan").PageSetup.RightHeader = "&""Porsche Next TT""&08" & DatumPoslednejAktualizacie

            If Worksheets("AIO_Plan").PageSetup.LeftHeader = "" Then
                Worksheets("AIO_Plan").PageSetup.LeftHeader = _
                "&""Porsche Next TT""&08" & DatumVytvorenia
'                MsgBox (DatumVytvorenia)
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
            Call Dr·ûky
        
            'Unload Me
            
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
            
        'Zabr·ni zobrezeniu systÈmov˝ch hl·öok
            Application.DisplayAlerts = False
            
''''''''''        'ULOZI AKO ".xlsm"--------------------------------------------------------
''''''''''            ActiveWorkbook.SaveAs FileName:=Fcesta & "\" & NazovPlanuUpinania & ".xlsm"
                
        'ULOZI AKO "Pdf"--------------------------------------------------------
            Worksheets("AIO_Plan").ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
                    FcestaPDF & "\" & NazovPlanuUpinania & ".pdf", _
                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
                    :=False, OpenAfterPublish:=True
                    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''        'ULOZI AKO "Pdf na sieù"--------------------------------------------------------
''''''''''            Worksheets("AIO_Plan").ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
''''''''''                    FcestaPdfNaSiet & "\" & NazovPlanuUpinania & ".pdf", _
''''''''''                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
''''''''''                    :=False, OpenAfterPublish:=True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''        'ULOZI AKO "Pdf do tabletu(na Cloud)"--------------------------------------------------------
''''''''''            Worksheets("AIO_Plan").ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
''''''''''                    FcestaPdfTablet & "\" & NazovPlanuUpinania & ".pdf", _
''''''''''                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
''''''''''                    :=False, OpenAfterPublish:=True

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
'            Call FormatNaSirku
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            
        'ULOZI AKO ".jpg"--------------------------------------------------------
            Dim cht As ChartObject
            Dim ActiveShape As Shape
            
            If ActiveWindow.DisplayGridlines = True Then
            
                ActiveWindow.DisplayGridlines = False
                'Confirm if a Cell Range is currently selected
                '  If TypeName(Selection) <> "Range" Then
                '    MsgBox "You do not have a single shape selected!"
                '    Exit Sub
                '  End If
                
                'Copy/Paste Cell Range as a Picture
                  Range("A1:AO63").Copy
                  Worksheets("AIO_Plan").Pictures.Paste(link:=False).Select
                  Selection.ShapeRange.ScaleWidth 1.7298371577, msoFalse, msoScaleFromTopLeft 'TU SA NASTAVUJE VELKOST SIRKY SKOPIROVANEJ OBLASTI
                  Selection.ShapeRange.ScaleHeight 1.7298371969, msoFalse, msoScaleFromTopLeft 'TU SA NASTAVUJE VELKOST VYSKY SKOPIROVANEJ OBLASTI
                  Set ActiveShape = Worksheets("AIO_Plan").Shapes(ActiveWindow.Selection.Name)
                  
                'Create a temporary chart object (same size as shape)
                  Set cht = Worksheets("AIO_Plan").ChartObjects.Add( _
                    Left:=1, _
                    Width:=966, _
                    Top:=1, _
                    Height:=1470) 'TU SA NASTAVUJE SIRKA A VYSKA GRAFU
                    'PovodneW=735,H=483
                
''''                'Copy/Paste Cell Range as a Picture
''''                  Range("A1:BO50").Copy
''''                  Worksheets("AIO_Plan").Pictures.Paste(link:=False).Select
''''                  Set ActiveShape = Worksheets("AIO_Plan").Shapes(ActiveWindow.Selection.Name)
''''
''''                'Create a temporary chart object (same size as shape)
''''                  Set cht = Worksheets("AIO_Plan").ChartObjects.Add( _
''''                    Left:=1, _
''''                    Width:=735, _
''''                    Top:=1, _
''''                    Height:=483)
                
                'Format temporary chart to have a transparent background
                  cht.ShapeRange.Fill.Visible = msoFalse
                  cht.ShapeRange.Line.Visible = msoFalse
                    
                'Copy/Paste Shape inside temporary chart
                  ActiveShape.Copy
                  cht.Activate
                  ActiveChart.Paste
                  
                'Save chart to User's Desktop as PNG File
                  cht.Chart.Export FileName:=FcestaJPG & "\" & NazovPlanuUpinania & ".jpg"
'                  cht.Chart.Export Filename:=FcestaPdfTablet & "\PU_JPG\" & NazovPlanuUpinania & ".jpg" 'UloûÌ do prieËinku "Pl·ny upÌnania tablet"
                'Delete temporary Chart
                  cht.Delete
                  ActiveShape.Delete
                  
                ActiveWindow.DisplayGridlines = True
            
            End If
            
            If ActiveWindow.DisplayGridlines = False Then
            
                'Confirm if a Cell Range is currently selected
                '  If TypeName(Selection) <> "Range" Then
                '    MsgBox "You do not have a single shape selected!"
                '    Exit Sub
                '  End If
                
                'Copy/Paste Cell Range as a Picture
                  Range("A1:AO63").Copy
                  Worksheets("AIO_Plan").Pictures.Paste(link:=False).Select
                  Selection.ShapeRange.ScaleWidth 1.7298371577, msoFalse, msoScaleFromTopLeft 'TU SA NASTAVUJE VELKOST SIRKY SKOPIROVANEJ OBLASTI
                  Selection.ShapeRange.ScaleHeight 1.7298371969, msoFalse, msoScaleFromTopLeft 'TU SA NASTAVUJE VELKOST VYSKY SKOPIROVANEJ OBLASTI
                  Set ActiveShape = Worksheets("AIO_Plan").Shapes(ActiveWindow.Selection.Name)
                  
                'Create a temporary chart object (same size as shape)
                  Set cht = Worksheets("AIO_Plan").ChartObjects.Add( _
                    Left:=1, _
                    Width:=966, _
                    Top:=1, _
                    Height:=1470) 'TU SA NASTAVUJE SIRKA A VYSKA GRAFU
                    'PovodneW=735,H=483
                
''''                'Copy/Paste Cell Range as a Picture
''''                  Range("A1:BO50").Copy
''''                  Worksheets("AIO_Plan").Pictures.Paste(link:=False).Select
''''                  Set ActiveShape = Worksheets("AIO_Plan").Shapes(ActiveWindow.Selection.Name)
''''
''''                'Create a temporary chart object (same size as shape)
''''                  Set cht = Worksheets("AIO_Plan").ChartObjects.Add( _
''''                    Left:=1, _
''''                    Width:=735, _
''''                    Top:=1, _
''''                    Height:=483)

                
                'Format temporary chart to have a transparent background
                  cht.ShapeRange.Fill.Visible = msoFalse
                  cht.ShapeRange.Line.Visible = msoFalse
                    
                'Copy/Paste Shape inside temporary chart
                  ActiveShape.Copy
                  cht.Activate
                  ActiveChart.Paste
                  
                'Save chart to User's Desktop as PNG File
                  cht.Chart.Export FileName:=FcestaJPG & "\" & NazovPlanuUpinania & ".jpg"
'                  cht.Chart.Export Filename:=FcestaPdfTablet & "\PU_JPG\" & NazovPlanuUpinania & ".jpg" 'UloûÌ do prieËinku "Pl·ny upÌnania tablet"
                'Delete temporary Chart
                  cht.Delete
                  ActiveShape.Delete
              
            End If
              
            Call SpocitaCerveneCentrovacieCapy
              
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
            
'            Call FormatNaVysku
            
        'OTVORI JPG SUBOR--------------------------------------------------------
            VBA.Shell "Explorer.exe " & FcestaJPG & "\" & NazovPlanuUpinania & ".jpg"
            
        'V SUBORE "AIO_Data" VYZNACI MOZNOST UPNUTIA NASTROJA DO LISU + ZE BOL VYTVORENY PLAN UPINANIA V NOVOM FORMULARI
        '---------------------------------------------------------------------------------------
            If Range("S10") = "" Then
                Call OtvorNajdiVyznacAll
'                Call DoplniOstatneUdajeDoAIO_Data
                Call OtvoriParamNastrAVytvoriHyperlinkyExcelPdf
            Else: 'MsgBox ("Spustam makro OtvoriParamNastrADoplniOstatneUdajeDoAIO_Data")
                Call OtvoriParamNastrADoplniOstatneUdajeDoAIO_Data
                
'                Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Activate
'
'                Application.WindowState = xlNormal
'                    Application.Left = 1 '226
'                    Application.Top = 1
'                    Application.Width = 601 '686 '(976)
'                    Application.Height = 870
                 
                Call OtvoriParamNastrAVytvoriHyperlinkyExcelPdf
            End If
            
    End Select
    
    Application.DisplayAlerts = True 'PovolÌ zobrezeniu systÈmov˝ch hl·öok
        
End Sub

'TLACIDLO ZLTA VYPLN V POZNAMKACH
Private Sub CommandButton8_Click()

    Call VyplnLis1Zlta

End Sub

'TLACIDLO CERVeNA VYPLN V POZNAMKACH
Private Sub CommandButton9_Click()

    Call VyplnLis2Cervena

End Sub

'SPUSTI SA PO ZMENE UDAJOV BUNKY...
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
'---------------------------------------------------------------------------------------
'POZNAMKY
'po zmene ˙dajov v "pozn·mkach k n·stroju" a zmena zobrazÌ aj v "Pozn·mkach k n·stroju" pre JPG aj v skryt˝ch
    Dim Pozn As Range
    Set Pozn = Range("Poznamky")
    
    If Not Application.Intersect(Pozn, Range(Target.Address)) Is Nothing Then
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        
    '1.riadok
        Range("AQ5").FormulaR1C1 = Range("I51").FormulaR1C1
        Range("I14").FormulaR1C1 = Range("AQ5").FormulaR1C1
        
    '2.riadok
        Range("AQ6").FormulaR1C1 = Range("B52").FormulaR1C1
        Range("B15").FormulaR1C1 = Range("AQ6").FormulaR1C1
        
    '3.riadok
        Range("AQ8").FormulaR1C1 = Range("B53").FormulaR1C1
        Range("B16").FormulaR1C1 = Range("AQ8").FormulaR1C1
    
    '4.riadok
        Range("AQ10").FormulaR1C1 = Range("B54").FormulaR1C1
        Range("B17").FormulaR1C1 = Range("AQ10").FormulaR1C1
    
    ' 5.riadok
        Range("AQ12").FormulaR1C1 = Range("B55").FormulaR1C1
        Range("B18").FormulaR1C1 = Range("AQ12").FormulaR1C1
    
    ' 6.riadok
        Range("AQ27").FormulaR1C1 = Range("B56").FormulaR1C1
        Range("B19").FormulaR1C1 = Range("AQ27").FormulaR1C1
    
    ' 7.riadok
        Range("AQ30").FormulaR1C1 = Range("B57").FormulaR1C1
        Range("B20").FormulaR1C1 = Range("AQ30").FormulaR1C1
    
    ' 8.riadok
        Range("AQ33").FormulaR1C1 = Range("B58").FormulaR1C1
        Range("B21").FormulaR1C1 = Range("AQ33").FormulaR1C1
    
    ' 9.riadok
        Range("AQ36").FormulaR1C1 = Range("B59").FormulaR1C1
        Range("B22").FormulaR1C1 = Range("AQ36").FormulaR1C1
        
    ' 10.riadok
        Range("AQ39").FormulaR1C1 = Range("B60").FormulaR1C1
        Range("B23").FormulaR1C1 = Range("AQ39").FormulaR1C1
        
    ' 11.riadok
        Range("AQ42").FormulaR1C1 = Range("B61").FormulaR1C1
        Range("B24").FormulaR1C1 = Range("AQ42").FormulaR1C1
        
    ' 12.riadok
        Range("AQ45").FormulaR1C1 = Range("B62").FormulaR1C1
        Range("B25").FormulaR1C1 = Range("AQ45").FormulaR1C1
        
        
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If
    
'---------------------------------------------------------------------------------------
'SUCET VAH
'Po zmene hodnoty bunky "A1" alebo "A2" spustÌ makro pre suËet s podmienkou
    Dim KlucoveBunky As Range
    Set KlucoveBunky = Range("G5:G6")
    
    If Not Application.Intersect(KlucoveBunky, Range(Target.Address)) Is Nothing Then
    
        If Range("G5") > 0 And Range("G6") > 0 Then
            Range("G7") = Application.WorksheetFunction.Sum(Range("G5:G6"))
        Else: Range("G7") = ""
        End If
    
    End If
    
'---------------------------------------------------------------------------------------
'UPINACIA VYSKA
'Po zmene hodnoty bunky "AK5" alebo "W6" spustÌ makro pre suËet s podmienkou
    Dim ZdvihGDF As Range
    Set ZdvihGDF = Range("AK5,W6")
    
    If Not Application.Intersect(ZdvihGDF, Range(Target.Address)) Is Nothing Then
    
        If Range("W6") > 0 Then
            Range("AK6") = Application.WorksheetFunction.Sum(Range("AK5,W6"))
        Else: Range("AK6") = ""
        End If
    
    End If
    
'---------------------------------------------------------------------------------------
'SUCET TLACNYCH CAPOV
'Po zmene v oblasti CelyStol vloûÌ do bunky AN28 funkciu "COUNTIF",
'ktor· spoËÌta bunky v danom rozsahu, ktorÈ splnaju urËit˙ podmienku
    Dim CelyStol As Range
    Set CelyStol = Range("F34:AJ48")
    
    If Not Application.Intersect(CelyStol, Range(Target.Address)) Is Nothing Then
        'MsgBox ("Doölo k zmene")
        Range("AN28").FormulaLocal = "=COUNTIF(F34:AJ48;B30)"
    '
    End If
    
'---------------------------------------------------------------------------------------
'CENTROVANIE SURADNICE
'Po zmene hodnoty bunky v oblasi "CentrSurad" zmeni hodnotu podla podmienok nizsie
    Dim Cntr As Range
    Set Cntr = Range("CentrSurad")
    
    If Not Application.Intersect(Cntr, Range(Target.Address)) Is Nothing Then
        For Each cell In Cntr
    '        If Cell.Value = 0 Then
    '        Cell.Value = 1
    '        End If
                If cell.Value = 150 Then
                cell.Value = 2
                End If
                    If cell.Value = 300 Then
                    cell.Value = 3
                    End If
                        If cell.Value = 450 Then
                        cell.Value = 4
                        End If
                            If cell.Value = 600 Then
                               cell.Value = 5
                               End If
            If cell.Value = 750 Then
            cell.Value = 6
            End If
                If cell.Value = 900 Then
                cell.Value = 7
                End If
                    If cell.Value = 1050 Then
                    cell.Value = 8
                    End If
                        If cell.Value = 1200 Then
                        cell.Value = 9
                        End If
                            If cell.Value = 1350 Then
                               cell.Value = 10
                               End If
            If cell.Value = 1500 Then
            cell.Value = 11
            End If
                If cell.Value = 1650 Then
                cell.Value = 12
                End If
                    If cell.Value = 1800 Then
                    cell.Value = 13
                    End If
                        If cell.Value = 1950 Then
                        cell.Value = 14
                        End If
                            If cell.Value = 2100 Then
                               cell.Value = 15
                               End If
                                   If cell.Value = 2250 Then
                                   cell.Value = 16
                                   End If
                                       If cell.Value = 2400 Then
                                       cell.Value = 17
                                       End If
        Next cell
                           
    End If
    On Error GoTo 0
'---------------------------------------------------------------------------------------
'SURADNICE CENTROVANIA LH
    Dim LHR As Range
    Set LHR = Range("T28")
    
    On Error Resume Next
    
    If Not Application.Intersect(LHR, Range(Target.Address)) Is Nothing Then
       Range("S29").Select
    End If
    
    Dim LHS As Range
    Set LHS = Range("S29")
    
    If Not Application.Intersect(LHS, Range(Target.Address)) Is Nothing Then
       Call SuradniceCentrovaniaLH
       Range("V28").Select
    End If
    'FUNGUJE Po zmene buniek "t29" Vyberie Bunku "s29", po zmene bunky "s29" _
    vyberie bunku "V28" a spusti makro "SuradniceCentrovaniaLH"
    
'---------------------------------------------------------------------------------------
'SURADNICE CENTROVANIA PH
    Dim PHR As Range
    Set PHR = Range("V28")
    
    If Not Application.Intersect(PHR, Range(Target.Address)) Is Nothing Then
       Range("W29").Select
    End If
    
    Dim PHS As Range
    Set PHS = Range("W29")
    
    If Not Application.Intersect(PHS, Range(Target.Address)) Is Nothing Then
       Call SuradniceCentrovaniaPH
       Range("T31").Select
    End If
    'FUNGUJE Po zmene buniek "V28" Vyberie Bunku "W29", po zmene bunky "W29" _
    vyberie bunku "T31" a spusti makro "SuradniceCentrovaniaPH"
    
'---------------------------------------------------------------------------------------
'SURADNICE CENTROVANIA LD
    Dim LDR As Range
    Set LDR = Range("T31")
    
    
    If Not Application.Intersect(LDR, Range(Target.Address)) Is Nothing Then
       Range("S30").Select
    End If
    
    Dim LDS As Range
    Set LDS = Range("S30")
    
    If Not Application.Intersect(LDS, Range(Target.Address)) Is Nothing Then
       Call SuradniceCentrovaniaLD
       Range("V31").Select
    End If
    'FUNGUJE Po zmene buniek "T31" Vyberie Bunku "S30", po zmene bunky "S30" _
    vyberie bunku "V31" a spusti makro "SuradniceCentrovaniaLD"
    
'---------------------------------------------------------------------------------------
'SURADNICE CENTROVANIA PD
    Dim PDR As Range
    Set PDR = Range("V31")
    
    If Not Application.Intersect(PDR, Range(Target.Address)) Is Nothing Then
       Range("W30").Select
    End If
    
    Dim PDS As Range
    Set PDS = Range("W30")
    
    If Not Application.Intersect(PDS, Range(Target.Address)) Is Nothing Then
       Call SuradniceCentrovaniaPD
       Range("T28").Select
    End If
    
    On Error GoTo 0
    'FUNGUJE Po zmene buniek "V31" Vyberie Bunku "W30", po zmene bunky "W30" _
    vyberie bunku "T28" a spusti makro "SuradniceCentrovaniaPD"
     
'---------------------------------------------------------------------------------------
''''''''''
'FUNGUJE
'Po vloûenÌ ˙daja do bunky S9-S13 sa ˙daj uloûÌ do koment·ra danej bunky, _
 po dalöej zmene tejto bunky sa do koment·ra doplnÌ nov˝ ˙daj, priËom predch·dzajuci tam zostane tieû _
 (Koment·r je tak v podstate histÛria parametrov)

'BARAN
    If Target.Address = Range("S10").Address Then
        On Error Resume Next
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            If Target.Comment Is Nothing Then
                Target.AddComment Text:=Now() & " / " & Target.Text & "kN" & vbLf
                Target.Comment.Shape.TextFrame.AutoSize = True
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                Call ZmeniTlakBarana
            Else
                Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "kN" & vbLf
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                Call ZmeniTlakBarana
                
            End If
        On Error GoTo 0
'        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If
        
'BRZDA
    If Target.Address = Range("S9").Address Then
        On Error Resume Next
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            If Target.Comment Is Nothing Then
                Target.AddComment Text:=Now() & " / " & Target.Text & "mm" & vbLf
                Target.Comment.Shape.TextFrame.AutoSize = True
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                Call ZmeniBrzda
            Else
                Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "mm" & vbLf
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                Call ZmeniBrzda
            End If
        On Error GoTo 0
'        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If
    
'PRIDRZIAVAC BARAN
    If Target.Address = Range("S13").Address Then
        On Error Resume Next
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            If Target.Comment Is Nothing Then
                Target.AddComment Text:=Now() & " / " & Target.Text & "kN" & vbLf
                Target.Comment.Shape.TextFrame.AutoSize = True
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                Call ZmeniTlakHP
            Else
                Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "kN" & vbLf
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                Call ZmeniTlakHP
            End If
        On Error GoTo 0
'        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If
        
'PRIDRZIAVAC STOL
    If Target.Address = Range("S12").Address Then
        On Error Resume Next
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            If Target.Comment Is Nothing Then
                Target.AddComment Text:=Now() & " / " & Target.Text & "kN" & vbLf
                Target.Comment.Shape.TextFrame.AutoSize = True
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                Call ZmeniTlakDP
            Else
                Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "kN" & vbLf
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                Call ZmeniTlakDP
            End If
        On Error GoTo 0
'        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If
    
'CAPY NAD STOL
    If Target.Address = Range("S11").Address Then
        On Error Resume Next
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            If Target.Comment Is Nothing Then
                Target.AddComment Text:=Now() & " / " & Target.Text & "mm" & vbLf
                Target.Comment.Shape.TextFrame.AutoSize = True
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                Call Zmeni»apyNadStÙl
            Else
                Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "mm" & vbLf
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                Call Zmeni»apyNadStÙl
            End If
        On Error GoTo 0
'        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If
    
    'VAHA GES
    If Target.Address = Range("G7").Address Then
        On Error Resume Next
'        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
'            If Target.Comment Is Nothing Then
'                Target.AddComment Text:=Now() & " / " & Target.Text & "mm" & vbLf
'                Target.Comment.Shape.TextFrame.AutoSize = True
'                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'                Call Zmeni»apyNadStÙl
'            Else
'                Target.Comment.Text Text:=Target.Comment.Text & Now() & " / " & Target.Text & "mm" & vbLf
'                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'                Call Zmeni»apyNadStÙl
'            End If
            If Range("G7").Value > 0 Then
'                MsgBox ("ZApisujem vahu do paremetre nastrojov")
                Call ZmeniVahuVAIO_Data
            Else
'                MsgBox ("NEZApisujem vahu do paremetre nastrojov")
            End If
        On Error GoTo 0
'        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If
    '''''''''
'---------------------------------------------------------------------------------------
'CISLO NASTROJA
'Po zmene hodnoty bunky "S1" spustÌ makro "DoplniUdaje
    Dim CisloNastroja As Range
    Set CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1")
    
    If Not Application.Intersect(CisloNastroja, Range(Target.Address)) Is Nothing Then
'        MsgBox ("spustam /harok1 /cislo nastroja change")
       
        If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Text > "" And _
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S3").Text = "" Then
'            MsgBox ("Udaje nie su doplnenÈ")
            Application.Wait Now + TimeValue("00:00:02")
            
            Call DoplniUdaje
            
        Else
        
            If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Text > "" And _
               Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S3").Text > "" And _
               Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("W5").Value = "" Then
'               MsgBox (" Otv·ram formul·r")
               Call ZavolaOtvoriùFormular
               
            Else
            
                If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Text > "" And _
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S3").Text > "" And _
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("W5").Value > 0 Then
                    MsgBox ("Udaje s˙ uû vyplnenÈ. Pre automatickÈ doplnenie ˙dajov vyËistite najskÙr pl·n upÌnania!")
                
                End If
                
            End If
            
        End If
         
    End If
 
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

'_________________________________________________________
'Sleduje oblasù "POZNAMKY V STREDE", ak sa aktÌvna bunka nach·dza v tejto oblasti, _
TlaËidl· pre zmenu v˝plne s˙ viditelnÈ, v opaËnom prÌpade s˙ neviditeænÈ

    Const WS_RANGE As String = "Pozn·mkyVStrede"     '<== change to suit

    On Error GoTo ws_exit
    Application.EnableEvents = False

    If Not Intersect(Target, Me.Range(WS_RANGE)) Is Nothing Then
        With Target
        
'            MsgBox ("AktÌvna bunka v pozn·mkach, tlacidl· viditelnÈ")
            Call ViditelnosùTlacidielVyplneVPoznamkach
        End With
    Else
'        MsgBox ("Nie je aktÌvna bunka v pozn·mkach, tlacidl· neviditelnÈ")
        Call NeviditelnostTlacidielVyplneVPozn·mkach
    End If

ws_exit:
    
    Application.EnableEvents = True
'_________________________________________________________

'Sleduje oblasù "CELY STOL", ak sa aktÌvna bunka nach·dza v tejto oblasti, _
TlaËidl· pre zmenu rastrastola s˙ viditelnÈ, v opaËnom prÌpade s˙ neviditeænÈ

    Const RS_RANGE As String = "CelyStol"     '<== change to suit

    On Error GoTo rs_exit
    Application.EnableEvents = False

    If Not Intersect(Target, Me.Range(RS_RANGE)) Is Nothing Then
        With Target

'            MsgBox ("AktÌvna bunka vrastri stola, tlacidl· viditelnÈ")
            Call ViditelnosùTlacidielPreRasterStola
            If ActiveCell.Value = Range("B29").Value Then
'               MsgBox ("Vidim tlacitko cerveny CentrovacÌ Ëap")
               Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").CommandButton19.Visible = True

            Else: 'MsgBox ("Nevidim tlacitko cerveny CentrovacÌ Ëap")
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").CommandButton19.Visible = False

            End If
        End With
    Else
'        MsgBox ("Nie je aktÌvna bunka v rastri stola, tlacidl· neviditelnÈ")
        Call NeviditelnostTlacidielPreRasterStola
'        MsgBox ("Nevidim tlacitko cerveny CentrovacÌ Ëap")
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").CommandButton19.Visible = False

    End If

rs_exit:

    Application.EnableEvents = True
'_________________________________________________________

End Sub


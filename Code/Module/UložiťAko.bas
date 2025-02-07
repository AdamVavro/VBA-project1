Attribute VB_Name = "UloûiùAko"
Sub UlozitAkoJpgNaVyskuPoznamkyVStrede()

''TLACITKO "ULOZIT PLAN UPINANIA"
'Private Sub CommandButton7_Click()

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
'''''    Fcesta = "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…"
    FcestaJPG =
'''''    "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…\PU_JPG"
    FcestaPDF =
''''''    "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…\PU_PDF"
''''''    FcestaPdfNaSiet = "T:\430_F\10_Verejne\10_planovanieLisov\Pl·ny upÌnania n·strojov"
''''''    FcestaPdfTablet = "C:\Users\lisy\Desktop\Pl·ny upÌnania tablet"
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
    
'    i = MsgBox("Uloûiù ako:  " & NazovPlanuUpinania & vbCrLf & _
'            "Pl·n upÌnania sa uloûÌ v PDF form·te na nasledovnÈ miesta:" & vbCrLf & _
'            "" & vbCrLf & _
'            "C:\Users\lisy\Desktop\Pl·ny upÌnania tablet" & vbCrLf & _
'            "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…\PU_PDF" & vbCrLf & _
'            "T:\430_F\10_Verejne\10_planovanieLisov\Pl·ny upÌnania n·strojov", vbYesNo + vbQuestion, "Uloûiù ako")
'
'    Select Case i
'        Case vbNo
'        '   Unload Me
'            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'    '       MsgBox ("Nie")
'
'        Case vbYes
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        'Ak je lava hlavicka pr·zdna vloûÌ do nej datum vytvorenia, _
'        upravÌ velkosù a font. V podstate doplni datum a cas vytvorenia _
'        pri prvom uloûenÌ pl·nu upÌnania a potom ho uû neprepisuje. _
'        Do pravej hlavicky vloûÌ vûdy pri uloûenÌ cez tlaËÌtko "ULOZIT PLAN UPINANIA" _
'        aktu·lny d·tum a Ëas.
'
'            date_test = Now()
'            Teraz = Format(date_test, "d.m.yyyy hh:mm") 'NastavÌ form·t Ëasu
'        '    MsgBox (Teraz)
'
'            DatumVytvorenia = "D·tum vytvorenia: " & Teraz
'            DatumPoslednejAktualizacie = "D·tum poslednej aktualiz·cie: " & Teraz
'
'            Worksheets("AIO_Plan").PageSetup.RightHeader = "&""Porsche Next TT""&08" & DatumPoslednejAktualizacie
'
'            If Worksheets("AIO_Plan").PageSetup.LeftHeader = "" Then
'                Worksheets("AIO_Plan").PageSetup.LeftHeader = _
'                "&""Porsche Next TT""&08" & DatumVytvorenia
''                MsgBox (DatumVytvorenia)
'            End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'            'Unload Me
'
'            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'
'        'Zabr·ni zobrezeniu systÈmov˝ch hl·öok
'            Application.DisplayAlerts = False
'
'        'ULOZI AKO ".xlsm"--------------------------------------------------------
'            ActiveWorkbook.SaveAs Filename:=Fcesta & "\" & NazovPlanuUpinania & ".xlsm"
'
'        'ULOZI AKO "Pdf"--------------------------------------------------------
'            Worksheets("AIO_Plan").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'                    FcestaPDF & "\" & NazovPlanuUpinania & ".pdf", _
'                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'                    :=False, OpenAfterPublish:=True
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        'ULOZI AKO "Pdf na sieù"--------------------------------------------------------
'            Worksheets("AIO_Plan").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'                    FcestaPdfNaSiet & "\" & NazovPlanuUpinania & ".pdf", _
'                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'                    :=False, OpenAfterPublish:=True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        'ULOZI AKO "Pdf do tabletu(na Cloud)"--------------------------------------------------------
'            Worksheets("AIO_Plan").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'                    FcestaPdfTablet & "\" & NazovPlanuUpinania & ".pdf", _
'                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'                    :=False, OpenAfterPublish:=True
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''            Call FormatNaSirku
'
'            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
'
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
                  Range("A1:AO50").Copy
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
                  Range("A1:AO50").Copy
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

            Worksheets("AIO_Plan").Protect Password:="Lis.0123"

'            Call FormatNaVysku

        'OTVORI JPG SUBOR--------------------------------------------------------
            VBA.Shell "Explorer.exe " & FcestaJPG & "\" & NazovPlanuUpinania & ".jpg"

'        'V SUBORE "AIO_Data" VYZNACI MOZNOST UPNUTIA NASTROJA DO LISU + ZE BOL VYTVORENY PLAN UPINANIA V NOVOM FORMULARI
'        '---------------------------------------------------------------------------------------
'            If Range("S10") = "" Then
'                Call OtvorNajdiVyznacAll
'
'                Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Activate
'
'                Application.WindowState = xlNormal
'                    Application.Left = 226
'                    Application.Top = 1
'                    Application.Width = 686 '(976)
'                    Application.Height = 870
'            End If
'
'    End Select
'
'    Application.DisplayAlerts = True 'PovolÌ zobrezeniu systÈmov˝ch hl·öok
    


End Sub


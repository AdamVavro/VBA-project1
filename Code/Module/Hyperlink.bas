Attribute VB_Name = "Hyperlink"
'Private Sub hyper()
'
'    FcestaPdfNaSiet = "T:\430_F\10_Verejne\10_planovanieLisov\Pl�ny up�nania n�strojov"
'    FCisloNastroja = Sheets("AIO_Plan").Range("S1").Text
'    FOperacia = Sheets("AIO_Plan").Range("AM1").Text
'    FKrok = Sheets("AIO_Plan").Range("AM3").Text
'    FCisloDielu = Sheets("AIO_Plan").Range("S3").Text
'
'    If Range("AM3").Text = "" Then
'        NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_Pl�n up�nania"
'        Range("AJ3:AL3").NumberFormat = ";;;"
'    Else
'        NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_S" & FKrok & "_Pl�n up�nania"
'        Range("AJ3:AL3").NumberFormat = "@"
'    End If
'
''    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''        'ULOZI AKO "Pdf na sie�"--------------------------------------------------------
''            Worksheets("AIO_Plan").ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
''                    FcestaPdfNaSiet & "\" & NazovPlanuUpinania & ".pdf", _
''                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
''                    :=False, OpenAfterPublish:=True
''    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'ActiveCell.Hyperlinks.Add Anchor:=Worksheets("AIO_Plan") _
'    .Range("I14"), Address:=ThisWorkbook.Name, _
'    SubAddress:="", ScreenTip:="EXCEL" & vbCrLf & _
'    ThisWorkbook.FullName _
'    , TextToDisplay:=""
'
'End Sub
'Sub HypertextovePrepojenieNaPlanUpinania()
'
''    Dim sXLFile As String
''
''    sXLFile = "C:\Users\lisy\Desktop\Pl�ny up�nania\PU_NOV�\ZMAZATTIEZ_OP60_TestUlozeniaDoJpg_S3_Pl�n up�nania.xlsm"
''
''    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").FollowHyperlink Address:=sFolder, NewWindow:=True 'Open Folder
'
''    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Hyperlinks.Add Anchor:=Selection, Address:= _
''    "C:\Users\lisy\Desktop\Pl�ny up�nania\PU_NOV�\ZMAZATTIEZ_OP60_TestUlozeniaDoJpg_S3_Pl�n up�nania.xlsm", TextToDisplay:="ZMAZATTIEZ"
'   '--------------------------------------------
'    JmenoSouboru = ThisWorkbook.Name 'ok funguje
'    MsgBox (JmenoSouboru) 'ok funguje
'    CestaVcetneNazvuSouboru = ThisWorkbook.FullName 'ok funguje
'    MsgBox (CestaVcetneNazvuSouboru) 'ok funguje
''    Worksheets("AIO_Plan").PageSetup.LeftFooter = _
'                "&""Porsche Next TT""&08&Z&F" 'OK funguje
'
''    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Hyperlinks.Add Range("I14"), "CestaVcetneNazvuSouboru"
'
'End Sub

Sub OtvoriParamNastrAVytvoriHyperlinkyExcelPdf()
        'V SUBORE "AIO_Data" Hypertextove prepojenie pre pl�n up�nania v exceli a v prf na stieti
        '---------------------------------------------------------------------------------------
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'            If Worksheets("AIO_Plan").PageSetup.RightHeader = "" Then
''                MsgBox ("prava hlavi�ka pr�zdna nerob�m ni�")
'            Else
''                MsgBox ("otvaram AIO_Data a spustam makro OtvoriParamNastrAVytvoriHyperlinkyExcelPdf")
                I = MsgBox("Prajete si doplni� hyperlink pre otvorenie pl�nu up�nania v exceli a v pdf v s�bore 'AIO_Data'?  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "Vyzna�i� mo�nos� upnutia?")
                    
                    Select Case I
                        Case vbNo
                            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                '           MsgBox ("Nie")
                        Case vbYes

                            'Fcesta = "C:\Users\lisy\Desktop\Pl�ny up�nania\PU_NOV�"
                            'FcestaJPG = "C:\Users\lisy\Desktop\Pl�ny up�nania\PU_NOV�\PU_JPG"
                            'FcestaPDF = "C:\Users\lisy\Desktop\Pl�ny up�nania\PU_NOV�\PU_PDF"
                            FCisloNastroja = Sheets("AIO_Plan").Range("S1").Text
                            FOperacia = Sheets("AIO_Plan").Range("AM1").Text
                            FKrok = Sheets("AIO_Plan").Range("AM3").Text
                            FCisloDielu = Sheets("AIO_Plan").Range("S3").Text

                            If Range("AM3").Text = "" Then
                                NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_Pl�n up�nania"
                            Else
                                NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_S" & FKrok & "_Pl�n up�nania"
                            End If
                            '    MsgBox (NazovPlanuUpinania)
                            'OK--------------------------------

                        'OTVORI SUBOR "AIO_Data"
                            On Error Resume Next
                            Set Zosit = Workbooks("AIO_Data")
                            ZositOtvoreny = Not Zosit Is Nothing
                            If ZositOtvoreny = True Then

                                'UserForm1.Unload
                                'Application.WindowState = xlMaximized

'                                MsgBox "S�bor AIO_Data je u� otvoren�"

                        'Zmens� okno excelu na lav� polovicu obrazovky
                                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                                Application.WindowState = xlNormal
'''''                                Application.Left = 1 '226
'''''                                Application.Top = 1
'''''                                Application.Width = 601 '668 '686 '(976)
'''''                                Application.Height = 870

                                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate

                        'Skryje riadok vzorcov
                '                Application.DisplayFormulaBar = False
                                '
                        'Skryje z�hlavia
                '                ActiveWindow.DisplayHeadings = False

                            'Zmens� okno excelu na prav� polovicu obrazovky
'''''                                Application.WindowState = xlNormal
'''''                                Application.Left = 602 '226
'''''                                Application.Top = 1
'''''                                Application.Width = 601 '754
'''''                                Application.Height = 870
                            Else

                                'UserForm1.Unload
                                'Application.WindowState = xlMaximized

'                                MsgBox "Otv�ram s�bor AIO_Data"

                            'Zmens� okno excelu na lav� polovicu obrazovky
                                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                                Application.WindowState = xlNormal
'''''                                Application.Left = 1 '226
'''''                                Application.Top = 1
'''''                                Application.Width = 601 '668 '686 '(976)
'''''                                Application.Height = 870


'                                Workbooks.Open FileName:="C:\Users\lisy\Desktop\Pl�ny up�nania\Parametre n�strojov\Parametre n�strojov.xlsm"
                                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate

                        'Skryje riadok vzorcov
                '                Application.DisplayFormulaBar = False
                                '
                        'Skryje z�hlavia
                '                ActiveWindow.DisplayHeadings = False

                            'Zmens� okno excelu na prav� polovicu obrazovky
'''''                                Application.WindowState = xlNormal
'''''                                Application.Left = 602 '226
'''''                                Application.Top = 1
'''''                                Application.Width = 601 '754
'''''                                Application.Height = 870
                            End If
                            'OK-------------------------------
'
'''                                Call DoplniOstatneUdajeDoAIO_Data
'''
'''                                If Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("AN28").Value > 0 Then
'''                    '                MsgBox ("Sp��tam DoplniRasterStolaDoAIO_Data")
'''                                    Call DoplniRasterStolaDoAIO_Data
'''                    '            Else
'''                    '                MsgBox ("Nep��tam DoplniRasterStolaDoAIO_Data")
'''                                End If
'--------------------------------------------------------------------------
'DO SUBORU "Paremetre n�strojov" VLOZI HYPERTEXTOVE PREPOJENIE NA PLAN UPINANIA V EXCELI A V PDF

                            Fcesta = "C:\Users\lisy\Desktop\Pl�ny up�nania\PU_NOV�"
                            FcestaPdfNaSiet = "T:\430_F\10_Verejne\10_planovanieLisov\Pl�ny up�nania n�strojov"
                            HyperlinkAdresaPdf = FcestaPdfNaSiet & "\" & NazovPlanuUpinania & ".pdf"
                            HyperlinkAdresaExcel = Fcesta & "\" & NazovPlanuUpinania & ".xlsm" 'ThisWorkbook.Name
                            
                            '----------
                            date_test = Now()
                            Teraz = Format(date_test, "d.m.yyyy hh:mm") 'Nastav� form�t �asu
                        '    MsgBox (Teraz)
                        
                            DatumUpravy = "D�tum �pravy: " & Teraz
'                            DatumPoslednejAktualizacie = "D�tum poslednej aktualiz�cie: " & Teraz
                            '-----------
                            
'                            MsgBox ("Excel Adresa: " & HyperlinkAdresaExcel)
'                            MsgBox ("PDF Adresa: " & HyperlinkAdresaPdf)
                            
                            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate

                            CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
                            Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
                            If OblastNajdi Is Nothing Then
                                I = MsgBox("��slo n�stroja sa nena�lo!", vbOKOnly + vbExclamation, "��slo n�stroja")
                            Else

'                               ExcelLink ok FUNGUJE
'                                MsgBox ("Spustam Excellink")
                                Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column).Select
                                Selection.Font.Underline = xlUnderlineStyleSingle 'pod�iarknutie textu v bunke
                                ActiveCell.Hyperlinks.Add Anchor:=Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column) _
                                , Address:=Fcesta & "\" & NazovPlanuUpinania & ".xlsm", _
                                SubAddress:="", ScreenTip:="Otvori� pl�n up�nania v EXCELI" & vbCrLf & _
                                Fcesta & "\" & NazovPlanuUpinania & ".xlsm" & vbCrLf & _
                                DatumUpravy ' _
                                , TextToDisplay:=CisloNastroja
                                
'                               PDFLink
'                                MsgBox ("Spustam PDFLink")
                                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                                Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column - 1).Select
                                Selection.Font.Underline = xlUnderlineStyleSingle 'pod�iarknutie textu v bunke
                                ActiveCell.Hyperlinks.Add Anchor:=Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column - 1) _
                                , Address:=FcestaPdfNaSiet & "\" & NazovPlanuUpinania & ".pdf", _
                                SubAddress:="", ScreenTip:="Otvori� pl�n up�nania v PDF" & vbCrLf & _
                                FcestaPdfNaSiet & "\" & NazovPlanuUpinania & ".pdf" & vbCrLf & _
                                DatumUpravy
                                
                                 
                            End If

                   End Select
                    
'            End If
End Sub

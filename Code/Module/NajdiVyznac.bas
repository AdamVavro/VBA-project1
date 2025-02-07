Attribute VB_Name = "NajdiVyznac"
'OTVORI SUBOR NAJDE VNOM CISLO NASTROJA AVYZNACI MOZNOS5 UPNUTIA DO LISU
Sub OtvorNajdiVyznacAll() 'FUNGUJE Pre NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_Plán upínania"
    I = MsgBox("Prajete si vyznaèi možnos upnutia nástroja v súbore 'AIO_Data'  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "Vyznaèi možnos upnutia?")
    Select Case I
        Case vbNo
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
            Call OtvoriParamNastrADoplniOstatneUdajeDoAIO_Data
'           MsgBox ("Nie")
        Case vbYes
            
            'Fcesta = "C:\Users\lisy\Desktop\Plány upínania\PU_NOVÉ"
            'FcestaJPG = "C:\Users\lisy\Desktop\Plány upínania\PU_NOVÉ\PU_JPG"
            'FcestaPDF = "C:\Users\lisy\Desktop\Plány upínania\PU_NOVÉ\PU_PDF"
            FCisloNastroja = Sheets("AIO_Plan").Range("S1").Text
            FOperacia = Sheets("AIO_Plan").Range("AM1").Text
            FKrok = Sheets("AIO_Plan").Range("AM3").Text
            FCisloDielu = Sheets("AIO_Plan").Range("S3").Text
            
            If Range("AM3").Text = "" Then
                NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_Plán upínania"
            Else
                NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_S" & FKrok & "_Plán upínania"
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
            
'                MsgBox "Súbor AIO_Data je už otvorený"
                
        'Zmensí okno excelu na lavú polovicu obrazovky
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                Application.WindowState = xlNormal
'''''                Application.Left = 1 '226
'''''                Application.Top = 1
'''''                Application.Width = 601 '668 '686 '(976)
'''''                Application.Height = 870
                
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                
        'Skryje riadok vzorcov
'                Application.DisplayFormulaBar = False
                '
        'Skryje záhlavia
'                ActiveWindow.DisplayHeadings = False
                
            'Zmensí okno excelu na pravú polovicu obrazovky
'''''                Application.WindowState = xlNormal
'''''                Application.Left = 602 '226
'''''                Application.Top = 1
'''''                Application.Width = 601 '754
'''''                Application.Height = 870
            Else
            
                'UserForm1.Unload
                'Application.WindowState = xlMaximized
            
'                MsgBox "Otváram súbor AIO_Data"
                
            'Zmensí okno excelu na lavú polovicu obrazovky
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                Application.WindowState = xlNormal
'''''                Application.Left = 1 '226
'''''                Application.Top = 1
'''''                Application.Width = 601 '668 '686 '(976)
'''''                Application.Height = 870

                
'                Workbooks.Open FileName:="C:\Users\lisy\Desktop\Plány upínania\Parametre nástrojov\Parametre nástrojov.xlsm"
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                
        'Skryje riadok vzorcov
'                Application.DisplayFormulaBar = False
                '
        'Skryje záhlavia
'                ActiveWindow.DisplayHeadings = False
                
            'Zmensí okno excelu na pravú polovicu obrazovky
'''''                Application.WindowState = xlNormal
'''''                Application.Left = 602 '226
'''''                Application.Top = 1
'''''                Application.Width = 601 '754
'''''                Application.Height = 870
            End If
            'OK-------------------------------
            
        'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
'            CisloNastroja = Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("S1").Value
            MsgBox ("H¾adané èíslo nástroja: " & FCisloNastroja)
            'OK-------------------------------
            Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(FCisloNastroja, LookIn:=xlValues, SearchFormat:=False)
            If OblastNajdi Is Nothing Then
                I = MsgBox("Èíslo nástroja sa nenašlo!", vbOKOnly + vbExclamation, "Èíslo nástroja")
            Else
                OblastNajdi.Select
                'MsgBox (OblastNajdi.Address)
                'MsgBox (OblastNajdi.Row)
                'MsgBox (OblastNajdi.Column)
                L1 = Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("Y7").Value
                L2 = Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("AA7").Value
                L3 = Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("AC7").Value
                L4 = Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("AE7").Value
                
                OblastNajdiL1 = (OblastNajdi.Column - 6)
                OblastNajdiL2 = (OblastNajdi.Column - 5)
                OblastNajdiL3 = (OblastNajdi.Column - 4)
                OblastNajdiL4 = (OblastNajdi.Column - 3)
                ''''''''
                OblastNajdiNPU = (OblastNajdi.Column - 2)
                
                'MsgBox (OblastNajdiL4)
                Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiNPU).Select
                Selection.Interior.Color = RGB(210, 245, 45)
                '''''''''
                
                If L1 = True Then
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiL1) = "z"
                Else: Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiL1) = "n"
                End If
                
                If L2 = True Then
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiL2) = "z"
                Else: Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiL2) = "n"
                End If
                
                If L3 = True Then
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiL3) = "z"
                Else: Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiL3) = "n"
                End If
                
                If L4 = True Then
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiL4) = "z"
                Else: Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiL4) = "n"
                End If
            
            End If
            'OK---------------------------
            'FUNGUJE
            
            Call DoplniOstatneUdajeDoAIO_Data
            
'            MsgBox (Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("AN28").Value)
            If Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("AN28").Value > 0 Or Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("AN29").Value > 0 Then
'                MsgBox ("Spúštam DoplniRasterStolaDoAIO_Data")
                Call DoplniRasterStolaDoAIO_Data
'            Else
'                MsgBox ("Nepúštam DoplniRasterStolaDoAIO_Data")
            End If
            
    End Select

End Sub

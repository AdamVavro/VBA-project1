Attribute VB_Name = "Import�dajov"
Public IpmortVsetkychUdajovZAIO_Data_Running As Boolean 'Premenn� v Sub IpmortVsetkychUdajovZAIO_Data


Sub DoplniUdaje()
'Do formul�ra dopln� nasleduj�ce �daje o n�stroji zo s�boru 'AIO_Data' _
                            Oper�cia: _
                            Stufe: _
                            ��slo dielu: _
                            VP: _
                            Ozna�enie dielu: _
                            N�zov projektu:

    If Worksheets("AIO_Plan").Range("S1") = "" Then
'            MsgBox ("Nerob�m ni�")
    Else:    I = MsgBox("Prajete si doplni� �daje o n�stroji zo s�boru 'AIO_Data'  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "Doplni� �daje?")
        
        Select Case I
            Case vbNo
                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    '           MsgBox ("Nie")
                Call ZavolaOtvori�Formular
            
            Case vbYes
            
                Application.ScreenUpdating = False   'vypne prekreslovanie obrazovky, t�m sa makro zr�chli
                
            'OTVORI SUBOR "AIO_Data"
                On Error Resume Next
                Set Zosit = Workbooks("AIO_Data")
                ZositOtvoreny = Not Zosit Is Nothing
                If ZositOtvoreny = True Then
    
'                    MsgBox "S�bor AIO_Data je u� otvoren�"
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                'Zmens� okno excelu
'''''                    Application.WindowState = xlNormal
'''''                    Application.Left = 226
'''''                    Application.Top = 362 '1
'''''                    Application.Width = 686
'''''                    Application.Height = 508 '870
                    
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=2 'Rozbal� zoskupen� udaje

                Else
    
'                    MsgBox "Otv�ram s�bor AIO_Data"
'                    Workbooks.Open FileName:="C:\Users\lisy\Desktop\Pl�ny up�nania\Parametre n�strojov\Parametre n�strojov.xlsm"
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                'Zmens� okno excelu
'''''                    Application.WindowState = xlNormal
'''''                    Application.Left = 226
'''''                    Application.Top = 362 '1
'''''                    Application.Width = 686
'''''                    Application.Height = 508 '870
                    
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=2 'Rozbal� zoskupen� udaje

                End If
    '
            'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
                Application.ScreenUpdating = True
                
                CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
'                MsgBox ("H�adan� ��slo n�stroja: " & CisloNastroja)
                Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
                If OblastNajdi Is Nothing Then
                    I = MsgBox("��slo n�stroja sa nena�lo!", vbOKOnly + vbExclamation, "��slo n�stroja")
                Else
                    OblastNajdi.Select
'                    MsgBox (CisloNastroja & vbCrLf & _
                            "Adresa h�adanej bunky:  " & OblastNajdi.Address & vbCrLf & _
                            "Riadok h�adanej bunky:  " & OblastNajdi.Row & vbCrLf & _
                            "St�pec h�adanej bunky:    " & OblastNajdi.Column)
                    
                    OblastNajdiOP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column + 1).Value
                    OblastNajdiS = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column + 2).Value
                    OblastNajdiCisloDielu = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column + 3).Value
                    OblastNajdiVP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column + 4).Value
                    OblastNajdiOznacenieDielu = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column + 5).Value
                    OblastNajdiNazovProjektu = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column + 6).Value
                    
'                    MsgBox (CisloNastroja & vbCrLf & _
                            "Oper�cia:                " & OblastNajdiOP & vbCrLf & _
                            "Stufe:                       " & OblastNajdiS & vbCrLf & _
                            "��slo dielu:              " & OblastNajdiCisloDielu & vbCrLf & _
                            "VP:                           " & OblastNajdiVP & vbCrLf & _
                            "Ozna�enie dielu:    " & OblastNajdiOznacenieDielu & vbCrLf & _
                            "N�zov projektu:      " & OblastNajdiNazovProjektu)
                    
                    If OblastNajdiCisloDielu = "" Then
                        
                        'Vyhlad� prv� �tvor��slie n�stroja v koment�roch
                        Dim PrveStvorcislieNastroja As String
                        PrveStvorcislieNastroja = Mid(CisloNastroja, 1, 4)
'                        MsgBox ("Spu��am hladaj prve �tvor�islie" & vbCrLf & _
                                "H�adan� ��slo n�stroja:" & vbCrLf & _
                                "" & vbCrLf & _
                                "" & PrveStvorcislieNastroja)
                        Set OblastNajdiDVA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(PrveStvorcislieNastroja, LookIn:=xlComments)
                        
                        If OblastNajdiDVA Is Nothing Then
                        I = MsgBox("Prv� �tvor��slie n�stroja sa nena�lo!", vbOKOnly + vbExclamation, "Prv� �tvor��slie n�stroja")
                        Else
                        OblastNajdiDVA.Select
                        
'                        MsgBox (PrveStvorcislieNastroja & vbCrLf & _
                                "Adresa h�adanej bunky:  " & OblastNajdiDVA.Address & vbCrLf & _
                                "Riadok h�adanej bunky:  " & OblastNajdiDVA.Row & vbCrLf & _
                                "St�pec h�adanej bunky:    " & OblastNajdiDVA.Column)
                        
                        OblastNajdiCisloDielu = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdiDVA.Row + 1, OblastNajdiDVA.Column + 3).Value
                        OblastNajdiVP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdiDVA.Row + 1, OblastNajdiDVA.Column + 4).Value
                        OblastNajdiOznacenieDielu = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdiDVA.Row + 1, OblastNajdiDVA.Column + 5).Value
                        OblastNajdiNazovProjektu = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdiDVA.Row + 1, OblastNajdiDVA.Column + 6).Value
                        
'                        MsgBox (PrveStvorcislieNastroja & vbCrLf & _
'                            "Oper�cia:                " & OblastNajdiOP & vbCrLf & _
'                            "Stufe:                       " & OblastNajdiS & vbCrLf & _
'                            "��slo dielu:              " & OblastNajdiCisloDielu & vbCrLf & _
'                            "VP:                           " & OblastNajdiVP & vbCrLf & _
'                            "Ozna�enie dielu:    " & OblastNajdiOznacenieDielu & vbCrLf & _
'                            "N�zov projektu:      " & OblastNajdiNazovProjektu)
                        End If
                    End If
                End If
                

                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            
                Application.Wait Now + TimeValue("00:00:01") 'zdr�anie 1 sekundu
        
                 MsgBox ("Pre n�stroj " & CisloNastroja & " sa doplnia nasledovn� �daje:" & vbCrLf & _
                                    "" & vbCrLf & _
                                    "Oper�cia:                " & OblastNajdiOP & vbCrLf & _
                                    "Stufe:                       " & OblastNajdiS & vbCrLf & _
                                    "��slo dielu:              " & OblastNajdiCisloDielu & vbCrLf & _
                                    "VP:                           " & OblastNajdiVP & vbCrLf & _
                                    "Ozna�enie dielu:    " & OblastNajdiOznacenieDielu & vbCrLf & _
                                    "N�zov projektu:      " & OblastNajdiNazovProjektu)
        '
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AM1").Value = OblastNajdiOP
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AM3").Value = OblastNajdiS
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S3").Value = OblastNajdiCisloDielu
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("C4").Value = OblastNajdiVP
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("K4").Value = OblastNajdiOznacenieDielu
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AB4").Value = OblastNajdiNazovProjektu
                
                Application.Wait Now + TimeValue("00:00:01") 'zdr�anie 1 sekundu
                Application.DisplayAlerts = False 'Zak�e zobrezeniu syst�mov�ch hl�ok
                Application.ScreenUpdating = False
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
                
                '--------------------------------------------------
                'Spyta sa �i chem doplnit vsetky ostatne udaje
    
    '           18 D�kaN�stroja
                OblastNajdiD�kaN�stroja = (OblastNajdi.Column + 18)
    
                If Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiD�kaN�stroja) = "" Then
'                    MsgBox ("Bunka ��D�ka n�stroja�� je pr�zna, nerob�m ni�")
                Else:
                    I = MsgBox("Prajete si doplni� v�etky ostatn� �daje o n�stroji zo s�boru 'AIO_Data'  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "Doplni� �daje?")
    
                    Select Case I
                        Case vbNo
                            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                '           MsgBox ("Nie")
            '
                        Case vbYes
    '                       NACITA VSETKY UDAJE O NASTROJI ZO SUBORU "AIO_Data" A DOPLNI ICH DO FORMULARA
                            Call IpmortVsetkychUdajovZAIO_Data
                    End Select
                End If
                '----------------------------------------------
'                Workbooks("AIO_Data").Close
                
                Application.ScreenUpdating = True
                Application.DisplayAlerts = True 'Povol� zobrezeniu syst�mov�ch hl�ok

        End Select
'
    End If
    
End Sub

Sub IpmortVsetkychUdajovZAIO_Data()
'NACITA VSETKY UDAJE O NASTROJI ZO SUBORU "AIO_Data" A DOPLNI ICH DO FORMULARA
'    MsgBox ("IpmortVsetkychUdajovZAIO_Data") 'OK


'   Po�as priebehu proced�ry "IpmortVsetkychUdajovZAIO_Data_Running" m� premenn� hodnotu "TRUE" po skonen� proced�ri _
    sa zmen� hodnota premenn�j na "FALSE"
    IpmortVsetkychUdajovZAIO_Data_Running = True
            
    Application.ScreenUpdating = False   'vypne prekreslovanie obrazovky, t�m sa makro zr�chli

'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

'    MsgBox ("OtvoriSuborAIO_Data") 'OK
    'OTVORENIE SUBORU "AIO_Data" resp. Overenie, �i u� nie je otvoren�

'OTVORI SUBOR "AIO_Data"
    On Error Resume Next
    Set Zosit = Workbooks("AIO_Data")
    ZositOtvoreny = Not Zosit Is Nothing
    If ZositOtvoreny = True Then

'        MsgBox "S�bor AIO_Data je u� otvoren�"
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
    'Zmens� okno excelu
'''''        Application.WindowState = xlNormal
'''''        Application.Left = 226
'''''        Application.Top = 362 '1
'''''        Application.Width = 686
'''''        Application.Height = 508 '870
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=2 'Rozbal� zoskupen� udaje

    Else

'        MsgBox "Otv�ram s�bor AIO_Data"
'        Workbooks.Open FileName:="C:\Users\lisy\Desktop\Pl�ny up�nania\Parametre n�strojov\Parametre n�strojov.xlsm"
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
    'Zmens� okno excelu
'''''        Application.WindowState = xlNormal
'''''        Application.Left = 226
'''''        Application.Top = 362 '1
'''''        Application.Width = 686
'''''        Application.Height = 508 '870
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=2 'Rozbal� zoskupen� udaje

    End If

'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

'    MsgBox ("ImportVsetkychUdajovNastrojaZAIO_Data") 'OK
     'IMPORT VSETKYCH UDAJOV O NASTROJI ZO SUBORU 'AIO_Data'

'---------------------------------------------------------------------------------------
'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
    Application.ScreenUpdating = True
    
    CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
'   MsgBox ("H�adan� ��slo n�stroja: " & CisloNastroja)
    Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
    If OblastNajdi Is Nothing Then
        I = MsgBox("��slo n�stroja sa nena�lo!", vbOKOnly + vbExclamation, "��slo n�stroja")
    Else
'---------------------------------------------------------------------------------------
        
'        MsgBox ("ImportPracovneTlakyANastavenia") 'OK
    
'Doplni Pracovn� tlaky a nastavenia
        'Zabr�ni zobrezeniu syst�mov�ch hl�ok
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
'8 OblastNajdiTlakBarana
        OblastNajdiTlakBarana = (OblastNajdi.Column + 8)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S10").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiTlakBarana)
'                            -------------------------------------------------
'                               Skop�rovanie a prilepenie koment�ra z AIO_Data
        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiTlakBarana).Copy
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S10").PasteSpecial Paste:=xlPasteComments, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'                            -------------------------------------------------
'9 OblastNajdiHP
        OblastNajdiHP = (OblastNajdi.Column + 9)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S13").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiHP)
'                            -------------------------------------------------
'                               Skop�rovanie a prilepenie koment�ra z AIO_Data
        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiHP).Copy
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S13").PasteSpecial Paste:=xlPasteComments, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'                            -------------------------------------------------
'10 OblastNajdiDP
        OblastNajdiDP = (OblastNajdi.Column + 10)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S12").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiDP)
'                            -------------------------------------------------
'                               Skop�rovanie a prilepenie koment�ra z AIO_Data
        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiDP).Copy
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S12").PasteSpecial Paste:=xlPasteComments, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'                            -------------------------------------------------
'16 Brzda
        OblastNajdiBrzda = (OblastNajdi.Column + 16)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S9").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiBrzda)
'                            -------------------------------------------------
'                               Skop�rovanie a prilepenie koment�ra z AIO_Data
        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiBrzda).Copy
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S9").PasteSpecial Paste:=xlPasteComments, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'                            -------------------------------------------------
'17 �apyNadSt�l
        OblastNajdi�apyNadSt�l = (OblastNajdi.Column + 17)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S11").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi�apyNadSt�l)
'                            -------------------------------------------------
'                               Skop�rovanie a prilepenie koment�ra z AIO_Data
        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi�apyNadSt�l).Copy
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S11").PasteSpecial Paste:=xlPasteComments, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'                            -------------------------------------------------

        Application.DisplayAlerts = True 'Povol� zobrezeniu syst�mov�ch hl�ok
        Application.ScreenUpdating = True
        '---------------------------------------------------------------------------------------
        
'        MsgBox ("ImportUdajovZUserForm1AVahy") 'OK
'        Doplni v�etky ostatn� �daje

'18 D�kaN�stroja
        OblastNajdiD�kaN�stroja = (OblastNajdi.Column + 18)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("W5").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiD�kaN�stroja)
'19 ��rkaN�stroja
        OblastNajdi��rkaN�stroja = (OblastNajdi.Column + 19)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AD5").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi��rkaN�stroja)
'20 V��kaN�stroja
        OblastNajdiV��kaN�stroja = (OblastNajdi.Column + 20)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AK5").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiV��kaN�stroja)
'21 Vzdialenos�MedziDr�kamiOT
        OblastNajdiVzdialenos�MedziDr�kamiOT = (OblastNajdi.Column + 21)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AI5").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVzdialenos�MedziDr�kamiOT)
'22 OdstavovaciePrvkyGDF
        OblastNajdiOdstavovaciePrvkyGDF = (OblastNajdi.Column + 22)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("R5").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiOdstavovaciePrvkyGDF)
'23 OdstavovaciePrvkyOB
        OblastNajdiOdstavovaciePrvkyOB = (OblastNajdi.Column + 23)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S5").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiOdstavovaciePrvkyOB)
'24 OdstavovaciePrvkyZdvih_Vy�ka
        OblastNajdiOdstavovaciePrvkyZdvih_Vy�ka = (OblastNajdi.Column + 24)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("W6").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiOdstavovaciePrvkyZdvih_Vy�ka)
'25 Up�naciaV��kaN�stroja
        OblastNajdiUpncV�kNstrj = (OblastNajdi.Column + 25)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AK6").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiUpncV�kNstrj)
'26 Pridr�iava�BaranBez
        OblastNajdiPridr�iava�BaranBez = (OblastNajdi.Column + 26)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("J13").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridr�iava�BaranBez)
'27 Pridr�iava�Baran�apy
        OblastNajdiPridr�iava�Baran�apy = (OblastNajdi.Column + 27)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("N13").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridr�iava�Baran�apy)
'28 Pridr�iava�BaranGDF
        OblastNajdiPridr�iava�BaranGDF = (OblastNajdi.Column + 28)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("R13").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridr�iava�BaranGDF)
'29 Pridr�iava�St�lBez
        OblastNajdiPridr�iava�St�lBez = (OblastNajdi.Column + 29)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("J12").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridr�iava�St�lBez)
'30 Pridr�iava�St�l�apy
        OblastNajdiPridr�iava�St�l�apy = (OblastNajdi.Column + 30)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("N12").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridr�iava�St�l�apy)
'31 OblastNajdiPridr�iava�St�lGDF
        OblastNajdiPridr�iava�St�lGDF = (OblastNajdi.Column + 31)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("R12").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridr�iava�St�lGDF)
'32 OblastNajdiM�ns�UpntNstrjDLs1 (Mo�nos� upnutia n�stroja do lisu1)
        OblastNajdiM�ns�UpntNstrjDLs1 = (OblastNajdi.Column + 32)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("Y7").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiM�ns�UpntNstrjDLs1)
'33 OblastNajdiM�ns�UpntNstrjDLs2 (Mo�nos� upnutia n�stroja do lisu2)
        OblastNajdiM�ns�UpntNstrjDLs2 = (OblastNajdi.Column + 33)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AA7").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiM�ns�UpntNstrjDLs2)
'34 OblastNajdiM�ns�UpntNstrjDLs3 (Mo�nos� upnutia n�stroja do lisu3)
        OblastNajdiM�ns�UpntNstrjDLs3 = (OblastNajdi.Column + 34)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AC7").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiM�ns�UpntNstrjDLs3)
'35 OblastNajdiM�ns�UpntNstrjDLs4 (Mo�nos� upnutia n�stroja do lisu4)
        OblastNajdiM�ns�UpntNstrjDLs4 = (OblastNajdi.Column + 35)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AE7").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiM�ns�UpntNstrjDLs4)
'36 OblastNajdiCntrvnNstrjPrmrLH (Centrovanie n�stroja priemer �av� horn�)
        OblastNajdiCntrvnNstrjPrmrLH = (OblastNajdi.Column + 36)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("P28").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiCntrvnNstrjPrmrLH)
'37 OblastNajdiCntrvnNstrjPrmrPH (Centrovanie n�stroja priemer prav� horn�)
        OblastNajdiCntrvnNstrjPrmrPH = (OblastNajdi.Column + 37)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("Z28").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiCntrvnNstrjPrmrPH)
'38 OblastNajdiCntrvnNstrjPrmrLD (Centrovanie n�stroja priemer �av� doln�)
        OblastNajdiCntrvnNstrjPrmrLD = (OblastNajdi.Column + 38)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("P31").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiCntrvnNstrjPrmrLD)
'39 OblastNajdiCntrvnNstrjPrmrPD (Centrovanie n�stroja priemer prav� doln�)
        OblastNajdiCntrvnNstrjPrmrPD = (OblastNajdi.Column + 39)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("Z31").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiCntrvnNstrjPrmrPD)
'40 OblastNajdiSrdncCntrvnLHR (S�radnice centrovania �av� horn� �tvrtina riadok)
        OblastNajdiSrdncCntrvnLHR = (OblastNajdi.Column + 40)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("T28").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnLHR)
'41 OblastNajdiSrdncCntrvnLHS (S�radnice centrovania �av� horn� �tvrtina st�pec)
        OblastNajdiSrdncCntrvnLHS = (OblastNajdi.Column + 41)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S29").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnLHS)
'42 OblastNajdiSrdncCntrvnPHR
        OblastNajdiSrdncCntrvnPHR = (OblastNajdi.Column + 42)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("V28").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnPHR)
'43 OblastNajdiSrdncCntrvnPHS
        OblastNajdiSrdncCntrvnPHS = (OblastNajdi.Column + 43)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("W29").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnPHS)
'44 OblastNajdiSrdncCntrvnLDR
        OblastNajdiSrdncCntrvnLDR = (OblastNajdi.Column + 44)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("T31").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnLDR)
'45 OblastNajdiSrdncCntrvnLDS
        OblastNajdiSrdncCntrvnLDS = (OblastNajdi.Column + 45)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S30").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnLDS)
'46 OblastNajdiSrdncCntrvnPDR
        OblastNajdiSrdncCntrvnPDR = (OblastNajdi.Column + 46)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("V31").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnPDR)
'47 OblastNajdiSrdncCntrvnPDS
        OblastNajdiSrdncCntrvnPDS = (OblastNajdi.Column + 47)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("W30").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnPDS)
'48 OblastNajdiSmrLsvnL (Smer lisovania �av�)
'        MsgBox ("Smer lisovania")
        OblastNajdiSmrLsvnL = (OblastNajdi.Column + 48)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("C41").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSmrLsvnL)
'        ------------------
'        Zobraz� alebo nezobraz� smer lisovania
        If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("C41").Value = True Then
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 22", _
            "Straight Arrow Connector 21")).Visible = msoTrue
        Else: Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 22", _
            "Straight Arrow Connector 21")).Visible = msoFalse
        End If
'        ------------------
'49 OblastNajdiSmrLsvnH (Smer lisovania hore)
        OblastNajdiSmrLsvnH = (OblastNajdi.Column + 49)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("U32").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSmrLsvnH)
'        ------------------
'        Zobraz� alebo nezobraz� smer lisovania
        If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("U32").Value = True Then
'            MsgBox ("Smer lisovania hore true")
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 3", _
            "Straight Arrow Connector 13")).Visible = msoTrue
        Else: Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 3", _
            "Straight Arrow Connector 13")).Visible = msoFalse
        End If
'        ------------------
'50 OblastNajdiSmrLsvnD (Smer lisovania dole)
        OblastNajdiSmrLsvnD = (OblastNajdi.Column + 50)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("U50").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSmrLsvnD)
'        ------------------
'        Zobraz� alebo nezobraz� smer lisovania
        If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("U50").Value = True Then
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 23", _
            "Straight Arrow Connector 24")).Visible = msoTrue
        Else: Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 23", _
            "Straight Arrow Connector 24")).Visible = msoFalse
        End If
'        ------------------
'51 OblastNajdiSmrLsvnP (Smer lisovania prav�)
        OblastNajdiSmrLsvnP = (OblastNajdi.Column + 51)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AM41").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSmrLsvnP)
'        ------------------
'        Zobraz� alebo nezobraz� smer lisovania
        If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AM41").Value = True Then
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 19", _
            "Straight Arrow Connector 20")).Visible = msoTrue
        Else: Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 19", _
            "Straight Arrow Connector 20")).Visible = msoFalse
        End If
'        ------------------
'52 OblastNajdiPznmkRdk1 (Pozn�mky k n�stroju )
        OblastNajdiPznmkRdk1 = (OblastNajdi.Column + 52)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("I14").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1)
        
'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach

        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("I14").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk1")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk1")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------

        
'53 OblastNajdiPznmkRdk2 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk2 = (OblastNajdi.Column + 53)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B15").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
'        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B15").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk2")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk2")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------

'54 OblastNajdiPznmkRdk3 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk3 = (OblastNajdi.Column + 54)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B16").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B16").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk3")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk3")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------

'55 OblastNajdiPznmkRdk4 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk4 = (OblastNajdi.Column + 55)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B17").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B17").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk4")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk4")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------
'56 OblastNajdiPznmkRdk5 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk5 = (OblastNajdi.Column + 56)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B18").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B18").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk5")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk5")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------

'57 OblastNajdiPznmkRdk6 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk6 = (OblastNajdi.Column + 57)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B19").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B19").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk6")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk6")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------
'58 OblastNajdiPznmkRdk7 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk7 = (OblastNajdi.Column + 58)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B20").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B20").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk7")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk7")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------

'59 OblastNajdiPznmkRdk8 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk8 = (OblastNajdi.Column + 59)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B21").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B21").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk8")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk8")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------

'60 OblastNajdiPznmkRdk9 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk9 = (OblastNajdi.Column + 60)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B22").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B22").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk9")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk9")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------

'61 OblastNajdiPznmkRdk10 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk10 = (OblastNajdi.Column + 61)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B23").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B23").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk10")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk10")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------

'62 OblastNajdiPznmkRdk11 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk11 = (OblastNajdi.Column + 62)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B24").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B24").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk11")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk11")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------

'63 OblastNajdiPznmkRdk12 (Pozn�mky k n�stroju )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiPznmkRdk12 = (OblastNajdi.Column + 63)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B25").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12)

'                    '------------
         'IMPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn�mkach
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        IC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12).Interior.Color
        FC = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12).Font.Color
        HA = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12).HorizontalAlignment
        IP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12).Interior.Pattern

'                    MsgBox ("Interior.Color: " & IC) 'OK
'                    MsgBox ("Font.Color: " & FC) 'OK
'                    MsgBox ("HorizontalAlignment: " & HA) 'OK
'                    MsgBox ("Iterior.Pattern: " & IP) 'OK
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B25").Select

        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        If IP <> 4000 Then
'           MsgBox ("Be�� If PznmkRdk12")
            ActiveCell.Interior.Color = IC
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
        Else:
'            MsgBox ("Be�� Else PznmkRdk12")
            ActiveCell.Font.Color = FC
            ActiveCell.HorizontalAlignment = HA
            Call VyplnZltoCervena
        End If
        Worksheets("AIO_Plan").Protect Password:="Lis.0123" '
'                        FUNGUJE
'                    '------------

'559 OblastNajdiVahaOT (V�ha OT)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiVahaOT = (OblastNajdi.Column + 559)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("G5").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVahaOT)
'560 OblastNajdiVahaUT (V�ha UT)
        OblastNajdiVahaUT = (OblastNajdi.Column + 560)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("G6").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVahaUT)
'561 OblastNajdiVahaGES (V�ha GES)
        OblastNajdiVahaGES = (OblastNajdi.Column + 561)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("G7").Value = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVahaGES)
'563 OblastNajdiGdfAleboBloky (Text vbunke "L6" Zdvih GDF/Vy�ka odstavovac�ch blokov )
        OblastNajdiGdfAleboBloky = (OblastNajdi.Column + 563)
        GdfAleboBloky = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiGdfAleboBloky).Value
'        MsgBox (GdfAleboBloky)
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("L6").Value = GdfAleboBloky
        
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'564 OblastNajdiLavaHlavicka (Lava Hlavicka "Datum vytvorenia" )
        
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        OblastNajdiLavaHlavicka = (OblastNajdi.Column + 564)
        LavaHlavicka = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiLavaHlavicka).Value
'        MsgBox (LavaHlavicka)

        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
        Worksheets("AIO_Plan").PageSetup.LeftHeader = _
        "&""Porsche Next TT""&08" & LavaHlavicka
                            
        Call SuradniceCentrovaniaLH
        Call SuradniceCentrovaniaPH
        Call SuradniceCentrovaniaLD
        Call SuradniceCentrovaniaPD
        
'---------------------------------------------------------------------------------------
        
'        MsgBox ("ImportRasterStola")

'562 OblastNajdiPctTl�nch�pv (Po�et tla�n�ch �apov )
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate

        OblastNajdiPctTl�nch�pv = (OblastNajdi.Column + 562)
        OblastNajdiPo�etCervenychCentrovacichCapov = (OblastNajdi.Column + 566)
        
        If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPctTl�nch�pv).Value = "" Or _
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPctTl�nch�pv).Value = "0" And _
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPo�etCervenychCentrovacichCapov).Value = "" Or _
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPo�etCervenychCentrovacichCapov).Value = "0" Then
        MsgBox ("Nekopirujem raster")
'        MsgBox (Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPo�etCervenychCentrovacichCapov).Value)
        
'            If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPo�etCervenychCentrovacichCapov).Value = "" Or _
'                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPo�etCervenychCentrovacichCapov).Value = "0" Then
'            MsgBox ("Nekop�rujem raster." & vbCrLf & _
'            "Po�et tla�n�ch �apov je 0." & vbCrLf & _
'            "Po�et �erven�ch centrovac�ch �apov je 0.")
'            End If
            
        Else
            MsgBox ("Kopirujem raster")
            Application.ScreenUpdating = False
            
          
'64 OblastNajdiRaster8H (Raster stola riadok 8 hore )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E34").Select
            
            OblastNajdiRaster8HoreZa�iatok = (OblastNajdi.Column + 64)
            OblastNajdiRaster8HoreKoniec = (OblastNajdi.Column + 96)

            AdresaOblastNajdiRaster8HoreZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster8HoreZa�iatok).Address
            AdresaOblastNajdiRaster8HoreKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster8HoreKoniec).Address
'                    MsgBox (AdresaOblastNajdiRaster8HoreZa�iatok)
'                    MsgBox (AdresaOblastNajdiRaster8HoreKoniec)
'                    MsgBox (AdresaOblastNajdiRaster8HoreZa�iatok & ":" & AdresaOblastNajdiRaster8HoreKoniec)
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster8HoreZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster8HoreZa�iatok & ":" & AdresaOblastNajdiRaster8HoreKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$34:$AK$34").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'                    -------
''97 OblastNajdiRaster7H (Raster stola riadok 7 hore )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E35").Select
            
            OblastNajdiRaster7HoreZa�iatok = (OblastNajdi.Column + 97)
            OblastNajdiRaster7HoreKoniec = (OblastNajdi.Column + 129)
            
            AdresaOblastNajdiRaster7HoreZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster7HoreZa�iatok).Address
            AdresaOblastNajdiRaster7HoreKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster7HoreKoniec).Address
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster7HoreZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster7HoreZa�iatok & ":" & AdresaOblastNajdiRaster7HoreKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$35:$AK$35").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
             
''130 OblastNajdiRaster6H (Raster stola riadok 6 hore )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E36").Select
            
            OblastNajdiRaster6HoreZa�iatok = (OblastNajdi.Column + 130)
            OblastNajdiRaster6HoreKoniec = (OblastNajdi.Column + 162)
            
            AdresaOblastNajdiRaster6HoreZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster6HoreZa�iatok).Address
            AdresaOblastNajdiRaster6HoreKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster6HoreKoniec).Address
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster6HoreZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster6HoreZa�iatok & ":" & AdresaOblastNajdiRaster6HoreKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$36:$AK$36").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"

''163 OblastNajdiRaster5H (Raster stola riadok 5 hore )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E37").Select
            
            OblastNajdiRaster5HoreZa�iatok = (OblastNajdi.Column + 163)
            OblastNajdiRaster5HoreKoniec = (OblastNajdi.Column + 195)
            
            AdresaOblastNajdiRaster5HoreZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster5HoreZa�iatok).Address
            AdresaOblastNajdiRaster5HoreKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster5HoreKoniec).Address
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster5HoreZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster5HoreZa�iatok & ":" & AdresaOblastNajdiRaster5HoreKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$37:$AK$37").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"

''196 OblastNajdiRaster4H (Raster stola riadok 4 hore )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E38").Select
            
            OblastNajdiRaster4HoreZa�iatok = (OblastNajdi.Column + 196)
            OblastNajdiRaster4HoreKoniec = (OblastNajdi.Column + 228)
            
            AdresaOblastNajdiRaster4HoreZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster4HoreZa�iatok).Address
            AdresaOblastNajdiRaster4HoreKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster4HoreKoniec).Address
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster4HoreZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster4HoreZa�iatok & ":" & AdresaOblastNajdiRaster4HoreKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$38:$AK$38").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"

''229 OblastNajdiRaster3H (Raster stola riadok 3 hore )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E39").Select
            
            OblastNajdiRaster3HoreZa�iatok = (OblastNajdi.Column + 229)
            OblastNajdiRaster3HoreKoniec = (OblastNajdi.Column + 261)
            
            AdresaOblastNajdiRaster3HoreZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster3HoreZa�iatok).Address
            AdresaOblastNajdiRaster3HoreKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster3HoreKoniec).Address
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster3HoreZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster3HoreZa�iatok & ":" & AdresaOblastNajdiRaster3HoreKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$39:$AK$39").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"

''262 OblastNajdiRaster2H (Raster stola riadok 2 hore )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E40").Select
            
            OblastNajdiRaster2HoreZa�iatok = (OblastNajdi.Column + 262)
            OblastNajdiRaster2HoreKoniec = (OblastNajdi.Column + 294)
            
            AdresaOblastNajdiRaster2HoreZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster2HoreZa�iatok).Address
            AdresaOblastNajdiRaster2HoreKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster2HoreKoniec).Address
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster2HoreZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster2HoreZa�iatok & ":" & AdresaOblastNajdiRaster2HoreKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$40:$AK$40").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"

''295 OblastNajdiRaster1S (Raster stola riadok 1 Stred )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E41").Select
            
            OblastNajdiRaster1StredZa�iatok = (OblastNajdi.Column + 295)
            OblastNajdiRaster1StredKoniec = (OblastNajdi.Column + 327)
            
            AdresaOblastNajdiRaster1StredZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster1StredZa�iatok).Address
            AdresaOblastNajdiRaster1StredKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster1StredKoniec).Address
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster1StredZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster1StredZa�iatok & ":" & AdresaOblastNajdiRaster1StredKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$41:$AK$41").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"

''328 OblastNajdiRaster2D (Raster stola riadok 2 dole )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E42").Select
            
            OblastNajdiRaster2DoleZa�iatok = (OblastNajdi.Column + 328)
            OblastNajdiRaster2DoleKoniec = (OblastNajdi.Column + 360)
            
            AdresaOblastNajdiRaster2DoleZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster2DoleZa�iatok).Address
            AdresaOblastNajdiRaster2DoleKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster2DoleKoniec).Address
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster2DoleZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster2DoleZa�iatok & ":" & AdresaOblastNajdiRaster2DoleKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$42:$AK$42").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"

'''361 OblastNajdiRaster3D (Raster stola riadok 3 dole )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E43").Select
            
            OblastNajdiRaster3DoleZa�iatok = (OblastNajdi.Column + 361)
            OblastNajdiRaster3DoleKoniec = (OblastNajdi.Column + 393)
            
            AdresaOblastNajdiRaster3DoleZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster3DoleZa�iatok).Address
            AdresaOblastNajdiRaster3DoleKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster3DoleKoniec).Address
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster3DoleZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster3DoleZa�iatok & ":" & AdresaOblastNajdiRaster3DoleKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$43:$AK$43").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
            
'''394 OblastNajdiRaster4D (Raster stola riadok 4 dole )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E44").Select
            
            OblastNajdiRaster4DoleZa�iatok = (OblastNajdi.Column + 394)
            OblastNajdiRaster4DoleKoniec = (OblastNajdi.Column + 426)
            
            AdresaOblastNajdiRaster4DoleZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster4DoleZa�iatok).Address
            AdresaOblastNajdiRaster4DoleKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster4DoleKoniec).Address
            
            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster4DoleZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster4DoleZa�iatok & ":" & AdresaOblastNajdiRaster4DoleKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$44:$AK$44").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'
'''427 OblastNajdiRaster5D (Raster stola riadok 5 dole )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E45").Select

            OblastNajdiRaster5DoleZa�iatok = (OblastNajdi.Column + 427)
            OblastNajdiRaster5DoleKoniec = (OblastNajdi.Column + 459)

            AdresaOblastNajdiRaster5DoleZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster5DoleZa�iatok).Address
            AdresaOblastNajdiRaster5DoleKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster5DoleKoniec).Address

            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster5DoleZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster5DoleZa�iatok & ":" & AdresaOblastNajdiRaster5DoleKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$45:$AK$45").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'
'''460 OblastNajdiRaster6D (Raster stola riadok 6 dole )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E46").Select

            OblastNajdiRaster6DoleZa�iatok = (OblastNajdi.Column + 460)
            OblastNajdiRaster6DoleKoniec = (OblastNajdi.Column + 492)

            AdresaOblastNajdiRaster6DoleZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster6DoleZa�iatok).Address
            AdresaOblastNajdiRaster6DoleKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster6DoleKoniec).Address

            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster6DoleZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster6DoleZa�iatok & ":" & AdresaOblastNajdiRaster6DoleKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$46:$AK$46").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'
'''493 OblastNajdiRaster7D (Raster stola riadok 7 dole )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E47").Select

            OblastNajdiRaster7DoleZa�iatok = (OblastNajdi.Column + 493)
            OblastNajdiRaster7DoleKoniec = (OblastNajdi.Column + 525)

            AdresaOblastNajdiRaster7DoleZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster7DoleZa�iatok).Address
            AdresaOblastNajdiRaster7DoleKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster7DoleKoniec).Address

            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster7DoleZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster7DoleZa�iatok & ":" & AdresaOblastNajdiRaster7DoleKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$47:$AK$47").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'
'''526 OblastNajdiRaster8D (Raster stola riadok 8 dole )
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("E48").Select

            OblastNajdiRaster8DoleZa�iatok = (OblastNajdi.Column + 526)
            OblastNajdiRaster8DoleKoniec = (OblastNajdi.Column + 558)

            AdresaOblastNajdiRaster8DoleZa�iatok = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster8DoleZa�iatok).Address
            AdresaOblastNajdiRaster8DoleKoniec = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster8DoleKoniec).Address

            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster8DoleZa�iatok).Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Range(AdresaOblastNajdiRaster8DoleZa�iatok & ":" & AdresaOblastNajdiRaster8DoleKoniec).Select
            Selection.Copy
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$48:$AK$48").Select
            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Paste
            Application.CutCopyMode = False
            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'
        End If

        
        '---------------------------------------------------------------------------------------
        Call SpocitaCerveneCentrovacieCapy
              
        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Select

'------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------

        Application.Wait Now + TimeValue("00:00:01") 'zdr�anie 1 sekundu
        Application.DisplayAlerts = False 'Zak�e zobrezeniu syst�mov�ch hl�ok
        
'        Workbooks("AIO_Data").Close
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True 'Povol� zobrezeniu syst�mov�ch hl�ok

    End If
    
    IpmortVsetkychUdajovZAIO_Data_Running = False

End Sub


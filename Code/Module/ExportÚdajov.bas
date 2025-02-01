Attribute VB_Name = "Export⁄dajov"
'PO VLOZENÌ ALEBO ZMENENÌ UDAJU TLAK BARANA VLOZI ALEBO ZMENÌ UDAJ TLAK BARANA AJ V S˙BORE "AIO_Data"
Sub ZmeniTlakBarana()

'    If Worksheets("AIO_Plan").PageSetup.LeftHeader = "" Then
'        MsgBox ("ºav· hlaviËka pr·zdnam, neprepisujem tlak barana v AIO_Data")

    If IpmortVsetkychUdajovZAIO_Data_Running = True Then

'       MsgBox ("Sub IpmortVsetkychUdajovZAIO_Data is running!Neprepisujem ˙daje TLAK BARANA v Paremetre n·strojov")

    Else
        I = MsgBox("Prajete si prepÌsaù tlak barana v s˙bore 'AIO_Data'  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "PrepÌsaù tlaky?")
            
            Select Case I
                Case vbNo
                    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
        '           MsgBox ("Nie")
                Case vbYes
        '
                'OTVORI SUBOR "AIO_Data"
                    On Error Resume Next
                    Set Zosit = Workbooks("AIO_Data")
                    ZositOtvoreny = Not Zosit Is Nothing
                    If ZositOtvoreny = True Then
        
    '                    MsgBox "S˙bor AIO_Data je uû otvoren˝"
                        
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
'''''
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
          
                    Else
        
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
        
    '                    MsgBox "Otv·ram s˙bor AIO_Data"
'                        Workbooks.Open FileName:="C:\Users\lisy\Desktop\Pl·ny upÌnania\Parametre n·strojov\Parametre n·strojov.xlsm"
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
                    
                    End If
        '
                'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
        
    '                CisloNastroja = Worksheets("AIO_Plan").Range("S1").Value
                    CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
    '                MsgBox ("HæadanÈ ËÌslo n·stroja: " & CisloNastroja)
                    Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
                    If OblastNajdi Is Nothing Then
                        I = MsgBox("»Ìslo n·stroja sa nenaölo!", vbOKOnly + vbExclamation, "»Ìslo n·stroja")
                    Else
                        OblastNajdi.Select
    '                    MsgBox (CisloNastroja & vbCrLf & _
                                "Adresa hæadanej bunky:  " & OblastNajdi.Address & vbCrLf & _
                                "Riadok hæadanej bunky:  " & OblastNajdi.Row & vbCrLf & _
                                "StÂpec hæadanej bunky:    " & OblastNajdi.Column)
                    
                        OblastNajdiTlakBaran = (OblastNajdi.Column + 8)
                        StaryTlakBaran = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiTlakBaran).Value
                        NovyTlakBaran = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S10").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiTlakBaran).Select
                        MsgBox ("Tlak na barana sa prepÌöe z " & StaryTlakBaran & " na " & NovyTlakBaran & "." & vbCrLf & _
                                "StarÈ tlaky s˙ zapÌsanÈ v koment·ri.")
                        
    '                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiTlakBaran) = Worksheets("AIO_Plan").Range("S10").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiTlakBaran) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S10").Value
                       
    '                    Workbooks("AIO_Data").Close
                    End If
            End Select
    End If

End Sub
'PO VLOZENÌ ALEBO ZMENENÌ UDAJU TLAK HORNEHO PRIDRZIAVACA VLOZI ALEBO ZMENÌ UDAJ TLAK HORNEHO PRIDRZIAVACA AJ V S˙BORE "AIO_Data"
Sub ZmeniTlakHP()

'    If Worksheets("AIO_Plan").PageSetup.LeftHeader = "" Then
'        MsgBox ("ºav· hlaviËka pr·zdnam, neprepisujem tlaky HP v AIO_Data")

    If IpmortVsetkychUdajovZAIO_Data_Running = True Then

'       MsgBox ("Sub IpmortVsetkychUdajovZAIO_Data is running!Neprepisujem ˙daje TLAK HP v Paremetre n·strojov")
       
    Else
    
        I = MsgBox("Prajete si prepÌsaù tlak hornÈho pridrûiavaËa v s˙bore 'AIO_Data'  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "PrepÌsaù tlaky?")
            
            Select Case I
                Case vbNo
                    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
        '           MsgBox ("Nie")
                Case vbYes
        '
                'OTVORI SUBOR "AIO_Data"
                    On Error Resume Next
                    Set Zosit = Workbooks("AIO_Data")
                    ZositOtvoreny = Not Zosit Is Nothing
                    If ZositOtvoreny = True Then
        
    '                    MsgBox "S˙bor AIO_Data je uû otvoren˝"
                        
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
                            
                            Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
        
                    Else
        
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
    
    '                    MsgBox "Otv·ram s˙bor AIO_Data"
'                        Workbooks.Open FileName:="C:\Users\lisy\Desktop\Pl·ny upÌnania\Parametre n·strojov\Parametre n·strojov.xlsm"
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
''''''                        Application.WindowState = xlNormal
''''''                        Application.Left = 602 '226
''''''                        Application.Top = 1
''''''                        Application.Width = 601 '754
''''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
        
                    End If
        '
                'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    
    '                CisloNastroja = Worksheets("AIO_Plan").Range("S1").Value
                    CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
    '                MsgBox ("HæadanÈ ËÌslo n·stroja: " & CisloNastroja)
                    Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
                    If OblastNajdi Is Nothing Then
                        I = MsgBox("»Ìslo n·stroja sa nenaölo!", vbOKOnly + vbExclamation, "»Ìslo n·stroja")
                    Else
                        OblastNajdi.Select
    '                    MsgBox (CisloNastroja & vbCrLf & _
                                "Adresa hæadanej bunky:  " & OblastNajdi.Address & vbCrLf & _
                                "Riadok hæadanej bunky:  " & OblastNajdi.Row & vbCrLf & _
                                "StÂpec hæadanej bunky:    " & OblastNajdi.Column)
                    
                        OblastNajdiHP = (OblastNajdi.Column + 9)
                        StaryTlakHP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiHP).Value
                        NovyTlakHP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S13").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiHP).Select
                        MsgBox ("Tlak na horn˝ pridrûiavaË sa prepÌöe z " & StaryTlakHP & " na " & NovyTlakHP & "." & vbCrLf & _
                                "StarÈ tlaky s˙ zapÌsanÈ v koment·ri.")
    '                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiHP) = Worksheets("AIO_Plan").Range("S13").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiHP) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S13").Value
                       
    '                    Workbooks("AIO_Data").Close
                    End If
            End Select
    End If

End Sub
'PO VLOZENÌ ALEBO ZMENENÌ UDAJU TLAK DOLNEHO PRIDRZIAVACA VLOZI ALEBO ZMENÌ UDAJ TLAK SPODNEHO PRIDRZIAVACA AJ V S˙BORE "AIO_Data"
Sub ZmeniTlakDP()

'    If Worksheets("AIO_Plan").PageSetup.LeftHeader = "" Then
'        MsgBox ("ºav· hlaviËka pr·zdnam, neprepisujem tlaky DP v AIO_Data")

    If IpmortVsetkychUdajovZAIO_Data_Running = True Then

'       MsgBox ("Sub IpmortVsetkychUdajovZAIO_Data is running!Neprepisujem ˙daje TLAK DP v Paremetre n·strojov")
       
    Else
        I = MsgBox("Prajete si prepÌsaù tlak dolnÈho pridrûiavaËa v s˙bore 'AIO_Data'  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "PrepÌsaù tlaky?")
            
            Select Case I
                Case vbNo
                    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
        '           MsgBox ("Nie")
                Case vbYes
        '
                'OTVORI SUBOR "AIO_Data"
                    On Error Resume Next
                    Set Zosit = Workbooks("AIO_Data")
                    ZositOtvoreny = Not Zosit Is Nothing
                    If ZositOtvoreny = True Then
        
    '                    MsgBox "S˙bor AIO_Data je uû otvoren˝"
                        
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
                    
                    Else
        
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
       
    '                    MsgBox "Otv·ram s˙bor AIO_Data"
'                        Workbooks.Open FileName:="C:\Users\lisy\Desktop\Pl·ny upÌnania\Parametre n·strojov\Parametre n·strojov.xlsm"
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
        
                    End If
        '
                'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    
    '                CisloNastroja = Worksheets("AIO_Plan").Range("S1").Value
                    CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
    '                MsgBox ("HæadanÈ ËÌslo n·stroja: " & CisloNastroja)
                    Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
                    If OblastNajdi Is Nothing Then
                        I = MsgBox("»Ìslo n·stroja sa nenaölo!", vbOKOnly + vbExclamation, "»Ìslo n·stroja")
                    Else
                        OblastNajdi.Select
    '                    MsgBox (CisloNastroja & vbCrLf & _
                                "Adresa hæadanej bunky:  " & OblastNajdi.Address & vbCrLf & _
                                "Riadok hæadanej bunky:  " & OblastNajdi.Row & vbCrLf & _
                                "StÂpec hæadanej bunky:    " & OblastNajdi.Column)
                    
                        OblastNajdiDP = (OblastNajdi.Column + 10)
                        StaryTlakDP = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiDP).Value
                        NovyTlakDP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S12").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiDP).Select
                        MsgBox ("Tlak na doln˝ pridrûiavaË sa prepÌöe z " & StaryTlakDP & " na " & NovyTlakDP & "." & vbCrLf & _
                                "StarÈ tlaky s˙ zapÌsanÈ v koment·ri.")
    '                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSP) = Worksheets("AIO_Plan").Range("S12").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiDP) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S12").Value
                       
    '                    Workbooks("AIO_Data").Close
                    End If
            End Select
    End If

End Sub
'PO VLOZENÌ ALEBO ZMENENÌ UDAJU BRZDA VLOZI ALEBO ZMENÌ UDAJ BRZDA AJ V S˙BORE "AIO_Data"
Sub ZmeniBrzda()

'    If Worksheets("AIO_Plan").PageSetup.LeftHeader = "" Then
'        MsgBox ("ºav· hlaviËka pr·zdnam, neprepisujem brzdu v AIO_Data")
    
    If IpmortVsetkychUdajovZAIO_Data_Running = True Then

'       MsgBox ("Sub IpmortVsetkychUdajovZAIO_Data is running!Neprepisujem ˙daje BRZDA v Paremetre n·strojov")
       
    Else
        I = MsgBox("Prajete si prepÌsaù brzdu v s˙bore 'AIO_Data'  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "PrepÌsaù parameter?")
            
            Select Case I
                Case vbNo
                    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
        '           MsgBox ("Nie")
                Case vbYes
        '
                'OTVORI SUBOR "AIO_Data"
                    On Error Resume Next
                    Set Zosit = Workbooks("AIO_Data")
                    ZositOtvoreny = Not Zosit Is Nothing
                    If ZositOtvoreny = True Then
        
    '                    MsgBox "S˙bor AIO_Data je uû otvoren˝"
                        
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
                    
                    Else
    
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
    
    '                    MsgBox "Otv·ram s˙bor AIO_Data"
'                        Workbooks.Open FileName:="C:\Users\lisy\Desktop\Pl·ny upÌnania\Parametre n·strojov\Parametre n·strojov.xlsm"
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
        
                    End If
        '
                'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    
    '                CisloNastroja = Worksheets("AIO_Plan").Range("S1").Value
                    CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
    '                MsgBox ("HæadanÈ ËÌslo n·stroja: " & CisloNastroja)
                    Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
                    If OblastNajdi Is Nothing Then
                        I = MsgBox("»Ìslo n·stroja sa nenaölo!", vbOKOnly + vbExclamation, "»Ìslo n·stroja")
                    Else
                        OblastNajdi.Select
    '                    MsgBox (CisloNastroja & vbCrLf & _
                                "Adresa hæadanej bunky:  " & OblastNajdi.Address & vbCrLf & _
                                "Riadok hæadanej bunky:  " & OblastNajdi.Row & vbCrLf & _
                                "StÂpec hæadanej bunky:    " & OblastNajdi.Column)
                    
                        OblastNajdiBrzda = (OblastNajdi.Column + 16)
                        StaraBrzda = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiBrzda).Value
                        NovaBrzda = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S9").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiBrzda).Select
                        MsgBox ("Parameter brzda sa prepÌöe z " & StaraBrzda & " na " & NovaBrzda & "." & vbCrLf & _
                                "StarÈ parametre s˙ zapÌsanÈ v koment·ri.")
    '                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSP) = Worksheets("AIO_Plan").Range("S12").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiBrzda) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S9").Value
    '                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, NovaBrzda).Select
    '                    Workbooks("AIO_Data").Close
                    End If
            End Select
    End If

End Sub
'PO VLOZENÌ ALEBO ZMENENÌ UDAJU CAPY NAD STOL VLOZI ALEBO ZMENÌ UDAJ CAPY NAD STOL AJ V S˙BORE "AIO_Data"
Sub Zmeni»apyNadStÙl()

'    If Worksheets("AIO_Plan").PageSetup.LeftHeader = "" Then
'        MsgBox ("ºav· hlaviËka pr·zdnam, neprepisujem Ëapy nad stÙl v AIO_Data")

    If IpmortVsetkychUdajovZAIO_Data_Running = True Then

'       MsgBox ("Sub IpmortVsetkychUdajovZAIO_Data is running!Neprepisujem ˙daje »APY NAD STOL v Paremetre n·strojov")
       
    Else
        I = MsgBox("Prajete si prepÌsaù Ëapy nad stÙl v s˙bore 'AIO_Data'  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "PrepÌsaù parameter?")
            
            Select Case I
                Case vbNo
                    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
        '           MsgBox ("Nie")
                Case vbYes
        '
                'OTVORI SUBOR "AIO_Data"
                    On Error Resume Next
                    Set Zosit = Workbooks("AIO_Data")
                    ZositOtvoreny = Not Zosit Is Nothing
                    If ZositOtvoreny = True Then
        
    '                    MsgBox "S˙bor AIO_Data je uû otvoren˝"
    
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
                    
                    Else
        
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870

        
    '                    MsgBox "Otv·ram s˙bor AIO_Data"
'                        Workbooks.Open FileName:="C:\Users\lisy\Desktop\Pl·ny upÌnania\Parametre n·strojov\Parametre n·strojov.xlsm"
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
        
                    End If
        '
                'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    
    '                CisloNastroja = Worksheets("AIO_Plan").Range("S1").Value
                    CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
    '                MsgBox ("HæadanÈ ËÌslo n·stroja: " & CisloNastroja)
                    Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
                    If OblastNajdi Is Nothing Then
                        I = MsgBox("»Ìslo n·stroja sa nenaölo!", vbOKOnly + vbExclamation, "»Ìslo n·stroja")
                    Else
                        OblastNajdi.Select
    '                    MsgBox (CisloNastroja & vbCrLf & _
                                "Adresa hæadanej bunky:  " & OblastNajdi.Address & vbCrLf & _
                                "Riadok hæadanej bunky:  " & OblastNajdi.Row & vbCrLf & _
                                "StÂpec hæadanej bunky:    " & OblastNajdi.Column)
                    
                        OblastNajdi»apyNadStÙl = (OblastNajdi.Column + 17)
                        Stare»apyNadStÙl = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi»apyNadStÙl).Value
                        Nove»apyNadStÙl = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S11").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi»apyNadStÙl).Select
                        MsgBox ("Parameter Ëapy nad stÙl sa prepÌöe z " & Stare»apyNadStÙl & " na " & Nove»apyNadStÙl & "." & vbCrLf & _
                                "StarÈ parametre s˙ zapÌsanÈ v koment·ri.")
    '                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSP) = Worksheets("AIO_Plan").Range("S12").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi»apyNadStÙl) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S11").Value
    '                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, NovaBrzda).Select
    '                    Workbooks("AIO_Data").Close
                    End If
            End Select
    End If

End Sub
     
     'PO VLOZENÌ ALEBO ZMENENÌ UDAJU VAHA GES VLOZI ALEBO ZMENÌ UDAJ VAHA OT, UT, GES AJ V S˙BORE "AIO_Data"
Sub ZmeniVahuVAIO_Data()

'    If Worksheets("AIO_Plan").PageSetup.LeftHeader = "" Then
'        MsgBox ("ºav· hlaviËka pr·zdnam, neprepisujem v·hy v AIO_Data")

    If IpmortVsetkychUdajovZAIO_Data_Running = True Then

'       MsgBox ("Sub IpmortVsetkychUdajovZAIO_Data is running!Neprepisujem ˙daje V¡HU v Paremetre n·strojov")

    Else
        I = MsgBox("Prajete si prepÌsaù v·hy v s˙bore 'AIO_Data'  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "PrepÌsaù parameter?")
            
            Select Case I
                Case vbNo
                    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
        '           MsgBox ("Nie")
                Case vbYes
        '
                'OTVORI SUBOR "AIO_Data"
                    On Error Resume Next
                    Set Zosit = Workbooks("AIO_Data")
                    ZositOtvoreny = Not Zosit Is Nothing
                    If ZositOtvoreny = True Then
        
    '                    MsgBox "S˙bor AIO_Data je uû otvoren˝"
                        
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
                    
                    Else
        
                    'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 1 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '668 '686 '(976)
'''''                        Application.Height = 870
        
    '                    MsgBox "Otv·ram s˙bor AIO_Data"
'                        Workbooks.Open FileName:="C:\Users\lisy\Desktop\Pl·ny upÌnania\Parametre n·strojov\Parametre n·strojov.xlsm"
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                        Application.WindowState = xlNormal
'''''                        Application.Left = 602 '226
'''''                        Application.Top = 1
'''''                        Application.Width = 601 '754
'''''                        Application.Height = 870
                        
                        Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'ZbalÌ zoskupenÈ udaje
        
                    End If
        '
                'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
                    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
                    
    '                CisloNastroja = Worksheets("AIO_Plan").Range("S1").Value
                    CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
    '                MsgBox ("HæadanÈ ËÌslo n·stroja: " & CisloNastroja)
                    Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
                    If OblastNajdi Is Nothing Then
                        I = MsgBox("»Ìslo n·stroja sa nenaölo!", vbOKOnly + vbExclamation, "»Ìslo n·stroja")
                    Else
                        OblastNajdi.Select
    '                    MsgBox (CisloNastroja & vbCrLf & _
                                "Adresa hæadanej bunky:  " & OblastNajdi.Address & vbCrLf & _
                                "Riadok hæadanej bunky:  " & OblastNajdi.Row & vbCrLf & _
                                "StÂpec hæadanej bunky:    " & OblastNajdi.Column)
                    
                        OblastNajdiVahaOT = (OblastNajdi.Column + 559)
                        OblastNajdiVahaUT = (OblastNajdi.Column + 560)
                        OblastNajdiVahaGES = (OblastNajdi.Column + 561)
    
                        StaraVahaOT = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVahaOT).Value
                        StaraVahaUT = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVahaUT).Value
                        StaraVahaGES = Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVahaGES).Value
    
                        NovaVahaOT = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("G5").Value
                        NovaVahaUT = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("G6").Value
                        NovaVahaGES = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("G7").Value
    
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVahaGES).Select
                       
    '                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSP) = Worksheets("AIO_Plan").Range("S12").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVahaOT) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("G5").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVahaUT) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("G6").Value
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVahaGES) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("G7").Value
    
    '                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, NovaBrzda).Select
    '                    Workbooks("AIO_Data").Close
                    End If
            End Select
    End If

End Sub
Sub DoplniOstatneUdajeDoAIO_Data()
'DOPLNI OSTATNE UDAJE DO SUBORU "AIO_Data"

                'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
    
                CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
                Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
                If OblastNajdi Is Nothing Then
                    I = MsgBox("»Ìslo n·stroja sa nenaölo!", vbOKOnly + vbExclamation, "»Ìslo n·stroja")
                Else
''16 Brzda
'                    OblastNajdiBrzda = (OblastNajdi.Column + 16)
'                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiBrzda) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S9").Value
''17 »apyNadStÙl
'                    OblastNajdi»apyNadStÙl = (OblastNajdi.Column + 17)
'                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi»apyNadStÙl) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S11").Value
'18 DÂûkaN·stroja
                    OblastNajdiDÂûkaN·stroja = (OblastNajdi.Column + 18)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiDÂûkaN·stroja) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("W5").Value
'19 äÌrkaN·stroja
                    OblastNajdiäÌrkaN·stroja = (OblastNajdi.Column + 19)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiäÌrkaN·stroja) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AD5").Value
'20 V˝ökaN·stroja
                    OblastNajdiV˝ökaN·stroja = (OblastNajdi.Column + 20)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiV˝ökaN·stroja) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AK5").Value
'21 VzdialenosùMedziDr·ûkamiOT
                    OblastNajdiVzdialenosùMedziDr·ûkamiOT = (OblastNajdi.Column + 21)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiVzdialenosùMedziDr·ûkamiOT) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AI5").Value
'22 OdstavovaciePrvkyGDF
                    OblastNajdiOdstavovaciePrvkyGDF = (OblastNajdi.Column + 22)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiOdstavovaciePrvkyGDF) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("R5").Value
'23 OdstavovaciePrvkyOB
                    OblastNajdiOdstavovaciePrvkyOB = (OblastNajdi.Column + 23)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiOdstavovaciePrvkyOB) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S5").Value
'24 OdstavovaciePrvkyZdvih_Vyöka
'                    OblastNajdi.Select
                    OblastNajdiOdstavovaciePrvkyZdvih_Vyöka = (OblastNajdi.Column + 24)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiOdstavovaciePrvkyZdvih_Vyöka) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("W6").Value
'25 UpÌnaciaV˝ökaN·stroja
                    OblastNajdiUpncVökNstrj = (OblastNajdi.Column + 25)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiUpncVökNstrj) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AK6").Value
'26 PridrûiavaËBaranBez
                    OblastNajdiPridrûiavaËBaranBez = (OblastNajdi.Column + 26)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridrûiavaËBaranBez) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("J13").Value
'27 PridrûiavaËBaran»apy
                    OblastNajdiPridrûiavaËBaran»apy = (OblastNajdi.Column + 27)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridrûiavaËBaran»apy) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("N13").Value
'28 PridrûiavaËBaranGDF
                    OblastNajdiPridrûiavaËBaranGDF = (OblastNajdi.Column + 28)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridrûiavaËBaranGDF) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("R13").Value
'29 PridrûiavaËStÙlBez
                    OblastNajdiPridrûiavaËStÙlBez = (OblastNajdi.Column + 29)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridrûiavaËStÙlBez) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("J12").Value
'30 PridrûiavaËStÙl»apy
                    OblastNajdiPridrûiavaËStÙl»apy = (OblastNajdi.Column + 30)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridrûiavaËStÙl»apy) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("N12").Value
'31 OblastNajdiPridrûiavaËStÙlGDF
                    OblastNajdiPridrûiavaËStÙlGDF = (OblastNajdi.Column + 31)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPridrûiavaËStÙlGDF) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("R12").Value
'32 OblastNajdiMûnsùUpntNstrjDLs1 (Moûnosù upnutia n·stroja do lisu1)
                    OblastNajdiMûnsùUpntNstrjDLs1 = (OblastNajdi.Column + 32)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiMûnsùUpntNstrjDLs1) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("Y7").Value
'33 OblastNajdiMûnsùUpntNstrjDLs2 (Moûnosù upnutia n·stroja do lisu2)
                    OblastNajdiMûnsùUpntNstrjDLs2 = (OblastNajdi.Column + 33)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiMûnsùUpntNstrjDLs2) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AA7").Value
'34 OblastNajdiMûnsùUpntNstrjDLs3 (Moûnosù upnutia n·stroja do lisu3)
                    OblastNajdiMûnsùUpntNstrjDLs3 = (OblastNajdi.Column + 34)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiMûnsùUpntNstrjDLs3) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AC7").Value
'35 OblastNajdiMûnsùUpntNstrjDLs4 (Moûnosù upnutia n·stroja do lisu4)
                    OblastNajdiMûnsùUpntNstrjDLs4 = (OblastNajdi.Column + 35)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiMûnsùUpntNstrjDLs4) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AE7").Value
'36 OblastNajdiCntrvnNstrjPrmrLH (Centrovanie n·stroja priemer æav˝ horn˝)
                    OblastNajdiCntrvnNstrjPrmrLH = (OblastNajdi.Column + 36)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiCntrvnNstrjPrmrLH) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("P28").Value
'37 OblastNajdiCntrvnNstrjPrmrPH (Centrovanie n·stroja priemer prav˝ horn˝)
                    OblastNajdiCntrvnNstrjPrmrPH = (OblastNajdi.Column + 37)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiCntrvnNstrjPrmrPH) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("Z28").Value
'38 OblastNajdiCntrvnNstrjPrmrLD (Centrovanie n·stroja priemer æav˝ doln˝)
                    OblastNajdiCntrvnNstrjPrmrLD = (OblastNajdi.Column + 38)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiCntrvnNstrjPrmrLD) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("P31").Value
'39 OblastNajdiCntrvnNstrjPrmrPD (Centrovanie n·stroja priemer prav˝ doln˝)
                    OblastNajdiCntrvnNstrjPrmrPD = (OblastNajdi.Column + 39)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiCntrvnNstrjPrmrPD) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("Z31").Value
'40 OblastNajdiSrdncCntrvnLHR (S˙radnice centrovania æav· horn· ötvrtina riadok)
                    OblastNajdiSrdncCntrvnLHR = (OblastNajdi.Column + 40)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnLHR) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("T28").Value
'41 OblastNajdiSrdncCntrvnLHS (S˙radnice centrovania æav· horn· ötvrtina stÂpec)
                    OblastNajdiSrdncCntrvnLHS = (OblastNajdi.Column + 41)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnLHS) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S29").Value
'42 OblastNajdiSrdncCntrvnPHR
                    OblastNajdiSrdncCntrvnPHR = (OblastNajdi.Column + 42)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnPHR) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("V28").Value
'43 OblastNajdiSrdncCntrvnPHS
                    OblastNajdiSrdncCntrvnPHS = (OblastNajdi.Column + 43)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnPHS) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("W29").Value
'44 OblastNajdiSrdncCntrvnLDR
                    OblastNajdiSrdncCntrvnLDR = (OblastNajdi.Column + 44)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnLDR) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("T31").Value
'45 OblastNajdiSrdncCntrvnLDS
                    OblastNajdiSrdncCntrvnLDS = (OblastNajdi.Column + 45)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnLDS) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S30").Value
'46 OblastNajdiSrdncCntrvnPDR
                    OblastNajdiSrdncCntrvnPDR = (OblastNajdi.Column + 46)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnPDR) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("V31").Value
'47 OblastNajdiSrdncCntrvnPDS
                    OblastNajdiSrdncCntrvnPDS = (OblastNajdi.Column + 47)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSrdncCntrvnPDS) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("W30").Value
'48 OblastNajdiSmrLsvnL (Smer lisovania æav·)
                    OblastNajdiSmrLsvnL = (OblastNajdi.Column + 48)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSmrLsvnL) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("C41").Value
'49 OblastNajdiSmrLsvnH (Smer lisovania hore)
                    OblastNajdiSmrLsvnH = (OblastNajdi.Column + 49)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSmrLsvnH) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("U32").Value
'50 OblastNajdiSmrLsvnD (Smer lisovania dole)
                    OblastNajdiSmrLsvnD = (OblastNajdi.Column + 50)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSmrLsvnD) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("U50").Value
'51 OblastNajdiSmrLsvnP (Smer lisovania prav·)
                    OblastNajdiSmrLsvnP = (OblastNajdi.Column + 51)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiSmrLsvnP) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AM41").Value
'52 OblastNajdiPznmkRdk1 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk1 = (OblastNajdi.Column + 52)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("I14").Value
                     
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("I14").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("I14").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("I14").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("I14").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("I14").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk1")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk1")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk1).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------
'53 OblastNajdiPznmkRdk2 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk2 = (OblastNajdi.Column + 53)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B15").Value
'
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B15").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B15").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B15").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B15").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B15").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk2")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk2")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk2).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------

'
'54 OblastNajdiPznmkRdk3 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk3 = (OblastNajdi.Column + 54)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B16").Value
'
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B16").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B16").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B16").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B16").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B16").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk3")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk3")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk3).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------
'
'55 OblastNajdiPznmkRdk4 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk4 = (OblastNajdi.Column + 55)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B17").Value
'
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B17").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B17").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B17").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B17").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B17").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk4")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk4")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk4).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------
'56 OblastNajdiPznmkRdk5 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk5 = (OblastNajdi.Column + 56)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B18").Value
'
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B18").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B18").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B18").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B18").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B18").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk5")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk5")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk5).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------
'
'57 OblastNajdiPznmkRdk6 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk6 = (OblastNajdi.Column + 57)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B19").Value
'
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B19").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B19").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B19").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B19").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B19").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk6")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk6")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk6).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------

'58 OblastNajdiPznmkRdk7 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk7 = (OblastNajdi.Column + 58)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B20").Value
                     
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B20").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B20").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B20").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B20").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B20").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk7")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk7")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk7).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------


'59 OblastNajdiPznmkRdk8 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk8 = (OblastNajdi.Column + 59)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B21").Value
                     
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B21").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B21").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B21").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B21").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B21").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk8")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk8")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk8).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------

'60 OblastNajdiPznmkRdk9 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk9 = (OblastNajdi.Column + 60)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B22").Value
                     
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B22").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B22").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B22").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B22").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B22").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk9")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk9")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk9).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------

'61 OblastNajdiPznmkRdk10 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk10 = (OblastNajdi.Column + 61)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B23").Value
                     
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B23").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B23").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B23").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B23").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B23").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk10")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk10")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk10).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------

'62 OblastNajdiPznmkRdk11 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk11 = (OblastNajdi.Column + 62)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B24").Value
'
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B24").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B24").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B24").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B24").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B24").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk11")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk11")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk11).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------

'63 OblastNajdiPznmkRdk12 (Pozn·mky k n·stroju )
                    OblastNajdiPznmkRdk12 = (OblastNajdi.Column + 63)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B25").Value
                     
'                     '-------------
                     'Sub EXPORTSkopirujeFarbuVyplnePismoZarovnanieZBunkyVPozn·mkach

                    IC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B25").Interior.Color
                    FC = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B25").Font.Color
                    HA = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B25").HorizontalAlignment
                    IP = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B25").Interior.Pattern

                    If Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("B25").Interior.Pattern <> 4000 Then
'                       MsgBox ("BeûÌ If PznmkRdk12")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12).Interior.Color = IC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12).HorizontalAlignment = HA
                    Else:
'                        MsgBox ("BeûÌ Else PznmkRdk12")
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12).Font.Color = FC
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12).HorizontalAlignment = HA
                        Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPznmkRdk12).Select
                            With Selection.Interior
                            .Pattern = xlPatternLinearGradient
                            .Gradient.ColorStops.Clear
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(0)
                                .Color = 65535
                            End With
                            With Selection.Interior.Gradient.ColorStops.Add(1)
                                .Color = 255
                            End With
                    End If

'                        FUNGUJE
'                    '------------
'
'562 OblastNajdiPctTlËnch»pv (PoËet tlaËn˝ch Ëapov )
                    OblastNajdiPctTlËnch»pv = (OblastNajdi.Column + 562)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPctTlËnch»pv) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AN28").Value
'563 OblastNajdiGdfAleboBloky (Text vbunke "L6" Zdvih GDF/Vyöka odstavovacÌch blokov )
                    OblastNajdiGdfAleboBloky = (OblastNajdi.Column + 563)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiGdfAleboBloky) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("L6").Value
'564 OblastNajdiLavaHlavicka (Lava Hlavicka "Datum vytvorenia")
                    OblastNajdiLavaHlavicka = (OblastNajdi.Column + 564)
                    LavaHlavicka = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").PageSetup.LeftHeader
                    DatumVytvorenia = Mid(LavaHlavicka, 22, 55)
'                    MsgBox (LavaHlavicka)
'                    MsgBox (DatumVytvorenia)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiLavaHlavicka) = DatumVytvorenia 'Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").PageSetup.LeftHeader
                   
'                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiDÂûkaN·stroja).Select
'                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiLavaHlavicka).Select

'565 OblastNajdiPravaHlavicka (Prava Hlavicka "Datum poslednej aktualiz·cie")
                    OblastNajdiPravaHlavicka = (OblastNajdi.Column + 565)
                    PravaHlavicka = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").PageSetup.RightHeader
                    DatumPoslednejAktualizacie = Mid(PravaHlavicka, 22, 67)
'                    MsgBox (PravaHlavicka)
'                    MsgBox (DatumPoslednejAktualizacie)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPravaHlavicka) = DatumPoslednejAktualizacie 'Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").PageSetup.RightHeader

'                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPravaHlavicka).Select

'566 OblastNajdiPoËetCervenychCentrovacichCapov (PoËet Ëerven˝ch centrovacÌch Ëapov)
                    OblastNajdiPoËetCervenychCentrovacichCapov = (OblastNajdi.Column + 566)
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiPoËetCervenychCentrovacichCapov) = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("AN29").Value

'-1 OblastNajdiDatumPoslednejAktualiz·cie (Prava Hlavicka "Datum poslednej aktualiz·cie")
                    OblastNajdiDatumPoslednejAktualiz·cie = (OblastNajdi.Column - 1)
                    PravaHlavickaCela = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").PageSetup.RightHeader
                    PravaHlavickaIbaDatum = Mid(PravaHlavickaCela, 52, 67)
'                    MsgBox (PravaHlavickaCela)
'                    MsgBox (PravaHlavickaIbaDatum)
                    VerziaPU = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("A64").Value

'                    MsgBox "EXPORT_Verzia pl·nu upÌnania: " & VerziaPU
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiDatumPoslednejAktualiz·cie) = PravaHlavickaIbaDatum & "  " & VerziaPU 'Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").PageSetup.RightHeader
                    
                End If
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column + 565).Select 'Vyberie bunku Prav· hlaviËka

End Sub

Sub DoplniRasterStolaDoAIO_Data()
'DOPLNI RASTER STOLA DO SUBORU "AIO_Data"

                'V OTVORENO SUBORE NAJDE CISLO NASTROJA A VYBERIE HO
                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
    
                CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
                Set OblastNajdi = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Columns(7).Find(CisloNastroja, LookIn:=xlValues, SearchFormat:=False)
                If OblastNajdi Is Nothing Then
                    I = MsgBox("»Ìslo n·stroja sa nenaölo!", vbOKOnly + vbExclamation, "»Ìslo n·stroja")
                Else
                    
'64 OblastNajdiRaster8H (Raster stola riadok 8 hore )
                    Dim riadok8H As Range
                    Set riadok8H = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$34:$AK$34")
                    
                    riadok8H.Copy
                    
                    OblastNajdiRaster8HoreZaËiatok = (OblastNajdi.Column + 64)
'                    OblastNajdiRaster8HoreKoniec = (OblastNajdi.Column + 96)
'                    MsgBox (Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster8HoreZaËiatok).Address) 'stlpec RS-CY pre jeden riadok rastra &
'                    MsgBox (Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster8HoreKoniec).Address)
'                    MsgBox (Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster8HoreZaËiatok).Address _
'                    & ":" & Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster8HoreKoniec).Address)
                    
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster8HoreZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'97 OblastNajdiRaster7H (Raster stola riadok 7 hore )
                    Dim Riadok7H As Range
                    Set Riadok7H = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$35:$AK$35")
                    
                    Riadok7H.Copy
                    
                    OblastNajdiRaster7HoreZaËiatok = (OblastNajdi.Column + 97)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster7HoreZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'130 OblastNajdiRaster6H (Raster stola riadok 6 hore )
                    Dim Riadok6H As Range
                    Set Riadok6H = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$36:$AK$36")
                    
                    Riadok6H.Copy
                    
                    OblastNajdiRaster6HoreZaËiatok = (OblastNajdi.Column + 130)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster6HoreZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'163 OblastNajdiRaster5H (Raster stola riadok 5 hore )
                    Dim Riadok5H As Range
                    Set Riadok5H = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$37:$AK$37")
                    
                    Riadok5H.Copy
                    
                    OblastNajdiRaster5HoreZaËiatok = (OblastNajdi.Column + 163)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster5HoreZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'196 OblastNajdiRaster4H (Raster stola riadok 4 hore )
                    Dim Riadok4H As Range
                    Set Riadok4H = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$38:$AK$38")
                    
                    Riadok4H.Copy
                    
                    OblastNajdiRaster4HoreZaËiatok = (OblastNajdi.Column + 196)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster4HoreZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'229 OblastNajdiRaster3H (Raster stola riadok 3 hore )
                    Dim Riadok3H As Range
                    Set Riadok3H = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$39:$AK$39")
                    
                    Riadok3H.Copy
                    
                    OblastNajdiRaster3HoreZaËiatok = (OblastNajdi.Column + 229)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster3HoreZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'262 OblastNajdiRaster2H (Raster stola riadok 2 hore )
                    Dim Riadok2H As Range
                    Set Riadok2H = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$40:$AK$40")
                    
                    Riadok2H.Copy
                    
                    OblastNajdiRaster2HoreZaËiatok = (OblastNajdi.Column + 262)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster2HoreZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'295 OblastNajdiRaster1S (Raster stola riadok 1 Stred )
                    Dim Riadok1S As Range
                    Set Riadok1S = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$41:$AK$41")
                    
                    Riadok1S.Copy
                    
                    OblastNajdiRaster1StredZaËiatok = (OblastNajdi.Column + 295)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster1StredZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'328 OblastNajdiRaster2D (Raster stola riadok 2 dole )
                    Dim Riadok2D As Range
                    Set Riadok2D = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$42:$AK$42")
                    
                    Riadok2D.Copy
                    
                    OblastNajdiRaster2DoleZaËiatok = (OblastNajdi.Column + 328)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster2DoleZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'361 OblastNajdiRaster3D (Raster stola riadok 3 dole )
                    Dim Riadok3D As Range
                    Set Riadok3D = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$43:$AK$43")
                    
                    Riadok3D.Copy
                    
                    OblastNajdiRaster3DoleZaËiatok = (OblastNajdi.Column + 361)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster3DoleZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'394 OblastNajdiRaster4D (Raster stola riadok 4 dole )
                    Dim Riadok4D As Range
                    Set Riadok4D = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$44:$AK$44")
                    
                    Riadok4D.Copy
                    
                    OblastNajdiRaster4DoleZaËiatok = (OblastNajdi.Column + 394)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster4DoleZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'427 OblastNajdiRaster5D (Raster stola riadok 5 dole )
                    Dim Riadok5D As Range
                    Set Riadok5D = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$45:$AK$45")
                    
                    Riadok5D.Copy
                    
                    OblastNajdiRaster5DoleZaËiatok = (OblastNajdi.Column + 427)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster5DoleZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'460 OblastNajdiRaster6D (Raster stola riadok 6 dole )
                    Dim Riadok6D As Range
                    Set Riadok6D = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$46:$AK$46")
                    
                    Riadok6D.Copy
                    
                    OblastNajdiRaster6DoleZaËiatok = (OblastNajdi.Column + 460)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster6DoleZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'493 OblastNajdiRaster7D (Raster stola riadok 7 dole )
                    Dim Riadok7D As Range
                    Set Riadok7D = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$47:$AK$47")
                    
                    Riadok7D.Copy
                    
                    OblastNajdiRaster7DoleZaËiatok = (OblastNajdi.Column + 493)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster7DoleZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
'526 OblastNajdiRaster8D (Raster stola riadok 8 dole )
                    Dim Riadok8D As Range
                    Set Riadok8D = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("$E$48:$AK$48")
                    
                    Riadok8D.Copy
                    
                    OblastNajdiRaster8DoleZaËiatok = (OblastNajdi.Column + 526)
'
                    Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdiRaster8DoleZaËiatok).Select
                    Worksheets("AIO_Data").Paste
                    Application.CutCopyMode = False
                    
                End If
                Worksheets("AIO_Data").Cells(OblastNajdi.Row, OblastNajdi.Column + 565).Select 'Vyberie bunku Prav· hlaviËka

End Sub

Sub OtvoriParamNastrADoplniOstatneUdajeDoAIO_Data()
        'V SUBORE "AIO_Data" Doplni vsetky ostatne ˙daje
        '---------------------------------------------------------------------------------------
            If Worksheets("AIO_Plan").PageSetup.RightHeader = "" Then
'                MsgBox ("prava hlaviËka pr·zdna nerobÌm niË")
            Else
'                MsgBox ("otvaram AIO_Data a spustam makro DoplniOstatneUdajeDoAIO_Data")
                I = MsgBox("Prajete si aktualizovaù vöetky ˙daje v s˙bore 'AIO_Data'  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "VyznaËiù moûnosù upnutia?")
                    Select Case I
                        Case vbNo
                            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
                '           MsgBox ("Nie")
                        Case vbYes

                            'Fcesta = "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…"
                            'FcestaJPG = "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…\PU_JPG"
                            'FcestaPDF = "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…\PU_PDF"
                            FCisloNastroja = Sheets("AIO_Plan").Range("S1").Text
                            FOperacia = Sheets("AIO_Plan").Range("AM1").Text
                            FKrok = Sheets("AIO_Plan").Range("AM3").Text
                            FCisloDielu = Sheets("AIO_Plan").Range("S3").Text

                            If Range("AM3").Text = "" Then
                                NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_Pl·n upÌnania"
                            Else
                                NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_S" & FKrok & "_Pl·n upÌnania"
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

                                MsgBox "S˙bor AIO_Data je uû otvoren˝"

                        'ZmensÌ okno excelu na lav˙ polovicu obrazovky
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
                        'Skryje z·hlavia
                '                ActiveWindow.DisplayHeadings = False

                            'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                                Application.WindowState = xlNormal
'''''                                Application.Left = 602 '226
'''''                                Application.Top = 1
'''''                                Application.Width = 601 '754
'''''                                Application.Height = 870
                            Else

                                'UserForm1.Unload
                                'Application.WindowState = xlMaximized

                                MsgBox "Otv·ram s˙bor AIO_Data"

                            'ZmensÌ okno excelu na lav˙ polovicu obrazovky
                                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
'''''                                Application.WindowState = xlNormal
'''''                                Application.Left = 1 '226
'''''                                Application.Top = 1
'''''                                Application.Width = 601 '668 '686 '(976)
'''''                                Application.Height = 870


'                                Workbooks.Open FileName:="C:\Users\lisy\Desktop\Pl·ny upÌnania\Parametre n·strojov\Parametre n·strojov.xlsm"
                                Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate

                        'Skryje riadok vzorcov
                '                Application.DisplayFormulaBar = False
                                '
                        'Skryje z·hlavia
                '                ActiveWindow.DisplayHeadings = False

                            'ZmensÌ okno excelu na prav˙ polovicu obrazovky
'''''                                Application.WindowState = xlNormal
'''''                                Application.Left = 602 '226
'''''                                Application.Top = 1
'''''                                Application.Width = 601 '754
'''''                                Application.Height = 870
                            End If
                            'OK-------------------------------
'
                                Call DoplniOstatneUdajeDoAIO_Data
                                
                                If Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("AN28").Value > 0 Or _
                                    Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Range("AN29").Value > 0 Then
                    '                MsgBox ("Sp˙ötam DoplniRasterStolaDoAIO_Data")
                                    Call DoplniRasterStolaDoAIO_Data
                    '            Else
                    '                MsgBox ("Nep˙ötam DoplniRasterStolaDoAIO_Data")
                                End If

                    End Select
                    
            End If
End Sub



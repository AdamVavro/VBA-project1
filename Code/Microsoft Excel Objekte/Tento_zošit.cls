VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tento_zošit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'PRED ULOZENIM ZABEZPECI HAROK
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'Zbalí zoskupenia
   
    ActiveWindow.ScrollColumn = 8 'posunie do¾ava

'    ActiveSheet.Protect Password:="Lis.0123"
    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"

End Sub
'PO OTVORENI NASTAVI VELKOST OKNA
Private Sub Workbook_Open()

    'Povolí zobrezeniu systémových hlášok
        Application.DisplayAlerts = True

    'Call Maxim
    Application.WindowState = xlNormal
        Application.Left = 226
        Application.Top = 1
        Application.Width = 668 '686 '(976)
        Application.Height = 870
    'Skryje riadok vzorcov
        Application.DisplayFormulaBar = False
    '
    'Skryje záhlavia
        ActiveWindow.DisplayHeadings = False
    'Skryje mriežku
        ActiveWindow.DisplayGridlines = False
    'Ak je pás s nástrojmi pripnutý, tak ho zbalí
        If Application.CommandBars("Ribbon").Height > 100 Then 'ActiveWindow.ToggleRibbon
'           MsgBox ("Pás je pripnutý")
        Application.CommandBars.ExecuteMso "MinimizeRibbon"
        End If
        
    'Skryje riadok vzorcov
    Application.DisplayFormulaBar = False


        
''''''''''''''''''''''''


'Po otvorení súboru sa spustí èasovaè, _
po uplynutí èasu vyskoèí hláška
'    Workbooks.Open Filename:="C:\Users\lisy\Desktop\Plány upínania\Parametre nástrojov\Countdown_KT05.xlsm"
'    Workbooks("Countdown_KT05.xlsm").Worksheets("Hárok1").Activate
   
'''''''''''''
'    Call NastavenieCasovaca
'''''''''''''
'    gCount = Now + TimeValue("00:10:00")
'    Application.OnTime gCount, "Hlaska"

   
'''''''''''''

'''''''''''''''''''''''''
'
''AUTOMATICKÝ ÈASOVAÈ FUNGUJE
''Funguje po spusteni zaène odpoèítava, po odpoèítani oznami _
'koniec a spýta sa ci chem reštartova odpoèítavanie
'    Range("BR28").Value = "0:10:00"
'    Call Timer
'
'''''''''''''''''''''''''

End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    'Zabráni zobrezeniu systémových hlášok
        Application.DisplayAlerts = False
    'Odokryje riadok vzorcov
        Application.DisplayFormulaBar = True
        '
    'Odokryje kryje záhlavia
        ActiveWindow.DisplayHeadings = True
    'Ak je pás s nástrojmi zbalený, tak ho pripne
        If Application.CommandBars("Ribbon").Height > 100 Then 'ActiveWindow.ToggleRibbon
'           MsgBox ("Pás je pripnutý")
        Else: Application.CommandBars.ExecuteMso "MinimizeRibbon"
        End If

End Sub

''''''''
''''''''
'Private Sub Workbook_AfterSave(ByVal Success As Boolean)
'''Po uložení automaticky zalohuje na dve miesta: _
''"C:\Users\lisy\Desktop\Plány upínania tablet\0_Databáza plánov upínania_Tablet.xlsm" _
''"C:\Users\lisy\Desktop\Plány upínania tablet\0_Databáza plánov upínania_Tablet.xlsm"
''    If ThisWorkbook.Name = "AIO_Data" Then
''            MsgBox ("Ukladam na tablet")
''
''            'Zabráni zobrezeniu systémových hlášok
''             Application.DisplayAlerts = False
''
''            Application.ScreenUpdating = False   'vypne prekreslovanie obrazovky, tým sa makro zrýchli
''
''            Workbooks("AIO_Data").SaveCopyAs "T:\430_F\10_Verejne\10_planovanieLisov\Plány upínania nástrojov\Databáza plánov upínania.xlsm"
''
''            'Zabezpeèí a uloží zošit "Databáza plánov upínania.xlsm"
''            Application.Workbooks.Open ("T:\430_F\10_Verejne\10_planovanieLisov\Plány upínania nástrojov\Databáza plánov upínania.xlsm")
''            Workbooks("Databáza plánov upínania.xlsm").Worksheets("AIO_Data").Protect Password:="Lis.0123"
''            Workbooks("Databáza plánov upínania.xlsm").Save
''            Workbooks("Databáza plánov upínania.xlsm").Close
''
''            Workbooks("AIO_Data").SaveCopyAs "C:\Users\lisy\Desktop\Plány upínania tablet\0_Databáza plánov upínania_Tablet.xlsm"
''
''            'Zabezpeèí a uloží zošit "0_Databáza plánov upínania_Tablet.xlsm"
''            Application.Workbooks.Open ("C:\Users\lisy\Desktop\Plány upínania tablet\0_Databáza plánov upínania_Tablet.xlsm")
''            Workbooks("0_Databáza plánov upínania_Tablet.xlsm").Worksheets("AIO_Data").Protect Password:="Lis.0123"
''            Workbooks("0_Databáza plánov upínania_Tablet.xlsm").Save
''            Workbooks("0_Databáza plánov upínania_Tablet.xlsm").Close
''
''            Application.ScreenUpdating = True   'zapne prekreslovanie obrazovky, tým sa makro zrýchli
''
''        Else: MsgBox ("Neukladam na tablet")
''             'Zabráni zobrezeniu systémových hlášok
''            Application.DisplayAlerts = False
''
'''            Workbooks("0_Databáza plánov upínania_Tablet.xlsm").Activate
'''            Workbooks("0_Databáza plánov upínania_Tablet.xlsm").Close
''        End If
'
'End Sub
'
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''Vytvorenie Zálohy na disk "T" a na tablet'''''''''''''''''''''''''''''''''''''
''        If ThisWorkbook.Name = "AIO_Data" Then
'''            MsgBox ("Ukladam na tablet")
''
''            'Zabráni zobrezeniu systémových hlášok
''             Application.DisplayAlerts = False
''
''             'Skryje riadok vzorcov
''             Application.DisplayFormulaBar = False
''    '
''            'Skryje záhlavia
''                ActiveWindow.DisplayHeadings = False
''            'Skryje mriežku
''                ActiveWindow.DisplayGridlines = False
''
''            'Uloží a zabezpeèí zošit "Databáza plánov upínania.xlsm"
''            Workbooks("AIO_Data").SaveAs "T:\430_F\10_Verejne\10_planovanieLisov\Plány upínania nástrojov\Databáza plánov upínania.xlsm"
''            Workbooks("Databáza plánov upínania.xlsm").Worksheets("AIO_Data").Protect Password:="Lis.0123"
''            Workbooks("Databáza plánov upínania.xlsm").Save
''
''            'Uloží a zabezpeèí zošit "0_Databáza plánov upínania_Tablet.xlsm"
''            Workbooks("Databáza plánov upínania.xlsm").SaveCopyAs "C:\Users\lisy\Desktop\Plány upínania tablet\0_Databáza plánov upínania_Tablet.xlsm"
''
''        Else:
''            'MsgBox ("Neukladam na tablet")
''             'Zabráni zobrezeniu systémových hlášok
''            Application.DisplayAlerts = False
''
'''            Workbooks("0_Databáza plánov upínania_Tablet.xlsm").Activate
'''            Workbooks("0_Databáza plánov upínania_Tablet.xlsm").Close
''        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'End Sub




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tento_zo�it"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'PRED ULOZENIM ZABEZPECI HAROK
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Activate
    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Data").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1 'Zbal� zoskupenia
   
    ActiveWindow.ScrollColumn = 8 'posunie do�ava

'    ActiveSheet.Protect Password:="Lis.0123"
    Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Activate
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"

End Sub
'PO OTVORENI NASTAVI VELKOST OKNA
Private Sub Workbook_Open()

    'Povol� zobrezeniu syst�mov�ch hl�ok
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
    'Skryje z�hlavia
        ActiveWindow.DisplayHeadings = False
    'Skryje mrie�ku
        ActiveWindow.DisplayGridlines = False
    'Ak je p�s s n�strojmi pripnut�, tak ho zbal�
        If Application.CommandBars("Ribbon").Height > 100 Then 'ActiveWindow.ToggleRibbon
'           MsgBox ("P�s je pripnut�")
        Application.CommandBars.ExecuteMso "MinimizeRibbon"
        End If
        
    'Skryje riadok vzorcov
    Application.DisplayFormulaBar = False


        
''''''''''''''''''''''''


'Po otvoren� s�boru sa spust� �asova�, _
po uplynut� �asu vysko�� hl�ka
'    Workbooks.Open Filename:="C:\Users\lisy\Desktop\Pl�ny up�nania\Parametre n�strojov\Countdown_KT05.xlsm"
'    Workbooks("Countdown_KT05.xlsm").Worksheets("H�rok1").Activate
   
'''''''''''''
'    Call NastavenieCasovaca
'''''''''''''
'    gCount = Now + TimeValue("00:10:00")
'    Application.OnTime gCount, "Hlaska"

   
'''''''''''''

'''''''''''''''''''''''''
'
''AUTOMATICK� �ASOVA� FUNGUJE
''Funguje po spusteni za�ne odpo��tava�, po odpo��tani oznami _
'koniec a sp�ta sa ci chem re�tartova� odpo��tavanie
'    Range("BR28").Value = "0:10:00"
'    Call Timer
'
'''''''''''''''''''''''''

End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    'Zabr�ni zobrezeniu syst�mov�ch hl�ok
        Application.DisplayAlerts = False
    'Odokryje riadok vzorcov
        Application.DisplayFormulaBar = True
        '
    'Odokryje kryje z�hlavia
        ActiveWindow.DisplayHeadings = True
    'Ak je p�s s n�strojmi zbalen�, tak ho pripne
        If Application.CommandBars("Ribbon").Height > 100 Then 'ActiveWindow.ToggleRibbon
'           MsgBox ("P�s je pripnut�")
        Else: Application.CommandBars.ExecuteMso "MinimizeRibbon"
        End If

End Sub

''''''''
''''''''
'Private Sub Workbook_AfterSave(ByVal Success As Boolean)
'''Po ulo�en� automaticky zalohuje na dve miesta: _
''"C:\Users\lisy\Desktop\Pl�ny up�nania tablet\0_Datab�za pl�nov up�nania_Tablet.xlsm" _
''"C:\Users\lisy\Desktop\Pl�ny up�nania tablet\0_Datab�za pl�nov up�nania_Tablet.xlsm"
''    If ThisWorkbook.Name = "AIO_Data" Then
''            MsgBox ("Ukladam na tablet")
''
''            'Zabr�ni zobrezeniu syst�mov�ch hl�ok
''             Application.DisplayAlerts = False
''
''            Application.ScreenUpdating = False   'vypne prekreslovanie obrazovky, t�m sa makro zr�chli
''
''            Workbooks("AIO_Data").SaveCopyAs "T:\430_F\10_Verejne\10_planovanieLisov\Pl�ny up�nania n�strojov\Datab�za pl�nov up�nania.xlsm"
''
''            'Zabezpe�� a ulo�� zo�it "Datab�za pl�nov up�nania.xlsm"
''            Application.Workbooks.Open ("T:\430_F\10_Verejne\10_planovanieLisov\Pl�ny up�nania n�strojov\Datab�za pl�nov up�nania.xlsm")
''            Workbooks("Datab�za pl�nov up�nania.xlsm").Worksheets("AIO_Data").Protect Password:="Lis.0123"
''            Workbooks("Datab�za pl�nov up�nania.xlsm").Save
''            Workbooks("Datab�za pl�nov up�nania.xlsm").Close
''
''            Workbooks("AIO_Data").SaveCopyAs "C:\Users\lisy\Desktop\Pl�ny up�nania tablet\0_Datab�za pl�nov up�nania_Tablet.xlsm"
''
''            'Zabezpe�� a ulo�� zo�it "0_Datab�za pl�nov up�nania_Tablet.xlsm"
''            Application.Workbooks.Open ("C:\Users\lisy\Desktop\Pl�ny up�nania tablet\0_Datab�za pl�nov up�nania_Tablet.xlsm")
''            Workbooks("0_Datab�za pl�nov up�nania_Tablet.xlsm").Worksheets("AIO_Data").Protect Password:="Lis.0123"
''            Workbooks("0_Datab�za pl�nov up�nania_Tablet.xlsm").Save
''            Workbooks("0_Datab�za pl�nov up�nania_Tablet.xlsm").Close
''
''            Application.ScreenUpdating = True   'zapne prekreslovanie obrazovky, t�m sa makro zr�chli
''
''        Else: MsgBox ("Neukladam na tablet")
''             'Zabr�ni zobrezeniu syst�mov�ch hl�ok
''            Application.DisplayAlerts = False
''
'''            Workbooks("0_Datab�za pl�nov up�nania_Tablet.xlsm").Activate
'''            Workbooks("0_Datab�za pl�nov up�nania_Tablet.xlsm").Close
''        End If
'
'End Sub
'
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''Vytvorenie Z�lohy na disk "T" a na tablet'''''''''''''''''''''''''''''''''''''
''        If ThisWorkbook.Name = "AIO_Data" Then
'''            MsgBox ("Ukladam na tablet")
''
''            'Zabr�ni zobrezeniu syst�mov�ch hl�ok
''             Application.DisplayAlerts = False
''
''             'Skryje riadok vzorcov
''             Application.DisplayFormulaBar = False
''    '
''            'Skryje z�hlavia
''                ActiveWindow.DisplayHeadings = False
''            'Skryje mrie�ku
''                ActiveWindow.DisplayGridlines = False
''
''            'Ulo�� a zabezpe�� zo�it "Datab�za pl�nov up�nania.xlsm"
''            Workbooks("AIO_Data").SaveAs "T:\430_F\10_Verejne\10_planovanieLisov\Pl�ny up�nania n�strojov\Datab�za pl�nov up�nania.xlsm"
''            Workbooks("Datab�za pl�nov up�nania.xlsm").Worksheets("AIO_Data").Protect Password:="Lis.0123"
''            Workbooks("Datab�za pl�nov up�nania.xlsm").Save
''
''            'Ulo�� a zabezpe�� zo�it "0_Datab�za pl�nov up�nania_Tablet.xlsm"
''            Workbooks("Datab�za pl�nov up�nania.xlsm").SaveCopyAs "C:\Users\lisy\Desktop\Pl�ny up�nania tablet\0_Datab�za pl�nov up�nania_Tablet.xlsm"
''
''        Else:
''            'MsgBox ("Neukladam na tablet")
''             'Zabr�ni zobrezeniu syst�mov�ch hl�ok
''            Application.DisplayAlerts = False
''
'''            Workbooks("0_Datab�za pl�nov up�nania_Tablet.xlsm").Activate
'''            Workbooks("0_Datab�za pl�nov up�nania_Tablet.xlsm").Close
''        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'End Sub




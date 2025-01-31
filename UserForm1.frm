VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Formul·r - Pl·n upÌnania do lisu"
   ClientHeight    =   16920
   ClientLeft      =   15
   ClientTop       =   465
   ClientWidth     =   4470
   OleObjectBlob   =   "UserForm1.frx":0000
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TlaËÌtko "MINIMALIZOVAç"
Private Sub CommandButton4_Click()

    Application.ScreenUpdating = False
    
        Unload Me
        
        Application.WindowState = xlMinimized
        
        UserForm1.Show
    
    Application.ScreenUpdating = True

End Sub
'TlaËÌtko "MAXIMALIZOVAT"
Private Sub CommandButton6_Click()

    Application.ScreenUpdating = False
        
        UserForm1.Hide
        
        Application.WindowState = xlMaximized
        
        Application.WindowState = xlNormal
        Application.Left = 226
        Application.Top = 1
        Application.Width = 686 '(976)
        Application.Height = 870
        
        UserForm1.Show
    
    Application.ScreenUpdating = True

End Sub
'PRIDRZIAVAC BARAN »APY
'Ked je zaökrtnutÈ pridrziavac baran bez nepovolÌ zaskrtnut Ëapy
Private Sub CheckBox1_Click()

    If CheckBox14.Value = True Then
        CheckBox1.Value = False
    End If
    
    If CheckBox1.Value = True Then
        I = MsgBox("Preverte, Ëi neb˙raj˙ tlaËnÈ Ëapy do barana!", vbOKOnly + vbExclamation, "POZOR!!!")
    End If

End Sub
'PRIDRZIAVAC STOL BEZ
Private Sub CheckBox13_Click()

    If CheckBox3.Value = True Or CheckBox4.Value = True Then
        CheckBox3.Value = False
        CheckBox4.Value = False
    End If

    If CheckBox13.Value = True Then
          CheckBox3.Locked = True
          CheckBox4.Locked = True
    Else: CheckBox3.Locked = False
          CheckBox4.Locked = False
    End If

End Sub
'PRIDRZIAVAC BARAN BEZ
Private Sub CheckBox14_Click()

    If CheckBox1.Value = True Or CheckBox2.Value = True Then
        CheckBox1.Value = False
        CheckBox2.Value = False
    End If

    If CheckBox14.Value = True Then
          CheckBox1.Locked = True
          CheckBox2.Locked = True
    Else: CheckBox1.Locked = False
          CheckBox2.Locked = False
    End If

End Sub
'ZDVIH OBVODOVYCH GDF
Private Sub CheckBox15_Click()
    On Error Resume Next
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"

    If CheckBox15.Value = True Then
          Range("L6").Value = "Zdvih obvodov˝ch GDF"
          CheckBox16.Locked = True
    Else: Range("L6").Value = ""
          CheckBox16.Locked = False
    End If
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    On Error GoTo 0
End Sub
'ZDVIH ODSTAVOVACICH BLOKOV
Private Sub CheckBox16_Click()
    
    On Error Resume Next
    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
        If CheckBox16.Value = True Then
              Range("L6").Value = "V˝öka odstavovacÌch blokov"
              CheckBox15.Locked = True
        Else: Range("L6").Value = ""
              CheckBox15.Locked = False
        End If
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    On Error GoTo 0
End Sub
'PRIDRZIAVAC BARAN GDF
'Ked je zaökrtnutÈ pridrziavac baran bez nepovolÌ zaskrtnut gdf
Private Sub CheckBox2_Click()

    If CheckBox14.Value = True Then
        CheckBox2.Value = False
    End If

End Sub
'PRIDRZIAVAC STOL GDF
'Ked je zaökrtnutÈ pridrziavac stol bez nepovolÌ zaskrtnut gdf
Private Sub CheckBox3_Click()

    If CheckBox13.Value = True Then
        CheckBox3.Value = False
    End If

End Sub

'PRIDRZIAVAC STOL »APY
Private Sub CheckBox4_Click()

'Ked je zaökrtnutÈ pridrziavac stÙl bez nepovolÌ zaskrtnut Ëapy
    If CheckBox13.Value = True Then
        CheckBox4.Value = False
    End If

    If CheckBox4.Value = True Then
    
        If IpmortVsetkychUdajovZAIO_Data_Running = True Or VyËistitPlanUpinania_Running = True Then
'
           MsgBox ("Sub IpmortVsetkychUdajovZAIO_Data or VyËistitPlanUpinania_Running is running!Neprajem si vyznaËiù polohu tlaËn˝ch Ëapov.")
'
        Else
            I = MsgBox("Prajete si vyznaËiù pozÌciu tlaËn˝ch Ëapov?", vbYesNo + vbQuestion, "PozÌcia tlaËn˝ch Ëapov")
            Select Case I
                Case vbNo
            '        MsgBox ("Nie")
                Case vbYes
                
                    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
                    
                    Application.ScreenUpdating = False
                '
                    Unload Me
                    
                    Application.WindowState = xlMaximized
                    
                    Application.WindowState = xlNormal
                        Application.Left = 226
                        Application.Top = 1
                        Application.Width = 686 '(976)
                        Application.Height = 870
                '    UserForm1.Show
                '
                'Nastavi raster stola do laveho horneho rohu okna
                    ActiveWindow.ScrollRow = 33
                    ActiveWindow.ScrollColumn = 4
                '   ActiveWindow.SmallScroll ToRight:=4
                '
                'Skryje riadok vzorcov
                    Application.DisplayFormulaBar = False
                    '
                'Skryje z·hlavia
                    ActiveWindow.DisplayHeadings = False
                
                'Ak je p·s s n·strojmi pripnut˝, tak ho zbalÌ
        
                    If Application.CommandBars("Ribbon").Height > 100 Then 'ActiveWindow.ToggleRibbon
                        MsgBox ("P·s je pripnut˝")
                        Application.CommandBars.ExecuteMso "MinimizeRibbon"
                    End If
    
    
                '
                'Vyberie tlaËidla pre vklad centrovania, tlacn˝ch Ëapov a volnÈho miesta _
                    a posunie ich dolava
                    Worksheets("AIO_Plan").Shapes.Range(Array("Group 1")).Select
                
                        With Selection.ShapeRange
                             .Left = 494
                             .Top = 590 '363Hodnota,ked su pozn·mky dole
                
                        End With
                'ZmensÌ okno excelu do laveho dolnÈho rohu
                    Application.WindowState = xlNormal
                        Application.Left = 1
                        Application.Top = 554.5
                        Application.Width = 502 '485
                        Application.Height = 318
                    
                    Range("U41").Select
                     
                '   UserForm1.Hide
                
                    Application.ScreenUpdating = True
                
                    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
            End Select
        End If
    End If

End Sub

Private Sub Label24_Click()

End Sub

''PO opusteni pola ËÌslo n·stroja(PotrebnÈ pre doplnenie udajov zo s˙boru ¥¥AIO_Data¥¥)
'Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'
''''    Call ZavolaZatvoritFormular
'    Unload Me
'    Application.WindowState = xlMaximized
'
'        Application.WindowState = xlNormal
'        Application.Left = 226
'        Application.Top = 130 '1
'        Application.Width = 686 '(976)
'        Application.Height = 740 '870
'
''    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'
'    Application.Wait Now + TimeValue("00:00:01") 'zdrûanie 1 sekunda
'
'    Range("S1").Value = TextBox1.Text
'
'End Sub

'Centrovanie LHR
Private Sub TextBox17_AfterUpdate()
'Funguje po Aktualiz·cii overenÌ ˙daje v textboxe17 (LH ötvrtina riadky)

    If TextBox17.Text = "" Or TextBox17.Text = "1" Or TextBox17.Text = "2" Or TextBox17.Text = "3" Or TextBox17.Text = "4" Or _
        TextBox17.Text = "5" Or TextBox17.Text = "6" Or TextBox17.Text = "7" Or TextBox17.Text = "8" Or _
        TextBox17.Text = "150" Or TextBox17.Text = "300" Or TextBox17.Text = "450" Or _
        TextBox17.Text = "600" Or TextBox17.Text = "750" Or TextBox17.Text = "900" Or TextBox17.Text = "1050" Then
        Range("T28").Value = TextBox17.Text
    Else: TextBox17.Text = ""
        I = MsgBox("Zadajte hodnotu:" & vbCrLf & _
        "1; 2; 3; 4; 5; 6; 7; 8" & vbCrLf & _
        "alebo" & vbCrLf & _
        "150; 300; 450; 600; 750; 900; 1050", vbOKOnly + vbCritical, "Neplatn˝ ˙daj!")
    End If

End Sub
'Centrovanie PHS
Private Sub TextBox19_AfterUpdate()
'Funguje po Aktualiz·cii overenÌ ˙daje v TextBox19 (PH ötvrtina stÂpce)

    If TextBox19.Text = "" Or TextBox19.Text = "1" Or TextBox19.Text = "2" Or TextBox19.Text = "3" Or TextBox19.Text = "4" Or _
        TextBox19.Text = "5" Or TextBox19.Text = "6" Or TextBox19.Text = "7" Or TextBox19.Text = "8" Or TextBox19.Text = "9" Or _
        TextBox19.Text = "10" Or TextBox19.Text = "11" Or TextBox19.Text = "12" Or TextBox19.Text = "13" Or TextBox19.Text = "14" Or _
        TextBox19.Text = "15" Or TextBox19.Text = "16" Or _
        TextBox19.Text = "150" Or TextBox19.Text = "300" Or TextBox19.Text = "450" Or _
        TextBox19.Text = "600" Or TextBox19.Text = "750" Or TextBox19.Text = "900" Or TextBox19.Text = "1050" Or TextBox19.Text = "1200" Or _
        TextBox19.Text = "1350" Or TextBox19.Text = "1500" Or TextBox19.Text = "1650" Or TextBox19.Text = "1800" Or TextBox19.Text = "1950" Or _
        TextBox19.Text = "2100" Or TextBox19.Text = "2250" Then
        Range("W29").Value = TextBox19.Text
    Else: TextBox19.Text = ""
        I = MsgBox("Zadajte hodnotu:" & vbCrLf & _
        "1; 2; 3; 4; 5; 6; 7; 8; 9; 10; 11; 12; 13; 14; 15; 16; 17" & vbCrLf & _
        "alebo" & vbCrLf & _
        "150; 300; 450; 600; 750; 900; 1050; 1200; 1350; 1500; 1650; 1800; 1950; 2100; 2250; 2400", vbOKOnly + vbCritical, "Neplatn˝ ˙daj!")
    End If

End Sub
'Centrovanie PDS
Private Sub TextBox21_AfterUpdate()
'Funguje po Aktualiz·cii overenÌ ˙daje v TextBox21 (PD ötvrtina stÂpce)

    If TextBox21.Text = "" Or TextBox21.Text = "1" Or TextBox21.Text = "2" Or TextBox21.Text = "3" Or TextBox21.Text = "4" Or _
        TextBox21.Text = "5" Or TextBox21.Text = "6" Or TextBox21.Text = "7" Or TextBox21.Text = "8" Or TextBox21.Text = "9" Or _
        TextBox21.Text = "10" Or TextBox21.Text = "11" Or TextBox21.Text = "12" Or TextBox21.Text = "13" Or TextBox21.Text = "14" Or _
        TextBox21.Text = "15" Or TextBox21.Text = "16" Or _
        TextBox21.Text = "150" Or TextBox21.Text = "300" Or TextBox21.Text = "450" Or _
        TextBox21.Text = "600" Or TextBox21.Text = "750" Or TextBox21.Text = "900" Or TextBox21.Text = "1050" Or TextBox21.Text = "1200" Or _
        TextBox21.Text = "1350" Or TextBox21.Text = "1500" Or TextBox21.Text = "1650" Or TextBox21.Text = "1800" Or TextBox21.Text = "1950" Or _
        TextBox21.Text = "2100" Or TextBox21.Text = "2250" Then
        Range("W30").Value = TextBox21.Text
    Else: TextBox21.Text = ""
        I = MsgBox("Zadajte hodnotu:" & vbCrLf & _
        "1; 2; 3; 4; 5; 6; 7; 8; 9; 10; 11; 12; 13; 14; 15; 16; 17" & vbCrLf & _
        "alebo" & vbCrLf & _
        "150; 300; 450; 600; 750; 900; 1050; 1200; 1350; 1500; 1650; 1800; 1950; 2100; 2250; 2400", vbOKOnly + vbCritical, "Neplatn˝ ˙daj!")
    End If
    
End Sub
'Centrovanie LDS
Private Sub TextBox22_AfterUpdate()
'Funguje po Aktualiz·cii overenÌ ˙daje v TextBox22 (LD ötvrtina stÂpce)

    If TextBox22.Text = "" Or TextBox22.Text = "1" Or TextBox22.Text = "2" Or TextBox22.Text = "3" Or TextBox22.Text = "4" Or _
        TextBox22.Text = "5" Or TextBox22.Text = "6" Or TextBox22.Text = "7" Or TextBox22.Text = "8" Or TextBox22.Text = "9" Or _
        TextBox22.Text = "10" Or TextBox22.Text = "11" Or TextBox22.Text = "12" Or TextBox22.Text = "13" Or TextBox22.Text = "14" Or _
        TextBox22.Text = "15" Or TextBox22.Text = "16" Or _
        TextBox22.Text = "150" Or TextBox22.Text = "300" Or TextBox22.Text = "450" Or _
        TextBox22.Text = "600" Or TextBox22.Text = "750" Or TextBox22.Text = "900" Or TextBox22.Text = "1050" Or TextBox22.Text = "1200" Or _
        TextBox22.Text = "1350" Or TextBox22.Text = "1500" Or TextBox22.Text = "1650" Or TextBox22.Text = "1800" Or TextBox22.Text = "1950" Or _
        TextBox22.Text = "2100" Or TextBox22.Text = "2250" Then
        Range("S30").Value = TextBox22.Text
    Else: TextBox22.Text = ""
        I = MsgBox("Zadajte hodnotu:" & vbCrLf & _
        "1; 2; 3; 4; 5; 6; 7; 8; 9; 10; 11; 12; 13; 14; 15; 16; 17" & vbCrLf & _
        "alebo" & vbCrLf & _
        "150; 300; 450; 600; 750; 900; 1050; 1200; 1350; 1500; 1650; 1800; 1950; 2100; 2250; 2400", vbOKOnly + vbCritical, "Neplatn˝ ˙daj!")
    End If

End Sub
'Centrovanie PDR
Private Sub TextBox23_AfterUpdate()
'Funguje po Aktualiz·cii overenÌ ˙daje v textboxe23 (PD ötvrtina riadky)

    If TextBox23.Text = "" Or TextBox23.Text = "1" Or TextBox23.Text = "2" Or TextBox23.Text = "3" Or TextBox23.Text = "4" Or _
        TextBox23.Text = "5" Or TextBox23.Text = "6" Or TextBox23.Text = "7" Or TextBox23.Text = "8" Or _
        TextBox23.Text = "150" Or TextBox23.Text = "300" Or TextBox23.Text = "450" Or _
        TextBox23.Text = "600" Or TextBox23.Text = "750" Or TextBox23.Text = "900" Or TextBox23.Text = "1050" Then
        Range("V31").Value = TextBox23.Text
    Else: TextBox23.Text = ""
        I = MsgBox("Zadajte hodnotu:" & vbCrLf & _
        "1; 2; 3; 4; 5; 6; 7; 8" & vbCrLf & _
        "alebo" & vbCrLf & _
        "150; 300; 450; 600; 750; 900; 1050", vbOKOnly + vbCritical, "Neplatn˝ ˙daj!")
    End If

End Sub
'Centrovanie LDR
Private Sub TextBox24_AfterUpdate()
'Funguje po Aktualiz·cii overenÌ ˙daje v textboxe24 (LD ötvrtina riadky)

    If TextBox24.Text = "" Or TextBox24.Text = "1" Or TextBox24.Text = "2" Or TextBox24.Text = "3" Or TextBox24.Text = "4" Or _
        TextBox24.Text = "5" Or TextBox24.Text = "6" Or TextBox24.Text = "7" Or TextBox24.Text = "8" Or _
        TextBox24.Text = "150" Or TextBox24.Text = "300" Or TextBox24.Text = "450" Or _
        TextBox24.Text = "600" Or TextBox24.Text = "750" Or TextBox24.Text = "900" Or TextBox24.Text = "1050" Then
        Range("T31").Value = TextBox24.Text
    Else: TextBox24.Text = ""
        I = MsgBox("Zadajte hodnotu:" & vbCrLf & _
        "1; 2; 3; 4; 5; 6; 7; 8" & vbCrLf & _
        "alebo" & vbCrLf & _
        "150; 300; 450; 600; 750; 900; 1050", vbOKOnly + vbCritical, "Neplatn˝ ˙daj!")
    End If

End Sub
Private Sub TextBox25_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'OVERI CI SA ROVNA POCET NAPOCITANYCH TLACNYCH CAPOV Z MODELU _
S POCTOM ZAKRESLENYCH TLACNYCH CAPOV

    If TextBox25.Text <> Range("AN28").Text Then
        I = MsgBox("Skontrolujte poËet tlaËn˝ch Ëapov!", vbOKOnly + vbExclamation, "PoËet tlaËn˝ch Ëapov")
    End If

End Sub
'Centrovanie PHR
Private Sub TextBox26_AfterUpdate()
'Funguje po Aktualiz·cii overenÌ ˙daje v textboxe26 (PH ötvrtina riadky)

    If TextBox26.Text = "" Or TextBox26.Text = "1" Or TextBox26.Text = "2" Or TextBox26.Text = "3" Or TextBox26.Text = "4" Or _
        TextBox26.Text = "5" Or TextBox26.Text = "6" Or TextBox26.Text = "7" Or TextBox26.Text = "8" Or _
        TextBox26.Text = "150" Or TextBox26.Text = "300" Or TextBox26.Text = "450" Or _
        TextBox26.Text = "600" Or TextBox26.Text = "750" Or TextBox26.Text = "900" Or TextBox26.Text = "1050" Then
        Range("V28").Value = TextBox26.Text
    Else: TextBox26.Text = ""
        I = MsgBox("Zadajte hodnotu:" & vbCrLf & _
        "1; 2; 3; 4; 5; 6; 7; 8" & vbCrLf & _
        "alebo" & vbCrLf & _
        "150; 300; 450; 600; 750; 900; 1050", vbOKOnly + vbCritical, "Neplatn˝ ˙daj!")
    End If

End Sub
'Centrovanie LHS
Private Sub TextBox27_AfterUpdate()
'Funguje po Aktualiz·cii overenÌ ˙daje v TextBox27 (LH ötvrtina stÂpce)

    If TextBox27.Text = "" Or TextBox27.Text = "1" Or TextBox27.Text = "2" Or TextBox27.Text = "3" Or TextBox27.Text = "4" Or _
        TextBox27.Text = "5" Or TextBox27.Text = "6" Or TextBox27.Text = "7" Or TextBox27.Text = "8" Or TextBox27.Text = "9" Or _
        TextBox27.Text = "10" Or TextBox27.Text = "11" Or TextBox27.Text = "12" Or TextBox27.Text = "13" Or TextBox27.Text = "14" Or _
        TextBox27.Text = "15" Or TextBox27.Text = "16" Or _
        TextBox27.Text = "150" Or TextBox27.Text = "300" Or TextBox27.Text = "450" Or _
        TextBox27.Text = "600" Or TextBox27.Text = "750" Or TextBox27.Text = "900" Or TextBox27.Text = "1050" Or TextBox27.Text = "1200" Or _
        TextBox27.Text = "1350" Or TextBox27.Text = "1500" Or TextBox27.Text = "1650" Or TextBox27.Text = "1800" Or TextBox27.Text = "1950" Or _
        TextBox27.Text = "2100" Or TextBox27.Text = "2250" Then
        Range("S29").Value = TextBox27.Text
    Else: TextBox27.Text = ""
        I = MsgBox("Zadajte hodnotu:" & vbCrLf & _
        "1; 2; 3; 4; 5; 6; 7; 8; 9; 10; 11; 12; 13; 14; 15; 16; 17" & vbCrLf & _
        "alebo" & vbCrLf & _
        "150; 300; 450; 600; 750; 900; 1050; 1200; 1350; 1500; 1650; 1800; 1950; 2100; 2250; 2400", vbOKOnly + vbCritical, "Neplatn˝ ˙daj!")
    End If

End Sub
'POZNAMKY 1.RIADOK
Private Sub TextBox33_Change()
    
    Range("I51").Value = TextBox33.Text

End Sub
'POZNAMKY 3.RIADOK
Private Sub TextBox34_Change()
    
    Range("B53").Value = TextBox34.Text

End Sub
'POZNAMKY 2.RIADOK
Private Sub TextBox35_Change()
    
    Range("B52").Value = TextBox35.Text

End Sub
'POZNAMKY 4.RIADOK
Private Sub TextBox36_Change()

    Range("B54").Value = TextBox36.Text

End Sub
'POZNAMKY 5.RIADOK
Private Sub TextBox37_Change()

    Range("B55").Value = TextBox37.Text

End Sub
'POZNAMKY 8.RIADOK
Private Sub TextBox38_Change()
    
    Range("B58").Value = TextBox38.Text

End Sub
'POZNAMKY 7.RIADOK
Private Sub TextBox39_Change()

    Range("B57").Value = TextBox39.Text

End Sub
'POZNAMKY 6.RIADOK
Private Sub TextBox40_Change()

    Range("B56").Value = TextBox40.Text

End Sub
'POZNAMKY 9.RIADOK
Private Sub TextBox41_Change()

    Range("B59").Value = TextBox41.Text

End Sub
'POZNAMKY 12.RIADOK
Private Sub TextBox42_Change()

    Range("B62").Value = TextBox42.Text

End Sub
'POZNAMKY 11.RIADOK
Private Sub TextBox43_Change()

    Range("B61").Value = TextBox43.Text

End Sub
'POZNAMKY 10.RIADOK
Private Sub TextBox44_Change()

    Range("B60").Value = TextBox44.Text

End Sub

 
Private Sub UserForm_Activate()

'--------------------------------------------------------------------
'Ak je po otvorenÌ formul·ra TB2 pr·zdny, potom je pr·zdny ja textbox1 a nastavÌ TB1 pre zad·vanie ˙dajov _
inak TB1 sa rovn· bunke S1
    If TextBox2.Text = "" Then
        TextBox1.Text = ""
        TextBox1.SetFocus 'NastavÌ dan˝ textbox pre zad·vanie ˙dajov
'        MsgBox ("Bingo!")
    Else:  TextBox1.Text = Range("S1").Value
'        MsgBox ("DruhÈ Bingo!")
    End If
'--------------------------------------------------------------------

End Sub

'ZATVORENIE CEZ KRIZIK
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        Cancel = False
    End If
    
'Po zatvorenÌ formal·ra cez krÌûik _
zmenÌ velkosù okna podæa parametrov niûöie
'''''    Application.WindowState = xlNormal
'''''        Application.Left = 226
'''''        Application.Top = 1
'''''        Application.Width = 686 '(976)
'''''        Application.Height = 870
        
'    Worksheets("AIO_Plan").Protect Password:="Lis.0123"

End Sub

'______________________________________________________________________________________________________________________________________________________
''TLACITKO "Uloûiù formul·r"
'Private Sub CommandButton2_Click()
'
'    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
'
'    Fcesta = "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…"
'    FcestaJPG = "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…\PU_JPG"
'    FcestaPDF = "C:\Users\lisy\Desktop\Pl·ny upÌnania\PU_NOV…\PU_PDF"
'    FcestaPdfNaSiet = "T:\430_F\10_Verejne\10_planovanieLisov\Pl·ny upÌnania n·strojov"
'    FcestaPdfTablet = "C:\Users\lisy\Desktop\Pl·ny upÌnania tablet"
'    FCisloNastroja = Sheets("AIO_Plan").Range("S1").Text
'    FOperacia = Sheets("AIO_Plan").Range("AM1").Text
'    FKrok = Sheets("AIO_Plan").Range("AM3").Text
'    FCisloDielu = Sheets("AIO_Plan").Range("S3").Text
'
'    If Range("AM3").Text = "" Then
'        NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_Pl·n upÌnania"
'        Range("AJ3:AL3").NumberFormat = ";;;"
'    Else
'        NazovPlanuUpinania = FCisloNastroja & "_OP" & FOperacia & "_" & FCisloDielu & "_S" & FKrok & "_Pl·n upÌnania"
'        Range("AJ3:AL3").NumberFormat = "@"
'    End If
'
'    i = MsgBox("Uloûiù ako:  " & NazovPlanuUpinania, vbYesNo + vbQuestion, "Uloûiù ako")
'
'    Select Case i
'        Case vbNo
'            Unload Me
'            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'    '       MsgBox ("Nie")
'        Case vbYes
'
'            Unload Me
'
'            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'            'Application.DisplayAlerts = False'Zabr·ni zobrezeniu systÈmov˝ch hl·öok
'
'        'ULOZI AKO ".xlsm"
'            ActiveWorkbook.SaveAs Filename:=Fcesta & "\" & NazovPlanuUpinania & ".xlsm"
'
'        'ULOZI AKO "Pdf"
'            Worksheets("AIO_Plan").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'                    FcestaPDF & "\" & NazovPlanuUpinania & ".pdf", _
'                    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
'                    :=False, OpenAfterPublish:=True
'
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
'            Call FormatNaSirku
'
'            Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
'
'        'ULOZI AKO ".jpg"----------------------------------------------------------------------
'            Dim cht As ChartObject
'            Dim ActiveShape As Shape
'
'            If ActiveWindow.DisplayGridlines = True Then
'
'                ActiveWindow.DisplayGridlines = False
'                'Confirm if a Cell Range is currently selected
'                '  If TypeName(Selection) <> "Range" Then
'                '    MsgBox "You do not have a single shape selected!"
'                '    Exit Sub
'                '  End If
'
'                'Copy/Paste Cell Range as a Picture
'                  Range("A1:BO50").Copy
'                  Worksheets("AIO_Plan").Pictures.Paste(link:=False).Select
'                  Set ActiveShape = Worksheets("AIO_Plan").Shapes(ActiveWindow.Selection.Name)
'
'                'Create a temporary chart object (same size as shape)
'                  Set cht = Worksheets("AIO_Plan").ChartObjects.Add( _
'                    Left:=1, _
'                    Width:=735, _
'                    Top:=1, _
'                    Height:=483)
'
'                'Format temporary chart to have a transparent background
'                  cht.ShapeRange.Fill.Visible = msoFalse
'                  cht.ShapeRange.Line.Visible = msoFalse
'
'                'Copy/Paste Shape inside temporary chart
'                  ActiveShape.Copy
'                  cht.Activate
'                  ActiveChart.Paste
'
'                'Save chart to User's Desktop as PNG File
'                  cht.Chart.Export Filename:=FcestaJPG & "\" & NazovPlanuUpinania & ".jpg"
'
'                'Delete temporary Chart
'                  cht.Delete
'                  ActiveShape.Delete
'
'                ActiveWindow.DisplayGridlines = True
'
'            End If
'
'            If ActiveWindow.DisplayGridlines = False Then
'
'                'Confirm if a Cell Range is currently selected
'                '  If TypeName(Selection) <> "Range" Then
'                '    MsgBox "You do not have a single shape selected!"
'                '    Exit Sub
'                '  End If
'
'                'Copy/Paste Cell Range as a Picture
'                  Range("A1:BO50").Copy
'                  Worksheets("AIO_Plan").Pictures.Paste(link:=False).Select
'                  Set ActiveShape = Worksheets("AIO_Plan").Shapes(ActiveWindow.Selection.Name)
'
'                'Create a temporary chart object (same size as shape)
'                  Set cht = Worksheets("AIO_Plan").ChartObjects.Add( _
'                    Left:=1, _
'                    Width:=735, _
'                    Top:=1, _
'                    Height:=483)
'
'                'Format temporary chart to have a transparent background
'                  cht.ShapeRange.Fill.Visible = msoFalse
'                  cht.ShapeRange.Line.Visible = msoFalse
'
'                'Copy/Paste Shape inside temporary chart
'                  ActiveShape.Copy
'                  cht.Activate
'                  ActiveChart.Paste
'
'                'Save chart to User's Desktop as PNG File
'                  cht.Chart.Export Filename:=FcestaJPG & "\" & NazovPlanuUpinania & ".jpg"
'
'                'Delete temporary Chart
'                  cht.Delete
'                  ActiveShape.Delete
'
'            End If
'
'            Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'
'            Call FormatNaVysku
'
'        'OTVORI JPG SUBOR----------------------------------------------------------------------
'            VBA.Shell "Explorer.exe " & FcestaJPG & "\" & NazovPlanuUpinania & ".jpg"
'
'        'V SUBORE "AIO_Data" VYZNACI MOZNOST UPNUTIA NASTROJA DO LISU + ZE BOL VYTVORENY PLAN UPINANIA V NOVOM FORMULARI
'        '---------------------------------------------------------------------------------------
'            Call OtvorNajdiVyznacAll
'
'            Workbooks(NazovPlanuUpinania & ".xlsm").Worksheets("AIO_Plan").Activate
'
'            Application.WindowState = xlNormal
'                Application.Left = 226
'                Application.Top = 1
'                Application.Width = 686 '(976)
'                Application.Height = 870
'
'            i = MsgBox("Prajete si n·vrat do formul·ra?", vbYesNo + vbQuestion, "N·vrat")
'
'            Select Case i
'                Case vbNo
'
'                Worksheets("AIO_Plan").Protect Password:="Lis.0123"
'            '   MsgBox ("Nie")
'
'                Case vbYes
'
'                UserForm1.Show
'            End Select
'
'    End Select
'
'    'Application.DisplayAlerts = True'PovolÌ zobrezeniu systÈmov˝ch hl·öok
'
'End Sub
'____________________________________________________________________________________________________________________________________________
'

'TLACITKO "Zatvoriù formul·r"
Private Sub CommandButton3_Click()

    Unload Me
    
'''''    Application.WindowState = xlNormal
'''''        Application.Left = 226
'''''        Application.Top = 1
'''''        Application.Width = 668 '686 '(976)
'''''        Application.Height = 870
        
'    Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    
End Sub
'SMER LISOVANIA HORE
Private Sub CheckBox5_Click()

    On Error Resume Next

    If CheckBox5.Value = True Then
        Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 3", _
        "Straight Arrow Connector 13")).Visible = msoTrue
    Else: Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 3", _
        "Straight Arrow Connector 13")).Visible = msoFalse
    End If
    
    On Error GoTo 0
    
    '    Selection.ShapeRange.Flip msoFlipHorizontal
    
End Sub
'SMER LISOVANIA DOLE
Private Sub CheckBox6_Click()

    On Error Resume Next

    If CheckBox6.Value = True Then
        Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 23", _
        "Straight Arrow Connector 24")).Visible = msoTrue
    Else: Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 23", _
        "Straight Arrow Connector 24")).Visible = msoFalse
    End If
    
    On Error GoTo 0
       
End Sub
'SMER LISOVANIA VPRAVO
Private Sub CheckBox7_Click()

    On Error Resume Next

    If CheckBox7.Value = True Then
        Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 19", _
        "Straight Arrow Connector 20")).Visible = msoTrue
    Else: Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 19", _
        "Straight Arrow Connector 20")).Visible = msoFalse
    End If
    
    On Error GoTo 0
       
End Sub
'SMER LISOVANIA VLAVO
Private Sub CheckBox8_Click()

    On Error Resume Next
    
    If CheckBox8.Value = True Then
        Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 22", _
        "Straight Arrow Connector 21")).Visible = msoTrue
    Else: Worksheets("AIO_Plan").Shapes.Range(Array("Straight Arrow Connector 22", _
        "Straight Arrow Connector 21")).Visible = msoFalse
    End If

    On Error GoTo 0

End Sub
'CISLO DIELU
Private Sub TextBox1_Change()

    'TextBox1.ControlTipText = "Napr.:123.456.789"
    ControlTipText = TextBox1.ControlTipText
    
'    Range("S1").Value = TextBox1.Text

End Sub
'OPERACIA
Private Sub TextBox2_Change()

    Range("AK1").Value = TextBox2.Text

End Sub
'DIALOGOVE OKNO "Smer lisovania"
Private Sub TextBox21_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    I = MsgBox("ProsÌm vyznaËte smer lisovania ", vbOKOnly + vbInformation, "Smer lisovania")

End Sub
'CISLO NASTROJA
Private Sub TextBox3_Change()

    ControlTipText = TextBox1.ControlTipText
    
    Range("S3").Value = TextBox3.Text

End Sub
'CISLO PROJEKTU (VP)
Private Sub TextBox4_Change()

    ControlTipText = TextBox1.ControlTipText
    
    Range("C4").Value = TextBox4.Text

End Sub
'OZNACENIE DIELU
Private Sub TextBox5_Change()

    ControlTipText = TextBox5.ControlTipText
    
    Range("K4").Value = TextBox5.Text

End Sub
'NAZOV PROJEKTU
Private Sub TextBox6_Change()

    ControlTipText = TextBox6.ControlTipText
    
    Range("AB4").Value = TextBox6.Text

End Sub
'ROZMER NASTROJA D
Private Sub TextBox10_Change()

    If TextBox10.Value > 3730 Then
        CheckBox9.Locked = True
        CheckBox9.ControlTipText = "N·stroj sa do lisu nezmestÌ! DÂûka stolu v lise 1 je 3700mm. Maxim·lna dÂûka n·stroja v lise 1 je 3730mm"
    Else: CheckBox9.Locked = False
        CheckBox9.ControlTipText = "Preveriù veækosù n·stroja, veækosù pridrûiavaËov, moûnosù upnutia OT"
    End If
    
    If TextBox10.Value > 4220 Then
        CheckBox10.Locked = True
        CheckBox10.ControlTipText = "N·stroj sa do lisu nezmestÌ! DÂûka stolu v lise 2 je 4200mm. Maxim·lna dÂûka n·stroja v lise 2 je 4220mm"
    Else: CheckBox10.Locked = False
        CheckBox10.ControlTipText = "Preveriù veækosù n·stroja, veækosù pridrûiavaËov, moûnosù upnutia OT"
    End If
    
    ControlTipText = TextBox10.ControlTipText
    
    Range("W5").Value = TextBox10.Text

End Sub
'ROZMER NASTROJA ä
Private Sub TextBox11_Change()

    ControlTipText = TextBox11.ControlTipText
    
    Range("AD5").Value = TextBox11.Text

End Sub
'ROZMER NASTROJA V
Private Sub TextBox12_Change()

    ControlTipText = TextBox12.ControlTipText
    
    Range("AK5").Value = TextBox12.Text
    
End Sub
'ROZMER NASTROJA V
'ak sa rovna bunka 0, tak po odchode z bunky sa vycistÌ
Private Sub TextBox12_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Range("AK5").Value = CDbl(Range("AK5").Value) 'KoneËne funguje tak ako chcem
    
    If Range("AK5").Value = 0 Then
       Range("AK5:AM5").ClearContents
    End If
    
    If TextBox12.Value < 900 Then
        CheckBox12.Locked = True
        CheckBox12.ControlTipText = "N·stroj je prÌliö nÌzky! Minim·lna v˝öka n·stroja v lise 4 je 900mm"
    Else: CheckBox12.Locked = False
        CheckBox12.ControlTipText = "Preveriù veækosù n·stroja, veækosù pridrûiavaËov, moûnosù upnutia OT"
    End If

End Sub
'ROZMER UPINACEJ PLOCHY OT D
Private Sub TextBox13_Change()

    ControlTipText = TextBox13.ControlTipText
    
    Range("W6").Value = TextBox13.Text
    
End Sub
'VZDIALENOST MEDZI POLOBLUKMI UPINACICH DRAZOK
Private Sub TextBox14_Change()

    Range("AI5").Value = TextBox14.Text

End Sub
'VZDIALENOST MEDZI POLOBLUKMI UPINACICH DRAZOK
Private Sub TextBox14_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
    Range("AD5:AF5").Select
    Range("AD5").Comment.Text Text:="Vzdialenosù medzi polobl˙kmi upÌnacÌch dr·ûok na prednej a zadnej strane OT je " & TextBox14.Text & "mm" 'Chr(10) &
    'Pre lisy Onapres je maximum 2170mm, idealne 2120mm vtedy by mal byù T-kameÚ cel˝ skryt˝ v dr·ûke
    
    Worksheets("AIO_Plan").Protect Password:="Lis.0123"

End Sub
'ZDVIH OBVODOVYCH GDF
Private Sub TextBox15_Change()

    ControlTipText = TextBox15.ControlTipText
    
    Range("W6").Value = TextBox15.Text
    
End Sub
'ZDVIH OBVODOVYCH GDF
'ak sa rovna bunka 0, tak po odchode z bunky sa vycistÌ
Private Sub TextBox15_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Range("W6").Value = CDbl(Range("W6").Value) 'KoneËne funguje tak ako chcem
    
    If Range("W6").Value = 0 Then
        Range("W6:Y6").ClearContents
    '   Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    '   Range("L6") = ""
    '   Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If

End Sub
'UPINACIA VYSKA NASTROJA V
Private Sub TextBox16_Change()

    ControlTipText = TextBox16.ControlTipText
    
    Range("AK6").Value = TextBox16.Text
    
End Sub
'UPINACIA VYSKA NASTROJA V
'ak sa rovna bunka 0, tak po odchode z bunky sa vycistÌ
Private Sub TextBox16_Exit(ByVal Cancel As MSForms.ReturnBoolean) 'DialogovÈ okno "PridrûiavaËe"

    Range("AK6").Value = CDbl(Range("AK6").Value) 'KoneËne funguje tak ako chcem
    
    If Range("AK6").Value = 0 Then
        Range("AK6:AM6").ClearContents
    End If
    
    I = MsgBox("ProsÌm vyznaËte prÌtomnosù pridrûiavaËa alebo GDF" & vbCrLf & _
    "a preverte moûnosù upnutia n·stroja do lisov PWS", vbOKOnly + vbInformation, "PridrûiavaËe")
    
End Sub

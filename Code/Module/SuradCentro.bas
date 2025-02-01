Attribute VB_Name = "SuradCentro"
Sub SuradniceCentrovaniaLH()
'Lavej hornej štvrtine rastra stola vyznaèí pozíciu centrovacieho èapu podla zadanıch suradníc
    
    If Range("T28") > 0 And Range("T28") <= 8 And Range("S29") > 0 And Range("S29") <= 16 Then
    
    'Ståpce
        If Range("S29").Value = 1 Then
            s = "U"
        End If
        If Range("S29").Value = 2 Then
            s = "T"
        End If
        If Range("S29").Value = 3 Then
            s = "S"
        End If
        If Range("S29").Value = 4 Then
            s = "R"
        End If
        If Range("S29").Value = 5 Then
            s = "Q"
        End If
        If Range("S29").Value = 6 Then
            s = "P"
        End If
        If Range("S29").Value = 7 Then
            s = "O"
        End If
        If Range("S29").Value = 8 Then
            s = "N"
        End If
        If Range("S29").Value = 9 Then
            s = "M"
        End If
        If Range("S29").Value = 10 Then
            s = "L"
        End If
        If Range("S29").Value = 11 Then
            s = "K"
        End If
        If Range("S29").Value = 12 Then
            s = "J"
        End If
        If Range("S29").Value = 13 Then
            s = "I"
        End If
        If Range("S29").Value = 14 Then
            s = "H"
        End If
        If Range("S29").Value = 15 Then
            s = "G"
        End If
        If Range("S29").Value = 16 Then
            s = "F"
        End If
        If Range("S29").Value = 17 Then
            s = "E"
        End If
        
    'Riadky
        If Range("T28").Value = 1 Then
            r = "41"
        End If
        If Range("T28").Value = 2 Then
            r = "40"
        End If
        If Range("T28").Value = 3 Then
            r = "39"
        End If
        If Range("T28").Value = 4 Then
            r = "38"
        End If
        If Range("T28").Value = 5 Then
            r = "37"
        End If
        If Range("T28").Value = 6 Then
            r = "36"
        End If
        If Range("T28").Value = 7 Then
            r = "35"
        End If
        If Range("T28").Value = 8 Then
            r = "34"
        End If
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
            
        Range("B29").Copy
            
        Range(s & r).Select
'        Worksheets("AIO_Plan").Paste
        Selection.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
   '    Range("T29").Select
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If

'FUNGUJE!!!
End Sub

Sub SuradniceCentrovaniaPH()
'V pravej hornej štvrtine rastra stola vyznaèí pozíciu centrovacieho èapu podla zadanıch suradníc

    If Range("V28") > 0 And Range("V28") <= 8 And Range("W29") > 0 And Range("W29") <= 16 Then
    
    'Ståpce
        If Range("W29").Value = 1 Then
            s = "U"
        End If
        If Range("W29").Value = 2 Then
            s = "V"
        End If
        If Range("W29").Value = 3 Then
            s = "W"
        End If
        If Range("W29").Value = 4 Then
            s = "X"
        End If
        If Range("W29").Value = 5 Then
            s = "Y"
        End If
        If Range("W29").Value = 6 Then
            s = "Z"
        End If
        If Range("W29").Value = 7 Then
            s = "AA"
        End If
        If Range("W29").Value = 8 Then
            s = "AB"
        End If
        If Range("W29").Value = 9 Then
            s = "AC"
        End If
        If Range("W29").Value = 10 Then
            s = "AD"
        End If
        If Range("W29").Value = 11 Then
            s = "AE"
        End If
        If Range("W29").Value = 12 Then
            s = "AF"
        End If
        If Range("W29").Value = 13 Then
            s = "AG"
        End If
        If Range("W29").Value = 14 Then
            s = "AH"
        End If
        If Range("W29").Value = 15 Then
            s = "AI"
        End If
        If Range("W29").Value = 16 Then
            s = "AJ"
        End If
        If Range("W29").Value = 17 Then
            s = "AK"
        End If
        
    'Riadky
        If Range("V28").Value = 1 Then
            r = "41"
            End If
        If Range("V28").Value = 2 Then
            r = "40"
        End If
        If Range("V28").Value = 3 Then
            r = "39"
        End If
        If Range("V28").Value = 4 Then
            r = "38"
        End If
        If Range("V28").Value = 5 Then
            r = "37"
        End If
        If Range("V28").Value = 6 Then
            r = "36"
        End If
        If Range("V28").Value = 7 Then
            r = "35"
        End If
        If Range("V28").Value = 8 Then
            r = "34"
        End If
        
    
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
        Range("B29").Copy
        
        Range(s & r).Select
'        Worksheets("AIO_Plan").Paste
        Selection.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    '   Range("V29").Select
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If
'FUNGUJE!!!
End Sub

Sub SuradniceCentrovaniaLD()
'Lavej dolnej štvrtine rastra stola vyznaèí pozíciu centrovacieho èapu podla zadanıch suradníc

    If Range("T31") > 0 And Range("T31") <= 8 And Range("S30") > 0 And Range("S30") <= 16 Then
    
    'Ståpce
        If Range("S30").Value = 1 Then
            s = "U"
        End If
        If Range("S30").Value = 2 Then
            s = "T"
        End If
        If Range("S30").Value = 3 Then
            s = "S"
        End If
        If Range("S30").Value = 4 Then
            s = "R"
        End If
        If Range("S30").Value = 5 Then
            s = "Q"
        End If
        If Range("S30").Value = 6 Then
            s = "P"
        End If
        If Range("S30").Value = 7 Then
            s = "O"
        End If
        If Range("S30").Value = 8 Then
            s = "N"
        End If
        If Range("S30").Value = 9 Then
            s = "M"
        End If
        If Range("S30").Value = 10 Then
            s = "L"
        End If
        If Range("S30").Value = 11 Then
            s = "K"
        End If
        If Range("S30").Value = 12 Then
            s = "J"
        End If
        If Range("S30").Value = 13 Then
            s = "I"
        End If
        If Range("S30").Value = 14 Then
            s = "H"
        End If
        If Range("S30").Value = 15 Then
            s = "G"
        End If
        If Range("S30").Value = 16 Then
            s = "F"
        End If
        If Range("S30").Value = 17 Then
            s = "E"
        End If
        
    'Riadky
        If Range("T31").Value = 1 Then
            r = "41"
        End If
        If Range("T31").Value = 2 Then
            r = "42"
        End If
        If Range("T31").Value = 3 Then
            r = "43"
        End If
        If Range("T31").Value = 4 Then
            r = "44"
        End If
        If Range("T31").Value = 5 Then
            r = "45"
        End If
        If Range("T31").Value = 6 Then
            r = "46"
        End If
        If Range("T31").Value = 7 Then
            r = "47"
        End If
        If Range("T31").Value = 8 Then
            r = "48"
        End If
        
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
    
        Range("B29").Copy
        
        Range(s & r).Select
'        Worksheets("AIO_Plan").Paste
        Selection.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    '   Range("T30").Select
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If
'FUNGUJE!!!
End Sub

Sub SuradniceCentrovaniaPD()
'V pravej dolnej štvrtine rastra stola vyznaèí pozíciu centrovacieho èapu podla zadanıch suradníc

    If Range("V31") > 0 And Range("V31") <= 8 And Range("W30") > 0 And Range("W30") <= 16 Then
        
    'Ståpce
        If Range("W30").Value = 1 Then
            s = "U"
        End If
        If Range("W30").Value = 2 Then
            s = "V"
        End If
        If Range("W30").Value = 3 Then
            s = "W"
        End If
        If Range("W30").Value = 4 Then
            s = "X"
        End If
        If Range("W30").Value = 5 Then
            s = "Y"
        End If
        If Range("W30").Value = 6 Then
            s = "Z"
        End If
        If Range("W30").Value = 7 Then
            s = "AA"
        End If
        If Range("W30").Value = 8 Then
            s = "AB"
        End If
        If Range("W30").Value = 9 Then
            s = "AC"
        End If
        If Range("W30").Value = 10 Then
            s = "AD"
        End If
        If Range("W30").Value = 11 Then
            s = "AE"
        End If
        If Range("W30").Value = 12 Then
            s = "AF"
        End If
        If Range("W30").Value = 13 Then
            s = "AG"
        End If
        If Range("W30").Value = 14 Then
            s = "AH"
        End If
        If Range("W30").Value = 15 Then
            s = "AI"
        End If
        If Range("W30").Value = 16 Then
            s = "AJ"
        End If
        If Range("W30").Value = 17 Then
            s = "AK"
        End If
        
    'Riadky
        If Range("V31").Value = 1 Then
            r = "41"
        End If
        If Range("V31").Value = 2 Then
            r = "42"
        End If
        If Range("V31").Value = 3 Then
            r = "43"
        End If
        If Range("V31").Value = 4 Then
            r = "44"
        End If
        If Range("V31").Value = 5 Then
            r = "45"
        End If
        If Range("V31").Value = 6 Then
            r = "46"
        End If
        If Range("V31").Value = 7 Then
            r = "47"
        End If
        If Range("V31").Value = 8 Then
            r = "48"
        End If
        
        Worksheets("AIO_Plan").Unprotect Password:="Lis.0123"
        Range("B29").Copy
        
        Range(s & r).Select
'        Worksheets("AIO_Plan").Paste
        Selection.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    '   Range("V30").Select
        Worksheets("AIO_Plan").Protect Password:="Lis.0123"
    End If
'FUNGUJE!!!
End Sub

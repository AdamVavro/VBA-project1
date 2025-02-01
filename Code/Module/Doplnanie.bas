Attribute VB_Name = "Doplnanie"

Sub ZavolaOtvoriùFormular()

   Worksheets("AIO_Plan").CommandButton1 = True

End Sub
Sub ZavolaZatvoritFormular()

   UserForm1.CommandButton3 = True

End Sub

Sub DoplniUdajeMsgBox()

    If Range("S1") = "" Then
        MsgBox ("NerobÌm niË")
    Else:    MsgBox ("Sp˙öùam makro ¥¥DoplniUdaje¥¥")
    End If
    
End Sub

Sub ZistiStavFormulara()

    If UserForm1.Visible = True Then
        MsgBox ("Zatv·ram formul·r")
    End If

End Sub
Sub NazovSuboru()

    MenoSuboru = ThisWorkbook.Name
    MsgBox (MenoSuboru)

End Sub
Sub ZmeniTabIndex()

    UserForm1.TextBox1.Text = "Bingo"
    
End Sub
Sub MsgB()

    MsgBox (Range("AN28").Value)
    
End Sub
Sub NajdiVKomentaroch()
            
'            CisloNastroja = Worksheets("AIO_Plan").Range("S1").Value
            CisloNastroja = Workbooks(ThisWorkbook.Name).Worksheets("AIO_Plan").Range("S1").Value
            MsgBox ("HæadanÈ ËÌslo n·stroja: " & CisloNastroja)
            'OK-------------------------------
            Dim PrveStvorcislieNastroja As String
            PrveStvorcislieNastroja = Mid(CisloNastroja, 1, 4)
            
            Set OblastNajdiDVA = Columns(7).Find(PrveStvorcislieNastroja, LookIn:=xlComments)
                       
            If OblastNajdiDVA Is Nothing Then
                I = MsgBox("PrvÈ ötvorËÌslie n·stroja sa nenaölo!", vbOKOnly + vbExclamation, "PrvÈ ötvorËÌslie n·stroja")
            Else
                OblastNajdiDVA.Select
                MsgBox (OblastNajdiDVA.Address)
                MsgBox (OblastNajdiDVA.Row)
                MsgBox (OblastNajdiDVA.Column)
            End If
End Sub

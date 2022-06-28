VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Accesorio 
   Caption         =   "Pérdida de Carga en Accesorio o Localizado"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8865.001
   OleObjectBlob   =   "Accesorio.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Accesorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AcAC_Change()
AcAC.Style = fmStyleDropDownList
End Sub

Private Sub AcAgregar_Click()
    If AcCelda.Text = "" Or AcCN.Text = "" Then
        MsgBox "La celda no es valida", vbCritical, "HF Riego Dice:"
    ElseIf AcPE.Text = "" Or AcVe.Text = "" Then
        MsgBox "Debe realizar primeramente un cálculo", vbCritical, "HF Riego Dice:"
    Else
        
                Celda = AcCelda.Text & AcCN.Text
                Range(Celda).Select
                ActiveCell.Offset(0, 0).Select
                ActiveCell.Value = "La pérdida de carga de " & AcCA.Text & " " & AcAC.Text & " de " & AcDN.Text & " mm con " & AcGA.Text & " lps es de="
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = AcPE.Text
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = "m"
                AcCN.Text = AcCN.Text + 1
            
    End If
End Sub

Private Sub AcBCN_Change()
AcCN.Text = AcBCN.Value
End Sub

Private Sub AcCA_Change()
Me.AcCA.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AcCA.Value)
End Sub
Private Sub AcCalcular_Click()
Dim ki, vel, q, Din, Cant, hf, area As Double
If AcGA.Text = "" Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
Else
    'poner valores en excel
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B36").Value = AcDN.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B20").Value = AcAC.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B22").Value = AcGA.Text
    'inicial variables
    Din = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B37").Value
    q = AcGA.Text
    Cant = AcCA.Text
    
    'inicia calculos
    If AcAC.ListIndex = 0 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C2").Value
    ElseIf AcAC.ListIndex = 1 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C3").Value
    ElseIf AcAC.ListIndex = 2 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C4").Value
    ElseIf AcAC.ListIndex = 3 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C5").Value
    ElseIf AcAC.ListIndex = 4 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C6").Value
    ElseIf AcAC.ListIndex = 5 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C7").Value
    ElseIf AcAC.ListIndex = 6 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C8").Value
    ElseIf AcAC.ListIndex = 7 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C9").Value
    ElseIf AcAC.ListIndex = 8 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C10").Value
    ElseIf AcAC.ListIndex = 9 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C11").Value
    ElseIf AcAC.ListIndex = 10 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C12").Value
    ElseIf AcAC.ListIndex = 11 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C13").Value
    ElseIf AcAC.ListIndex = 12 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C14").Value
    ElseIf AcAC.ListIndex = 13 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C15").Value
    ElseIf AcAC.ListIndex = 14 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C16").Value
    ElseIf AcAC.ListIndex = 15 Then
    ki = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C17").Value
    Else
    
    End If
    area = (WorksheetFunction.pi) * (Din) ^ 2 / 4
    vel = (q / 1000) / area
    hf = Cant * ki * (vel) ^ 2 / (2 * 9.81)

    AcPE.Text = FormatNumber(CDbl(hf), 4)
    AcVe.Text = FormatNumber(CDbl(vel), 4)
    'Workbooks("RegisterU2DF7.xlam").Save
    
                    
End If


End Sub

Private Sub AcCAn_Change()
AcCA.Text = AcCAn.Value
End Sub

Private Sub AcCelda_Change()
Me.AcCelda.Value = Workbooks("RegisterU2DF7.xlam").SoloTexto(Me.AcCelda.Value)
End Sub

Private Sub AcCN_Change()
Me.AcCN.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AcCN.Value)
End Sub

Private Sub AcDN_Change()
AcDN.Style = fmStyleDropDownList
End Sub

Private Sub AcGA_Change()
Me.AcGA.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AcGA.Value)
End Sub
Private Sub SpinButton2_Change()
AcCA.Text = AcCAn.Value
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub PAcAgregar_Click()
If AcGA.Text = "" Or AcPE.Text = "" Or AcVe.Text = "" Then
    MsgBox "No hay suficientes datos para agregar" & vbNewLine & "debe realizar un cálculo", vbCritical, "HF Riego Dice:"
Else
    Dim PTotal As Double
    ListBoxPAc.AddItem Acontador.Caption & "       " & Mid(AcAC.Text, 1, 20) & "                        " & FormatNumber(CDbl(AcDN.Text), 0) & "                             " & FormatNumber(CDbl(AcGA.Text), 2) & "                          " & FormatNumber(CDbl(AcCA.Text), 0) & "                            " & AcPE.Text
    PTotal = PAAc.Text * 1
    PAAc.Text = PTotal + AcPE * 1
    
    'Celda Inicial en la hoja de Excel
    Celda = "A" & Acontador.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAccesorios").Range(Celda).Value = Acontador.Caption
    Celda = "B" & Acontador.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAccesorios").Range(Celda).Value = AcAC.Text
    Celda = "C" & Acontador.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAccesorios").Range(Celda).Value = AcDN.Text
    Celda = "D" & Acontador.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAccesorios").Range(Celda).Value = AcGA.Text
    Celda = "E" & Acontador.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAccesorios").Range(Celda).Value = AcCA.Text
    Celda = "F" & Acontador.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAccesorios").Range(Celda).Value = AcVe.Text
    Celda = "G" & Acontador.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAccesorios").Range(Celda).Value = AcPE.Text

    Acontador.Caption = Acontador.Caption + 1

End If
End Sub

Private Sub PAcExportar_Click()
    If ListBoxPAc.ListCount < 2 Then
        MsgBox "No hay Suficientes Valores para Exportar", vbCritical, "HF Riego Dice:"
    Else
        Workbooks("RegisterU2DF7.xlam").Worksheets("RAccesorios").Range("B6").Value = (PAAc.Text) * 1
    
        '3.- Importamos la hoja de Excel del complemento
                    hojas = ActiveSheet.Name
                    Workbooks("RegisterU2DF7.xlam").Worksheets("RAccesorios").Copy _
                           after:=ActiveWorkbook.Sheets(hojas)
    End If
End Sub

Private Sub UserForm_Initialize()
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B2").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B3").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B4").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B5").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B6").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B7").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B8").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B9").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B10").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B11").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B12").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B13").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B14").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B15").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B16").Value
AcAC.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B17").Value

AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A4").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A5").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A6").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A7").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A8").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
AcDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value

AcAC.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B20").Value
AcDN.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B36").Value
AcGA.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B22").Value

ListBoxPAc.AddItem "#" & "   |             " & "Accesorio" & "              |     " & "Diametro (m)" & "       |       " & "Caudal (lps)" & "       |       " & "Cantidad" & "       |       " & "Pérdida (m)"


End Sub

Private Sub UserForm_Terminate()
Workbooks("RegisterU2DF7.xlam").Worksheets("RAccesorios").Range("A10:G50").Value = ""
    Workbooks("RegisterU2DF7.xlam").Save
End Sub

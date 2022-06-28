VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ETO 
   Caption         =   "Evapotranspiración"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7650
   OleObjectBlob   =   "ETO.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ETO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ETAgregar_Click()
If ETEv.Text = "" Or ETVe.Text = "" Or ETHr.Text = "" Or ETCv.Text = "" Or ETCt.Text = "" Or ETET.Text = "" Then
    MsgBox "HF Riego Dice:" & vbNewLine & "No hay datos Suficientes para agregar a la lista", , "Error"
Else
    ListBoxETO.AddItem ContaETO.Caption & "                         " & FormatNumber(CDbl(ETEv.Text), 3) & "                                       " & FormatNumber(CDbl(ETCt.Text), 3) & "                                       " & FormatNumber(CDbl(ETET.Text), 3)
    'Celda Inicial en la hoja de Excel
    Celda = "A" & ContaETO.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RETo").Range(Celda).Value = ContaETO.Caption
    Celda = "B" & ContaETO.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RETo").Range(Celda).Value = ETEv.Text
    Celda = "C" & ContaETO.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RETo").Range(Celda).Value = ETCt.Text
    Celda = "D" & ContaETO.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RETo").Range(Celda).Value = ETET.Text
    ContaETO.Caption = ContaETO.Caption + 1
End If
End Sub

Private Sub ETCalcular_Click()
If ETEv.Text = "" Or ETVe.Text = "" Or ETHr.Text = "" Or ETCv.Text = "" Then
    MsgBox "HF Riego Dice:" & vbNewLine & "Faltan datos o son irreales: Debe ingresar todos los datos", , "Error"
ElseIf ETEv.Text = 0 Or ETVe.Text = 0 Or ETHr.Text = 0 Or ETCv.Text = 0 Then
    MsgBox "HF Riego Dice:" & vbNewLine & "Faltan datos o son irreales: Debe ingresar todos los datos", , "Error"
ElseIf ETHr.Text >= 100 Then
    MsgBox "HF Riego Dice:" & vbNewLine & "La Humedad Realtiva debe ser menor a 100%", , "Error"
    ETHr.Text = 80
Else
    Dim Ev, U2, HR, d, Kt, Eto As Double
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B64").Value = ETEv.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B65").Value = ETVe.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B66").Value = ETHr.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B67").Value = ETCv.Text
    Ev = ETEv.Text
    U2 = ETVe.Text
    HR = ETHr.Text
    d = ETCv.Text
    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 1 Then
        U2 = U2 * 86400 / 1000
        Kt = 0.475 - 0.00024 * U2 + 0.00516 * HR + 0.00118 * d - 0.000016 * (HR) ^ 2 - 0.101 * (10) ^ -5 * (d) ^ 2 - 0.8 * (10) ^ -8 * (HR) ^ 2 * U2 - 1 * (10) ^ -8 * (HR) ^ 2 * d
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 2 Then
        Kt = 0.108 - 0.0286 * U2 + 0.0422 * WorksheetFunction.Ln(d) + 0.1434 * WorksheetFunction.Ln(HR) - 0.000631 * (WorksheetFunction.Ln(d)) ^ 2 * WorksheetFunction.Ln(HR)
    Else
        Kt = 0.61 + 0.00341 * HR - 0.000162 * U2 * HR - 0.00000959 * U2 * d + 0.00327 * U2 * WorksheetFunction.Ln(d) - 0.00289 * U2 * WorksheetFunction.Ln(86.4 * d) - 0.0106 * WorksheetFunction.Ln(86.4 * U2) * WorksheetFunction.Ln(d) + 0.00063 * (WorksheetFunction.Ln(d)) ^ 2 * WorksheetFunction.Ln(86.4 * U2)
    End If
    Eto = Kt * Ev
    ETCt.Text = FormatNumber(CDbl(Kt), 3)
    ETET.Text = FormatNumber(CDbl(Eto), 3)
        'Workbooks("RegisterU2DF7.xlam").Save
End If
End Sub

Private Sub ETCv_Change()
Me.ETCv.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.ETCv.Value)
End Sub

Private Sub ETEv_Change()
Me.ETEv.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.ETEv.Value)
End Sub

Private Sub ETExportar_Click()
    If ListBoxETO.ListCount < 2 Then
        MsgBox "HF Riego Dice:" & vbNewLine & "No hay Suficientes Valores para Exportar a Excel", , "Error"
    Else
        If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 1 Then
            Workbooks("RegisterU2DF7.xlam").Worksheets("RETo").Range("B4").Value = "Cuenca y Jensen(1989) - Evaporimetro rodeado de Cobertura vegetal"
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 2 Then
            Workbooks("RegisterU2DF7.xlam").Worksheets("RETo").Range("B4").Value = "Allen et al., (1998) - Evaporimetro rodeado de Cobertura vegetal"
    Else
            Workbooks("RegisterU2DF7.xlam").Worksheets("RETo").Range("B4").Value = "Allen et al., (1998) - Evaporimetro rodeado suelo desnudo"
    End If
        '3.- Importamos la hoja de Excel del complemento
                   hojas = ActiveSheet.Name
                    Workbooks("RegisterU2DF7.xlam").Worksheets("RETo").Copy _
                           after:=ActiveWorkbook.Sheets(hojas)
                          MsgBox "Se realizo con exito"
    End If
End Sub

Private Sub ETHr_Change()
Me.ETHr.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.ETHr.Value)

End Sub

Private Sub ETVe_Change()
Me.ETVe.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.ETVe.Value)
End Sub
Private Sub UserForm_Initialize()
ETEv.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B64").Value
ETVe.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B65").Value
ETHr.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B66").Value
ETCv.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B67").Value
ListBoxETO.AddItem "Núm" & "   |      " & "Evaporación (mm)" & "      |      " & "Coeficiente de Tanque" & "       |        " & "Evapotranspiración (mm)"
End Sub

Private Sub UserForm_Terminate()
Workbooks("RegisterU2DF7.xlam").Worksheets("RETo").Range("A10:D80").Value = ""
    Workbooks("RegisterU2DF7.xlam").Save
End Sub

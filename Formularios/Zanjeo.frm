VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Zanjeo 
   Caption         =   "Cálculo de Volumenes de Excavación y Relleno"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7575
   OleObjectBlob   =   "Zanjeo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Zanjeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExportarP_Click()
    If ZaLV.ListCount < 3 Then
        MsgBox " No hay suficientes valores para exportar ", vbCritical, "HF Riego Dice:"
    Else
   
        '3.- Importamos la hoja de Excel del complemento
                    hojas = ActiveSheet.Name
                    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Copy _
                           after:=ActiveWorkbook.Sheets(hojas)
                          MsgBox "Se realizo con exito"
    End If
End Sub

Private Sub PerAgregar2_Click()

If ZaAEx.Text = "" Or ZaAPl.Text = "" Or ZaARC.Text = "" Or ZaARV.Text = "" Then
    MsgBox "Primero, debe realizar un cálculo", vbCritical, "HF Riego Dice:"
ElseIf ZaVEx.Text = "" Or ZaVPl.Text = "" Or ZaVRC.Text = "" Or ZaVRV.Text = "" Then
    MsgBox "Primero, debe realizar un cálculo", vbCritical, "HF Riego Dice:"
Else
    ZaLV.AddItem " " & FormatNumber(CDbl(ContaZ.Caption), 0) & "          " & FormatNumber(CDbl(ZaLT.Text), 1) & "               " & FormatNumber(CDbl(ZaDT.Text), 0) & "                   " & FormatNumber(CDbl(ZaVEx.Text), 2) & "             " & FormatNumber(CDbl(ZaVPl.Text), 2) & "                " & FormatNumber(CDbl(ZaVRC.Text), 2) & "                 " & FormatNumber(CDbl(ZaVRV.Text), 2)

    
    'Celda Inicial en la hoja de Excel
    Celda = "A" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ContaZ.Caption
    Celda = "B" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ZaLT.Text
    Celda = "C" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ZaDT.Text
    Celda = "D" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B53").Value
    Celda = "E" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B54").Value
    Celda = "F" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B55").Value
    Celda = "G" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B56").Value
    'INRESAMOS LAS AREA A EXCEL
    Celda = "H" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ZaAEx.Text
    Celda = "I" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ZaAPl.Text
    Celda = "J" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ZaARC.Text
    Celda = "K" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ZaARV.Text
    'INRESAMOS LOS VOLUMENES A EXCEL
    Celda = "L" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ZaVEx.Text
    Celda = "M" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ZaVPl.Text
    Celda = "N" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ZaVRC.Text
    Celda = "O" & ContaZ.Caption + 10
    Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range(Celda).Value = ZaVRV.Text
    ContaZ.Caption = ContaZ.Caption + 1

End If
End Sub

Private Sub PerCalcular_Click()
Dim AnchoZ, AltoZ, PlantillaZ, EspesorZ, ExcaZ, PlanZ, ReComp, ReVolt, DiametroE, ZaLo As Double
If ZaLT.Text = "" Or ZaDT.Text = "" Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
ElseIf ZaLT.Text = 0 Or ZaDT.Text = 0 Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
Else
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B52").Value = ZaDT.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B51").Value = ZaLT.Text
    AnchoZ = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B53").Value / 100
    AltoZ = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B54").Value / 100
    EspesorZ = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B55").Value / 100
    PlantillaZ = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B56").Value / 100
    DiametroE = ZaDT.Text * 1 / 1000
    ExcaZ = AnchoZ * AltoZ
    PlanZ = AnchoZ * PlantillaZ
    ReComp = ((DiametroE + EspesorZ) * AnchoZ) - (WorksheetFunction.pi * (DiametroE) ^ 2 / 4)
    ReVolt = AnchoZ * (AltoZ - EspesorZ - PlantillaZ - DiametroE)
    ZaL = ZaLT.Text * 1
    ZaAEx.Text = FormatNumber(CDbl(ExcaZ), 3)
    ZaAPl.Text = FormatNumber(CDbl(PlanZ), 3)
    ZaARC.Text = FormatNumber(CDbl(ReComp), 3)
    ZaARV.Text = FormatNumber(CDbl(ReVolt), 3)
    ZaVEx.Text = FormatNumber(CDbl(ExcaZ * ZaL), 3)
    ZaVPl.Text = FormatNumber(CDbl(PlanZ * ZaL), 3)
    ZaVRC.Text = FormatNumber(CDbl(ReComp * ZaL), 3)
    ZaVRV.Text = FormatNumber(CDbl(ReVolt * ZaL), 3)
End If
    

End Sub
Private Sub UserForm_Initialize()
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
ZaDT.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value

ZaLT.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B51").Value
ZaDT.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B52").Value
ZaLV.AddItem "                                                                                       " & "VOLUMENES EN M3"
ZaLV.AddItem "#" & "   |   " & "L Tub.(m)" & "    |    " & "Diám(mm)" & "    |    " & "Excavación" & "        " & "Plantilla" & "        " & "R. Compactado" & "         " & "R. Volteo"

End Sub

Private Sub UserForm_Terminate()

Workbooks("RegisterU2DF7.xlam").Worksheets("RZanjeo").Range("A11:O200").Value = ""
    Workbooks("RegisterU2DF7.xlam").Save


End Sub

Private Sub ZaDT_Change()
ZaDT.Style = fmStyleDropDownList
End Sub

Private Sub ZaLT_Change()
Me.ZaLT.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.ZaLT.Value)
End Sub

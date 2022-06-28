VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PerdidaX 
   Caption         =   "Pérdida de Carga por Fricción en tuberías ciegas o simples"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9030.001
   OleObjectBlob   =   "PerdidaX.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PerdidaX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Ajustes.Show
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub ExportarP_Click()
    If ListBoxPerdida.ListCount < 2 Then
        MsgBox " No hay suficientes valores para exportar ", vbCritical, "HF Riego Dice:"
    Else
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Range("B4").Value = (PerLA.Text)
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Range("B5").Value = PerPA.Text
    
        '3.- Importamos la hoja de Excel del complemento
                    hojas = ActiveSheet.Name
                    Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Copy _
                           after:=ActiveWorkbook.Sheets(hojas)
                          MsgBox "Se realizo con exito"
    End If
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label15_Click()

End Sub

Private Sub PerAgregar_Click()
Dim Ltotal, PTotal As Double
If PerGA.Text = "" Or PerLO.Text = "" Or PerPE.Text = "" Or PerVE.Text = "" Or PerOK.Text = "" Then
    MsgBox "Primero, debe realizar un cálculo", vbCritical, "HF Riego Dice:"
Else
    ListBoxPerdida.AddItem "    " & FormatNumber(CDbl(PerGA.Text), 2) & "                 " & FormatNumber(CDbl(PerDI.Text), 0) & "                         " & FormatNumber(CDbl(PerLO.Text), 2) & "                        " & FormatNumber(CDbl(PerPE.Text), 3) & "                       " & PerOK.Text
    Ltotal = PerLA.Text
    PTotal = PerPA.Text
    PerLA.Text = (PerLO.Text) * 1 + Ltotal
    PerPA.Text = PTotal + PerPE.Text
    
    'Celda Inicial en la hoja de Excel
    Celda = "A" & ContaP.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Range(Celda).Value = ContaP.Caption
    Celda = "B" & ContaP.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Range(Celda).Value = PerGA.Text
    Celda = "C" & ContaP.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Range(Celda).Value = PerDI.Text
    Celda = "D" & ContaP.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Range(Celda).Value = PerLO.Text
    Celda = "E" & ContaP.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Range(Celda).Value = PerPE.Text
    Celda = "F" & ContaP.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Range(Celda).Value = PerVE.Text
    Celda = "G" & ContaP.Caption + 9
    Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Range(Celda).Value = PerOK.Text

    ContaP.Caption = ContaP.Caption + 1

End If

End Sub

Private Sub PerCalcular_Click()
Dim Coefp, DIp, Qp, Longp, Hfp, velp, area As Double
Dim VMax, VMin, perm, perl, relacion As Double
Dim Rey, fdw As Double

If PerGA.Text = "" Or PerDI.Text = "" Or PerLO.Text = "" Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
ElseIf PerGA.Text <= 0 Or PerDI.Text <= 0 Or PerLO.Text <= 0 Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
Else

    Coefp = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E1").Value
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B40").Value = PerGA.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B41").Value = PerDI.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B42").Value = PerLO.Text
    perm = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C47").Value
    perl = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E47").Value
    VMin = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C48").Value
    VMax = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E48").Value
    
    DIp = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B43").Value
    Qp = PerGA.Text / 1000
    Longp = PerLO.Text

    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
        Hfp = 10.674 * 1 / (Coefp) ^ (1.852) * (Qp) ^ (1.852) / (DIp) ^ (4.871) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
        Hfp = 10.294 * (Coefp) ^ (2) * (Qp) ^ (2) / (DIp) ^ (16 / 3) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
        Hfp = 0.004098 * Coefp * (Qp) ^ (1.9) / (DIp) ^ (4.9) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 4 Then
        Rey = Workbooks("RegisterU2DF7.xlam").NReynoldsP(Qp * 1000, DIp * 1000) * 1
        If Rey <= 2000 Then
            fdw = 64 / Rey
        Else
        
                If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E2").Value = 0 Then
                    fdw = Workbooks("RegisterU2DF7.xlam").CoeFriccionDWP(Rey, Coefp, DIp * 1000) * 1
                Else
                    fdw = Workbooks("RegisterU2DF7.xlam").CoeFriccionSJ(Rey, Coefp, DIp * 1000) * 1
                End If
        End If
        
        Hfp = 0.0827 * fdw * (Qp) ^ (2) / (DIp) ^ (5) * Longp
    End If

    area = (WorksheetFunction.pi) * (DIp) ^ 2 / 4
    velp = (Qp) / area
    
    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B46").Value = 1 Then
        relacion = Hfp / (Longp / perl)
        If relacion > perm Then
            PerOK.Text = "Aumenta Diámetro"
        ElseIf velp < VMin Then
            PerOK.Text = "Disminuye Diámetro"
        Else
        PerOK.Text = "Ok. " & "Pierdes " & FormatNumber(CDbl(relacion), 2) & " m en " & FormatNumber(CDbl(perl), 1) & " m"
        End If
    Else
        If velp < VMax And velp > VMin Then
            PerOK.Text = "Ok"
        ElseIf velp < VMin Then
            PerOK.Text = "Disminuye Diámetro"
        ElseIf velp > VMax Then
            PerOK.Text = "Aumenta Diámetro"
        End If
    End If


    PerPE.Text = FormatNumber(CDbl(Hfp), 6)
    PerVE.Text = FormatNumber(CDbl(velp), 7)
    'Workbooks("RegisterU2DF7.xlam").Save
End If
End Sub

Private Sub PerDI_Change()
PerDI.Style = fmStyleDropDownList

End Sub

Private Sub PerGA_Change()
Me.PerGA.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.PerGA.Value)
End Sub
Private Sub PerLO_Change()
Me.PerLO.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.PerLO.Value)
End Sub

Private Sub PerPE_Change()

End Sub

Private Sub UserForm_Initialize()

PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A4").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A5").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A6").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A7").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A8").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
PerDI.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value

PerGA.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B40").Value
PerDI.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B41").Value
PerLO.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B42").Value
If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B46").Value = 1 Then
Metodo.Caption = "Pérdida de Carga Unitaria"
Else
Metodo.Caption = "Velocidad Permisible"
End If

ListBoxPerdida.AddItem "Gasto (lps)" & "   |   " & "Diametro (mm)" & "    |     " & "Longitud (m)" & "       |       " & "Pérdida (m)"
End Sub

Private Sub UserForm_Terminate()
Workbooks("RegisterU2DF7.xlam").Worksheets("RTubCiega").Range("A10:G80").Value = ""
    Workbooks("RegisterU2DF7.xlam").Save
End Sub

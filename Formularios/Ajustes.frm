VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ajustes 
   Caption         =   "Ajustes HF Riego"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10575
   OleObjectBlob   =   "Ajustes.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Ajustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aceptar_Click()
'Agregamos los valores a la hoja de excel Metodo
If Di12.Text = "" Or Di16.Text = "" Or Di17.Text = "" Or Di19.Text = "" Or Di20.Text = "" Or Di25.Text = "" Or Di32.Text = "" Or Di38.Text = "" Or Di50.Text = "" Or Di60.Text = "" Or Di75.Text = "" Or Di100.Text = "" Or Di160.Text = "" Or Di200.Text = "" Or Di250.Text = "" Or Di315.Text = "" Then
    MsgBox "Faltan diámetros interno por agregar", vbCritical, "HF Riego Dice:"
ElseIf NPVC.Text = "" Or NPOLI.Text = "" Or NALUM.Text = "" Or NASB.Text = "" Or NACE.Text = "" Or HWPVC.Text = "" Or HWPOLI.Text = "" Or HWALUM.Text = "" Or HWASB.Text = "" Or HWACE.Text = "" Or SPVC.Text = "" Or SPOLI.Text = "" Or SALUM.Text = "" Or SASB.Text = "" Or SACE.Text = "" Or DWPVC.Text = "" Or DWPOLI.Text = "" Or DWALUM.Text = "" Or DWASB.Text = "" Or DWACE.Text = "" Then
    MsgBox "Faltan coeficientes de fricción por agregar", vbCritical, "HF Riego Dice:"
ElseIf PF.Text < 0 Or PF.Text > 100 Then
    MsgBox "Porcentaje debe ser entre 0 y 100" & vbNewLine & "Porcentaje Fijo", vbCritical, "HF Riego Dice:"
ElseIf KxVP.Text = "" Or KxAG.Text = "" Or KxAB.Text = "" Or KxRG.Text = "" Or KxRB.Text = "" Or KxC9.Text = "" Or KxC4.Text = "" Or KxCR9.Text = "" Or KxCR4.Text = "" Or KxCR2.Text = "" Or KxTE.Text = "" Or KxVC.Text = "" Or KxVG.Text = "" Or KxVA.Text = "" Or KxCG.Text = "" Or KxVM.Text = "" Then
    MsgBox "Faltan coeficientes de accesorios por agregar", vbCritical, "HF Riego Dice:"
ElseIf AjMD.Text = "" Or DtPerder.Text = "" Or DtCada.Text = "" Or DtVM.Text = "" Or DtVMx.Text = "" Or EfiG.Text = "" Or EfiM.Text = "" Then
    MsgBox "Faltan algun valor por agregar en la ventana de diseño", vbCritical, "HF Riego Dice:"
ElseIf AZAB25.Text = "" Or AZAB32.Text = "" Or AZAB38.Text = "" Or AZAB50.Text = "" Or AZAB75.Text = "" Or AZAB100.Text = "" Or AZAB160.Text = "" Or AZAB200.Text = "" Or AZAB250.Text = "" Or AZAB315.Text = "" Or AZAB60.Text = "" Then
    MsgBox "Faltan algun valor por agregar de ancho de Zanja", vbCritical, "HF Riego Dice:"
ElseIf AZAH25.Text = "" Or AZAH32.Text = "" Or AZAH38.Text = "" Or AZAH50.Text = "" Or AZAH75.Text = "" Or AZAH100.Text = "" Or AZAH160.Text = "" Or AZAH200.Text = "" Or AZAH250.Text = "" Or AZAH315.Text = "" Or AZAH60.Text = "" Then
    MsgBox "Faltan algun valor por agregar de profundidad de Zanja", vbCritical, "HF Riego Dice:"
ElseIf AZAE25.Text = "" Or AZAE32.Text = "" Or AZAE38.Text = "" Or AZAE50.Text = "" Or AZAE75.Text = "" Or AZAE100.Text = "" Or AZAE160.Text = "" Or AZAE200.Text = "" Or AZAE250.Text = "" Or AZAE315.Text = "" Or AZAE60.Text = "" Then
    MsgBox "Faltan algun valor por agregar de espesor de Zanja", vbCritical, "HF Riego Dice:"
ElseIf AZAL25.Text = "" Or AZAL32.Text = "" Or AZAL38.Text = "" Or AZAL50.Text = "" Or AZAL75.Text = "" Or AZAL100.Text = "" Or AZAL160.Text = "" Or AZAL200.Text = "" Or AZAB250.Text = "" Or AZAL315.Text = "" Or AZAL60.Text = "" Then
    MsgBox "Faltan algun valor por agregar de plantilla de Zanja", vbCritical, "HF Riego Dice:"
ElseIf ETOKh.Text = "" Or ETOeh.Text = "" Or ETOKt.Text = "" Then
    MsgBox "Faltan algun valor por agregar en Evapotranspiración", vbCritical, "HF Riego Dice:"
ElseIf ETOKh.Text = 0 Or ETOeh.Text = 0 Or ETOKt.Text = 0 Then
    MsgBox "Faltan algun valor por agregar en Evapotranspiración", vbCritical, "HF Riego Dice:"
ElseIf OptionButton1.Value = True And TextKrs.Text = "" Then
    MsgBox "Faltan algun valor por agregar en Evapotranspiración o datos no validos", vbCritical, "HF Riego Dice:"
ElseIf OptionButton1.Value = True And TextKrs.Text <= 0 Then
    MsgBox "Faltan algun valor por agregar en Evapotranspiración o datos no validos", vbCritical, "HF Riego Dice:"
Else
    '1. ponemos las variables
    If Formula.Text = "Hazen-Williams" Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1
    ElseIf Formula.Text = "Manning" Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2
    ElseIf Formula.Text = "Scobey" Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3
    ElseIf Formula.Text = "Darcy-Weisbach" Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 4
    End If
    '2. Ponemos los materiales
    If Material.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A22").Value Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B2").Value = 1
    ElseIf Material.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A23").Value Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B2").Value = 2
    ElseIf Material.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A24").Value Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B2").Value = 3
    ElseIf Material.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A25").Value Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B2").Value = 4
    ElseIf Material.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A26").Value Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B2").Value = 5
    End If
    
    'nos encargamos ahora de los diametros internos
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B4").Value = Di12.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B5").Value = Di16.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B6").Value = Di17.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B7").Value = Di19.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B8").Value = Di20.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B9").Value = Di25.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B10").Value = Di32.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B11").Value = Di38.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B12").Value = Di50.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B13").Value = Di60.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B14").Value = Di75.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B15").Value = Di100.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B16").Value = Di160.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B17").Value = Di200.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B18").Value = Di250.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B19").Value = Di315.Text
    
    ' Nos encargamos de los coefientes de Manning
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B22").Value = NPVC.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B23").Value = NPOLI.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B24").Value = NALUM.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B25").Value = NASB.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B26").Value = NACE.Text
    
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C22").Value = HWPVC.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C23").Value = HWPOLI.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C24").Value = HWALUM.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C25").Value = HWASB.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C26").Value = HWACE.Text
    
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("D22").Value = SPVC.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("D23").Value = SPOLI.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("D24").Value = SALUM.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("D25").Value = SASB.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("D26").Value = SACE.Text
    
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E22").Value = DWPVC.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E23").Value = DWPOLI.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E24").Value = DWALUM.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E25").Value = DWASB.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E26").Value = DWACE.Text
    
    'AQUI ES DEL MENU DE EVAPOTRANSPIRACION

    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = MetodoPE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B6").Value = AFE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("E6").Value = BFE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B7").Value = CFE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("E7").Value = DFE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("H6").Value = FECONDI.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B9").Value = PF.Text
    FECONDI2.Text = FECONDI.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B18").Value = ETOKh.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B19").Value = ETOeh.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B20").Value = ETOKt.Text
    
    If OptionButton1.Value = True Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B23").Value = 1
    Else
        Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B23").Value = 2
    End If
    
    If OptionCW.Value = True Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E2").Value = 0
    Else
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E2").Value = 1
    End If
    
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B26").Value = TRocio.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B24").Value = TextKrs.Text
    
    
    ' Nos encargamos de los Kx de Accesorios
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C2").Value = KxVP.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C3").Value = KxAG.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C4").Value = KxAB.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C5").Value = KxRG.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C6").Value = KxRB.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C7").Value = KxC9.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C8").Value = KxC4.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C9").Value = KxCR9.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C10").Value = KxCR4.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C11").Value = KxCR2.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C12").Value = KxTE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C13").Value = KxVC.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C14").Value = KxVG.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C15").Value = KxVA.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C16").Value = KxCG.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C17").Value = KxVM.Text
    
    'Agregamos los datos de eficiencia del sistema de riego
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B14").Value = EfiG.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B15").Value = EfiM.Text
    'Agregamos los datos del diseño
    If AjMD.Text = "Pérdida de Carga Unitaria" Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B46").Value = 1
    Else
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B46").Value = 2
    End If
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C47").Value = DtPerder.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E47").Value = DtCada.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C48").Value = DtVM.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E48").Value = DtVMx.Text
    
    If MeCT1.Value = True Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 1
    ElseIf MeCT2.Value = True Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 2
    Else
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 3
    End If

    'guardamos el archivo
    Workbooks("RegisterU2DF7.xlam").Save
    Unload Me
End If

End Sub

Private Sub AFE_Change()
Me.AFE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AFE.Value)
End Sub

Private Sub AjMD_Change()
AjMD.Style = fmStyleDropDownList
If AjMD.Text = "Pérdida de Carga Unitaria" Then
    DtPerder.Enabled = True
    DtCada.Enabled = True
    DtVM.Enabled = False
    DtVMx.Enabled = False
Else
    DtPerder.Enabled = False
    DtCada.Enabled = False
    DtVM.Enabled = True
    DtVMx.Enabled = True

End If
End Sub

Private Sub AZAB100_Change()
Me.AZAB100.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB100.Value)
End Sub

Private Sub AZAB160_Change()
Me.AZAB160.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB160.Value)
End Sub

Private Sub AZAB200_Change()
Me.AZAB200.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB200.Value)
End Sub

Private Sub AZAB25_Change()
Me.AZAB25.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB25.Value)
End Sub

Private Sub AZAB250_Change()
Me.AZAB250.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB250.Value)
End Sub

Private Sub AZAB315_Change()
Me.AZAB315.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB315.Value)
End Sub

Private Sub AZAB32_Change()
Me.AZAB32.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB32.Value)
End Sub

Private Sub AZAB38_Change()
Me.AZAB38.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB38.Value)
End Sub

Private Sub AZAB50_Change()
Me.AZAB50.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB50.Value)
End Sub

Private Sub AZAB60_Change()
Me.AZAB60.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB60.Value)
End Sub

Private Sub AZAB75_Change()
Me.AZAB75.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.AZAB75.Value)
End Sub

Private Sub BFE_Change()

End Sub

Private Sub CFE_Change()
Me.CFE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.CFE.Value)
End Sub

Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
EZanja.Show
End Sub

Private Sub CommandButton3_Click()
RenDiametro.Show
End Sub

Private Sub CommandButton4_Click()
ReAcce.Show
End Sub

Private Sub CommandButton5_Click()
EToHS.Show
End Sub

Private Sub CommandButton6_Click()
Rs.Show
End Sub

Private Sub DFE_Change()
Me.DFE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DFE.Value)
End Sub

Private Sub Di100_Change()
Me.Di100.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di100.Value)
End Sub

Private Sub Di12_Change()
Me.Di12.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di12.Value)
End Sub

Private Sub Di16_Change()
Me.Di16.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di16.Value)
End Sub

Private Sub Di160_Change()
Me.Di160.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di160.Value)
End Sub

Private Sub Di17_Change()
Me.Di17.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di17.Value)
End Sub

Private Sub Di19_Change()
Me.Di19.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di19.Value)
End Sub

Private Sub Di20_Change()
Me.Di20.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di20.Value)
End Sub

Private Sub Di200_Change()
Me.Di200.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di200.Value)
End Sub

Private Sub Di25_Change()
Me.Di25.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di25.Value)
End Sub

Private Sub Di250_Change()
Me.Di250.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di250.Value)
End Sub

Private Sub Di315_Change()
Me.Di315.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di315.Value)
End Sub

Private Sub Di32_Change()
Me.Di32.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di32.Value)
End Sub

Private Sub Di38_Change()
Me.Di38.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di38.Value)
End Sub

Private Sub Di50_Change()
Me.Di50.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di50.Value)
End Sub

Private Sub Di60_Change()
Me.Di60.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di60.Value)
End Sub

Private Sub Di75_Change()
Me.Di75.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Di75.Value)
End Sub


Private Sub DtCada_Change()
Me.DtCada.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DtCada.Value)
End Sub

Private Sub DtPerder_Change()
Me.DtPerder.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DtPerder.Value)
End Sub

Private Sub DtVM_Change()
Me.DtVM.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DtVM.Value)
End Sub

Private Sub DtVMx_Change()
Me.DtVMx.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DtVMx.Value)
End Sub

Private Sub EfiG_Change()
Me.EfiG.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.EfiG.Value)
End Sub

Private Sub EfiM_Change()
Me.EfiM.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.EfiM.Value)
End Sub

Private Sub Estados_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0

End Sub

Private Sub ETOeh_Change()
Me.ETOeh.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.ETOeh.Value)
End Sub

Private Sub ETOKh_Change()
Me.ETOKh.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.ETOKh.Value)
End Sub

Private Sub ETOKt_Change()
Me.ETOKt.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.ETOKt.Value)
End Sub

Private Sub FECONDI_Change()
Me.FECONDI.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.FECONDI.Value)
End Sub

Private Sub Formula_Change()
Formula.Style = fmStyleDropDownList
End Sub

Private Sub Formula_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub Frame15_Click()

End Sub

Private Sub HWACE_Change()
Me.HWACE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.HWACE.Value)
End Sub

Private Sub HWALUM_Change()
Me.HWALUM.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.HWALUM.Value)
End Sub

Private Sub HWASB_Change()
Me.HWASB.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.HWASB.Value)
End Sub

Private Sub HWPOLI_Change()
Me.HWPOLI.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.HWPOLI.Value)
End Sub

Private Sub HWPVC_Change()
Me.HWPVC.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.HWPVC.Value)
End Sub
Private Sub KxAB_Change()
Me.KxAB.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxAB.Value)
End Sub

Private Sub KxAG_Change()
Me.KxAG.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxAG.Value)
End Sub

Private Sub KxC4_Change()
Me.KxC4.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxC4.Value)
End Sub

Private Sub KxC9_Change()
Me.KxC9.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxC9.Value)
End Sub

Private Sub KxCG_Change()
Me.KxCG.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxCG.Value)
End Sub

Private Sub KxCR2_Change()
Me.KxCR2.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxCR2.Value)
End Sub

Private Sub KxCR4_Change()
Me.KxCR4.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxCR4.Value)
End Sub

Private Sub KxCR9_Change()
Me.KxCR9.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxCR9.Value)
End Sub

Private Sub KxRB_Change()
Me.KxRB.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxRB.Value)
End Sub

Private Sub KxRG_Change()
Me.KxRG.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxRG.Value)
End Sub

Private Sub KxTE_Change()
Me.KxTE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxTE.Value)
End Sub

Private Sub KxVA_Change()
Me.KxVA.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxVA.Value)
End Sub

Private Sub KxVG_Change()
Me.KxVG.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxVG.Value)
End Sub

Private Sub KxVM_Change()
Me.KxVM.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxVM.Value)
End Sub

Private Sub KxVP_Change()
Me.KxVP.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.KxVP.Value)
End Sub

Private Sub Material_Change()
Material.Style = fmStyleDropDownList
End Sub

Private Sub Material_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub
Private Sub MetodoPE_Change()
MetodoPE.Style = fmStyleDropDownList

If MetodoPE.Text = "Formula empirica" Then
AFE.Enabled = True
BFE.Enabled = True
CFE.Enabled = True
DFE.Enabled = True
FECONDI.Enabled = True
ElseIf MetodoPE.Text = "Porcentaje fijo" Then
PF.Enabled = True
Else
AFE.Enabled = False
BFE.Enabled = False
CFE.Enabled = False
DFE.Enabled = False
FECONDI.Enabled = False
PF.Enabled = False
End If


End Sub
Private Sub MetodoPE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub NACE_Change()
Me.NACE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.NACE.Value)
End Sub

Private Sub NALUM_Change()
Me.NALUM.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.NALUM.Value)
End Sub

Private Sub NASB_Change()
Me.NASB.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.NASB.Value)
End Sub

Private Sub NPOLI_Change()
Me.NPOLI.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.NPOLI.Value)
End Sub

Private Sub NPVC_Change()
Me.NPVC.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.NPVC.Value)
End Sub

Private Sub PF_Change()
Me.PF.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.PF.Value)
End Sub

Private Sub SACE_Change()
Me.SACE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.SACE.Value)
End Sub

Private Sub SALUM_Change()
Me.SALUM.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.SALUM.Value)
End Sub

Private Sub SASB_Change()
Me.SASB.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.SASB.Value)
End Sub

Private Sub SEfiG_Change()
EfiG.Text = SEfiG.Value
End Sub

Private Sub SEfiM_Change()
EfiM.Text = SEfiM.Value
End Sub

Private Sub SPOLI_Change()
Me.SPOLI.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.SPOLI.Value)
End Sub

Private Sub SPVC_Change()
Me.SPVC.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.SPVC.Value)
End Sub

Private Sub TextBox13_Change()

End Sub

Private Sub TextKrs_Change()
Me.TextKrs.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.TextKrs.Value)
End Sub

Private Sub TRocio_Change()
Me.TRocio.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimalNega(Me.TRocio.Value)

End Sub

Private Sub UserForm_Initialize()
' CRea las de las formulas
Formula.AddItem "Hazen-Williams"
Formula.AddItem "Manning"
Formula.AddItem "Scobey"
Formula.AddItem "Darcy-Weisbach"

'crea las opciones del material
Material.AddItem "PVC"
Material.AddItem "Polietileno"
Material.AddItem "Aluminio"
Material.AddItem "Asbesto-Cemento"
Material.AddItem "Acero Galvanizado"

'crea las opciones de metodo precipitacion efectiva
MetodoPE.AddItem "Porcentaje fijo"
MetodoPE.AddItem "Precipitacion Confiable"
MetodoPE.AddItem "Formula empirica"
MetodoPE.AddItem "USDA"

'crea las opciones de Estados
'Estados.AddItem "Aguascalientes"
'Estados.AddItem "Baja California"
'Estados.AddItem "Baja California Sur"
'Estados.AddItem "Campeche"
'Estados.AddItem "Chiapas"
'Estados.AddItem "Chihuahua"
'Estados.AddItem "Distrito Federal"
'Estados.AddItem "Coahuila de Zaragoza"
'Estados.AddItem "Colima"
'Estados.AddItem "Durango"
'Estados.AddItem "Guanajuato"
'Estados.AddItem "Guerrero"
'Estados.AddItem "Hidalgo"
'Estados.AddItem "Jalisco"
'Estados.AddItem "México"
'Estados.AddItem "Michoacán de Ocampo"
'Estados.AddItem "Morelos"
'Estados.AddItem "Nayarit"
'Estados.AddItem "Nuevo León"
'Estados.AddItem "Oaxaca"
'Estados.AddItem "Puebla"
'Estados.AddItem "Querétaro de Arteaga"
'Estados.AddItem "Quintana Roo"
'Estados.AddItem "San Luis Potosí"
'Estados.AddItem "Sinaloa"
'Estados.AddItem "Sonora"
'Estados.AddItem "Tabasco"
'Estados.AddItem "Tamaulipas"
'Estados.AddItem "Tlaxcala"
'Estados.AddItem "Veracruz de Ignacio de la Llave"
'Estados.AddItem "Yucatán"
'Estados.AddItem "Zacatecas"

'Agregamos los tipos de metodos de diseño hidraulico
AjMD.AddItem "Pérdida de Carga Unitaria"
AjMD.AddItem "Velocidad Pérmisible"


'AQUI ES DEL MENU DE EVAPOTRANSPIRACION
'Estados.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B1").Value
MetodoPE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value
AFE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B6").Value
BFE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("E6").Value
CFE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B7").Value
DFE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("E7").Value
FECONDI.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("H6").Value
FECONDI2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("H6").Value
PF.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B9").Value

ETOKh.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B18").Value
ETOeh.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B19").Value
ETOKt.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B20").Value

If Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B23").Value = 1 Then
OptionButton1.Value = True
Else
OptionButton2.Value = True
End If
TRocio.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B26").Value
TextKrs.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B24").Value

If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E2").Value = 0 Then
    OptionCW.Value = True
Else
    OptionSJ.Value = True
End If


'Agregamos los datos de Diseño Tuberia
EfiG.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B14").Value
EfiM.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B15").Value

If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B46").Value = 1 Then
    AjMD.Text = "Pérdida de Carga Unitaria"
Else
    AjMD.Text = "Velocidad Pérmisible"
End If
DtPerder.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C47").Value
DtCada.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E47").Value
DtVM.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C48").Value
DtVMx.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E48").Value

'AQUI ES DEL MENU DE HIDRAULICO
'1. ponemos las variables
If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
    Formula.Text = "Hazen-Williams"
ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
    Formula.Text = "Manning"
ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
    Formula.Text = "Scobey"
Else
    Formula.Text = "Darcy-Weisbach"
End If
'2. Ponemos los materiales
If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B2").Value = 1 Then
    Material.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A22").Value
ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B2").Value = 2 Then
    Material.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A23").Value
ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B2").Value = 3 Then
    Material.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A24").Value
ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B2").Value = 4 Then
    Material.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A25").Value
ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B2").Value = 5 Then
    Material.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A26").Value
End If



'nos encargamos ahora de los diametros internos
Di12.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B4").Value
Di16.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B5").Value
Di17.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B6").Value
Di19.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B7").Value
Di20.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B8").Value
Di25.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B9").Value
Di32.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B10").Value
Di38.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B11").Value
Di50.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B12").Value
Di60.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B13").Value
Di75.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B14").Value
Di100.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B15").Value
Di160.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B16").Value
Di200.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B17").Value
Di250.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B18").Value
Di315.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B19").Value
'nos encargamos de los diámetros nominales
Ad1.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A4").Value
Ad2.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A5").Value
Ad3.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A6").Value
Ad4.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A7").Value
Ad5.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A8").Value
Ad6.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
Ad7.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
Ad8.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
Ad9.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
Ad10.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
Ad11.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
Ad12.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
Ad13.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
Ad14.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
Ad15.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
Ad16.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value
'nos encargamos de los diametros de la zanja
ZAd6.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
ZAd7.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
ZAd8.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
ZAd9.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
ZAd10.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
ZAd11.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
ZAd12.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
ZAd13.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
ZAd14.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
ZAd15.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
ZAd16.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value
'nos encargamos de los diámetros nominales del cemento
'Cad1.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
'Cad2.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
'Cad3.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
'Cad4.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
'Cad5.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
'Cad6.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
'Cad7.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
'Cad8.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
'Cad9.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
'Cad10.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
'Cad11.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value

' Nos encargamos de los coefientes de Manning
NPVC.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B22").Value
NPOLI.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B23").Value
NALUM.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B24").Value
NASB.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B25").Value
NACE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B26").Value

HWPVC.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C22").Value
HWPOLI.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C23").Value
HWALUM.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C24").Value
HWASB.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C25").Value
HWACE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C26").Value

SPVC.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("D22").Value
SPOLI.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("D23").Value
SALUM.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("D24").Value
SASB.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("D25").Value
SACE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("D26").Value

DWPVC.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E22").Value
DWPOLI.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E23").Value
DWALUM.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E24").Value
DWASB.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E25").Value
DWACE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E26").Value

' Nos encargamos de los Kx de Accesorios
KxVP.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C2").Value
KxAG.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C3").Value
KxAB.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C4").Value
KxRG.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C5").Value
KxRB.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C6").Value
KxC9.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C7").Value
KxC4.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C8").Value
KxCR9.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C9").Value
KxCR4.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C10").Value
KxCR2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C11").Value
KxTE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C12").Value
KxVC.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C13").Value
KxVG.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C14").Value
KxVA.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C15").Value
KxCG.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C16").Value
KxVM.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("C17").Value

'nos encargamos de los nombre de los accesorios
AcN1.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B2").Value
AcN2.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B3").Value
AcN3.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B4").Value
AcN4.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B5").Value
AcN5.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B6").Value
AcN6.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B7").Value
AcN7.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B8").Value
AcN8.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B9").Value
AcN9.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B10").Value
AcN10.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B11").Value
AcN11.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B12").Value
AcN12.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B13").Value
AcN13.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B14").Value
AcN14.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B15").Value
AcN15.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B16").Value
AcN16.Caption = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B17").Value


'Este es para el menu de Zanja
AZAB25.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E9").Value
AZAB32.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E10").Value
AZAB38.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E11").Value
AZAB50.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E12").Value
AZAB60.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E13").Value
AZAB75.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E14").Value
AZAB100.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E15").Value
AZAB160.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E16").Value
AZAB200.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E17").Value
AZAB250.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E18").Value
AZAB315.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E19").Value

AZAH25.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F9").Value
AZAH32.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F10").Value
AZAH38.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F11").Value
AZAH50.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F12").Value
AZAH60.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F13").Value
AZAH75.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F14").Value
AZAH100.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F15").Value
AZAH160.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F16").Value
AZAH200.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F17").Value
AZAH250.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F18").Value
AZAH315.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F19").Value

AZAE25.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G9").Value
AZAE32.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G10").Value
AZAE38.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G11").Value
AZAE50.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G12").Value
AZAE60.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G13").Value
AZAE75.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G14").Value
AZAE100.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G15").Value
AZAE160.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G16").Value
AZAE200.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G17").Value
AZAE250.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G18").Value
AZAE315.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("G19").Value

AZAL25.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H9").Value
AZAL32.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H10").Value
AZAL38.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H11").Value
AZAL50.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H12").Value
AZAL60.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H13").Value
AZAL75.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H14").Value
AZAL100.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H15").Value
AZAL160.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H16").Value
AZAL200.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H17").Value
AZAL250.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H18").Value
AZAL315.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("H18").Value

If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 1 Then
    MeCT1.Value = True
ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 2 Then
    MeCT2.Value = True
Else
    MeCT3.Value = True
End If
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Workbooks("RegisterU2DF7.xlam").ClicDerecho
End Sub

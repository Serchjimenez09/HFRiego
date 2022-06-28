Attribute VB_Name = "GENERAL2"
'LLAMANDO LAS FUNCIONES DE HIRAULICA
Sub llamarInterno(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=dinterno()"
End Sub

'Callback for AGROLAMINA onAction
Sub llamarLAMINA(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=LaminaHoraria()"
End Sub

'Callback for LLateral onAction
Sub llamarLATERAL(control As IRibbonControl)
Regante.Show
End Sub


'Callback for Ccompuestos onAction
Sub llamarCONFIGURACION(control As IRibbonControl)
Ajustes.Show
End Sub

'Callback for ayuda onAction
Sub llamarAyuda(control As IRibbonControl)
RiegoAyuda.Show
End Sub

'Callback for AcercaDE onAction
Sub llamarAcercaDe(control As IRibbonControl)
ACERCA_DE.Show
End Sub


'Callback for KC onAction
Sub llamarKC(control As IRibbonControl)
KC.Show

End Sub

'Callback for PE onAction
Sub llamarPE(control As IRibbonControl)
PreEfectiva.Show
End Sub

'Callback for RR onAction
Sub llamarRR(control As IRibbonControl)
End Sub
'Callback for QSISTEMAA onAction
Sub llamarQSISTEMA(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=Qtotalreq()"
End Sub

'Callback for QMINIMO onAction
Sub llamarQMINIMO(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=Qminimoxseccion()"
    SendKeys "+{F3}", True
End Sub

'Callback for FSALIDAS onAction
Sub llamarSALIDASM(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=FChristiansen()"
End Sub

'Callback for DAGRONOMICO onAction
Sub llamaAGRONOMICO(control As IRibbonControl)
Agronomico.Show
End Sub
'Callback for FSALIDAS2 onAction
Sub llamarSALIDAS2(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=FJensen()"
End Sub

'Callback for FSALIDAS3 onAction
Sub llamarSALIDASM3(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=FScaloppi()"
End Sub

'Callback for BPerfilESREc onAction
Sub HPrincipal(control As IRibbonControl)
PerdidaX.Show
End Sub

'Callback for haccesorios onAction
Sub llamarACCESORIOS(control As IRibbonControl)
Accesorio.Show
End Sub
Sub llamarTeles(control As IRibbonControl)
Secundaria.Show
End Sub
Sub llamarZANJA(control As IRibbonControl)
Zanjeo.Show
End Sub
Sub llamarBOMBEO(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=PotenciaBomba()"
End Sub
Sub llamarTEXTURA(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=TexturaSuelo()"
End Sub
Sub llamarETO(control As IRibbonControl)
    Eto.Show
End Sub
Sub llamarHFacil(control As IRibbonControl)
End Sub
Sub llamarETOPM(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=EToPM()"
End Sub
Sub llamarETOPMDL(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=PMDatosLimitados()"
End Sub

Sub llamarETO2(control As IRibbonControl) 'Tanque tipoAformula
    ActiveCell.FormulaR1C1 = "=EvapotranspiracionA()"
End Sub
Sub llamarETOHS(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=EToHargreavesSamani()"
End Sub
Sub llamarETOPT(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=EToPriestleTaylor()"
End Sub
Sub llamarJuliano(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=aDiaJulianoo()"
End Sub

Sub llamarVelocidad(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=Windspeed()"
End Sub
Sub llamarRE(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=RadiacionExtraterrestres()"
End Sub
'formulas
Sub llamarLM(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=LongMaxRegante()"
End Sub
Sub llamarPCF(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=Perdida()"
End Sub
Sub llamarVeFlujo(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=VelocidadFlujo()"
End Sub
Sub llamarReynolds(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=NReynolds()"
End Sub
Sub llamarCoeficienteF(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=CoeFriccionDW()"
End Sub
Sub llamarCoeficienteSJ(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=CoeFriccionSJ()"
End Sub
Sub llamarRs(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=Rs_SkyCover()"
End Sub
Sub llamarHR(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=VP_SpecHumid()"
End Sub
Sub llamarME(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=MeanError()"
End Sub
Sub llamarSDE(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=StandarDeviationError()"
End Sub
Sub llamarDW(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=dWilmmott()"
End Sub
Sub llamarRMSE(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=RMSE()"
End Sub
Sub llamarVPA(control As IRibbonControl)
    ActiveCell.FormulaR1C1 = "=ActualVaporRocio()"
End Sub

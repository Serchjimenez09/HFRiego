Attribute VB_Name = "Hidraulicas"
Option Private Module
'ULTIMA MODIFICACIÓN 03 DE FEBRERO DE 2014



' *********************************************************************
' Module  : MExample
' Purpose : Demonstrates the use of CUdfHelper to load and unload
'           function descriptions from a worksheet.
'
' Notes   : Copy these routines to your own workbook
'           and call them from Workbook_Open/Close
' *********************************************************************

' ---------------------------------------------------------------------
' Date        Developer                   Action
' 2006-02-23  Jurgen Volkerink            Created

Option Explicit

Sub LoadFunctionDescriptions()
'Register the functions in the range
    With New CUdfHelper
        .ProcessRange shFuncList.Range("FuncList"), bRegister:=True
    End With
End Sub
Sub UnLoadFunctionDescriptions()
'UnRegister the functions in the range
    With New CUdfHelper
        .ProcessRange shFuncList.Range("FuncList"), bRegister:=False
    End With
End Sub

Private Sub foo()
    Dim v
    v = Application.RegisteredFunctions
    Stop                                     'open Locals Window to view
    'i've noted that when ATP is loaded
    '1st of 'our' functions appears in the top
    'while the rest appear the end... hmm???
End Sub

Function PerdidaX2(gasto, diametro, longitud As Double) As Double
'Determina la perdida por fricción en la tuberia con PVC
Dim Coefp, DIp, Qp, Longp, Hfp As Double
    Qp = gasto / 1000
    Longp = longitud
    DIp = diametro / 1000
    Coefp = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E1").Value
    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
        Hfp = 10.648 * 1 / (Coefp) ^ (1.852) * (Qp) ^ (1.852) / (DIp) ^ (4.871) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
        Hfp = 10.3 * (Coefp) ^ (2) * (Qp) ^ (2) / (DIp) ^ (16 / 3) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
        Hfp = 0.004098 * Coefp * (Qp) ^ (1.9) / (DIp) ^ (4.9) * Longp
    End If

    PerdidaX = 1
    
End Function

Function FChristiansen(NS As Integer) As Double
'Determina el Factor para salidas multiples
Rem REVISAR Y CORREGIR 03 DE
Dim N As Double

    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
    N = 1.852
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
    N = 2
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
    N = 1.9
    Else
    N = 2
    End If
    If NS > 0 Then
        FChristiansen = 1 / (N + 1) + 1 / (2 * NS) + (Sqr(N - 1)) / (6 * (NS) ^ 2)
    Else
        FChristiansen = CVErr(xlErrDiv0)
    End If
End Function
Function FJensen(NS As Integer) As Double
'Determina el Factor para salidas multiples con jensen

Dim N As Double

    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
    N = 1.852
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
    N = 2
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
    N = 1.9
    Else
    N = 2
    End If
    If NS > 0 Then
        FJensen = (2 * NS / (2 * NS - 1)) * (1 / (N + 1) + (Sqr(N - 1)) / (6 * (NS) ^ 2))
    Else
        FJensen = CVErr(xlErrDiv0)
    End If
End Function
Function FScaloppi(NS As Integer, s, So As Double) As Double
'Determina el Factor para salidas multiples con jensen

Dim N, f1, Rs As Double

    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
    N = 1.852
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
    N = 2
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
    N = 1.9
    Else
    N = 2
    End If
    If NS > 0 Then
        f1 = 1 / (N + 1) + 1 / (2 * NS) + (Sqr(N - 1)) / (6 * (NS) ^ 2)
        Rs = So / s
        FScaloppi = (NS * f1 + Rs - 1) / (NS + Rs - 1)
    Else
        FScaloppi = CVErr(xlErrDiv0)
    End If
End Function
Function Qminimoxseccion(area As Double, ETc As Double, Trd As Double)
'Calcula el caudal minimo necesario para regar una sección de riego en un area total
area = area * 10000
ETc = ETc / 1000
Qminimoxseccion = (area * ETc) / (Trd * 3.6)
End Function

Function LaminaHoraria(Qemisor As Double, SepEmisor As Double, SepRegantes As Double)
'Estima la lamina horaria aplicada por el emisor
Dim areamojada As Double
    areamojada = SepEmisor * SepRegantes
    LaminaHoraria = Qemisor / (areamojada)
End Function
Function LaminaHorariaGoteo(Qemisor As Double, SepEmisor As Double, SepRegantes As Double, areaMojado As Double)
'Estima la lamina horaria aplicada por el emisor
Dim areamojada As Double
    areamojada = SepEmisor * SepRegantes
    LaminaHorariaGoteo = Qemisor / (areamojada) / (areaMojado / 100)
End Function
Public Function LaminaHorariaMicro(diametro, gasto As Double)
Dim area, q As Double
area = WorksheetFunction.pi * (diametro) ^ 2 / 4 'm2
q = gasto / 1000 ' m3/hora
LaminaHorariaMicro = q / area * 1000
End Function
Function Qtotalreq(lh As Double, area As Double)
    lh = lh / 1000
    area = area * 10000
    Qtotalreq = (lh * area) / 3.6
End Function
Function dinterno(diametro As Double) As Double
'Estima el diametro interno en función del diametro nominal y redondea
'al diametro siguiente en función del valor
Dim DN(16) As Double
    DN(1) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A4").Value
    DN(2) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A5").Value
    DN(3) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A6").Value
    DN(4) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A7").Value
    DN(5) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A8").Value
    DN(6) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
    DN(7) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
    DN(8) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
    DN(9) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
    DN(10) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
    DN(11) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
    DN(12) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
    DN(13) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
    DN(14) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
    DN(15) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
    DN(16) = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value

Select Case diametro
   Case 1 To DN(1)
        'dinterno = 18.1
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B4").Value
   Case DN(1) + 0.0000001 To DN(2)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B5").Value
   Case DN(2) + 0.0000000001 To DN(3)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B6").Value
   Case DN(3) + 0.0000000001 To DN(4)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B7").Value
   Case DN(4) + 0.0000000001 To DN(5)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B8").Value
   Case DN(5) + 0.0000000001 To DN(6)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B9").Value
   Case DN(6) + 0.0000000001 To DN(7)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B10").Value
   Case DN(7) + 0.0000000001 To DN(8)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B11").Value
   Case DN(8) + 0.0000000001 To DN(9)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B12").Value
   Case DN(9) + 0.0000000001 To DN(10)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B13").Value
   Case DN(10) + 0.0000000001 To DN(11)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B14").Value
   Case DN(11) + 0.0000000001 To DN(12)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B15").Value
   Case DN(12) + 0.0000000001 To DN(13)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B16").Value
   Case DN(13) + 0.0000000001 To DN(14)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B17").Value
   Case DN(14) + 0.0000000001 To DN(15)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B18").Value
   Case DN(15) + 0.0000000001 To DN(16)
        dinterno = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B19").Value
End Select
End Function
Function Dcalculado(caudal As Double)
'Sugiere el valor del diametro de un tubo en función del caudal a pasar por el
'no toma en cuenta para nada el valor de la fricción
Dim DiaMat As Double
caudal = caudal / 1000
DiaMat = (Sqr(caudal)) * 0.9213 * 1000
Select Case DiaMat
   Case 1 To 18.1
        Dcalculado = 13
   Case 18.1 To 22.7
      Dcalculado = 19
   Case 22.7 To 30.4
      Dcalculado = 25
   Case 30.4 To 39
      Dcalculado = 32
   Case 39 To 55.7
      Dcalculado = 50
   Case 55.7 To 82.1
      Dcalculado = 75
    Case 82.1 To 105.5
      Dcalculado = 100
    Case 105.5 To 154.4
      Dcalculado = 160
    Case 154.4 To 193
      Dcalculado = 200
    Case 193 To 241.2
      Dcalculado = 250
    Case 241.2 To 303.8
        Dcalculado = 315
    Case 303.8 To 342.6
        Dcalculado = 355
    Case 342.6 To 386
        Dcalculado = 400
    Case 386 To 434.2
        Dcalculado = 450
    Case 434.2 To 482.4
        Dcalculado = 500
    Case 482.4 To 607.8
        Dcalculado = 630
End Select

End Function

Function TirNormal(gasto, Ancho, Pendiente, Talud, Manning As Double)
'Sugiere el valor del diametro de un tubo en función del caudal a pasar por el
'no toma en cuenta para nada el valor de la fricción
        Dim yn, pn1, an, Pn, RH, pn2, pn3, pn4, ptn, Theta, So As Double
        Theta = Math.Atn(Pendiente)
        So = Pendiente
        If Ancho = 0 Then
        yn = ((gasto * Manning * (2 * (1 + Talud * Talud) ^ 0.5) ^ (2 / 3)) / ((So) ^ 0.5 * (Talud) ^ (5 / 3))) ^ (3 / 8)
        Else
        If So >= 0.001 Then
        yn = 1
        Else
        yn = 50
        End If
        pn1 = (gasto * Manning) / (So) ^ (0.5)
        an = (Ancho + Talud * yn) * yn
        Pn = Ancho + 2 * yn * ((Talud) ^ (2) + 1) ^ (0.5)
        RH = (an) / (Pn)
        pn2 = an * (RH) ^ (2 / 3)
        pn3 = 2 * (RH) ^ (2 / 3) - 2 * (RH) ^ (5 / 3) * (1 + (Talud) ^ (2)) ^ (0.5)
        Do While Abs(pn1 - pn2) >= 0.00011
            pn1 = (gasto * Manning) / (So) ^ (0.5)
            an = (Ancho + Talud * yn) * yn
            Pn = Ancho + 2 * yn * ((Talud) ^ (2) + 1) ^ (0.5)
            RH = (an) / (Pn)
            pn2 = an * ((RH) ^ (2 / 3))
            pn3 = 2 * Ancho + 2 * Talud * yn * (RH) ^ (2 / 3) - 2 * (RH) ^ (5 / 3) * (1 + (Talud) ^ (2)) ^ (0.5)
            ptn = (pn2 - pn1) / pn3
            yn = yn - ptn
        Loop
        End If
TirNormal = yn
End Function

Function LongMaxRegante(GastoEmisor, s, hf, diametro As Double)
Dim Coeficiente, DI, nmi, a, b, C, Sem, q As Double
Dim L, Qt, a1, B1, C1, F, HP, ah, res, Rey, fdw As Double
    
    Coeficiente = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E1").Value
    DI = diametro / 1000
    Sem = s
    q = GastoEmisor
    HP = hf
    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
        nmi = 1.852
        a = 0
        b = 10000
        ah = 10.648 / ((Coeficiente) ^ 1.852 * (DI) ^ 4.871 * (3600000) ^ 1.852)
        res = 1
        Do While Abs(res) > 0.0000001
            'calculamos C
            C = (a + b) / 2
            'Evaluamos en C
            L = Sem * C 'calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            a1 = 1 / (nmi + 1)
            B1 = 1 / (2 * C)
            C1 = (nmi - 1) ^ 0.5 / (6 * (C) ^ 2)
            F = a1 + B1 + C1 'calculo del Factor de salidas Multiples
            res = ah * F * L * Qt - HP 'calculo de la perdida de carga
                If res > 0 Then
                    b = C
                Else
                    a = C
                End If
        Loop
    
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
    
        nmi = 2
        ah = 10.3 * (Coeficiente) ^ 2 / ((DI) ^ (16 / 3) * (3600000) ^ 2)
        a = 0
        b = 10000
        res = 1
        Do While Abs(res) > 0.0000001
            C = (a + b) / 2
            L = Sem * C ' calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            a1 = 1 / (nmi + 1)
            B1 = 1 / (2 * C)
            C1 = (nmi - 1) ^ 0.5 / (6 * (C) ^ 2)
            F = a1 + B1 + C1 'calculo del Factor de salidas Multiples
            res = ah * F * L * Qt - HP 'calculo de la perdida de carga
                If res > 0 Then
                    b = C
                Else
                    a = C
                End If
        Loop
    
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
        nmi = 1.9
        ah = 4.098 * (10) ^ (-3) * Coeficiente / ((DI) ^ (4.9) * (3600000) ^ nmi)
        a = 0
        b = 10000
        res = 1
        Do While Abs(res) > 0.0000001
            C = (a + b) / 2
            L = Sem * C ' calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            a1 = 1 / (nmi + 1)
            B1 = 1 / (2 * C)
            C1 = (nmi - 1) ^ 0.5 / (6 * (C) ^ 2)
            F = a1 + B1 + C1 'calculo del Factor de salidas Multiples
            res = ah * F * L * Qt - HP 'calculo de la perdida de carga
                If res > 0 Then
                    b = C
                Else
                    a = C
                End If
        Loop
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 4 Then
        nmi = 2
        ah = 0.0827 / ((DI) ^ (5) * (3600000) ^ 2)
        a = 0
        b = 10000
        res = 1
        Do While Abs(res) > 0.00000001
            C = (a + b) / 2
            L = Sem * C ' calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            a1 = 1 / (nmi + 1)
            B1 = 1 / (2 * C)
            C1 = (nmi - 1) ^ 0.5 / (6 * (C) ^ 2)
            F = a1 + B1 + C1 'calculo del Factor de salidas Multiples
            Rey = Workbooks("RegisterU2DF7.xlam").NReynoldsP((q * C) / 3600, DI * 1000) * 1
                If Rey <= 2000 Then
                        fdw = 64 / Rey
                Else
                    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E2").Value = 0 Then
                        fdw = Workbooks("RegisterU2DF7.xlam").CoeFriccionDWP(Rey, Coeficiente, DI * 1000)
                    Else
                        fdw = Workbooks("RegisterU2DF7.xlam").CoeFriccionSJ(Rey, Coeficiente, DI * 1000)
                    End If
                End If
            res = ah * fdw * F * L * Qt - HP 'calculo de la perdida de carga
                If res > 0 Then
                    b = C
                Else
                    a = C
                End If
        Loop
    
    
    End If
    LongMaxRegante = L
    

End Function

Function PrecipitacionEfectiva(Precipitacion As Double)
Dim P, PE As Double
    P = Precipitacion
    If Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "Porcentaje fijo" Then
        PE = P * (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B9").Value) / 100
    
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "Precipitacion Confiable" Then
        If P <= 70 Then
            PE = 0.6 * P - 10
        Else
            PE = 0.8 * P - 24
        End If
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "Formula empirica" Then
        If P <= (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("H6").Value) Then
            PE = (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B6").Value) * P + ((Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("E6").Value)) * 1
        Else
            PE = (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B7").Value) * P + ((Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("E7").Value)) * 1
        End If
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "USDA" Then
        If P <= 250 Then
            PE = (P * (125 - 0.2 * P)) / 125
        Else
            PE = 0.1 * P + 125
        End If
    Else
        PE = 0
    End If
    If PE <= 0 Then
    PrecipitacionEfectiva = 0
    Else
    PrecipitacionEfectiva = PE
    End If
End Function

Function EvapotranspiracionA(Evaporacion, VelocidadViento, HumedadRelativa, CoberturaTanque As Double)

    Dim Ev, U2, HR, d, Kt, Eto As Double
    Ev = Evaporacion
    U2 = VelocidadViento
    HR = HumedadRelativa
    d = CoberturaTanque
    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 1 Then
        U2 = U2 * 86400 / 1000
        Kt = 0.475 - 0.00024 * U2 + 0.00516 * HR + 0.00118 * d - 0.000016 * (HR) ^ 2 - 0.101 * (10) ^ -5 * (d) ^ 2 - 0.8 * (10) ^ -8 * (HR) ^ 2 * U2 - 1 * (10) ^ -8 * (HR) ^ 2 * d
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B63").Value = 2 Then
        Kt = 0.108 - 0.0286 * U2 + 0.0422 * WorksheetFunction.Ln(d) + 0.1434 * WorksheetFunction.Ln(HR) - 0.000631 * (WorksheetFunction.Ln(d)) ^ 2 * WorksheetFunction.Ln(HR)
    Else
        Kt = 0.61 + 0.00341 * HR - 0.000162 * U2 * HR - 0.00000959 * U2 * d + 0.00327 * U2 * WorksheetFunction.Ln(d) - 0.00289 * U2 * WorksheetFunction.Ln(86.4 * d) - 0.0106 * WorksheetFunction.Ln(86.4 * U2) * WorksheetFunction.Ln(d) + 0.00063 * (WorksheetFunction.Ln(d)) ^ 2 * WorksheetFunction.Ln(86.4 * U2)
    End If
    Eto = Kt * Ev
    EvapotranspiracionA = Eto
End Function

Function PotenciaBomba(gasto, Presion, EfiBomba, EfiMotor As Double)
    Dim Qa, Pr, EB, EM, pot As Double
    Qa = gasto
    Pr = Presion
    EB = EfiBomba / 100
    EM = EfiMotor / 100
    pot = Qa * Pr / (76 * EB * EM)
    PotenciaBomba = pot
End Function

Function perdida(gasto, diametro, longitud As Double)
'Determina la perdida por fricción en la tuberia con PVC
Dim Coefp, DIp, Qp, Longp, Hfp, Rey, fdw As Double
    Qp = gasto / 1000
    Longp = longitud
    DIp = diametro / 1000
    Coefp = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E1").Value
    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
        Hfp = 10.648 * 1 / (Coefp) ^ (1.852) * (Qp) ^ (1.852) / (DIp) ^ (4.871) * Longp
    ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
        Hfp = 10.3 * (Coefp) ^ (2) * (Qp) ^ (2) / (DIp) ^ (16 / 3) * Longp
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
    
    


    perdida = Hfp
    
End Function
'ULTIMA MODIFICACIÓN 12 DE septiembre DE 2014 by sergio ivan jimenez jimenez
Public Function GalonphLitroph(GPH As Double) As Double
'CONVIERTE LAS UNIDADES DE GALONES POR HORA A LITROS POR HORA
If GPH = 0 Then
GalonphLitroph = CVErr(xlErrDiv0)
End If
If IsNumeric(GPH > 0) = True Then
GalonphLitroph = 3.7854118 * GPH
End If
End Function

Public Function LitrophGalonph(LPH As Double) As Double
'CONVIERTE LAS UNIDADES DE LITROS POR HORA A GALONES POR HORA
If LPH = 0 Then
LitrophGalonph = CVErr(xlErrDiv0)
End If
If IsNumeric(LPH > 0) = True Then
LitrophGalonph = LPH / 3.7854118
End If
End Function
Public Function TexturaSuelo(arena, limo, arcilla As Double) As String
'Determina la textura del suelo
Dim suma As Double
suma = arena + limo + arcilla
If suma > 101 Or suma < 99 Then
    TexturaSuelo = "Suma diferente a 100%"
Else
    If (arena >= 0 And arena <= 45) And (limo >= 0 And limo <= 40) And (arcilla >= 40 And arcilla <= 100) Then
        TexturaSuelo = "Arcillosa"
    ElseIf (arena >= 45 And arena < 65) And (limo >= 0 And limo < 20) And (arcilla >= 35 And arcilla < 55) Then
        TexturaSuelo = "Arcillosa Arenosa"
    ElseIf (arena >= 20 And arena < 45) And (limo >= 15 And limo < 52) And (arcilla >= 27 And arcilla < 40) Then
        TexturaSuelo = "Franco Arcillosa"
    ElseIf (arena >= 0 And arena <= 20) And (limo >= 40 And limo < 60) And (arcilla >= 40 And arcilla < 60) Then
        TexturaSuelo = "Arcillosa Limosa"
    ElseIf (arena >= 0 And arena <= 20) And (limo >= 40 And limo <= 73) And (arcilla >= 27 And arcilla <= 40) Then
        TexturaSuelo = "Franco Arcillosa Limosa"
    ElseIf (arena >= 45 And arena <= 80) And (limo >= 0 And limo <= 28) And (arcilla >= 20 And arcilla <= 35) Then
        TexturaSuelo = "Franco Arcillosa Arenosa"
    ElseIf (arena >= 23 And arena < 52) And (limo >= 28 And limo < 50) And (arcilla >= 7 And arcilla < 27) Then
        TexturaSuelo = "Franco"
    ElseIf (arena >= 0 And arena < 50) And (limo >= 50 And limo < 100) And (arcilla >= 0 And arcilla < 27) Then
        If (arena >= 0 And arena < 20) And (limo >= 80 And limo <= 100) And (arcilla >= 0 And arcilla < 12) Then
            TexturaSuelo = "Limosa"
        Else
            TexturaSuelo = "Franco Limosa"
        End If
    ElseIf ((arena >= 42 And arena < 52) And (limo >= 42 And limo <= 50) And (arcilla >= 0 And arcilla < 7)) Or ((arena >= 52 And arena <= 70) And (limo >= 10 And limo < 50) And (arcilla >= 0 And arcilla < 20)) Or ((arena >= 70 And arena <= 80) And (limo >= 0 And limo < 20) And (arcilla >= 10 And arcilla < 20)) Then
        TexturaSuelo = "Franco Arenosa"
    ElseIf (arena >= 70 And arena <= 80) And (limo >= 10 And limo <= 30) And (arcilla >= 0 And arcilla <= 10) Then
        If Abs(80 - arena) + arcilla >= 10 Then
            TexturaSuelo = "Franco Arenosa"
        Else
            TexturaSuelo = "Arenosa Franco"
        End If
    ElseIf (arena >= 80 And arena < 90) And (limo >= 0 And limo < 10) And (arcilla >= 10 And arcilla < 20) Then
        If Abs(90 - arena) + Abs(10 - arcilla) >= 10 Then
            TexturaSuelo = "Franco Arenosa"
        Else
            TexturaSuelo = "Arenosa Franco"
        End If
    ElseIf (arena >= 80 And arena < 85) And (limo >= 0 And limo < 20) And (arcilla >= 0 And arcilla < 10) Then
        TexturaSuelo = "Arenosa Franco"
    ElseIf (arena >= 90 And arena <= 100) And (limo >= 0 And limo < 10) And (arcilla >= 0 And arcilla <= 10) Then
        TexturaSuelo = "Arenosa"
    ElseIf (arena >= 85 And arena <= 90) And (limo >= 0 And limo < 15) And (arcilla >= 0 And arcilla <= 10) Then
        If (Abs(90 - arena) * 2 + arcilla) >= 10 Then
            TexturaSuelo = "Arenosa Franco"
        Else
            TexturaSuelo = "Arenosa"
        End If
    Else
        TexturaSuelo = "Error"
    End If
End If
End Function
'Eto Priestley Taylor model
Public Function EToPriestleTaylor(Juliano, Tmax, Tmin, Tmean, RH, Rs, elevation, Latitud) As Double
Dim LE, a, cs, P, Es, Ea, Ra, Rnl, RT, Rn, AT, Rso, Rns As Double
'Tmean = (Tmax + Tmin) / 2
    If Juliano > 366 Or Juliano <= 0 Or elevation < 0 Then
        EToPriestleTaylor = 0 / 0
    ElseIf Tmax < Tmin Or RH > 100 Then
        EToPriestleTaylor = 0 / 0
    Else
        LE = LatentHVaporization(Tmean)
        a = SlopeSaturationVP(Tmean)
        P = AtmosphericP(elevation)
        cs = PsychrometricC(Tmean, P)
        Ea = ActualVaporP(RH, Tmean)
        Ra = RadiacionExtraterrestres(Juliano, Latitud)
        Rso = ClearSkySR(elevation, Ra)
        Rns = NetShortwaveR(0.23, Rs)
        Rnl = Rnlarga(Tmax, Tmin, Ea, Rs, Rso)
            If Rnl <= 0 Then
                Rnl = 0
            Else
            End If
        Rn = Rns - Rnl
        EToPriestleTaylor = 1 / LE * 1.26 * (a) / (a + cs) * (Rn - 0)
    End If
End Function
Public Function EToPM(Juliano, Tmax, Tmin, Tmean, v, RH, Rs, elevation, Latitud) As Double
Dim a, cs, P, Es, Ea, Ra, Rnl, RT, Rn, AT, Rso, Rns As Double
'Tmean = (Tmax + Tmin) / 2
If Juliano > 366 Or Juliano <= 0 Or elevation < 0 Then
    EToPM = 0 / 0
ElseIf Tmax < Tmin Or RH > 100 Then
    EToPM = 0 / 0
Else
    a = SlopeSaturationVP(Tmean)
    P = AtmosphericP(elevation)
    cs = PsychrometricC(Tmean, P)
    Es = SaturationVaporP(Tmax, Tmin)
    Ea = ActualVaporP(RH, Tmean)
    Ra = RadiacionExtraterrestres(Juliano, Latitud)
    Rso = ClearSkySR(elevation, Ra)
    Rns = NetShortwaveR(0.23, Rs)
    Rnl = Rnlarga(Tmax, Tmin, Ea, Rs, Rso)
    If Rnl <= 0 Then
        Rnl = 0
    Else
    End If
    
    Rn = Rns - Rnl
    RT = RadiationTerm(a, cs, v, Rn)
    AT = AdvectionTerm(a, cs, v, Tmean, Es, Ea)
    EToPM = RT + AT
End If
End Function
Public Function Windspeed(Velocidad, altura As Double)
    Windspeed = Velocidad * 4.87 / Math.Log(67.8 * altura - 5.42)
End Function
Function aDiaJulianoo(Fecha As Long) As Integer
' verificar cuantos dias julianos tiene el año
Dim año As Integer
Dim dia As Integer
Dim DiaJ As String

año = Year(Fecha)
dia = DateDiff("d", DateSerial(año, 1, 0), Fecha)

If Fecha = DateSerial(año, 2, 29) Then
    dia = 59
        DiaJ = Format(dia, "000")
    aDiaJulianoo = DiaJ

End If

If Fecha = DateSerial(año, 3, 1) Then

    dia = 60
    DiaJ = Format(dia, "000")
    aDiaJulianoo = DiaJ

Else
    If dia > 60 Then
        dia = dia
        DiaJ = Format(dia, "000")
        aDiaJulianoo = DiaJ
    End If
End If
        DiaJ = Format(dia, "000")
        aDiaJulianoo = DiaJ
End Function

Function EToHargreavesSamani(Juliano, Latitud, TemMaxima, TemMinima, Tm As Double) As Double
Dim Ho, Tx, Tn, Krs, Eto, Cl, Kh, eh, Kt As Double

If Juliano > 366 Or Juliano <= 0 Then
    EToHargreavesSamani = 0 / 0
ElseIf TemMaxima < TemMinima Then
    EToHargreavesSamani = 0 / 0
Else
    Kh = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B18").Value
    eh = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B19").Value
    Kt = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B20").Value
    
    Ho = RadiacionExtraterrestres(Juliano, Latitud)
    Tx = TemMaxima * 1
    Tn = TemMinima * 1
    
    'Tm = (Tx + Tn) / 2
    'Cl = 2.501 - (2.361 * (10) ^ (-3) * Tm)
    If Tx <= Tn Or Tm >= Tx Then
        Eto = 0
    Else
        Eto = 0.408 * Kh * Ho * (Tm + Kt) * (Tx - Tn) ^ (eh)
    End If
    EToHargreavesSamani = Eto
End If
End Function

Function PMDatosLimitados(Juliano, Tmax, Tmin, Tmean, v, elevation, Latitud) As Double
Dim a, cs, P, Es, Ea, Ra, Rnl, RT, Rn, AT, Kr, Rs, RH, Rso, Rns, Tresta, TRocio As Double
'Tmean = (Tmax + Tmin) / 2
If Juliano > 366 Or Juliano <= 0 Or elevation <= 0 Then
    PMDatosLimitados = 0 / 0
ElseIf Tmax < Tmin Then
    PMDatosLimitados = 0 / 0
Else
    If Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B23").Value = 1 Then
        Kr = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B24").Value
    Else
        Kr = Krs(elevation)
    End If
    Tresta = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B26").Value
    TRocio = Tmin - Tresta
    RH = HRsindatos(Tmin, Tmean)
    a = SlopeSaturationVP(Tmean)
    P = AtmosphericP(elevation)
    cs = PsychrometricC(Tmean, P)
    Es = SaturationVaporP(Tmax, Tmin)
    Ea = SVP(TRocio)
    Ra = RadiacionExtraterrestres(Juliano, Latitud)
    Rs = SolarRadiation(Tmax, Tmin, Ra, Kr)
    Rso = ClearSkySR(elevation, Ra)
    Rns = NetShortwaveR(0.23, Rs)
    Rnl = Rnlarga(Tmax, Tmin, Ea, Rs, Rso)
    Rn = Rns - Rnl
    RT = RadiationTerm(a, cs, v, Rn)
    AT = AdvectionTerm(a, cs, v, Tmean, Es, Ea)
    PMDatosLimitados = RT + AT
End If
End Function
Function RadiacionExtraterrestres(Juliano, Latitud)
Dim J As String
Dim P, dr, s, Ph, Ws, Ra, pi As Double
J = Juliano
P = Latitud
pi = WorksheetFunction.pi()
dr = 1 + 0.033 * Cos(2 * pi * J / 365)
s = 0.409 * Sin(2 * pi * J / 365 - 1.39)
Ph = pi / 180 * P
Ws = WorksheetFunction.Acos(-Tan(Ph) * Tan(s))
Ra = 24 * 60 / pi * 0.082 * dr * (Ws * Sin(Ph) * Sin(s) + Cos(Ph) * Cos(s) * Sin(Ws))
RadiacionExtraterrestres = Ra
End Function
Public Function VelocidadFlujo(gasto, DIp) As Double
Dim area, Qp As Double
    Qp = gasto / 1000
    area = (WorksheetFunction.pi) * (DIp / 1000) ^ 2 / 4
    VelocidadFlujo = (Qp) / area
End Function

Public Function NReynolds(gasto, DI) As Double
Dim vel, v As Double
    v = 0.000001007
    vel = VelocidadFlujo(gasto, DI)
    NReynolds = (DI / 1000) * vel / v
End Function

Public Function CoeFriccionDW(re, e, d) As Double
' proponemos valores de a y b
Dim a, b, C, F, fa, fb As Double
a = 0
b = 10
' Proponemos un valor de f inicial para que entre al ciclo
F = 1
Do While Abs(F) > 0.00001
    ' Calculamos el valor de C
    C = (a + b) / 2
    ' Evaluamos en C
    fa = 1 / (C) ^ 0.5
    fb = -2 * (Math.Log((e / d) / 3.71 + 2.51 / (re * (C) ^ 0.5))) * 0.434294481903252
    F = fb - fa
    If F > 0 Then
        b = C
    Else
        a = C
    End If

Loop
    CoeFriccionDW = C
End Function

Public Function CoeFriccionSJ(re, e, d) As Double
' proponemos valores de a y b
Dim a, b, C, F, fa, fb As Double
CoeFriccionSJ = 0.25 / ((Math.Log((e / d) / 3.7 + 5.74 / ((re) ^ 0.9))) * 0.434294481903252) ^ 2
End Function


'----- aqui inicia formulas complementarias--------

Public Function RadMJ(Radiacion As Double)
    RadMJ = Radiacion * 0.0864
End Function


Public Function SlopeSaturationVP(Tmean)
SlopeSaturationVP = 4098 * (0.6108 * Math.Exp((17.27 * Tmean) / (Tmean + 237.3))) / (Tmean + 237.3) ^ 2
End Function
Public Function AtmosphericP(elevation)
    AtmosphericP = 101.3 * ((293 - 0.0065 * elevation) / 293) ^ 5.26
End Function

Public Function LatendHVaporization(Tmean)
LatendHVaporization = 2.501 - (2.361 * (10) ^ (-3) * Tmean)
End Function

Public Function PsychrometricC(Tmean, AtmosfericP)
    PsychrometricC = 1.013 * (10) ^ (-3) * AtmosfericP / (LatendHVaporization(Tmean) * 0.622)
End Function
Public Function SVP(T)
    SVP = 0.6108 * Math.Exp((17.27 * T) / (T + 237.3))
End Function
Public Function SaturationVaporP(Tmax, Tmin)
    SaturationVaporP = (SVP(Tmax) + SVP(Tmin)) / 2
End Function

Public Function ActualVaporP(RH, Tmean)
    ActualVaporP = RH / 100 * SVP(Tmean) '((SVP(Tmax) + SVP(Tmin)) / 2)
End Function

Public Function ClearSkySR(Ele, Ra)
    ClearSkySR = (0.75 + 2 * (10) ^ (-5) * Ele) * Ra
End Function
Public Function NetShortwaveR(a, Rs)
    NetShortwaveR = (1 - a) * Rs
End Function
Public Function Rnlarga(Tmax, Tmin, Ea, Rs, Rso)
Dim a1, a2, a3, Sb, Tmk4, Tmink4 As Double
    Sb = 4.903 * (10) ^ (-9)
    Tmk4 = (Tmax + 273.16) ^ 4
    Tmink4 = (Tmin + 273.16) ^ 4
    a1 = Sb * (Tmk4 + Tmink4) / 2
    a2 = 0.34 - 0.14 * (Ea) ^ 0.5
    a3 = 1.35 * Rs / Rso - 0.35
    Rnlarga = a1 * a2 * a3
End Function

Public Function RadiationTerm(slope, cs, Velocity, Rn)
Dim Apoyo As Double
    Apoyo = slope + cs * (1 + 0.34 * Velocity)
    RadiationTerm = 0.408 * slope * (Rn - 0) / Apoyo
End Function
Public Function AdvectionTerm(slope, cs, Velocity, T, Es, Ea) As Double
Dim Apoyo As Double
    Apoyo = slope + cs * (1 + 0.34 * Velocity)
    AdvectionTerm = cs * (900 / (T + 273)) * Velocity * (Es - Ea) / Apoyo
End Function

Private Function EToPM_RE(Tmax, Tmin, v, RH, Rs, elevation, Ra, Tmean) As Double
Dim a, cs, P, Es, Ea, Rnl, RT, Rn, AT, Rso, Rns As Double
'Tmean = (Tmax + Tmin) / 2
a = SlopeSaturationVP(Tmean)
P = AtmosphericP(elevation)
cs = PsychrometricC(Tmean, P)
Es = SaturationVaporP(Tmax, Tmin)
Ea = RH
'Ra = RadiacionExtraterrestres(Juliano, Latitud)
Rso = ClearSkySR(elevation, Ra)
Rns = NetShortwaveR(0.23, Rs)
Rnl = Rnlarga(Tmax, Tmin, Ea, Rs, Rso)
Rn = Rns - Rnl
RT = RadiationTerm(a, cs, v, Rn)
AT = AdvectionTerm(a, cs, v, Tmean, Es, Ea)
EToPM_RE = RT + AT
End Function

Public Function LatentHVaporization(Tm)
LatentHVaporization = 2.501 - (2.361 * (10) ^ (-3) * Tm)
End Function

Public Function SolarRadiation(Tmax, Tmin, Radiation, Krs)
SolarRadiation = (Tmax - Tmin) ^ 0.5 * Radiation * Krs
End Function

Public Function Krs(elevation)
Dim P As Double
P = AtmosphericP(elevation)
Krs = 0.17 * (P / 101.3) ^ 0.5
End Function

Public Function HRsindatos(Tmin, Tmean)
Dim Ea, m As Double
Ea = SVP(Tmin)
m = SVP(Tmean)
HRsindatos = Ea / m * 100
End Function

Public Function ActualVaporRocio(T) '
    ActualVaporRocio = 0.6108 * Math.Exp((17.27 * T) / (T + 237.3))
End Function
Public Function Rs_SkyCover(SkyCover, Juliano, Latitud)
Dim J As Integer
Dim P, dr, s, Ph, Ws, Ra, pi As Double
J = Juliano
P = Latitud
pi = WorksheetFunction.pi()
dr = 1 + 0.033 * Cos(2 * pi * J / 365)
s = 0.409 * Sin(2 * pi * J / 365 - 1.39)
Ph = pi / 180 * P
Ws = WorksheetFunction.Acos(-Tan(Ph) * Tan(s))
Ra = 24 * 60 / pi * 0.082 * dr * (Ws * Sin(Ph) * Sin(s) + Cos(Ph) * Cos(s) * Sin(Ws))
Rs_SkyCover = (0.25 + (0.5 * ((-0.0083 * SkyCover) + 0.9659))) * Ra

End Function
Public Function VP_SpecHumid(SpecificHumidity, elevation)
Dim P, a As Double
P = AtmosphericP(elevation) * 10
a = (SpecificHumidity * P / (0.622 + 0.378 * SpecificHumidity)) / 10
VP_SpecHumid = a '/ SaturationVaporP(Tmax, Tmin) * 100
' formulas en https://archive.eol.ucar.edu/projects/ceop/dm/documents/refdata_report/eqns.html
End Function
Function MeanError(Medido As Range, Estimado As Range) As Double
    Dim Tfilas, TColumnas, QFilas, Qcolumnas As Integer

        'Para el ia

        Tfilas = Medido.Rows.Count
        TColumnas = Medido.Columns.Count
        'Para el variable conocida
        QFilas = Estimado.Rows.Count
        Qcolumnas = Estimado.Columns.Count
        
            If TColumnas <> 1 Or Qcolumnas <> 1 Then
                MeanError = 0 / 0
            ElseIf Tfilas < 3 Or QFilas < 3 Then
                MeanError = 0 / 0
            ElseIf Tfilas <> QFilas Then
                MeanError = 0 / 0
            Else
                Dim MatrizT(), MatrizQ() As Variant
                Dim XY() As Double
                ReDim XY(Tfilas) As Double
                Dim Sumaxy, diferencia As Double
                
                MatrizT = Medido
                MatrizQ = Estimado
                Sumaxy = 0
                diferencia = 0
                'Pasamos los datos
                For i = 1 To Tfilas
                    If MatrizT(i, 1) = "" Or MatrizQ(i, 1) = "" Then
                        XY(i) = 0
                        diferencia = diferencia + 1
                    Else
                    
                        XY(i) = MatrizQ(i, 1) - MatrizT(i, 1)
                    End If
                        Sumaxy = Sumaxy + XY(i)
                Next

                MeanError = Sumaxy / (Tfilas - diferencia)
            End If
End Function
Function StandarDeviationError(Medido As Range, Estimado As Range, Merror As Double) As Double
    Dim Tfilas, TColumnas, QFilas, Qcolumnas As Integer

        'Para el ia

        Tfilas = Medido.Rows.Count
        TColumnas = Medido.Columns.Count
        'Para el variable conocida
        QFilas = Estimado.Rows.Count
        Qcolumnas = Estimado.Columns.Count
        
            If TColumnas <> 1 Or Qcolumnas <> 1 Then
                StandarDeviationError = 0 / 0
            ElseIf Tfilas < 3 Or QFilas < 3 Then
                StandarDeviationError = 0 / 0
            ElseIf Tfilas <> QFilas Then
                StandarDeviationError = 0 / 0
            Else
                Dim MatrizT(), MatrizQ() As Variant
                Dim XY() As Double
                ReDim XY(Tfilas) As Double
                Dim Sumaxy, diferencia As Double
                
                MatrizT = Medido
                MatrizQ = Estimado
                Sumaxy = 0
                diferencia = 0
                'Pasamos los datos
                For i = 1 To Tfilas
                    If MatrizT(i, 1) = "" Or MatrizQ(i, 1) = "" Then
                        XY(i) = 0
                        diferencia = diferencia + 1
                    Else
                    
                        XY(i) = ((MatrizQ(i, 1) - MatrizT(i, 1)) - Merror) ^ 2
                    End If
                        Sumaxy = Sumaxy + XY(i)
                Next

                StandarDeviationError = (Sumaxy / (Tfilas - diferencia - 1)) ^ 0.5
            End If
End Function
Function dWilmmott(Medido As Range, Estimado As Range) As Double
    Dim Tfilas, TColumnas, QFilas, Qcolumnas As Integer

        'Para el ia

        Tfilas = Medido.Rows.Count
        TColumnas = Medido.Columns.Count
        'Para el variable conocida
        QFilas = Estimado.Rows.Count
        Qcolumnas = Estimado.Columns.Count
        
            If TColumnas <> 1 Or Qcolumnas <> 1 Then
                dWilmmott = 0 / 0
            ElseIf Tfilas < 3 Or QFilas < 3 Then
                dWilmmott = 0 / 0
            ElseIf Tfilas <> QFilas Then
                dWilmmott = 0 / 0
            Else
                Dim MatrizT(), MatrizQ() As Variant
                Dim XY(), XYA() As Double
                ReDim XY(Tfilas), XYA(Tfilas) As Double
                Dim Sumaxy, diferencia, MeanA As Double
                
                MatrizT = Medido
                MatrizQ = Estimado
                Sumaxy = 0
                Sumaxy2 = 0
                diferencia = 0
                'Pasamos los datos

                
                MeanA = WorksheetFunction.Average(Medido)
                
                
                For i = 1 To Tfilas
                    If MatrizT(i, 1) = "" Or MatrizQ(i, 1) = "" Then
                        XY(i) = 0
                        diferencia = diferencia + 1
                    Else
                    
                        XY(i) = (MatrizT(i, 1) - MatrizQ(i, 1)) ^ 2
                        XYA(i) = (Math.Abs(MatrizQ(i, 1) - MeanA) + Math.Abs(MatrizT(i, 1) - MeanA)) ^ 2
                    End If
                        Sumaxy = Sumaxy + XY(i)
                        Sumaxy2 = Sumaxy2 + XYA(i)
                Next

                dWilmmott = 1 - Sumaxy / Sumaxy2
            End If
End Function
Function RMSE(Medido As Range, Estimado As Range) As Double
    Dim Tfilas, TColumnas, QFilas, Qcolumnas As Integer

        'Para el ia

        Tfilas = Medido.Rows.Count
        TColumnas = Medido.Columns.Count
        'Para el variable conocida
        QFilas = Estimado.Rows.Count
        Qcolumnas = Estimado.Columns.Count
        
            If TColumnas <> 1 Or Qcolumnas <> 1 Then
                RMSE = 0 / 0
            ElseIf Tfilas < 3 Or QFilas < 3 Then
                RMSE = 0 / 0
            ElseIf Tfilas <> QFilas Then
                RMSE = 0 / 0
            Else
                Dim MatrizT(), MatrizQ() As Variant
                Dim XY() As Double
                ReDim XY(Tfilas) As Double
                Dim Sumaxy, diferencia As Double
                
                MatrizT = Medido
                MatrizQ = Estimado
                Sumaxy = 0
                diferencia = 0
                'Pasamos los datos
                For i = 1 To Tfilas
                    If MatrizT(i, 1) = "" Or MatrizQ(i, 1) = "" Then
                        XY(i) = 0
                        diferencia = diferencia + 1
                    Else
                    
                        XY(i) = ((MatrizQ(i, 1) - MatrizT(i, 1))) ^ 2
                    End If
                        Sumaxy = Sumaxy + XY(i)
                Next

                RMSE = (Sumaxy / (Tfilas - diferencia)) ^ 0.5
            End If
End Function


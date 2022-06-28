VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Regante 
   Caption         =   "Tuberías con Salidas Multiples de Servicio Mixto"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8880.001
   OleObjectBlob   =   "Regante.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Regante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BCalcular_Click()

Dim Coeficiente, DI, nmi, a, b, C, Sem, q As Double
Dim L, Qt, a1, B1, C1, F, HP, ah, res, HP1 As Double
Dim Rey, fdw As Double

If SE.Text = "" Or QE.Text = "" Or PresionEmisor.Text = "" Or MaxPresion.Text = "" Or PendienteTerreno.Text = "" Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
ElseIf SE.Text = 0 Or QE.Text = 0 Or PresionEmisor.Text = 0 Then
    MsgBox "Ningun valor debe ser igual a cero", vbCritical, "HF Riego Dice:"
ElseIf SE.Text > 50 Or PresionEmisor.Text > 100 Or MaxPresion.Text > 50 Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
ElseIf PendienteTerreno.Text > 50 Or PendienteTerreno.Text < -50 Then
    MsgBox "Pendiente incorrecta", vbCritical, "HF Riego Dice:"
Else
    Coeficiente = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E1").Value
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B32").Value = DN.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B29").Value = QE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B30").Value = SE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B31").Value = PresionEmisor.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B34").Value = ComboBoxSo.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F29").Value = MaxPresion.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F30").Value = PendienteTerreno.Text
    
    DI = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B33").Value
    Sem = (SE.Text) * 1
    If ComboBox2.Text = "lph" Then
            q = (QE.Text)
    ElseIf ComboBox2.Text = "lps" Then
            q = (QE.Text) * 3600
    Else
            q = (QE.Text) * 3600 * 1000
    End If
    
    'Q = (QE.Text) * 1
    HP1 = (PresionEmisor.Text) * (MaxPresion.Text) / 100
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
            'L = Sem * C 'calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida

            If ComboBoxSo.Text = "S0=S" Then
                F = 1 / (nmi + 1) + 1 / (2 * C) + (nmi - 1) ^ 0.5 / (6 * (C) ^ 2) 'calculo del Factor de salidas Multiples
                L = Sem * C
            Else
                F = (2 * C / (2 * C - 1)) * (1 / (nmi + 1) + (Sqr(nmi - 1)) / (6 * (C) ^ 2))
                L = Sem * C - Sem / 2
            End If
            HP = HP1 - (PendienteTerreno.Text / 100 * L)
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
            'L = Sem * C ' calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            If ComboBoxSo.Text = "S0=S" Then
                F = 1 / (nmi + 1) + 1 / (2 * C) + (nmi - 1) ^ 0.5 / (6 * (C) ^ 2) 'calculo del Factor de salidas Multiples
                L = Sem * C
            Else
                F = (2 * C / (2 * C - 1)) * (1 / (nmi + 1) + (Sqr(nmi - 1)) / (6 * (C) ^ 2))
                L = Sem * C - Sem / 2
            End If
            HP = HP1 - (PendienteTerreno.Text / 100 * L)
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
            'L = Sem * C ' calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            If ComboBoxSo.Text = "S0=S" Then
                F = 1 / (nmi + 1) + 1 / (2 * C) + (nmi - 1) ^ 0.5 / (6 * (C) ^ 2) 'calculo del Factor de salidas Multiples
                L = Sem * C
            Else
                F = (2 * C / (2 * C - 1)) * (1 / (nmi + 1) + (Sqr(nmi - 1)) / (6 * (C) ^ 2))
                L = Sem * C - Sem / 2
            End If
            HP = HP1 - (PendienteTerreno.Text / 100 * L)
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
            'L = Sem * C ' calculo de la longitud por cada salida
            Qt = (q * C) ^ (nmi) ' calculo del gasto por cada  salida
            If ComboBoxSo.Text = "S0=S" Then
                F = 1 / (nmi + 1) + 1 / (2 * C) + (nmi - 1) ^ 0.5 / (6 * (C) ^ 2) 'calculo del Factor de salidas Multiples
                L = Sem * C
            Else
                F = (2 * C / (2 * C - 1)) * (1 / (nmi + 1) + (Sqr(nmi - 1)) / (6 * (C) ^ 2))
                L = Sem * C - Sem / 2
            End If
            
            
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
                
            HP = HP1 - (PendienteTerreno.Text / 100 * L)
            res = ah * fdw * F * L * Qt - HP 'calculo de la perdida de carga
                If res > 0 Then
                    b = C
                Else
                    a = C
                End If
        Loop
    
        ah = ah * fdw
    
    End If
    C = Fix(C)
    If ComboBoxSo.Text = "S0=S" Then
        L = Sem * C
    Else
        L = Sem * C - Sem / 2
    End If
    Qt = (q * C) ^ (nmi)
    If ComboBoxSo.Text = "S0=S" Then
            F = 1 / (nmi + 1) + 1 / (2 * C) + (nmi - 1) ^ 0.5 / (6 * (C) ^ 2) 'calculo del Factor de salidas Multiples
    Else
            F = (2 * C / (2 * C - 1)) * (1 / (nmi + 1) + (Sqr(nmi - 1)) / (6 * (C) ^ 2))
    End If
    NS.Text = FormatNumber(CDbl(C), 0)
    LM.Text = FormatNumber(CDbl(L), 3)
    LmPE.Text = FormatNumber(CDbl(ah * F * L * Qt + PendienteTerreno.Text / 100 * L), 3)
    LmPF.Text = FormatNumber(CDbl(ah * F * L * Qt), 3)
    LmPresion.Text = FormatNumber(CDbl(PresionEmisor.Text * 1 + LmPE.Text * 1), 3)
    LmCaudal.Text = Qt
    If ComboBox2.Text = "lph" Then
            LmCaudal.Text = FormatNumber(CDbl(C * QE.Text / 3600), 3)
    ElseIf ComboBox2.Text = "lps" Then
            LmCaudal.Text = FormatNumber(CDbl(C * QE.Text), 3)
    Else
            LmCaudal.Text = FormatNumber(CDbl(C * QE.Text * 1000), 3)
    End If
    'Workbooks("RegisterU2DF7.xlam").Save

End If
End Sub

Private Sub BCerrar_Click()
Unload Me
End Sub

Private Sub BLimpiar_Click()
QE.Text = ""
SE.Text = ""
Hfp.Text = ""
NS.Text = ""
LM.Text = ""
LmPE.Text = ""
End Sub

Private Sub ComboBox1_Change()
ComboBox1.Style = fmStyleDropDownList
End Sub

Private Sub ComboBoxSo_Change()
ComboBoxSo.Style = fmStyleDropDownList
End Sub
Private Sub CommandButton1_Click()
Ajustes.Show
End Sub
Private Sub CommandButton2_Click()
Salida.Show
End Sub
Private Sub CommandButton3_Click()
Salida.Show
End Sub

Private Sub CommandButton4_Click()
Pendientex.Show
End Sub

Private Sub CommandButton5_Click()
Pendientex.Show
End Sub

Private Sub CommandButton6_Click()
HelpTSM1.Show
End Sub

Private Sub CommandButton7_Click()
HelpTSM2.Show
End Sub

Private Sub CommanExportar_Click()
    If NS.Text = "" Or LM.Text = "" Or LmPE.Text = "" Or SE.Text = "" Or MaxPresion.Text = "" Then
        MsgBox "Debe realizar un calculo antes de exportar a excel", vbCritical, "HF Riego Dice:"
    Else
        '3.- Importamos la hoja de Excel del complemento
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("B3").Value = QE.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("B4").Value = SE.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("B5").Value = PresionEmisor.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("B6").Value = MaxPresion.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("B7").Value = PendienteTerreno.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("B8").Value = DN.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("B9").Value = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B33").Value
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("B10").Value = ComboBoxSo.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("C3").Value = ComboBox2.Text
        '4.-Exportamos los resultados generales a excel
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("F3").Value = NS.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("F4").Value = LM.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("F5").Value = LmPE.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("F6").Value = LmPF.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("F7").Value = LmPresion.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Range("F8").Value = LmCaudal.Text
                
                    hojas = ActiveSheet.Name
                    Workbooks("RegisterU2DF7.xlam").Worksheets("LMax").Copy _
                           after:=ActiveWorkbook.Sheets(hojas)
                          MsgBox "Se realizo con exito"
    End If
End Sub

Private Sub DN_Change()
DN.Style = fmStyleDropDownList
End Sub

Private Sub Frame3_Click()

End Sub

Private Sub HFP_Change()
Me.Hfp.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Hfp.Value)
End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label23_Click()

End Sub

Private Sub Label33_Click()

End Sub

Private Sub Label47_Click()

End Sub

Private Sub MaxPresion_Change()
Me.MaxPresion.Value = Workbooks("RegisterU2DF7.xlam").SoloNumero(Me.MaxPresion.Value)
End Sub



Private Sub PECALCULAR_Click()
    Dim Coefp, DI, nmi, Sem, q, Pendien As Double
    Dim Longp, Hfp, F, C As Double
    Dim area, velp, Rey, fdw As Double
    
    If PESE.Text = "" Or PEQE.Text = "" Or PEPE.Text = "" Or Pendiente.Text = "" Then
        MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
    ElseIf PESE.Text = 0 Or PEQE.Text = 0 Or PEPE.Text = 0 Then
        MsgBox "Ningun valor debe ser igual a cero", vbCritical, "HF Riego Dice:"
    Else
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("A10:H500").Value = ""
        
        Coefp = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E1").Value
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C32").Value = PEDN.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C29").Value = PEQE.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C30").Value = PESE.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C31").Value = PEPE.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C34").Value = PEPS.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C35").Value = Pendiente.Text
        DI = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C33").Value
        
        Sem = (PESE.Text) * 1
        Pendien = (Pendiente.Text) * 1
        If ComboBox1.Text = "lph" Then
            q = (PEQE.Text)
        ElseIf ComboBox1.Text = "lps" Then
            q = (PEQE.Text) * 3600
        Else
            q = (PEQE.Text) * 3600 * 1000
        End If
        C = (PEPE.Text) * 1
        If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 1 Then
            nmi = 1.852
           
            For i = 1 To C
                'Celda Inicial en la hoja de Excel
                Celda = "A" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = i
                If PEPS.Text = "S0=S" Then
                    F = 1 / (nmi + 1) + 1 / (2 * i) + (nmi - 1) ^ 0.5 / (6 * (i) ^ 2) 'calculo del Factor de salidas Multiples
                    Longp = i * Sem
                ElseIf PEPS.Text = "S0=S/2" And i <= C - 1 Then
                    F = 1 / (nmi + 1) + 1 / (2 * i) + (nmi - 1) ^ 0.5 / (6 * (i) ^ 2)
                    Longp = i * Sem
                Else
                    F = (2 * i / (2 * i - 1)) * (1 / (nmi + 1) + (Sqr(nmi - 1)) / (6 * (i) ^ 2))
                    Longp = i * Sem - Sem / 2
                End If
                Celda = "B" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Longp
                Celda = "C" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = q * i
                Celda = "D" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = F
                Celda = "E" & i + 9
                Hfp = 10.648 * 1 / (Coefp) ^ (1.852) * (q / (3600000) * i) ^ (1.852) / (DI) ^ (4.871) * Longp * F
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Coefp
                Celda = "F" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Hfp
                Celda = "G" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Pendien / 100 * Longp
                Celda = "H" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Hfp + Pendien / 100 * Longp
            Next i
            
    
    
        ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 2 Then
            nmi = 2
           For i = 1 To C
                'Celda Inicial en la hoja de Excel
                Celda = "A" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = i
                If PEPS.Text = "S0=S" Then
                    F = 1 / (nmi + 1) + 1 / (2 * i) + (nmi - 1) ^ 0.5 / (6 * (i) ^ 2) 'calculo del Factor de salidas Multiples
                    Longp = i * Sem
                ElseIf PEPS.Text = "S0=S/2" And i <= C - 1 Then
                    F = 1 / (nmi + 1) + 1 / (2 * i) + (nmi - 1) ^ 0.5 / (6 * (i) ^ 2)
                    Longp = i * Sem
                Else
                    F = (2 * i / (2 * i - 1)) * (1 / (nmi + 1) + (Sqr(nmi - 1)) / (6 * (i) ^ 2))
                    Longp = i * Sem - Sem / 2
                End If

                Celda = "B" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Longp
                Celda = "C" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = q * i
                Celda = "D" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = F
                Celda = "E" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Coefp
                Celda = "F" & i + 9
                Hfp = 10.3 * (Coefp) ^ (2) * (q / (3600000) * i) ^ (2) / (DI) ^ (16 / 3) * Longp * F
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Hfp
                Celda = "G" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Pendien / 100 * Longp
                Celda = "H" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Hfp + Pendien / 100 * Longp
            Next i
        
        
        ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 3 Then
            nmi = 1.9
            
            For i = 1 To C
                'Celda Inicial en la hoja de Excel
                Celda = "A" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = i
                If PEPS.Text = "S0=S" Then
                    F = 1 / (nmi + 1) + 1 / (2 * i) + (nmi - 1) ^ 0.5 / (6 * (i) ^ 2) 'calculo del Factor de salidas Multiples
                    Longp = i * Sem
                ElseIf PEPS.Text = "S0=S/2" And i <= C - 1 Then
                    F = 1 / (nmi + 1) + 1 / (2 * i) + (nmi - 1) ^ 0.5 / (6 * (i) ^ 2)
                    Longp = i * Sem
                Else
                    F = (2 * i / (2 * i - 1)) * (1 / (nmi + 1) + (Sqr(nmi - 1)) / (6 * (i) ^ 2))
                    Longp = i * Sem - Sem / 2
                End If
                Celda = "B" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Longp
                Celda = "C" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = q * i
                Celda = "D" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = F
                Celda = "E" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Coefp
                Celda = "F" & i + 9
                Hfp = 0.004098 * Coefp * (q / (3600000) * i) ^ (1.9) / (DI) ^ (4.9) * Longp * F
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Hfp
                Celda = "G" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Pendien / 100 * Longp
                Celda = "H" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Hfp + Pendien / 100 * Longp
            Next i
         
         
         ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B1").Value = 4 Then '//Darcy Weisbach
            nmi = 2
            
            For i = 1 To C
                'Celda Inicial en la hoja de Excel
                Celda = "A" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = i
                If PEPS.Text = "S0=S" Then
                    F = 1 / (nmi + 1) + 1 / (2 * i) + (nmi - 1) ^ 0.5 / (6 * (i) ^ 2) 'calculo del Factor de salidas Multiples
                    Longp = i * Sem
                ElseIf PEPS.Text = "S0=S/2" And i = 1 Then
                    F = 1 / (nmi + 1) + 1 / (2 * i) + (nmi - 1) ^ 0.5 / (6 * (i) ^ 2) 'calculo del Factor de salidas Multiples
                    Longp = i * Sem
                ElseIf PEPS.Text = "S0=S/2" And i <= C - 1 Then
                    F = 1 / (nmi + 1) + 1 / (2 * i) + (nmi - 1) ^ 0.5 / (6 * (i) ^ 2)
                    Longp = i * Sem
                Else
                    F = (2 * i / (2 * i - 1)) * (1 / (nmi + 1) + (Sqr(nmi - 1)) / (6 * (i) ^ 2))
                    Longp = i * Sem - Sem / 2
                End If
                Celda = "B" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Longp
                Celda = "C" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = q * i
                Celda = "D" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = F
                Celda = "E" & i + 9
                
                
                Rey = Workbooks("RegisterU2DF7.xlam").NReynoldsP(q * i / 3600, DI * 1000) * 1
                If Rey <= 2000 Then
                        fdw = 64 / Rey
                Else
                    
                    If Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("E2").Value = 0 Then
                        fdw = Workbooks("RegisterU2DF7.xlam").CoeFriccionDWP(Rey, Coefp, DI * 1000)
                    Else
                        fdw = Workbooks("RegisterU2DF7.xlam").CoeFriccionSJ(Rey, Coefp, DI * 1000)
                    End If
                End If
                
                Hfp = 0.0827 * fdw * (q / (3600000) * i) ^ (2) / (DI) ^ (5) * Longp * F
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = fdw
                Celda = "F" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Hfp
                Celda = "G" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Pendien / 100 * Longp
                Celda = "H" & i + 9
                Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range(Celda).Value = Hfp + Pendien / 100 * Longp
            Next i
            
        End If
        

        VPresion.Text = FormatNumber(CDbl(Hfp + (Pendien / 100 * Longp)), 3)
        Dtotal.Text = FormatNumber(CDbl((Pendien / 100 * Longp)), 3)
        
        PEPerdida.Text = FormatNumber(CDbl(Hfp), 3)
        PELongitud.Text = FormatNumber(CDbl(Longp), 1)
        PEFSalidas.Text = FormatNumber(CDbl(F), 3)
        
        If ComboBox1.Text = "lph" Then
            QTot.Text = FormatNumber(CDbl(C * PEQE.Text / 3600), 3)
        ElseIf ComboBox1.Text = "lps" Then
            QTot.Text = FormatNumber(CDbl(C * PEQE.Text), 3)
        Else
            QTot.Text = FormatNumber(CDbl(C * PEQE.Text * 1000), 3)
        End If
        
        'UnidadQ.Caption = ComboBox1.Text
        'Workbooks("RegisterU2DF7.xlam").Save

    
    End If
End Sub

Private Sub PEDN_Change()
PEDN.Style = fmStyleDropDownList
End Sub

Private Sub PEITERRAR_Click()
    If PESE.Text = "" Or PEQE.Text = "" Or PEPE.Text = "" Or PEPerdida = "" Or PELongitud = "" Or PEFSalidas = "" Then
        MsgBox "Debe realizar un calculo antes de exportar a excel las iteraciones", vbCritical, "HF Riego Dice:"
    Else
        '3.- Importamos la hoja de Excel del complemento
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("B3").Value = PEQE.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("C3").Value = ComboBox1.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("B4").Value = PESE.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("B5").Value = PEDN.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("B6").Value = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C33").Value
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("B7").Value = PEPS.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("B8").Value = Pendiente.Text
        '4.-Exportamos los resultados generales a excel
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("F3").Value = VPresion.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("F4").Value = PEPerdida.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("F5").Value = Dtotal.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("F6").Value = QTot.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Range("F7").Value = PELongitud.Text
                
                
                    hojas = ActiveSheet.Name
                    Workbooks("RegisterU2DF7.xlam").Worksheets("RTuberiaSM").Copy _
                           after:=ActiveWorkbook.Sheets(hojas)
                          MsgBox "Se realizo con exito"
    End If
    
    

End Sub
Private Sub PELIMPIAR_Click()
PEQE.Text = ""
PESE.Text = ""
PEPE.Text = ""
PEPerdida.Text = ""
PELongitud.Text = ""
PEFSalidas.Text = ""
End Sub

Private Sub Pendiente_Change()
Me.Pendiente.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimalNega(Me.Pendiente.Value)
'Me.Pendiente.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Pendiente.Value)
End Sub

Private Sub PendienteTerreno_Change()
Me.PendienteTerreno.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimalNega(Me.PendienteTerreno.Value)
End Sub

Private Sub PEPE_Change()
Me.PEPE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumero(Me.PEPE.Value)
End Sub

Private Sub PEPS_Change()
PEPS.Style = fmStyleDropDownList
End Sub

Private Sub PEQE_Change()
Me.PEQE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.PEQE.Value)
End Sub
Private Sub PESE_Change()
Me.PESE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.PESE.Value)
End Sub

Private Sub PresionEmisor_Change()
Me.PresionEmisor.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.PresionEmisor.Value)
End Sub

Private Sub QE_Change()
Me.QE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.QE.Value)
End Sub
Private Sub SE_Change()
Me.SE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.SE.Value)
End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Initialize()
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A4").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A5").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A6").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A7").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A8").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
DN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value


ComboBoxSo.AddItem "S0=S"
ComboBoxSo.AddItem "S0=S/2"
'COMBOBOX DE LA OTRA VENTAJA
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A4").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A5").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A6").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A7").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A8").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
PEDN.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value

PEPS.AddItem "S0=S"
PEPS.AddItem "S0=S/2"
'Reconoce el ultimo calculo hecho
QE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B29").Value
SE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B30").Value
PresionEmisor.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B31").Value
DN.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B32").Value
ComboBoxSo.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B34").Value
MaxPresion.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F29").Value
PendienteTerreno.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("F30").Value
'RECONOCE LOS VALORES DE LA SEGUNDA VENTANA
PEQE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C29").Value
PESE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C30").Value
PEPE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C31").Value
PEDN.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C32").Value
PEPS.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C34").Value
Pendiente.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("C35").Value
'Iniciar con las unidades
ComboBox1.AddItem "lph"
ComboBox1.AddItem "lps"
ComboBox1.AddItem "m3/s"
ComboBox1.Text = "lph"
'Iniciar con las unidades longitud maxima
ComboBox2.AddItem "lph"
ComboBox2.AddItem "lps"
ComboBox2.AddItem "m3/s"
ComboBox2.Text = "lph"

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Workbooks("RegisterU2DF7.xlam").ClicDerecho
End Sub

Private Sub UserForm_Terminate()
Workbooks("RegisterU2DF7.xlam").Save
End Sub

Private Sub VPresión_Change()

End Sub

Private Sub VPresion_Change()

End Sub

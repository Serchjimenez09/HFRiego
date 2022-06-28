VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Agronomico 
   Caption         =   "Diseño Agronómico de sistemas de riego localizados"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7995
   OleObjectBlob   =   "Agronomico.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Agronomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
If DaSu.Text = "" Or DaTD.Text = "" Or DAQE.Text = "" Or DaSR.Text = "" Or DaQD.Text = "" Or DaET.Text = "" Or DaSE.Text = "" Then
    MsgBox "Primero, debe realizar un cálculo", vbCritical, "HF Riego Dice:"
ElseIf DaAM.Text = "" Or DaLH.Text = "" Or DaLr.Text = "" Or DaQN.Text = "" Then
    MsgBox "Primero, debe realizar un cálculo", vbCritical, "HF Riego Dice:"
Else
    'Ponemos Datos de entrada
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B4").Value = DaTR.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B5").Value = DaSu.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B6").Value = DaQD.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B7").Value = DaTD.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B8").Value = DaET.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B9").Value = DAQE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B10").Value = DaSE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B11").Value = DaSR.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B12").Value = Porcentaje.Text
    If DaDL.Value = True Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B13").Value = "SI"
    Else
        Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B13").Value = "NO"
    End If
    
    'Exportamos los Resultados
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B15").Value = DaAM.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B16").Value = DaLH.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B17").Value = DaLr.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B18").Value = DaQN.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B19").Value = DaSM.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B20").Value = DaQT.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B21").Value = DaNM.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B22").Value = DaSC.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B23").Value = DaQS.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B24").Value = DaTS.Text
    '3.- Importamos la hoja de Excel del complemento
                hojas = ActiveSheet.Name
                Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Copy _
                       after:=ActiveWorkbook.Sheets(hojas)
                      MsgBox "El archivo se exporto con éxito a Excel"
    End If

End Sub

Private Sub CommandButton2_Click()
HelpAgronomico.Show
End Sub

Private Sub DaBorrar_Click()
'Borrar Datos de entrada
DaSu.Text = ""
DaTD.Text = ""
DAQE.Text = ""
DaSR.Text = ""
DaQD.Text = ""
DaET.Text = ""
DaSE.Text = ""
'Borrar Resultados
DaAM.Text = ""
DaQN.Text = ""
DaQT.Text = ""
DaLr.Text = ""
DaSC.Text = ""
DaTS.Text = ""
DaLH.Text = ""
DaSM.Text = ""
DaNM.Text = ""
DaQS.Text = ""
End Sub

Private Sub DaCalcular_Click()
Dim sup, Td, QE, sl, Qd, Eto, SE, Porce, ai As Double
Dim am, eficiencia, lr, lh, qha, smax, smaxt, smax2, smaxt2, Qt, maxS, mimS, numS, SupS, Qs, Ts As Double

If DaSu.Text = "" Or DaTD.Text = "" Or DAQE.Text = "" Or DaSR.Text = "" Or DaQD.Text = "" Or DaET.Text = "" Or DaSE.Text = "" Or Porcentaje.Text = "" Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
ElseIf DaSu.Text <= 0 Or DaTD.Text <= 0 Or DAQE.Text <= 0 Or DaSR.Text <= 0 Or DaQD.Text <= 0 Or DaET.Text <= 0 Or DaSE.Text <= 0 Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
ElseIf Porcentaje.Text <= 3 Or Porcentaje.Text > 100 Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
Else
    
    'Es para que el tiempo disponible de riego sea menor a 24 hr
    If DaTD >= 24 Then
        DaTD.Text = 22
    Else
    End If
    
    If DaTR.Text = "Goteo" Then
        eficiencia = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B14").Value
    Else
        eficiencia = Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B15").Value
    End If
    
    'Colocando las variables a los campo
    sup = DaSu.Text * 1
    Td = DaTD.Text * 1
    QE = DAQE.Text * 1
    sl = DaSR.Text * 1
    Qd = DaQD.Text * 1
    Eto = DaET.Text * 1
    SE = DaSE.Text * 1
    Porce = Porcentaje.Text / 100
    
    If DaDL.Value = True Then
        sl = sl / 2
    Else
        sl = sl
    End If
    
    ai = sl * SE                    'Area de influencia del emisor
    am = ai * Porce                 ' Area de mojado del emisor
    lh = QE / am                    ' Lamina horaria o velocidad de aplicacion
    qha = lh * 10 / 3.6             'Gasto total del sistema por ha
    
    Qt = (lh * sup) / 0.36          ' Gasto total de todo el sistema
    minS = (Qt / Qd) 'ns2

    lr = Eto / (eficiencia / 100)   ' Lamina bruta de riego
    Ts = lr / lh                    ' Tiempo de riego
    maxS = Int(Td / Ts) 'ns1
    maxS2 = Td / Ts
    smax = Qd / qha * minS          ' superficie maxima con ese caudal
    smaxt = Qd / qha * maxS        ' superficie maxima con ese tiempo de riego
            
    
    
    If maxS = 0 Then
        MsgBox "El tiempo de riego no es suficiente para regar toda la superficie." & vbNewLine & "Solo se pueden regar " & FormatNumber(CDbl(Td / Ts), 4) & " ha con ese tiempo", vbCritical
        Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B26").Value = "El tiempo de riego no es suficiente para regar toda la superficie." & vbNewLine & "Solo se pueden regar " & FormatNumber(CDbl(Td / Ts), 4) & " ha con ese tiempo"
        'smax = Qd / ((lh * 1) / 0.36)
        'smaxt = Qd / qha * maxS
        'DaAM.Text = FormatNumber(CDbl(am), 4)
        'DaLH.Text = FormatNumber(CDbl(lh), 5)
        'DaLr.Text = FormatNumber(CDbl(lr), 5)
        'DaQN.Text = FormatNumber(CDbl(qha), 5)
        DaSM.Text = FormatNumber(CDbl(smax), 5)
        DaQT.Text = FormatNumber(CDbl(Qt), 5)
        DaNM.Text = 1
        DaSC.Text = FormatNumber(CDbl(Td / Ts), 5)
        DaQS.Text = ((lh * Td / Ts) / 0.36)
        DaTS.Text = FormatNumber(CDbl(Ts), 5)
        
    ElseIf smax < sup Then
        numS = minS
        MsgBox "El gasto disponible no es suficiente para regar toda la superficie." & vbNewLine & "Solo se pueden regar " & FormatNumber(CDbl(smax), 4) & " ha con ese caudal" & vbNewLine & "Para " & FormatNumber(CDbl(sup), 4) & " ha, con el caudal y tiempo de riego disponible, el emisor y su arreglo, se necesitan: " & FormatNumber(CDbl(Qt / maxS), 4) & " lps", vbCritical
        Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B26").Value = "El gasto disponible no es suficiente para regar toda la superficie." & vbNewLine & "Solo se pueden regar " & FormatNumber(CDbl(smax), 4) & " ha con ese caudal" & vbNewLine & "Para " & FormatNumber(CDbl(sup), 4) & " ha, con el caudal y tiempo de riego disponible, el emisor y su arreglo, se necesitan: " & FormatNumber(CDbl(Qt / maxS), 4) & " lps"
        'smax = Qd / ((lh * 1) / 0.36)
        'smaxt = Qd / qha * maxS
        'DaAM.Text = FormatNumber(CDbl(am), 4)
        'DaLH.Text = FormatNumber(CDbl(lh), 5)
        'DaLr.Text = FormatNumber(CDbl(lr), 5)
        'DaQN.Text = FormatNumber(CDbl(qha), 5)
        DaSM.Text = FormatNumber(CDbl(smax), 5)
       
        DaQT.Text = FormatNumber(CDbl(Qt), 5)
        DaNM.Text = ""
        DaSC.Text = ""
        DaQS.Text = ""
        DaTS.Text = FormatNumber(CDbl(Ts), 5)
    ElseIf smaxt < sup Then
        numS = minS
        'smax = Qd / ((lh * 1) / 0.36)
        
        MsgBox "El tiempo disponible no es suficiente para regar toda la superficie." & vbNewLine & "Solo se pueden regar " & FormatNumber(CDbl(smaxt), 4) & " ha con ese tiempo" & vbNewLine & "" & vbNewLine & "Para " & FormatNumber(CDbl(sup), 4) & " ha, con el tiempo y caudal de riego disponible, el emisor y su arreglo, se necesitan aumentar el caudal a: " & FormatNumber(CDbl(Qt / maxS), 4) & " lps, o dar riegos deficitarios reduciendo la lamina ", vbCritical
        Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B26").Value = "El tiempo disponible no es suficiente para regar toda la superficie." & vbNewLine & "Solo se pueden regar " & FormatNumber(CDbl(smaxt), 4) & " ha con ese tiempo" & vbNewLine & "" & vbNewLine & "Para " & FormatNumber(CDbl(sup), 4) & " ha, con el tiempo y caudal de riego disponible, el emisor y su arreglo, se necesitan aumentar el caudal a: " & FormatNumber(CDbl(Qt / maxS), 4) & " lps, o dar riegos deficitarios reduciendo la lamina "
        'DaAM.Text = FormatNumber(CDbl(am), 4)
        'DaLH.Text = FormatNumber(CDbl(lh), 5)
        'DaLr.Text = FormatNumber(CDbl(lr), 5)
        'DaQN.Text = FormatNumber(CDbl(qha), 5)
        DaSM.Text = FormatNumber(CDbl(smax), 5)
        
        DaQT.Text = FormatNumber(CDbl(Qt), 5)
        DaNM.Text = ""
        DaSC.Text = ""
        DaQS.Text = ""
        DaTS.Text = FormatNumber(CDbl(Ts), 5)
        
    Else
        numS = Round(maxS)
        
        Workbooks("RegisterU2DF7.xlam").Worksheets("RAgronomico").Range("B26").Value = "El gasto Disponible es suficiente para regar toda la superficie de riego"
        SupS = sup / numS
        Qs = (lh * SupS) / 0.36

        DaSM.Text = FormatNumber(CDbl(smax), 3)
       
        DaQT.Text = FormatNumber(CDbl(Qt), 3)
        DaNM.Text = FormatNumber(CDbl(numS), 3)
        DaSC.Text = FormatNumber(CDbl(SupS), 3)
        DaQS.Text = FormatNumber(CDbl(Qs), 3)
        DaTS.Text = FormatNumber(CDbl(Ts), 3)
        DaLn.Text = FormatNumber(CDbl(Eto), 3)
    End If
        
        DaAM.Text = FormatNumber(CDbl(ai), 3)
        AreaHumedecida.Text = FormatNumber(CDbl(am), 3)
        DaLH.Text = FormatNumber(CDbl(lh), 3)
        DaLr.Text = FormatNumber(CDbl(lr), 3)
        DaQN.Text = FormatNumber(CDbl(qha), 3)
        DaLn.Text = FormatNumber(CDbl(Eto), 3)
        smax2 = Qd / qha * minS         ' superficie maxima con ese caudal
        smaxt2 = Qd / qha * maxS
        DaSM.Text = FormatNumber(CDbl(smaxt2), 3)
        
        
        
    Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B1").Value = DaTR.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B2").Value = DaSu.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B4").Value = DaTD.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B6").Value = DAQE.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B8").Value = DaSR.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B3").Value = DaQD.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B5").Value = DaET.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B7").Value = DaSE.Text
    If DaDL.Value = True Then
        Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B9").Value = 1
    Else
        Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B9").Value = 0
    End If
    

End If
End Sub

Private Sub DaET_Change()
    Me.DaET.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DaET.Value)
End Sub

Private Sub DaQD_Change()
    Me.DaQD.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DaQD.Value)
End Sub

Private Sub DAQE_Change()
    Me.DAQE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DAQE.Value)
End Sub

Private Sub DaSE_Change()
    Me.DaSE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DaSE.Value)
End Sub

Private Sub DaSR_Change()
    Me.DaSR.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DaSR.Value)
End Sub

Private Sub DaSu_Change()
    Me.DaSu.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DaSu.Value)
End Sub

Private Sub DaTD_Change()
    Me.DaTD.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.DaTD.Value)
    
End Sub


Private Sub DaTR_Change()
DaTR.Style = fmStyleDropDownList
If DaTR.Value = "Microaspersión" Then
Label22.Caption = "Díametro de Mojado del microaspersor"
Else
Label22.Caption = "Porcentaje de mojado del emisor(%):"
End If

End Sub

Private Sub Label18_Click()

End Sub

Private Sub Label23_Click()

End Sub

Private Sub UserForm_Initialize()
DaTR.AddItem "Goteo"
DaTR.AddItem "Microaspersión"
DaTR.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B1").Value
DaSu.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B2").Value
DaTD.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B4").Value
DAQE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B6").Value
DaSR.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B8").Value
DaQD.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B3").Value
DaET.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B5").Value
DaSE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B7").Value
Porcentaje.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B10").Value
If Workbooks("RegisterU2DF7.xlam").Worksheets("Agronomico").Range("B9").Value = 1 Then
    DaDL.Value = True
Else
    DaDL.Value = False
End If


End Sub

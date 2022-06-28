VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PreEfectiva 
   Caption         =   "Precipitación Efectiva"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7395
   OleObjectBlob   =   "PreEfectiva.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PreEfectiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox2_Change()
ComboBox2.Style = fmStyleDropDownList
End Sub

Private Sub Mes_Change()
Mes.Style = fmStyleDropDownList
End Sub
Private Sub PECALCULAR_Click()

Dim DPENE, penel, EmpiA, EmpiB, EmpiC, EmpiD, EmpiE, Porcentaje As Double
    
    If PENE.Text = "" Then
            MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
            PEENE.Text = ""
    Else
        
        EmpiA = (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B6").Value) * 1
        EmpiB = (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("E6").Value) * 1
        EmpiC = (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B7").Value) * 1
        EmpiD = (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("E7").Value) * 1
        EmpiE = (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("H6").Value) * 1
        Porcentaje = (Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B9").Value) * 1
        DPENE = PENE.Text * 1
        
        If Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "Porcentaje fijo" Then
            pene1 = DPENE * Porcentaje / 100
        ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "Precipitacion Confiable" Then
                If DPENE <= 70 Then
                    pene1 = 0.6 * DPENE - 10

                Else
                    pene1 = 0.8 * DPENE - 24
                End If
        ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "Formula empirica" Then
            If DPENE <= EmpiE Then
                pene1 = EmpiA * DPENE + EmpiB * 1
                                    
            Else
                pene1 = EmpiC * DPENE + EmpiD * 1
            End If
        ElseIf Workbooks("RegisterU2DF7.xlam").Worksheets("PE").Range("B3").Value = "USDA" Then
            If PENE <= 250 Then
                pene1 = (PENE * (125 - 0.2 * PENE)) / 125
            Else
                pene1 = 0.1 * DPENE + 125
            End If
        End If
    
        If pene1 <= 0 Then
            PEENE.Text = 0
        Else
            PEENE.Text = FormatNumber(CDbl(pene1), 3)
        End If
        
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B59").Value = Mes.Text
        Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B60").Value = PENE.Text
        'Workbooks("RegisterU2DF7.xlam").Save
    End If

End Sub

Private Sub PELIMPIAR_Click()
    If PENE.Text = "" Or PEENE = "" Then
            MsgBox "Primero, debe realizar un cálculo", vbCritical, "HF Riego Dice:"

    Else
        ListBoxPE.AddItem " " & FormatNumber(CDbl(PEconta.Caption), 0) & "                        " & Mes.Text & "                    " & FormatNumber(CDbl(PENE.Text), 2) & "                       " & FormatNumber(CDbl(PEENE.Text), 2)
        PET.Caption = FormatNumber(CDbl(PET.Caption + (PEENE.Text) * 1), 0)
        PTotal.Caption = FormatNumber(CDbl(PTotal + (PENE.Text) * 1), 0)
        
        'Celda Inicial en la hoja de Excel
        Celda = "A" & PEconta.Caption + 10
        Workbooks("RegisterU2DF7.xlam").Worksheets("RPE").Range(Celda).Value = PEconta.Caption
        Celda = "B" & PEconta.Caption + 10
        Workbooks("RegisterU2DF7.xlam").Worksheets("RPE").Range(Celda).Value = Mes.Text
        Celda = "C" & PEconta.Caption + 10
        Workbooks("RegisterU2DF7.xlam").Worksheets("RPE").Range(Celda).Value = PENE.Text
        Celda = "D" & PEconta.Caption + 10
        Workbooks("RegisterU2DF7.xlam").Worksheets("RPE").Range(Celda).Value = PEENE.Text
        PEconta.Caption = PEconta.Caption + 1
        
    End If
End Sub

Private Sub PENE_Change()
Me.PENE.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.PENE.Value)
End Sub

Private Sub PEsalir_Click()
    If ListBoxPE.ListCount < 2 Then
        MsgBox " No hay suficientes valores para exportar ", vbCritical, "HF Riego Dice:"
    Else
        Workbooks("RegisterU2DF7.xlam").Worksheets("RPE").Range("C10").Value = PTotal.Caption
        Workbooks("RegisterU2DF7.xlam").Worksheets("RPE").Range("D10").Value = PET.Caption
    
        '3.- Importamos la hoja de Excel del complemento
                    hojas = ActiveSheet.Name
                    Workbooks("RegisterU2DF7.xlam").Worksheets("RPE").Copy _
                           after:=ActiveWorkbook.Sheets(hojas)
                          MsgBox "Se realizo con exito"
    End If
End Sub

Private Sub UserForm_Initialize()
Mes.AddItem "Enero"
Mes.AddItem "Febrero"
Mes.AddItem "Marzo"
Mes.AddItem "Abril"
Mes.AddItem "Mayo"
Mes.AddItem "Junio"
Mes.AddItem "Julio"
Mes.AddItem "Agosto"
Mes.AddItem "Septiembre"
Mes.AddItem "Octubre"
Mes.AddItem "Noviembre"
Mes.AddItem "Diciembre"

Mes.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B59").Value
PENE.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("B60").Value
ListBoxPE.AddItem " # " & "         |            " & "Mes" & "             |     " & "Precipitación (mm)" & "     |     " & "P. Efectiva (mm)"

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Workbooks("RegisterU2DF7.xlam").ClicDerecho
End Sub

Private Sub UserForm_Terminate()
Workbooks("RegisterU2DF7.xlam").Worksheets("RPE").Range("A11:D40").Value = ""
    Workbooks("RegisterU2DF7.xlam").Save
End Sub

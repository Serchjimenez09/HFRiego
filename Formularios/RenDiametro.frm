VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenDiametro 
   Caption         =   "Renombrar Diámetros Nominales"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6450
   OleObjectBlob   =   "RenDiametro.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "RenDiametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Rd1_Change()
Me.Rd1.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd1.Value)
End Sub

Private Sub Rd10_Change()
Me.Rd10.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd10.Value)
End Sub

Private Sub Rd11_Change()
Me.Rd11.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd11.Value)
End Sub

Private Sub Rd12_Change()
Me.Rd12.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd12.Value)
End Sub

Private Sub Rd13_Change()
Me.Rd13.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd13.Value)
End Sub

Private Sub Rd14_Change()
Me.Rd14.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd14.Value)
End Sub

Private Sub Rd15_Change()
Me.Rd15.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd15.Value)
End Sub

Private Sub Rd16_Change()
Me.Rd16.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd16.Value)
End Sub

Private Sub Rd2_Change()
Me.Rd2.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd2.Value)
End Sub

Private Sub Rd3_Change()
Me.Rd3.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd3.Value)
End Sub

Private Sub Rd4_Change()
Me.Rd4.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd4.Value)
End Sub

Private Sub Rd5_Change()
Me.Rd5.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd5.Value)
End Sub

Private Sub Rd6_Change()
Me.Rd6.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd6.Value)
End Sub

Private Sub Rd7_Change()
Me.Rd7.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd7.Value)
End Sub

Private Sub Rd8_Change()
Me.Rd8.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd8.Value)
End Sub

Private Sub Rd9_Change()
Me.Rd9.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.Rd9.Value)
End Sub

Private Sub RdAceptar_Click()
If Rd1.Text = "" Or Rd2.Text = "" Or Rd3.Text = "" Or Rd4.Text = "" Or Rd5.Text = "" Or Rd6.Text = "" Or Rd7.Text = "" Or Rd8.Text = "" Or Rd9.Text = "" Or Rd10.Text = "" Or Rd11.Text = "" Or Rd12.Text = "" Or Rd13.Text = "" Or Rd14.Text = "" Or Rd15.Text = "" Or Rd16.Text = "" Then
    MsgBox "HF Riego Dice:" & vbNewLine & "Faltan datos: Debe ingresar todos los datos", vbCritical, "Error"
ElseIf Rd1.Text = 0 Or Rd2.Text = 0 Or Rd3.Text = 0 Or Rd4.Text = 0 Or Rd5.Text = 0 Or Rd6.Text = 0 Or Rd7.Text = 0 Or Rd8.Text = 0 Or Rd9.Text = 0 Or Rd10.Text = 0 Or Rd11.Text = 0 Or Rd12.Text = 0 Or Rd13.Text = 0 Or Rd14.Text = 0 Or Rd15.Text = 0 Or Rd16.Text = 0 Then
    MsgBox "HF Riego Dice:" & vbNewLine & "Faltan datos: Debe ingresar todos los datos", vbCritical, "Error"
Else
    Dim DN(16), temp, cambio As Double
    DN(1) = Rd1.Text * 1
    DN(2) = Rd2.Text * 1
    DN(3) = Rd3.Text * 1
    DN(4) = Rd4.Text * 1
    DN(5) = Rd5.Text * 1
    DN(6) = Rd6.Text * 1
    DN(7) = Rd7.Text * 1
    DN(8) = Rd8.Text * 1
    DN(9) = Rd9.Text * 1
    DN(10) = Rd10.Text * 1
    DN(11) = Rd11.Text * 1
    DN(12) = Rd12.Text * 1
    DN(13) = Rd13.Text * 1
    DN(14) = Rd14.Text * 1
    DN(15) = Rd15.Text * 1
    DN(16) = Rd16.Text * 1
    'codigo para ordenar vector
    cambio = 0
    For i = 1 To 15
        For J = 1 To 15
            If DN(J) > DN(J + 1) Then
                tmp = DN(J)
                DN(J) = DN(J + 1)
                DN(J + 1) = tmp
                cambio = 1
            Else
            End If
        Next
    Next
    'MsgBox "HF Riego Dice:" & vbNewLine & dn(14) & cambio, , "Error"
    If cambio = 1 Then
        MsgBox "HF Riego Dice:" & vbNewLine & "Se ordenarón los diámetros nominales de menor a mayor", , "¡¡¡ATENCIÓN!!!"
    Else
    End If
    Rd1.Text = DN(1)
    Rd2.Text = DN(2)
    Rd3.Text = DN(3)
    Rd4.Text = DN(4)
    Rd5.Text = DN(5)
    Rd6.Text = DN(6)
    Rd7.Text = DN(7)
    Rd8.Text = DN(8)
    Rd9.Text = DN(9)
    Rd10.Text = DN(10)
    Rd11.Text = DN(11)
    Rd12.Text = DN(12)
    Rd13.Text = DN(13)
    Rd14.Text = DN(14)
    Rd15.Text = DN(15)
    Rd16.Text = DN(16)
    Ajustes.Ad1.Caption = Rd1.Text
    Ajustes.Ad2.Caption = Rd2.Text
    Ajustes.Ad3.Caption = Rd3.Text
    Ajustes.Ad4.Caption = Rd4.Text
    Ajustes.Ad5.Caption = Rd5.Text
    Ajustes.Ad6.Caption = Rd6.Text
    Ajustes.Ad7.Caption = Rd7.Text
    Ajustes.Ad8.Caption = Rd8.Text
    Ajustes.Ad9.Caption = Rd9.Text
    Ajustes.Ad10.Caption = Rd10.Text
    Ajustes.Ad11.Caption = Rd11.Text
    Ajustes.Ad12.Caption = Rd12.Text
    Ajustes.Ad13.Caption = Rd13.Text
    Ajustes.Ad14.Caption = Rd14.Text
    Ajustes.Ad15.Caption = Rd15.Text
    Ajustes.Ad16.Caption = Rd16.Text
    'para los diametros de la zanja
    Ajustes.ZAd6.Caption = Rd6.Text
    Ajustes.ZAd7.Caption = Rd7.Text
    Ajustes.ZAd8.Caption = Rd8.Text
    Ajustes.ZAd9.Caption = Rd9.Text
    Ajustes.ZAd10.Caption = Rd10.Text
    Ajustes.ZAd11.Caption = Rd11.Text
    Ajustes.ZAd12.Caption = Rd12.Text
    Ajustes.ZAd13.Caption = Rd13.Text
    Ajustes.ZAd14.Caption = Rd14.Text
    Ajustes.ZAd15.Caption = Rd15.Text
    Ajustes.ZAd16.Caption = Rd16.Text
    
    'Ajustes.Cad1.Caption = Rd6.Text
    'Ajustes.Cad2.Caption = Rd7.Text
    'Ajustes.Cad3.Caption = Rd8.Text
    'Ajustes.Cad4.Caption = Rd9.Text
    'Ajustes.Cad5.Caption = Rd10.Text
    'Ajustes.Cad6.Caption = Rd11.Text
    'Ajustes.Cad7.Caption = Rd12.Text
    'Ajustes.Cad8.Caption = Rd13.Text
    'Ajustes.Cad9.Caption = Rd14.Text
    'Ajustes.Cad10.Caption = Rd15.Text
    'Ajustes.Cad11.Caption = Rd16.Text
    MsgBox "HF Riego Dice:" & vbNewLine & "Se actualizaron los valores, ahora solo" & vbNewLine & "actualiza los diámetros internos", , "¡¡¡ATENCIÓN!!!"
    
    'se pasan los diámetros a la hoja de excel
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A4").Value = Rd1.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A5").Value = Rd2.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A6").Value = Rd3.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A7").Value = Rd4.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A8").Value = Rd5.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value = Rd6.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value = Rd7.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value = Rd8.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value = Rd9.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value = Rd10.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value = Rd11.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value = Rd12.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value = Rd13.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value = Rd14.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value = Rd15.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value = Rd16.Text
    Workbooks("RegisterU2DF7.xlam").Save
End If
End Sub

Private Sub UserForm_Initialize()
Rd1.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A4").Value
Rd2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A5").Value
Rd3.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A6").Value
Rd4.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A7").Value
Rd5.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A8").Value
Rd6.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
Rd7.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
Rd8.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
Rd9.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
Rd10.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
Rd11.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
Rd12.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
Rd13.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
Rd14.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
Rd15.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
Rd16.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value
End Sub


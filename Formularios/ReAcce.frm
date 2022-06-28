VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReAcce 
   Caption         =   "Renombrar nombre de accesorios"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7320
   OleObjectBlob   =   "ReAcce.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ReAcce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub RaAceptar_Click()

If Ra1.Text = "" Or Ra2.Text = "" Or Ra3.Text = "" Or Ra4.Text = "" Or Ra5.Text = "" Or Ra6.Text = "" Or Ra7.Text = "" Or Ra8.Text = "" Or Ra9.Text = "" Or Ra10.Text = "" Or Ra11.Text = "" Or Ra12.Text = "" Or Ra13.Text = "" Or Ra14.Text = "" Or Ra15.Text = "" Or Ra16.Text = "" Then
    MsgBox "Faltan datos: Debe ingresar todos los datos", vbCritical, "HF Riego Dice:"
Else
    Ajustes.AcN1.Caption = Ra1.Text
    Ajustes.AcN2.Caption = Ra2.Text
    Ajustes.AcN3.Caption = Ra3.Text
    Ajustes.AcN4.Caption = Ra4.Text
    Ajustes.AcN5.Caption = Ra5.Text
    Ajustes.AcN6.Caption = Ra6.Text
    Ajustes.AcN7.Caption = Ra7.Text
    Ajustes.AcN8.Caption = Ra8.Text
    Ajustes.AcN9.Caption = Ra9.Text
    Ajustes.AcN10.Caption = Ra10.Text
    Ajustes.AcN11.Caption = Ra11.Text
    Ajustes.AcN12.Caption = Ra12.Text
    Ajustes.AcN13.Caption = Ra13.Text
    Ajustes.AcN14.Caption = Ra14.Text
    Ajustes.AcN15.Caption = Ra15.Text
    Ajustes.AcN16.Caption = Ra16.Text

    MsgBox "Se actualizaron los valores, ahora solo" & vbNewLine & "actualiza los coeficientes", vbExclamation, "HF Riego Dice:"
    
    'se pasan los diámetros a la hoja de excel
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B2").Value = Ra1.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B3").Value = Ra2.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B4").Value = Ra3.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B5").Value = Ra4.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B6").Value = Ra5.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B7").Value = Ra6.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B8").Value = Ra7.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B9").Value = Ra8.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B10").Value = Ra9.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B11").Value = Ra10.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B12").Value = Ra11.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B13").Value = Ra12.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B14").Value = Ra13.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B15").Value = Ra14.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B16").Value = Ra15.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B17").Value = Ra16.Text
    Workbooks("RegisterU2DF7.xlam").Save
End If
End Sub

Private Sub UserForm_Initialize()
Ra1.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B2").Value
Ra2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B3").Value
Ra3.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B4").Value
Ra4.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B5").Value
Ra5.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B6").Value
Ra6.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B7").Value
Ra7.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B8").Value
Ra8.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B9").Value
Ra9.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B10").Value
Ra10.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B11").Value
Ra11.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B12").Value
Ra12.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B13").Value
Ra13.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B14").Value
Ra14.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B15").Value
Ra15.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B16").Value
Ra16.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("Acce").Range("B17").Value
End Sub

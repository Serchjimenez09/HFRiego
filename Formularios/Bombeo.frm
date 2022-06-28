VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bombeo 
   Caption         =   "Cantidad de Cemento o Lubricante para PVC"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8370.001
   OleObjectBlob   =   "Bombeo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Bombeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
CTipo.AddItem "Con bocina o Cementado"
CTipo.AddItem "Con Campana o Anillo"

DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A4").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A5").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A6").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A7").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A8").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A9").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A10").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A11").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A12").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A13").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A14").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A15").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A16").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A17").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A18").Value
DCem.AddItem Workbooks("RegisterU2DF7.xlam").Worksheets("Metodo").Range("A19").Value

End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calcularlamina 
   Caption         =   "Lámina Horaria"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   OleObjectBlob   =   "Calcularlamina.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Calcularlamina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LH_calcular_Click()

Dim areamojada, SE, sl, QE As Double
If LH_qe.Text = "" Or LH_se.Text = "" Or LH_sl.Text = "" Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
ElseIf LH_qe.Text = 0 Or LH_se.Text = 0 Or LH_sl.Text = 0 Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
Else
    QE = LH_qe.Text
    SE = LH_se.Text
    If Dlinea.Value = True Then
        sl = LH_sl.Text / 2
    Else
        sl = LH_sl.Text
    End If
    areamojada = SE * sl
    LaminaH.Text = FormatNumber(CDbl(QE / areamojada), 3)
    
End If
End Sub
Private Sub LH_qe_Change()
Me.LH_qe.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.LH_qe.Value)
End Sub

Private Sub LH_se_Change()
Me.LH_se.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.LH_se.Value)
End Sub

Private Sub LH_sl_Change()
Me.LH_sl.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.LH_sl.Value)
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()
If LaminaH.Text <> "" Then
    Secundaria.Slh.Text = LaminaH.Text
    Secundaria.Sse.Text = LH_sl.Text
Else
End If
End Sub


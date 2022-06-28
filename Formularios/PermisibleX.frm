VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PermisibleX 
   Caption         =   "HF Permisible"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5190
   OleObjectBlob   =   "PermisibleX.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PermisibleX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim Hfp, pen, SERGIO As Double


If PPresion.Text = "" Or PVariacion.Text = "" Or PLongitud.Text = "" Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
ElseIf OptionButton1.Value = True And PPendiente.Text = "" Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
ElseIf OptionButton2.Value = True And PPendiente2.Text = "" Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"
ElseIf PPresion.Text = 0 Or PVariacion.Text = 0 Or PLongitud.Text = 0 Then
    MsgBox "Ningun valor debe ser igual a cero", vbCritical, "HF Riego Dice:"
ElseIf PPresion.Text > 100 Or PVariacion.Text > 50 Or PLongitud.Text < 10 Then
    MsgBox "Faltan datos o son irreales", vbCritical, "HF Riego Dice:"

    
Else
    If OptionButton1.Value = True Then
        If (PPendiente.Text > 50 Or PPendiente.Text < -50) Then
            MsgBox "Pendiente incorrecta", vbCritical, "HF Riego Dice:"
            SERGIO = 0
        Else
            pen = (PPendiente.Text / 100 * PLongitud)
            SERGIO = 1
        End If
    ElseIf OptionButton2.Value = True Then
        If (PPendiente2.Text > 50 Or PPendiente2.Text < -50) Then
            MsgBox "Pendiente incorrecta", vbCritical, "HF Riego Dice:"
            SERGIO = 0
        Else
            pen = -1 * PPendiente2.Text
            SERGIO = 1
        End If
    End If
    
    If SERGIO = 1 Then
        Hfp = PPresion * (PVariacion / 100) - pen
        If Hfp <= 0 Then
            MsgBox "Debe aumentar la maxima variación entre emisores o disminuir la pendiente", vbCritical, "HF Riego Dice:"
            PPermisible.Text = ""
        Else
            PPermisible.Text = FormatNumber(CDbl(Hfp), 4)
        End If
    Else
    End If
End If
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub CommandButton5_Click()
Pendientex.Show
End Sub



Private Sub PLongitud_Change()
Me.PLongitud.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.PLongitud.Value)
End Sub

Private Sub PPendiente_Change()
Me.PPendiente.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimalNega(Me.PPendiente.Value)
End Sub

Private Sub PPresion_Change()
Me.PPresion.Value = Workbooks("RegisterU2DF7.xlam").SoloNumeroDecimal(Me.PPresion.Value)
End Sub

Private Sub PVariacion_Change()
Me.PVariacion.Value = Workbooks("RegisterU2DF7.xlam").SoloNumero(Me.PVariacion.Value)
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()
If PPermisible.Text <> "" Then
    If Num.Caption = 1 Then
        Secundaria.Spe.Text = PPermisible.Text
        Secundaria.Slt.Text = PLongitud.Text
    Else
        Secundaria.PermisibleLateral.Text = PPermisible.Text
        Secundaria.LTotalLateral.Text = PLongitud.Text
    End If
Else
End If
End Sub

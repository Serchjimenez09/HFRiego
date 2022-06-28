VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AddinInstall()
   'Cualquier código de inicialización debe colocarse aquí
End Sub

Function SoloNumeroDecimal(Texto As Variant)
Dim Caracter As Variant
    Dim Largo As Integer
    On Error Resume Next
    Punto = 0
    Largo = Len(Texto)
    '
    For i = 1 To Largo
        Caracter = Mid(CStr(Texto), i, 1)
        If Caracter <> "" Then
            '
            If Caracter = Chr(46) Then
                Punto = Punto + 1
                If Punto > 1 Then
                    Texto = WorksheetFunction.Replace(Texto, i, 1, "")
                    SoloNumeroDecimal = Texto
                    Punto = 0
                End If
            Else
                If Caracter < Chr(48) Or Caracter > Chr(57) Then
                    Texto = Replace(Texto, Caracter, "")
                    SoloNumeroDecimal = Texto
                Else
                End If
                '
            End If
            '
        End If
    Next i
    '
    SoloNumeroDecimal = Texto
    On Error GoTo 0
    '
End Function
Function SoloTexto(Texto As Variant)
'
    Dim Caracter As Variant
    Dim Largo As String
    On Error Resume Next
    Largo = Len(Texto)
    '
    For i = 1 To Largo
        Caracter = CInt(Mid(Texto, i, 1))
        '
        If Caracter <> "" Then
            If Not Application.WorksheetFunction.IsText(Caracter) Then
                Texto = Replace(Texto, Caracter, "")
                SoloTexto = Texto
            Else
            End If
        End If
        '
    Next i
    '
    SoloTexto = Texto
    On Error GoTo 0
    '
End Function
Function SoloNumero(Texto As Variant)
'
    Dim Caracter As Variant
    Dim Largo As Integer
    On Error Resume Next
    Largo = Len(Texto)
    '
    For i = 1 To Largo
        Caracter = Mid(CStr(Texto), i, 1)
        '
        If Caracter <> "" Then
            If Caracter < Chr(48) Or Caracter > Chr(57) Then
                Texto = Replace(Texto, Caracter, "")
                SoloNumero = Texto
            Else
            End If

        End If
        '
    Next i
    '
    SoloNumero = Texto
    On Error GoTo 0
    '
End Function

Function SoloNumeroDecimalNega(Texto As Variant)
Dim Caracter As Variant
    Dim Largo As Integer
    On Error Resume Next
    Punto = 0
    Resta = 0
    Largo = Len(Texto)
    '
    For i = 1 To Largo
        Caracter = Mid(CStr(Texto), i, 1)
        If Caracter <> "" Then
            '
            If Caracter = Chr(45) Then
                Resta = Resta + 1
                If Resta > 1 Then
                    Texto = WorksheetFunction.Replace(Texto, i, 1, "")
                    SoloNumeroDecimalNega = Texto
                    Resta = 0
                End If
            ElseIf Caracter = Chr(46) Then
                Punto = Punto + 1
                If Punto > 1 Then
                    Texto = WorksheetFunction.Replace(Texto, i, 1, "")
                    SoloNumeroDecimalNega = Texto
                    Punto = 0
                End If
            Else
                If Caracter < Chr(48) Or Caracter > Chr(57) Then
                    Texto = Replace(Texto, Caracter, "")
                    SoloNumeroDecimalNega = Texto
                Else
                End If
                '
            End If
            '
        End If
    Next i
    '
    SoloNumeroDecimalNega = Texto
    On Error GoTo 0
    '
End Function
Function NReynoldsP(gasto, DI) As Double
Dim vel, v As Double
    v = 0.000001007
    vel = VelocidadFlujo(gasto, DI)
    NReynoldsP = (DI / 1000) * vel / v
End Function

Function CoeFriccionDWP(re, e, d) As Double
' proponemos valores de a y b
Dim a, b, C, F, fa, fb As Double
a = 0
b = 10
' Proponemos un valor de f inicial para que entre al ciclo
F = 1
Do While Abs(F) > 0.00000000001
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
    CoeFriccionDWP = C
End Function

Function CoeFriccionSJ(re, e, d) As Double
' proponemos valores de a y b
Dim a, b, C, F, fa, fb As Double
CoeFriccionSJ = 0.25 / ((Math.Log((e / d) / 3.7 + 5.74 / ((re) ^ 0.9))) * 0.434294481903252) ^ 2
End Function

Function ClicDerecho()
MsgBox "HF Riego fue Programado por:" & vbNewLine & "M.C. Sergio Iván Jiménez Jiménez. " & vbNewLine & vbNewLine & vbNewLine & "Derechos Reservados." & vbNewLine & vbNewLine & "hidraulica.facil@gmail.com", , "Acerca Del Complemento"

End Function
Function FSMultiples(N, m)
FSMultiples = 1 / (m + 1) + 1 / (2 * N) + (m - 1) ^ 0.5 / (6 * (N) ^ 2)
End Function




Private Sub Workbook_Open()
    LoadFunctionDescriptions
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    UnLoadFunctionDescriptions
End Sub

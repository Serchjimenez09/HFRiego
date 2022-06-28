VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KC 
   Caption         =   "Buscador de Coeficientes de Cultivo"
   ClientHeight    =   8760.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9300.001
   OleObjectBlob   =   "KC.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "KC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CCultivo_Change()
CCultivo.Style = fmStyleDropDownList

If CCultivo.ListIndex <> -1 Then

    Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("B3").Value = CCultivo.Text
    KCINI.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("B4").Value
    KCMED.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("B5").Value
    KCFIN.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("B6").Value
    ALT.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("B7").Value
    KcEI1.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("D1").Value
    KcED1.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("D2").Value
    KcEM1.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("D3").Value
    KcEN1.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("D4").Value
    KcET1.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("D5").Value
    KcEF1.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("D6").Value
    KcER1.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("D7").Value
    KcEI2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("E1").Value
    KcED2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("E2").Value
    KcEM2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("E3").Value
    KcEN2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("E4").Value
    KcET2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("E5").Value
    KcEF2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("E6").Value
    KcER2.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("E7").Value
    KcEI3.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("F1").Value
    KcED3.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("F2").Value
    KcEM3.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("F3").Value
    KcEN3.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("F4").Value
    KcET3.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("F5").Value
    KcEF3.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("F6").Value
    KcER3.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("F7").Value
Else
End If
End Sub

Private Sub CTipo_Change()
CTipo.Style = fmStyleDropDownList

If CTipo.Text = "a. Hortalizas Pequeñas" Then
    CCultivo.Clear
    CCultivo.AddItem "Brécol (Brócoli)"
    CCultivo.AddItem "Col de Bruselas"
    CCultivo.AddItem "Repollo"
    CCultivo.AddItem "Zanahoria"
    CCultivo.AddItem "Coliflor"
    CCultivo.AddItem "Apio (Céleri)"
    CCultivo.AddItem "Ajo"
    CCultivo.AddItem "Lechuga"
    CCultivo.AddItem "Cebolla–seca"
    CCultivo.AddItem "Cebolla–verde"
    CCultivo.AddItem "Cebolla–semilla"
    CCultivo.AddItem "Espinaca"
    CCultivo.AddItem "Rábano"
ElseIf CTipo.Text = "b. Hortalizas–Familia de la Solanáceas" Then
    CCultivo.Clear
    CCultivo.AddItem "Berenjena"
    CCultivo.AddItem "Pimiento Dulce(campana)"
    CCultivo.AddItem "Tomate"
ElseIf CTipo.Text = "c. Hortalizas–Familia de las Cucurbitáceas" Then
    CCultivo.Clear
    CCultivo.AddItem "Melón"
    CCultivo.AddItem "Pepino-Cosechado Fresco"
    CCultivo.AddItem "Pepino-Cosechado a Máquina"
    CCultivo.AddItem "Calabaza de Invierno"
    CCultivo.AddItem "Calabacín (zucchini)"
    CCultivo.AddItem "Melón dulce"
    CCultivo.AddItem "Sandía"
ElseIf CTipo.Text = "d. Raíces y Tubérculos" Then
    CCultivo.Clear
    CCultivo.AddItem "Remolacha. mesa"
    CCultivo.AddItem "Yuca o Mandioca-año 1"
    CCultivo.AddItem "Yuca o Mandioca-año 2"
    CCultivo.AddItem "Chirivía"
    CCultivo.AddItem "Chirivía"
    CCultivo.AddItem "Camote o Batata"
    CCultivo.AddItem "Nabos (Rutabaga)"
    CCultivo.AddItem "Remolacha Azucarera"
ElseIf CTipo.Text = "e. Leguminosas(Leguminosae)" Then
    CCultivo.Clear
    CCultivo.AddItem "Frijoles o judías-verdes"
    CCultivo.AddItem "Frijoles o judías-secos y frescos"
    CCultivo.AddItem "Garbanzo (chick pea)"
    CCultivo.AddItem "Habas-Fresco"
    CCultivo.AddItem "Habas-Seco/Semilla"
    CCultivo.AddItem "Garbanzo hindú"
    CCultivo.AddItem "Caupís (cowpeas)"
    CCultivo.AddItem "Maní o Cacahuete"
    CCultivo.AddItem "Lentejas"
    CCultivo.AddItem "Guisantes o arveja-Frescos"
    CCultivo.AddItem "Guisantes o arveja-Secos/Semilla"
    CCultivo.AddItem "Soya o soja"""
ElseIf CTipo.Text = "f. Hortalizas perennes(con letargo invernal y suelo inicialmente desnudo)" Then
    CCultivo.Clear
    CCultivo.AddItem "Alcachofa"
    CCultivo.AddItem "Espárragos"
    CCultivo.AddItem "Menta"
    CCultivo.AddItem "Fresas"
ElseIf CTipo.Text = "g. Cultivos Textiles" Then
    CCultivo.Clear
    CCultivo.AddItem "Algodón"
    CCultivo.AddItem "Lino"
    CCultivo.AddItem "Sisa"
ElseIf CTipo.Text = "h. Cultivos Oleaginosos" Then
    CCultivo.Clear
    CCultivo.AddItem "Ricino"
    CCultivo.AddItem "Canola (colza)"
    CCultivo.AddItem "Cártamo"
    CCultivo.AddItem "Sésamo (ajonjolí)"
    CCultivo.AddItem "Girasol"
ElseIf CTipo.Text = "i. Cereales" Then
    CCultivo.Clear
    CCultivo.AddItem "Cebada"
    CCultivo.AddItem "Avena"
    CCultivo.AddItem "Trigo de Primavera"
    CCultivo.AddItem "Trigo de Invierno-con suelos congelados"
    CCultivo.AddItem "Trigo de Invierno-con suelos no congelados"
    CCultivo.AddItem "Maíz (grano)"
    CCultivo.AddItem "Maíz (dulce)"
    CCultivo.AddItem "Mijo"
    CCultivo.AddItem "Sorgo-grano"
    CCultivo.AddItem "Sorgo-dulce"
    CCultivo.AddItem "Arroz"
ElseIf CTipo.Text = "j. Forrajes" Then
    CCultivo.Clear
    CCultivo.AddItem "Alfalfa (heno)-efecto promedio de los cortes"
    CCultivo.AddItem "Alfalfa (heno)-períodos individuales de corte"
    CCultivo.AddItem "Alfalfa (heno)-para semilla"
    CCultivo.AddItem "Bermuda (heno)-efecto promedio de los cortes"
    CCultivo.AddItem "Bermuda (heno)-cultivo para semilla (primavera)"
    CCultivo.AddItem "Trébol heno. Bersím-efecto promedio de los cortes"
    CCultivo.AddItem "Trébol heno. Bersím-períodos individuales de corte"
    CCultivo.AddItem "Rye Grass (heno)-efecto promedio de los cortes"
    CCultivo.AddItem "Pasto del Sudán (anual)-efecto promedio de los cortes"
    CCultivo.AddItem "Pasto del Sudán (anual)-período individual de corte"
    CCultivo.AddItem "Pastos de Pastoreo-pastos de rotación"
    CCultivo.AddItem "Pastos de Pastoreo-pastoreo extensivo"
    CCultivo.AddItem "Pastos (césped. turfgrass)-época fría"
    CCultivo.AddItem "Pastos (césped. turfgrass)-época caliente"
ElseIf CTipo.Text = "k. Caña de azúcar" Then
    CCultivo.Clear
    CCultivo.AddItem "Caña de Azúcar"
ElseIf CTipo.Text = "l. Frutas Tropicales y Árboles" Then
    CCultivo.Clear
    CCultivo.AddItem "Banana-1er año"
    CCultivo.AddItem "Banana-2do año"
    CCultivo.AddItem "Cacao"
    CCultivo.AddItem "Café-suelo sin cobertura"
    CCultivo.AddItem "Café-con malezas"
    CCultivo.AddItem "Palma Datilera"
    CCultivo.AddItem "Palmas"
    CCultivo.AddItem "Piña-suelo sin cobertura"
    CCultivo.AddItem "Piña-con cobertura de gramíneas"
    CCultivo.AddItem "Árbol del Caucho"
    CCultivo.AddItem "Té-no sombreado"
    CCultivo.AddItem "Té-sombreado"
ElseIf CTipo.Text = "m. Uvas y Moras" Then
    CCultivo.Clear
    CCultivo.AddItem "Moras (arbusto)"
    CCultivo.AddItem "Uvas-Mesa o secas (pasas)"
    CCultivo.AddItem "Uvas-Vino"
    CCultivo.AddItem "Lúpulo"
ElseIf CTipo.Text = "n. Árboles Frutales" Then
    CCultivo.Clear
    CCultivo.AddItem "Almendras. sin cobertura del suelo"
    CCultivo.AddItem "Manzanas. Cerezas. Peras"
    CCultivo.AddItem "Manzanas,Cerezas,Peras-sin cobertura del suelo. con fuertes heladas"
    CCultivo.AddItem "Manzanas,Cerezas,Peras-sin cobertura del suelo. sin heladas"
    CCultivo.AddItem "Manzanas,Cerezas,Peras-cobertura activa del suelo. con fuertes heladas"
    CCultivo.AddItem "Manzanas,Cerezas,Peras-cobertura activa del suelo. sin heladas"
    CCultivo.AddItem "Albaricoque. Melocotón o Durazno. Drupas-sin cobertura del suelo. con fuertes heladas"
    CCultivo.AddItem "Albaricoque. Melocotón o Durazno. Drupas-sin cobertura del suelo. sin heladas"
    CCultivo.AddItem "Albaricoque. Melocotón o Durazno. Drupas-cobertura activa del suelo. con fuertes heladas"
    CCultivo.AddItem "Albaricoque. Melocotón o Durazno. Drupas-cobertura activa del suelo. sin heladas"
    CCultivo.AddItem "Aguacate. sin cobertura del suelo"
    CCultivo.AddItem "Cítricos. sin cobertura del suelo-sin cobertura del suelo"
    CCultivo.AddItem "Cítricos. sin cobertura del suelo-70% cubierta vegetativa"
    CCultivo.AddItem "Cítricos. sin cobertura del suelo-50% cubierta vegetativa"
    CCultivo.AddItem "Cítricos. sin cobertura del suelo-20% cubierta vegetativa"
    CCultivo.AddItem "Cítricos. con cobertura activa del suelo o malezas-70% cubierta vegetativa"
    CCultivo.AddItem "Cítricos. con cobertura activa del suelo o malezas-50% cubierta vegetativa"
    CCultivo.AddItem "Cítricos. con cobertura activa del suelo o malezas-20% cubierta vegetativa"
    CCultivo.AddItem "Coníferas"
    CCultivo.AddItem "Kiwi"
    CCultivo.AddItem "Olivos (40 a 60% de cobertura del suelo por el dosel)"
    CCultivo.AddItem "Pistachos. sin cobertura del suelo"
    CCultivo.AddItem "Huerto de Nogal"

ElseIf CTipo.Text = "o. Humedales–clima templado" Then
    CCultivo.Clear
    CCultivo.AddItem "Anea (Typha). Junco (Scirpus). muerte por heladas"
    CCultivo.AddItem "Anea. Junco. sin heladas"
    CCultivo.AddItem "Vegetación pequeña. sin heladas"
    CCultivo.AddItem "Carrizo (Phragmites). con agua sobre el suelo"
    CCultivo.AddItem "Carrizo. suelo húmedo"
Else
End If
Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("B2").Value = CTipo.Text
End Sub

Private Sub KcEX_Click()

If CCultivo.Text = "" Then
    MsgBox "Debe seleccionar un cultivo"
Else
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("C2").Value = CTipo.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("C3").Value = CCultivo.Text
    
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("B8").Value = KCINI.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("C8").Value = KCMED.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("D8").Value = KCFIN.Text
    
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("D10").Value = ALT.Text
    
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("C15").Value = KcEI1.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("D15").Value = KcED1.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("E15").Value = KcEM1.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("F15").Value = KcEN1.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("G15").Value = KcET1.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("H15").Value = KcEF1.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("I15").Value = KcER1.Text
    
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("C16").Value = KcEI2.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("D16").Value = KcED2.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("E16").Value = KcEM2.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("F16").Value = KcEN2.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("G16").Value = KcET2.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("H16").Value = KcEF2.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("I16").Value = KcER2.Text
    
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("C17").Value = KcEI3.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("D17").Value = KcED3.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("E17").Value = KcEM3.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("F17").Value = KcEN3.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("G17").Value = KcET3.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("H17").Value = KcEF3.Text
    Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Range("I17").Value = KcER3.Text
    '3.- Importamos la hoja de Excel del complemento
                hojas = ActiveSheet.Name
                Workbooks("RegisterU2DF7.xlam").Worksheets("KCExport").Copy _
                       after:=ActiveWorkbook.Sheets(hojas)
                      MsgBox "Se realizo con exito"
End If
End Sub

Private Sub UserForm_Initialize()
CTipo.AddItem "a. Hortalizas Pequeñas"
CTipo.AddItem "b. Hortalizas–Familia de la Solanáceas"
CTipo.AddItem "c. Hortalizas–Familia de las Cucurbitáceas"
CTipo.AddItem "d. Raíces y Tubérculos"
CTipo.AddItem "e. Leguminosas(Leguminosae)"
CTipo.AddItem "f. Hortalizas perennes(con letargo invernal y suelo inicialmente desnudo)"
CTipo.AddItem "g. Cultivos Textiles"
CTipo.AddItem "h. Cultivos Oleaginosos"
CTipo.AddItem "i. Cereales"
CTipo.AddItem "j. Forrajes"
CTipo.AddItem "k. Caña de azúcar"
CTipo.AddItem "l. Frutas Tropicales y Árboles"
CTipo.AddItem "m. Uvas y Moras"
CTipo.AddItem "n. Árboles Frutales"
CTipo.AddItem "o. Humedales–clima templado"
CTipo.AddItem "p. Especial"

CTipo.Text = Workbooks("RegisterU2DF7.xlam").Worksheets("KC").Range("B2").Value

End Sub

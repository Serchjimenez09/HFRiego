# HFRiego Excel

HFRiego es una extensión de Excel 2013 y 2016 que sirve para calcular las pérdidas de carga por fricción en tuberías ciegas y con salidas múltiples. 

Dentro de otras funciones tiene un módulo para calcular la evapotranspiración potencial, otro para diseño agronómico y diversas funciones útiles cuando se diseñan sistemas de riego.

Pueden descargar el complemento en: https://www.hidraulicafacil.com/p/extension-hf-riego.html
<img src="/Images/hfriego png.JPG" alt="HFRiego en Excel"/>

# Iniciar a editar "HFRiego"
1. Para poder comenzar a editar "HFRiego" debe utilizar el archivo "RegisterU2DF7.xlsx", puede abrir este archivo con “Custom UI Editor for Microsoft Office” para poder cambiar nombre de la pestaña, botones, iconos, etc. Dentro del archivo “PestaniaHFRiego.xml” puede encontrar de igual manera el código XML de la pestaña “HFRiego”.
<img src="/Images/Captura2.PNG" alt="HFRiego en Excel"/>

2. Las funciones de cada botón se mandan a llamar desde la propiedad de “onAction= GENERAL2.llamarETOPM“ del archivo XML, y estas lo pueden encontrar dentro de la carpeta  “Funciones”
<img src="/Images/Captura2.PNG" alt="XML"/>
<img src="/Images/Captura4.PNG" alt="Excel"/>

3. Copie las funciones y los formularios a su archivo de Excel y edítelos según los métodos que desee emplear.
 <img src="/Images/Captura5.PNG" alt="Excel"/>
 
Dentro de cada carpeta hay un archivo “Leeme”, debe revisarlo para tener más información sobre los archivos 

# CrearCronograma

`CrearCronograma` es un script en Python 3 para crear automaticamente 
cronogramas para materias. El script genera, a partir de un archivo de texto, 
un archivo Excel con las fechas de cursada (ya sean teoricas, practicas o 
laboratorio) en orden cronologico y marcando los feriados segun el calendario 
academico de la FCEN, UBA. 

## Como usarlo:

1. Crear un archivo de texto `MiCursadaSimple.txt` consignando los siguientes
   datos, uno por linea:

   - de que cuatrimestre se trata (se asume que el año es el actual),
   - la direccion `URL` del calendario exactas, y
   - los diferentes turnos de teoricas, practicas y laboratorio; especificando
     los dias de cursada (en ingles).

Un ejemplo de los contenidos posibles de ese archivo es el siguiente:

    primer cuatrimestre
    http://exactas.uba.ar/academico/display.php?estructura=2&desarrollo=1&id_caja=31&nivel_caja=1
    Teoricas:Monday,Wednesday
    Laboratorio 1:Tuesday
    Laboratorio 2:Friday

1. En la linea de comandos (terminal) ejecutar:

    `./crearcronograma.py MiCursadaSimple.txt Salida.xlsx`

El script genera entonces el archivo de salida `Salida.xlsx`.

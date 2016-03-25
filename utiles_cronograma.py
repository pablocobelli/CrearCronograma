import re
import datetime
import xlsxwriter
import codecs
from bs4 import BeautifulSoup
from urllib.request import urlopen
from collections import OrderedDict

def parsear_datos_iniciales(archivo_de_datos_de_cursada):
    
    f = open(archivo_de_datos_de_cursada, 'r') 
    datos = []
    for line in f:
        datos.append(line.rstrip('\n').strip())
    f.close()

    #  fecha_inicio = datetime.datetime.strptime(
        #  datos[0],'%d-%m-%Y').date()
    #  fecha_final = datetime.datetime.strptime(
        #  datos[1],'%d-%m-%Y').date()
    #  lapso_cursada = [fecha_inicio, fecha_final]
    cursada = datos[0]
    pagina = datos[1]

    datos = datos[2:]

    horarios = OrderedDict()
    for linea in datos:
        turno, dia = linea.split(':')
        dias = dia.split(',')
        horarios.update({turno: dias})
    
    return cursada, pagina, horarios


def lista_de_dias_de_clase(horarios, turno, fecha_inicio, fecha_final, feriados):
    
    delta = fecha_final - fecha_inicio

    lista_dias = []
    contador = 0

    for i in range(delta.days + 1):
        current = fecha_inicio + datetime.timedelta(days=i)
        if current.strftime("%A") in horarios[turno]:
            lista_dias.append(current)
            contador += 1
           
    return lista_dias


def lista_de_feriados(pagina_web_calendario_exactas, guardar=False):

    response = urlopen(pagina_web_calendario_exactas)
    html = response.read()
    soup = BeautifulSoup(html,"lxml")

    # el 1 al final de la siguiente linea es porque la table que buscamos
    # en el sitio web de la FCEN es la *segunda tabla* que aparece con la
    # descripcion de clase "tabla_persona", esperemos que no la cambien!
    tabla_feriados = soup.findAll("table", {"class" : "tabla_persona"})[1] 

    feriados = []

    records = []
    for row in tabla_feriados.findAll('tr')[1:]:
        col = row.findAll('td')
        fecha_feriado = col[0].string.strip()
        razon_feriado = col[1].string.strip()
        record = '%s,%s' % (fecha_feriado, razon_feriado) 
        records.append(record)
        feriados.append(fecha_feriado)
    
    if guardar == True:
        fl = codecs.open('Lista_de_Feriados.txt', 'wb', 'utf8')
        line = ';'.join(records)
        fl.write(line + u'\r\n')
        fl.close()
        
    return feriados


def escribir_cronograma_excel(archivo, horarios, fecha_inicio, fecha_final, feriados):
     
    workbook = xlsxwriter.Workbook(archivo + '.xlsx')
    worksheet = workbook.add_worksheet('Cronograma')

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    formato_fecha = workbook.add_format({'num_format': 'd mmm yyyy'})

    # Escribamos los encabezados de la primera linea.
    worksheet.write('A1', 'Clase', bold)
    worksheet.write('B1', 'Fecha', bold)

    # Titulos de turnos
    j = 2
    for elem in horarios:
        worksheet.write(0, j, list(horarios.keys())[j-2])
        j += 1

    # Escribimos cada turno, despues se ordena la columna fecha
    # para tener el cronograma integrado.
    linea = 1

    for turno in horarios:
        contador_clase = 1
        dias_del_turno = lista_de_dias_de_clase(horarios, turno, fecha_inicio, fecha_final, feriados) 
        for dia in dias_del_turno:
            # Escribimos la fecha
            worksheet.write(linea, 1, dia.strftime('%d-%m-%Y'), formato_fecha)
 
            # Chequeamos si es un feriado:
            #   si lo es, lo advertimos y no numeramos la clase;
            #   de otra forma numeramos la clase
            if dia.strftime('%d-%m-%Y') in feriados:
                worksheet.write(linea, 2, "Feriado")
            else:
                worksheet.write(linea, 0, turno + " " + '{0:02d}'.format(contador_clase))
                contador_clase += 1
            linea += 1 

    workbook.close()

def determinar_lapso_cursada(pagina, cursada):

    contents = urlopen(pagina).read()
    soup = BeautifulSoup(contents, 'lxml')

    inis = soup.body.findAll(text=re.compile('Fecha Inicio:'))
    fins = soup.body.findAll(text=re.compile('Fecha Fin:'))

    inis = [inis[i] for i in [0,2,5]]
    fins = [fins[i] for i in [0,2,5]]

    inis = [texto.split(': ')[1] for texto in inis]
    fins = [texto.split(': ')[1] for texto in fins]

    if cursada.strip() in ['verano', 'Verano']:
        indice = 0
    elif cursada.strip() in ['Primer cuatrimestre', 
            'Primer Cuatrimestre', 'primer cuatrimestre']:
        indice = 1
    elif cursada.strip() in ['Segundo cuatrimestre', 
            'Segundo Cuatrimestre', 'segundo cuatrimestre']:
        indice = 2

    fecha_inicio, fecha_final = inis[indice], fins[indice]
    
    fecha_inicio = datetime.datetime.strptime(
            fecha_inicio,'%d-%m-%Y').date()
    fecha_final = datetime.datetime.strptime(
            fecha_final,'%d-%m-%Y').date()

    return fecha_inicio, fecha_final

def crear_cronograma(archivo_de_datos_de_cursada, archivo_salida):
    
    cursada, pagina, horarios = parsear_datos_iniciales(archivo_de_datos_de_cursada)
    fecha_inicio, fecha_final = determinar_lapso_cursada(pagina, cursada)
    feriados = lista_de_feriados(pagina)
    escribir_cronograma_excel(archivo_salida, horarios, fecha_inicio, fecha_final, feriados)
    print('Cronograma creado en ' + archivo_salida + '.xlsx.') 



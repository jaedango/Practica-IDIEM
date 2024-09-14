# imports
import sqlite3
import pandas as pd
import re
from datetime import date 
import requests
import time

# ------------ request ------------ #
# Extrae la informacion de la API y la retorna como dataframe
def json_to_df():
    url = ''
    payload = {}
    headers = {
        'Authorization': 'Bearer '
    }

    response = requests.request("GET", url, headers=headers, data=payload)
    resp = response.json()
    try:
        for key in resp:
            if key=='data':
                data = pd.DataFrame(resp[key])
                return data
    except:
        print('Ocurrio un error con la API')
    print(response)
    print(f'No hay datos')
    return

# ------------ Funciones relacionadas a BBDD ------------ #
# Create database
def create_sqlite_database(filename):
    conn = None
    try:
        conn = sqlite3.connect(filename)
    except sqlite3.Error as e:
        print(e)
    finally:
        if conn:
            print("Se ha creado db")
            conn.close()

# Funcion para conectarse a db
def conectarse(dir_db):
    conn = sqlite3.connect(dir_db)
    cursor = conn.cursor()
    return conn, cursor

# Funcion para desconectarse de db
def desconectarse(conn):
    conn.commit()
    conn.close()

# Elimina una tabla
def drop_table_fechas(dir_db):
    conn, cursor = conectarse(dir_db)
    cursor.execute(f'DROP TABLE IF EXISTS fechas_db')

    desconectarse(conn)
    print(f'Se ha eliminado correctamente la tabla relacionada fechas')

# Eliminar tabla datos
def drom_table_datos(dir_db):
    conn, cursor = conectarse(dir_db)
    cursor.execute(f'DROP TABLE IF EXISTS datos_db')

# ------------ BBDD fechas ------------ #
# Funcion para crear la tabla de fechas
# contiene id, nombre, fecha de inicio, fecha de termino, diferencia de fechas
# Todos los datos se extraen directamente desde la API
def crear_tabla_fechas(dir_db):
    conn, cursor = conectarse(dir_db)

    # crear table
    cursor.execute(f'''CREATE TABLE IF NOT EXISTS fechas_db
                    (numero_informe TEXT,
                   url TEXT PRIMARY KEY,
                   centro_costo INT,
                   fecha_ingreso TEXT,
                   fecha_termino TEXT,
                   diferencia_fechas TEXT,
                   revision INT,
                   cliente TEXT)''')
    desconectarse(conn)
    print(f'Se ha creado la tabla')

# agrega elementos nuevos mientras no se repita el num de informe
def add_no_rep_num(dir_db, df):
    # chequear si se enceuntran las columnas necesarias
    columns = ["Número Informe", "Link Documento", "Centro de Costos", "Fecha Fin Ensayo", "Fecha Publicación", "Cliente"]
    for column in columns:
        if column not in df.columns:
            print(f"Documento no válido, falta la columna {column}")

    # Seleccionar url distintas
    conn, cursor = conectarse(dir_db)
    cursor.execute(f'SELECT url FROM fechas_db')
    data = cursor.fetchall()
    url_existente = [row[0] for row in data]

    # Revisar si las nuevas urls estan entre los datos anteriores, sino agregar
    for index, row in df.iterrows():
        if row['Link Documento'] in url_existente:
            # st.write(row['Link Documento'])
            continue
        try:
            init_time = row['Fecha Fin Ensayo'].strftime("%d-%m-%Y")
        except:
            init_time = eliminate_symbols(str(row['Fecha Fin Ensayo']))
        
        final_time = row['Fecha Publicación'].split(" ")[0]
        final_time = final_time.split("-")
        end_time = final_time[2] + '-' + final_time[1] + '-' + final_time[0]

        # agregar diferencia fechas
        # Check
        if init_time is None or end_time is None:
            date_dif = "Incorrecto"
        else:
            date_dif = get_difference(init_time, end_time)
            date_dif = date_ranges(date_dif)

        # Revisar si es una revision
        if es_revision(row["Número Informe"]):
            revision = True
        else:
            revision = False

        cliente = row['Cliente'].lstrip().rstrip()
        
        cursor.execute(f'INSERT INTO fechas_db (numero_informe, url, centro_costo, fecha_ingreso, fecha_termino, diferencia_fechas, revision, cliente) VALUES (?, ?, ?, ?, ?, ?, ?, ?)', (row['Número Informe'], row['Link Documento'], row['Centro de Costos'], init_time, end_time, date_dif, revision, cliente))
        url_existente.append(row['Link Documento'])

    desconectarse(conn)
    print(f'Se han agregado los datos')

# muestra toda la info contenida en la tabla
# eliminar
def mostrar_toda_info_fechas(dir_db):
    conn, cursor = conectarse(dir_db)
    cursor.execute(f'SELECT * FROM fechas_db')
    rows = cursor.fetchall()
    desconectarse(conn)
    # Agregar columnas
    data = pd.DataFrame(rows, columns = ["Número de Informe", "Url","Centro de Costo", "Fecha de Último Ensayo", "Fecha de Publicación", "Diferencia de Fechas", "Revision", "Clientes"])
    return data

# Revisa si la tabla esta vacia
def check_table(dir_db):
    try:
        conn, cursor = conectarse(dir_db)
        cursor.execute(f'SELECT * FROM fechas_db')
        items = cursor.fetchall()
        return False if len(items) == 0 else True
    except:
        print(f'Ocurrio un error con la base de datos')
        return False
    
# ------------ Funciones adicionales fechas ------------ #
# Ayuda a formatear fechas, eliminando simbolos, y dejando solo fechas validas
def eliminate_symbols(date):
    if pd.isna(date):
        return None
    else:
        date1 = re.sub('[\.|\-|\/ ]+', '-', date)
        date2 = re.sub('[a-z][A-Z]*', '', date1)
        if re.fullmatch(r'([0-3]\d[\-][0-1]\d[-]20[0-2]\d|[0-3]\d[\-][0-1]\d[-][0-2]\d)', date2):
            return date2
        return None

def es_revision(numero_informe):
    regex_pattern = r'.*(?:rev|v\d|v \d| r2).*'
    if re.match(regex_pattern, numero_informe, re.IGNORECASE):
        return True
    return False
    
# ------------ Arreglo de fechas ------------ #
# una fecha tiene dia d, mes m, año y
class Date:
    def __init__(self, d, m ,y):
        self.d = d
        self.m = m
        self.y = y

# guardar los días del mes
month_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

# años biciestos
def countLeapYears(d):
    years = d.y

    # revisar si el mes debe ser considerado
    if (d.m <= 2):
        years -= 1

    return int(years/4) - int(years/100) + int(years/400)

# Diferencia entre 2 fechas
def get_difference(dt1, dt2):
    # formatear la fecha
    dt1 = format_date(dt1)
    dt2 = format_date(dt2)
    
    # Contar el numero de dias antes de dt1
    n1 = dt1.y * 365 + dt1.d
    # agregar dias segun meses
    for i in range(0, dt1.m - 1):
        n1 += month_days[i]
    
    # agregar dias segun año biciesto
    n1 += countLeapYears(dt1)

    # contar el numero de dias antes de dt2
    n2 = dt2.y * 365 + dt2.d

    for i in range(0, dt2.m -1):
        n2 += month_days[i]

    n2 += countLeapYears(dt2)

    # return dif
    return (n2 - n1)

# test
# revisar si es util
def format_date(date):
    date_list = date.split('-')
    return Date(int(date_list[0]), int(date_list[1]), int(date_list[2]))

# Entrega una mecha en formato "yyyy-mm"
def format_month(fecha):
    fecha_lista = fecha.split('-')
    fecha_format = Date(int(fecha_lista[0]), int(fecha_lista[1]), int(fecha_lista[2]))
    if fecha_format.m < 10:
        fecha = str(fecha_format.y) + '-0' + str(fecha_format.m)
    else:
        fecha = str(fecha_format.y) + '-' + str(fecha_format.m)
    return fecha

# Rangos de datos de dif fechas
def date_ranges(date_dif):
    if date_dif < 0:
        return "Incorrecto"
    elif date_dif <= 1:
        return "0-1"
    elif date_dif <= 3:
        return "2-3"
    elif date_dif <= 5:
        return "4-5"
    elif date_dif <= 8:
        return "6-8"
    elif date_dif <= 10:
        return "9-10"
    elif date_dif <= 14:
        return "11-14"
    elif date_dif > 14:
        return "14+"
    else:
        return "Incorrecto"

# ------------ BBDD de datos procesados ------------ #
# funcion para crear tabla de datos
# contiene nombre=cc-mes-anho, cc, fecha, cuenta total, revisiones, total correctas
# Los datos vienen de fechas_db
# nombre = fecha - centro de costo
# centro de costo son 1817, 2339, 2340, 2341
# fecha se crea mediante una funcion, sacando las ultimas 36 fechas en formato yyyy-mm
# revisiones se extrae desde el numero de informe usando regex
# rangos es una lista que se crea a partir de diferencia de fechas y fecha de termino
def crear_tabla_datos(dir_db):
    conn, cursor = conectarse(dir_db)

    # crear tabla
    cursor.execute(f'''CREATE TABLE IF NOT EXISTS datos_db
                   (nombre TEXT PRIMARY KEY,
                   centro_costo INT,
                   fecha TEXT,
                   total INT,
                   revisiones TEXT,
                   rangos TEXT)''')
    desconectarse(conn)
    print(f'Se ha creado la tabla de datos procesados')

def drop_table_datos(dir_db):
    conn, cursor = conectarse(dir_db)
    cursor.execute(f'DROP TABLE IF EXISTS datos_db')
    desconectarse(conn)

# Actualizar tabla de datos
def update_tabla_datos(dir_db):
    # conseguir lista de meses
    meses = lista_meses(dir_db, 36)

    # lista de centros de costos
    cc_list = [1817, 2339, 2340, 2341]

    # crear el dataframe que va a ser utilizado al final
    final_df = pd.DataFrame(columns=["nombre", "fecha", "cc"])

    # extraer el centro de costo y fecha para hacer el nombre
    for cc in cc_list:
        nombre_fecha = [(fecha + '-' + str(cc), fecha, cc) for fecha in meses]
        nombre_fecha_df = pd.DataFrame(nombre_fecha, columns=["nombre", "fecha", "cc"])
        final_df = pd.concat([final_df, nombre_fecha_df], ignore_index=True, axis=0)
    
    # Obtener la cantidad de informes por periodo
    periodos = get_informes_periodo(dir_db, 36)
    total_informes_df = pd.DataFrame(periodos, columns=['nombre', 'informe', 'cc', 'fecha', 'rango', 'revision'])
    total_informes_df = total_informes_df[["nombre", "informe", "cc", "fecha", "rango"]]
    total_informes_df = total_informes_df.groupby("nombre").count()
    total_informes_df = total_informes_df[["rango"]].rename(columns={"rango":"total"})

    # unir al df final
    final_df = pd.merge(final_df, total_informes_df, on="nombre", how='left')

    # Agregar los rangos
    rangos_total_df = pd.DataFrame(periodos, columns=["nombre", "informe", "cc", "fecha", "rango", "revision"])
    rangos_total_df = rangos_total_df[["nombre", "informe", "cc", "fecha", "rango"]]
    rangos_df = format_rangos(rangos_total_df)

    # unir al df final
    final_df = pd.merge(final_df, rangos_df, on="nombre", how='left')

    # Agregar conteo de revisiones
    revision_total = pd.DataFrame(periodos, columns=["nombre", "informe", "cc", "fecha", "rango", "revision"])
    revision_df = format_revisiones(revision_total)

    # unir al df final
    final_df = pd.merge(final_df, revision_df, on="nombre", how='left')

    # fill none type
    final_df[["rangos", "revisiones"]] = final_df[["rangos", "revisiones"]].fillna("")
    final_df[["total"]] = final_df[["total"]].fillna(0)

    # conectarse a db y actualizar data
    conn, cursor = conectarse(dir_db)
    for index, row in final_df.iterrows():
        cursor.execute(f'INSERT INTO datos_db (nombre, centro_costo, fecha, total, revisiones, rangos) VALUES (?, ?, ?, ?, ?, ?)', (row["nombre"], row["cc"], row["fecha"], row["total"], row["revisiones"], row["rangos"]))
    desconectarse(conn)
    

# ------------ Funciones datos ------------ #
# retorna una lista con los informes segun la cantidad de meses pedida
def get_informes_periodo(dir_db, meses):
    conn, cursor = conectarse(dir_db)
    cursor.execute(f'SELECT numero_informe, centro_costo, fecha_termino, diferencia_fechas, revision FROM fechas_db')
    lista = cursor.fetchall()
    desconectarse(conn)

    # listas con los datos segun columna
    num_informe = [row[0] for row in lista]
    centro_costo = [row[1] for row in lista]
    fecha_termino = [row[2] for row in lista]
    dif_fechas = [row[3] for row in lista]
    revision = [row[4] for row in lista]

    # sacar fechas null
    union = [(item1, item2, item3, item4, item5) for item1, item2, item3, item4, item5 in zip(num_informe, centro_costo, fecha_termino, dif_fechas, revision) if item3 is not None]

    meses = lista_meses(dir_db, meses)
    periodos = [(format_month(item3) + '-' + str(item2), item1, item2, format_month(item3), item4, item5) for item1, item2, item3, item4, item5 in union]
    
    return periodos

# funcion para formatear los rangos y ser puestos en la db
def format_rangos(dataframe):
    # columnas: fecha, dif_dias, id
    rangos = dataframe.groupby(['nombre', 'rango']).size().to_frame()

    rangos_dict = {}
    
    # operaciones infernales para separar numero, fecha y rango
    for index, row in rangos.iterrows():
        item = row.to_string(name=True).split("\n")
        num=int(item[0][-4:])
        item2 = item[1].split('(')
        item2 = item2[1].split(',')
        fecha= item2[0]
        rango = item2[1][1:-1]
        if fecha not in rangos_dict:
            rangos_dict.update({fecha: rango + ':' + str(num)})
        else:
            info = rangos_dict.get(fecha)
            rangos_dict.update({fecha: info +'|'+ rango + ':' + str(num)})
    rangos = pd.DataFrame(rangos_dict.items(), columns=['nombre', 'rangos'])
    return rangos

# funcion para formatear revisiones y ser puestas en db
def format_revisiones(dataframe):
    revision = dataframe[dataframe["revision"]==1]
    revision = revision.groupby(['nombre', 'rango']).size().to_frame()

    revision_dict = {}

    # operaciones para separa numero fecha y rango
    for index, row in revision.iterrows():
        item = row.to_string(name=True).split("\n")
        num = int(item[0][-3:])
        item2 = item[1].split('(')
        item2 = item2[1].split(',')
        fecha = item2[0]
        revision = item2[1][1:-1]
        if fecha not in revision_dict:
            revision_dict.update({fecha: revision + ':' + str(num)})
        else:
            info = revision_dict.get(fecha)
            revision_dict.update({fecha: info +'|'+ revision + ':' + str(num)})
    revisiones = pd.DataFrame(revision_dict.items(), columns=['nombre', 'revisiones'])
    return revisiones

# ------------ Arreglo de fechas ------------ #
# Crea una lista con los meses formateados
# Nota: Agregue el tema de la fecha today() posteriormente debido a que habian datos mal ingresados en fechas fuera de lugar
def lista_meses(dir_db, meses):
    conn, cursor = conectarse(dir_db)
    cursor.execute(f'SELECT fecha_termino FROM fechas_db')
    periodo = set(row[0] for row in cursor.fetchall())
    desconectarse(conn) 

    periodo_list = [format_date(fecha) for fecha in periodo if fecha]

    mayor_mes = 0
    mayor_anho = 0

    # buscar ultimo mes y año
    for fecha in periodo_list:
        if fecha.y > mayor_anho:
            mayor_anho = fecha.y
            mayor_mes = 1
        if fecha.m > mayor_mes and fecha.y == mayor_anho:
            mayor_mes = fecha.m

    # revisar que el ultimo mes sea menor o igual a la fecha actual
    today = date.today().strftime("%m-%Y")
    hoy = today.split('-')
    mayor_anho = mayor_anho if mayor_anho < int(hoy[1]) else int(hoy[1])
    mayor_mes = mayor_mes if mayor_mes < int(hoy[0]) and mayor_anho <= int(hoy[1]) else int(hoy[0])

    if mayor_mes < 10:
        ultimo_mes = str(mayor_anho) + '-0' + str(mayor_mes)
    else:
        ultimo_mes = str(mayor_anho) + '-' + str(mayor_mes)
    
    # crear lista meses
    lista_meses = [ultimo_mes]

    for mes in range(meses - 1):
        mayor_anho = mayor_anho - 1 if mayor_mes == 1 else mayor_anho
        mayor_mes = 12 if mayor_mes == 1 else mayor_mes - 1
        if mayor_mes < 10:
            fecha = str(mayor_anho) + '-0' + str(mayor_mes)
        else:
            fecha = str(mayor_anho) + '-' + str(mayor_mes)
        lista_meses.append(fecha)
    
    return lista_meses

# ------------ Tablas ------------ #
# mostrar en pantalla toda la info de las tablas
# eliminar
def mostrar_toda_info_datos(dir_db):
    conn, cursor = conectarse(dir_db)
    cursor.execute(f'SELECT * FROM datos_db')
    rows = cursor.fetchall()
    desconectarse(conn)
    data = pd.DataFrame(rows, columns=["Nombre", "Centro de Costo", "Fecha", "Total Mensual", "Total Revisiones Mes", "Total por Rangos"])
    return data

# ------------ Programa ------------ #
# usar en C:\Users\rodrigo.jara\anaconda3\scripts>
# .\streamlit run ~\Documents\Javier\dgl_practica\streamlit\data_check_prototype.py
# dir_db = "../data/KPI.db"
dir_db = "./KPI.db"

# Crear base de datos
create_sqlite_database(dir_db)

# dropear tabla de datos - Buscar forma de optimizar esto
drop_table_datos(dir_db)
time.sleep(1)
drop_table_fechas(dir_db)
time.sleep(1)

# Crear tablas
crear_tabla_fechas(dir_db)
crear_tabla_datos(dir_db)

# Update data
df = json_to_df()
add_no_rep_num(dir_db, df)
time.sleep(1)
update_tabla_datos(dir_db)
print("Se ha completado la actualización")
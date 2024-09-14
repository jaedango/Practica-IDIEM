# imports
import sqlite3
import pandas as pd
import streamlit as st
from datetime import date
from io import BytesIO
import sys
sys.path.append("../")

# ------------ Funciones relacionadas a BBDD ------------ #
# Funcion para conectarse a db
def conectarse(dir_db):
    conn = sqlite3.connect(dir_db)
    cursor = conn.cursor()
    return conn, cursor

# Funcion para desconectarse de db
def desconectarse(conn):
    conn.commit()
    conn.close()

# ------------ Funciones Adicionales ------------ #
# Crea una lista con los meses formateados
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

# Funcion para ordenar rangos, tambien sirve para revisiones
def ordenar_rangos(rangos):
    # crear lista vacia para los rangos
    lista_rangos = [0]*8

    # ordenar segun el rango
    for rango in rangos:
        if rango == '':
            continue
        ran = rango.split(':')[0]
        num = int(rango.split(':')[1])
        if ran == '0-1':
            lista_rangos[0] = num
        elif ran == '2-3':
            lista_rangos[1] = num
        elif ran == '4-5':
            lista_rangos[2] = num
        elif ran == '6-8':
            lista_rangos[3] = num
        elif ran == '9-10':
            lista_rangos[4] = num
        elif ran == '11-14':
            lista_rangos[5] = num
        elif ran == '14+':
            lista_rangos[6] = num
        elif ran == 'Incorrecto':
            lista_rangos[7] = num
    return lista_rangos

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

# ------------ Arreglo de fechas ------------ #
# una fecha tiene dia d, mes m, año y
class Date:
    def __init__(self, d, m ,y):
        self.d = d
        self.m = m
        self.y = y

# Retorna las fechas en un formato mas indicado
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

# ------------ Tablas ------------ #
# mostrar en pantalla toda la info de las tablas
def mostrar_toda_info_datos(dir_db):
    conn, cursor = conectarse(dir_db)
    cursor.execute(f'SELECT * FROM datos_db')
    rows = cursor.fetchall()
    desconectarse(conn)
    data = pd.DataFrame(rows, columns=["Nombre", "Centro de Costo", "Fecha", "Total Mensual", "Total Revisiones Mes", "Total por Rangos"])
    st.write(data)

def mostrar_toda_info_fechas(dir_db, cc, meses):
    mes_lista = lista_meses(dir_db, meses)

    conn, cursor = conectarse(dir_db)
    cursor.execute(f'SELECT * FROM fechas_db')
    rows = cursor.fetchall()
    desconectarse(conn)
    # Agregar columnas
    data = pd.DataFrame(rows, columns = ["Número de Informe", "Url","Centro de Costo", "Fecha de Último Ensayo", "Fecha de Publicación", "Diferencia de Fechas", "Revision", "Cliente"])
    
    data[['dia', 'mes', "año"]] = data["Fecha de Publicación"].str.split('-', n=2, expand=True)
    data['Fecha'] = data['año'] + '-' + data['mes']
    
    data["Centro de Costo"] = data["Centro de Costo"].astype(str)
    data = data[data['Centro de Costo'].isin(cc)]
    data = data[data['Fecha'].isin(mes_lista)]

    data = data[["Número de Informe", "Centro de Costo", "Fecha de Último Ensayo", "Fecha de Publicación", "Cliente", "Url"]]
    return data


# Entrega el total segun el mes
# Cantidad de informes por mes
def tabla_cc_mes(dir_db, meses):
    mes_lista = lista_meses(dir_db, meses)

    conn,cursor = conectarse(dir_db)
    df_fechas = pd.DataFrame({'cc':['1817', '2339', '2340', '2341', 'total mes']})

    for mes in mes_lista:
        cursor.execute(f'SELECT centro_costo, total FROM datos_db WHERE fecha = "{mes}"')
        datos_mes = cursor.fetchall()
        cc = [str(row[0]) for row in datos_mes]
        total = [row[1] for row in datos_mes]
        total_mes = sum(total)

        df_mes = pd.DataFrame({'cc': cc, mes:total})

        df_fechas = pd.merge(df_fechas, df_mes, on='cc', how='left')
        df_fechas = df_fechas.fillna(0)
        df_fechas.at[4, mes] = total_mes

    desconectarse(conn)
    st.write(df_fechas)

# Entrega total segun el mes sin revisiones
def tabla_cc_mes_sr(dir_db, meses):
    mes_lista = lista_meses(dir_db, meses)

    conn,cursor = conectarse(dir_db)
    df_fechas = pd.DataFrame({'cc':['1817', '2339', '2340', '2341', 'total mes']})

    for mes in mes_lista:
        cursor.execute(f'SELECT centro_costo, total, revisiones FROM datos_db WHERE fecha = "{mes}"')
        datos_mes = cursor.fetchall()
        cc = [str(row[0]) for row in datos_mes]
        total = [row[1] for row in datos_mes]
        revisiones = [row[2] for row in datos_mes]
        lista_rev = []
        for rev in revisiones:
            if rev == "":
                lista_rev.append(0)
            else:
                if "|" in rev:
                    rev = rev.split("|")
                    count = 0
                    for r in rev:
                        r = r.split(":")
                        count += int(r[1])
                    lista_rev.append(count)
                else:
                    r = int(rev.split(":")[1])
                    lista_rev.append(r)

        total_final = [0] * 4
        for idx in range(4):
            total_final[idx] = total[idx] - lista_rev[idx]
        
        total_r = sum(lista_rev)
        
        total_s = sum(total)
        total_mes = total_s - total_r

        df_mes = pd.DataFrame({'cc': cc, mes:total_final})

        df_fechas = pd.merge(df_fechas, df_mes, on='cc', how='left')
        df_fechas = df_fechas.fillna(0)
        df_fechas.at[4, mes] = total_mes

    desconectarse(conn)
    st.write(df_fechas)

# porcentaje cumplimiento por mes
# entrega los porcentajes de cumplimiento segun el mes
def tabla_resumen_porcentajes(dir_db, meses):
    conn, cursor = conectarse(dir_db)
    mes_lista = lista_meses(dir_db, meses)

    df_porcentajes = pd.DataFrame({'cc': ['1817', '2339', '2340', '2341']})
    for mes in mes_lista:
        cursor.execute(f'SELECT centro_costo, total, rangos, revisiones FROM datos_db WHERE fecha="{mes}"')
        datos_mes = cursor.fetchall()
        cc = [str(row[0]) for row in datos_mes]
        total = [row[1] for row in datos_mes]
        correctas = []
        revisiones = []

        # Extraer revisiones
        for row in datos_mes:
            rango = ordenar_rangos(row[2].split('|'))
            rev = ordenar_rangos(row[3].split('|'))
            total_sumable = sum(rango[:4]) - sum(rev[:4])
            total_revisiones = sum(rev)
            correctas.append(total_sumable)
            revisiones.append(total_revisiones)

        porcentaje_list = []
        for index in range(len(total)):
            correcta_final = correctas[index]
            revisiones_final = revisiones[index]
            total_final = total[index]

            try:
                porcentaje = round(correcta_final/(total_final-revisiones_final), 3)
            except:
                porcentaje = 0.0
            porcentaje = str(porcentaje*100)[:4] + '%'
            porcentaje_list.append(porcentaje)

        df_dato = pd.DataFrame({'cc':cc, mes:porcentaje_list})
        df_porcentajes = pd.merge(df_porcentajes, df_dato, on='cc', how='left')
    df_porcentajes = df_porcentajes.fillna('0.0%')
    desconectarse(conn)
    st.write(df_porcentajes)

# Entrega un resumen como la pag de formantoRenan
# Cantidad de informes segun tiempo de entrega
def tabla_resumen_total(dir_db, meses, cc):
    conn, cursor = conectarse(dir_db)

    mes_lista = lista_meses(dir_db, meses)
    datos=[]

    for mes in mes_lista:
        cursor.execute(f'SELECT * FROM datos_db WHERE fecha = "{mes}"')
        datos_mes = cursor.fetchall()

        for dato in datos_mes:
            dato = list(dato)
            try:
                rangos = ordenar_rangos(dato[-1].split("|"))
            except:
                rangos = [0]*8
            try:
                revisiones = ordenar_rangos(dato[-2].split("|"))
            except:
                revisiones = [0]*8
            for idx, row in enumerate(rangos):
                rangos[idx] -= revisiones[idx]
            dato[4] = sum(revisiones)
            dato[1] = str(dato[1])
            dato.extend(rangos)
            # st.write(dato)
            datos.append(dato)

    desconectarse(conn)

    df_datos = pd.DataFrame(datos, columns=['Nombre', 'Centro de Costo', 'Periodo', 'Total de Informes', 'Revisiones', 'rangos', '0<=t<=1', '1<t<=3', '3<t<=5', '5<t<=8', '8<t<=10', '10<t<=14', 't>14', 'Incorrectas'])
    df_datos = df_datos.fillna(0)
    lst_cols = ['0<=t<=1', '1<t<=3', '3<t<=5', '5<t<=8', '8<t<=10', '10<t<=14', 't>14']
    df_datos["Suma"] = df_datos[lst_cols].sum(axis=1)
    df_datos = df_datos[['Centro de Costo', 'Periodo', 'Total de Informes', 'Revisiones', '0<=t<=1', '1<t<=3', '3<t<=5', '5<t<=8', '8<t<=10', '10<t<=14', 't>14', 'Suma', 'Incorrectas']]

    if cc == "Todos":
        st.write(df_datos)
    else:
        st.write(df_datos[df_datos["Centro de Costo"]==str(cc)])

# Entrega el total segun el mes
# Cantidad de informes por mes
def tabla_resumen_cc(dir_db, mes):
    mes_lista = lista_meses(dir_db, mes)
    cc = ['1817', '2339', '2340', '2341']

    # Nombres filas
    fila1 = "N° Informes Emitidos"
    fila2 = "N° Informes en Plazo"
    fila3 = "N° Informes Fuera de Plazo"
    fila4 = "N° Revisiones"
    fila5 = "N° de Informes Incorrectos"
    fila6 = "% Informes dentro de plazo"

    for centro in cc:
        df_fechas = pd.DataFrame({centro: [fila1, fila2, fila3, fila4, fila5, fila6]})
        for mes in mes_lista:
            conn, cursor = conectarse(dir_db)
            cursor.execute(f'SELECT total, revisiones, rangos FROM datos_db WHERE fecha = "{mes}" AND centro_costo = {centro}')
            try:
                dato = cursor.fetchall()[0]
            except:
                dato = [0, '', '']

            total = dato[0]
            try:
                rangos = ordenar_rangos(dato[2].split('|'))
            except:
                rangos = [0] * 8
            try:
                revisiones = ordenar_rangos(dato[1].split('|'))
            except:
                revisiones = [0] * 8

            total_correctas = sum(rangos[:4]) - sum(revisiones[:4])
            total_revisiones = sum(revisiones)
            fuera_de_plazo = sum(rangos[4:-1])
            incorrectas = rangos[-1]
            try:
                porcentaje = round(total_correctas/(total-total_revisiones), 3)
            except:
                porcentaje = 0.0
            porcentaje = str(porcentaje * 100)[:4] + '%'

            df_dato = pd.DataFrame({centro: [fila1, fila2, fila3, fila4, fila5, fila6], mes: [total, total_correctas, fuera_de_plazo, total_revisiones, incorrectas, porcentaje]})
            df_dato = df_dato.fillna(0)
            df_fechas = pd.merge(df_fechas, df_dato, on=centro, how='left')
            desconectarse(conn)
        st.write(df_fechas)

# Muestra en pantalla las revisiones
def buscar_revisiones(dir_db, meses):
    periodos = get_informes_periodo(dir_db, 36)
    revisiones_total = pd.DataFrame(periodos, columns=["nombre", "Informe", "Centro de Costos", "Fecha", "rango", "revision"])
    revisiones = revisiones_total[revisiones_total["revision"]==1]
    #revision_df = format_revisiones(revisiones_total)
    return revisiones[["Informe", "Centro de Costos", "Fecha"]]

# ------------ Ventanas ------------ #
# Para seleccionar los centros de costos en consultas
def select_cc():
    option = st.selectbox("Filtro según Centro de Costo", ("Todos", "1817", "2339", "2340", "2341"))
    if option == "Todos":
        return "Todos"
    if option == "1817":
        return 1817
    if option == "2339":
        return 2339
    if option == "2340":
        return 2340
    if option == "2341":
        return 2341

# Para seleccionar los centros de costos en consultas
def select_cc2():
    options = st.multiselect("Seleccionar Centro de Costo", ["1817", "2339", "2340", "2341"])
    return options
    
# Formatear revisiones para mostrar en pantalla
def revisiones(dir_db, meses):
    df = buscar_revisiones(dir_db, meses)

    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Hoja1')
    workbook = writer.book
    worksheet = writer.sheets['Hoja1']
    format1 = workbook.add_format({'num_format':'0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def to_excel(df: pd.DataFrame):
    in_memory_fp = BytesIO()
    df.to_excel(in_memory_fp, index=False)
    in_memory_fp.seek(0, 0)
    return in_memory_fp.read()

# ------------ Programa ------------ #
# usar en C:\Users\rodrigo.jara\anaconda3\scripts>
# .\streamlit run ~\Documents\Javier\dgl_practica\streamlit\data_check_prototype.py

# dir_db = "data/KPI.db"
dir_db = "./KPI.db"

st.title("Datos Centros de Costos")

meses = st.slider("Seleccionar la cantidad de meses", 0, 36, 12)

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Resumen", "Porcentaje de Cumplimiento", "Detalles", 'Resumen por Centro de Costo', 'Revisiones', 'Descargas'])

with tab1:
    st.subheader("Cantidad de Informes por mes, incluye revisiones")
    tabla_cc_mes(dir_db, meses)
    # st.caption("\* Incluye Revisiones")
    st.subheader("Cantidad de Informes por mes, sin revisiones")
    tabla_cc_mes_sr(dir_db, meses)

with tab2:
    st.subheader("Porcentaje de Cumplimiento por mes")
    tabla_resumen_porcentajes(dir_db, meses)
    st.caption("\* Cumplimiento para informes entregados 8 días después del último ensayo")
    st.caption("** Porcentajes no incluyen revisiones")

with tab3:
    st.subheader("Cantidad de Informes según tiempo de entrega")
    centro = select_cc()
    tabla_resumen_total(dir_db, meses, centro)
    st.caption('\* Incorrectas se refiere a que tiene un problema con la fecha')

with tab4:
    st.subheader("Resumen de Informes según el Centro de Costo")
    tabla_resumen_cc(dir_db, meses)

with tab5:
    st.subheader("Revisiones según mes")
    rev = buscar_revisiones(dir_db, meses)
    # to_excel = revisiones(dir_db, meses)
    data = to_excel(rev)
    file_name = "Revisiones.xlsx"
    # st.download_button(label = 'Excel', data = to_excel, file_name='revisiones.xlsx')
    st.download_button(
        f"Descargar Listado de Revisiones",
        data,
        file_name,
        f"text/{file_name}",
        key=file_name
    )

    st.write(rev)

with tab6:
    st.subheader("Descarga Datos")
#    mostrar_toda_info_datos(dir_db)
    cc = select_cc2()
    info = mostrar_toda_info_fechas(dir_db, cc, meses)
    info_xlsx = to_excel(info)
    info_name = "Datos.xlsx"
    st.download_button(
        f"Descargar Datos Finales",
        info_xlsx,
        info_name,
        f"text\{info_name}",
        key=info_name
    )
    st.write(info)
    #b = mostrar_toda_info_datos(dir_db)
    #st.write(b)
import streamlit as st
import desarrollo_finanzas as ff
import pandas as pd
import io
from datetime import datetime, timedelta
import requests
from io import BytesIO
from dateutil.relativedelta import relativedelta
import numpy_financial as npf
st.set_page_config(page_title="D5", layout="wide")

# --- Credenciales hardcodeadas ---
USERNAME = "Contabilidad"
PASSWORD = "Esgari2025"   # cámbiala por la que quieras

# Inicializar variables de sesión
def init_session_state():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
        st.session_state["username"] = None

# --- Validación ---
def validar_credenciales(username, password):
    if username == USERNAME and password == PASSWORD:
        return True
    return False

# --- MAIN ---
init_session_state()

if not st.session_state["logged_in"]:
    st.title("🔐 Inicio de Sesión ESGARI 360")

    with st.form("login_form"):
        username = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        submitted = st.form_submit_button("Iniciar sesión")

        if submitted:
            if validar_credenciales(username, password):
                st.session_state["logged_in"] = True
                st.session_state["username"] = username
                st.success("¡Inicio de sesión exitoso!")
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")

else:
    st.sidebar.success(f"✅ Bienvenido, {st.session_state['username']}")

    if st.sidebar.button("Cerrar sesión"):
        st.session_state.clear()
        st.rerun()

    # --- Configuración de página ---
    st.title("📊 Reportes desde Oracle OTM")
    st.write("Selecciona un rango de fechas para ejecutar y mostrar los reportes automáticamente.")

    # --- Parámetros de conexión ---
    USER = 'rolmedo'
    PASS = 'Mexico.2022'
    SERVER = 'ekck.fa.us6'

    # --- Fechas por defecto ---
    hoy = datetime.now().date()
    ayer = hoy - timedelta(days=1)

    # --- Inputs de fechas ---
    col1, col2 = st.columns(2)
    with col1:
        fecha_ini = st.date_input("📅 Fecha inicial", value=ayer)
    with col2:
        fecha_fin = st.date_input("📅 Fecha final", value=hoy)

    # --- Funciones cacheadas ---
    @st.cache_data(show_spinner="⏳ Cargando reporte de Facturas UUID...")
    def get_rf(fecha_ini, fecha_fin):
        user = ff.Sesion(USER, PASS, SERVER)
        params = [
            {
                'dataType': 'Date',
                'name': 'P_FECHA_INI',
                'dateFormatString': 'DD-MM-YYYY',
                'values': fecha_ini.strftime('%m-%d-%Y'),
                'multiValuesAllowed': False,
                'refreshParamOnChange': False,
                'selectAll': False,
                'templateParam': False,
                'useNullForAll': False
            },
            {
                'dataType': 'Date',
                'name': 'P_FECHA_FIN',
                'dateFormatString': 'DD-MM-YYYY',
                'values': fecha_fin.strftime('%m-%d-%Y'),
                'multiValuesAllowed': False,
                'refreshParamOnChange': False,
                'selectAll': False,
                'templateParam': False,
                'useNullForAll': False
            }
        ]
        ruta = "/Custom/ESGARI/Qlik/ReportesFinanzas/XXRO_EXTRACTOR_GL_REP.xdo"
        h = user.runReport(ruta, params=params)
        return pd.read_csv(io.BytesIO(h.reportBytes))

    rf = get_rf(fecha_ini, fecha_fin)
    rf = rf.fillna({'CREDIT': 0, 'DEBIT': 0})
    rf['Neto'] = rf['DEBIT'] - rf['CREDIT']

    @st.cache_data
    def cargar_datos(url):
        response = requests.get(url)
        response.raise_for_status()
        archivo_excel = BytesIO(response.content)
        return pd.read_excel(archivo_excel, engine="openpyxl")

    contratos = 'https://docs.google.com/spreadsheets/d/1MsLQEZXYj60eGqp-Fq8_I6BDESJ97AVf/export?format=xlsx'
    arrendadoras = 'https://docs.google.com/spreadsheets/d/1kVKUntKgQ-B5NPNbXFnAuXCQC9G2KK7is_aKhGsec4Q/export?format=xlsx'

    df_contratos = cargar_datos(contratos)
    df_con_orig = df_contratos.copy()
    df_contratos.fillna(0, inplace=True)
    df_arrendadoras = cargar_datos(arrendadoras)

    with st.expander("Ver detalles de Contratos"):
        st.subheader("📋 Datos de Contratos")
        st.dataframe(df_contratos)

    with st.expander("Ver detalles de Facturas UUID"):
        st.subheader("📄 Reporte de Facturas UUID")
        st.write(rf)

    st.sidebar.subheader("Carga archivo de Excel de estatus de unidades")
    uploaded_file = st.sidebar.file_uploader("Sube un archivo .xlsx", type=["xlsx"])
    if uploaded_file is not None:
        unidades_activas = pd.read_excel(uploaded_file)
    distribucion_david = st.sidebar.file_uploader("Sube archivo de distribucion", type=["xlsx"])
    columnas_leer = ['unidad', 'MANZANILLO2', 'CONTINENTAL3', 'CENTRAL4', 'FLEX SPOT5', 'CHALCO6', 'ARRAYANES7', 'FLEX DEDICADO8', 'INTERNACIONAL FWD9']
    if distribucion_david is not None:
        dit_ca = pd.read_excel(distribucion_david, sheet_name='camiones', engine='openpyxl', header=1, usecols=columnas_leer)


    if uploaded_file is not None and distribucion_david is not None:
        st.sidebar.success("Archivo cargado correctamente.")
        def calcular_meses(row):
            delta = relativedelta(row['FECHA INIO '], row['FECHA FIN'])
            return (delta.years * 12 + delta.months)*-1

        # Aplicar por fila
        df_contratos['meses contrato'] = df_contratos.apply(calcular_meses, axis=1)
        def calcular_vp(row):
            return npf.pv(rate=row['TASA']/12, nper=row['meses contrato'], pmt=-row['MENSUALIDAD'], fv=0) + row['PAGO INICIAL ']

        df_contratos['vp contrato'] = df_contratos.apply(calcular_vp, axis=1)

        unidades_activas['UNIDAD'] = unidades_activas['UNIDAD_GID']

        df_contratos = df_contratos.merge(
            unidades_activas[['UNIDAD', 'TIPO_UNIDAD']],
            on='UNIDAD',
            how='left'  # Simula BUSCARV: busca por clave y trae el valor
        )

        df_contratos = df_contratos.merge(
            unidades_activas[['UNIDAD', 'ACTIVO__Y_N_']],
            on='UNIDAD',
            how='left'  # Simula BUSCARV: busca por clave y trae el valor
        )
        
        df_contratos['baja'] = df_contratos['ACTIVO__Y_N_']
        df_contratos.drop(columns=['ACTIVO__Y_N_'], inplace=True)

        rf = rf[rf['SEGMENT5'] == 510100070]

        rf['UNIDAD ORACLE'] = rf['SEGMENT7']
        rf = rf[~rf['DESCRIPTION'].str.contains('ARRENDAMIENTO')]
        rf = rf[~rf['DESCRIPTION'].str.contains('AJUSTE')]
        rf_gro = rf.groupby(['UNIDAD ORACLE'], as_index=False).agg({
        'Neto': 'sum',
        'DEBIT': 'sum',
        'CREDIT': 'sum'
    })

        df_contratos = df_contratos.merge(
            rf_gro[['UNIDAD ORACLE', 'Neto', 'DEBIT', 'CREDIT']],
            on='UNIDAD ORACLE',
            how='left'  # Simula BUSCARV: busca por clave y trae el valor
        )
        rf_drop = rf.drop_duplicates(subset=['SEGMENT1', 'SEGMENT2', 'SEGMENT3', 'SEGMENT4', 'SEGMENT7'])
        df_contratos = df_contratos.merge(
            rf_drop[['UNIDAD ORACLE', 'SEGMENT1', 'SEGMENT2', 'SEGMENT3', 'SEGMENT4', 'SEGMENT7']],
            on='UNIDAD ORACLE',
            how='left'  # Simula BUSCARV: busca por clave y trae el valor
        )

        df_contratos['DEBIT D5'] = df_contratos["MENSUALIDAD"]*-1
        df_contratos['CREDIT D5'] = 0

        df_contratos['vp amortizacion'] = df_contratos['vp contrato'] + df_contratos['PAGO INICIAL '] + df_contratos['MENSUALIDAD']

        df_contratos['amortizacion'] = df_contratos['vp amortizacion'] / (df_contratos['meses contrato'] + 1)

        def calcular_capital_pagado(row):
            tasa_mensual = row['TASA'] / 12
            meses_transcurridos = relativedelta(fecha_fin, row['FECHA INIO '])
            periodo_actual = meses_transcurridos.years * 12 + meses_transcurridos.months
            
            # Asegura que el periodo actual no exceda el número de pagos
            periodo_actual = min(periodo_actual, row['meses contrato'])
            
            capital_mes = -npf.ppmt(rate=tasa_mensual, per=periodo_actual, nper=row['meses contrato'], pv=row['vp contrato'])
            
            return capital_mes

        df_contratos['vp mensual contrato'] = df_contratos.apply(calcular_capital_pagado, axis=1)

        def calcular_interes_pagado(row):
            tasa_mensual = row['TASA'] / 12
            meses_transcurridos = relativedelta(fecha_fin, row['FECHA INIO '])
            periodo_actual = meses_transcurridos.years * 12 + meses_transcurridos.months

            # Asegura que el periodo actual no exceda el número de pagos
            periodo_actual = min(periodo_actual, row['meses contrato'])

            interes_mes = -npf.ipmt(rate=tasa_mensual, per=periodo_actual, nper=row['meses contrato'], pv=row['vp contrato'])

            return interes_mes
        
        df_contratos['interes pagado'] = df_contratos.apply(calcular_interes_pagado, axis=1)

        df_contratos['intercompañia'] = "00"

        df_contratos['futuro 2'] = "0000"

        df_contratos['Moneda'] = "MXN"

        df_contratos['diferencia'] = df_contratos['MENSUALIDAD'] - df_contratos['Neto']

        mes_a_texto = {
            1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio",
            7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
        }

        nombre_mes = mes_a_texto.get(fecha_fin.month, str(fecha_fin.month))  # por si acaso

        df_contratos['descripcion arrendamiento'] = df_contratos.apply(
            lambda row: f"ARRENDAMIENTO / {row['PROVEEDOR']} / {row['ANEXO']} / {nombre_mes} {fecha_fin.year}",
            axis=1
        )

        df_contratos['descripcion amortizacion'] = df_contratos.apply(
            lambda row: f"AMORTIZACION / {row['PROVEEDOR']} / {row['ANEXO']} / {nombre_mes} {fecha_fin.year}",
            axis=1
        )

        df_contratos['descripcion intereses'] = df_contratos.apply(
            lambda row: f"INTERESES / {row['PROVEEDOR']} / {row['ANEXO']} / {nombre_mes} {fecha_fin.year}",
            axis=1
        )

        df_contratos['descripcion capital'] = df_contratos.apply(
            lambda row: f"CAPITAL / {row['PROVEEDOR']} / {row['ANEXO']} / {nombre_mes} {fecha_fin.year}",
            axis=1
        )   

        df_contratos['SEGMENT7'] = df_contratos['SEGMENT7'].astype(str).str.pad(width=4, side='left', fillchar='0')

        df_arrendadoras.rename(columns={'EMPRESA': 'PROVEEDOR'}, inplace=True)

        df_contratos = df_contratos.merge(
            df_arrendadoras[['PROVEEDOR','PASIVO', 'ACTIVO']],
            on='PROVEEDOR',
            how = 'left'
        )
        dit_ca = dit_ca[dit_ca.ne(1).all(axis=1)]
        dit_ca = dit_ca.dropna().reset_index(drop=True)

        uni_comp = dit_ca['unidad'].unique().tolist()
        df_uni_comp = df_contratos[df_contratos['UNIDAD'].isin(uni_comp)]
    
        st.write(df_contratos)

        #falta la cuenta
        df_para_plantilla = df_contratos.copy()
        df_para_plantilla = df_para_plantilla[~df_para_plantilla['UNIDAD'].isin(uni_comp)]
        df_para_plantilla['SEGMENT3'] = df_contratos['SEGMENT3'].astype(str).str.pad(width=5, side='left', fillchar='0')
        amortizacion = df_para_plantilla.copy()
        amortizacion['SEGMENT5'] = 510100070
        crear_columnas = ['Fecha de cambio', 'Clase de tipo de cambio', 'tipo de cambio', 'Debito contabilizado', 'credito contabilizado']
        amortizacion['DEBIT'] = amortizacion['MENSUALIDAD']*-1
        amortizacion['CREDIT'] = 0
        for col in crear_columnas:
            amortizacion[col] = ""
        orden_columnas = ['SEGMENT1', 'SEGMENT2', 'SEGMENT3', 'SEGMENT4', 'SEGMENT5', 'intercompañia',
        'SEGMENT7', 'futuro 2','Moneda', 'DEBIT', 'CREDIT', 'Fecha de cambio', 'Clase de tipo de cambio', 'tipo de cambio', 'Debito contabilizado', 'credito contabilizado', 'descripcion arrendamiento', 
        'descripcion amortizacion', 'descripcion intereses', 'descripcion capital']
        amortizacion = amortizacion[orden_columnas]
        amortizacion.drop(columns=['descripcion intereses', 'descripcion capital', 'descripcion arrendamiento'], inplace=True)
        arrendamientos = df_para_plantilla.copy()
        arrendamientos['SEGMENT5'] = 540004000
        arrendamientos['DEBIT'] = arrendamientos['amortizacion']
        arrendamientos['CREDIT'] = 0
        for col in crear_columnas:
            arrendamientos[col] = ""
        arrendamientos = arrendamientos[orden_columnas]
        arrendamientos.drop(columns=['descripcion amortizacion', 'descripcion intereses', 'descripcion capital'], inplace=True)

        intereses = df_para_plantilla.copy()
        intereses['SEGMENT5'] = 520003000
        intereses['DEBIT'] = intereses['interes pagado']
        intereses['CREDIT'] = 0
        for col in crear_columnas:
            intereses[col] = ""
        intereses = intereses[orden_columnas]
        intereses.drop(columns=['descripcion arrendamiento', 'descripcion amortizacion', 'descripcion capital'], inplace=True)

        amortizacion_pasivo = df_para_plantilla.copy()
        amortizacion_pasivo['SEGMENT5'] = amortizacion_pasivo['PASIVO']
        amortizacion_pasivo['DEBIT'] = amortizacion_pasivo['vp mensual contrato']
        amortizacion_pasivo['CREDIT'] = 0
        for col in crear_columnas:
            amortizacion_pasivo[col] = ""
        amortizacion_pasivo = amortizacion_pasivo[orden_columnas]
        amortizacion_pasivo.drop(columns=['descripcion arrendamiento', 'descripcion intereses', 'descripcion capital'], inplace=True)

        amortizacion_activo = df_para_plantilla.copy()
        amortizacion_activo['SEGMENT5'] = 135000004
        amortizacion_activo['DEBIT'] = 0
        amortizacion_activo['CREDIT'] = amortizacion_activo['amortizacion']
        for col in crear_columnas:
            amortizacion_activo[col] = ""
        amortizacion_activo = amortizacion_activo[orden_columnas]
        amortizacion_activo.drop(columns=['descripcion arrendamiento', 'descripcion intereses', 'descripcion capital'], inplace=True)

        nuevos_pasivo = df_para_plantilla.copy()
        nuevos_pasivo = nuevos_pasivo[nuevos_pasivo['NUEVO'] == 'si']
        nuevos_pasivo['SEGMENT5'] = nuevos_pasivo['PASIVO']
        nuevos_pasivo['CREDIT'] = nuevos_pasivo['vp amortizacion']
        nuevos_pasivo['DEBIT'] = 0
        for col in crear_columnas:
            nuevos_pasivo[col] = "" 
        nuevos_pasivo = nuevos_pasivo[orden_columnas]
        nuevos_pasivo.drop(columns=['descripcion arrendamiento', 'descripcion intereses', 'descripcion amortizacion'], inplace=True)

        nuevos_activo = df_para_plantilla.copy()
        nuevos_activo = nuevos_activo[nuevos_activo['NUEVO'] == 'si']
        nuevos_activo['SEGMENT5'] = nuevos_activo['ACTIVO']
        nuevos_activo['CREDIT'] = 0
        nuevos_activo['DEBIT'] = nuevos_activo['vp amortizacion']
        for col in crear_columnas:
            nuevos_activo[col] = ""
        nuevos_activo = nuevos_activo[orden_columnas]
        nuevos_activo.drop(columns=['descripcion arrendamiento', 'descripcion intereses', 'descripcion amortizacion'], inplace=True)
        
        #mostrar tablas en expander
        with st.expander("Ver detalles de Arrendamientos"):
            st.write(arrendamientos)
        with st.expander("Ver detalles de Intereses"):
            st.write(intereses)
        with st.expander("Ver detalles de Amortización"):   
            st.write(amortizacion)
        with st.expander("Ver detalles de Amortización Pasivo"):
            st.write(amortizacion_pasivo)
        with st.expander("Ver detalles de Amortización Activo"):
            st.write(amortizacion_activo)
        with st.expander("Ver detalles de Nuevos Pasivos"):
            st.write(nuevos_pasivo)
        with st.expander("Ver detalles de Nuevos Activos"):
            st.write(nuevos_activo)

        #repartir entre proyectos unidades compartidas
        dit_ca['UNIDAD'] = dit_ca['unidad']
        df_uni_comp = df_uni_comp.merge(on='UNIDAD', right=dit_ca, how='left')

        df_uni_comp['descripcion ajuste'] = df_uni_comp.apply(
            lambda row: f"AJUSTE / {row['PROVEEDOR']} / {row['ANEXO']} / {nombre_mes} {fecha_fin.year}",
            axis=1
        )

        df_uni_comp['05001'] = df_uni_comp['MANZANILLO2']
        df_uni_comp['03201'] = df_uni_comp['CONTINENTAL3']
        df_uni_comp['03002'] = df_uni_comp['CENTRAL4']
        df_uni_comp['02003'] = df_uni_comp['FLEX SPOT5']
        df_uni_comp['01001'] = df_uni_comp['CHALCO6']
        df_uni_comp['01003'] = df_uni_comp['ARRAYANES7']
        df_uni_comp['02001'] = df_uni_comp['FLEX DEDICADO8']
        df_uni_comp['07806'] = df_uni_comp['INTERNACIONAL FWD9']



        #movimientos amortizacion compartidas

        amortizacion_com = df_uni_comp.copy()
        amortizacion_com['SEGMENT5'] = 510100070
        crear_columnas = ['Fecha de cambio', 'Clase de tipo de cambio', 'tipo de cambio', 'Debito contabilizado', 'credito contabilizado']
        amortizacion_com['DEBIT'] = amortizacion_com['MENSUALIDAD']*-1
        amortizacion_com['CREDIT'] = 0
        for col in crear_columnas:
            amortizacion_com[col] = ""
        orden_columnas = ['SEGMENT1', 'SEGMENT2', 'SEGMENT3', 'SEGMENT4', 'SEGMENT5', 'intercompañia',
        'SEGMENT7', 'futuro 2','Moneda', 'DEBIT', 'CREDIT', 'Fecha de cambio', 'Clase de tipo de cambio', 'tipo de cambio', 'Debito contabilizado', 'credito contabilizado', 'descripcion arrendamiento', 
        'descripcion amortizacion', 'descripcion intereses', 'descripcion capital', 'descripcion ajuste']
        amortizacion_com = amortizacion_com[orden_columnas]
        amortizacion_com.drop(columns=['descripcion arrendamiento', 'descripcion intereses', 'descripcion capital', 'descripcion ajuste'], inplace=True)


        df_distri = df_uni_comp.copy()

        # Quitar columnas de proyectos "manuales"
        df_distri = df_distri.drop(columns=[
            'MANZANILLO2', 'CONTINENTAL3', 'CENTRAL4', 'FLEX SPOT5',
            'CHALCO6', 'ARRAYANES7', 'FLEX DEDICADO8', 'INTERNACIONAL FWD9'
        ], errors="ignore")

        cols_distribucion = ['05001', '03201', '03002', '02003', '01001', '01003', '02001', '07806']

        # Si existe un SEGMENT3 original, renombrarlo o quitarlo
        if "SEGMENT3" in df_distri.columns:
            df_distri = df_distri.rename(columns={"SEGMENT3": "SEGMENT3_ORIGINAL"})

        # Columnas base (todas menos las de distribución)
        base_cols = [c for c in df_distri.columns if c not in cols_distribucion]

        # Derretir las columnas de distribución → cada fila será un proyecto
        df_melt = df_distri.melt(
            id_vars=base_cols,
            value_vars=cols_distribucion,
            var_name="SEGMENT3",
            value_name="porcentaje"
        )

        # Solo quedarnos con filas con porcentaje > 0
        df_melt = df_melt[df_melt["porcentaje"] > 0]

        # Distribuir valores financieros
        for col in ["amortizacion", "interes pagado", "vp mensual contrato"]:
            df_melt[col] = df_melt[col] * df_melt["porcentaje"]

        # Eliminar columna auxiliar "porcentaje"
        df_final = df_melt.drop(columns=["porcentaje"])

        #intereses compartidas
        intereses_com = df_final.copy()
        intereses_com['SEGMENT5'] = 520003000
        intereses_com['DEBIT'] = intereses_com['interes pagado']
        intereses_com['CREDIT'] = 0
        for col in crear_columnas:
            intereses_com[col] = ""
        intereses_com = intereses_com[orden_columnas]
        intereses_com.drop(columns=['descripcion arrendamiento', 'descripcion amortizacion', 'descripcion capital', 'descripcion ajuste'], inplace=True)

        #arrendamientos compartidas
        arrendamientos_com = df_final.copy()
        arrendamientos_com['SEGMENT5'] = 540004000
        arrendamientos_com['DEBIT'] = arrendamientos_com['amortizacion']
        arrendamientos_com['CREDIT'] = 0
        for col in crear_columnas:
            arrendamientos_com[col] = ""
        arrendamientos_com = arrendamientos_com[orden_columnas]
        arrendamientos_com.drop(columns=['descripcion amortizacion', 'descripcion intereses', 'descripcion capital', 'descripcion ajuste'], inplace=True)

        #amortizacion pasivo compartidas
        amortizacion_pasivo_com = df_final.copy()
        amortizacion_pasivo_com['SEGMENT5'] = amortizacion_pasivo_com['PASIVO']
        amortizacion_pasivo_com['DEBIT'] = amortizacion_pasivo_com['vp mensual contrato']
        amortizacion_pasivo_com['CREDIT'] = 0
        for col in crear_columnas:
            amortizacion_pasivo_com[col] = ""
        amortizacion_pasivo_com = amortizacion_pasivo_com[orden_columnas]
        amortizacion_pasivo_com.drop(columns=['descripcion arrendamiento', 'descripcion intereses', 'descripcion capital', 'descripcion ajuste'], inplace=True)
        #amortizacion activo compartidas
        amortizacion_activo_com = df_final.copy()
        amortizacion_activo_com['SEGMENT5'] = 135000004
        amortizacion_activo_com['DEBIT'] = 0
        amortizacion_activo_com['CREDIT'] = amortizacion_activo_com['amortizacion']
        for col in crear_columnas:
            amortizacion_activo_com[col] = ""
        amortizacion_activo_com = amortizacion_activo_com[orden_columnas]
        amortizacion_activo_com.drop(columns=['descripcion arrendamiento', 'descripcion intereses', 'descripcion capital', 'descripcion ajuste'], inplace=True)
        #nuevos pasivo compartidas
        nuevos_pasivo_com = df_uni_comp.copy()
        nuevos_pasivo_com = nuevos_pasivo_com[nuevos_pasivo_com['NUEVO'] == 'si']
        nuevos_pasivo_com['SEGMENT5'] = nuevos_pasivo_com['PASIVO']
        nuevos_pasivo_com['CREDIT'] = nuevos_pasivo_com['vp amortizacion']
        nuevos_pasivo_com['DEBIT'] = 0
        for col in crear_columnas:
            nuevos_pasivo_com[col] = ""
        nuevos_pasivo_com = nuevos_pasivo_com[orden_columnas]
        nuevos_pasivo_com.drop(columns=['descripcion arrendamiento', 'descripcion intereses', 'descripcion amortizacion', 'descripcion ajuste'], inplace=True)
        #nuevos activo compartidas
        nuevos_activo_com = df_uni_comp.copy()
        nuevos_activo_com = nuevos_activo_com[nuevos_activo_com['NUEVO'] == 'si']
        nuevos_activo_com['SEGMENT5'] = nuevos_activo_com['ACTIVO']
        nuevos_activo_com['CREDIT'] = 0
        nuevos_activo_com['DEBIT'] = nuevos_activo_com['vp amortizacion']
        for col in crear_columnas:
            nuevos_activo_com[col] = ""
        nuevos_activo_com = nuevos_activo_com[orden_columnas]
        nuevos_activo_com.drop(columns=['descripcion arrendamiento', 'descripcion intereses', 'descripcion amortizacion', 'descripcion ajuste'], inplace=True)
        #mostrar tablas en expander
        with st.expander("Ver detalles de Arrendamientos Compartidas"):
            st.write(arrendamientos_com)
        with st.expander("Ver detalles de Intereses Compartidas"):
            st.write(intereses_com)
        with st.expander("Ver detalles de Amortización Compartidas"):   
            st.write(amortizacion_com)
        with st.expander("Ver detalles de Amortización Pasivo Compartidas"):
            st.write(amortizacion_pasivo_com)
        with st.expander("Ver detalles de Amortización Activo Compartidas"):
            st.write(amortizacion_activo_com)
        with st.expander("Ver detalles de Nuevos Pasivos Compartidas"):
            st.write(nuevos_pasivo_com)
        with st.expander("Ver detalles de Nuevos Activos Compartidas"):
            st.write(nuevos_activo_com)
        # unir todas las tablas
        resultado_final = pd.concat([
            arrendamientos, intereses, amortizacion, amortizacion_pasivo,
            amortizacion_activo, nuevos_pasivo, nuevos_activo,
            arrendamientos_com, intereses_com, amortizacion_com,
            amortizacion_pasivo_com, amortizacion_activo_com,
            nuevos_pasivo_com, nuevos_activo_com
        ], ignore_index=True)

        # ✅ Unificar descripciones en una sola columna
        cols_desc = [
            "descripcion arrendamiento",
            "descripcion amortizacion",
            "descripcion intereses",
            "descripcion capital",
        ]
        # Crear columna "descripcion" tomando la primera no vacía
        resultado_final["descripcion"] = resultado_final[cols_desc].bfill(axis=1).iloc[:, 0]

        # Eliminar las columnas de descripción originales
        resultado_final = resultado_final.drop(columns=cols_desc, errors="ignore")

        # descargar resultado final
        st.subheader("📥 Descargar Resultado Final")

        def convertir_a_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Reporte')
                writer.save()
            processed_data = output.getvalue()
            return processed_data

        excel_data = convertir_a_excel(resultado_final)

        import gspread
        from oauth2client.service_account import ServiceAccountCredentials

        def actualizar_google_sheet(df, json_file, spreadsheet_id, worksheet_name="Hoja1"):
            # Configurar acceso
            scope = ["https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name(json_file, scope)
            client = gspread.authorize(creds)

            # Abrir el spreadsheet por ID
            sh = client.open_by_key(spreadsheet_id)

            # Seleccionar la pestaña (worksheet)
            worksheet = sh.worksheet(worksheet_name)

            # Limpiar hoja
            worksheet.clear()

            # 🔑 Convertir todas las columnas a string (especialmente fechas)
            df_str = df.copy()
            for col in df_str.select_dtypes(include=["datetime64[ns]"]).columns:
                df_str[col] = df_str[col].dt.strftime("%Y-%m-%d")  # formato limpio de fecha

            # Convertir también cualquier otro tipo raro (Timestamp, NaT, etc.)
            df_str = df_str.astype(str)

            # Subir DataFrame (encabezados + datos)
            worksheet.update([df_str.columns.values.tolist()] + df_str.values.tolist())

        # --- USO en tu app ---
        json_file = dict(st.secrets["gcp_service_account"])
        spreadsheet_id = "1Vzw1lWuWaC0uvbqxh6JrNqu9UEnjHXywqwG9cDIowRk"

        if st.download_button(
            label="Descargar archivo Excel",
            data=excel_data,
            file_name=f'd5_contabilidad_{fecha_fin}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ):
            df_con_orig['NUEVO'] = 'no'
            actualizar_google_sheet(df_con_orig, json_file, spreadsheet_id, worksheet_name="Hoja 1")
            st.success("✅ Google Sheet actualizado correctamente")




    else:
        st.sidebar.warning("Por favor, sube los archivo para continuar.")






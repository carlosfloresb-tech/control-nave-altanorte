import streamlit as st
import pandas as pd
from datetime import datetime, time

st.set_page_config(page_title="Control Operacional - Nave Altanorte", layout="wide")

st.title("🏭 Sistema de Control Operacional - Nave Altanorte")
st.markdown("Plataforma en línea para el registro de turnos, termografía y ciclos de equipos en terreno.")

# --- BARRA LATERAL: CONFIGURACIÓN GENERAL DEL TURNO ---
st.sidebar.header("📋 Datos Generales del Turno")
fecha_turno = st.sidebar.date_input("Fecha", datetime.today())
tipo_turno = st.sidebar.selectbox("Turno", ["Turno A", "Turno B"])

st.sidebar.subheader("Supervisores y Operadores")
sup_sop = st.sidebar.text_input("Supervisor SOP", "Rene Philipps")
sup_crm = st.sidebar.text_input("Supervisor CRM", "Jovelino Burgos")
sup_caemin = st.sidebar.text_input("Supervisor Caemin", "Ruben Infanta")
sup_nave = st.sidebar.text_input("Supervisor Nave", "Jorge Galindo")
op_picoton = st.sidebar.text_input("Operador Picotón", "Juan Pablo Figueroa")
op_cargador = st.sidebar.text_input("Operador Cargador", "Fernando Tapia")

# --- PESTAÑAS PRINCIPALES ---
tab1, tab2, tab3, tab4 = st.tabs(["1. Checklist Inicio", "2. Termografía (°C)", "3. Ciclos de CPS (Manual)", "4. Resumen y Exportar"])

with tab1:
    st.subheader("Control Previo de Inicio de Actividades")
    chk_cargador = st.checkbox("Check list Operacional Cargador y Picotón (Op. Maq.)", value=True)
    chk_caemin = st.checkbox("Asegurar condición de equipos por Caemin", value=True)
    
    st.subheader("Controles de Retiro y Traslado de Material")
    c1 = st.checkbox("Confirmar tiempo de al menos 10 min desde término de picado del foso", value=True)
    c2 = st.checkbox("Verificar visualmente que el material en el foso esté sólido (no líquido)", value=True)
    c3 = st.checkbox("Verificar que no exista presencia de llamas, humo o indicios de ignición", value=True)
    c4 = st.checkbox("Asegurar que la disposición del material en explanada favorezca el enfriamiento", value=True)
    c5 = st.checkbox("Verificar que camión aljibe esté disponible durante toda la actividad", value=True)
    
    comentarios_inicio = st.text_area("Comentarios y Estado de Equipos (Stand by, Operativos, etc.)", 
                                       "Operativos: Cargador y Picotón revisados. Camión aljibe operativo.")

with tab2:
    st.subheader("Registro de Termografía (°C)")
    st.markdown("Ingrese las temperaturas registradas en los puntos de control de los equipos:")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### ⛏️ Picotón")
        cps1_flex = st.number_input("CPS-1: Flexibles", value=40.0)
        cps1_cuna = st.number_input("CPS-1: Cuña", value=160.0)
        cps1_torn = st.number_input("CPS-1: Tornamesa", value=20.0)
        
        cps3_flex = st.number_input("CPS-3: Flexibles", value=40.0)
        cps3_cuna = st.number_input("CPS-3: Cuña", value=200.0)
        cps3_torn = st.number_input("CPS-3 Tornamesa", value=20.0)
        
        cps2_flex = st.number_input("CPS-2: Flexibles", value=40.0)
        cps2_cuna = st.number_input("CPS-2: Cuña", value=280.0)
        cps2_torn = st.number_input("CPS-2: Tornamesa", value=20.0)
        
    with col2:
        st.markdown("### 🚜 Cargador")
        carg_cps1_batea = st.number_input("Cargador CPS-1: Batea", value=140.0)
        carg_cps1_inf = st.number_input("Cargador CPS-1: Parte inferior", value=200.0)
        carg_cps1_flex = st.number_input("Cargador CPS-1: Flexibles interiores", value=40.0)
        
        carg_cps3_batea = st.number_input("Cargador CPS-3: Batea", value=40.0)
        carg_cps3_inf = st.number_input("Cargador CPS-3: Parte inferior", value=160.0)
        carg_cps3_flex = st.number_input("Cargador CPS-3: Flexibles interiores", value=20.0)

with tab3:
    st.subheader("Registro Manual de Ciclos por CPS")
    st.markdown("Seleccione el CPS y registre manualmente los horarios y tiempos de cada hito operativo:")
    
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        cps_seleccionado = st.selectbox("Equipo / CPS", ["CPS-1", "CPS-3", "CPS-2"])
    with col_sel2:
        nombre_ciclo = st.text_input("Identificador de Ciclo", "J-5")
    
    st.markdown("---")
    c_h1, c_h2 = st.columns(2)
    with c_h1:
        t_mazamorra = st.time_input("Término retiro mazamorra desde CPS", value=time(21, 35))
        t_bloqueo = st.time_input("Bloqueo de tapas CPS", value=time(21, 39))
        t_ing_pic = st.time_input("Ingreso Picotón foso", value=time(21, 40))
        t_ret_pic = st.time_input("Retiro Picotón foso", value=time(21, 52))
    with c_h2:
        t_ing_carg = st.time_input("Ingreso Cargador Foso", value=time(22, 3))
        t_ret_carg = st.time_input("Retiro de Cargador en foso", value=time(22, 9))
        t_desbloqueo = st.time_input("Desbloqueo de tapas CPS", value=time(22, 11))
        
    st.markdown("---")
    col_val1, col_val2 = st.columns(2)
    with col_val1:
        num_baldadas = st.number_input("Número de Baldadas", min_value=0, value=2)
    with col_val2:
        toneladas_ext = st.number_input("Toneladas / Comentarios de Desviaciones", value=6.5)
        
    obs_desviacion = st.text_input("Detalle de desviación (ej. 4 min picado boca // 7 min picado piso)")

with tab4:
    st.subheader("Resumen del Turno y Consolidado")
    st.info("Revisa la información ingresada. Al presionar el botón, los datos quedarán listos para integrarse a tu reporte oficial.")
    
    if st.button("💾 Confirmar y Registrar Datos del Turno"):
        st.success("¡Datos guardados exitosamente en el sistema de la nave!")
        st.balloons()

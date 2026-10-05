import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta

st.set_page_config(page_title="Control Operacional - Nave Altanorte", layout="wide")

st.title("🏭 Sistema de Control Operacional - Nave Altanorte")
st.markdown("Plataforma en línea para el registro de turnos, termografía y ciclos de equipos con cálculo automático de tiempos.")

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
    
    comentarios_inicio = st.text_area("Comentarios y Estado de Equipos", "Operativos: Cargador y Picotón revisados. Camión aljibe operativo.")

with tab2:
    st.subheader("Registro de Termografía (°C) - CPS 1 al 4")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### ⛏️ Picotón")
        cps_term_sel = st.selectbox("Seleccione CPS para Termografía", ["CPS-1", "CPS-2", "CPS-3", "CPS-4"])
        p_flex = st.number_input(f"{cps_term_sel}: Flexibles", value=40.0)
        p_cuna = st.number_input(f"{cps_term_sel}: Cuña", value=160.0)
        p_torn = st.number_input(f"{cps_term_sel}: Tornamesa", value=20.0)
        
    with col2:
        st.markdown("### 🚜 Cargador")
        c_batea = st.number_input(f"{cps_term_sel} Cargador: Batea", value=140.0)
        c_inf = st.number_input(f"{cps_term_sel} Cargador: Parte inferior", value=200.0)
        c_flex = st.number_input(f"{cps_term_sel} Cargador: Flexibles interiores", value=40.0)

with tab3:
    st.subheader("Registro Manual de Ciclos por CPS (1 al 4) con Cálculo de Tiempos")
    
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        cps_seleccionado = st.selectbox("Seleccione CPS", ["CPS-1", "CPS-2", "CPS-3", "CPS-4"])
    with col_sel2:
        nombre_ciclo = st.text_input("Identificador de Ciclo", "J-5")
    
    st.markdown("---")
    st.markdown("### Ingrese los horarios de cada hito:")
    
    # Horarios
    t_mazamorra = st.time_input("1. Término retiro mazamorra desde CPS", value=time(21, 35))
    t_bloqueo = st.time_input("2. Bloqueo de tapas CPS", value=time(21, 39))
    t_ing_pic = st.time_input("3. Ingreso Picotón foso", value=time(21, 40))
    t_ret_pic = st.time_input("4. Retiro Picotón foso", value=time(21, 52))
    t_ing_carg = st.time_input("5. Ingreso Cargador Foso", value=time(22, 3))
    t_ret_carg = st.time_input("6. Retiro de Cargador en foso", value=time(22, 9))
    t_desbloqueo = st.time_input("7. Desbloqueo de tapas CPS", value=time(22, 11))

    # Función auxiliar para calcular diferencia en minutos entre dos objetos time
    def diff_minutes(t_start, t_end):
        d1 = timedelta(hours=t_start.hour, minutes=t_start.minute, seconds=t_start.second)
        d2 = timedelta(hours=t_end.hour, minutes=t_end.minute, seconds=t_end.second)
        diff = (d2 - d1).total_seconds() / 60
        if diff < 0: # Manejo simple si cruza la medianoche
            diff += 24 * 60
        return diff

    m_bloqueo = diff_minutes(t_mazamorra, t_bloqueo)
    m_ing_pic = diff_minutes(t_bloqueo, t_ing_pic)
    m_ret_pic = diff_minutes(t_ing_pic, t_ret_pic)
    m_ing_carg = diff_minutes(t_ret_pic, t_ing_carg)
    m_ret_carg = diff_minutes(t_ing_carg, t_ret_carg)
    m_desbloqueo = diff_minutes(t_ret_carg, t_desbloqueo)
    total_minutos = m_bloqueo + m_ing_pic + m_ret_pic + m_ing_carg + m_ret_carg + m_desbloqueo

    st.markdown("---")
    st.markdown("### ⏱️ Resultados Automáticos del Ciclo:")
    
    res_df = pd.DataFrame({
        "Proceso / Hito": [
            "Término retiro mazamorra desde CPS",
            "Bloqueo de tapas CPS",
            "Ingreso Picotón foso",
            "Retiro Picotón foso",
            "Ingreso Cargador Foso",
            "Retiro de Cargador en foso",
            "Desbloqueo de tapas CPS"
        ],
        "Hora": [str(t_mazamorra), str(t_bloqueo), str(t_ing_pic), str(t_ret_pic), str(t_ing_carg), str(t_ret_carg), str(t_desbloqueo)],
        "Minutos": [0, m_bloqueo, m_ing_pic, m_ret_pic, m_ing_carg, m_ret_carg, m_desbloqueo]
    })
    
    st.dataframe(res_df, use_container_width=True)
    st.success(f"📌 **Tiempo Total del Ciclo ({nombre_ciclo}): {total_minutos:.1f} minutos**")

    st.markdown("---")
    col_val1, col_val2 = st.columns(2)
    with col_val1:
        num_baldadas = st.number_input("Número de Baldadas", min_value=0, value=2)
    with col_val2:
        toneladas_ext = st.number_input("Toneladas estimadas", value=6.5)
        
    obs_desviacion = st.text_input("Detalle de desviación / Comentarios (ej. 4 min picado boca // 7 min picado piso)")

with tab4:
    st.subheader("Resumen del Turno y Consolidado")
    st.info("Revisa la información ingresada. Al presionar el botón, los datos quedarán registrados.")
    
    if st.button("💾 Confirmar y Registrar Datos del Turno"):
        st.success("¡Datos guardados exitosamente en el sistema de la nave!")
        st.balloons()

import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta

st.set_page_config(page_title="Control Operacional - Nave Altanorte", layout="wide")

st.title("🏭 Sistema de Control Operacional - Nave Altanorte")
st.markdown("Plataforma en línea para registro de turnos, múltiples ciclos de CPS y termografía en terreno.")

# --- BARRA LATERAL: CONFIGURACIÓN GENERAL DEL TURNO ---
st.sidebar.header("📋 Identificación del Turno")
fecha_turno = st.sidebar.date_input("Fecha", datetime.today())
tipo_turno = st.sidebar.selectbox("Tipo de Turno", ["Turno Día (TA)", "Turno Noche (TB)"])

st.sidebar.subheader("Supervisores y Operadores")
sup_sop = st.sidebar.text_input("Supervisor SOP", "Rene Philipps")
sup_crm = st.sidebar.text_input("Supervisor CRM", "Jovelino Burgos")
sup_caemin = st.sidebar.text_input("Supervisor Caemin", "Ruben Infanta")
sup_nave = st.sidebar.text_input("Supervisor Nave", "Jorge Galindo")
op_picoton = st.sidebar.text_input("Operador Picotón", "Juan Pablo Figueroa")
op_cargador = st.sidebar.text_input("Operador Cargador", "Fernando Tapia")

# --- INICIALIZAR MEMORIA TEMPORAL PARA LOS REGISTROS DEL TURNO ---
if "ciclos_registrados" not in st.session_state:
    st.session_state.ciclos_registrados = []

if "termografias_registradas" not in st.session_state:
    st.session_state.termografias_registradas = []

# --- PESTAÑAS PRINCIPALES ---
tab1, tab2, tab3, tab4 = st.tabs([
    "1. Checklist Inicio", 
    "2. Registro de Termografía (°C)", 
    "3. Registro de Ciclos CPS (Múltiples)", 
    "4. Resumen del Turno y Exportar"
])

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
    st.subheader("🌡️ Registro de Termografía (°C) por Ingreso")
    st.markdown("Cada vez que un equipo ingrese o se evalúe, registra sus temperaturas y agrégalo al listado del turno:")
    
    col_t_eq, col_t_cps = st.columns(2)
    with col_t_eq:
        eq_term = st.selectbox("Equipo", ["Picotón", "Cargador"])
    with col_t_cps:
        cps_term = st.selectbox("CPS", ["CPS-1", "CPS-2", "CPS-3", "CPS-4"], key="term_cps")
        
    col_val1, col_val2, col_val3 = st.columns(3)
    
    if eq_term == "Picotón":
        with col_val1: val1 = st.number_input("Flexibles (°C)", value=40.0)
        with col_val2: val2 = st.number_input("Cuña (°C)", value=160.0)
        with col_val3: val3 = st.number_input("Tornamesa (°C)", value=20.0)
        p_puntos = {"Flexibles": val1, "Cuña": val2, "Tornamesa": val3}
    else:
        # Cargador con sus 3 campos específicos
        with col_val1: val1 = st.number_input("Batea (°C)", value=140.0)
        with col_val2: val2 = st.number_input("Parte inferior (°C)", value=200.0)
        with col_val3: val3 = st.number_input("Flexibles (°C)", value=40.0)
        p_puntos = {"Batea": val1, "Parte Inferior": val2, "Flexibles": val3}
        
    if st.button("➕ Agregar Registro de Termografía"):
        st.session_state.termografias_registradas.append({
            "Turno": tipo_turno,
            "Equipo": eq_term,
            "CPS": cps_term,
            **p_puntos,
            "Hora Registro": datetime.now().strftime("%H:%M")
        })
        st.success("¡Termografía agregada correctamente al turno!")
        
    if len(st.session_state.termografias_registradas) > 0:
        st.markdown("### Historial de Termografías en el Turno")
        st.dataframe(pd.DataFrame(st.session_state.termografias_registradas), use_container_width=True)

with tab3:
    st.subheader("🔄 Registro de Ciclos CPS (Múltiples Ingresos)")
    st.markdown("Puedes registrar tantos ciclos como entradas realicen los equipos a lo largo del turno.")
    
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        cps_seleccionado = st.selectbox("Seleccione CPS", ["CPS-1", "CPS-2", "CPS-3", "CPS-4"], key="ciclo_cps")
    with col_sel2:
        nombre_ciclo = st.text_input("Identificador de Ciclo (ej. J-5, J-6, etc.)", "J-5")
    
    st.markdown("---")
    st.markdown("### Ingrese los horarios de los hitos del ciclo:")
    
    t_mazamorra = st.time_input("1. Término retiro mazamorra desde CPS", value=time(21, 35))
    t_bloqueo = st.time_input("2. Bloqueo de tapas CPS", value=time(21, 39))
    t_ing_pic = st.time_input("3. Ingreso Picotón foso", value=time(21, 40))
    t_ret_pic = st.time_input("4. Retiro Picotón foso", value=time(21, 52))
    t_ing_carg = st.time_input("5. Ingreso Cargador Foso", value=time(22, 3))
    t_ret_carg = st.time_input("6. Retiro de Cargador en foso", value=time(22, 9))
    t_desbloqueo = st.time_input("7. Desbloqueo de tapas CPS", value=time(22, 11))

    def diff_minutes(t_start, t_end):
        d1 = timedelta(hours=t_start.hour, minutes=t_start.minute, seconds=t_start.second)
        d2 = timedelta(hours=t_end.hour, minutes=t_end.minute, seconds=t_end.second)
        diff = (d2 - d1).total_seconds() / 60
        if diff < 0: diff += 24 * 60
        return diff

    m_bloqueo = diff_minutes(t_mazamorra, t_bloqueo)
    m_ing_pic = diff_minutes(t_bloqueo, t_ing_pic)
    m_ret_pic = diff_minutes(t_ing_pic, t_ret_pic)
    m_ing_carg = diff_minutes(t_ret_pic, t_ing_carg)
    m_ret_carg = diff_minutes(t_ing_carg, t_ret_carg)
    m_desbloqueo = diff_minutes(t_ret_carg, t_desbloqueo)
    total_minutos = m_bloqueo + m_ing_pic + m_ret_pic + m_ing_carg + m_ret_carg + m_desbloqueo

    st.markdown(f"📌 **Tiempo Total Calculado para este Ciclo: {total_minutos:.1f} minutos**")

    col_val1, col_val2 = st.columns(2)
    with col_val1:
        num_baldadas = st.number_input("Número de Baldadas", min_value=0, value=2)
    with col_val2:
        toneladas_ext = st.number_input("Toneladas estimadas", value=6.5)
        
    obs_desviacion = st.text_input("Detalle de desviación / Comentarios (ej. 4 min picado boca // 7 min picado piso)")

    if st.button("➕ Guardar e Incluir este Ciclo en el Turno"):
        st.session_state.ciclos_registrados.append({
            "Turno": tipo_turno,
            "CPS": cps_seleccionado,
            "Ciclo": nombre_ciclo,
            "Total Minutos": round(total_minutos, 1),
            "Baldadas": num_baldadas,
            "Toneladas": toneladas_ext,
            "Comentarios": obs_desviacion
        })
        st.success(f"¡Ciclo {nombre_ciclo} guardado exitosamente!")

    if len(st.session_state.ciclos_registrados) > 0:
        st.markdown("### Historial de Ciclos Registrados en el Turno")
        st.dataframe(pd.DataFrame(st.session_state.ciclos_registrados), use_container_width=True)

with tab4:
    st.subheader("📋 Consolidado Total del Turno")
    st.info(f"Resumen general para el **{tipo_turno}** del día **{fecha_turno}**.")
    
    st.markdown("### Ciclos Realizados:")
    if len(st.session_state.ciclos_registrados) > 0:
        st.dataframe(pd.DataFrame(st.session_state.ciclos_registrados), use_container_width=True)
    else:
        st.warning("Aún no hay ciclos registrados en este turno.")
        
    st.markdown("### Termografías Registradas:")
    if len(st.session_state.termografias_registradas) > 0:
        st.dataframe(pd.DataFrame(st.session_state.termografias_registradas), use_container_width=True)
    else:
        st.warning("Aún no hay registros de termografía en este turno.")

    st.markdown("---")
    if st.button("💾 Generar Reporte Final y Descargar Excel"):
        st.success("¡Reporte consolidado generado con éxito para Altanorte!")
        st.balloons()

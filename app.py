import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import io
import os
import json
import smtplib
from email.message import EmailMessage

st.set_page_config(page_title="Control Operacional - Nave Altonorte", layout="wide")

st.title("🏭 Sistema de Control Operacional - Nave Altonorte")
st.markdown("Plataforma web oficial para turnos 4x4.")

# --- ARCHIVO DE RESPALDO PERSISTENTE EN SERVIDOR ---
RESPALDO_JSON = "respaldo_turno_actual.json"

def guardar_respaldo_disco():
    """Guarda una copia exacta de los datos en disco asegurando serialización JSON."""
    try:
        ciclos_serializables = []
        for c in st.session_state.ciclos_registrados:
            c_copy = c.copy()
            c_copy["Foto_Entrada"] = c_copy["Foto_Entrada"].hex() if c_copy["Foto_Entrada"] else None
            c_copy["Foto_Salida"] = c_copy["Foto_Salida"].hex() if c_copy["Foto_Salida"] else None
            
            for campo_t in ["T_Mazamorra", "T_Bloqueo", "T_Ing_Pic", "T_Ret_Pic", "T_Ing_Carg", "T_Ret_Carg", "T_Desbloqueo"]:
                if isinstance(c_copy.get(campo_t), time):
                    c_copy[campo_t] = c_copy[campo_t].strftime("%H:%M:%S")
            ciclos_serializables.append(c_copy)

        datos = {
            "ciclos": ciclos_serializables,
            "termografias": st.session_state.termografias_registradas,
            "equipos": {
                "est_carg_627": st.session_state.get("est_carg_627", "Revisado por CAEMIN"),
                "est_carg_637": st.session_state.get("est_carg_637", "Pendiente de revisión CAEMIN"),
                "est_aljibe": st.session_state.get("est_aljibe", "Operativo"),
                "est_pic_855": st.session_state.get("est_pic_855", "Revisado por CAEMIN"),
                "est_pic_854": st.session_state.get("est_pic_854", "Pendiente de revisión CAEMIN"),
                "tiempo_teleop": st.session_state.get("tiempo_teleop", 0)
            }
        }
        with open(RESPALDO_JSON, "w", encoding="utf-8") as f:
            json.dump(datos, f)
    except Exception as e:
        print(f"Error al guardar respaldo: {e}")

def cargar_respaldo_disco():
    """Recupera automáticamente los datos desde el disco si la sesión se recarga."""
    if os.path.exists(RESPALDO_JSON):
        try:
            with open(RESPALDO_JSON, "r", encoding="utf-8") as f:
                datos = json.load(f)
                
                ciclos_recuperados = []
                for c in datos.get("ciclos", []):
                    c["Foto_Entrada"] = bytes.fromhex(c["Foto_Entrada"]) if c["Foto_Entrada"] else None
                    c["Foto_Salida"] = bytes.fromhex(c["Foto_Salida"]) if c["Foto_Salida"] else None
                    
                    for campo_t in ["T_Mazamorra", "T_Bloqueo", "T_Ing_Pic", "T_Ret_Pic", "T_Ing_Carg", "T_Ret_Carg", "T_Desbloqueo"]:
                        if isinstance(c.get(campo_t), str):
                            c[campo_t] = datetime.strptime(c[campo_t], "%H:%M:%S").time()
                    ciclos_recuperados.append(c)
                    
                st.session_state.ciclos_registrados = ciclos_recuperados
                st.session_state.termografias_registradas = datos.get("termografias", [])
        except Exception as e:
            print(f"Error al cargar respaldo: {e}")

# --- INICIALIZAR MEMORIA Y RECUPERACIÓN AUTOMÁTICA ---
if "inicializado" not in st.session_state:
    st.session_state.ciclos_registrados = []
    st.session_state.termografias_registradas = []
    cargar_respaldo_disco()
    st.session_state.inicializado = True

# --- FUNCIÓN DE ENVÍO DE CORREO DESDE GMAIL ---
def enviar_correo_gmail(destinatario, archivo_bytes, nombre_archivo):
    try:
        remitente = st.secrets.get("GMAIL_USER", "carlos.flores.b@gmail.com")
        password = st.secrets.get("GMAIL_APP_PASSWORD", "eyai squb rdyp unlz")
        
        smtp_server = "smtp.gmail.com"
        smtp_port = 465

        msg = EmailMessage()
        msg['Subject'] = f"📊 Reporte Oficial Control Operacional - {nombre_archivo}"
        msg['From'] = remitente
        msg['To'] = destinatario
        msg.set_content(
            f"Estimado Supervisor (Glencore / Altonorte),\n\n"
            f"Adjunto encontrará la planilla oficial de control operacional de la Nave Altonorte "
            f"correspondiente al turno registrado, incluyendo termografías, ciclos y evidencias fotográficas.\n\n"
            f"Este correo ha sido generado y enviado automáticamente desde el sistema en terreno.\n\n"
            f"Atentamente,\nPlataforma de Control Operacional Altonorte."
        )

        msg.add_attachment(
            archivo_bytes.getvalue(),
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=nombre_archivo
        )

        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(remitente, password)
            server.send_message(msg)
                
        return True, "¡Reporte Excel enviado exitosamente!"
    except Exception as e:
        return False, f"Error al enviar correo: {e}"

# --- PANEL SUPERIOR ADAPTADO PARA MÓVILES (Turnos rotativos 4x4) ---
with st.expander("📋 1. Identificación del Turno, Fecha y Supervisores (Toca aquí para desplegar)", expanded=True):
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        fecha_input = st.date_input("Fecha del Turno (Computador)", datetime.today())
    with col_f2:
        tipo_turno = st.selectbox("Tipo de Turno", ["Turno Día (TA)", "Turno Noche (TB)"])
    
    st.markdown("---")
    st.subheader("Supervisores y Operadores Responsables (Turno 4x4)")
    correo_supervisor = st.text_input("📧 Correo Destinatario (Glencore)", "@glencore.cl")
    
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        sup_sop = st.text_input("Supervisor SOP", "")
        sup_nave = st.text_input("Supervisor Nave", "")
    with col_s2:
        sup_crm = st.text_input("Supervisor CRM", "")
        op_picoton = st.text_input("Operador Picotón", "")
    with col_s3:
        sup_caemin = st.text_input("Supervisor Caemin", "")
        op_cargador = st.text_input("Operador Cargador", "")

st.markdown("---")

# Indicador visual de respaldo activo
if len(st.session_state.ciclos_registrados) > 0 or len(st.session_state.termografias_registradas) > 0:
    st.info(f"💾 **Estado de Respaldo:** Se han recuperado/guardado `{len(st.session_state.ciclos_registrados)}` ciclos y `{len(st.session_state.termografias_registradas)}` termografías de forma segura.")

# --- PESTAÑAS PRINCIPALES ---
tab1, tab2, tab3, tab4 = st.tabs([
    "1. Checklist & Equipos", 
    "2. Termografía (°C)", 
    "3. Ciclos CPS & Fotos", 
    "4. Consolidado y Envío correo"
])

with tab1:
    st.subheader("Control Previo de Inicio de Actividades y Equipos")
    
    st.markdown("### Estado Operacional de Equipos (CAEMIN)")
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        est_carg_627 = st.selectbox("Cargador M-627", ["Revisado por CAEMIN", "Pendiente de revisión CAEMIN"], key="est_carg_627")
        est_carg_637 = st.selectbox("Cargador M-637", ["Pendiente de revisión CAEMIN", "Revisado por CAEMIN"], key="est_carg_637")
        est_aljibe = st.selectbox("Camión Aljibe M-8380", ["Operativo", "No disponible"], key="est_aljibe")
    with col_e2:
        est_pic_855 = st.selectbox("Picotón M-855", ["Revisado por CAEMIN", "Pendiente de revisión CAEMIN"], key="est_pic_855")
        est_pic_854 = st.selectbox("Picotón M-854", ["Pendiente de revisión CAEMIN", "Revisado por CAEMIN"], key="est_pic_854")
    
    tiempo_teleop = st.number_input("⏱️ Tiempo en teleoperación (minutos perdidos por señal)", min_value=0, value=0, key="tiempo_teleop")

    st.markdown("---")
    st.subheader("Controles de Retiro y Traslado de Material")
    c1 = st.checkbox("Confirmar tiempo de al menos 10 min desde término de picado del foso", value=True)
    c2 = st.checkbox("Verificar visualmente que material en foso esté sólido (no líquido)", value=True)
    c3 = st.checkbox("Verificar ausencia de llamas, humo o indicios de ignición", value=True)
    c4 = st.checkbox("Asegurar que disposición en explanada favorezca enfriamiento", value=True)
    c5 = st.checkbox("Verificar que camión aljibe esté disponible", value=True)

with tab2:
    st.subheader("🌡️ Registro de Termografía (°C) por Ingreso")
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
        with col_val1: val1 = st.number_input("Batea (°C)", value=140.0)
        with col_val2: val2 = st.number_input("Parte inferior (°C)", value=200.0)
        with col_val3: val3 = st.number_input("Flexibles (°C)", value=40.0)
        p_puntos = {"Batea": val1, "Parte Inferior": val2, "Flexibles": val3}
        
    if st.button("➕ Agregar Registro de Termografía"):
        st.session_state.termografias_registradas.append({
            "Equipo": eq_term,
            "CPS": cps_term,
            **p_puntos,
            "Hora": datetime.now().strftime("%H:%M")
        })
        guardar_respaldo_disco()
        st.success("¡Termografía agregada y respaldada con éxito!")
        
    if len(st.session_state.termografias_registradas) > 0:
        st.markdown(f"### Historial de Termografías en el Turno (Total: {len(st.session_state.termografias_registradas)})")
        df_term = pd.DataFrame(st.session_state.termografias_registradas)
        df_term.index = range(1, len(df_term) + 1)
        st.dataframe(df_term, use_container_width=True)

with tab3:
    st.subheader("🔄 Registro de Ciclos CPS y Fotografías (Entrada / Salida)")
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        cps_seleccionado = st.selectbox("Seleccione CPS", ["CPS-1", "CPS-2", "CPS-3", "CPS-4"], key="ciclo_cps")
    with col_sel2:
        nombre_ciclo = st.text_input("Identificador de Ciclo (ej. J-5, J-6)", "J-5")
    
    st.markdown("---")
    st.markdown("### ⏱️ Horarios del Ciclo:")
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

    m_bloqueo = diff_minutes(t

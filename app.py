import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage, ImageOps
import io
import os
import json
import base64
import smtplib
from email.message import EmailMessage

st.set_page_config(page_title="Control Operacional - Nave Altonorte", layout="wide")

st.title("🏭 Sistema de Control Operacional - Nave Altonorte")
st.markdown("Plataforma web oficial para turnos 4x4 con **persistencia total, formato de fecha Altonorte y envío por Gmail**.")

# --- ARCHIVO DE RESPALDO PERSISTENTE EN SERVIDOR ---
RESPALDO_JSON = "respaldo_turno_actual.json"

def guardar_respaldo_disco():
    """Guarda una copia exacta de los datos, supervisores y fotos (en Base64) en disco."""
    try:
        ciclos_serializables = []
        for c in st.session_state.ciclos_registrados:
            c_copy = c.copy()
            if c_copy.get("Foto_Entrada"):
                c_copy["Foto_Entrada"] = base64.b64encode(c_copy["Foto_Entrada"]).decode('utf-8')
            else:
                c_copy["Foto_Entrada"] = None

            if c_copy.get("Foto_Salida"):
                c_copy["Foto_Salida"] = base64.b64encode(c_copy["Foto_Salida"]).decode('utf-8')
            else:
                c_copy["Foto_Salida"] = None
            
            for campo_t in ["T_Mazamorra", "T_Bloqueo", "T_Ing_Pic", "T_Ret_Pic", "T_Ing_Carg", "T_Ret_Carg", "T_Desbloqueo"]:
                if isinstance(c_copy.get(campo_t), time):
                    c_copy[campo_t] = c_copy[campo_t].strftime("%H:%M:%S")
            ciclos_serializables.append(c_copy)

        datos = {
            "ciclos": ciclos_serializables,
            "termografias": st.session_state.termografias_registradas,
            "supervisores": {
                "correo_supervisor": st.session_state.get("correo_supervisor", ""),
                "sup_sop": st.session_state.get("sup_sop", ""),
                "sup_nave": st.session_state.get("sup_nave", ""),
                "sup_crm": st.session_state.get("sup_crm", ""),
                "op_picoton": st.session_state.get("op_picoton", ""),
                "sup_caemin": st.session_state.get("sup_caemin", ""),
                "op_cargador": st.session_state.get("op_cargador", ""),
                "num_regador": st.session_state.get("num_regador", "M-8358")
            },
            "equipos": {
                "est_carg_627": st.session_state.get("est_carg_627", "Revisado por CAEMIN"),
                "est_carg_637": st.session_state.get("est_carg_637", "Pendiente de revisión CAEMIN"),
                "est_aljibe": st.session_state.get("est_aljibe", "Operativo"),
                "est_pic_855": st.session_state.get("est_pic_855", "Revisado por CAEMIN"),
                "est_pic_854": st.session_state.get("est_pic_854", "Pendiente de revisión CAEMIN"),
                "tiempo_teleop": st.session_state.get("tiempo_teleop", 0)
            },
            "comentarios_controles": {
                "obs_ctrl_10": st.session_state.get("obs_ctrl_10", ""),
                "obs_ctrl_11": st.session_state.get("obs_ctrl_11", ""),
                "obs_ctrl_12": st.session_state.get("obs_ctrl_12", ""),
                "obs_ctrl_13": st.session_state.get("obs_ctrl_13", ""),
                "obs_ctrl_14": st.session_state.get("obs_ctrl_14", "")
            }
        }
        with open(RESPALDO_JSON, "w", encoding="utf-8") as f:
            json.dump(datos, f)
    except Exception as e:
        print(f"Error al guardar respaldo: {e}")

def cargar_respaldo_disco():
    """Recupera automáticamente los datos, supervisores y fotos desde el disco reconvirtiendo Base64."""
    if os.path.exists(RESPALDO_JSON):
        try:
            with open(RESPALDO_JSON, "r", encoding="utf-8") as f:
                datos = json.load(f)
                
                ciclos_recuperados = []
                for c in datos.get("ciclos", []):
                    if c.get("Foto_Entrada"):
                        try:
                            c["Foto_Entrada"] = base64.b64decode(c["Foto_Entrada"])
                        except Exception:
                            c["Foto_Entrada"] = None
                    else:
                        c["Foto_Entrada"] = None

                    if c.get("Foto_Salida"):
                        try:
                            c["Foto_Salida"] = base64.b64decode(c["Foto_Salida"])
                        except Exception:
                            c["Foto_Salida"] = None
                    else:
                        c["Foto_Salida"] = None
                    
                    for campo_t in ["T_Mazamorra", "T_Bloqueo", "T_Ing_Pic", "T_Ret_Pic", "T_Ing_Carg", "T_Ret_Carg", "T_Desbloqueo"]:
                        if isinstance(c.get(campo_t), str):
                            c[campo_t] = datetime.strptime(c[campo_t], "%H:%M:%S").time()
                    ciclos_recuperados.append(c)
                    
                st.session_state.ciclos_registrados = ciclos_recuperados
                st.session_state.termografias_registradas = datos.get("termografias", [])
                
                sup_data = datos.get("supervisores", {})
                for k, v in sup_data.items():
                    st.session_state[k] = v

                ctrl_data = datos.get("comentarios_controles", {})
                for k, v in ctrl_data.items():
                    st.session_state[k] = v
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

# --- FUNCIÓN PARA COMPRIMIR Y OPTIMIZAR FOTOGRAFÍAS EN MEMORIA ---
def optimizar_imagen(img_bytes, max_width=800, quality=75):
    """Comprime y redimensiona la imagen para evitar correos pesados (máx ~100KB por foto)."""
    try:
        img = PILImage.open(io.BytesIO(img_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
            
        if img.width > max_width:
            ratio = max_width / float(img.width)
            new_height = int(float(img.height) * ratio)
            img = img.resize((max_width, new_height), PILImage.Resampling.LANCZOS)
            
        img_con_marco = ImageOps.expand(img, border=6, fill='black')
        
        buf = io.BytesIO()
        img_con_marco.save(buf, format="JPEG", quality=quality)
        buf.seek(0)
        return buf
    except Exception:
        return io.BytesIO(img_bytes)

# --- PANEL SUPERIOR ADAPTADO PARA MÓVILES ---
with st.expander("📋 1. Identificación del Turno, Fecha y Supervisores (Toca aquí para desplegar)", expanded=True):
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        fecha_input = st.date_input("Fecha del Turno (Computador)", datetime.today())
    with col_f2:
        tipo_turno = st.selectbox("Tipo de Turno", ["Turno Día (TA)", "Turno Noche (TB)"])
    
    st.markdown("---")
    st.subheader("Supervisores y Operadores Responsables (Turno 4x4)")
    correo_supervisor = st.text_input("📧 Correo Destinatario (Glencore)", value=st.session_state.get("correo_supervisor", "@glencore.cl"), key="correo_supervisor", on_change=guardar_respaldo_disco)
    
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        sup_sop = st.text_input("Supervisor SOP", value=st.session_state.get("sup_sop", ""), key="sup_sop", on_change=guardar_respaldo_disco)
        sup_nave = st.text_input("Supervisor Nave", value=st.session_state.get("sup_nave", ""), key="sup_nave", on_change=guardar_respaldo_disco)
    with col_s2:
        sup_crm = st.text_input("Supervisor CRM", value=st.session_state.get("sup_crm", ""), key="sup_crm", on_change=guardar_respaldo_disco)
        op_picoton = st.text_input("Operador Picotón", value=st.session_state.get("op_picoton", ""), key="op_picoton", on_change=guardar_respaldo_disco)
    with col_s3:
        sup_caemin = st.text_input("Supervisor Caemin", value=st.session_state.get("sup_caemin", ""), key="sup_caemin", on_change=guardar_respaldo_disco)
        op_cargador = st.text_input("Operador Cargador", value=st.session_state.get("op_cargador", ""), key="op_cargador", on_change=guardar_respaldo_disco)

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
        est_carg_627 = st.selectbox("Cargador M-627", ["Revisado por CAEMIN", "Pendiente de revisión CAEMIN", "Fuera de servicio"], key="est_carg_627")
        est_carg_637 = st.selectbox("Cargador M-637", ["Pendiente de revisión CAEMIN", "Revisado por CAEMIN", "Fuera de servicio"], key="est_carg_637")
        num_regador = st.text_input("🚰 Número de Regador / Aljibe", value=st.session_state.get("num_regador", "M-8358"), key="num_regador", on_change=guardar_respaldo_disco)
        est_aljibe = st.selectbox(f"Camión Aljibe ({num_regador})", ["Operativo", "No disponible", "Fuera de servicio"], key="est_aljibe")
    with col_e2:
        est_pic_855 = st.selectbox("Picotón M-855", ["Revisado por CAEMIN", "Pendiente de revisión CAEMIN", "Fuera de servicio"], key="est_pic_855")
        est_pic_854 = st.selectbox("Picotón M-854", ["Pendiente de revisión CAEMIN", "Revisado por CAEMIN", "Fuera de servicio"], key="est_pic_854")
    
    tiempo_teleop = st.number_input("⏱️ Tiempo en teleoperación (minutos perdidos por señal)", min_value=0, value=0, key="tiempo_teleop")

    st.markdown("---")
    st.subheader("Controles de Retiro y Traslado de Material (Comentarios por Actividad)")
    
    obs_ctrl_10 = st.text_input("Observación Control 10 (Confirmar tiempo de al menos 10 min desde término de picado)", 
                                value=st.session_state.get("obs_ctrl_10", "Se evalua el ingreso inmediato de cargador, por disminucion de material en foso."), key="obs_ctrl_10", on_change=guardar_respaldo_disco)
    obs_ctrl_11 = st.text_input("Observación Control 11 (Verificar visualmente que material esté sólido)", 
                                value=st.session_state.get("obs_ctrl_11", "No se detecta material en estado líquido."), key="obs_ctrl_11", on_change=guardar_respaldo_disco)
    obs_ctrl_12 = st.text_input("Observación Control 12 (Verificar ausencia de llamas, humo o ignición)", 
                                value=st.session_state.get("obs_ctrl_12", "No se aprecian fosos con materiales liquidos ni presencia de llamas en equipos"), key="obs_ctrl_12", on_change=guardar_respaldo_disco)
    obs_ctrl_13 = st.text_input("Observación Control 13 (Disposición en explanada favorezca enfriamiento)", 
                                value=st.session_state.get("obs_ctrl_13", "Se mantiene area despejada, se realiza retiro de material de explanada"), key="obs_ctrl_13", on_change=guardar_respaldo_disco)
    obs_ctrl_14 = st.text_input("Observación Control 14 (Camión aljibe disponible durante la actividad)", 
                                value=st.session_state.get("obs_ctrl_14", f"Camión Aljibe {num_regador} {est_aljibe.lower()}"), key="obs_ctrl_14", on_change=guardar_respaldo_disco)

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
        nombre_ciclo = st.text_input("Identificador de Ciclo (ej. J-117, J-118)", "J-117")
    
    st.markdown("---")
    st.markdown("### ⏱️ Horarios del Ciclo:")
    t_mazamorra = st.time_input("1. Término retiro mazamorra desde CPS", value=time(20, 25))
    t_bloqueo = st.time_input("2. Bloqueo de tapas CPS", value=time(20, 31))
    t_ing_pic = st.time_input("3. Ingreso Picotón foso", value=time(20, 32))
    t_ret_pic = st.time_input("4. Retiro Picotón foso", value=time(20, 40))
    t_ing_carg = st.time_input("5. Ingreso Cargador Foso", value=time(20, 42))
    t_ret_carg = st.time_input("6. Retiro de Cargador en foso", value=time(20, 48))
    t_desbloqueo = st.time_input("7. Desbloqueo de tapas CPS", value=time(20, 53))

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

    st.markdown(f"📌 **Tiempo Total Calculado: {total_minutos:.1f} minutos**")

    col_val1, col_val2 = st.columns(2)
    with col_val1: num_baldadas = st.number_input("Número de Baldadas", min_value=0, value=2)
    with col_val2: toneladas_ext = st.number_input("Toneladas estimadas / Temp Material", value=8.0)
        
    obs_desviacion = st.text_input("Detalle de desviación / Comentarios del Ciclo", key=f"obs_ciclo_{nombre_ciclo}")

    st.markdown("---")
    st.markdown(f"### 📸 Evidencia Fotográfica ({cps_seleccionado} - {nombre_ciclo})")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        foto_entrada = st.file_uploader(f"Foto ANTES del Ingreso ({cps_seleccionado})", type=["jpg", "jpeg", "png"], key=f"f_ent_{nombre_ciclo}")
        if foto_entrada:
            st.image(foto_entrada, caption=f"Entrada {cps_seleccionado} - {nombre_ciclo}", width=200)
    with col_f2:
        foto_salida = st.file_uploader(f"Foto a la SALIDA ({cps_seleccionado})", type=["jpg", "jpeg", "png"], key=f"f_sal_{nombre_ciclo}")
        if foto_salida:
            st.image(foto_salida, caption=f"Salida {cps_seleccionado} - {nombre_ciclo}", width=200)

    if st.button("➕ Guardar Ciclo y Sus Fotografías en el Turno"):
        b_ent = foto_entrada.getvalue() if foto_entrada else None
        b_sal = foto_salida.getvalue() if foto_salida else None

        st.session_state.ciclos_registrados.append({
            "CPS": cps_seleccionado,
            "Ciclo": nombre_ciclo,
            "T_Mazamorra": t_mazamorra,
            "T_Bloqueo": t_bloqueo,
            "T_Ing_Pic": t_ing_pic,
            "T_Ret_Pic": t_ret_pic,
            "T_Ing_Carg": t_ing_carg,
            "T_Ret_Carg": t_ret_carg,
            "T_Desbloqueo": t_desbloqueo,
            "Total Minutos": round(total_minutos, 1),
            "Baldadas": num_baldadas,
            "Toneladas": toneladas_ext,
            "Comentarios": obs_desviacion,
            "Foto_Entrada": b_ent,
            "Foto_Salida": b_sal
        })
        guardar_respaldo_disco()
        st.success(f"¡Ciclo {nombre_ciclo} para {cps_seleccionado} guardado y respaldado en disco con éxito!")

    if len(st.session_state.ciclos_registrados) > 0:
        st.markdown("### Historial de Ciclos Registrados en el Turno")
        historial_display = []
        for c in st.session_state.ciclos_registrados:
            dc = {k: v for k, v in c.items() if not k.startswith("Foto")}
            historial_display.append(dc)
        st.dataframe(pd.DataFrame(historial_display), use_container_width=True)

with tab4:
    st.subheader("📋 Consolidado y Envío por correo")
    
    # --- LÓGICA REGLAMENTARIA DE FECHA ALTONORTE ---
    if "Noche" in tipo_turno:
        sufijo_turno = "TB"
        primer_hito_hora = st.session_state.ciclos_registrados[0]["T_Mazamorra"] if len(st.session_state.ciclos_registrados) > 0 else datetime.now().time()
        if isinstance(primer_hito_hora, str):
            primer_hito_hora = datetime.strptime(primer_hito_hora, "%H:%M:%S").time()

        if time(20, 0) <= primer_hito_hora <= time(23, 59):
            fecha_turno_oficial = fecha_input
        else:
            fecha_turno_oficial = fecha_input + timedelta(days=1)
    else:
        sufijo_turno = "TA"
        fecha_turno_oficial = fecha_input

    str_fecha = fecha_turno_oficial.strftime("%d-%m")
    nombre_nueva_hoja = f"{str_fecha} {sufijo_turno}"
    nombre_archivo_salida = f"Control_Nave_{str_fecha}_{sufijo_turno}.xlsx"

    st.info(f"Correo Destinatario (Glencore): **{correo_supervisor}** | Turno Oficial: **{str_fecha} {sufijo_turno}**")
    
    st.markdown(f"💾 **Resumen Respaldado:** `{len(st.session_state.ciclos_registrados)}` ciclos y `{len(st.session_state.termografias_registradas)}` termografías guardadas de forma segura.")
    
    if st.button("🗑️ Finalizar y Limpiar Turno Actual (Borrar Respaldo)"):
        st.session_state.ciclos_registrados = []
        st.session_state.termografias_registradas = []
        if os.path.exists(RESPALDO_JSON):
            os.remove(RESPALDO_JSON)
        st.success("Turno finalizado y memoria limpiada.")
        st.rerun()

    st.markdown("---")
    if st.button("📧 Generar y Enviar Planilla Excel por correo"):
        base_excel = 'plantilla.xlsx'
        
        try:
            if os.path.exists(base_excel):
                wb = openpyxl.load_workbook(base_excel)
                hoja_plantilla = wb.sheetnames[0]
                ws_source = wb[hoja_plantilla]
                
                if nombre_nueva_hoja in wb.sheetnames:
                    del wb[nombre_nueva_hoja]
                
                ws = wb.copy_worksheet(ws_source)
                ws.title = nombre_nueva_hoja
            else:
                st.error("No se encontró el archivo 'plantilla.xlsx' en el repositorio de GitHub.")
                st.stop()

            # 1. Cabecera exacta
            ws['B3'] = f"Fecha: {str_fecha} {sufijo_turno}"
            ws['C4'] = sup_sop
            ws['E4'] = sup_crm
            ws['F4'] = f"Operador Picotón: {op_picoton}" if "Operador" not in op_picoton else op_picoton
            
            ws['C5'] = sup_caemin
            ws['E5'] = sup_nave
            ws['F5'] = f"Operador Cargador: {op_cargador}" if "Operador" not in op_cargador else op_cargador

            # 2. Ticks de estatus (✓)
            for r_chk in range(8, 15):
                ws.cell(row=r_chk, column=5, value="✓")

            # 3. Comentarios institucionales (con número de regador dinámico)
            comentarios_generales = (
                f"Operativos\n"
                f"Cargador M-627 {est_carg_627}. \n"
                f"Picoton M-855 {est_pic_855}. \n"
                f"Camión Aljibe ({num_regador}) {est_aljibe}.\n\n"
                f"Stand by:\n"
                f"Picoton 854 {est_pic_854}.\n"
                f"Cargador 637 {est_carg_637}.\n\n"
                f"Tiempo en teleoperación: {tiempo_teleop} min por perdida de señal."
            )
            ws.cell(row=8, column=6, value=comentarios_generales)

            # Inyección de observaciones específicas por cada control de actividad (Filas 10 a 14)
            if st.session_state.get("obs_ctrl_10"): ws['F10'] = st.session_state.get("obs_ctrl_10")
            if st.session_state.get("obs_ctrl_11"): ws['F11'] = st.session_state.get("obs_ctrl_11")
            if st.session_state.get("obs_ctrl_12"): ws['F12'] = st.session_state.get("obs_ctrl_12")
            if st.session_state.get("obs_ctrl_13"): ws['F13'] = st.session_state.get("obs_ctrl_13")
            ws['F14'] = st.session_state.get("obs_ctrl_14", f"Camión Aljibe {num_regador} {est_aljibe.lower()}")

            # 4. Termografías (Centradas)
            p_pic_rows = [18, 19, 20]
            c_car_rows = [23, 24, 25]
            
            idx_p = 0
            idx_c = 0
            for term in st.session_state.termografias_registradas:
                if term["Equipo"] == "Picotón" and idx_p < len(p_pic_rows):
                    r = p_pic_rows[idx_p]
                    ws.cell(row=r, column=2, value=term["CPS"]).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=r, column=3, value=term["Flexibles"]).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=r, column=4, value=term["Cuña"]).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=r, column=5, value=term["Tornamesa"]).alignment = Alignment(horizontal="center", vertical="center")
                    idx_p += 1
                elif term["Equipo"] == "Cargador" and idx_c < len(c_car_rows):
                    r = c_car_rows[idx_c]
                    ws.cell(row=r, column=2, value=term["CPS"]).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=r, column=3, value=term["Batea"]).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=r, column=4, value=term["Parte Inferior"]).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(row=r, column=5, value=term["Flexibles"]).alignment = Alignment(horizontal="center", vertical="center")
                    idx_c += 1

            # --- LIMPIAR ESTRICTAMENTE TODAS LAS CELDAS DE CABECERA DE FOTOS Y PALABRAS CPS (Filas 63 a 105) ---
            for col_l in ['B', 'C', 'D', 'E', 'F']:
                for row_l in [63, 64, 65, 66, 79, 80, 96, 97]:
                    cell_target = ws[f'{col_l}{row_l}']
                    cell_target.value = None

            # 5. Ciclos y Fotografías garantizando ambas fotos (Antes y Después) y comentarios de desviaciones
            ciclo_start_rows = [28, 40, 52]
            image_target_rows = [67, 81, 98]

            font_titulo = Font(name="Aptos Narrow", size=16, bold=True, color="000000")
            borde_negro_fino = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )

            for idx, ciclo in enumerate(st.session_state.ciclos_registrados):
                if idx >= len(ciclo_start_rows):
                    break
                
                r = ciclo_start_rows[idx]
                img_row = image_target_rows[idx]
                
                ws.cell(row=r, column=3, value=ciclo["CPS"])
                ws.cell(row=r+1, column=3, value=ciclo["Ciclo"])
                
                def parse_time(t_val):
                    if isinstance(t_val, str):
                        return datetime.strptime(t_val, "%H:%M:%S").time()
                    return t_val

                t_maz = parse_time(ciclo["T_Mazamorra"])
                t_blo = parse_time(ciclo["T_Bloqueo"])
                t_i_pic = parse_time(ciclo["T_Ing_Pic"])
                t_r_pic = parse_time(ciclo["T_Ret_Pic"])
                t_i_car = parse_time(ciclo["T_Ing_Carg"])
                t_r_car = parse_time(ciclo["T_Ret_Carg"])
                t_des = parse_time(ciclo["T_Desbloqueo"])

                hitos = [
                    (datetime.combine(fecha_turno_oficial, t_maz), r+3),
                    (datetime.combine(fecha_turno_oficial, t_blo), r+4),
                    (datetime.combine(fecha_turno_oficial, t_i_pic), r+5),
                    (datetime.combine(fecha_turno_oficial, t_r_pic), r+6),
                    (datetime.combine(fecha_turno_oficial, t_i_car), r+7),
                    (datetime.combine(fecha_turno_oficial, t_r_car), r+8),
                    (datetime.combine(fecha_turno_oficial, t_des), r+9)
                ]
                
                for h_dt, h_row in hitos:
                    cell_h = ws.cell(row=h_row, column=3, value=h_dt)
                    cell_h.number_format = 'DD-MM-YYYY HH:MM:SS'
                
                # Baldadas y Toneladas / Temp
                ret_carg_row = r + 8
                ws.cell(row=ret_carg_row, column=5, value=ciclo["Baldadas"])
                ws.cell(row=ret_carg_row, column=6, value=ciclo["Toneladas"])

                # Inyectar comentario de desviación del ciclo en la fila de Término retiro mazamorra (Col F)
                if ciclo.get("Comentarios"):
                    ws.cell(row=r+3, column=6, value=ciclo["Comentarios"])

                titulo_row = img_row - 1

                # Inserción de Foto de ENTRADA (Antes)
                if ciclo.get("Foto_Entrada") is not None:
                    try:
                        buf_ent = optimizar_imagen(ciclo["Foto_Entrada"])
                        
                        cell_t_ent = ws.cell(row=titulo_row, column=2, value=f"📸 FOTO ANTES DE LIMPIEZA - {ciclo['CPS']} ({ciclo['Ciclo']})")
                        cell_t_ent.font = font_titulo
                        cell_t_ent.alignment = Alignment(horizontal="left", vertical="center")
                        cell_t_ent.border = borde_negro_fino
                        
                        img_ent = XLImage(buf_ent)
                        img_ent.width = 180
                        img_ent.height = 140
                        ws.add_image(img_ent, f"B{img_row}")
                    except Exception as img_err:
                        st.warning(f"Error foto entrada {ciclo['CPS']}: {img_err}")

                # Inserción de Foto de SALIDA (Después)
                if ciclo.get("Foto_Salida") is not None:
                    try:
                        buf_sal = optimizar_imagen(ciclo["Foto_Salida"])
                        
                        cell_t_sal = ws.cell(row=titulo_row, column=5, value=f"📸 FOTO DESPUES DE LIMPIEZA - {ciclo['CPS']} ({ciclo['Ciclo']})")
                        cell_t_sal.font = font_titulo
                        cell_t_sal.alignment = Alignment(horizontal="left", vertical="center")
                        cell_t_sal.border = borde_negro_fino
                        
                        img_sal = XLImage(buf_sal)
                        img_sal.width = 180
                        img_sal.height = 140
                        ws.add_image(img_sal, f"E{img_row}")
                    except Exception as img_err:
                        st.warning(f"Error foto salida {ciclo['CPS']}: {img_err}")

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            exito, mensaje = enviar_correo_gmail(correo_supervisor, output, nombre_archivo_salida)
            if exito:
                st.success(mensaje)
            else:
                st.error(mensaje)
                st.download_button(
                    label="📥 Descarga Alternativa de Planilla Excel",
                    data=output,
                    file_name=nombre_archivo_salida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error al generar o enviar el archivo: {e}")

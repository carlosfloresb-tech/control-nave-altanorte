import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import io
import os

st.set_page_config(page_title="Control Operacional - Nave Altonorte", layout="wide")

st.title("🏭 Sistema de Control Operacional - Nave Altonorte")
st.markdown("Plataforma web oficial para termografía, ciclos de CPS, registro fotográfico con trazabilidad y planilla oficial de Altonorte.")

# --- BARRA LATERAL: CONFIGURACIÓN GENERAL DEL TURNO ---
st.sidebar.header("📋 Identificación del Turno")
fecha_turno = st.sidebar.date_input("Fecha", datetime.today())
tipo_turno = st.sidebar.selectbox("Tipo de Turno", ["Turno Día (TA)", "Turno Noche (TB)"])

st.sidebar.subheader("Supervisores y Operadores")
correo_supervisor = st.sidebar.text_input("📧 Correo Supervisor de Nave", "carlos.flores@altonorte.cl")
sup_sop = st.sidebar.text_input("Supervisor SOP", "Rene Philipps")
sup_crm = st.sidebar.text_input("Supervisor CRM", "Jovelino Burgos")
sup_caemin = st.sidebar.text_input("Supervisor Caemin", "Ruben Infanta")
sup_nave = st.sidebar.text_input("Supervisor Nave", "Jorge Galindo")
op_picoton = st.sidebar.text_input("Operador Picotón", "Juan Pablo Figueroa")
op_cargador = st.sidebar.text_input("Operador Cargador", "Fernando Tapia")

# --- INICIALIZAR MEMORIA TEMPORAL DEL TURNO ---
if "ciclos_registrados" not in st.session_state:
    st.session_state.ciclos_registrados = []

if "termografias_registradas" not in st.session_state:
    st.session_state.termografias_registradas = []

# --- PESTAÑAS PRINCIPALES ---
tab1, tab2, tab3, tab4 = st.tabs([
    "1. Checklist & Equipos", 
    "2. Termografía (°C)", 
    "3. Ciclos de CPS & Fotos", 
    "4. Consolidado y Excel Oficial"
])

with tab1:
    st.subheader("Control Previo de Inicio de Actividades y Equipos")
    
    st.markdown("### Estado Operacional de Equipos (CAEMIN)")
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        est_carg_627 = st.selectbox("Cargador M-627", ["Revisado por CAEMIN", "Pendiente de revisión CAEMIN"])
        est_carg_637 = st.selectbox("Cargador M-637", ["Pendiente de revisión CAEMIN", "Revisado por CAEMIN"])
        est_aljibe = st.selectbox("Camión Aljibe M-8380", ["Operativo", "No disponible"])
    with col_e2:
        est_pic_855 = st.selectbox("Picotón M-855", ["Revisado por CAEMIN", "Pendiente de revisión CAEMIN"])
        est_pic_854 = st.selectbox("Picotón M-854", ["Pendiente de revisión CAEMIN", "Revisado por CAEMIN"])
    
    tiempo_teleop = st.number_input("⏱️ Tiempo en teleoperación (minutos perdidos por señal)", min_value=0, value=0)

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
        st.success("¡Termografía agregada correctamente!")
        
    if len(st.session_state.termografias_registradas) > 0:
        st.markdown("### Historial de Termografías en el Turno")
        st.dataframe(pd.DataFrame(st.session_state.termografias_registradas), use_container_width=True)

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
    with col_val2: toneladas_ext = st.number_input("Toneladas estimadas", value=6.5)
        
    obs_desviacion = st.text_input("Detalle de desviación / Comentarios")

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
            "Foto_Entrada": foto_entrada.getvalue() if foto_entrada else None,
            "Foto_Salida": foto_salida.getvalue() if foto_salida else None
        })
        st.success(f"¡Ciclo {nombre_ciclo} para {cps_seleccionado} guardado con éxito!")

    if len(st.session_state.ciclos_registrados) > 0:
        st.markdown("### Historial de Ciclos Registrados en el Turno")
        st.dataframe(pd.DataFrame([{k: v for k, v in c.items() if not k.startswith("Foto")} for c in st.session_state.ciclos_registrados]), use_container_width=True)

with tab4:
    st.subheader("📋 Consolidado Total y Generación de Planilla Oficial")
    st.info(f"Correo Supervisor: **{correo_supervisor}** | Turno: **{tipo_turno}** | Fecha: **{fecha_turno}**")
    
    st.markdown("---")
    if st.button("🔄 Generar Planilla Excel Oficial de Altonorte"):
        str_fecha = fecha_turno.strftime("%d-%m")
        sufijo_turno = "TA" if "Día" in tipo_turno else "TB"
        nombre_nueva_hoja = f"{str_fecha} {sufijo_turno}"

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
            ws['F4'] = f"Operador Picotón: {op_picoton}"
            
            ws['C5'] = sup_caemin
            ws['E5'] = f"{sup_nave} | Correo: {correo_supervisor}"
            ws['F5'] = f"Operador Cargador: {op_cargador}"

            # 2. Asignar ticks (✓) en las celdas de estatus (Columna E, filas 8 a 14)
            for r_chk in range(8, 15):
                ws.cell(row=r_chk, column=5, value="✓")

            # 3. Generar bloque de comentarios exacto de Altonorte
            comentarios_generales = (
                f"Operativos\n"
                f"Cargador M-627 {est_carg_627}. \n"
                f"Picoton M-855 {est_pic_855}. \n"
                f"Aljibe M-8380 {est_aljibe}.\n\n"
                f"Stand by:\n"
                f"Picoton 854 {est_pic_854}.\n"
                f"Cargador 637 {est_carg_637}.\n\n"
                f"Tiempo en teleoperación: {tiempo_teleop} min por perdida de señal."
            )
            ws.cell(row=8, column=6, value=comentarios_generales)

            # 4. Termografías
            p_pic_rows = [18, 19, 20]
            c_car_rows = [23, 24, 25]
            
            idx_p = 0
            idx_c = 0
            for term in st.session_state.termografias_registradas:
                if term["Equipo"] == "Picotón" and idx_p < len(p_pic_rows):
                    r = p_pic_rows[idx_p]
                    ws.cell(row=r, column=2, value=term["CPS"])
                    ws.cell(row=r, column=3, value=term["Flexibles"])
                    ws.cell(row=r, column=4, value=term["Cuña"])
                    ws.cell(row=r, column=5, value=term["Tornamesa"])
                    idx_p += 1
                elif term["Equipo"] == "Cargador" and idx_c < len(c_car_rows):
                    r = c_car_rows[idx_c]
                    ws.cell(row=r, column=2, value=term["CPS"])
                    ws.cell(row=r, column=3, value=term["Batea"])
                    ws.cell(row=r, column=4, value=term["Parte Inferior"])
                    ws.cell(row=r, column=5, value=term["Flexibles"])
                    idx_c += 1

            # 5. Ciclos, Toneladas y Fotografías con trazabilidad de CPS
            start_row = 28
            for idx, ciclo in enumerate(st.session_state.ciclos_registrados):
                r = start_row + (idx * 15) # Espacio optimizado para incluir imágenes abajo
                
                ws.cell(row=r, column=3, value=ciclo["CPS"])
                ws.cell(row=r+1, column=3, value=ciclo["Ciclo"])
                
                hitos = [
                    (ciclo["T_Mazamorra"], r+3),
                    (ciclo["T_Bloqueo"], r+4),
                    (ciclo["T_Ing_Pic"], r+5),
                    (ciclo["T_Ret_Pic"], r+6),
                    (ciclo["T_Ing_Carg"], r+7),
                    (ciclo["T_Ret_Carg"], r+8),
                    (ciclo["T_Desbloqueo"], r+9)
                ]
                
                for h_time, h_row in hitos:
                    ws.cell(row=h_row, column=3, value=h_time)
                
                ret_carg_row = r + 8
                ws.cell(row=ret_carg_row, column=5, value=ciclo["Baldadas"])
                ws.cell(row=ret_carg_row, column=6, value=ciclo["Toneladas"])

                tot_row = r+10

                # Inyectar Fotografías con etiqueta del CPS correspondiente
                if ciclo["Foto_Entrada"]:
                    ws.cell(row=tot_row+1, column=2, value=f"Foto Entrada - {ciclo['CPS']} ({ciclo['Ciclo']})")
                    img_ent = XLImage(io.BytesIO(ciclo["Foto_Entrada"]))
                    img_ent.width = 150
                    img_ent.height = 110
                    ws.add_image(img_ent, f"B{tot_row+2}")
                
                if ciclo["Foto_Salida"]:
                    ws.cell(row=tot_row+1, column=5, value=f"Foto Salida - {ciclo['CPS']} ({ciclo['Ciclo']})")
                    img_sal = XLImage(io.BytesIO(ciclo["Foto_Salida"]))
                    img_sal.width = 150
                    img_sal.height = 110
                    ws.add_image(img_sal, f"E{tot_row+2}")

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("¡Planilla oficial de Altonorte generada con éxito!")
            st.download_button(
                label="📥 Descargar Planilla Excel Oficial Actualizada",
                data=output,
                file_name=f"Control_Nave_{str_fecha}_{sufijo_turno}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error al generar el archivo: {e}")

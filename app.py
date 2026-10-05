import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import openpyxl
import io
import os

st.set_page_config(page_title="Control Operacional - Nave Altonorte", layout="wide")

st.title("🏭 Sistema de Control Operacional - Nave Altonorte")
st.markdown("Plataforma en línea para el registro de turnos, termografía, ciclos de equipos y generación de planilla oficial de Altonorte.")

# --- BARRA LATERAL: CONFIGURACIÓN GENERAL DEL TURNO ---
st.sidebar.header("📋 Identificación del Turno")
fecha_turno = st.sidebar.date_input("Fecha", datetime.today())
tipo_turno = st.sidebar.selectbox("Tipo de Turno", ["Turno Día (TA)", "Turno Noche (TB)"])

st.sidebar.subheader("Supervisores y Operadores")
correo_supervisor = st.sidebar.text_input("📧 Correo Supervisor de Nave", "supervisor.nave@altonorte.cl")
sup_sop = st.sidebar.text_input("Supervisor SOP", "Rene Philipps")
sup_crm = st.sidebar.text_input("Supervisor CRM", "Jovelino Burgos")
sup_caemin = st.sidebar.text_input("Supervisor Caemin", "Ruben Infanta")
sup_nave = st.sidebar.text_input("Supervisor Nave", "Jorge Galindo")
op_picoton = st.sidebar.text_input("Operador Picotón", "Juan Pablo Figueroa")
op_cargador = st.sidebar.text_input("Operador Cargador", "Fernando Tapia")

# --- INICIALIZAR MEMORIA TEMPORAL PARA EL TURNO ---
if "ciclos_registrados" not in st.session_state:
    st.session_state.ciclos_registrados = []

if "termografias_registradas" not in st.session_state:
    st.session_state.termografias_registradas = []

# --- PESTAÑAS PRINCIPALES ---
tab1, tab2, tab3, tab4 = st.tabs([
    "1. Checklist Inicio", 
    "2. Registro de Termografía (°C)", 
    "3. Registro de Ciclos CPS (Múltiples)", 
    "4. Resumen y Exportar Excel"
])

with tab1:
    st.subheader("Control Previo de Inicio de Actividades")
    chk_cargador = st.checkbox("Check list Operacional Cargador y Picotón (Op. Maq.)", value=True)
    chk_caemin = st.checkbox("Asegurar condición de equipos por Caemin", value=True)
    
    st.subheader("Controles de Retiro y Traslado de Material")
    c1 = st.checkbox("Confirmar tiempo de al menos 10 min desde término de picado del foso", value=True)
    c2 = st.checkbox("Verificar visualmente que material en foso esté sólido (no líquido)", value=True)
    c3 = st.checkbox("Verificar ausencia de llamas, humo o indicios de ignición", value=True)
    c4 = st.checkbox("Asegurar que disposición en explanada favorezca enfriamiento", value=True)
    c5 = st.checkbox("Verificar que camión aljibe esté disponible", value=True)
    
    comentarios_inicio = st.text_area("Comentarios y Estado de Equipos", "Operativos: Cargador y Picotón revisados. Camión aljibe operativo.")

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
    st.subheader("🔄 Registro de Ciclos CPS (Múltiples Ingresos)")
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        cps_seleccionado = st.selectbox("Seleccione CPS", ["CPS-1", "CPS-2", "CPS-3", "CPS-4"], key="ciclo_cps")
    with col_sel2:
        nombre_ciclo = st.text_input("Identificador de Ciclo (ej. J-5, J-6)", "J-5")
    
    st.markdown("---")
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

    if st.button("➕ Guardar Ciclo en el Turno"):
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
            "Comentarios": obs_desviacion
        })
        st.success(f"¡Ciclo {nombre_ciclo} guardado!")

    if len(st.session_state.ciclos_registrados) > 0:
        st.markdown("### Historial de Ciclos en el Turno")
        st.dataframe(pd.DataFrame(st.session_state.ciclos_registrados), use_container_width=True)

with tab4:
    st.subheader("📋 Consolidado Total del Turno y Generación de Excel")
    st.info(f"Correo Supervisor: **{correo_supervisor}** | Turno: **{tipo_turno}** | Fecha: **{fecha_turno}**")
    
    st.markdown("---")
    if st.button("🔄 Generar Planilla Excel Oficial de Altonorte"):
        str_fecha = fecha_turno.strftime("%d-%m")
        sufijo_turno = "TA" if "Día" in tipo_turno else "TB"
        nombre_nueva_hoja = f"{str_fecha} {sufijo_turno}"

        base_excel = 'Control Nave 02-10-26 TB.xlsx'
        
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
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = nombre_nueva_hoja

            # 1. Actualizar Cabecera con los datos ingresados
            ws['B3'] = f"Fecha: {str_fecha} {sufijo_turno}"
            ws['C4'] = sup_sop
            ws['E4'] = sup_crm
            ws['F4'] = f"Operador Picotón: {op_picoton} | Correo: {correo_supervisor}"
            
            ws['C5'] = sup_caemin
            ws['E5'] = sup_nave
            ws['F5'] = f"Operador Cargador: {op_cargador}"

            # 2. Limpiar bloques de ciclos anteriores en la plantilla clonada para evitar datos residuales
            for r in range(28, 95):
                for c in range(2, 8):
                    ws.cell(row=r, column=c, value=None)

            # 3. Inyectar dinámicamente los ciclos registrados en las filas correspondientes de la hoja
            row_pointer = 28
            for ciclo in st.session_state.ciclos_registrados:
                ws.cell(row=row_pointer, column=2, value="CPS")
                ws.cell(row=row_pointer, column=3, value=ciclo["CPS"])
                row_pointer += 1
                
                ws.cell(row=row_pointer, column=2, value="Ciclo")
                ws.cell(row=row_pointer, column=3, value=ciclo["Ciclo"])
                row_pointer += 1
                
                ws.cell(row=row_pointer, column=2, value="CPS")
                ws.cell(row=row_pointer, column=3, value="Fecha")
                ws.cell(row=row_pointer, column=4, value="Minutos")
                ws.cell(row=row_pointer, column=5, value="Baldadas")
                ws.cell(row=row_pointer, column=6, value="Comentarios desviaciones")
                row_pointer += 1
                
                hitos = [
                    ("Termino retiro mazamorra desde CPS", ciclo["T_Mazamorra"]),
                    ("Bloqueo de tapas CPS", ciclo["T_Bloqueo"]),
                    ("Ingreso Picoton foso", ciclo["T_Ing_Pic"]),
                    ("Retiro picoton foso", ciclo["T_Ret_Pic"]),
                    ("Ingreso Cargador Foso", ciclo["T_Ing_Carg"]),
                    ("Retiro de Cargador en foso", ciclo["T_Ret_Carg"]),
                    ("Desbloqueo de tapas CPS", ciclo["T_Desbloqueo"])
                ]
                
                start_hito_row = row_pointer
                for h_text, h_time in hitos:
                    ws.cell(row=row_pointer, column=2, value=h_text)
                    ws.cell(row=row_pointer, column=3, value=h_time)
                    # Fórmula de minutos idéntica a la original de Altonorte
                    ws.cell(row=row_pointer, column=4, value=f"=+(C{row_pointer}-C{start_hito_row})*1440")
                    row_pointer += 1
                
                # Fila Total con fórmula SUM
                ws.cell(row=row_pointer, column=2, value="Total")
                ws.cell(row=row_pointer, column=4, value=f"=SUM(D{start_hito_row}:D{row_pointer-1})")
                ws.cell(row=row_pointer, column=5, value=ciclo["Baldadas"])
                ws.cell(row=row_pointer, column=6, value=f"{ciclo['Toneladas']} Ton. // {ciclo['Comentarios']}")
                row_pointer += 3 # Espacio entre ciclos

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("¡Planilla oficial generada y actualizada correctamente con todos los registros!")
            st.download_button(
                label="📥 Descargar Archivo Excel Maestro Actualizado",
                data=output,
                file_name=f"Control_Nave_{str_fecha}_{sufijo_turno}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error al generar el archivo: {e}")

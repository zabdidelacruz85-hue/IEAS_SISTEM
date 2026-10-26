import os
import pandas as pd
from datetime import datetime
import streamlit as st
import io

st.set_page_config(
    page_title="IEAS - Sistema Operativo Central",
    page_icon="⚡",
    layout="wide"
)

DB_PATH = "IEAS_Control_Maestro.xlsx"

def inicializar_bd():
    if not os.path.exists(DB_PATH):
        df_inicial = pd.DataFrame(columns=[
            "Folio_Contacto", "Fecha", "Cliente", "Telefono", "Ubicacion", 
            "Servicio", "Estatus", "N_Proyecto", "CDT_MXN", "Precio_Final_MXN", "Ingeniero_Responsable", "Notas"
        ])
        datos = [
            ["IEAS-26-002", "2026-09-05", "Clínica Dental Dr. Víctor Hugo", "228-234-5678", "Xalapa, Ver.", "Solar Fotovoltaico / MT", "PA", "PROYECTO 002/26", 93840.0, 138000.0, "101 - Ing. Zabdi (Director)", "12 paneles ReneSola 630W - Precio fijo acordado"]
        ]
        df_inicial = pd.DataFrame(datos, columns=df_inicial.columns)
        df_historial = pd.DataFrame(columns=["Fecha_Hora", "Usuario_ID", "Accion", "Detalles"])
        
        with pd.ExcelWriter(DB_PATH, engine='openpyxl') as writer:
            df_inicial.to_excel(writer, sheet_name="Control_Prospectos", index=False)
            df_historial.to_excel(writer, sheet_name="Historial_Auditoria", index=False)

inicializar_bd()

st.markdown("<h1 style='color: #78A68B;'>⚡ IEAS - Sistema Operativo y Trazabilidad de Proyectos</h1>", unsafe_allow_html=True)
st.markdown("<p style='color: #A0A5A8;'>Ingeniería Eléctrica Aplicada y Sustentable | Región Centro de Veracruz</p>", unsafe_allow_html=True)

# Padrón Oficial de Operadores
PADRON_OPERADORES = [
    "101 - Ing. Zabdi (Director)",
    "102 - Alonso",
    "103 - Aldo",
    "104 - Vicente",
    "105 - Said",
    "106 - Técnico por Asignar",
    "107 - Técnico por Asignar",
    "108 - Técnico por Asignar",
    "109 - Técnico por Asignar"
]

st.sidebar.header("🔑 Control de Acceso Operativo")
usuario_activo = st.sidebar.selectbox("Selecciona tu Identidad / ID:", PADRON_OPERADORES)

menu = st.sidebar.selectbox("Menú de Navegación:", [
    "📊 Dashboard y Proyectos", 
    "➕ Alta y Calculador Solar Dinámico", 
    "📥 Descarga de Reportes (Impresión)",
    "📜 Historial de Auditoría"
])

df = pd.read_excel(DB_PATH, sheet_name="Control_Prospectos")
df_hist = pd.read_excel(DB_PATH, sheet_name="Historial_Auditoria")

def registrar_auditoria(usuario, accion, detalles):
    global df_hist
    nuevo_cambio = {
        "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Usuario_ID": usuario,
        "Accion": accion,
        "Detalles": detalles
    }
    df_hist = pd.concat([df_hist, pd.DataFrame([nuevo_cambio])], ignore_index=True)
    
    with pd.ExcelWriter(DB_PATH, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, sheet_name="Control_Prospectos", index=False)
        df_hist.to_excel(writer, sheet_name="Historial_Auditoria", index=False)

if menu == "📊 Dashboard y Proyectos":
    st.subheader("📋 Base de Datos de Proyectos Activos")
    st.dataframe(df, use_container_width=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="Valor Cartera Total (Con IVA)", value=f"${df['Precio_Final_MXN'].sum():,.2f} MXN")
    with col2:
        st.metric(label="Total Proyectos Registrados", value=f"{len(df)}")

elif menu == "➕ Alta y Calculador Solar Dinámico":
    st.subheader("☀️ Calculador de Ingeniería Solar y Apertura de Proyecto")
    st.info(f"Operador activo firmando el registro: **{usuario_activo}**")
    
    with st.form("form_solar"):
        st.markdown("### 1. Parámetros de Diseño y Consumo (Modificables)")
        c1, c2, c3 = st.columns(3)
        with c1:
            consumo_kwh_dia = st.number_input("Consumo Promedio (kWh/día)", value=25.0, step=1.0)
            hsp = st.number_input("Horas Sol Pico (HSP Región)", value=4.5, step=0.1)
        with c2:
            pr = st.number_input("Factor de Rendimiento (PR)", value=0.80, step=0.01)
            potencia_panel_w = st.number_input("Potencia Panel (Watts)", value=630.0, step=10.0)
        with c3:
            margen_bruto = st.number_input("Margen Bruto Objetivo", value=0.32, step=0.01)
            iva = st.number_input("Factor IVA", value=1.16, step=0.01)

        # Cálculo dinámico en tiempo real dentro del formulario
        energia_mes_requerida = consumo_kwh_dia * 30
        kwp_necesarios = energia_mes_requerida / (hsp * pr * 30) if (hsp * pr * 30) > 0 else 0
        paneles_necesarios = int((kwp_necesarios * 1000) // potencia_panel_w) + (1 if (kwp_necesarios * 1000) % potencia_panel_w > 0 else 0)
        
        st.markdown(f"👉 **Resultado de Ingeniería Sugerido:** Se requieren **{paneles_necesarios} paneles de {potencia_panel_w}W** (~{kwp_necesarios:.2f} kWp).")

        st.markdown("### 2. Datos del Cliente y Costeo")
        col1, col2 = st.columns(2)
        with col1:
            cliente = st.text_input("Nombre del Cliente / Titular")
            telefono = st.text_input("Teléfono de Contacto")
            ubicacion = st.text_input("Ubicación Exacta (Ej. Xalapa, Ver.)")
        with col2:
            servicio = st.text_input("Descripción del Sistema", value=f"Sistema Fotovoltaico {paneles_necesarios} Módulos {potencia_panel_w}W")
            cdt = st.number_input("Costo Directo Total (CDT en MXN)", min_value=0.0, step=1000.0)
            notas = st.text_area("Notas Técnicas / Alertas en Sitio")
            
        submit = st.form_submit_button("⚡ Generar Proyecto y Guardar en Sistema")
        
        if submit:
            if not cliente or cdt <= 0:
                st.error("Por favor ingresa un cliente válido y un Costo Directo Total (CDT) mayor a cero.")
            else:
                # Aplicación de la fórmula institucional IEAS con los parámetros configurados
                subtotal_venta = cdt / (1 - margen_bruto)
                precio_final = subtotal_venta * iva
                
                n_id = len(df) + 1
                folio_contacto = f"IEAS-26-{n_id:03d}"
                n_proyecto = f"PROYECTO {n_id:03d}/26"
                
                nuevo_row = {
                    "Folio_Contacto": folio_contacto,
                    "Fecha": datetime.now().strftime("%Y-%m-%d"),
                    "Cliente": cliente,
                    "Telefono": telefono,
                    "Ubicacion": ubicacion,
                    "Servicio": servicio,
                    "Estatus": "PA",
                    "N_Proyecto": n_proyecto,
                    "CDT_MXN": cdt,
                    "Precio_Final_MXN": round(precio_final, 2),
                    "Ingeniero_Responsable": usuario_activo,
                    "Notas": f"{paneles_necesarios} paneles de {potencia_panel_w}W | {notas}"
                }
                
                df = pd.concat([df, pd.DataFrame([nuevo_row])], ignore_index=True)
                
                with pd.ExcelWriter(DB_PATH, engine='openpyxl', mode='w') as writer:
                    df.to_excel(writer, sheet_name="Control_Prospectos", index=False)
                    df_hist.to_excel(writer, sheet_name="Historial_Auditoria", index=False)
                
                registrar_auditoria(usuario_activo, "ALTA DE PROYECTO SOLAR", f"Creado {n_proyecto} para {cliente} por {usuario_activo}")
                st.success(f"¡{n_proyecto} registrado con éxito y firmado por **{usuario_activo}**!")
                st.balloons()

elif menu == "📥 Descarga de Reportes (Impresión)":
    st.subheader("🖨️ Generación de Reportes Listos para Impresión")
    st.markdown("Descarga el archivo Excel con el padrón limpio para imprimir y entregar en obra.")
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Control_Prospectos", index=False)
        df_hist.to_excel(writer, sheet_name="Historial_Auditoria", index=False)
    excel_data = output.getvalue()
    
    st.download_button(
        label="📥 Descargar Base de Datos Depurada (Excel)",
        data=excel_data,
        file_name=f"IEAS_Reporte_Operativo_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif menu == "📜 Historial de Auditoría":
    st.subheader("🔍 Registro Histórico de Modificaciones y Trazabilidad")
    st.dataframe(df_hist, use_container_width=True)

import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os
import random
import math
import io

# Intentar importar la librería OPC UA
try:
    from opcua import Client
    OPCUA_AVAILABLE = True
except ImportError:
    OPCUA_AVAILABLE = False

# Intentar importar openpyxl para escribir archivos Excel
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# 1. Configuración de página
st.set_page_config(page_title="OPC UA SCADA", layout="wide")

# ==============================================================================
# 2. ESTILOS (PALETA PROFESIONAL DARK SCADA + FIX DESPLEGABLES)
# ==============================================================================
st.markdown("""
<style>
    /* Ocultar menús por defecto de Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Fondo principal y tipografía base */
    html, body, .stApp { 
        background-color: #0F172A !important; 
        color: #E2E8F0 !important;
        font-family: 'Inter', system-ui, -apple-system, sans-serif !important;
    }

    /* Encabezados y Títulos */
    h1, h2, h3, h4, h5, h6, label {
        color: #F8FAFC !important;
        font-weight: 700 !important;
        text-shadow: 0 1px 3px rgba(0,0,0,0.5) !important;
    }

    /* Subtítulos de sección en Ámbar Dorado suave */
    [data-testid="stMarkdownContainer"] h2, 
    [data-testid="stMarkdownContainer"] h3 {
        color: #F59E0B !important;
    }

    /* Textos generales y párrafos */
    [data-testid="stMarkdownContainer"] p, 
    [data-testid="stMarkdownContainer"] li {
        color: #CBD5E1 !important;
        font-size: 15px !important;
        line-height: 1.6 !important;
    }

    /* FIX CORRECCIÓN DE DESPLEGABLES / SELECTBOX EN MODO OSCURO */
    div[data-testid="stSelectbox"] > div > div {
        background-color: #1E293B !important;
        color: #F8FAFC !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
    }

    div[data-testid="stSelectbox"] div[role="button"] p,
    div[data-testid="stSelectbox"] span {
        color: #F8FAFC !important;
        font-weight: 600 !important;
    }

    div[data-testid="stSelectbox"] svg {
        fill: #38BDF8 !important;
    }

    div[data-baseweb="select"] > div,
    div[data-baseweb="popover"],
    div[data-baseweb="menu"],
    ul[role="listbox"],
    div[role="listbox"] {
        background-color: #1E293B !important;
        border: 1px solid #334155 !important;
    }

    li[role="option"],
    div[role="option"] {
        background-color: #1E293B !important;
        color: #E2E8F0 !important;
        font-weight: 500 !important;
    }

    li[role="option"]:hover,
    div[role="option"]:hover,
    li[aria-selected="true"],
    div[role="option"][aria-selected="true"] {
        background-color: #334155 !important;
        color: #38BDF8 !important;
    }

    [data-testid="stDataEditor"] input,
    [data-testid="stDataEditor"] select {
        background-color: #1E293B !important;
        color: #F8FAFC !important;
    }

    /* Tarjetas SCADA */
    .card-scada {
        background-color: #1E293B !important;
        border-radius: 12px !important;
        padding: 18px !important;
        margin-bottom: 15px !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3) !important;
    }

    /* Botones generales */
    div.stButton > button {
        background: linear-gradient(135deg, #0EA5E9, #0284C7) !important;
        color: #FFFFFF !important;
        border: 1px solid #38BDF8 !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 12px rgba(14, 165, 233, 0.3) !important;
        transition: all 0.2s ease-in-out !important;
    }

    div.stButton > button:hover {
        background: linear-gradient(135deg, #38BDF8, #0EA5E9) !important;
        box-shadow: 0 6px 18px rgba(56, 189, 248, 0.5) !important;
        transform: translateY(-2px) !important;
    }

    button[data-baseweb="tab"] p {
        color: #94A3B8 !important;
        font-weight: 600 !important;
        font-size: 15px !important;
    }

    button[aria-selected="true"][data-baseweb="tab"] p {
        color: #38BDF8 !important;
        font-weight: 800 !important;
    }

    div[data-testid="stImage"] {
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        height: 160px !important;
        margin-bottom: 10px !important;
    }

    div[data-testid="stImage"] > img {
        max-height: 100% !important;
        max-width: 100% !important;
        object-fit: contain !important;
        border-radius: 8px !important;
    }
</style>
""", unsafe_allow_html=True)

# Generación Interna de NodeID Adaptativa por Marca
def generar_node_id_interno(bloque_o_prog, variable, marca="Siemens (OPC UA)"):
    bloque_clean = str(bloque_o_prog).strip() if pd.notna(bloque_o_prog) and str(bloque_o_prog).strip() != "" else ""
    var_clean = str(variable).strip()

    if "Siemens" in marca:
        db_clean = bloque_clean.upper() if bloque_clean else "DB1"
        return f'ns=3;s="{db_clean}"."{var_clean}"'
    elif "Allen-Bradley" in marca:
        if bloque_clean:
            return f'ns=2;s={bloque_clean}.{var_clean}'
        return f'ns=2;s={var_clean}'
    elif "Schneider" in marca:
        if bloque_clean:
            return f'ns=2;s={bloque_clean}.{var_clean}'
        return f'ns=2;s=Application.GVL.{var_clean}'
    else:
        return f'ns=2;s={var_clean}'

def obtener_variables_iniciales(marca):
    if "Siemens" in marca:
        col_header = "Bloque (DB)"
        default_val = "DB1"
    elif "Allen-Bradley" in marca:
        col_header = "Programa / Scope (Opcional)"
        default_val = ""
    else:
        col_header = "Prefijo / GVL (Opcional)"
        default_val = "Application.GVL"

    return pd.DataFrame([
        {col_header: default_val, "Nombre de Variable": "Temperatura", "Tipo de Dato": "FLOAT", "Unidad": "°C"},
        {col_header: default_val, "Nombre de Variable": "Presion", "Tipo de Dato": "FLOAT", "Unidad": "PSI"},
        {col_header: default_val, "Nombre de Variable": "Conteo_Piezas", "Tipo de Dato": "INT", "Unidad": "pzs"},
        {col_header: default_val, "Nombre de Variable": "Bomba_Activa", "Tipo de Dato": "BOOL", "Unidad": "Estado"}
    ])

def generar_valor_simulado(var_nombre, tipo_var, estado_anterior):
    t = datetime.now().timestamp()
    
    if tipo_var == "BOOL":
        prev = bool(estado_anterior) if estado_anterior is not None else False
        if random.random() < 0.15:
            return not prev
        return prev

    elif tipo_var in ["INT", "DINT"]:
        prev = int(estado_anterior) if estado_anterior is not None and isinstance(estado_anterior, (int, float)) else 100
        inc = random.choice([0, 1, 1, 2, -1])
        nuevo_val = max(0, prev + inc)
        return nuevo_val

    else:  # FLOAT
        frecuencia = 0.1
        amplitud = 25.0
        offset = 50.0
        offset_fase = sum(ord(c) for c in var_nombre) % 10
        ruido = random.uniform(-1.2, 1.2)
        
        val_simulado = offset + amplitud * math.sin(frecuencia * t + offset_fase) + ruido
        return round(val_simulado, 2)

# Inicialización de Estados de Sesión
if 'idioma' not in st.session_state:
    st.session_state['idioma'] = "ES"
if 'pantalla' not in st.session_state:
    st.session_state['pantalla'] = "INICIO"
if 'estoy_conectado' not in st.session_state:
    st.session_state['estoy_conectado'] = False
if 'modo_simulacion' not in st.session_state:
    st.session_state['modo_simulacion'] = True
if 'marca_plc' not in st.session_state:
    st.session_state['marca_plc'] = "Siemens (OPC UA)"
if 'variables_config' not in st.session_state:
    st.session_state['variables_config'] = obtener_variables_iniciales(st.session_state['marca_plc'])
if 'mi_lista_de_datos' not in st.session_state:
    cols = ["Fecha_Hora"] + list(st.session_state['variables_config']['Nombre de Variable'])
    st.session_state['mi_lista_de_datos'] = pd.DataFrame(columns=cols)
if 'ultimos_valores_sim' not in st.session_state:
    st.session_state['ultimos_valores_sim'] = {}

# Encabezado Fijo
top_col1, top_col2, top_col3 = st.columns([6, 2.5, 1.5])
with top_col1:
    st.title("🌐 PROTOCOLO INDUSTRIAL OPC UA")
with top_col2:
    if st.button("📖 Guía de Configuración", key="btn_global_manual", use_container_width=True):
        st.session_state['pantalla'] = "GUIA"
        st.rerun()
with top_col3:
    st.session_state['idioma'] = st.radio(
        "🌐 Idioma", 
        ["ES", "EN"], 
        horizontal=True, 
        index=0 if st.session_state['idioma']=="ES" else 1,
        key="selector_idioma_global"
    )

st.markdown("---")

# PANTALLA 1: SELECCIÓN DE SERVIDOR
if st.session_state['pantalla'] == "INICIO":
    st.subheader("🎓 SELECCIÓN DE SERVIDOR OPC UA" if st.session_state['idioma'] == "ES" else "🎓 OPC UA SERVER SELECTION")
    st.markdown(
        "Estándar OPC UA. Seleccione la marca del controlador para continuar o revise la guía de configuración."
        if st.session_state['idioma'] == "ES" else
        "OPC UA standard. Select the PLC brand or open the setup guide."
    )

    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.image("Siemens 1.jpg", use_container_width=True) 
        if st.button("🚀 Conectar Siemens (OPC UA)" if st.session_state['idioma'] == "ES" else "🚀 Select Siemens (OPC UA)", key="btn_s7", use_container_width=True):
            st.session_state['marca_plc'] = "Siemens (OPC UA)"
            st.session_state['variables_config'] = obtener_variables_iniciales("Siemens (OPC UA)")
            st.session_state['pantalla'] = "CONFIGURACION"
            st.rerun()

    with col2:
        st.image("schneider 1.jpg", use_container_width=True)
        if st.button("🚀 Conectar Schneider (OPC UA)" if st.session_state['idioma'] == "ES" else "🚀 Select Schneider (OPC UA)", key="btn_sch", use_container_width=True):
            st.session_state['marca_plc'] = "Schneider (OPC UA)"
            st.session_state['variables_config'] = obtener_variables_iniciales("Schneider (OPC UA)")
            st.session_state['pantalla'] = "CONFIGURACION"
            st.rerun()

    with col3:
        st.image("Allen bradley 1.png", use_container_width=True)
        if st.button("🚀 Conectar Allen-Bradley (OPC UA)" if st.session_state['idioma'] == "ES" else "🚀 Select Allen-Bradley (OPC UA)", key="btn_ab", use_container_width=True):
            st.session_state['marca_plc'] = "Allen-Bradley (OPC UA)"
            st.session_state['variables_config'] = obtener_variables_iniciales("Allen-Bradley (OPC UA)")
            st.session_state['pantalla'] = "CONFIGURACION"
            st.rerun()

# GUÍA DE CONFIGURACIÓN
elif st.session_state['pantalla'] == "GUIA":
    st.subheader("📚 Guía de Configuración y Direccionamiento OPC UA" if st.session_state['idioma'] == "ES" else "📚 OPC UA Configuration & Tagging Guide")

    if st.button("⬅️ Volver a la Aplicación", key="btn_guide_back"):
        st.session_state['pantalla'] = "INICIO"
        st.rerun()

    tab_guide_siemens, tab_guide_ab, tab_guide_sch = st.tabs([
        "💻 Siemens", "🎛️ Allen-Bradley", "⚡ Schneider Electric"
    ])

    with tab_guide_siemens:
        st.markdown("### 💻 Siemens (S7-1200 FW ≥ 4.2 y S7-1500)\n* **Sintaxis NodeID:** `ns=3;s=\"<NOMBRE_DB>\".\"<NOMBRE_VARIABLE>\"`")

    with tab_guide_ab:
        st.markdown("### 🎛️ Allen-Bradley (ControlLogix / CompactLogix)\n* **Sintaxis NodeID:** `ns=2;s=<VARIABLE>` o `ns=2;s=<PROGRAMA>.<VARIABLE>`")

    with tab_guide_sch:
        st.markdown("### ⚡ Schneider Electric (Modicon M241 / M251 / M262 / M580)\n* **Sintaxis NodeID:** `ns=2;s=Application.GVL.<VARIABLE>`")

# PANTALLA 2: CONFIGURACIÓN
elif st.session_state['pantalla'] == "CONFIGURACION":
    marca_actual = st.session_state.get('marca_plc', 'Universal')
    st.subheader(f"⚙️ Configuración OPC UA: {marca_actual}" if st.session_state['idioma'] == "ES" else f"⚙️ OPC UA Configuration: {marca_actual}")
    
    col_ip, col_sim = st.columns([7, 3])
    with col_ip:
        plc_ip_input = st.text_input(
            "🌐 Dirección IP del PLC:" if st.session_state['idioma'] == "ES" else "🌐 PLC IP Address:",  
            value="192.168.0.2",
            key="cfg_plc_ip"
        )
    with col_sim:
        usar_sim = st.checkbox("🧪 Habilitar Simulación Automática" if st.session_state['idioma'] == "ES" else "🧪 Enable Auto Simulation", value=True)

    endpoint_calculado = f"opc.tcp://{plc_ip_input.strip()}:4840"

    st.markdown("#### 📋 Tags a Monitorear" if st.session_state['idioma'] == "ES" else "#### 📋 Monitored Tags")
    
    col_bloque_nombre = st.session_state['variables_config'].columns[0]
    
    df_editado = st.data_editor(
        st.session_state['variables_config'],
        column_config={
            col_bloque_nombre: st.column_config.TextColumn(col_bloque_nombre),
            "Nombre de Variable": st.column_config.TextColumn("Nombre de Variable", required=True),
            "Tipo de Dato": st.column_config.SelectboxColumn("Tipo de Dato", options=["FLOAT", "BOOL", "INT", "DINT"], required=True),
            "Unidad": st.column_config.TextColumn("Unidad")
        },
        num_rows="dynamic",
        use_container_width=True,
        key="cfg_editor_opc"
    )

    c_btn1, c_btn2 = st.columns(2)
    with c_btn1:
        if st.button("🟢 Iniciar Estación SCADA OPC UA" if st.session_state['idioma'] == "ES" else "🟢 Start OPC UA SCADA Station", key="btn_start", use_container_width=True):
            st.session_state['opc_url'] = endpoint_calculado
            st.session_state['modo_simulacion'] = usar_sim
            st.session_state['variables_config'] = df_editado.copy()
            
            cols = ["Fecha_Hora"] + list(st.session_state['variables_config']['Nombre de Variable'])
            st.session_state['mi_lista_de_datos'] = pd.DataFrame(columns=cols)
            
            st.session_state['estoy_conectado'] = True
            st.session_state['pantalla'] = "SCADA_MAIN"
            st.rerun()
            
    with c_btn2:
        if st.button("🏠 Inicio / Cambiar Servidor" if st.session_state['idioma'] == "ES" else "🏠 Switch Server", key="btn_back", use_container_width=True):
            st.session_state['pantalla'] = "INICIO"
            st.rerun()

# PANTALLA 3: ENTORNO SCADA
elif st.session_state['pantalla'] == "SCADA_MAIN":
    marca_plc = st.session_state.get('marca_plc', 'PLC_OPC_UA')
    opc_url = st.session_state.get('opc_url', 'opc.tcp://127.0.0.1:4840')
    modo_sim = st.session_state.get('modo_simulacion', True)
    df_vars = st.session_state['variables_config']
    
    col_nav, col_btn = st.columns([8, 2])
    with col_btn:
        if st.button("🏠 Cambiar Servidor" if st.session_state['idioma'] == "ES" else "🏠 Switch Server", key="btn_exit", use_container_width=True):
            st.session_state['estoy_conectado'] = False
            st.session_state['pantalla'] = "INICIO"
            st.rerun()

    titulos_tabs = ["📊 Monitor & Diagnóstico", "📈 Tendencias", "💾 Histórico y Adquisición"] if st.session_state['idioma'] == "ES" else ["📊 Monitor & Diagnostics", "📈 Trends", "💾 History & Acquisition"]
    with col_nav:
        tab_mon, tab_tre, tab_his = st.tabs(titulos_tabs)

    @st.fragment(run_every=1.0)
    def render_scada_live():
        lista_diagnostico = []
        valores_actuales = {}
        
        txt_conn = "🟢 Conectado (OPC UA)" if st.session_state['idioma'] == "ES" else "🟢 Connected (OPC UA)"
        txt_sim = "🧪 Simulación Activa" if st.session_state['idioma'] == "ES" else "🧪 Simulated Data"
        txt_fail = "🔴 Fallo de Lectura" if st.session_state['idioma'] == "ES" else "🔴 Read Fail"

        client_opc = None
        if not modo_sim and OPCUA_AVAILABLE:
            try:
                client_opc = Client(opc_url)
                client_opc.timeout = 0.2
                client_opc.connect()
            except Exception:
                client_opc = None

        col_bloque_nombre = df_vars.columns[0]

        for _, row in df_vars.iterrows():
            bloque_val = row.get(col_bloque_nombre, '')
            var_nombre = str(row.get('Nombre de Variable', ''))
            tipo_var = str(row.get('Tipo de Dato', 'FLOAT')).upper()
            unidad = str(row['Unidad']) if pd.notna(row['Unidad']) else ""
            
            node_id_str = generar_node_id_interno(bloque_val, var_nombre, marca_plc)

            val = None
            estado = txt_fail

            if modo_sim:
                val_anterior = st.session_state['ultimos_valores_sim'].get(var_nombre, None)
                val = generar_valor_simulado(var_nombre, tipo_var, val_anterior)
                st.session_state['ultimos_valores_sim'][var_nombre] = val
                estado = txt_sim

            elif client_opc:
                try:
                    node = client_opc.get_node(node_id_str)
                    val = node.get_value()
                    estado = txt_conn
                except Exception:
                    val = None
                    estado = txt_fail

            valores_actuales[var_nombre] = val
            
            if tipo_var == "BOOL":
                val_str = "ON (1)" if bool(val) else "OFF (0)"
            elif tipo_var in ["INT", "DINT"]:
                val_str = f"{int(val)} {unidad}" if val is not None else "---"
            else:
                val_str = f"{float(val):.2f} {unidad}" if val is not None else "---"

            lista_diagnostico.append({
                col_bloque_nombre: bloque_val if bloque_val else "---",
                "Variable": var_nombre,
                "Tipo de Dato": tipo_var,
                "Valor Actual": val_str,
                "NodeID Generado": node_id_str,
                "Estado de Conexión": estado
            })

        if client_opc:
            try:
                client_opc.disconnect()
            except Exception:
                pass

        if valores_actuales:
            hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            registro_dict = {"Fecha_Hora": hora, **valores_actuales}
            row_df = pd.DataFrame([registro_dict])
            st.session_state['mi_lista_de_datos'] = pd.concat([st.session_state['mi_lista_de_datos'], row_df], ignore_index=True).tail(500)

        with tab_mon:
            st.markdown(f"### Estado de Variables OPC UA ({marca_plc})" if st.session_state['idioma'] == "ES" else f"### OPC UA Variable Status ({marca_plc})")
            
            df_diag = pd.DataFrame(lista_diagnostico)
            st.dataframe(df_diag, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            st.markdown("#### Tarjetas de Medición Identificadas por Tipo" if st.session_state['idioma'] == "ES" else "#### Measurement Cards by Type")
            
            num_vars = len(df_vars)
            if num_vars > 0:
                cols_cards = st.columns(4)
                accent_colors = ["#38BDF8", "#34D399", "#FBBF24", "#C084FC", "#F87171"]

                for idx, (_, row) in enumerate(df_vars.iterrows()):
                    var_name = row['Nombre de Variable']
                    tipo_var = str(row.get('Tipo de Dato', 'FLOAT')).upper()
                    unidad = row['Unidad'] if pd.notna(row['Unidad']) else ""
                    val = valores_actuales.get(var_name, False if tipo_var == "BOOL" else 0)
                    
                    estado = txt_fail
                    for item in lista_diagnostico:
                        if item["Variable"] == var_name:
                            estado = item["Estado de Conexión"]
                            break

                    glow_color = accent_colors[idx % len(accent_colors)] if (txt_conn in estado or txt_sim in estado) else "#64748B"
                    col_target = cols_cards[idx % 4]
                    
                    if tipo_var == "BOOL":
                        estado_bool = bool(val)
                        badge_color = "#10B981" if estado_bool else "#EF4444"
                        val_str = "ON" if estado_bool else "OFF"
                        col_target.markdown(f"""
                        <div class="card-scada" style="border: 1px solid {glow_color};">
                            <div style="font-weight: bold; color: {glow_color}; font-size: 13px;">{var_name} <span style="font-size:11px; opacity:0.8;">(BOOL)</span></div>
                            <div style="font-weight: 800; font-size: 20px; margin: 8px 0px; color: {badge_color};">{val_str}</div>
                            <div style="font-size: 11px; color: #94A3B8;">{estado}</div>
                        </div>
                        """, unsafe_allow_html=True)

                    elif tipo_var in ["INT", "DINT"]:
                        val_int = int(val) if isinstance(val, (int, float, bool)) else 0
                        col_target.markdown(f"""
                        <div class="card-scada" style="border: 1px solid {glow_color};">
                            <div style="font-weight: bold; color: {glow_color}; font-size: 13px;">{var_name} <span style="font-size:11px; opacity:0.8;">({tipo_var})</span></div>
                            <div style="font-weight: 800; font-size: 24px; margin: 8px 0px; color: #F8FAFC;">{val_int} <span style='font-size:14px; color:#94A3B8;'>{unidad}</span></div>
                            <div style="font-size: 11px; color: #94A3B8;">{estado}</div>
                        </div>
                        """, unsafe_allow_html=True)

                    else:  # FLOAT
                        val_num = float(val) if isinstance(val, (int, float)) else 0.0
                        col_target.markdown(f"""
                        <div class="card-scada" style="border: 1px solid {glow_color};">
                            <div style="font-weight: bold; color: {glow_color}; font-size: 13px;">{var_name} <span style="font-size:11px; opacity:0.8;">(FLOAT)</span></div>
                            <div style="font-weight: 800; font-size: 24px; margin: 8px 0px; color: #F8FAFC;">{val_num:.2f} <span style='font-size:14px; color:#94A3B8;'>{unidad}</span></div>
                            <div style="font-size: 11px; color: #94A3B8;">{estado}</div>
                        </div>
                        """, unsafe_allow_html=True)

        with tab_tre:
            st.markdown("### 📈 Tendencias OPC UA en Tiempo Real" if st.session_state['idioma'] == "ES" else "### 📈 Real-Time OPC UA Trends")
            opciones_vars = list(df_vars['Nombre de Variable'])
            if opciones_vars:
                var_sel = st.selectbox("Seleccione la variable:" if st.session_state['idioma'] == "ES" else "Select Variable:", opciones_vars, key="sb_trend_opc")
                if not st.session_state['mi_lista_de_datos'].empty and var_sel in st.session_state['mi_lista_de_datos'].columns:
                    fig = px.line(st.session_state['mi_lista_de_datos'].tail(30), x="Fecha_Hora", y=var_sel, template="plotly_dark")
                    fig.update_layout(
                        paper_bgcolor="#1E293B",
                        plot_bgcolor="#1E293B",
                        font_color="#E2E8F0"
                    )
                    st.plotly_chart(fig, use_container_width=True)

        with tab_his:
            st.markdown("### 📋 Histórico Registrado en Memoria" if st.session_state['idioma'] == "ES" else "### 📋 Recorded Data History")
            
            if not st.session_state['mi_lista_de_datos'].empty:
                st.dataframe(st.session_state['mi_lista_de_datos'], use_container_width=True)
                
                # Conversión del DataFrame a archivo Excel en memoria
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    st.session_state['mi_lista_de_datos'].to_excel(writer, index=False, sheet_name='Historico_SCADA')
                
                st.markdown("---")
                # Botón de Descarga Oficial de Streamlit
                st.download_button(
                    label="📥 Descargar Histórico en Excel (.xlsx)" if st.session_state['idioma'] == "ES" else "📥 Download History Excel (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"historico_scada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info("Aún no hay datos registrados en el historial." if st.session_state['idioma'] == "ES" else "No data recorded in history yet.")

    render_scada_live()

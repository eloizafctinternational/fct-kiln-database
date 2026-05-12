import streamlit as st
import pandas as pd
import plotly.express as px
import io

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Rotary Kiln Database", page_icon="⚙️", layout="wide")
st.title("⚙️ Rotary Kiln Comparison Tool")

# --- 2. CONEXÃO COM O DROPBOX ---
URL_DROPBOX = "https://www.dropbox.com/scl/fi/o3txb6fkr7lo4v1wlkdus/Rotary-Kill-Database.xlsx?rlkey=ra6g7p0tenr7zu74u5f0j4e7x&dl=1"

@st.cache_data(ttl=3600)
def carregar_dados_nuvem(url):
    return pd.read_excel(url)

try:
    df = carregar_dados_nuvem(URL_DROPBOX)
except Exception as e:
    st.error(f"⚠️ Erro ao conectar com a planilha: {e}")
    st.stop()

# --- 3. DICIONÁRIO DE APELIDOS ---
apelidos = {
    'client': 'Client', 'Year': 'Year', 'capacity_tpd': 'Capacity (TPD)', 'capacity_tpy': 'Capacity (TPY)',
    'heat_cons_kcal_kg': 'Heat Consumption (kcal/kg)', 'shell_diameter_m': 'Shell Diameter (m)',
    'length_m': 'Kiln Length (m)', 'l_d_ratio': 'L/D Ratio', 'speed_rpm': 'Speed (RPM)',
    'kiln_heat_system': 'Kiln Heat System', 'kiln_fuel_main': 'Main Fuel', 'cooler_type': 'Cooler Type',
    'total_power_kw': 'Total Power (kW)', 'material_type': 'Material Type', 'moisture_pct': 'Moisture (%)',
    'calcined_temp_out_C': 'Calcined Temp Out (°C)', 'total_exhaust_nm3_kg': 'Total Exhaust (Nm³/kg)',
    'primary_air_flow_rate_kg_h': 'Primary Air Flow Rate (kg/h)', 'kiln_exhaust_flow_rate_nm3_h': 'Kiln Exhaust Flow Rate to FGT (Nm³/h)',
    'dryer_exhaust_flow_rate_nm3_h': 'Dryer Exhaust Flow Rate (Nm³/h)', 'dryer_exhaust_temp_C': 'Dryer Exhaust Temperature (°C)',
    'cooler_exhaust_flow_rate_nm3_h': 'Cooler Exhaust Flow Rate (Nm³/h)', 'cooler_exhaust_temp_C': 'Cooler Exhaust Temperature (°C)',
    'fraction_to_gas_treatment_pct': 'Fraction to Gas Treatment (%)', 'fan_elec_cons_kw': 'Fan Elec Cons (kW)',
    'drives_elec_cons_kw': 'Drives Elec Cons (kW)'
}

# --- 4. BARRA LATERAL (FILTROS) ---
st.sidebar.header("🔍 1. Primary Filters")
Material = st.sidebar.selectbox("Material", ["Todos", "Clay", "Limestone", "Cement", "Other"])
Min_Capacity_TPD = st.sidebar.slider("Min Capacity (TPD)", 0, 5000, 0, 50)
Max_Consumption_kcal_kg = st.sidebar.slider("Max Consumption (kcal/kg)", 0, 3000, 3000, 10)

st.sidebar.header("2. Material Composition (Max %)")
Kaolinite_pct = st.sidebar.slider("Kaolinite (%)", 0, 100, 100, 1)
Illite_pct = st.sidebar.slider("Illite (%)", 0, 100, 100, 1)
Moisture_pct = st.sidebar.slider("Moisture (%)", 0, 100, 100, 1)
LOI_pct = st.sidebar.slider("LOI (%)", 0, 100, 100, 1)
Fe2O3_pct = st.sidebar.slider("Fe2O3 (%)", 0, 100, 100, 1)
Color_Control = st.sidebar.selectbox("Color Control", ["Todos", "yes", "no", "IM"])

st.sidebar.header("3. Kiln Dimensions (Max Values)")
Max_Diameter_m = st.sidebar.slider("Max Diameter (m)", 0.0, 10.0, 10.0, 0.1)
Max_Length_m = st.sidebar.slider("Max Length (m)", 0, 150, 150, 1)

st.sidebar.header("4. System Configuration")
Dryer = st.sidebar.selectbox("Dryer", ["Todos", "yes", "no"])
Dryer_Heat_System = st.sidebar.selectbox("Dryer Heat System", ["Todos", "HGG", "Burner"])
Kiln_Status = st.sidebar.selectbox("Kiln Status", ["Todos", "Novo", "Existente"])
Kiln_Heat_System = st.sidebar.selectbox("Kiln Heat System", ["Todos", "Burner", "Gasifier", "HGG"])
Cooler_Type = st.sidebar.selectbox("Cooler Type", ["Todos", "Rotary", "Grate"])

st.sidebar.header("📥 5. Export Options")
Download_Presentation_Data_Sheet = st.sidebar.checkbox("Show Presentation Download Button")

# --- 5. LÓGICA DE FILTRAGEM ---
df_res = df.copy()

if Material != "Todos": df_res = df_res[df_res['material_type'].str.contains(Material, case=False, na=False)]
if Min_Capacity_TPD > 0: df_res = df_res[df_res['capacity_tpd'] >= Min_Capacity_TPD]
if Max_Consumption_kcal_kg < 3000: df_res = df_res[df_res['heat_cons_kcal_kg'] <= Max_Consumption_kcal_kg]

if Kaolinite_pct < 100: df_res = df_res[(df_res['kaolinite_pct'] <= Kaolinite_pct) | df_res['kaolinite_pct'].isna()]
if Illite_pct < 100: df_res = df_res[(df_res['illite_pct'] <= Illite_pct) | df_res['illite_pct'].isna()]
if Moisture_pct < 100: df_res = df_res[(df_res['moisture_pct'] <= Moisture_pct) | df_res['moisture_pct'].isna()]
if LOI_pct < 100: df_res = df_res[(df_res['LOI_pct'] <= LOI_pct) | df_res['LOI_pct'].isna()]
if Fe2O3_pct < 100: df_res = df_res[(df_res['fe2o3_pct'] <= Fe2O3_pct) | df_res['fe2o3_pct'].isna()]
if Color_Control != "Todos": df_res = df_res[df_res['color_control'].str.contains(Color_Control, case=False, na=False)]

if Max_Diameter_m < 10: df_res = df_res[df_res['shell_diameter_m'] <= Max_Diameter_m]
if Max_Length_m < 150: df_res = df_res[df_res['length_m'] <= Max_Length_m]

if Dryer != "Todos": df_res = df_res[df_res['dryer'].str.lower() == Dryer.lower()]
if Dryer_Heat_System != "Todos": df_res = df_res[df_res['dryer_heat_system'].str.contains(Dryer_Heat_System, case=False, na=False)]
if Kiln_Status != "Todos": df_res = df_res[df_res['kiln_status'].str.contains(Kiln_Status, case=False, na=False)]
if Kiln_Heat_System != "Todos": df_res = df_res[df_res['kiln_heat_system'].str.contains(Kiln_Heat_System, case=False, na=False)]
if Cooler_Type != "Todos": df_res = df_res[df_res['cooler_type'].str.contains(Cooler_Type, case=False, na=False)]

# --- 6. EXIBIR RESULTADOS ---
if df_res.empty:
    st.warning("⚠️ No projects met the criteria.")
else:
    st.success(f"✅ Filter Complete! Found {len(df_res)} projects.")

    st.subheader("📈 EFFICIENCY ANALYSIS")
    resultado_graf = df_res.rename(columns=apelidos)
    fig = px.scatter(resultado_graf, 
                     x=apelidos['capacity_tpd'], 
                     y=apelidos['heat_cons_kcal_kg'], 
                     color=apelidos['client'],
                     hover_data=[apelidos['shell_diameter_m'], apelidos['length_m']], 
                     title="Capacity vs Consumption")
    fig.update_layout(hoverlabel=dict(bgcolor="white", font_color="black"))
    fig.update_traces(marker=dict(size=12, line=dict(width=1, color='Black')))
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("📋 DATA SHEET FOR SLIDES (Summary)")
    cols_summary = [
        'client', 'capacity_tpy', 'heat_cons_kcal_kg', 'shell_diameter_m', 'length_m',
        'calcined_temp_out_C', 'total_exhaust_nm3_kg', 'primary_air_flow_rate_kg_h',
        'kiln_exhaust_flow_rate_nm3_h', 'dryer_exhaust_flow_rate_nm3_h', 'dryer_exhaust_temp_C',
        'cooler_exhaust_flow_rate_nm3_h', 'cooler_exhaust_temp_C', 'fraction_to_gas_treatment_pct',
        'fan_elec_cons_kw', 'drives_elec_cons_kw', 'total_power_kw'
    ]
    col_pres = [c for c in cols_summary if c in df_res.columns]
    tabela_resumo_transposta = df_res[col_pres].rename(columns=apelidos).set_index(apelidos['client']).T
    st.dataframe(tabela_resumo_transposta)

    st.markdown("---")
    col1, col2 = st.columns(2)
    
    output_detail = io.BytesIO()
    with pd.ExcelWriter(output_detail, engine='openpyxl') as writer:
        df_res.to_excel(writer, index=False)
    
    with col1:
        st.download_button(
            label="📥 Download Detailed Results",
            data=output_detail.getvalue(),
            file_name="Rotary_Kiln_Filtered_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    if Download_Presentation_Data_Sheet:
        output_summary = io.BytesIO()
        with pd.ExcelWriter(output_summary, engine='openpyxl') as writer:
            tabela_resumo_transposta.to_excel(writer)
            
        with col2:
            st.download_button(
                label="📋 Download Presentation Data Sheet",
                data=output_summary.getvalue(),
                file_name="Presentation_Data_Sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

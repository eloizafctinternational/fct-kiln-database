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

# --- 3. DICIONÁRIO DE APELIDOS (COMPLETO) ---
apelidos = {
    # Informações Gerais do Projeto
    'client': 'Client',
    'Year': 'Year',
    'description': 'Description',
    'status': 'Status',
    'proposal_id': 'Proposal ID',
    'project_type': 'Project Type',
    'project_id': 'Project ID',
    'presentation_name': 'Presentation Name',
    'pfd_file_name': 'PFD File Name',
    'calc_file_name': 'Calculation File Name',
    
    # Material e Composição
    'material_type': 'Material Type',
    'kaolinite_pct': 'Kaolinite (%)',
    'illite_pct': 'Illite (%)',
    'montmori_pct': 'Montmorillonite (%)',
    'goetite_pct': 'Goethite (%)',
    'moisture_pct': 'Moisture (%)',
    'LOI_pct': 'Loss on Ignition (LOI %)',
    'fe2o3_pct': 'Fe2O3 (%)',
    'color_control': 'Color Control',
    
    # Desempenho e Dimensões do Forno
    'capacity_tpd': 'Capacity (TPD)',
    'capacity_tpy': 'Capacity (TPY)',
    'heat_cons_kcal_kg': 'Heat Consumption (kcal/kg)',
    'shell_diameter_m': 'Shell Diameter (m)',
    'length_m': 'Kiln Length (m)',
    'l_d_ratio': 'L/D Ratio',
    'slope_pct': 'Slope (%)',
    'gas_velocity_inlet_kiln_ms': 'Gas Velocity at Kiln Inlet (m/s)',
    'gas_velocity_cross_section_kiln_ms': 'Cross-Sectional Gas Velocity (m/s)',
    'filling_degree_pct': 'Filling Degree (%)',
    'retention_time_min': 'Retention Time (min)',
    'speed_rpm': 'Kiln Speed (RPM)',
    'kiln_dimension_m': 'Kiln Dimensions (m)',
    'lifters_present': 'Lifters Present',
    'kiln_status': 'Kiln Status',
    
    # Secador (Dryer)
    'dryer': 'Dryer Present',
    'dryer_type': 'Dryer Type',
    'dryer_heat_system': 'Dryer Heat System',
    'dryer_fuel': 'Dryer Fuel',
    'dryer_biomass_pct': 'Dryer Biomass (%)',
    'dryer_natural_gas_pct': 'Dryer Natural Gas (%)',
    'dryer_biocoal_pct': 'Dryer Biocoal (%)',
    'dryer_coal_pct': 'Dryer Coal (%)',
    'dryer_fuel_flow_kg_h': 'Dryer Fuel Flow Rate (kg/h)',
    
    # Sistema Térmico e Combustível do Forno
    'gasifier': 'Gasifier',
    'kiln_heat_system': 'Kiln Heat System',
    'kiln_fuel_main': 'Kiln Main Fuel(s)',
    'kiln_fuel_flow_kg_h': 'Kiln Fuel Flow Rate (kg/h)',
    'kiln_fuel_biomass_pct': 'Kiln Biomass (%)',
    'kiln_fuel_petcoke_pct': 'Kiln Petcoke (%)',
    'kiln_fuel_natural_gas_pct': 'Kiln Natural Gas (%)',
    'kiln_fuel_rdf_pct': 'Kiln RDF (%)',
    'kiln_fuel_coal_pct': 'Kiln Coal (%)',
    'kiln_fuel_biocoal_pct': 'Kiln Biocoal (%)',
    'HFO_pct': 'HFO (%)',
    'kiln_fuel_rice_husk_pct': 'Kiln Rice Husk (%)',
    
    # Resfriador (Cooler)
    'cooler_type': 'Cooler Type',
    'cooler_dimension_m': 'Cooler Dimensions (m)',
    'cooler_external_diameter_m': 'Cooler External Diameter (m)',
    'cooler_length_m': 'Cooler Length (m)',
    'cooler_mass_flow_kg_h': 'Cooler Mass Flow (kg/h)',
    'cooler_product_mass_flow_kg_h': 'Cooler Product Mass Flow (kg/h)',
    'water_injection': 'Water Injection',
    'water_injection_flow_m3_h': 'Water Injection Flow (m³/h)',
    
    # Temperaturas, Vazões e Exaustão
    'calcined_temp_out_C': 'Calcined Product Temp Out (°C)',
    'secondary_air_temp_C': 'Secondary Air Temp (°C)',
    'kiln_exhaust_temp_C': 'Kiln Exhaust Temp (°C)',
    'kiln_O2_pct': 'Kiln Exhaust O2 (%)',
    'total_exhaust_nm3_kg': 'Total Exhaust (Nm³/kg)',
    'primary_air_flow_rate_kg_h': 'Primary Air Flow Rate (kg/h)',
    'kiln_exhaust_flow_rate_kg_h': 'Kiln Exhaust Flow to FGT (kg/h)',
    'dryer_exhaust_flow_rate_kg_h': 'Dryer Exhaust Flow Rate (kg/h)',
    'dryer_exhaust_temp_C': 'Dryer Exhaust Temp (°C)',
    'cooler_exhaust_flow_rate_kg_h': 'Cooler Exhaust Flow Rate (kg/h)',
    'cooler_exhaust_temp_C': 'Cooler Exhaust Temp (°C)',
    'fraction_to_gas_treatment_pct': 'Fraction to Gas Treatment (%)',
    
    # Consumo Elétrico
    'main_drive_power_kw': 'Main Drive Power (kW)',
    'fans_power_kw': 'Fans Power (kW)',
    'total_power_kw': 'Total Installed Power (kW)',
    'fan_elec_cons_kwh_t': 'Fan Elec Cons (kWh/t)',
    'drives_elec_cons_kwh_t': 'Drives Elec Cons (kWh/t)'
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
    
    # Traduz as colunas
    resultado_graf = df_res.rename(columns=apelidos)
    
    # Prepara o hover 
    info_hover = []
    
    if 'description' in df_res.columns: 
        info_hover.append(apelidos['description'])
        
    if 'shell_diameter_m' in df_res.columns:
        info_hover.append(apelidos['shell_diameter_m'])
        
    if 'length_m' in df_res.columns:
        info_hover.append(apelidos['length_m'])
        
    if 'kiln_fuel_main' in df_res.columns: 
        info_hover.append(apelidos['kiln_fuel_main'])
        
    if 'moisture_pct' in df_res.columns: 
        info_hover.append(apelidos['moisture_pct'])

    # Cria o gráfico
    fig = px.scatter(resultado_graf, 
                     x=apelidos['capacity_tpd'], 
                     y=apelidos['heat_cons_kcal_kg'], 
                     color=apelidos['client'],
                     hover_data=info_hover, 
                     title="Capacity vs Consumption")
    
    fig.update_layout(hoverlabel=dict(bgcolor="white", font_color="black"))
    fig.update_traces(marker=dict(size=12, line=dict(width=1, color='Black')))
    st.plotly_chart(fig, use_container_width=True)

    # --- Tabela Resumo (Transposta) ---
    st.subheader("📊 TECHNICAL DATASHEET")
    
    cols_summary = [
        'client', 'description', 'Year', 
        'heat_cons_kcal_kg', 
        'kiln_fuel_main',      
        'kiln_fuel_flow_kg_h',
        'dryer_fuel',               
        'dryer_fuel_flow_kg_h',    
        'kiln_O2_pct',
        'primary_air_flow_rate_kg_h',
        'kiln_exhaust_flow_rate_kg_h', #to FGT
        'kiln_exhaust_temp_C',
        'dryer_exhaust_flow_rate_kg_h',
        'dryer_exhaust_temp_C',
        'cooler_exhaust_flow_rate_kg_h',
        'cooler_exhaust_temp_C',
        'capacity_tpy',
        'calcined_temp_out_C',
        'kiln_dimension_m', 
        'cooler_dimension_m', 
        'fraction_to_gas_treatment_pct',
        'total_exhaust_nm3_kg',
        'fan_elec_cons_kwh_t',
        'drives_elec_cons_kwh_t'
    ]
    
    col_pres = [c for c in cols_summary if c in df_res.columns]
    
    # O .reset_index(drop=True) limpa os números das linhas antigas
    tabela_resumo = df_res[col_pres].copy().reset_index(drop=True)
    
    # Geramos a etiqueta base (CLIENTE | Descrição ou Ano)
    def gerar_etiqueta_base(row):
        cliente = str(row['client']).upper()
        if 'description' in row and pd.notna(row['description']) and str(row['description']).strip() != "":
            return f"{cliente} | {str(row['description'])}"
        return f"{cliente} ({row['Year']})"

    tabela_resumo['Project_Header'] = tabela_resumo.apply(gerar_etiqueta_base, axis=1)

    # Numeração Automática de Duplicados
    counts = tabela_resumo.groupby('Project_Header').cumcount() + 1
    duplicated_mask = tabela_resumo['Project_Header'].duplicated(keep=False)
    
    tabela_resumo['Project_Header'] = tabela_resumo.apply(
        lambda x: f"{x['Project_Header']} ({counts[x.name]})" if duplicated_mask[x.name] else x['Project_Header'],
        axis=1
    )

    # Aplicamos os nomes em Inglês (Apelidos)
    tabela_resumo_final = tabela_resumo.rename(columns=apelidos)
    
# 5. Transposição e Limpeza
    # Removemos o que já está no título da coluna pegando os nomes traduzidos!
    colunas_remover = [
        apelidos.get('client', 'Client'), 
        apelidos.get('description', 'Description'), 
        apelidos.get('Year', 'Year'), 
        'Project_Header'
    ]
    
    # Definimos a etiqueta como índice e transpomos (.T)
    tabela_resumo_transposta = tabela_resumo_final.set_index(tabela_resumo_final.columns[-1]).drop(
        columns=[c for c in colunas_remover if c in tabela_resumo_final.columns], 
        errors='ignore'
    ).T
    
    # Exibimos a tabela
    st.dataframe(tabela_resumo_transposta, use_container_width=True)

    # --- Botões de Download ---
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

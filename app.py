import streamlit as st
import pandas as pd
import numpy as np
import io

# --- Configuración de la Página ---
st.set_page_config(
    page_title="Conversor a MobilServ",
    page_icon="🔄",
    layout="wide"
)

st.title("🔄 Conversor de Formato a MobilServ")

st.markdown("""
**Flujo:**
1. Sube uno o varios archivos Excel.
2. Visualiza los datos originales combinados.
3. Transforma el archivo al formato MobilServ.
4. Descarga el resultado en Excel.
""")

# --- Funciones auxiliares ---

def letter_to_index(letter):
    """Convierte letra de columna a índice base 0."""
    letter = letter.upper()
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord('A')) + 1
    return result - 1

def process_excel_file(df):
    """Transforma el DataFrame combinado al formato MobilServ."""
    
    # 1️⃣ Mapeo de columnas origen → destino
    movimientos = [
        ("A", "W"), ("Y", "B"), ("H", "C"), ("U", "E"), ("X", "F"), ("Z", "J"),
        ("V", "L"), ("W", "O"), ("E", "AA"), ("F", "AB"), ("C", "K"), ("D", "AH"),
        ("G", "AC"), ("I", "BB"), ("J", "BC"), ("K", "BD"), ("L", "BE"), ("M", "BF"),
        ("N", "BG"), ("O", "I"), ("B", "R"),
        ("IP", "FW"), ("MJ", "CC"), ("AJ", "CG"), ("FL", "CY"), ("BW", "DA"),
        ("IE", "DS"), ("PA", "GT"), ("MM", "FS"), ("JR", "ES"), ("JL", "EM"),
        ("OD", "GH"), ("OG", "EQ"), ("MO", "EE"), ("PE", "GX"),
        ("BJ", "CK"), ("BD", "CM"), ("BN", "CO"), ("BL", "CQ"), ("JF", "EI"),
        ("JG", "EK"), ("HQ", "FA"), ("PP", "HN"), ("BZ", "FK"),
        ("FB", "FM"), ("FC", "FO"), ("FA", "FQ"),
        ("KC", "EW"), ("JS", "EU"), ("JV", "GN"), ("JX", "GP"), ("JW", "GR"),
        ("IG", "GL"), ("GO", "DY"), ("AE", "HH"), ("CS", "HJ"), ("ER", "PI"),
        ("PH", "GZ"), ("PI", "HB"), ("C", "K"), ("CE", "EP")
    ]

    origen_indices = [letter_to_index(m[0]) for m in movimientos]
    destino_indices = [letter_to_index(m[1]) for m in movimientos]

    # Crear DataFrame destino suficientemente grande
    max_col_index = max(destino_indices)
    df_nuevo = pd.DataFrame(np.nan, index=df.index, columns=range(max_col_index + 1))

    # Mover datos
    for orig_idx, dest_idx in zip(origen_indices, destino_indices):
        if orig_idx < df.shape[1]:
            df_nuevo.iloc[:, dest_idx] = df.iloc[:, orig_idx].values

    # Encabezados MobilServ
    header_string = """Sample Status,Report Status,Date Reported,Asset ID,Unit ID,Unit Description,
Asset Class,Position,Tested Lubricant,Service Level,Sample Bottle ID,Manufacturer,
Alt Manufacturer,Model,Alt Model,Model Year,Serial Number,Account Name,Account ID,
Oil Rating,Contamination Rating,Equipment Rating,Parent Account Name,Parent Account ID,
ERP Account Number,Days Since Sampled,Date Sampled,Date Registered,Date Received,
Country,Laboratory,Business Lines,Fully Qualified,LIMS Sample ID,Schedule,
Tested Lubricant ID,Registered Lubricant,Registered Lubricant ID,Zone,Work ID,Sampler,
IMO No,Service Type,Component Type,Fuel Type,RPM,Cycles,Pressure,kW Rating,Cylinder Number,
Target PC 4,Target PC 6,Target PC 14,Equipment Age,Equipment UOM,Oil Age,Oil Age UOM,
Makeup Volume,MakeUp Volume UOM,Oil Changed,Filter Changed,Oil Temp In,Oil Temp Out,
Oil Temp UOM,Coolant Temp In,Coolant Temp Out,Coolant Temp UOM,Reservoir Temp,
Reservoir Temp UOM,Total Engine Hours,Hrs. Since Last Overhaul,Oil Service Hours,
Used Oil Volume,Used Oil Volume UOM,Oil Used in Last 24Hrs,Oil Used in Last 24Hrs UOM,
Sulphur %,Engine Power at Sampling,Date Landed,Port Landed,Ag (Silver),RESULT_Ag,
Al (Aluminum),RESULT_Al,B (Boron),RESULT_B,Ba (Barium),RESULT_Ba,Ca (Calcium),
RESULT_Ca,Cd (Cadmium),RESULT_Cd,Cl (Chlorine ppm - Xray),RESULT_Cl,Cr (Chromium),
RESULT_Cr,Cu (Copper),RESULT_Cu,K (Potassium),RESULT_K,Mg (Magnesium),RESULT_Mg,
Mn (Manganese),RESULT_Mn,Mo (Molybdenum),RESULT_Mo,Na (Sodium),RESULT_Na,
Ni (Nickel),RESULT_Ni,P (Phosphorus),RESULT_P,Zn (Zinc),RESULT_Zn,
ISO Code (4/6/14),RESULT_ISO Code (4/6/14),Particle Count >4um,RESULT_Particle Count >4um,
Particle Count >6um,RESULT_Particle Count >6um,Particle Count >14um,RESULT_Particle Count >14um,
Oxidation (Ab/cm),RESULT_Oxidation,Nitration (Ab/cm),RESULT_Nitration,
TAN (mg KOH/g),RESULT_TAN,TBN (mg KOH/g),RESULT_TBN,Soot (Wt%),RESULT_Soot,
Fuel Dilut. (Vol%),RESULT_Fuel Dilut.,Water (IR),RESULT_Water,Water KF,RESULT_Water KF,
Glycol %,RESULT_Glycol,Visc@100C (cSt),RESULT_Visc@100C,Visc@40C (cSt),RESULT_Visc@40C,
Sample ID"""
    headers = [h.strip() for h in header_string.split(",")]

    # Ajustar columnas al tamaño de headers
    df_final = df_nuevo.reindex(columns=range(len(headers)))
    df_final.columns = headers

    # Formatear fechas
    for col in ["Date Reported", "Date Sampled", "Date Registered", "Date Received"]:
        if col in df_final.columns:
            df_final[col] = pd.to_datetime(df_final[col], errors='coerce')

    # Marcar Sample Status
    if "Report Status" in df_final.columns and "Sample Status" in df_final.columns:
        mask = df_final["Report Status"].notna() & (df_final["Report Status"] != "")
        df_final.loc[mask, "Sample Status"] = "Completed"

    return df_final

def to_excel(df):
    """Exporta a Excel con formato básico."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', datetime_format='yyyy-mm-dd') as writer:
        df.to_excel(writer, index=False, sheet_name='MobilServ_Data')
    return output.getvalue()

# --- Interfaz ---

uploaded_files = st.file_uploader("📤 Sube uno o varios archivos Excel", type=["xlsx","xls"], accept_multiple_files=True)

if uploaded_files:
    dfs = []
    for f in uploaded_files:
        df = pd.read_excel(f, header=0, dtype=str)
        df["Archivo_Origen"] = f.name
        dfs.append(df)

    df_combined = pd.concat(dfs, ignore_index=True)
    st.subheader("📊 Vista previa de datos combinados")
    st.dataframe(df_combined.head(10))

    df_final = process_excel_file(df_combined)
    st.subheader("✅ Vista previa MobilServ")
    st.dataframe(df_final.head(10))

    excel_data = to_excel(df_final)
    st.download_button(
        label="📥 Descargar Excel MobilServ",
        data=excel_data,
        file_name="mobilserv_ordenado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

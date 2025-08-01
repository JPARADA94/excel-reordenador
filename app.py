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
**Flujo de la aplicación:**
1. Sube **uno o varios archivos Excel**.
2. Visualiza los datos originales combinados.
3. Transforma los datos al formato **MobilServ completo**.
4. Descarga el resultado final en Excel.
""")

# --- Utilitarios ---

def letter_to_index(letter):
    letter = letter.upper()
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord('A')) + 1
    return result - 1

def process_excel_file(df):
    # 1️⃣ Mapeo origen → destino
    movimientos = [
        ("A","W"),("Y","B"),("H","C"),("U","E"),("X","F"),("Z","J"),
        ("V","L"),("W","O"),("E","AA"),("F","AB"),("C","K"),("D","AH"),
        ("G","AC"),("I","BB"),("J","BC"),("K","BD"),("L","BE"),("M","BF"),
        ("N","BG"),("O","I"),("B","R"),
        ("IP","FW"),("MJ","CC"),("AJ","CG"),("FL","CY"),("BW","DA"),
        ("IE","DS"),("PA","GT"),("MM","FS"),("JR","ES"),("JL","EM"),
        ("OD","GH"),("OG","EQ"),("MO","EE"),("PE","GX"),
        ("BJ","CK"),("BD","CM"),("BN","CO"),("BL","CQ"),("JF","EI"),
        ("JG","EK"),("HQ","FA"),("PP","HN"),("BZ","FK"),
        ("FB","FM"),("FC","FO"),("FA","FQ"),
        ("KC","EW"),("JS","EU"),("JV","GN"),("JX","GP"),("JW","GR"),
        ("IG","GL"),("GO","DY"),("AE","HH"),("CS","HJ"),("ER","PI"),
        ("PH","GZ"),("PI","HB"),("C","K"),("CE","EP")
    ]
    orig_idx = [letter_to_index(o) for o,_ in movimientos]
    dest_idx = [letter_to_index(d) for _,d in movimientos]

    max_dest = max(dest_idx)
    df_nuevo = pd.DataFrame(np.nan, index=df.index, columns=range(max_dest+1))

    for oi,di in zip(orig_idx,dest_idx):
        if oi < df.shape[1]:
            df_nuevo.iloc[:,di] = df.iloc[:,oi].values

    # 2️⃣ Lista completa de encabezados
    full_headers = [
        "Sample Status","Report Status","Date Reported","Asset ID","Unit ID","Unit Description","Asset Class",
        "Position","Tested Lubricant","Service Level","Sample Bottle ID","Manufacturer","Alt Manufacturer",
        "Model","Alt Model","Model Year","Serial Number","Account Name","Account ID","Oil Rating",
        "Contamination Rating","Equipment Rating","Parent Account Name","Parent Account ID","ERP Account Number",
        "Days Since Sampled","Date Sampled","Date Registered","Date Received","Country","Laboratory","Business Lines",
        "Fully Qualified","LIMS Sample ID","Schedule","Tested Lubricant ID","Registered Lubricant",
        "Registered Lubricant ID","Zone","Work ID","Sampler","IMO No","Service Type","Component Type",
        "Fuel Type","RPM","Cycles","Pressure","kW Rating","Cylinder Number","Target PC 4","Target PC 6","Target PC 14",
        "Equipment Age","Equipment UOM","Oil Age","Oil Age UOM","Makeup Volume","MakeUp Volume UOM","Oil Changed",
        "Filter Changed","Oil Temp In","Oil Temp Out","Oil Temp UOM","Coolant Temp In","Coolant Temp Out",
        "Coolant Temp UOM","Reservoir Temp","Reservoir Temp UOM","Total Engine Hours","Hrs. Since Last Overhaul",
        "Oil Service Hours","Used Oil Volume","Used Oil Volume UOM","Oil Used in Last 24Hrs","Oil Used in Last 24Hrs UOM",
        "Sulphur %","Engine Power at Sampling","Date Landed","Port Landed","Ag (Silver)","RESULT_Ag",
        "Air Release @50 (min)","RESULT_Air Release @50 (min)","Al (Aluminum)","RESULT_Al","Appearance","RESULT_Appearance",
        "B (Boron)","RESULT_ B","Ba (Barium)","RESULT_Ba","Ca (Calcium)","RESULT_Ca","Cd (Cadmium)","RESULT_Cd",
        "Cl (Chlorine ppm - Xray)","RESULT_Cl (Chlorine ppm - Xray)","Compatibility","RESULT_Compatibility",
        "Coolant Indicator","RESULT_Coolant Indicator","Cr (Chromium)","RESULT_Cr","Cu (Copper)","RESULT_Cu",
        "DAC(%Asphalt.)","RESULT_DAC(%Asphalt.)","Demul@54C  (min)","RESULT_Demul@54C  (min)","Demul@54C (min)","RESULT_Demul@54C (min)",
        "Demul@54C (Oil/Water/Emul/Time)","RESULT_Demul@54C (Oil/Water/Emul/Time)","Demulsibility @54C (time-min)","RESULT_Demulsibility @54C (time-min)",
        "Demulsibility @54C (oil)","RESULT_Demulsibility @54C (oil)","Demulsibility @54C (water)","RESULT_Demulsibility @54C (water)",
        "Demulsibility @54C (emulsion)","RESULT_Demulsibility @54C (emulsion)","Fe (Iron)","RESULT_Fe (Iron)",
        "Foam Seq 1 - stability (ml)","RESULT_Foam Seq 1 - stability (ml)","Foam Seq 1 - tendency (ml)","RESULT_Foam Seq 1 - tendency (ml)",
        "Fuel Dilut. (Vol%)","RESULT_Fuel Dilut. (Vol%)","Initial pH","RESULT_Initial pH","Insolubles 5u","RESULT_Insolubles 5u",
        "K (Potassium)","RESULT_K","MCR","RESULT_MCR","Mg (Magnesium)","RESULT_Mg","Mn (Manganese)","RESULT_Mn (Manganese)",
        "Mo (Molybdenum)","RESULT_Mo","MPC delta E","RESULT_MPC delta E","Na (Sodium)","RESULT_Na","Ni (Nickel)","RESULT_Ni",
        "Nitration (Ab/cm)","RESULT_Nitration (Ab/cm)","Oxidation (Ab/cm)","RESULT_Oxidation (Ab/cm)","P  (Phosphorus)","RESULT_P  (Phosphorus)",
        "P (Phosphorus)","RESULT_P (Phosphorus)","ISO Code (4/6/14)","RESULT_ISO Code (4/6/14)","Particle Count  >4um","RESULT_Particle Count  >4um",
        "Particle Count  >6um","RESULT_Particle Count  >6um","Particle Count>14um","RESULT_Particle Count>14um","Diluted ISO Code (4/6/14)","RESULT_Diluted ISO Code (4/6/14)",
        "Particle Count (Diluted) >4um","RESULT_Particle Count (Diluted) >4um","Particle Count (Diluted) >6um","RESULT_Particle Count (Diluted) >6um",
        "Particle Count (Diluted) >14um","RESULT_Particle Count (Diluted) >14um","Pb (Lead)","RESULT_Pb","PM Flash Pt.(°C)","RESULT_PM Flash Pt.(°C)",
        "PQ Index","RESULT_PQ Index","RESULT_Product","RPVOT (Min)","RESULT_RPVOT (Min)","RULER-Amine (% vs new)","RESULT_RULER-Amine (% vs new)",
        "RULER-Phenol (% vs new)","RESULT_RULER-Phenol (% vs new)","SAE Viscosity Grade","RESULT_SAE Viscosity Grade","Si (Silicon)","RESULT_Si",
        "Sn (Tin)","RESULT_Sn","Soot (Wt%)","RESULT_Soot (Wt%)","TAN (mg KOH/g)","RESULT_TAN (mg KOH/g)","TBN (mg KOH/g)","RESULT_TBN (mg KOH/g) 2",
        "TBN (mg KOH/g)","RESULT_TBN (mg KOH/g) 2","Ti (Titanium)","RESULT_Ti","UC Rating","RESULT_UC Rating","V (Vanadium)","RESULT_V",
        "Visc@100C (cSt)","RESULT_Visc@100C (cSt)","Visc@40C (cSt)","RESULT_Visc@40C (cSt)","Water (Hot Plate)","RESULT_Water (Hot Plate)",
        "Water (Vol %)","RESULT_Water (Vol%)","Water (Vol%)","RESULT_Water (Vol%) 2","Water (Vol.)","RESULT_Water (Vol%) 3","Water Free est.","RESULT_Water Free est.","Zn (Zinc)"
    ]

    # 2️⃣ Reindexar y renombrar
    df_final = df_nuevo.reindex(columns=range(len(full_headers)))
    df_final.columns = full_headers

    # 3️⃣ Formatear fechas
    for col in ["Date Reported","Date Sampled","Date Registered","Date Received"]:
        if col in df_final.columns:
            df_final[col] = pd.to_datetime(df_final[col], errors='coerce')

    # 4️⃣ Sample Status automático
    if "Report Status" in df_final.columns and "Sample Status" in df_final.columns:
        mask = df_final["Report Status"].notna() & (df_final["Report Status"]!="")
        df_final.loc[mask,"Sample Status"] = "Completed"

    return df_final

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', datetime_format='yyyy-mm-dd') as writer:
        df.to_excel(writer, index=False, sheet_name='MobilServ_Data')
    return output.getvalue()

# --- UI ---

uploaded = st.file_uploader("📤 Sube archivos Excel", type=["xlsx"], accept_multiple_files=True)
if uploaded:
    dfs = []
    for f in uploaded:
        tmp = pd.read_excel(f, header=0, dtype=str)
        tmp["Archivo_Origen"] = f.name
        dfs.append(tmp)
    df_comb = pd.concat(dfs, ignore_index=True)

    st.subheader("📊 Datos combinados originales")
    st.dataframe(df_comb.head(10))

    df_final = process_excel_file(df_comb)

    # 🌟 Renombrar duplicados solo para vista previa
    seen={}
    preview_cols=[]
    for c in df_final.columns:
        if c not in seen:
            seen[c]=0; preview_cols.append(c)
        else:
            seen[c]+=1; preview_cols.append(f"{c} ({seen[c]})")
    preview_df = pd.DataFrame(df_final.head(10).values, columns=preview_cols)

    st.subheader("✅ Preview MobilServ (sin error de duplicados)")
    st.dataframe(preview_df)

    data = to_excel(df_final)
    st.download_button("📥 Descargar MobilServ", data=data,
                       file_name="mobilserv.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


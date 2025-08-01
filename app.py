import streamlit as st
import pandas as pd
from io import BytesIO

# —————— Configuración de la página ——————
st.set_page_config(page_title="Reordenador Excel a formato MobilServ", layout="wide")

st.markdown("**Creado por:** Javier Parada  \n**Ingeniero de Soporte en Campo**")
st.title("Reordenador Excel a formato MobilServ – Múltiples Archivos")

# —————— Instrucciones ——————
st.markdown("""
**Cómo usar esta herramienta:**
1. Sube uno o varios archivos Excel (.xlsx).
2. Revisa la vista previa de los datos originales.
3. El sistema aplicará el reordenamiento a todos los archivos.
4. Descarga el archivo consolidado final.
""")

# —————— Utilitario: columna letra → índice ——————
def col_letter_to_index(letter: str) -> int:
    idx = 0
    for c in letter.upper():
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx - 1

# —————— Mapeo actualizado (según tabla final) ——————
mapping_text = """
A W
Y B
H C
U E
X F
Z J
V L
W O
E AA
F AB
G AC
I BB
J BC
K BD
L BE
M BF
N BG
O I
B R
IP FW
MJ CC
AJ CG
FL CY
BW DA
IE DS
PA GT
MM FS
JR ES
JL EM
OD GH
OG EQ
MO EE
PE GX
BJ CK
BD CM
BN CO
BL CQ
JF EI
JG EK
HQ FA
PP HN
BZ FK
FB FM
FC FO
FA FQ
KC EW
JS EU
JV GN
JX GP
JW GR
IG GL
GO DY
AE HH
CS HJ
ER PI
PH GZ
PI HB
C K
CE EP
""".strip()

MOVIMIENTOS = [tuple(line.split()) for line in mapping_text.splitlines()]

# —————— Encabezados destino ——————
headerString = """Sample Status,Report Status,Date Reported,Asset ID,Unit ID,Unit Description,
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
Al (Aluminum),RESULT_Al,B (Boron),RESULT_Ba,Ba (Barium),RESULT_Ba,Ca (Calcium),
RESULT_Ca,Cd (Cadmium),RESULT_Cd,Cl (Chlorine ppm - Xray),RESULT_Cl,Cr (Chromium),
RESULT_Cr,Cu (Copper),RESULT_Cu,K (Potassium),RESULT_K,Mg (Magnesium),RESULT_Mg,
Mn (Manganese),RESULT_Mn,Mo (Molybdenum),RESULT_Mo,Na (Sodium),RESULT_Na,
Ni (Nickel),RESULT_Ni,P  (Phosphorus),RESULT_P,Zn (Zinc),RESULT_Zn,
ISO Code (4/6/14),RESULT_ISO Code (4/6/14),Particle Count >4um,RESULT_Particle Count >4um,
Particle Count >6um,RESULT_Particle Count >6um,Particle Count >14um,RESULT_Particle Count >14um,
Oxidation (Ab/cm),RESULT_Oxidation,Nitration (Ab/cm),RESULT_Nitration,
TAN (mg KOH/g),RESULT_TAN,TBN (mg KOH/g),RESULT_TBN,Soot (Wt%),RESULT_Soot,
Fuel Dilut. (Vol%),RESULT_Fuel Dilut.,Water (IR),RESULT_Water,Water KF,RESULT_Water KF,
Glycol %,RESULT_Glycol,Visc@100C (cSt),RESULT_Visc@100C,Visc@40C (cSt),RESULT_Visc@40C,
Sample ID
""".replace("\n", "")

header_list = [h.strip() for h in headerString.split(",")]

# —————— Columnas especiales ——————
DATE_COLS   = ["Date Reported","Date Sampled","Date Registered","Date Received"]
INT_LETTERS = ["BB","BD","BF","CC","CG","CK","CM","CO","CQ","CY","DA","DS","EE","EI","EK","EM","EQ","ES","EW","FA","FM","FO","FQ","FS","FW","GH","GT","GX","HN"]
DEC_LETTERS = ["DY","GL","GN","GP","GR","GZ","HB","HH","HJ"]

# —————— Subida de múltiples archivos ——————
uploaded_files = st.file_uploader("📤 Sube uno o varios archivos Excel (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    df_consolidado = pd.DataFrame()

    for uploaded in uploaded_files:
        df = pd.read_excel(uploaded, header=0, dtype=str)
        df["Archivo_Origen"] = uploaded.name

        st.subheader(f"📌 Vista previa – {uploaded.name}")
        st.dataframe(df.head(5))

        max_dest = max(col_letter_to_index(d) for _, d in MOVIMIENTOS)
        result = pd.DataFrame(index=df.index, columns=range(max_dest + 1))

        for orig, dest in MOVIMIENTOS:
            i = col_letter_to_index(orig)
            j = col_letter_to_index(dest)
            if i < df.shape[1]:
                result.iloc[:, j] = df.iloc[:, i]
            else:
                result.iloc[:, j] = None

        if result.shape[1] > len(header_list):
            result = result.iloc[:, :len(header_list)]
        result.columns = header_list[:result.shape[1]]

        # Evita duplicados en encabezados
        seen = {}
        unique_cols = []
        for col in result.columns:
            if col not in seen:
                seen[col] = 0
                unique_cols.append(col)
            else:
                seen[col] += 1
                unique_cols.append(f"{col} ({seen[col]})")
        result.columns = unique_cols

        # Tipos: fechas, enteros, decimales
        for c in DATE_COLS:
            if c in result:
                result[c] = pd.to_datetime(result[c], errors="coerce").dt.date
        for letter in INT_LETTERS:
            idx = col_letter_to_index(letter)
            if idx < result.shape[1]:
                result.iloc[:, idx] = pd.to_numeric(result.iloc[:, idx], errors="coerce").astype("Int64")
        for letter in DEC_LETTERS:
            idx = col_letter_to_index(letter)
            if idx < result.shape[1]:
                result.iloc[:, idx] = pd.to_numeric(result.iloc[:, idx], errors="coerce").round(2)

        if "Report Status" in result and "Sample Status" in result:
            result.loc[result["Report Status"].notna(), "Sample Status"] = "Completed"

        # Agregar columna de origen al DataFrame final
        result["Archivo_Origen"] = uploaded.name
        df_consolidado = pd.concat([df_consolidado, result], ignore_index=True)

    # Vista previa consolidada
    st.subheader("✅ Vista previa – Archivo consolidado final")
    st.dataframe(df_consolidado.head(10))

    # Descarga del archivo consolidado
    buffer = BytesIO()
    df_consolidado.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.download_button(
        label="📥 Descargar Excel consolidado",
        data=buffer,
        file_name="mobilserv_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

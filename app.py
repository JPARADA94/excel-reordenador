import streamlit as st
import pandas as pd
from io import BytesIO

# —————— Configuración de la página ——————
st.set_page_config(page_title="Reordenador Excel a formato MobilServ", layout="wide")

st.markdown("**Creado por:** Javier Parada  \n**Ingeniero de Soporte en Campo**")
st.title("Reordenador Excel a formato MobilServ – Múltiples Archivos (Encabezados RESULT vacíos)")

# —————— Instrucciones ——————
st.markdown("""
**Flujo de la herramienta:**
1. Sube **uno o varios archivos Excel (.xlsx)**.
2. El sistema combinará todos los archivos en un solo DataFrame.
3. Verás una **vista previa de los datos combinados originales**.
4. Luego, el sistema aplicará el **reordenamiento MobilServ** (las columnas `RESULT_XXX` solo son encabezados).
5. Se mostrará la **vista previa final sin errores**.
6. Finalmente podrás **descargar el Excel final MobilServ**.
""")

# —————— Utilitario: columna letra → índice ——————
def col_letter_to_index(letter: str) -> int:
    idx = 0
    for c in letter.upper():
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx - 1

# —————— Mapeo columnas origen → MobilServ ——————
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

# —————— Lista de encabezados MobilServ final ——————
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

# —————— Subida de múltiples archivos ——————
uploaded_files = st.file_uploader("📤 Sube uno o varios archivos Excel (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    # 1️⃣ Combinar todos los archivos en un solo DataFrame
    dataframes = []
    for uploaded in uploaded_files:
        df = pd.read_excel(uploaded, header=0, dtype=str)
        df["Archivo_Origen"] = uploaded.name
        dataframes.append(df)

    df_consolidado = pd.concat(dataframes, ignore_index=True)

    # 2️⃣ Vista previa de datos originales combinados
    st.subheader("📌 Vista previa – Datos combinados originales")
    st.dataframe(df_consolidado.head(10))

    # 3️⃣ Aplicar reordenamiento MobilServ
    max_dest = max(col_letter_to_index(d) for _, d in MOVIMIENTOS)
    result = pd.DataFrame(index=df_consolidado.index, columns=range(max_dest + 1))

    for orig, dest in MOVIMIENTOS:
        i = col_letter_to_index(orig)
        j = col_letter_to_index(dest)
        if i < df_consolidado.shape[1]:
            result.iloc[:, j] = df_consolidado.iloc[:, i]
        else:
            result.iloc[:, j] = None

    # Ajustar encabezados al formato MobilServ
    if result.shape[1] > len(header_list):
        result = result.iloc[:, :len(header_list)]
    result.columns = header_list[:result.shape[1]]

    # 4️⃣ Crear versión temporal sin duplicados para la vista previa
    preview_cols = []
    seen = {}
    for col in result.columns:
        if col not in seen:
            seen[col] = 0
            preview_cols.append(col)
        else:
            seen[col] += 1
            preview_cols.append(f"{col} ({seen[col]})")

    st.subheader("✅ Vista previa – Archivo ya reordenado MobilServ")
    st.dataframe(pd.DataFrame(result.head(10).values, columns=preview_cols))

    # 5️⃣ Agregar columna de origen al final y descargar Excel final
    result["Archivo_Origen"] = df_consolidado["Archivo_Origen"]

    buffer = BytesIO()
    result.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.download_button(
        label="📥 Descargar Excel MobilServ final",
        data=buffer,
        file_name="mobilserv_ordenado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

import streamlit as st
import pandas as pd
from io import BytesIO

# —————— Configuración de la página ——————
st.set_page_config(page_title="Reordenador Excel a formato MobilServ", layout="wide")

st.markdown("**Creado por:** Javier Parada  \n**Ingeniero de Soporte en Campo**")
st.title("Reordenador Excel a formato MobilServ – Validación y Fechas Formateadas")

st.markdown("""
**Flujo de la herramienta:**
1. Sube **uno o varios archivos Excel (.xlsx)**.
2. La app **validará que los encabezados coincidan con los esperados**.
3. Se combinarán todos los archivos en un solo DataFrame.
4. Vista previa **original** de los datos combinados.
5. Reordenamiento MobilServ, vista previa final sin errores.
6. Descarga de **Excel final MobilServ** con:
   - Columnas `RESULT_XXX` vacías
   - Columnas de fecha en formato `yyyy-mm-dd`
   - Resto de datos **idénticos a los originales**
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

# —————— Encabezados MobilServ Final ——————
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

# Columnas de fecha
DATE_COLS = ["Date Reported", "Date Sampled", "Date Registered", "Date Received"]

# —————— Subida de múltiples archivos ——————
uploaded_files = st.file_uploader("📤 Sube uno o varios archivos Excel (.xlsx)", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    dataframes = []

    for uploaded in uploaded_files:
        df = pd.read_excel(uploaded, header=0, dtype=str)

        # 1️⃣ Validación de encabezados
        expected_cols = [c for c, _ in MOVIMIENTOS]
        file_cols = list(df.columns[:len(expected_cols)])
        if len(file_cols) < len(expected_cols):
            st.error(f"❌ El archivo `{uploaded.name}` tiene menos columnas de las esperadas.")
            st.stop()

        # Convertimos letras a nombres esperados para validación
        errores = []
        for idx, (letter, _) in enumerate(MOVIMIENTOS):
            if idx < len(file_cols):
                if df.columns[idx].strip() != df.columns[idx].strip():
                    errores.append(f"Columna {letter}: Esperada `{df.columns[idx]}`, encontrada `{file_cols[idx]}`")

        if errores:
            st.error(f"❌ Encabezados incorrectos en `{uploaded.name}`:\n" + "\n".join(errores))
            st.stop()

        df["Archivo_Origen"] = uploaded.name
        dataframes.append(df)

    # 2️⃣ Combinar DataFrames
    df_consolidado = pd.concat(dataframes, ignore_index=True)

    st.subheader("📌 Vista previa – Datos combinados originales")
    st.dataframe(df_consolidado.head(10))

    # 3️⃣ Crear DataFrame reordenado MobilServ
    max_dest = max(col_letter_to_index(d) for _, d in MOVIMIENTOS)
    result = pd.DataFrame(index=df_consolidado.index, columns=range(max_dest + 1))

    for orig, dest in MOVIMIENTOS:
        i = col_letter_to_index(orig)
        j = col_letter_to_index(dest)
        if i < df_consolidado.shape[1]:
            result.iloc[:, j] = df_consolidado.iloc[:, i]
        else:
            result.iloc[:, j] = None

    # Ajustar encabezados MobilServ
    if result.shape[1] > len(header_list):
        result = result.iloc[:, :len(header_list)]
    result.columns = header_list[:result.shape[1]]

    # 4️⃣ Vista previa sin errores por duplicados
    preview_cols = []
    seen = {}
    for col in result.columns:
        if col not in seen:
            seen[col] = 0
            preview_cols.append(col)
        else:
            seen[col] += 1
            preview_cols.append(f"{col} ({seen[col]})")

    st.subheader("✅ Vista previa – Archivo reordenado MobilServ")
    st.dataframe(pd.DataFrame(result.head(10).values, columns=preview_cols))

    # 5️⃣ Exportar con formato de fechas y datos originales
    result["Archivo_Origen"] = df_consolidado["Archivo_Origen"]

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        result.to_excel(writer, index=False, sheet_name="MobilServ")

        workbook = writer.book
        worksheet = writer.sheets["MobilServ"]

        # Formato de fecha para columnas DATE_COLS
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

        for col_idx, col_name in enumerate(result.columns):
            if col_name in DATE_COLS:
                worksheet.set_column(col_idx, col_idx, 15, date_format)
            else:
                worksheet.set_column(col_idx, col_idx, 20)

    buffer.seek(0)

    st.download_button(
        label="📥 Descargar Excel MobilServ final",
        data=buffer,
        file_name="mobilserv_ordenado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# -*- coding: utf-8 -*-
"""APP
Reordenador Excel a formato MobilServ con preview de origen y destino,
y manejo dinámico de encabezados para evitar ValueError al asignar columnas.
"""

import streamlit as st
import pandas as pd
from io import BytesIO

# —————— Configuración de la página ——————
st.set_page_config(page_title="Reordenador Excel a formato MobilServ", layout="wide")

# —————— Crédito del autor ——————
st.markdown("**Creado por:** Javier Parada  \n**Ingeniero de Soporte en Campo**")
st.title("Reordenador Excel a formato MobilServ")

# —————— Instrucciones de uso ——————
st.markdown("**Cómo usar esta herramienta:**")
st.markdown(
    """
    1. Sube tu archivo Excel (.xlsx).
    2. Revisa la vista previa de los datos originales.
    3. Espera a que se procese y revisa la vista previa del archivo reordenado.
    4. Haz clic en "📥 Descargar Excel reordenado" para obtener tu archivo.
    """
)

# —————— Utilitario: Columna Excel → índice 0-based ——————
def col_letter_to_index(letter: str) -> int:
    idx = 0
    for c in letter.upper():
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx - 1

# —————— Mapeo actualizado (columna origen → columna destino) ——————
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
C K    # N_MUESTRA → Sample Bottle ID
I BB
J BC
K BD
L BE
M BF
N BG
O I
B R
IO FW
MI CC
AJ CG
FK CY
BV DA
IE DS
OZ GT
MK FS
JQ ES
JJ EM
OB GH
OG EQ
MM EE
PD GX
BI CK
BD CM
BM CO
BL CQ
JE EI
JF EK
HQ FA
PO HN
BZ FK
FB FM
FC FO
FA FQ
KB EW
JR EU
JU GN
JW GP
JV GR
IG GL
GO DY
AE HH
CS HJ
ER PI
PG GZ
PH HB
""".strip()

MOVIMIENTOS = [tuple(line.split()[0:2]) for line in mapping_text.splitlines()]

# —————— Encabezados destino (pegados de tu macro VBA) ——————
headerString = """
Sample Status,Report Status,Date Reported,Asset ID,Unit ID,Unit Description,Asset Class,Position,
Tested Lubricant,Service Level,Sample Bottle ID,Manufacturer,Alt Manufacturer,Model,Alt Model,
Model Year,Serial Number,Account Name,Account ID,Oil Rating,Contamination Rating,Equipment Rating,
Parent Account Name,Parent Account ID,ERP Account Number,Days Since Sampled,Date Sampled,
Date Registered,Date Received,Country,Laboratory,Business Lines,Fully Qualified,LIMS Sample ID,
Schedule,Tested Lubricant ID,Registered Lubricant,Registered Lubricant ID,Zone,Work ID,Sampler,
IMO No,Service Type,Component Type,Fuel Type,RPM,Cycles,Pressure,kW Rating,Cylinder Number,
Target PC 4,Target PC 6,Target PC 14,Equipment Age,Equipment UOM,Oil Age,Oil Age UOM,Makeup Volume,
MakeUp Volume UOM,Oil Changed,Filter Changed,Oil Temp In,Oil Temp Out,Oil Temp UOM,Coolant Temp In,
Coolant Temp Out,Coolant Temp UOM,Reservoir Temp,Reservoir Temp UOM,Total Engine Hours,
Hrs. Since Last Overhaul,Oil Service Hours,Used Oil Volume,Used Oil Volume UOM,
Oil Used in Last 24Hrs,Oil Used in Last 24Hrs UOM,Sulphur %,Engine Power at Sampling,Date Landed,
Port Landed,Ag (Silver),RESULT_Ag,Air Release @50 (min),RESULT_Air Release @50 (min),Al (Aluminum),
RESULT_Al,Appearance,RESULT_Appearance,B (Boron),RESULT_ B,Ba (Barium),RESULT_Ba,Ca (Calcium),
RESULT_Ca,Cd (Cadmium),RESULT_Cd,Cl (Chlorine ppm - Xray),RESULT_Cl (Chlorine ppm - Xray),
Compatibility,RESULT_Compatibility,Coolant Indicator,RESULT_Coolant Indicator,Cr (Chromium),RESULT_Cr,
Cu (Copper),RESULT_Cu,DAC(%Asphalt.),RESULT_DAC(%Asphalt.),Demul@54C  (min),RESULT_Demul@54C  (min),
Demul@54C (min),RESULT_Demul@54C (min),Demul@54C (Oil/Water/Emul/Time),RESULT_Demul@54C (Oil/Water/Emul/Time),
Demulsibility @54C (time-min),RESULT_Demulsibilidad @54C (time-min),Demulsibility @54C (oil),RESULT_Demulsibilidad @54C (oil),
Demulsibility @54C (water),RESULT_Demulsibilidad @54C (water),Demulsibility @54C (emulsion),RESULT_Demulsibilidad @54C (emulsion),
Fe (Iron),RESULT_Fe (Iron),Foam Seq 1 - stability (ml),RESULT_Foam Seq 1 - stability (ml),Foam Seq 1 - tendency (ml),
RESULT_Foam Seq 1 - tendency (ml),Fuel Dilut. (Vol%),RESULT_Fuel Dilut. (Vol%),Initial pH,RESULT_Initial pH,
Insolubles 5u,RESULT_Insolubles 5u,K (Potassium),RESULT_K,MCR,RESULT_MCR,Mg (Magnesium),RESULT_Mg,
Mn (Manganese),RESULT_Mn (Manganese),Mo (Molybdenum),RESULT_Mo,MPC delta E,RESULT_MPC delta E,Na (Sodium),
RESULT_Na,Ni (Nickel),RESULT_Ni,Nitration (Ab/cm),RESULT_Nitration (Ab/cm),Oxidation (Ab/cm),
RESULT_Oxidation (Ab/cm),P  (Phosphorus),RESULT_P  (Phosphorus),P (Phosphorus),RESULT_P (Phosphorus),
ISO Code (4/6/14),RESULT_ISO Code (4/6/14),Particle Count  >4um,RESULT_Particle Count  >4um,
Particle Count  >6um,RESULT_Particle Count  >6um,Particle Count>14um,RESULT_Particle Count>14um,
Diluted ISO Code (4/6/14),RESULT_Diluted ISO Code (4/6/14),Particle Count (Diluted) >4um,
RESULT_Particle Count (Diluted) >4um,Particle Count (Diluted) >6um,RESULT_Particle Count (Diluted) >6um,
Particle Count (Diluted) >14um,RESULT_Particle Count (Diluted) >14um,Pb (Lead),RESULT_Pb,PM Flash Pt.(°C),
RESULT_PM Flash Pt.(°C),PQ Index,RESULT_PQ Index,RESULT_Product,RPVOT (Min),RESULT_RPVOT (Min),
RULER-Amine (% vs new),RESULT_RULER-Amine (% vs new),RULER-Phenol (% vs new),RESULT_RULER-Phenol (% vs new),
SAE Viscosity Grade,RESULT_SAE Viscosity Grade,Si (Silicon),RESULT_Si,Sn (Tin),RESULT_Sn,Soot (Wt%),RESULT_Soot (Wt%),
TAN (mg KOH/g),RESULT_TAN (mg KOH/g),TBN (mg KOH/g),RESULT_TBN (mg KOH/g) 2,TBN (mg KOH/g),
RESULT_TBN (mg KOH/g) 2,Ti (Titanium),RESULT_Ti,UC Rating,RESULT_UC Rating,V (Vanadium),RESULT_V,
Visc@100C (cSt),RESULT_Visc@100C (cSt),Visc@40C (cSt),RESULT_Visc@40C (cSt),Water (Hot Plate),
RESULT_Water (Hot Plate),Water (Vol %),RESULT_Water (Vol%),Water (Vol%),RESULT_Water (Vol%) 2,
Water (Vol.%),RESULT_Water (Vol%) 3,Water Free est.,RESULT_Water Free est.,Zn (Zinc),RESULT_Zn,
Zn (Zinc) 2,RESULT_Zn 2,Soot (Wt%)- No Ref,RESULT_Soot (Wt%)- No Ref,Oxidation (Abs/cm)- no Ref,
RESULT_Oxidation (Abs/cm)- no Ref,Nitration (Abs/cm)- no Ref,RESULT_Nitration (Abs/cm)- no Ref,
Water (Abs/cm)- no Ref,RESULT_Water (Abs/cm) - no Ref,Aluminum - GR,RESULT_Aluminum - GR,
Antimony - gr,RESULT_Antimony - gr,Appearance - gr,RESULT_Appearance - gr,Barium - GR,RESULT_Barium - GR,
Boron - GR,RESULT_Boron - GR,Cadmium - gr,RESULT_Cadmium - gr,Calcium - GR,RESULT_Calcium - GR,
Chromium - gr,RESULT_Chromium - gr,Copper - GR,RESULT_Copper - GR,IR Correlation - gr,RESULT_IR Correlation - gr,
Ferrous Debris - gr,RESULT_Ferrous Debris - gr,Stress Index - Gr,RESULT_Stress Index - Gr,Grease Thief Video,
RESULT_Grease Thief Video,Iron - GR,RESULT_Iron - GR,Lead - gr,RESULT_Lead - gr,Magnesium - GR,
RESULT_Magnesium - GR,Manganese - GR,RESULT_Manganese - GR,Molybdneum - gr,RESULT_Molybdneum - gr,
Nickel - gr,RESULT_Nickel - gr,Phosphorus - GR,RESULT_Phosphorus - GR,Potassium - Gr,RESULT_Potassium - Gr,
Silicon - gr,RESULT_Silicon - gr,Silver - Grease,RESULT_Silver - Grease,Sodium - Gr,RESULT_Sodium - Gr,
Tin - gr,RESULT_Tin - gr,Titanium - gr,RESULT_Titanium - gr,Vanadium - gr,RESULT_Vanadium - gr,
Water - Gr,RESULT_Water - Gr,Zinc - gr,RESULT_Zinc - gr,Fuel Dilution - INDO,RESULT_Fuel Dilution - INDO,
TBN - INDO,RESULT_TBN - INDO,Soot - INDO,RESULT_Soot - INDO,Water - INDO,RESULT_Water - INDO,
Oxidation - INDO,RESULT_Oxidation - INDO,Nitration - INDO,RESULT_Nitration - INDO,
Boron,RESULT_Boron,Barium,RESULT_Barium,Calcium,RESULT_Calcium,Magnesium,RESULT_Magnesium,
Lithium -gr,RESULT_Lithium -gr,Color -gr,RESULT_Color -gr,Chlorine,RESULT_Chlorine,Lithium,RESULT_Lithium,
Antimony,RESULT_Antimony,Sulfur,RESULT_Sulfur,Insolubles,RESULT_Insolubles,Aluminum - gr - ICP,
RESULT_Aluminum - gr - ICP,Antimony - gr- ICP,RESULT_Antimony - gr- ICP,Barium - gr - ICP,
RESULT_Barium - gr - ICP,Boron - gr - ICP,RESULT_Boron - gr - ICP,Cadmium - gr - ICP,
RESULT_Cadmium - gr - ICP,Calcium - gr - ICP,RESULT_Calcium - gr - ICP,Chromium - gr - ICP,
RESULT_Chromium - gr - ICP,Copper - gr - ICP,RESULT_Copper - gr - ICP,Iron - gr - ICP,
RESULT_Iron - gr - ICP,Lead - gr - ICP,RESULT_Lead - gr - ICP,Lithium - gr - ICP,
RESULT_Lithium - gr - ICP,Magnesium - gr - ICP,RESULT_Magnesium - gr - ICP,Manganese - gr - ICP,
RESULT_N NAS particles 5-15um,RESULT_NAS particles 5-15um,NAS particles 15-25um,RESULT_NAS particles 15-25um,NAS particles 25-50um,RESULT_NAS particles 25-50um,NAS particles 50-100um,RESULT_NAS particles 50-100um,NAS particles > 100um,RESULT_NAS particles > 100um,Glycol %,RESULT_Glycol %,Blotter Spot C-Index,RESULT_Blotter Spot C-Index,Blotter Spot Diameter,RESULT_Blotter Spot Diameter,Blotter Spot Dispersancy,RESULT_Blotter Spot Dispersancy,Blotter Spot Opacity,RESULT_Blotter Spot Opacity,Blotter Spot Note,RESULT_Blotter Spot Note
""".replace("\n", "")

header_list = [h.strip() for h in headerString.split(",")]

# —————— Columnas especiales ——————
DATE_COLS   = ["Date Reported","Date Sampled","Date Registered","Date Received"]
INT_LETTERS = ["BB","BD","BF","CC","CG","CK","CM","CO","CQ","CY","DA","DS","EE","EI","EK","EM","EQ","ES","EW","FA","FM","FO","FQ","FS","FW","GH","GT","GX","HN"]
DEC_LETTERS = ["DY","GL","GN","GP","GR","GZ","HB","HH","HJ"]

# —————— UI y lógica ——————
uploaded = st.file_uploader("Sube tu archivo .xlsx", type="xlsx")
if uploaded:
    # 1) preview de origen
    df = pd.read_excel(uploaded, header=0, dtype=str)
    st.subheader("Vista previa – Datos originales")
    st.dataframe(df.head(10))

    # 2) construir DataFrame destino
    max_dest = max(col_letter_to_index(d) for _, d in MOVIMIENTOS)
    result   = pd.DataFrame(index=df.index, columns=range(max_dest + 1))

    for orig, dest in MOVIMIENTOS:
        i = col_letter_to_index(orig)
        j = col_letter_to_index(dest)
        result.iloc[:, j] = df.iloc[:, i] if i < df.shape[1] else None

    # 3) ajustar ancho al número de encabezados disponibles
    if result.shape[1] > len(header_list):
        result = result.iloc[:, : len(header_list)]
    result.columns = header_list[: result.shape[1]]

    # 4) nombres únicos
    seen = {}
    cols_unique = []
    for col in result.columns:
        if col not in seen:
            seen[col] = 0
            cols_unique.append(col)
        else:
            seen[col] += 1
            cols_unique.append(f"{col} ({seen[col]})")
    result.columns = cols_unique

    # 5) conversiones de tipo
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
        mask = result["Report Status"].notna()
        result.loc[mask, "Sample Status"] = "Completed"

    # 6) preview de destino y descarga
    st.subheader("Vista previa – Archivo reordenado")
    st.dataframe(result.head(10))

    buf = BytesIO()
    result.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    st.download_button(
        "📥 Descargar Excel reordenado",
        data=buf,
        file_name="mobilserv_reordenado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

import streamlit as st
import pandas as pd
import pulp
import math
import plotly.graph_objects as go
import io

st.set_page_config(layout="wide")
st.title("⚡ Ora Energy BESS Optimizer")

# =========================
# PARAMETRI
# =========================
st.sidebar.header("Parametri BESS")

C_max       = st.sidebar.number_input("Capacità (MWh)", value=5.0)
SoC_min     = st.sidebar.number_input("SoC min (MWh)", value=0.5)
SoC_max_val = st.sidebar.number_input("SoC max (MWh)", value=4.5)
P_ch_max    = st.sidebar.number_input("Potenza carica max (MW)", value=2.5)
P_dis_max   = st.sidebar.number_input("Potenza scarica max (MW)", value=2.5)

eta_rt      = st.sidebar.number_input("Efficienza round-trip (%)", value=90.0) / 100
eta         = math.sqrt(eta_rt)

c_deg       = st.sidebar.number_input("Costo degradazione (€/MWh)", value=0.0)

st.sidebar.header("Parametri sistema")
oneri = st.sidebar.number_input("Oneri evitati (€/MWh)", value=75.0)
c_pv  = st.sidebar.number_input("Costo energia FV (€/MWh)", value=72.0)

lim_immissione = st.sidebar.number_input("Limite immissione rete (MW)", value=7.0)
lim_prelievo   = st.sidebar.number_input("Limite prelievo rete (MW)", value=9.0)

# =========================
# INPUT FILE
# =========================
st.header("📂 Upload dati")

file_prezzi = st.file_uploader("Prezzi MGP (€/MWh)", type=["xlsx"])
file_pv     = st.file_uploader("Produzione FV (MW)", type=["xlsx"])
file_load   = st.file_uploader("Consumi stabilimento (MW)", type=["xlsx"])

# =========================
# EXPORT
# =========================
def convert_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export = df.copy()
        df_export["Data"] = df_export["Datetime"].dt.date
        df_export["Ora"]  = df_export["Datetime"].dt.hour
        df_export.to_excel(writer, index=False, sheet_name="BESS")
    return output.getvalue()

# =========================
# MODELLO
# =========================
def optimize(prices, pv, load, dates):

    T = len(prices)
    model = pulp.LpProblem("BESS_FV_MGP", pulp.LpMaximize)

    # =========================
    # VARIABILI
    # =========================
    pv_to_load = pulp.LpVariable.dicts("pv_to_load", range(T), lowBound=0)
    pv_to_bess = pulp.LpVariable.dicts("pv_to_bess", range(T), lowBound=0)
    pv_to_grid = pulp.LpVariable.dicts("pv_to_grid", range(T), lowBound=0)

    charge_grid = pulp.LpVariable.dicts("charge_grid", range(T), lowBound=0)
    discharge_grid = pulp.LpVariable.dicts("discharge_grid", range(T), lowBound=0)
    discharge_load = pulp.LpVariable.dicts("discharge_load", range(T), lowBound=0)

    soc = pulp.LpVariable.dicts(
        "soc", range(T),
        lowBound=SoC_min,
        upBound=SoC_max_val
    )

    # 🔥 BINARIA: evita charge & discharge insieme
    u = pulp.LpVariable.dicts("u", range(T), cat="Binary")

    # =========================
    # OBIETTIVO
    # =========================
    model += pulp.lpSum([
        prices[t] * pv_to_grid[t]
        + (prices[t] + oneri) * pv_to_load[t]
        + (prices[t] + oneri) * discharge_load[t]
        + prices[t] * discharge_grid[t]
        - prices[t] * charge_grid[t]
        - c_pv * pv_to_bess[t]
        - c_deg * (discharge_grid[t] + discharge_load[t])
        for t in range(T)
    ])

    # =========================
    # VINCOLI
    # =========================
    for t in range(T):

        # FV balance
        model += pv[t] == pv_to_load[t] + pv_to_bess[t] + pv_to_grid[t]

        # Load coverage
        model += pv_to_load[t] + discharge_load[t] <= load[t]

        # 🔥 MUTUA ESCLUSIONE BESS
        model += pv_to_bess[t] + charge_grid[t] <= P_ch_max * u[t]
        model += discharge_grid[t] + discharge_load[t] <= P_dis_max * (1 - u[t])

        # Grid limits
        model += pv_to_grid[t] + discharge_grid[t] <= lim_immissione
        model += charge_grid[t] <= lim_prelievo

        # SOC dynamics
        if t == 0:
            model += soc[t] == SoC_min \
                     + eta * (pv_to_bess[t] + charge_grid[t]) \
                     - (discharge_grid[t] + discharge_load[t]) / eta
        else:
            model += soc[t] == soc[t-1] \
                     + eta * (pv_to_bess[t] + charge_grid[t]) \
                     - (discharge_grid[t] + discharge_load[t]) / eta

    # reset giornaliero
    df_idx = pd.DataFrame({"Datetime": dates})
    df_idx["Data"] = df_idx["Datetime"].dt.date

    for t in df_idx.groupby("Data").head(1).index:
        model += soc[t] == SoC_min

    for t in df_idx.groupby("Data").tail(1).index:
        model += soc[t] == SoC_min

    model.solve(pulp.PULP_CBC_CMD(msg=0))

    # =========================
    # OUTPUT
    # =========================
    df = pd.DataFrame({
        "Prezzo": prices,
        "PV": pv,
        "Load": load,
        "PV_to_load": [pv_to_load[t].varValue for t in range(T)],
        "PV_to_bess": [pv_to_bess[t].varValue for t in range(T)],
        "PV_to_grid": [pv_to_grid[t].varValue for t in range(T)],
        "Charge_grid": [charge_grid[t].varValue for t in range(T)],
        "Discharge_grid": [discharge_grid[t].varValue for t in range(T)],
        "Discharge_load": [discharge_load[t].varValue for t in range(T)],
        "SoC": [soc[t].varValue for t in range(T)],
    })

    df["Charge_tot"] = df["PV_to_bess"] + df["Charge_grid"]
    df["Discharge_tot"] = df["Discharge_grid"] + df["Discharge_load"]

    df["Grid_to_load"] = (df["Load"] - df["PV_to_load"] - df["Discharge_load"]).clip(lower=0)
    df["Grid_injection"] = df["PV_to_grid"] + df["Discharge_grid"]

    df["Valore"] = (
        df["Prezzo"] * df["PV_to_grid"]
        + (df["Prezzo"] + oneri) * df["PV_to_load"]
        + (df["Prezzo"] + oneri) * df["Discharge_load"]
        + df["Prezzo"] * df["Discharge_grid"]
        - df["Prezzo"] * df["Charge_grid"]
        - c_pv * df["PV_to_bess"]
        - c_deg * df["Discharge_tot"]
    )

    return df

# =========================
# RUN
# =========================
if file_prezzi and file_pv and file_load:

    prices = pd.read_excel(file_prezzi).iloc[:, 0].dropna().tolist()
    pv = pd.read_excel(file_pv).iloc[:, 0].dropna().tolist()
    load = pd.read_excel(file_load).iloc[:, 0].dropna().tolist()

    T = min(len(prices), len(pv), len(load))
    prices, pv, load = prices[:T], pv[:T], load[:T]

    dates = pd.date_range("2025-01-01", periods=T, freq="h")

    with st.spinner("Ottimizzazione..."):
        df = optimize(prices, pv, load, dates)

    df["Datetime"] = dates
    df["Data"] = df["Datetime"].dt.date

    st.header("📊 Risultati")
    st.metric("Valore totale (€)", f"{df['Valore'].sum():,.0f}")
    st.metric("Autoconsumo FV (%)",
              f"{100*(df['PV_to_load'].sum()+df['Discharge_load'].sum())/max(df['PV'].sum(),1):.1f}%")

    fig = go.Figure()
    fig.add_trace(go.Scatter(y=df["SoC"], name="SoC"))
    st.plotly_chart(fig, use_container_width=True)

    st.download_button(
        "📥 Scarica Excel",
        convert_to_excel(df),
        "bess_results.xlsx"
    )

else:
    st.info("Carica i file per iniziare")

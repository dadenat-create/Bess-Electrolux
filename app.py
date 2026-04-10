import streamlit as st
import pandas as pd
import pulp
import math
import plotly.graph_objects as go
import io

st.set_page_config(layout="wide")
st.title("⚡ Ora Energy BESS Optimizer - Advanced")

# =========================
# PARAMETRI
# =========================
st.sidebar.header("Parametri BESS")

C_max = st.sidebar.number_input("Capacità (MWh)", value=5.0)
SoC_min = st.sidebar.number_input("SoC min", value=0.5)
SoC_max = st.sidebar.number_input("SoC max", value=4.5)

P_charge_max = st.sidebar.number_input("P carica (MW)", value=2.5)
P_discharge_max = st.sidebar.number_input("P scarica (MW)", value=2.5)

eta_rt = st.sidebar.number_input("Efficienza (%)", value=90.0)/100
eta = math.sqrt(eta_rt)

c_deg = st.sidebar.number_input("Costo degradazione", value=0.0)

st.sidebar.header("Parametri sistema")
oneri = st.sidebar.number_input("Oneri evitati", value=75.0)
c_pv = st.sidebar.number_input("Costo FV", value=72.0)

# =========================
# INPUT FILE
# =========================
file_prezzi = st.file_uploader("Prezzi MGP")
file_pv = st.file_uploader("Produzione FV")
file_load = st.file_uploader("Consumi")

# =========================
# EXPORT EXCEL
# =========================
def export_excel(df, daily, monthly, summary):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Dettaglio", index=False)
        daily.to_excel(writer, sheet_name="Giornaliero")
        monthly.to_excel(writer, sheet_name="Mensile")
        summary.to_excel(writer, sheet_name="Totale")
    return output.getvalue()

# =========================
# OTTIMIZZAZIONE
# =========================
def optimize(prices, pv, load, dates):

    T = len(prices)
    model = pulp.LpProblem("BESS", pulp.LpMaximize)

    cg = pulp.LpVariable.dicts("cg", range(T), 0)
    cpv = pulp.LpVariable.dicts("cpv", range(T), 0)
    dg = pulp.LpVariable.dicts("dg", range(T), 0)
    dl = pulp.LpVariable.dicts("dl", range(T), 0)

    pvl = pulp.LpVariable.dicts("pvl", range(T), 0)
    pvg = pulp.LpVariable.dicts("pvg", range(T), 0)
    gtl = pulp.LpVariable.dicts("gtl", range(T), 0)

    soc = pulp.LpVariable.dicts("soc", range(T), SoC_min, SoC_max)

    # OBIETTIVO (solo BESS)
    model += pulp.lpSum([
        prices[t]*dg[t]
        - prices[t]*cg[t]
        + oneri*dl[t]
        - c_pv*cpv[t]
        - c_deg*(dg[t] + dl[t])
        for t in range(T)
    ])

    for t in range(T):

        # bilanci
        model += pv[t] == pvl[t] + cpv[t] + pvg[t]
        model += pvl[t] + dl[t] + gtl[t] == load[t]

        # potenze
        model += cg[t] + cpv[t] <= P_charge_max
        model += dg[t] + dl[t] <= P_discharge_max

        # limiti rete
        model += pvg[t] + dg[t] <= 7
        model += cg[t] + gtl[t] <= 9

        # SOC
        if t == 0:
            model += soc[t] == SoC_min + eta*(cg[t]+cpv[t]) - (dg[t]+dl[t])/eta
        else:
            model += soc[t] == soc[t-1] + eta*(cg[t]+cpv[t]) - (dg[t]+dl[t])/eta

    # vincoli giornalieri
    df_idx = pd.DataFrame({"Datetime": dates})
    df_idx["Data"] = df_idx["Datetime"].dt.date

    for t in df_idx.groupby("Data").head(1).index:
        model += soc[t] == SoC_min

    for t in df_idx.groupby("Data").tail(1).index:
        model += soc[t] == SoC_min

    model.solve(pulp.PULP_CBC_CMD(msg=0))

    df = pd.DataFrame({
        "Datetime": dates,
        "Prezzo": prices,
        "Charge_grid": [cg[t].value() for t in range(T)],
        "Charge_PV": [cpv[t].value() for t in range(T)],
        "Discharge_grid": [dg[t].value() for t in range(T)],
        "Discharge_load": [dl[t].value() for t in range(T)],
        "PV_to_grid": [pvg[t].value() for t in range(T)],
        "SoC": [soc[t].value() for t in range(T)]
    })

    return df

# =========================
# RUN
# =========================
if file_prezzi and file_pv and file_load:

    prices = pd.to_numeric(pd.read_excel(file_prezzi).iloc[:,0], errors='coerce').dropna()
    pv = pd.to_numeric(pd.read_excel(file_pv).iloc[:,0], errors='coerce').dropna()
    load = pd.to_numeric(pd.read_excel(file_load).iloc[:,0], errors='coerce').dropna()

    T = min(len(prices), len(pv), len(load))

    prices = prices[:T].tolist()
    pv = pv[:T].tolist()
    load = load[:T].tolist()

    dates = pd.date_range("2025-01-01", periods=T, freq="h")

    df = optimize(prices, pv, load, dates)

    df["Data"] = df["Datetime"].dt.date
    df["Mese"] = df["Datetime"].dt.to_period("M")

    # =========================
    # BREAKDOWN
    # =========================
    total_charge = df["Charge_grid"] + df["Charge_PV"]
    ratio_grid = df["Charge_grid"] / total_charge.replace(0,1)
    ratio_pv = df["Charge_PV"] / total_charge.replace(0,1)

    df["Auto_BESS_PV"] = df["Discharge_load"] * ratio_pv
    df["Auto_BESS_Grid"] = df["Discharge_load"] * ratio_grid
    df["Sell_BESS_Grid"] = df["Discharge_grid"] * ratio_grid
    df["Sell_FV"] = df["PV_to_grid"]

    cols = ["Auto_BESS_PV","Auto_BESS_Grid","Sell_BESS_Grid","Sell_FV"]

    daily = df.groupby("Data")[cols].sum()
    monthly = df.groupby("Mese")[cols].sum()
    summary = df[cols].sum().to_frame("Totale")

    # KPI cicli
    total_discharge = df["Discharge_grid"].sum() + df["Discharge_load"].sum()
    cycles = total_discharge / C_max if C_max > 0 else 0

    st.header("📊 Dashboard")

    col1, col2, col3 = st.columns(3)
    col1.metric("🔋 Cicli", round(cycles,2))
    col2.metric("⚡ Energia scaricata (MWh)", round(total_discharge,2))
    col3.metric("📅 Giorni simulati", len(daily))

    st.subheader("Totale")
    st.dataframe(summary)

    fig = go.Figure()
    for c in cols:
        fig.add_bar(name=c, x=monthly.index.astype(str), y=monthly[c])
    st.plotly_chart(fig, use_container_width=True)

    # =========================
    # EXPORT
    # =========================
    excel = export_excel(df, daily, monthly, summary)

    st.download_button(
        "📥 Scarica Excel completo",
        data=excel,
        file_name="bess_analysis.xlsx"
    )

else:
    st.info("Carica tutti i file per partire")

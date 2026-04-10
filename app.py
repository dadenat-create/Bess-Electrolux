import streamlit as st
import pandas as pd
import pulp
import math
import plotly.graph_objects as go

st.set_page_config(layout="wide")
st.title("⚡ Ora Energy BESS Optimizer (BESS Only Value)")

# =========================
# PARAMETRI
# =========================
st.sidebar.header("Parametri BESS")

C_max = st.sidebar.number_input("Capacità (MWh)", value=5.0)
SoC_0 = st.sidebar.number_input("SoC iniziale", value=2.5)
SoC_min = st.sidebar.number_input("SoC min", value=0.5)
SoC_max = st.sidebar.number_input("SoC max", value=4.5)

P_charge_max = st.sidebar.number_input("P carica (MW)", value=2.5)
P_discharge_max = st.sidebar.number_input("P scarica (MW)", value=2.5)

eta_rt = st.sidebar.number_input("Efficienza round-trip (%)", value=90.0)/100
eta = math.sqrt(eta_rt)

c_deg = st.sidebar.number_input("Costo degradazione (€/MWh)", value=0.0)

st.sidebar.header("Parametri sistema")
c_pv = st.sidebar.number_input("Costo FV (€/MWh)", value=72.0)
oneri = st.sidebar.number_input("Oneri evitati (€/MWh)", value=75.0)

# =========================
# INPUT FILE
# =========================
st.header("Upload dati")

file_prezzi = st.file_uploader("Prezzi MGP", type=["xlsx"])
file_pv = st.file_uploader("Produzione FV", type=["xlsx"])
file_load = st.file_uploader("Consumi stabilimento", type=["xlsx"])

# =========================
# OTTIMIZZAZIONE
# =========================
def optimize(prices, pv, load):

    T = len(prices)
    model = pulp.LpProblem("BESS_only", pulp.LpMaximize)

    # Variabili
    charge_grid = pulp.LpVariable.dicts("charge_grid", range(T), lowBound=0)
    charge_pv = pulp.LpVariable.dicts("charge_pv", range(T), lowBound=0)

    discharge_grid = pulp.LpVariable.dicts("discharge_grid", range(T), lowBound=0)
    discharge_load = pulp.LpVariable.dicts("discharge_load", range(T), lowBound=0)

    pv_to_load = pulp.LpVariable.dicts("pv_to_load", range(T), lowBound=0)
    pv_to_grid = pulp.LpVariable.dicts("pv_to_grid", range(T), lowBound=0)

    grid_to_load = pulp.LpVariable.dicts("grid_to_load", range(T), lowBound=0)

    soc = pulp.LpVariable.dicts("soc", range(T), lowBound=SoC_min, upBound=SoC_max)

    # =========================
    # OBIETTIVO → SOLO BESS
    # =========================
    model += pulp.lpSum([

        # arbitraggio rete
        prices[t]*discharge_grid[t]
        - prices[t]*charge_grid[t]

        # valore autoconsumo
        + oneri*discharge_load[t]

        # costo opportunità FV
        - prices[t]*charge_pv[t]

        # degradazione
        - c_deg*(discharge_grid[t] + discharge_load[t])

        for t in range(T)
    ])

    # =========================
    # VINCOLI
    # =========================
    for t in range(T):

        # FV balance
        model += pv[t] == pv_to_load[t] + charge_pv[t] + pv_to_grid[t]

        # Load balance
        model += pv_to_load[t] + discharge_load[t] + grid_to_load[t] == load[t]

        # Limiti potenza
        model += charge_grid[t] + charge_pv[t] <= P_charge_max
        model += discharge_grid[t] + discharge_load[t] <= P_discharge_max

        # Limiti connessione
        model += pv_to_grid[t] + discharge_grid[t] <= 7
        model += charge_grid[t] + grid_to_load[t] <= 9

        # SOC
        if t == 0:
            model += soc[t] == SoC_0 + eta*(charge_grid[t]+charge_pv[t]) - (discharge_grid[t]+discharge_load[t])/eta
        else:
            model += soc[t] == soc[t-1] + eta*(charge_grid[t]+charge_pv[t]) - (discharge_grid[t]+discharge_load[t])/eta

    model.solve(pulp.PULP_CBC_CMD(msg=0))

    # =========================
    # OUTPUT
    # =========================
    df = pd.DataFrame({
        "Prezzo": prices,
        "PV": pv,
        "Load": load,
        "Charge_grid": [charge_grid[t].varValue for t in range(T)],
        "Charge_PV": [charge_pv[t].varValue for t in range(T)],
        "Discharge_grid": [discharge_grid[t].varValue for t in range(T)],
        "Discharge_load": [discharge_load[t].varValue for t in range(T)],
        "PV_to_grid": [pv_to_grid[t].varValue for t in range(T)],
        "PV_to_load": [pv_to_load[t].varValue for t in range(T)],
        "Grid_to_load": [grid_to_load[t].varValue for t in range(T)],
        "SoC": [soc[t].varValue for t in range(T)]
    })

    # PROFITTO BESS
    df["Profitto_BESS"] = (
        df["Prezzo"]*df["Discharge_grid"]
        - df["Prezzo"]*df["Charge_grid"]
        + oneri*df["Discharge_load"]
        - df["Prezzo"]*df["Charge_PV"]
        - c_deg*(df["Discharge_grid"] + df["Discharge_load"])
    )

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

    dates = pd.date_range(start="2025-01-01", periods=int(T), freq="h")

    df = optimize(prices, pv, load)
    df["Datetime"] = dates

    df["Data"] = df["Datetime"].dt.date
    df["Mese"] = df["Datetime"].dt.to_period("M")

    # KPI
    daily = df.groupby("Data")["Profitto_BESS"].sum()
    monthly = df.groupby("Mese")["Profitto_BESS"].sum()

    st.header("📊 Dashboard BESS")

    col1, col2 = st.columns(2)
    col1.metric("💰 Valore totale BESS (€)", round(df["Profitto_BESS"].sum(),2))
    col2.metric("📅 Valore medio giorno (€)", round(daily.mean(),2))

    # Mensile
    fig_m = go.Figure()
    fig_m.add_trace(go.Bar(x=monthly.index.astype(str), y=monthly.values))
    fig_m.update_layout(title="Valore mensile BESS")
    st.plotly_chart(fig_m, use_container_width=True)

    # Selezione mese
    selected_month = st.selectbox("Seleziona mese", monthly.index.astype(str))
    df_m = df[df["Mese"].astype(str)==selected_month]

    fig_month = go.Figure()
    fig_month.add_trace(go.Scatter(y=df_m["Prezzo"], name="Prezzo"))
    fig_month.add_trace(go.Bar(y=df_m["Charge_grid"], name="Charge"))
    fig_month.add_trace(go.Bar(y=df_m["Discharge_grid"], name="Discharge"))
    st.plotly_chart(fig_month, use_container_width=True)

    # Giorno
    selected_day = st.selectbox("Seleziona giorno", df_m["Data"].unique())
    df_d = df[df["Data"]==selected_day]

    st.subheader(f"Giorno {selected_day}")

    fig_day = go.Figure()
    fig_day.add_trace(go.Scatter(y=df_d["Prezzo"], name="Prezzo"))
    fig_day.add_trace(go.Bar(y=df_d["Charge_grid"], name="Charge"))
    fig_day.add_trace(go.Bar(y=df_d["Discharge_grid"], name="Discharge"))
    st.plotly_chart(fig_day, use_container_width=True)

    fig_soc = go.Figure()
    fig_soc.add_trace(go.Scatter(y=df_d["SoC"], name="SoC"))
    st.plotly_chart(fig_soc, use_container_width=True)

else:
    st.info("Carica tutti i file per iniziare")

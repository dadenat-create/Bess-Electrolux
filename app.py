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

C_max = st.sidebar.number_input("Capacità (MWh)", value=5.0)
SoC_min = st.sidebar.number_input("SoC min", value=0.5)
SoC_max = st.sidebar.number_input("SoC max", value=4.5)

P_charge_max = st.sidebar.number_input("P carica (MW)", value=2.5)
P_discharge_max = st.sidebar.number_input("P scarica (MW)", value=2.5)

eta_rt = st.sidebar.number_input("Efficienza round-trip (%)", value=90.0)/100
eta = math.sqrt(eta_rt)

c_deg = st.sidebar.number_input("Costo degradazione (€/MWh)", value=0.0)

st.sidebar.header("Parametri sistema")
oneri = st.sidebar.number_input("Oneri evitati (€/MWh)", value=75.0)
c_pv = st.sidebar.number_input("Costo energia FV (€/MWh)", value=72.0)

# =========================
# INPUT FILE
# =========================
st.header("Upload dati")

file_prezzi = st.file_uploader("Prezzi MGP", type=["xlsx"])
file_pv = st.file_uploader("Produzione FV", type=["xlsx"])
file_load = st.file_uploader("Consumi stabilimento", type=["xlsx"])

# =========================
# EXPORT EXCEL
# =========================
def convert_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export = df.copy()
        df_export["Data"] = df_export["Datetime"].dt.date
        df_export["Ora"] = df_export["Datetime"].dt.hour
        df_export.to_excel(writer, index=False, sheet_name="BESS")
    return output.getvalue()

# =========================
# OTTIMIZZAZIONE
# =========================
def optimize(prices, pv, load, dates):

    T = len(prices)
    model = pulp.LpProblem("BESS_only", pulp.LpMaximize)

    charge_grid = pulp.LpVariable.dicts("charge_grid", range(T), lowBound=0)
    charge_pv = pulp.LpVariable.dicts("charge_pv", range(T), lowBound=0)

    discharge_grid = pulp.LpVariable.dicts("discharge_grid", range(T), lowBound=0)
    discharge_load = pulp.LpVariable.dicts("discharge_load", range(T), lowBound=0)

    pv_to_load = pulp.LpVariable.dicts("pv_to_load", range(T), lowBound=0)
    pv_to_grid = pulp.LpVariable.dicts("pv_to_grid", range(T), lowBound=0)

    grid_to_load = pulp.LpVariable.dicts("grid_to_load", range(T), lowBound=0)

    soc = pulp.LpVariable.dicts("soc", range(T), lowBound=SoC_min, upBound=SoC_max)

    # =========================
    # OBIETTIVO
    # =========================
    model += pulp.lpSum([

        prices[t]*discharge_grid[t]
        - prices[t]*charge_grid[t]

        + oneri*discharge_load[t]

        - c_pv*charge_pv[t]

        - c_deg*(discharge_grid[t] + discharge_load[t])

        for t in range(T)
    ])

    # =========================
    # VINCOLI
    # =========================
    for t in range(T):

        model += pv[t] == pv_to_load[t] + charge_pv[t] + pv_to_grid[t]

        model += pv_to_load[t] + discharge_load[t] + grid_to_load[t] == load[t]

        model += charge_grid[t] + charge_pv[t] <= P_charge_max
        model += discharge_grid[t] + discharge_load[t] <= P_discharge_max

        model += pv_to_grid[t] + discharge_grid[t] <= 7
        model += charge_grid[t] + grid_to_load[t] <= 9

        if t == 0:
            model += soc[t] == SoC_min + eta*(charge_grid[t]+charge_pv[t]) - (discharge_grid[t]+discharge_load[t])/eta
        else:
            model += soc[t] == soc[t-1] + eta*(charge_grid[t]+charge_pv[t]) - (discharge_grid[t]+discharge_load[t])/eta

    # =========================
    # VINCOLI GIORNALIERI SOC
    # =========================
    df_index = pd.DataFrame({"Datetime": dates})
    df_index["Data"] = df_index["Datetime"].dt.date

    start_idx = df_index.groupby("Data").head(1).index
    end_idx = df_index.groupby("Data").tail(1).index

    for t in start_idx:
        model += soc[t] == SoC_min

    for t in end_idx:
        model += soc[t] == SoC_min

    model.solve(pulp.PULP_CBC_CMD(msg=0))

    df = pd.DataFrame({
        "Prezzo": prices,
        "PV": pv,
        "Load": load,
        "Charge_grid": [charge_grid[t].varValue for t in range(T)],
        "Charge_PV": [charge_pv[t].varValue for t in range(T)],
        "Discharge_grid": [discharge_grid[t].varValue for t in range(T)],
        "Discharge_load": [discharge_load[t].varValue for t in range(T)],
        "SoC": [soc[t].varValue for t in range(T)]
    })

    df["Profitto_BESS"] = (
        df["Prezzo"]*df["Discharge_grid"]
        - df["Prezzo"]*df["Charge_grid"]
        + oneri*df["Discharge_load"]
        - c_pv*df["Charge_PV"]
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

    df = optimize(prices, pv, load, dates)
    df["Datetime"] = dates

    df["Data"] = df["Datetime"].dt.date
    df["Mese"] = df["Datetime"].dt.to_period("M")

    # KPI
    daily = df.groupby("Data")["Profitto_BESS"].sum()
    monthly = df.groupby("Mese")["Profitto_BESS"].sum()

    total_discharge = df["Discharge_grid"].sum() + df["Discharge_load"].sum()
    cycles = total_discharge / C_max if C_max > 0 else 0

    st.header("📊 Dashboard BESS")

    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Valore totale (€)", round(df["Profitto_BESS"].sum(),2))
    col2.metric("📅 Valore medio giorno (€)", round(daily.mean(),2))
    col3.metric("🔋 Cicli equivalenti", round(cycles,2))

    # Grafico mensile
    fig_m = go.Figure()
    fig_m.add_trace(go.Bar(x=monthly.index.astype(str), y=monthly.values))
    fig_m.update_layout(title="Valore mensile BESS")
    st.plotly_chart(fig_m, use_container_width=True)

    # Giorno
    selected_day = st.selectbox("Seleziona giorno", df["Data"].unique())
    df_d = df[df["Data"]==selected_day]

    fig_day = go.Figure()
    fig_day.add_trace(go.Scatter(y=df_d["Prezzo"], name="Prezzo"))
    fig_day.add_trace(go.Bar(y=df_d["Charge_grid"], name="Charge"))
    fig_day.add_trace(go.Bar(y=df_d["Discharge_grid"], name="Discharge"))
    st.plotly_chart(fig_day, use_container_width=True)

    fig_soc = go.Figure()
    fig_soc.add_trace(go.Scatter(y=df_d["SoC"], name="SoC"))
    st.plotly_chart(fig_soc, use_container_width=True)

    # DOWNLOAD EXCEL
    excel_data = convert_to_excel(df)

    st.download_button(
        label="📥 Scarica risultati BESS (Excel)",
        data=excel_data,
        file_name="bess_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Carica tutti i file per iniziare")import streamlit as st
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

C_max = st.sidebar.number_input("Capacità (MWh)", value=5.0)
SoC_min = st.sidebar.number_input("SoC min", value=0.5)
SoC_max = st.sidebar.number_input("SoC max", value=4.5)

P_charge_max = st.sidebar.number_input("P carica (MW)", value=2.5)
P_discharge_max = st.sidebar.number_input("P scarica (MW)", value=2.5)

eta_rt = st.sidebar.number_input("Efficienza round-trip (%)", value=90.0)/100
eta = math.sqrt(eta_rt)

c_deg = st.sidebar.number_input("Costo degradazione (€/MWh)", value=0.0)

st.sidebar.header("Parametri sistema")
oneri = st.sidebar.number_input("Oneri evitati (€/MWh)", value=75.0)
c_pv = st.sidebar.number_input("Costo energia FV (€/MWh)", value=72.0)

# =========================
# INPUT FILE
# =========================
st.header("Upload dati")

file_prezzi = st.file_uploader("Prezzi MGP", type=["xlsx"])
file_pv = st.file_uploader("Produzione FV", type=["xlsx"])
file_load = st.file_uploader("Consumi stabilimento", type=["xlsx"])

# =========================
# EXPORT EXCEL
# =========================
def convert_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export = df.copy()
        df_export["Data"] = df_export["Datetime"].dt.date
        df_export["Ora"] = df_export["Datetime"].dt.hour
        df_export.to_excel(writer, index=False, sheet_name="BESS")
    return output.getvalue()

# =========================
# OTTIMIZZAZIONE
# =========================
def optimize(prices, pv, load, dates):

    T = len(prices)
    model = pulp.LpProblem("BESS_only", pulp.LpMaximize)

    charge_grid = pulp.LpVariable.dicts("charge_grid", range(T), lowBound=0)
    charge_pv = pulp.LpVariable.dicts("charge_pv", range(T), lowBound=0)

    discharge_grid = pulp.LpVariable.dicts("discharge_grid", range(T), lowBound=0)
    discharge_load = pulp.LpVariable.dicts("discharge_load", range(T), lowBound=0)

    pv_to_load = pulp.LpVariable.dicts("pv_to_load", range(T), lowBound=0)
    pv_to_grid = pulp.LpVariable.dicts("pv_to_grid", range(T), lowBound=0)

    grid_to_load = pulp.LpVariable.dicts("grid_to_load", range(T), lowBound=0)

    soc = pulp.LpVariable.dicts("soc", range(T), lowBound=SoC_min, upBound=SoC_max)

    # =========================
    # OBIETTIVO
    # =========================
    model += pulp.lpSum([

        prices[t]*discharge_grid[t]
        - prices[t]*charge_grid[t]

        + oneri*discharge_load[t]

        - c_pv*charge_pv[t]

        - c_deg*(discharge_grid[t] + discharge_load[t])

        for t in range(T)
    ])

    # =========================
    # VINCOLI
    # =========================
    for t in range(T):

        model += pv[t] == pv_to_load[t] + charge_pv[t] + pv_to_grid[t]

        model += pv_to_load[t] + discharge_load[t] + grid_to_load[t] == load[t]

        model += charge_grid[t] + charge_pv[t] <= P_charge_max
        model += discharge_grid[t] + discharge_load[t] <= P_discharge_max

        model += pv_to_grid[t] + discharge_grid[t] <= 7
        model += charge_grid[t] + grid_to_load[t] <= 9

        if t == 0:
            model += soc[t] == SoC_min + eta*(charge_grid[t]+charge_pv[t]) - (discharge_grid[t]+discharge_load[t])/eta
        else:
            model += soc[t] == soc[t-1] + eta*(charge_grid[t]+charge_pv[t]) - (discharge_grid[t]+discharge_load[t])/eta

    # =========================
    # VINCOLI GIORNALIERI SOC
    # =========================
    df_index = pd.DataFrame({"Datetime": dates})
    df_index["Data"] = df_index["Datetime"].dt.date

    start_idx = df_index.groupby("Data").head(1).index
    end_idx = df_index.groupby("Data").tail(1).index

    for t in start_idx:
        model += soc[t] == SoC_min

    for t in end_idx:
        model += soc[t] == SoC_min

    model.solve(pulp.PULP_CBC_CMD(msg=0))

    df = pd.DataFrame({
        "Prezzo": prices,
        "PV": pv,
        "Load": load,
        "Charge_grid": [charge_grid[t].varValue for t in range(T)],
        "Charge_PV": [charge_pv[t].varValue for t in range(T)],
        "Discharge_grid": [discharge_grid[t].varValue for t in range(T)],
        "Discharge_load": [discharge_load[t].varValue for t in range(T)],
        "SoC": [soc[t].varValue for t in range(T)]
    })

    df["Profitto_BESS"] = (
        df["Prezzo"]*df["Discharge_grid"]
        - df["Prezzo"]*df["Charge_grid"]
        + oneri*df["Discharge_load"]
        - c_pv*df["Charge_PV"]
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

    df = optimize(prices, pv, load, dates)
    df["Datetime"] = dates

    df["Data"] = df["Datetime"].dt.date
    df["Mese"] = df["Datetime"].dt.to_period("M")

    # KPI
    daily = df.groupby("Data")["Profitto_BESS"].sum()
    monthly = df.groupby("Mese")["Profitto_BESS"].sum()

    total_discharge = df["Discharge_grid"].sum() + df["Discharge_load"].sum()
    cycles = total_discharge / C_max if C_max > 0 else 0

    st.header("📊 Dashboard BESS")

    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Valore totale (€)", round(df["Profitto_BESS"].sum(),2))
    col2.metric("📅 Valore medio giorno (€)", round(daily.mean(),2))
    col3.metric("🔋 Cicli equivalenti", round(cycles,2))

    # Grafico mensile
    fig_m = go.Figure()
    fig_m.add_trace(go.Bar(x=monthly.index.astype(str), y=monthly.values))
    fig_m.update_layout(title="Valore mensile BESS")
    st.plotly_chart(fig_m, use_container_width=True)

    # Giorno
    selected_day = st.selectbox("Seleziona giorno", df["Data"].unique())
    df_d = df[df["Data"]==selected_day]

    fig_day = go.Figure()
    fig_day.add_trace(go.Scatter(y=df_d["Prezzo"], name="Prezzo"))
    fig_day.add_trace(go.Bar(y=df_d["Charge_grid"], name="Charge"))
    fig_day.add_trace(go.Bar(y=df_d["Discharge_grid"], name="Discharge"))
    st.plotly_chart(fig_day, use_container_width=True)

    fig_soc = go.Figure()
    fig_soc.add_trace(go.Scatter(y=df_d["SoC"], name="SoC"))
    st.plotly_chart(fig_soc, use_container_width=True)

    # DOWNLOAD EXCEL
    excel_data = convert_to_excel(df)

    st.download_button(
        label="📥 Scarica risultati BESS (Excel)",
        data=excel_data,
        file_name="bess_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Carica tutti i file per iniziare")

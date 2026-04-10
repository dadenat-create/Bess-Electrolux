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

    model += pulp.lpSum([
        prices[t]*discharge_grid[t]
        - prices[t]*charge_grid[t]
        + oneri*discharge_load[t]
        - c_pv*charge_pv[t]
        - c_deg*(discharge_grid[t] + discharge_load[t])
        for t in range(T)
    ])

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
        "PV_to_load": [pv_to_load[t].varValue for t in range(T)],
        "PV_to_grid": [pv_to_grid[t].varValue for t in range(T)],
        "Grid_to_load": [grid_to_load[t].varValue for t in range(T)],
        "SoC": [soc[t].varValue for t in range(T)]
    })

    df["Profitto_BESS"] = (
        df["Prezzo"]*df["Discharge_grid"]
        - df["Prezzo"]*df["Charge_grid"]
        + oneri*df["Discharge_load"]
        - c_pv*df["Charge_PV"]
        - c_deg*(df["Discharge_grid"] + df["Discharge_load"])
    )

    # -------------------------------------------------------
    # TRACCIAMENTO ORIGINE ENERGIA NELLA BESS
    # Per ogni ora calcoliamo la quota di Charge_PV e Charge_grid
    # sul totale caricato, e la applichiamo proporzionalmente alle
    # due modalità di scarica (discharge_load e discharge_grid).
    # -------------------------------------------------------
    total_charge = df["Charge_PV"] + df["Charge_grid"]

    # Quota FV nella carica totale (0 se nessuna carica)
    df["frac_pv_in_charge"] = df["Charge_PV"] / total_charge.where(total_charge > 0, other=1)
    df["frac_grid_in_charge"] = df["Charge_grid"] / total_charge.where(total_charge > 0, other=1)

    # Scarica verso carico: quota originata da FV vs Rete
    df["Discharge_load_from_PV"]  = df["Discharge_load"] * df["frac_pv_in_charge"]
    df["Discharge_load_from_grid"] = df["Discharge_load"] * df["frac_grid_in_charge"]

    # Scarica verso rete: quota originata da FV vs Rete
    df["Discharge_grid_from_PV"]  = df["Discharge_grid"] * df["frac_pv_in_charge"]
    df["Discharge_grid_from_grid"] = df["Discharge_grid"] * df["frac_grid_in_charge"]

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

    # KPI principali
    daily = df.groupby("Data")["Profitto_BESS"].sum()
    monthly = df.groupby("Mese")["Profitto_BESS"].sum()

    total_discharge = df["Discharge_grid"].sum() + df["Discharge_load"].sum()
    cycles = total_discharge / C_max if C_max > 0 else 0

    st.header("📊 Dashboard BESS")

    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Valore totale (€)", round(df["Profitto_BESS"].sum(), 2))
    col2.metric("📅 Valore medio giorno (€)", round(daily.mean(), 2))
    col3.metric("🔋 Cicli equivalenti", round(cycles, 2))

    # Grafico mensile
    fig_m = go.Figure()
    fig_m.add_trace(go.Bar(x=monthly.index.astype(str), y=monthly.values))
    fig_m.update_layout(title="Valore mensile BESS")
    st.plotly_chart(fig_m, use_container_width=True)

    # =========================================================
    # SEZIONE NUOVA: ANALISI FLUSSI ENERGETICI
    # =========================================================
    st.header("📈 Analisi Flussi Energetici")

    # --- Aggregati totali ---
    tot_pv                   = df["PV"].sum()
    tot_pv_to_load           = df["PV_to_load"].sum()          # autoconsumo diretto FV
    tot_pv_to_bess           = df["Charge_PV"].sum()           # FV → BESS (carica)
    tot_pv_to_grid_direct    = df["PV_to_grid"].sum()          # FV → rete diretta (senza BESS)
    tot_bess_load_from_pv    = df["Discharge_load_from_PV"].sum()   # FV via BESS → carico
    tot_bess_grid_from_pv    = df["Discharge_grid_from_PV"].sum()   # FV via BESS → rete
    tot_bess_load_from_grid  = df["Discharge_load_from_grid"].sum() # Rete via BESS → carico
    tot_bess_grid_from_grid  = df["Discharge_grid_from_grid"].sum() # Rete via BESS → rete (arbitraggio)
    tot_grid_to_load         = df["Grid_to_load"].sum()        # Rete → carico diretto
    tot_charge_grid          = df["Charge_grid"].sum()         # Rete → BESS (carica)
    tot_prelievo_rete        = tot_grid_to_load + tot_charge_grid   # Prelievo rete totale

    st.subheader("Totali periodo (MWh)")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("☀️ FV prodotto", f"{tot_pv:.1f}")
    col2.metric("🏭 Carico totale", f"{df['Load'].sum():.1f}")
    col3.metric("🔌 Prelievo da rete", f"{tot_prelievo_rete:.1f}")
    col4.metric("📤 Immesso in rete", f"{(tot_pv_to_grid_direct + tot_bess_grid_from_pv + tot_bess_grid_from_grid):.1f}")

    st.subheader("Flussi FV (MWh)")
    col1, col2, col3 = st.columns(3)
    col1.metric(
        "🟢 Autoconsumo diretto FV",
        f"{tot_pv_to_load:.1f}",
        help="PV → Carico direttamente, senza passare dalla BESS"
    )
    col2.metric(
        "🔋 Autoconsumo FV via BESS",
        f"{tot_bess_load_from_pv:.1f}",
        help="FV caricato in BESS e poi scaricato verso il carico"
    )
    col3.metric(
        "📤 Venduto in rete FV via BESS",
        f"{tot_bess_grid_from_pv:.1f}",
        help="FV caricato in BESS e poi venduto in rete (time-shifting)"
    )

    st.subheader("Flussi Rete (MWh)")
    col1, col2, col3 = st.columns(3)
    col1.metric(
        "⚡ Rete → Carico diretto",
        f"{tot_grid_to_load:.1f}",
        help="Energia prelevata dalla rete e consumata direttamente"
    )
    col2.metric(
        "🔋 Rete via BESS → Carico",
        f"{tot_bess_load_from_grid:.1f}",
        help="Energia caricata in BESS dalla rete e poi scaricata verso il carico"
    )
    col3.metric(
        "💹 Arbitraggio puro (Rete → BESS → Rete)",
        f"{tot_bess_grid_from_grid:.1f}",
        help="Energia acquistata dalla rete, stoccata in BESS e rivenduta in rete"
    )

    # --- Grafico a torta: destinazione energia FV ---
    st.subheader("Destinazione energia FV")
    fig_pv_pie = go.Figure(go.Pie(
        labels=[
            "Autoconsumo diretto",
            "Autoconsumo via BESS",
            "Venduto via BESS (time-shift)",
            "Venduto in rete diretto"
        ],
        values=[
            tot_pv_to_load,
            tot_bess_load_from_pv,
            tot_bess_grid_from_pv,
            tot_pv_to_grid_direct
        ],
        hole=0.4,
        marker_colors=["#2ecc71", "#27ae60", "#f39c12", "#3498db"]
    ))
    fig_pv_pie.update_layout(title="Come viene utilizzata l'energia FV")
    st.plotly_chart(fig_pv_pie, use_container_width=True)

    # --- Grafico a torta: origine energia prelevata dalla rete ---
    st.subheader("Utilizzo energia prelevata dalla rete")
    fig_grid_pie = go.Figure(go.Pie(
        labels=[
            "Rete → Carico diretto",
            "Rete via BESS → Carico",
            "Rete via BESS → Rete (arbitraggio)"
        ],
        values=[
            tot_grid_to_load,
            tot_bess_load_from_grid,
            tot_bess_grid_from_grid
        ],
        hole=0.4,
        marker_colors=["#e74c3c", "#c0392b", "#8e44ad"]
    ))
    fig_grid_pie.update_layout(title="Come viene utilizzata l'energia prelevata dalla rete")
    st.plotly_chart(fig_grid_pie, use_container_width=True)

    # --- Grafico mensile dei flussi principali ---
    st.subheader("Andamento mensile flussi energetici")
    monthly_flows = df.groupby("Mese").agg(
        PV_diretto=("PV_to_load", "sum"),
        PV_via_BESS_carico=("Discharge_load_from_PV", "sum"),
        PV_via_BESS_rete=("Discharge_grid_from_PV", "sum"),
        PV_rete_diretta=("PV_to_grid", "sum"),
        Rete_carico_diretto=("Grid_to_load", "sum"),
        Rete_BESS_carico=("Discharge_load_from_grid", "sum"),
        Arbitraggio=("Discharge_grid_from_grid", "sum"),
    ).reset_index()

    mesi = monthly_flows["Mese"].astype(str)

    fig_flows = go.Figure()
    fig_flows.add_trace(go.Bar(name="FV diretto → Carico",        x=mesi, y=monthly_flows["PV_diretto"],          marker_color="#2ecc71"))
    fig_flows.add_trace(go.Bar(name="FV via BESS → Carico",       x=mesi, y=monthly_flows["PV_via_BESS_carico"],  marker_color="#27ae60"))
    fig_flows.add_trace(go.Bar(name="FV via BESS → Rete",         x=mesi, y=monthly_flows["PV_via_BESS_rete"],    marker_color="#f39c12"))
    fig_flows.add_trace(go.Bar(name="FV diretto → Rete",          x=mesi, y=monthly_flows["PV_rete_diretta"],     marker_color="#3498db"))
    fig_flows.add_trace(go.Bar(name="Rete → Carico diretto",      x=mesi, y=monthly_flows["Rete_carico_diretto"], marker_color="#e74c3c"))
    fig_flows.add_trace(go.Bar(name="Rete via BESS → Carico",     x=mesi, y=monthly_flows["Rete_BESS_carico"],    marker_color="#c0392b"))
    fig_flows.add_trace(go.Bar(name="Arbitraggio Rete→BESS→Rete", x=mesi, y=monthly_flows["Arbitraggio"],         marker_color="#8e44ad"))

    fig_flows.update_layout(
        barmode="stack",
        title="Flussi energetici mensili (MWh)",
        yaxis_title="MWh",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig_flows, use_container_width=True)

    # =========================================================
    # GRAFICI GIORNALIERI (originali + flussi)
    # =========================================================
    st.header("📅 Dettaglio giornaliero")

    selected_day = st.selectbox("Seleziona giorno", df["Data"].unique())
    df_d = df[df["Data"] == selected_day]

    fig_day = go.Figure()
    fig_day.add_trace(go.Scatter(y=df_d["Prezzo"], name="Prezzo (€/MWh)", yaxis="y2", line=dict(dash="dot")))
    fig_day.add_trace(go.Bar(y=df_d["Charge_grid"] + df_d["Charge_PV"], name="Carica totale BESS"))
    fig_day.add_trace(go.Bar(y=df_d["Discharge_grid"] + df_d["Discharge_load"], name="Scarica totale BESS"))
    fig_day.update_layout(
        title="Cariche/Scariche BESS e prezzo",
        yaxis=dict(title="MW"),
        yaxis2=dict(title="€/MWh", overlaying="y", side="right"),
        barmode="group"
    )
    st.plotly_chart(fig_day, use_container_width=True)

    # Grafico flussi giornalieri dettagliato
    fig_day_flows = go.Figure()
    fig_day_flows.add_trace(go.Bar(y=df_d["PV_to_load"],               name="FV → Carico diretto",     marker_color="#2ecc71"))
    fig_day_flows.add_trace(go.Bar(y=df_d["Discharge_load_from_PV"],   name="FV via BESS → Carico",    marker_color="#27ae60"))
    fig_day_flows.add_trace(go.Bar(y=df_d["Discharge_grid_from_PV"],   name="FV via BESS → Rete",      marker_color="#f39c12"))
    fig_day_flows.add_trace(go.Bar(y=df_d["PV_to_grid"],               name="FV → Rete diretto",       marker_color="#3498db"))
    fig_day_flows.add_trace(go.Bar(y=df_d["Grid_to_load"],             name="Rete → Carico diretto",   marker_color="#e74c3c"))
    fig_day_flows.add_trace(go.Bar(y=df_d["Discharge_load_from_grid"], name="Rete via BESS → Carico",  marker_color="#c0392b"))
    fig_day_flows.add_trace(go.Bar(y=df_d["Discharge_grid_from_grid"], name="Arbitraggio Rete→BESS→Rete", marker_color="#8e44ad"))
    fig_day_flows.update_layout(
        title="Flussi energetici orari",
        yaxis_title="MW",
        barmode="stack"
    )
    st.plotly_chart(fig_day_flows, use_container_width=True)

    fig_soc = go.Figure()
    fig_soc.add_trace(go.Scatter(y=df_d["SoC"], name="SoC (MWh)", fill="tozeroy"))
    fig_soc.update_layout(title="State of Charge", yaxis_title="MWh")
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

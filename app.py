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
# EXPORT EXCEL
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
# OTTIMIZZAZIONE
# =========================
def optimize(prices, pv, load, dates):
    """
    Variabili decisionali (per ogni ora t):
      Flussi FV:
        pv_to_load[t]  : FV → carico diretto
        pv_to_bess[t]  : FV → BESS  (= charge_pv)
        pv_to_grid[t]  : FV → rete diretta

      BESS:
        charge_grid[t]    : Rete → BESS
        charge_pv[t]      : FV → BESS  (alias pv_to_bess)
        discharge_grid[t] : BESS → rete
        discharge_load[t] : BESS → carico

      Stato:
        soc[t]            : State of Charge (MWh)

    Funzione obiettivo (dalla specifica §7):
      max Σ_t [
          P[t]            * pv_to_grid[t]           (vendita FV diretto)
        + (P[t]+oneri)    * pv_to_load[t]            (autoconsumo FV diretto)
        + (P[t]+oneri)    * discharge_load[t]         (autoconsumo BESS)
        + P[t]            * discharge_grid[t]         (vendita BESS in rete)
        - P[t]            * charge_grid[t]            (acquisto energia da rete)
        - c_pv            * charge_pv[t]              (costo energia FV caricata)
        - c_deg           * (discharge_grid+discharge_load)[t]  (degradazione)
      ]

    Vincoli principali:
      - Bilancio FV:  pv[t] = pv_to_load + pv_to_bess + pv_to_grid
      - Bilancio carico: pv_to_load + discharge_load = load[t]
          (il carico residuo è soddisfatto dalla rete: grid_to_load = load - pv_to_load - discharge_load)
      - Carica totale: charge_grid + charge_pv ≤ P_ch_max
      - Scarica totale: discharge_grid + discharge_load ≤ P_dis_max
      - Immissione: pv_to_grid + discharge_grid ≤ lim_immissione
      - Prelievo: charge_grid ≤ lim_prelievo
      - SoC dynamics con reset giornaliero a SoC_min
    """
    T = len(prices)
    model = pulp.LpProblem("BESS_FV_MGP", pulp.LpMaximize)

    # --- Variabili ---
    pv_to_load     = pulp.LpVariable.dicts("pv_to_load",     range(T), lowBound=0)
    pv_to_bess     = pulp.LpVariable.dicts("pv_to_bess",     range(T), lowBound=0)
    pv_to_grid     = pulp.LpVariable.dicts("pv_to_grid",     range(T), lowBound=0)

    charge_grid    = pulp.LpVariable.dicts("charge_grid",    range(T), lowBound=0)
    discharge_grid = pulp.LpVariable.dicts("discharge_grid", range(T), lowBound=0)
    discharge_load = pulp.LpVariable.dicts("discharge_load", range(T), lowBound=0)

    soc            = pulp.LpVariable.dicts("soc",            range(T),
                                           lowBound=SoC_min, upBound=SoC_max_val)

    # --- Funzione obiettivo (specifica §7) ---
    model += pulp.lpSum([
          prices[t]           * pv_to_grid[t]
        + (prices[t] + oneri) * pv_to_load[t]
        + (prices[t] + oneri) * discharge_load[t]
        + prices[t]           * discharge_grid[t]
        - prices[t]           * charge_grid[t]
        - c_pv                * pv_to_bess[t]
        - c_deg               * (discharge_grid[t] + discharge_load[t])
        for t in range(T)
    ])

    # --- Vincoli ---
    for t in range(T):

        # Bilancio FV (§6.1)
        model += pv[t] == pv_to_load[t] + pv_to_bess[t] + pv_to_grid[t]

        # Bilancio carico:
        # il carico viene coperto da FV diretto + BESS scarica
        # la parte restante è prelevata dalla rete (variabile implicita, sempre ≥ 0)
        model += pv_to_load[t] + discharge_load[t] <= load[t]

        # Carica totale BESS (§6.2)
        model += charge_grid[t] + pv_to_bess[t] <= P_ch_max

        # Scarica totale BESS (§6.3)
        model += discharge_grid[t] + discharge_load[t] <= P_dis_max

        # Vincolo immissione lato produzione (§6.5)
        model += pv_to_grid[t] + discharge_grid[t] <= lim_immissione

        # Vincolo prelievo rete (§6.6)
        model += charge_grid[t] <= lim_prelievo

        # Dinamica SoC (§4.2)
        if t == 0:
            model += soc[t] == SoC_min \
                     + eta * (charge_grid[t] + pv_to_bess[t]) \
                     - (discharge_grid[t] + discharge_load[t]) / eta
        else:
            model += soc[t] == soc[t-1] \
                     + eta * (charge_grid[t] + pv_to_bess[t]) \
                     - (discharge_grid[t] + discharge_load[t]) / eta

    # Vincoli di reset SoC giornaliero
    df_idx = pd.DataFrame({"Datetime": dates})
    df_idx["Data"] = df_idx["Datetime"].dt.date
    for t in df_idx.groupby("Data").head(1).index:
        model += soc[t] == SoC_min
    for t in df_idx.groupby("Data").tail(1).index:
        model += soc[t] == SoC_min

    model.solve(pulp.PULP_CBC_CMD(msg=0))

    # --- Costruzione DataFrame risultati ---
    df = pd.DataFrame({
        "Prezzo":          prices,
        "PV":              pv,
        "Load":            load,
        "PV_to_load":      [pv_to_load[t].varValue      for t in range(T)],
        "PV_to_BESS":      [pv_to_bess[t].varValue      for t in range(T)],
        "PV_to_grid":      [pv_to_grid[t].varValue      for t in range(T)],
        "Charge_grid":     [charge_grid[t].varValue      for t in range(T)],
        "Discharge_grid":  [discharge_grid[t].varValue   for t in range(T)],
        "Discharge_load":  [discharge_load[t].varValue   for t in range(T)],
        "SoC":             [soc[t].varValue              for t in range(T)],
    })

    # Carica/scarica totale aggregata
    df["Charge_tot"]    = df["Charge_grid"] + df["PV_to_BESS"]
    df["Discharge_tot"] = df["Discharge_grid"] + df["Discharge_load"]

    # Prelievo rete implicito (residuo del carico non coperto da FV/BESS)
    df["Grid_to_load"] = (df["Load"] - df["PV_to_load"] - df["Discharge_load"]).clip(lower=0)

    # Immissione totale in rete
    df["Grid_injection"] = df["PV_to_grid"] + df["Discharge_grid"]

    # Valore economico ora per ora (funzione obiettivo §7)
    df["Valore"] = (
          df["Prezzo"]           * df["PV_to_grid"]
        + (df["Prezzo"] + oneri) * df["PV_to_load"]
        + (df["Prezzo"] + oneri) * df["Discharge_load"]
        + df["Prezzo"]           * df["Discharge_grid"]
        - df["Prezzo"]           * df["Charge_grid"]
        - c_pv                   * df["PV_to_BESS"]
        - c_deg                  * df["Discharge_tot"]
    )

    # --- Tracciamento origine energia nella BESS ---
    # Quota proporzionale FV/rete nella carica → applicata alle scariche
    total_ch = df["Charge_tot"]
    df["frac_pv"]   = (df["PV_to_BESS"]  / total_ch.where(total_ch > 0, other=1)).clip(0, 1)
    df["frac_grid"] = (df["Charge_grid"] / total_ch.where(total_ch > 0, other=1)).clip(0, 1)

    df["Dis_load_from_PV"]   = df["Discharge_load"] * df["frac_pv"]
    df["Dis_load_from_grid"] = df["Discharge_load"] * df["frac_grid"]
    df["Dis_grid_from_PV"]   = df["Discharge_grid"] * df["frac_pv"]
    df["Dis_grid_from_grid"] = df["Discharge_grid"] * df["frac_grid"]

    return df

# =========================
# RUN
# =========================
if file_prezzi and file_pv and file_load:

    prices = pd.to_numeric(pd.read_excel(file_prezzi).iloc[:, 0], errors='coerce').dropna()
    pv     = pd.to_numeric(pd.read_excel(file_pv).iloc[:, 0],     errors='coerce').dropna()
    load   = pd.to_numeric(pd.read_excel(file_load).iloc[:, 0],   errors='coerce').dropna()

    T = min(len(prices), len(pv), len(load))
    prices = prices[:T].tolist()
    pv     = pv[:T].tolist()
    load   = load[:T].tolist()

    dates = pd.date_range(start="2025-01-01", periods=int(T), freq="h")

    with st.spinner("Ottimizzazione in corso..."):
        df = optimize(prices, pv, load, dates)

    df["Datetime"] = dates
    df["Data"]     = df["Datetime"].dt.date
    df["Mese"]     = df["Datetime"].dt.to_period("M")

    # =========================================================
    # KPI PRINCIPALI
    # =========================================================
    st.header("📊 Dashboard BESS")

    daily   = df.groupby("Data")["Valore"].sum()
    monthly = df.groupby("Mese")["Valore"].sum()
    cycles  = df["Discharge_tot"].sum() / C_max if C_max > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Valore totale (€)",       f"{df['Valore'].sum():,.0f}")
    col2.metric("📅 Valore medio giorno (€)",  f"{daily.mean():,.0f}")
    col3.metric("🔋 Cicli equivalenti",         f"{cycles:.1f}")
    col4.metric("☀️ Autoconsumo FV (%)",
                f"{100*(df['PV_to_load'].sum()+df['Dis_load_from_PV'].sum()) / max(df['PV'].sum(),1):.1f}%")

    # Grafico valore mensile
    fig_m = go.Figure()
    fig_m.add_trace(go.Bar(x=monthly.index.astype(str), y=monthly.values, marker_color="#2563eb"))
    fig_m.update_layout(title="Valore mensile (€)", yaxis_title="€")
    st.plotly_chart(fig_m, use_container_width=True)

    # =========================================================
    # ANALISI FLUSSI ENERGETICI
    # =========================================================
    st.header("📈 Analisi Flussi Energetici")

    tot_pv               = df["PV"].sum()
    tot_pv_to_load       = df["PV_to_load"].sum()
    tot_pv_to_bess       = df["PV_to_BESS"].sum()
    tot_pv_to_grid       = df["PV_to_grid"].sum()
    tot_dis_load_pv      = df["Dis_load_from_PV"].sum()
    tot_dis_grid_pv      = df["Dis_grid_from_PV"].sum()
    tot_dis_load_grid    = df["Dis_load_from_grid"].sum()
    tot_dis_grid_grid    = df["Dis_grid_from_grid"].sum()
    tot_grid_to_load     = df["Grid_to_load"].sum()
    tot_charge_grid      = df["Charge_grid"].sum()

    st.subheader("Totali periodo (MWh)")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("☀️ FV prodotto",          f"{tot_pv:.0f}")
    col2.metric("🏭 Carico totale",         f"{df['Load'].sum():.0f}")
    col3.metric("🔌 Prelievo da rete",      f"{tot_grid_to_load + tot_charge_grid:.0f}")
    col4.metric("📤 Immissione in rete",    f"{df['Grid_injection'].sum():.0f}")

    st.subheader("Flussi FV (MWh)")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("🟢 FV → Carico diretto",       f"{tot_pv_to_load:.0f}",
                help="pv_to_load: FV autoconsumato senza passare per la BESS")
    col2.metric("🔋 FV via BESS → Carico",       f"{tot_dis_load_pv:.0f}",
                help="FV caricato in BESS, poi scaricato sul carico (autoconsumo ritardato)")
    col3.metric("📤 FV via BESS → Rete",         f"{tot_dis_grid_pv:.0f}",
                help="FV c

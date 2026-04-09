import streamlit as st
import pandas as pd
import pulp
import math

st.set_page_config(layout="wide")
st.title("⚡ Ora Energy BESS + FV Optimizer")

# =========================
# INPUT PARAMETRI
# =========================
st.sidebar.header("Parametri BESS")

C_max = st.sidebar.number_input("Capacità (MWh)", value=5.0)
SoC_0 = st.sidebar.number_input("SoC iniziale", value=1.0)
SoC_min = st.sidebar.number_input("SoC min", value=0.5)
SoC_max = st.sidebar.number_input("SoC max", value=4.5)

P_charge_max = st.sidebar.number_input("P carica (MW)", value=2.5)
P_discharge_max = st.sidebar.number_input("P scarica (MW)", value=2.5)

eta_rt = st.sidebar.number_input("Efficienza (%)", value=90.0)/100
eta = math.sqrt(eta_rt)

c_deg = st.sidebar.number_input("Costo degradazione", value=0.0)

st.sidebar.header("Parametri economici")
c_pv = st.sidebar.number_input("Costo FV (€/MWh)", value=72.0)
oneri = st.sidebar.number_input("Oneri evitati (€/MWh)", value=75.0)

# =========================
# FILE INPUT
# =========================
st.header("Upload dati")

file_prezzi = st.file_uploader("Prezzi MGP", type=["xlsx"])
file_pv = st.file_uploader("Produzione FV", type=["xlsx"])
file_load = st.file_uploader("Consumi stabilimento", type=["xlsx"])

if file_prezzi and file_pv and file_load:

    prices = pd.to_numeric(pd.read_excel(file_prezzi).iloc[:,0], errors='coerce').dropna().tolist()
    pv = pd.to_numeric(pd.read_excel(file_pv).iloc[:,0], errors='coerce').dropna().tolist()
    load = pd.to_numeric(pd.read_excel(file_load).iloc[:,0], errors='coerce').dropna().tolist()

    T = min(len(prices), len(pv), len(load))

    prices = prices[:T]
    pv = pv[:T]
    load = load[:T]

    if st.button("🚀 Ottimizza"):

        model = pulp.LpProblem("BESS_FV", pulp.LpMaximize)

        # Variabili
        charge_grid = pulp.LpVariable.dicts("charge_grid", range(T), lowBound=0)
        charge_pv = pulp.LpVariable.dicts("charge_pv", range(T), lowBound=0)

        discharge_grid = pulp.LpVariable.dicts("discharge_grid", range(T), lowBound=0)
        discharge_load = pulp.LpVariable.dicts("discharge_load", range(T), lowBound=0)

        pv_to_load = pulp.LpVariable.dicts("pv_to_load", range(T), lowBound=0)
        pv_to_grid = pulp.LpVariable.dicts("pv_to_grid", range(T), lowBound=0)

        soc = pulp.LpVariable.dicts("soc", range(T), lowBound=SoC_min, upBound=SoC_max)

        # =========================
        # OBIETTIVO
        # =========================
        model += pulp.lpSum([
            # vendita FV
            prices[t] * pv_to_grid[t]

            # autoconsumo FV
            + (prices[t] + oneri) * pv_to_load[t]

            # scarica BESS
            + prices[t] * discharge_grid[t]
            + (prices[t] + oneri) * discharge_load[t]

            # costi
            - prices[t] * charge_grid[t]
            - c_pv * charge_pv[t]
            - c_deg * (discharge_grid[t] + discharge_load[t])

            for t in range(T)
        ])

        # =========================
        # VINCOLI
        # =========================
        for t in range(T):

            # bilancio FV
            model += pv[t] == pv_to_load[t] + charge_pv[t] + pv_to_grid[t]

            # soddisfacimento carico
            model += pv_to_load[t] + discharge_load[t] <= load[t]

            # limiti potenza
            model += charge_grid[t] + charge_pv[t] <= P_charge_max
            model += discharge_grid[t] + discharge_load[t] <= P_discharge_max

            # vincolo immissione
            model += pv_to_grid[t] + discharge_grid[t] <= 7

            # vincolo prelievo
            model += charge_grid[t] <= 9

            # SOC
            if t == 0:
                model += soc[t] == SoC_0 + eta*(charge_grid[t]+charge_pv[t]) - (discharge_grid[t]+discharge_load[t])/eta
            else:
                model += soc[t] == soc[t-1] + eta*(charge_grid[t]+charge_pv[t]) - (discharge_grid[t]+discharge_load[t])/eta

        model += soc[T-1] == SoC_0

        model.solve(pulp.PULP_CBC_CMD(msg=0))

        # =========================
        # RISULTATI
        # =========================
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

        df["Profitto"] = (
            df["Prezzo"]*df["Discharge_grid"]
            + (df["Prezzo"]+oneri)*df["Discharge_load"]
            + df["Prezzo"]*(pv_to_grid[0].varValue if T>0 else 0)
        )

        st.success("Ottimizzazione completata")

        st.dataframe(df)

        st.download_button(
            "Scarica risultati",
            df.to_csv(index=False),
            "results.csv"
        )

else:
    st.info("Carica tutti i file per iniziare")

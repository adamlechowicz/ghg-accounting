import streamlit as st
import pandas as pd
import numpy as np
import io
import plotly.express as px
import xlsxwriter
import requests

# --- NEW: STATE TO EGRID MAPPING ---
# A simplified mapping of US State -> Primary eGRID Subregion.
# Note: Real grid boundaries cross state lines. This is a heuristic.
STATE_TO_EGRID = {
    "AL": "SRMV (SERC Mississippi Valley)", "AR": "SRMV (SERC Mississippi Valley)",
    "AZ": "AZNM (WECC Southwest)", "CA": "CAMX (California)",
    "CO": "RMPA (WECC Rockies)", "CT": "NEWE (New England)",
    "DC": "RFCE (RFC East)", "DE": "RFCE (RFC East)",
    "FL": "FRCC (FRCC All)", "GA": "SRMV (SERC Mississippi Valley)", # Split, but SRMV/SRSO cover parts
    "IA": "MROW (MRO West)", "ID": "NWPP (WECC Northwest)",
    "IL": "RFCW (RFC West)", "IN": "RFCW (RFC West)",
    "KS": "SPNO (SPP North)", "KY": "SRMV (SERC Mississippi Valley)",
    "LA": "SRMV (SERC Mississippi Valley)", "MA": "NEWE (New England)",
    "MD": "RFCE (RFC East)", "ME": "NEWE (New England)",
    "MI": "RFCM (RFC Michigan)", "MN": "MROE (MRO East)",
    "MO": "SPNO (SPP North)", "MS": "SRMV (SERC Mississippi Valley)",
    "MT": "NWPP (WECC Northwest)", "NC": "SRMV (SERC Mississippi Valley)",
    "ND": "MROW (MRO West)", "NE": "MROW (MRO West)",
    "NH": "NEWE (New England)", "NJ": "RFCE (RFC East)",
    "NM": "AZNM (WECC Southwest)", "NV": "NWPP (WECC Northwest)",
    "NY": "NYUP (Upstate NY)", # Defaulting NY to Upstate (covers more area)
    "OH": "RFCW (RFC West)", "OK": "SPSO (SPP South)",
    "OR": "NWPP (WECC Northwest)", "PA": "RFCE (RFC East)",
    "RI": "NEWE (New England)", "SC": "SRMV (SERC Mississippi Valley)",
    "SD": "MROW (MRO West)", "TN": "SRMV (SERC Mississippi Valley)",
    "TX": "ERCT (ERCOT All)", "UT": "NWPP (WECC Northwest)",
    "VA": "SRMV (SERC Mississippi Valley)", "VT": "NEWE (New England)",
    "WA": "NWPP (WECC Northwest)", "WI": "MROE (MRO East)",
    "WV": "RFCW (RFC West)", "WY": "RMPA (WECC Rockies)"
}

def guess_region_from_ip():
    """
    Attempts to guess the eGRID region based on the user's IP address.
    Returns the region key (str) or None if failed.
    """
    try:
        # 1. Get IP from New Streamlit Context
        # st.context.headers returns a dictionary-like object of headers
        headers = st.context.headers
        if not headers:
            return None
        
        # 'X-Forwarded-For' usually contains: "Client_IP, Proxy1, Proxy2"
        # We want the first one. Keys are often case-insensitive but usually lower-case in dicts.
        x_forwarded = headers.get("X-Forwarded-For")
        
        if x_forwarded:
            client_ip = x_forwarded.split(",")[0].strip()
        else:
            return None 

        # 2. Geolocation (Using free API)
        # Using ip-api.com (free for non-commercial use)
        response = requests.get(f"http://ip-api.com/json/{client_ip}", timeout=2)
        data = response.json()
        
        if data.get('status') == 'success' and data.get('countryCode') == 'US':
            state_code = data.get('region') # e.g., "MA"
            return STATE_TO_EGRID.get(state_code)
            
    except Exception as e:
        # Fail silently
        return None
    return None

# --- CONFIGURATION & DATA ---

# 1. Amtrak Electrified Stations (NEC Spine + Keystone)
NEC_STATIONS = {
    "BOS", "BBY", "RTE", "PVD", "KIN", "WLY", "MYS", "NLC", "OSB", 
    "NHV", "BRP", "STM", "NRO", "NYP", "NWK", "EWR", "MET", "NBK", 
    "PJC", "TRE", "CWH", "PHN", "PHL", "WIL", "NRK", "ABE", "BAL", 
    "BWI", "NCR", "WAS", "ARD", "PAO", "EXT", "DTW", "CTZ", "PAR", 
    "LNC", "MDT", "HAR"
}

# 2. EPA eGRID 2023 Data (lb/MWh -> converted later)
EGRID_DATA = {
    "NEWE (New England)": {"co2": 539.275, "ch4": 0.063, "n2o": 0.008},
    "NYUP (Upstate NY)": {"co2": 242.089, "ch4": 0.011, "n2o": 0.001},
    "NYCW (NYC/Westchester)": {"co2": 864.469, "ch4": 0.022, "n2o": 0.002},
    "NYLI (Long Island)": {"co2": 1180.672, "ch4": 0.140, "n2o": 0.018},
    "CAMX (California)": {"co2": 428.464, "ch4": 0.025, "n2o": 0.003},
    "AZNM (WECC Southwest)": {"co2": 703.703, "ch4": 0.039, "n2o": 0.005},
    "ERCT (ERCOT All)": {"co2": 733.862, "ch4": 0.043, "n2o": 0.006},
    "FRCC (FRCC All)": {"co2": 782.262, "ch4": 0.041, "n2o": 0.005},
    "MROE (MRO East)": {"co2": 1397.313, "ch4": 0.116, "n2o": 0.017},
    "MROW (MRO West)": {"co2": 920.130, "ch4": 0.097, "n2o": 0.014},
    "NWPP (WECC Northwest)": {"co2": 631.735, "ch4": 0.054, "n2o": 0.008},
    "RFCE (RFC East)": {"co2": 596.904, "ch4": 0.036, "n2o": 0.005},
    "RFCM (RFC Michigan)": {"co2": 970.617, "ch4": 0.082, "n2o": 0.012},
    "RFCW (RFC West)": {"co2": 911.424, "ch4": 0.071, "n2o": 0.010},
    "RMPA (WECC Rockies)": {"co2": 1036.601, "ch4": 0.090, "n2o": 0.013},
    "SPNO (SPP North)": {"co2": 861.999, "ch4": 0.087, "n2o": 0.012},
    "SPSO (SPP South)": {"co2": 872.042, "ch4": 0.054, "n2o": 0.008},
    "SRMV (SERC Mississippi Valley)": {"co2": 739.720, "ch4": 0.032, "n2o": 0.004},
    "US Average": {"co2": 830.0, "ch4": 0.060, "n2o": 0.008}
}

# 3. Base Emission Factors
FACTORS = {
    "gasoline": { "co2_kg_per_gal": 8.78, "ch4_g_per_gal": 0.38, "n2o_g_per_gal": 0.08 },
    "diesel": { "co2_kg_per_gal": 10.21, "ch4_g_per_gal": 0.41, "n2o_g_per_gal": 0.08 },
    "fuel_oil": { "co2_kg_per_gal": 10.21, "ch4_g_per_gal": 0.41, "n2o_g_per_gal": 0.08 },
    "natural_gas": { "co2_kg_per_therm": 5.3, "ch4_g_per_therm": 0.1, "n2o_g_per_therm": 0.01 },
    "propane": { "co2_kg_per_gal": 5.74, "ch4_g_per_gal": 0.12, "n2o_g_per_gal": 0.02 },
    "air_short": { "co2_kg_per_mi": 0.207, "ch4_g_per_mi": 0.0064, "n2o_g_per_mi": 0.0066 },
    "air_medium": { "co2_kg_per_mi": 0.12926, "ch4_g_per_mi": 0.00064, "n2o_g_per_mi": 0.00410 },
    "air_long": { "co2_kg_per_mi": 0.16256, "ch4_g_per_mi": 0.00064, "n2o_g_per_mi": 0.00518 },
    "rail_ne": { "co2_kg_per_mi": 0.058, "ch4_g_per_mi": 0.0055, "n2o_g_per_mi": 0.0007 },
    "rail_other": { "co2_kg_per_mi": 0.15, "ch4_g_per_mi": 0.0117, "n2o_g_per_mi": 0.0038 },
    "gwp": { "co2": 1, "ch4": 28, "n2o": 265 }
}

# --- HELPER FUNCTIONS ---

@st.cache_data
def load_data_sources():
    # 1. Airports
    airports = {}
    try:
        df_air = pd.read_csv("https://davidmegginson.github.io/ourairports-data/airports.csv")
        df_air = df_air.dropna(subset=['iata_code'])
        airports = df_air.set_index('iata_code')[['latitude_deg', 'longitude_deg']].T.to_dict('list')
    except: pass
    
    # 2. Amtrak Stations
    amtrak = {}
    try:
        url_amtrak = "https://raw.githubusercontent.com/csunlab/PAWU/main/Amtrak_Stations_2020.csv"
        df_rail = pd.read_csv(url_amtrak)
        if 'STNCODE' in df_rail.columns:
            amtrak = df_rail.set_index('STNCODE')[['Latitude', 'Longitude']].T.to_dict('list')
    except: pass
    return airports, amtrak

def get_electricity_factor(region_key):
    data = EGRID_DATA[region_key]
    lb_to_kg = 0.453592
    return {
        "co2_kg_per_mwh": data["co2"] * lb_to_kg,
        "ch4_g_per_mwh": data["ch4"] * lb_to_kg * 1000,
        "n2o_g_per_mwh": data["n2o"] * lb_to_kg * 1000
    }

def haversine_distance(lat1, lon1, lat2, lon2):
    R_miles = 3958.8
    phi1, lambda1, phi2, lambda2 = map(np.radians, [lat1, lon1, lat2, lon2])
    a = np.sin((phi2-phi1)/2)**2 + np.cos(phi1) * np.cos(phi2) * np.sin((lambda2-lambda1)/2)**2
    c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1-a)) 
    return R_miles * c

def calculate_co2e(amount, factor_dict, unit_conversion=1):
    k_co2 = [k for k in factor_dict.keys() if 'co2' in k][0]
    k_ch4 = [k for k in factor_dict.keys() if 'ch4' in k][0]
    k_n2o = [k for k in factor_dict.keys() if 'n2o' in k][0]
    
    co2 = amount * unit_conversion * factor_dict[k_co2]
    ch4_g = amount * unit_conversion * factor_dict[k_ch4]
    n2o_g = amount * unit_conversion * factor_dict[k_n2o]
    
    co2e_kg = (co2 * FACTORS["gwp"]["co2"]) + \
              (ch4_g / 1000 * FACTORS["gwp"]["ch4"]) + \
              (n2o_g / 1000 * FACTORS["gwp"]["n2o"])
    return co2e_kg

# --- STREAMLIT APP ---

st.set_page_config(page_title="Personal Carbon Calculator", layout="wide")
st.title("🌱 Yearly Personal GHG Estimate Calculator")
st.markdown("""
This calculator uses [emissions factors published by the U.S. EPA](https://www.epa.gov/climateleadership/ghg-emission-factors-hub)
to estimate personal greenhouse gas (GHG) emissions (expressed as metric tons of CO2 equivalent) based on vehicle use, home energy consumption, and travel by air and rail.
Estimates do not include emissions from food, goods, waste, and other activities not explicitly captured here. 
This calculator's data is specific to the United States. 
""")

#Everything is open source!

if 'flights' not in st.session_state: st.session_state.flights = []
if 'rail_trips' not in st.session_state: st.session_state.rail_trips = []

airport_db, amtrak_db = load_data_sources()

# --- SIDEBAR INPUTS ---

st.sidebar.header("1. Vehicle (Scope 1)")
vehicle_type = st.sidebar.selectbox("Vehicle Type", ["Gasoline", "Diesel", "Electric (EV)"], help="If your vehicle is a hybrid or plug-in hybrid, select the non-electric fuel used and estimate the fuel consumption accordingly.")


gallons_gas = 0
gallons_diesel = 0
if vehicle_type == "Gasoline":
    car_miles = st.sidebar.number_input("Miles Driven", value=1000)
    car_mpg = st.sidebar.number_input("Average MPG (miles per gallon)", value=30.0, help="If you have a hybrid vehicle, use the combined MPG rating.")
    if car_mpg > 0:
        gallons_gas = car_miles / car_mpg
elif vehicle_type == "Diesel":
    car_miles = st.sidebar.number_input("Miles Driven", value=1000)
    car_mpg = st.sidebar.number_input("Average MPG (miles per gallon)", value=30.0, help="If you have a hybrid vehicle, use the combined MPG rating.")
    if car_mpg > 0:
        gallons_diesel = car_miles / car_mpg
else:
    st.sidebar.info("EV emissions are captured in your Electricity usage (Scope 2).")
    car_miles = st.sidebar.number_input("Miles Driven (for record)", value=1000, help="This assumes you mostly charge your EV at home.")
    gallons_gas = 0

st.sidebar.header("2. Home Heating (Scope 1)")
heat_source = st.sidebar.selectbox("Heating Source Fuel", ["Electric", "Fuel Oil", "Natural Gas", "Propane"], help="Select your primary home heating fuel.")
oil_gallons = 0; nat_gas_therms = 0; propane_gallons = 0
if heat_source == "Fuel Oil": oil_gallons = st.sidebar.number_input("Gallons of Oil", value=0)
elif heat_source == "Natural Gas": nat_gas_therms = st.sidebar.number_input("Therms of Gas", value=0)
elif heat_source == "Propane": propane_gallons = st.sidebar.number_input("Gallons of Propane", value=0)
else:
    st.sidebar.info("Heating emissions are captured in your Electricity usage (Scope 2).")

st.sidebar.header("3. Electricity (Scope 2)")
egrid_help_url = "https://www.epa.gov/system/files/images/2024-05/egrid-subregion-map.png"

# Initialize session state for the region index
if "egrid_index" not in st.session_state:
    # Default to 0 (NEWE)
    st.session_state.egrid_index = 0
    
    # Try to guess
    guessed_region = guess_region_from_ip()
    if guessed_region:
        # Find the index of the guessed region in our data keys
        all_regions = list(EGRID_DATA.keys())
        if guessed_region in all_regions:
            st.session_state.egrid_index = all_regions.index(guessed_region)
            st.toast(f"📍 Detected location: {guessed_region.split(' ')[0]}. Updated grid region!")

# Render the selectbox using the session state index
egrid_region = st.sidebar.selectbox(
    "eGRID Subregion", 
    list(EGRID_DATA.keys()), 
    index=st.session_state.egrid_index, # Uses the auto-detected index
    help=f"What subregion am I in? [Map of subregions]({egrid_help_url})"
)
elec_factor = get_electricity_factor(egrid_region)

elec_str = st.sidebar.text_area("Electricity Used (kWh)", value="200, 200, 200, 150, 150, 100, 150, 150, 150, 100, 200, 200", help=f"Input a single value for your total annual kWh, or a comma-separated list of monthly kWh values.")
elec_values = [float(x.strip()) for x in elec_str.split(',') if x.strip()]
try: elec_kwh = sum(elec_values) 
except: elec_kwh = 0
st.sidebar.caption(f"**Total Parsed:** {elec_kwh:,.0f} kWh ({len(elec_values)} entries)")

# --- LOGGERS (Updated with Editable Tables) ---

# st.sidebar.markdown("---")
st.sidebar.header("4. Air Travel Log (Scope 3)")
with st.sidebar.form("flight_form"):
    c1, c2 = st.columns(2)
    f_orig = c1.text_input("Origin", max_chars=3).upper()
    f_dest = c2.text_input("Dest", max_chars=3).upper()
    if st.form_submit_button("Add Flight ✈️"):
        if f_orig in airport_db and f_dest in airport_db:
            dist = haversine_distance(*airport_db[f_orig], *airport_db[f_dest])
            cat = "Short Haul" if dist < 300 else ("Medium Haul" if dist < 2300 else "Long Haul")
            st.session_state.flights.append({"Origin": f_orig, "Dest": f_dest, "Miles": dist, "Category": cat})
        else: st.error("Invalid Airport Code")

# Display Editable Flight Log
if st.session_state.flights:
    st.sidebar.caption("✎ Edit mileage or delete rows below:")
    edited_flights = st.sidebar.data_editor(
        pd.DataFrame(st.session_state.flights),
        column_config={
            "Origin": st.column_config.TextColumn(disabled=True),
            "Dest": st.column_config.TextColumn(disabled=True),
            "Miles": st.column_config.NumberColumn(min_value=0, format="%.0f"),
            "Category": st.column_config.SelectboxColumn(
                options=["Short Haul", "Medium Haul", "Long Haul"],
                required=True
            )
        },
        num_rows="dynamic", # Allows users to delete rows
        hide_index=True,
        key="flight_editor" # Unique key for state management
    )
    # Sync edits back to session state immediately
    st.session_state.flights = edited_flights.to_dict('records')


st.sidebar.header("5. Amtrak Travel Log (Scope 3)")
with st.sidebar.form("rail_form"):
    r1, r2 = st.columns(2)
    r_orig = r1.text_input("Origin", max_chars=3).upper()
    r_dest = r2.text_input("Dest", max_chars=3).upper()
    if st.form_submit_button("Add Rail Trip 🚆"):
        if r_orig in amtrak_db and r_dest in amtrak_db:
            dist = haversine_distance(*amtrak_db[r_orig], *amtrak_db[r_dest]) * 1.25
            is_nec = (r_orig in NEC_STATIONS) and (r_dest in NEC_STATIONS)
            cat = "Northeast Corridor" if is_nec else "Other Routes"
            st.session_state.rail_trips.append({"Origin": r_orig, "Dest": r_dest, "Miles": dist, "Category": cat})
        else: st.error("Invalid Station Code")

# Display Editable Rail Log
if st.session_state.rail_trips:
    st.sidebar.caption("✎ Edit mileage or delete rows below:")
    edited_rail = st.sidebar.data_editor(
        pd.DataFrame(st.session_state.rail_trips),
        column_config={
            "Origin": st.column_config.TextColumn(disabled=True),
            "Dest": st.column_config.TextColumn(disabled=True),
            "Miles": st.column_config.NumberColumn(min_value=0, format="%.0f"),
            "Category": st.column_config.SelectboxColumn(
                options=["Northeast Corridor", "Other Routes"],
                required=True
            )
        },
        num_rows="dynamic",
        hide_index=True,
        key="rail_editor"
    )
    # Sync edits back to session state immediately
    st.session_state.rail_trips = edited_rail.to_dict('records')

# Global Clear Button
if st.session_state.flights or st.session_state.rail_trips:
    st.sidebar.markdown("---")
    if st.sidebar.button("Clear All Logs"):
        st.session_state.flights = []
        st.session_state.rail_trips = []
        st.rerun()

# --- CALCULATIONS ---

air_s = sum(f['Miles'] for f in st.session_state.flights if f['Category'] == "Short Haul")
air_m = sum(f['Miles'] for f in st.session_state.flights if f['Category'] == "Medium Haul")
air_l = sum(f['Miles'] for f in st.session_state.flights if f['Category'] == "Long Haul")

rail_ne = sum(r['Miles'] for r in st.session_state.rail_trips if r['Category'] == "Northeast Corridor")
rail_other = sum(r['Miles'] for r in st.session_state.rail_trips if r['Category'] == "Other Routes")

results = []
# Scope 1
results.append({"Category": "Vehicle", "Scope": "Scope 1", "CO2e (kg)": calculate_co2e(gallons_gas, FACTORS["gasoline"]) + calculate_co2e(gallons_diesel, FACTORS["diesel"])})
if heat_source == "Fuel Oil": results.append({"Category": "Heating (Oil)", "Scope": "Scope 1", "CO2e (kg)": calculate_co2e(oil_gallons, FACTORS["fuel_oil"])})
elif heat_source == "Natural Gas": results.append({"Category": "Heating (Gas)", "Scope": "Scope 1", "CO2e (kg)": calculate_co2e(nat_gas_therms, FACTORS["natural_gas"])})
elif heat_source == "Propane": results.append({"Category": "Heating (Propane)", "Scope": "Scope 1", "CO2e (kg)": calculate_co2e(propane_gallons, FACTORS["propane"])})

# Scope 2
results.append({"Category": "Electricity", "Scope": "Scope 2", "CO2e (kg)": calculate_co2e(elec_kwh, elec_factor, 0.001)})

# Scope 3
results.append({"Category": "Air (short-haul)", "Scope": "Scope 3", "CO2e (kg)": calculate_co2e(air_s, FACTORS["air_short"])})
results.append({"Category": "Air (medium-haul)", "Scope": "Scope 3", "CO2e (kg)": calculate_co2e(air_m, FACTORS["air_medium"])})
results.append({"Category": "Air (long-haul)", "Scope": "Scope 3", "CO2e (kg)": calculate_co2e(air_l, FACTORS["air_long"])})
results.append({"Category": "Rail (NE Corridor)", "Scope": "Scope 3", "CO2e (kg)": calculate_co2e(rail_ne, FACTORS["rail_ne"])})
results.append({"Category": "Rail (other)", "Scope": "Scope 3", "CO2e (kg)": calculate_co2e(rail_other, FACTORS["rail_other"])})

df_res = pd.DataFrame(results)
df_res["CO2e (MT)"] = df_res["CO2e (kg)"] / 1000
total_mt = df_res["CO2e (MT)"].sum()

# --- DISPLAY ---

col1, col2 = st.columns([2, 1])
with col1:
    st.subheader("Emissions Breakdown")
    fig = px.pie(df_res, values='CO2e (MT)', names='Category', hole=0.4, color_discrete_sequence=px.colors.qualitative.Prism)
    # hide 0% slices
    fig.update_traces(textinfo='percent+label', textposition='inside', insidetextfont_size=14,
                      hovertemplate='%{label}: %{value:.2f} MT CO2e (%{percent})',
                      texttemplate='%{percent:.1%}<br>%{label}',
                      marker=dict(line=dict(color='#000000', width=1)))
    st.plotly_chart(fig, width='stretch')

    st.info(f"eGRID subregion: **{egrid_region}**", icon="⚡")
    if vehicle_type == "Electric (EV)":
        st.success("Vehicle: **EV** (Emissions in Scope 2)", icon="🚗")
        # make this info box green


with col2:
    st.subheader("Summary")
    st.metric("Total Emissions", f"{total_mt:.2f} MT CO2e")
    st.markdown("### By Scope")
    
    # FIXED LINE BELOW: Specify the column name in the format dictionary
    st.dataframe(
        df_res.groupby("Scope")["CO2e (MT)"].sum().reset_index()
        .style.format({"CO2e (MT)": "{:.2f}"}), 
        hide_index=True
    )

    st.markdown('### By Category')
    st.dataframe(
        df_res[["Category", "CO2e (MT)"]]
        .style.format({"CO2e (MT)": "{:.2f}"}), 
        hide_index=True
    )


st.subheader("Export")
# --- EXCEL EXPORT (Revised for All Features) ---

def generate_excel():
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Formats
    fmt_bold = wb.add_format({'bold': True})
    fmt_head = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    fmt_num = wb.add_format({'num_format': '0.00'})
    
    # --- SHEET 1: CALCULATOR ---
    ws = wb.add_worksheet("Calculator")
    
    # 1. EMISSION FACTORS
    ws.write("A1", "EMISSION FACTORS", fmt_bold)
    ws.write_row(1, 0, ["Factor ID", "CO2 Factor", "Unit", "CH4 Factor", "N2O Factor"], fmt_head)
    
    row = 2
    factor_map = {}
    
    # Static Factors
    for key, val in FACTORS.items():
        if key == "gwp": continue
        k_co2 = [k for k in val.keys() if 'co2' in k][0]
        # Write Row
        ws.write(row, 0, key)
        ws.write(row, 1, val[k_co2], fmt_num)
        ws.write(row, 2, k_co2.split("_")[-1])
        ws.write(row, 3, val[[k for k in val.keys() if 'ch4' in k][0]], fmt_num)
        ws.write(row, 4, val[[k for k in val.keys() if 'n2o' in k][0]], fmt_num)
        # Store Excel Cell Refs
        factor_map[key] = {'co2': f"B{row+1}", 'ch4': f"D{row+1}", 'n2o': f"E{row+1}"}
        row += 1
        
    # Dynamic Electricity Factor
    ws.write(row, 0, f"electricity ({egrid_region})", fmt_bold)
    ws.write(row, 1, elec_factor["co2_kg_per_mwh"], fmt_num)
    ws.write(row, 2, "MWh")
    ws.write(row, 3, elec_factor["ch4_g_per_mwh"], fmt_num)
    ws.write(row, 4, elec_factor["n2o_g_per_mwh"], fmt_num)
    factor_map["electricity"] = {'co2': f"B{row+1}", 'ch4': f"D{row+1}", 'n2o': f"E{row+1}"}
    row += 1
    
    # GWP Constants
    row += 1
    ws.write(row, 0, "GWP Constants", fmt_bold)
    gwp_start = row + 1
    ws.write(gwp_start, 0, "CO2"); ws.write(gwp_start, 1, FACTORS['gwp']['co2'])
    ws.write(gwp_start+1, 0, "CH4"); ws.write(gwp_start+1, 1, FACTORS['gwp']['ch4'])
    ws.write(gwp_start+2, 0, "N2O"); ws.write(gwp_start+2, 1, FACTORS['gwp']['n2o'])
    
    # Refs for formulas
    ref_gwp_ch4 = f"B{gwp_start+2}"
    ref_gwp_n2o = f"B{gwp_start+3}"

    # 2. CALCULATIONS
    calc_start = row + 5
    ws.write(calc_start, 0, "CALCULATIONS", fmt_bold)
    headers = ["Category", "Input Amount", "Unit", "CO2 (kg)", "CH4 (g)", "N2O (g)", "Total CO2e (MT)"]
    ws.write_row(calc_start+1, 0, headers, fmt_head)
    
    curr = calc_start + 2
    
    # Helper to write calculation rows
    def write_calc(name, amt, unit, f_key, mult=1, comment=None):
        nonlocal curr
        f = factor_map[f_key]
        
        ws.write(curr, 0, name)
        ws.write(curr, 1, amt, fmt_num)
        if comment:
            ws.write_comment(curr, 1, comment)
            
        ws.write(curr, 2, unit)
        
        # Formulas
        ws.write_formula(curr, 3, f"=B{curr+1}*{mult}*{f['co2']}", fmt_num)
        ws.write_formula(curr, 4, f"=B{curr+1}*{mult}*{f['ch4']}", fmt_num)
        ws.write_formula(curr, 5, f"=B{curr+1}*{mult}*{f['n2o']}", fmt_num)
        
        # Total CO2e (MT)
        f_mt = f"=(D{curr+1} + (E{curr+1}/1000*{ref_gwp_ch4}) + (F{curr+1}/1000*{ref_gwp_n2o})) / 1000"
        ws.write_formula(curr, 6, f_mt, fmt_num)
        curr += 1

    # Write Data
    if vehicle_type == "Gasoline":
        write_calc("Vehicle (Gas)", gallons_gas, "gallons", "gasoline")
    else:
        write_calc("Vehicle (EV)", 0, "gallons", "gasoline")
        
    if heat_source == "Fuel Oil":
        write_calc("Heating (Oil)", oil_gallons, "gallons", "fuel_oil")
    elif heat_source == "Natural Gas":
        write_calc("Heating (Gas)", nat_gas_therms, "therms", "natural_gas")
        
    write_calc("Electricity", elec_kwh, "kWh", "electricity", mult=0.001, comment=f"Raw Inputs: {elec_str}")
    
    write_calc("Air Short", air_s, "miles", "air_short")
    write_calc("Air Med", air_m, "miles", "air_medium")
    write_calc("Air Long", air_l, "miles", "air_long")
    write_calc("Rail NE", rail_ne, "miles", "rail_ne")
    write_calc("Rail Other", rail_other, "miles", "rail_other")
    
    # Grand Total
    ws.write(curr, 5, "TOTAL (MT):", fmt_bold)
    ws.write_formula(curr, 6, f"=SUM(G{calc_start+3}:G{curr})", wb.add_format({'bold': True, 'num_format': '0.00'}))

    # --- SHEET 2: TRAVEL LOGS ---
    if st.session_state.flights or st.session_state.rail_trips:
        ws2 = wb.add_worksheet("Travel Logs")
        ws2.write_row(0, 0, ["Type", "Origin", "Destination", "Miles", "Category"], fmt_head)
        
        r2 = 1
        for f in st.session_state.flights:
            ws2.write_row(r2, 0, ["Air", f['Origin'], f['Dest'], f['Miles'], f['Category']], fmt_num)
            r2 += 1
            
        for t in st.session_state.rail_trips:
            ws2.write_row(r2, 0, ["Rail", t['Origin'], t['Dest'], t['Miles'], t['Category']], fmt_num)
            r2 += 1
            
    wb.close()
    return output.getvalue()

st.download_button("Download Report (.xlsx)", generate_excel(), "carbon_footprint.xlsx")

st.subheader("About This Calculator")
st.markdown("""
I created this GHG inventory calculator mostly to speed up my own personal yearly estimates -- see my [blog post](https://adamlechowicz.github.io/blog) for more details.
This calculator is fully open source: code is available on the associated [GitHub repository](https://github.com/adamlechowicz/ghg-accounting), and licensed under the [MIT License](https://raw.githubusercontent.com/adamlechowicz/ghg-accounting/refs/heads/main/LICENSE).
""")

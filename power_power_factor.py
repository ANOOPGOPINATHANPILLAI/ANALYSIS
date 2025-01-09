import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import numpy as np

def read_and_process_data(file_path):
    df = pd.read_excel(file_path)
    df['Power_DC'] = df['Vdc'] * df['Idc'] / 1000
    df['Power_AC'] = ((df['Euv'] * df['Iu'] + df['Evw'] * df['Iv'] + df['Ewu'] * df['Iw']) / 3) * np.sqrt(3) / 1000
    df['Efficiency'] = (df['Power_DC'] / df['Power_AC']) * 100
    
    if isinstance(df['TM'].iloc[0], str):
        df['TM'] = pd.to_datetime(df['TM'], format='%H:%M:%S').dt.time

    return df

def process_wind_speed_data(df):
    df = df[df['WSD'] >= 3]
    #df['WSD_rounded'] = np.floor(df['WSD'])
    df['WSD_rounded'] = df['WSD'].round(1)
    wind_speed_groups = df.groupby('WSD_rounded').agg({
        'Power_DC': 'mean',
        'Power_AC': 'mean'
    }).reset_index()
    return wind_speed_groups

def create_power_vs_wind_speed_plot(wind_speed_groups):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=wind_speed_groups['WSD_rounded'], 
        y=wind_speed_groups['Power_DC'], 
        name='DC Power Output', 
        line=dict(color='blue'), 
        mode='lines+markers'
    ))
    fig.add_trace(go.Scatter(
        x=wind_speed_groups['WSD_rounded'], 
        y=wind_speed_groups['Power_AC'], 
        name='AC Power Input', 
        line=dict(color='red'), 
        mode='lines+markers'
    ))
    fig.update_layout(
        title='Power vs Wind Speed (Averaged)',
        xaxis_title='Wind Speed (m/s)',
        yaxis_title='Power (kW)',
        plot_bgcolor='white'
    )
    return fig

def error_analysis(df):
    error_codes = {
        # Error codes mapping as provided
        1: "Generator overcurrent",
        2: "Three-phase voltage imbalance",
        3: "DC voltage output to the inverter is too high", 
        4: "DC current output to the inverter is too high", 
        5: "Overcurrent during motor operation", 
        6: "Motor circuit anomaly", 
        7: "Disk brake anomaly", 
        8: "Generator stator overheating",
        9: "Dynamic unbalance in the wind turbine", 
        10: "Wind controller overheating", 
        11: "Warning from the internal ECU module of the wind controller", 
        12: "Brake command from the internal ECU module of the wind controller", 
        13: "Generator speed too high", 
        25: "Timer switch off", 
        26: "Oil system anomaly", 
        27: "Internal communication anomaly", 
        28: "Pneumatic brake failure", 
        29: "High wind speed", 
        30: "Generator over-speed (possible inverter issue)", 
        33: "Exceeded pushing limit within time", 
        34: "Disk brake needs replacement", 
        39: "Memory reset or battery failure", 
        40: "Add gearbox lubrication oil", 
        41: "No gearbox oil left", 
        44: "Generator mechanical fault", 
        45: "Brake caliper line pressure increase", 
        47: "Oil pressure motor overload protection",
        0: "System in standby", 
        14: "Wind speed reaching turbine start-up speed", 
        15: "Waiting for turbine speed to reach boost circuit start-up speed", 
        16: "System generating power normally", 
        17: "No-load short brake", 
        18: "No-load long brake", 
        19: "No-load brake", 
        20: "No-load brake with generator output three-phase short circuit", 
        21: "Loaded short brake", 
        22: "Loaded long brake", 
        23: "Loaded brake", 
        24: "Loaded brake with generator output three-phase short circuit", 
        31: "Waiting for manual reset", 
        32: "Automatic reset",
        42: "Manual stop, local forced stop", 
        43: "Remote stop", 
        46: "Cooling fan activated"

    }
    
    df['Error Name'] = df['CODE 1'].map(error_codes)
    unique_errors = df[['CODE 1', 'Error Name']].drop_duplicates()
    return unique_errors

def calculate_power_factors(df):
    df = df[(df['WSD'] >= 3) & (df['Power_DC'] > 0) & (df['Power_AC'] > 0)]
    df['WSD_interval'] = np.floor(df['WSD'])
    df['PWo_PWi'] = df['Power_DC'] / df['Power_AC']
    df = df[df['PWo_PWi'] < 1]  # Exclude power factors >= 1

    power_factors = df.groupby('WSD_interval').agg({
        'PWo_PWi': 'mean'
    }).reset_index()
    return power_factors
    


def main():
    st.title("Wind Power Analysis")
    
    uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")
    
    if uploaded_file is not None:
        try:
            df = read_and_process_data(uploaded_file)
            wind_speed_groups = process_wind_speed_data(df)
            fig = create_power_vs_wind_speed_plot(wind_speed_groups)
            st.plotly_chart(fig)
            
            error_data = error_analysis(df)
            st.write("Error Code and Name:")
            st.write(error_data)

            power_factors = calculate_power_factors(df)
            st.write("Mean Power Factor by Wind Speed:")
            st.table(power_factors)
            
        except Exception as e:
            st.error(f"An error occurred while processing the data: {str(e)}")

if __name__ == "__main__":
    main()

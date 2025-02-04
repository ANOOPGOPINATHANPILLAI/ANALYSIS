import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import streamlit as st
from datetime import datetime, time
from collections import defaultdict

def calculate_power_coefficient(power_ac, wind_speed, air_density, rotor_diameter,rotor_height):
    """Calculate power coefficient (Cp)"""
    # Calculate swept area
    area = rotor_height * rotor_diameter
    
    # Avoid division by zero and very low wind speeds
    wind_speed = np.where(wind_speed < 0.1, 0.1, wind_speed)
    
    # Calculate power coefficient
    denominator = 0.5 * air_density * area * wind_speed**3
    cp = np.where(denominator > 0, power_ac * 1000 / denominator, 0)  # Convert power from kW to W
    
    # Cap Cp at theoretical Betz limit (0.593) and remove negative values
    cp = np.clip(cp, 0, 0.593)
    #cp = np.clip(cp, 0, 1)
    
    return cp

def read_and_process_data(file_path, air_density, rotor_diameter, rotor_height):
    """Read and process the Excel data"""
    df = pd.read_excel(file_path)
    
    # Calculate powers
    df['Power_DC'] = df['Vdc'] * df['Idc'] / 1000
    df['Power_AC'] = ((df['Euv'] * df['Iu'] + df['Evw'] * df['Iv'] + df['Ewu'] * df['Iw']) / 3) * np.sqrt(3) / 1000
    
    #maximum value of power
    max_power_dc = df['Power_DC'].max()
    st.write(f"Maximum Power: {max_power_dc:.2f} kW")
    max_power_ac = df['Power_AC'].max()
    st.write(f"Maximum Power AC: {max_power_ac:.2f} kW")
    # Calculate efficiency
    df['Efficiency'] = (df['Power_DC'] / df['Power_AC']) * 100
    
    # Calculate power coefficient
    df['Cp'] = calculate_power_coefficient(df['Power_AC'], df['WSD'], air_density, rotor_diameter, rotor_height)
    
    # Convert time strings to time objects if they aren't already
    if isinstance(df['TM'].iloc[0], str):
        df['TM'] = pd.to_datetime(df['TM'], format='%H:%M:%S').dt.time


    return df

def process_wind_speed_data(df):
    """Process data by wind speed intervals"""
    df['WSD_rounded'] = round(df['WSD'] * 10) / 10
    wind_speed_groups = df.groupby('WSD_rounded').agg({
        'Power_DC': 'mean',
        'Power_AC': 'mean',
        'Cp': 'mean'
    }).reset_index()
    return wind_speed_groups
def process_time_data(df):
    try:
        # Create a copy to avoid modifying the original dataframe
        df_copy = df.copy()
        
        # Check if 'TM' column contains time objects or datetime objects
        if isinstance(df_copy['TM'].iloc[0], time):
            # Combine with a dummy date
            df_copy['minute'] = df_copy['TM'].apply(
                lambda x: datetime.combine(datetime.today(), x)
            )
        elif isinstance(df_copy['TM'].iloc[0], datetime):
            # Extract only the time portion and set to a common date
            df_copy['minute'] = df_copy['TM'].apply(
                lambda x: datetime.combine(datetime.today(), x.time())
            )
        else:
            # Convert to datetime if it's not already
            df_copy['minute'] = pd.to_datetime(df_copy['TM'])
        
        # Round to the nearest minute
        df_copy['minute'] = df_copy['minute'].dt.round('1min')
        
        # Group by minute and calculate means
        time_groups = df_copy.groupby('minute').agg({
            'Power_DC': 'mean',
            'Power_AC': 'mean',
            'WSD': 'mean',
            'Cp': 'mean'
        }).reset_index()
        
        return time_groups
        
    except Exception as e:
        raise Exception(f"Error processing time data: {str(e)}")


def create_plots(wind_speed_groups, time_groups, max_power_dc, max_power_ac):
    """Create interactive plots"""
    fig = make_subplots(
        rows=4,
        cols=1,
        subplot_titles=('Power vs Wind Speed (Averaged)', 
                       'Power vs Time (Minute Intervals)',
                       'Wind Speed vs Time (Minute Intervals)',
                       'Power Coefficient vs Wind Speed'),
        vertical_spacing=0.2
    )
    
    # First plot: Power vs Wind Speed
    fig.add_trace(
        go.Scatter(
            x=wind_speed_groups['WSD_rounded'], 
            y=wind_speed_groups['Power_DC'], 
            name='DC Power Output', 
            line=dict(color='blue'), 
            mode='lines+markers'
        ), 
        row=1, col=1
    )
    fig.add_trace(
        go.Scatter(
            x=wind_speed_groups['WSD_rounded'], 
            y=wind_speed_groups['Power_AC'], 
            name='AC Power Input', 
            line=dict(color='red'), 
            mode='lines+markers'
        ), 
        row=1, col=1
    )
    
    # Second plot: Power vs Time
    fig.add_trace(
        go.Scatter(
            x=time_groups['minute'], 
            y=time_groups['Power_DC'], 
            name='DC Power Output', 
            line=dict(color='blue')
        ), 
        row=2, col=1
    )
    fig.add_trace(
        go.Scatter(
            x=time_groups['minute'], 
            y=time_groups['Power_AC'], 
            name='AC Power Input', 
            line=dict(color='red')
        ), 
        row=2, col=1
    )
    
    # Third plot: Wind Speed vs Time
    fig.add_trace(
        go.Scatter(
            x=time_groups['minute'],
            y=time_groups['WSD'],
            name='Wind Speed',
            line=dict(color='green')
        ),
        row=3, col=1
    )
    
    # Fourth plot: Power Coefficient vs Wind Speed
    fig.add_trace(
        go.Scatter(
            x=wind_speed_groups['WSD_rounded'],
            y=wind_speed_groups['Cp'],
            name='Power Coefficient',
            line=dict(color='purple'),
            mode='lines+markers'
        ),
        row=4, col=1
    )
    
    # Update layout
    fig.update_layout(
        title='Wind Power Analysis',
        showlegend=True,
        plot_bgcolor='white',
        height=1500  # Increased height for fourth plot
    )
    
    # Update axes
    fig.update_xaxes(title_text="Wind Speed (m/s)", row=1, col=1,range=[wind_speed_groups['WSD_rounded'].min(), wind_speed_groups['WSD_rounded'].max()])
    fig.update_xaxes(title_text="Time", row=2, col=1, tickformat="%H:%M", tickangle=45,range=[time_groups['minute'].min(), time_groups['minute'].max()])
    fig.update_xaxes(title_text="Time", row=3, col=1, tickformat="%H:%M", tickangle=45,range=[time_groups['minute'].min(), time_groups['minute'].max()])
    fig.update_xaxes(title_text="Wind Speed (m/s)", row=4, col=1,range=[wind_speed_groups['WSD_rounded'].min(), wind_speed_groups['WSD_rounded'].max()])
    
    fig.update_yaxes(title_text="Power (kW)", row=1, col=1, range=[0, max_power_ac])
    fig.update_yaxes(title_text="Power (kW)", row=2, col=1, range=[0, max_power_ac])
    fig.update_yaxes(title_text="Wind Speed (m/s)", row=3, col=1, range=[0, 'WSD'])
    fig.update_yaxes(title_text="Power Coefficient (Cp)", row=4, col=1, range=[0, 'Cp'])
    
    # Add grid to all plots
    for i in range(1, 5):
        fig.update_xaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            row=i, col=1
        )
        fig.update_yaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            row=i, col=1
        )
    
    return fig

def find_error_time_ranges(df, error_codes):
    """Find time ranges for each error code"""
    error_ranges = defaultdict(list)
    current_code = None
    start_time = None
    
    # Convert series to list for easier processing
    times = df['TM'].tolist()
    codes = df['CODE 1'].tolist()
    
    for i in range(len(df)):
        code = codes[i]
        if code != current_code:
            if current_code is not None and current_code != 16:  # Save previous range
                error_ranges[current_code].append((start_time, times[i-1]))
            if code != 16:  # Start new range for non-16 code
                start_time = times[i]
            current_code = code
    
    # Handle the last range
    if current_code is not None and current_code != 16:
        error_ranges[current_code].append((start_time, times[-1]))
    
    return error_ranges

def error_analysis(df):
    """Analyze error codes from CODE 1 column"""
    error_codes = {
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
    try:
        if 'CODE 1' not in df.columns:
            return "No error code data available"
        
        # Get error time ranges
        error_ranges = find_error_time_ranges(df, error_codes)
        
        # Calculate efficiency statistics for non-zero power
        valid_efficiency = df[(df['Efficiency'] < 100) & (df['Power_AC'] > 0)]['Efficiency']
        avg_efficiency = valid_efficiency.mean() if len(valid_efficiency) > 0 else 0
        
        # Create analysis message
        analysis = "\nSystem Analysis:\n"
        analysis += f"Average Efficiency (excluding values ≥100% and zero power): {avg_efficiency:.2f}%\n\n"
        analysis += "Error Code Timeline:\n"
        
        # Add error code ranges to analysis
        for code in sorted(error_ranges.keys()):
            if code != 16:  # Skip normal operation code
                ranges = error_ranges[code]
                analysis += f"\nCode {code} - {error_codes.get(code, 'Unknown')}:\n"
                for start, end in ranges:
                    analysis += f"  {start.strftime('%H:%M:%S')} to {end.strftime('%H:%M:%S')}\n"
        
        return analysis
        
    except Exception as e:
        return f"Error in analysis: {str(e)}"

def main():
    st.title("Wind Power Analysis")
    
    # Create a sidebar for parameter inputs
    st.sidebar.header("Turbine Parameters")
    air_density = st.sidebar.number_input(
        "Air Density (kg/m³)",
        min_value=0.0,
        value=1.225,
        help="Standard air density at sea level is approximately 1.225 kg/m³"
    )
    
    rotor_diameter = st.sidebar.number_input(
        "Rotor Diameter (m)",
        min_value=0.0,
        value=10.0,
        help="The diameter of the wind turbine rotor"
    )

    rotor_height = st.sidebar.number_input(
        "Rotor Height (m)",
        min_value=0.0,
        value=10.0,
        help="The height of the wind turbine rotor"
    )
    
    uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")
    
    if uploaded_file is not None:
        try:
            # Read and process data with the new parameters
            df = read_and_process_data(uploaded_file, air_density, rotor_diameter, rotor_height)
            
            # Process data for plots
            wind_speed_groups = process_wind_speed_data(df)
            time_groups = process_time_data(df)
            
            # Get maximum power  value
            max_power_dc = df['Power_DC'].max()
            max_power_ac = df['Power_AC'].max()
            # Create and display plots
            fig = create_plots(wind_speed_groups, time_groups, max_power_dc, max_power_ac)
            st.plotly_chart(fig)
            
            # Display error analysis
            error_message = error_analysis(df)
            st.write(error_message)
            
            # Display additional statistics
            st.write("\nData Statistics:")
            st.write(f"Total number of records: {len(df)}")
            st.write(f"Wind speed range: {df['WSD'].min():.1f} - {df['WSD'].max():.1f} m/s")
            st.write(f"Time range: {min(df['TM']).strftime('%H:%M:%S')} - {max(df['TM']).strftime('%H:%M:%S')}")
            #st.write(f"Maximum Power Coefficient: {df['Cp'].max():.3f}")
            
        except Exception as e:
            st.error(f"An error occurred while processing the data: {str(e)}")

if __name__ == "__main__":
    main()
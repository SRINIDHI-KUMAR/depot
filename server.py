"""
Freight & Volume Analytics Dashboard
Comprehensive dashboard with charts, filters, and KPIs
"""

import os
import re
import pandas as pd
import streamlit as st
import plotly.express as px
import numpy as np

# ---------- Helper to normalize depot names ----------
def normalize_depot(name):
    if pd.isna(name):
        return name
    name = str(name).strip().lower()
    mapping = {
        "coimbatore": "COIMBATORE",
        "bangalore": "BANGALORE",
        "bengaluru": "BANGALORE",
        "parsons": "PARSONS",
        "vijayawada": "VIJAYAWADA",
        "chennai": "CHENNAI",
        "hubli": "HUBLI",
        "cochin": "COCHIN",
        "hyderabad": "HYDERABAD",
    }
    return mapping.get(name, name.upper())

def safe_to_float(val):
    """Safely convert value to float, handling various formats"""
    if pd.isna(val):
        return np.nan
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        val = val.replace(",", "").replace("₹", "").replace("$", "").strip()
        try:
            return float(val)
        except ValueError:
            return np.nan
    try:
        return float(val)
    except (ValueError, TypeError):
        return np.nan

@st.cache_data
def parse_all_freight(file_path):
    """Parse all months from Depot wise Freight Cost sheet."""
    df_raw = pd.read_excel(file_path, sheet_name="Depot wise Freight Cost",
                           header=None, dtype=str)

    depot_desc_cells = []
    for r in range(df_raw.shape[0]):
        for c in range(df_raw.shape[1]):
            val = df_raw.iat[r, c]
            if isinstance(val, str) and "depot desc" in val.strip().lower():
                depot_desc_cells.append((r, c))

    if not depot_desc_cells:
        raise ValueError("No 'Depot Desc' found in the sheet.")

    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    year_match = re.search(r"202[56]", file_path)
    year = int(year_match.group()) if year_match else None

    all_data = {}

    for idx, (header_row, start_col) in enumerate(depot_desc_cells[:12]):
        month = months[idx] if idx < len(months) else f"M{idx+1}"
        records = []
        row = header_row + 1

        while row < df_raw.shape[0]:
            depot_val = df_raw.iat[row, start_col]
            if pd.isna(depot_val) or (isinstance(depot_val, str) and depot_val.strip().lower() == "south"):
                break

            try:
                ded_fc = safe_to_float(df_raw.iat[row, start_col + 1])
                ded_var = safe_to_float(df_raw.iat[row, start_col + 2])
                market = safe_to_float(df_raw.iat[row, start_col + 3])
                parcel = safe_to_float(df_raw.iat[row, start_col + 4])
                inter_depot = safe_to_float(df_raw.iat[row, start_col + 5])
                total_frt = safe_to_float(df_raw.iat[row, start_col + 6])
                vol = safe_to_float(df_raw.iat[row, start_col + 7])
                rpt = safe_to_float(df_raw.iat[row, start_col + 8])
            except IndexError:
                break

            records.append({
                "Depot": normalize_depot(depot_val),
                "Ded_FC": ded_fc,
                "Ded_VAR": ded_var,
                "Market": market,
                "Parcel": parcel,
                "Inter_Depot": inter_depot,
                "Total_FRT": total_frt,
                "Vol": vol,
                "RPT": rpt
            })
            row += 1

        df_month = pd.DataFrame(records)
        df_month = df_month.dropna(subset=["Depot"])
        if not df_month.empty:
            all_data[month] = df_month

    return year, all_data

@st.cache_data
def parse_all_volume(file_path):
    """Parse all months' Primary Net Weight from Last 2 days & first 29 days sheet."""
    df_raw = pd.read_excel(file_path, sheet_name="Last 2 days & first 29 days",
                           header=None, dtype=str)

    primary_rows = []
    for r in range(df_raw.shape[0]):
        val_a = df_raw.iat[r, 0] if df_raw.shape[1] > 0 else None
        if isinstance(val_a, str) and "primary" in val_a.strip().lower():
            primary_rows.append((r, 0))
            continue
        val_b = df_raw.iat[r, 1] if df_raw.shape[1] > 1 else None
        if isinstance(val_b, str) and "primary" in val_b.strip().lower():
            primary_rows.append((r, 1))

    if not primary_rows:
        raise ValueError("No 'Primary' rows found.")

    year_match = re.search(r"202[56]", file_path)
    year = int(year_match.group()) if year_match else None

    depot_cols = {
        3: "COIMBATORE",
        4: "BANGALORE",
        5: "PARSONS",
        6: "VIJAYAWADA",
        7: "CHENNAI",
        8: "HUBLI",
        9: "COCHIN",
        10: "HYDERABAD"
    }

    all_volume = {}

    for primary_row, col_pos in primary_rows:
        net_weight_row = None
        for offset in range(1, 5):
            r = primary_row + offset
            if r >= df_raw.shape[0]:
                break
            val_b = df_raw.iat[r, 1] if df_raw.shape[1] > 1 else None
            if isinstance(val_b, str) and "net weight" in val_b.strip().lower():
                net_weight_row = r
                break
            val_c = df_raw.iat[r, 2] if df_raw.shape[1] > 2 else None
            if isinstance(val_c, str) and "net weight" in val_c.strip().lower():
                net_weight_row = r
                break

        if net_weight_row is None:
            net_weight_row = primary_row + 1

        month_cell = df_raw.iat[primary_row, col_pos]
        month_part = str(month_cell).split()[0]

        records = []
        for col_idx, depot_name in depot_cols.items():
            if col_idx < df_raw.shape[1]:
                val = df_raw.iat[net_weight_row, col_idx]
                if not pd.isna(val):
                    try:
                        net_wt = float(val)
                        records.append({"Depot": depot_name, "Net_Weight": net_wt})
                    except (ValueError, TypeError):
                        pass

        df_month = pd.DataFrame(records)
        if not df_month.empty:
            all_volume[month_part] = df_month

    return year, all_volume

@st.cache_data
def load_data(selected_year):
    """Load data for the selected year"""
    file_pattern = f"Master file for Cost working - South {selected_year}.xlsx"
    
    if not os.path.exists(file_pattern):
        st.error(f"❌ File not found: {file_pattern}")
        st.info("Please ensure the file is in the current directory.")
        return None, None
    
    year, freight_data = parse_all_freight(file_pattern)
    vol_year, volume_data = parse_all_volume(file_pattern)
    
    all_months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                 "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    
    combined_data = []
    for month in all_months:
        if month in freight_data:
            df = freight_data[month].copy()
            df['Month'] = month
            df['Year'] = year
            combined_data.append(df)
    
    if combined_data:
        master_data = pd.concat(combined_data, ignore_index=True)
        
        for month in all_months:
            if month in volume_data:
                volume_df = volume_data[month]
                for idx, row in master_data[master_data['Month'] == month].iterrows():
                    depot = row['Depot']
                    vol_match = volume_df[volume_df['Depot'] == depot]
                    if not vol_match.empty:
                        master_data.loc[idx, 'Net_Weight'] = vol_match.iloc[0]['Net_Weight']
        
        if 'Net_Weight' not in master_data.columns:
            master_data['Net_Weight'] = None
        
        numeric_cols = ['Ded_FC', 'Ded_VAR', 'Market', 'Parcel', 'Inter_Depot', 'Total_FRT', 'Vol', 'RPT']
        for col in numeric_cols:
            if col in master_data.columns:
                master_data[col] = pd.to_numeric(master_data[col], errors='coerce')
        
        return master_data, volume_data
    
    return None, None

def create_dashboard(data, volume_data, selected_year):
    """Create the dashboard with all visualizations"""
    
    st.title(f"📊 Freight & Volume Analytics Dashboard - {selected_year}")
    
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    depots = sorted(data['Depot'].unique())
    
    col1, col2 = st.columns(2)
    
    with col1:
        available_months = sorted(data['Month'].unique(), 
                                 key=lambda x: months.index(x))
        selected_month = st.selectbox("📆 Select Month", available_months)
    
    with col2:
        selected_depot = st.selectbox("🏭 Select Depot", ["All"] + depots)
    
    st.divider()
    
    filtered_data = data[data['Month'] == selected_month]
    if selected_depot != "All":
        filtered_data = filtered_data[filtered_data['Depot'] == selected_depot]
    
    # KPIs
    col1, col2, col3 = st.columns(3)
    
    total_volume = pd.to_numeric(filtered_data['Vol'], errors='coerce').sum() if not filtered_data.empty else 0
    avg_rpt = pd.to_numeric(filtered_data['RPT'], errors='coerce').mean() if not filtered_data.empty else 0
    total_freight = pd.to_numeric(filtered_data['Total_FRT'], errors='coerce').sum() if not filtered_data.empty else 0
    
    with col1:
        st.metric("📦 Total Volume", f"{total_volume:,.0f}" if not pd.isna(total_volume) else "0")
    with col2:
        st.metric("📊 Average RPT", f"{avg_rpt:,.2f}" if not pd.isna(avg_rpt) else "0.00")
    with col3:
        st.metric("💰 Total Freight Cost", f"₹{total_freight:,.2f}" if not pd.isna(total_freight) else "₹0.00")
    
    st.divider()
    
    # Trend data for charts
    trend_data = data.copy()
    if selected_depot != "All":
        trend_data = trend_data[trend_data['Depot'] == selected_depot]
    
    trend_data['RPT'] = pd.to_numeric(trend_data['RPT'], errors='coerce')
    trend_data['Vol'] = pd.to_numeric(trend_data['Vol'], errors='coerce')
    trend_data['Total_FRT'] = pd.to_numeric(trend_data['Total_FRT'], errors='coerce')
    
    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                   "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    trend_data['Month_Num'] = trend_data['Month'].apply(lambda x: month_order.index(x) if x in month_order else 0)
    trend_data = trend_data.sort_values('Month_Num')
    
    col1, col2 = st.columns(2)
    
    # Line Chart
    with col1:
        st.subheader("📈 RPT Trend Over Time")
        if not trend_data.empty and trend_data['RPT'].notna().any():
            fig_line = px.line(trend_data, x='Month', y='RPT', 
                              title=f'RPT Trend - {selected_year}' + (f' - {selected_depot}' if selected_depot != "All" else ''),
                              markers=True, color='Depot' if selected_depot == "All" else None)
            fig_line.update_layout(xaxis_title="Month", yaxis_title="RPT", hovermode='x unified')
            st.plotly_chart(fig_line, width='stretch')  # updated
        else:
            st.info("No data available for RPT trend")
    
    # Bar Chart
    with col2:
        st.subheader("📊 Total Volume per Depot")
        if selected_depot == "All":
            filtered_data['Vol'] = pd.to_numeric(filtered_data['Vol'], errors='coerce')
            depot_volume = filtered_data.groupby('Depot')['Vol'].sum().reset_index()
            depot_volume = depot_volume[depot_volume['Vol'].notna()]
            if not depot_volume.empty:
                fig_bar = px.bar(depot_volume, x='Depot', y='Vol', 
                                title=f'Total Volume by Depot - {selected_month} {selected_year}',
                                color='Depot', text='Vol')
                fig_bar.update_traces(texttemplate='%{text:.0f}', textposition='outside')
                fig_bar.update_layout(xaxis_title="Depot", yaxis_title="Total Volume", showlegend=False)
                st.plotly_chart(fig_bar, width='stretch')
            else:
                st.info("No volume data available")
        else:
            if not trend_data.empty and trend_data['Vol'].notna().any():
                fig_bar = px.bar(trend_data, x='Month', y='Vol', 
                                title=f'Total Volume Trend - {selected_depot}',
                                color='Month', text='Vol')
                fig_bar.update_traces(texttemplate='%{text:.0f}', textposition='outside')
                fig_bar.update_layout(xaxis_title="Month", yaxis_title="Volume", showlegend=False)
                st.plotly_chart(fig_bar, width='stretch')
            else:
                st.info("No volume data available")
    
    # Scatter plot
    st.subheader("🔄 RPT vs Volume Analysis")
    
    if selected_depot == "All":
        scatter_data = filtered_data.copy()
        title_text = f'RPT vs Volume - {selected_month} {selected_year}'
    else:
        scatter_data = trend_data.copy()
        title_text = f'RPT vs Volume - {selected_depot} ({selected_year})'
    
    scatter_data['Vol'] = pd.to_numeric(scatter_data['Vol'], errors='coerce')
    scatter_data['RPT'] = pd.to_numeric(scatter_data['RPT'], errors='coerce')
    scatter_data['Total_FRT'] = pd.to_numeric(scatter_data['Total_FRT'], errors='coerce')
    scatter_data = scatter_data.dropna(subset=['Vol', 'RPT', 'Total_FRT'])
    
    if not scatter_data.empty:
        fig_scatter = px.scatter(scatter_data, x='Vol', y='RPT', 
                                size='Total_FRT', 
                                color='Depot' if selected_depot == "All" else 'Month',
                                hover_name='Depot' if selected_depot == "All" else None,
                                title=title_text,
                                labels={'Vol': 'Volume', 'RPT': 'RPT', 'Total_FRT': 'Freight Cost'},
                                size_max=60)
        fig_scatter.update_layout(xaxis_title="Volume", yaxis_title="RPT", hovermode='closest')
        st.plotly_chart(fig_scatter, width='stretch')
    else:
        st.info("No data available for scatter plot")
    
    # Data table
    st.subheader("📋 Data Table")
    display_data = filtered_data.copy()
    for col in ['RPT', 'Vol', 'Total_FRT', 'Ded_FC', 'Ded_VAR', 'Market', 'Parcel', 'Inter_Depot']:
        if col in display_data.columns:
            display_data[col] = pd.to_numeric(display_data[col], errors='coerce')
    
    st.dataframe(
        display_data,
        width='stretch',  # updated
        hide_index=True,
        column_config={
            "Year": st.column_config.TextColumn("Year"),
            "Month": st.column_config.TextColumn("Month"),
            "Depot": st.column_config.TextColumn("Depot"),
            "RPT": st.column_config.NumberColumn("RPT", format="%.2f"),
            "Vol": st.column_config.NumberColumn("Volume", format="%.0f"),
            "Total_FRT": st.column_config.NumberColumn("Freight Cost", format="₹%.2f"),
            "Ded_FC": st.column_config.NumberColumn("Dedicated FC", format="%.2f"),
            "Ded_VAR": st.column_config.NumberColumn("Dedicated VAR", format="%.2f"),
            "Market": st.column_config.NumberColumn("Market", format="%.2f"),
            "Parcel": st.column_config.NumberColumn("Parcel", format="%.2f"),
            "Inter_Depot": st.column_config.NumberColumn("Inter Depot", format="%.2f"),
        }
    )
    
    csv = display_data.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Filtered Data (CSV)",
        data=csv,
        file_name=f"freight_data_{selected_year}_{selected_month}_{selected_depot}.csv",
        mime='text/csv',
    )

def main():
    st.set_page_config(page_title="Freight & Volume Dashboard", layout="wide")
    
    pattern = re.compile(r"Master file for Cost working - South 202[56]\.xlsx")
    files = [f for f in os.listdir(".") if pattern.match(f)]
    
    if not files:
        st.error("❌ No Excel files found. Please ensure the 2025/2026 files are in the current directory.")
        st.info("Expected files: 'Master file for Cost working - South 2025.xlsx' or 'Master file for Cost working - South 2026.xlsx'")
        return
    
    available_years = sorted([int(re.search(r"202[56]", f).group()) for f in files])
    selected_year = st.selectbox("📅 Select Year", available_years, index=0)
    
    master_data, volume_data = load_data(selected_year)
    
    if master_data is not None:
        create_dashboard(master_data, volume_data, selected_year)
    else:
        st.error("Failed to load data. Please check the file format.")

if __name__ == "__main__":
    main()
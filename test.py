"""
Freight & Volume Analytics Dashboard
Mondelez-inspired Professional Dashboard - Fixed Comparison View
"""

import os
import re
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from datetime import datetime

# ---------- Helper functions ----------
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

# ---------- Parsing functions ----------
@st.cache_data
def parse_all_freight(file_path):
    df_raw = pd.read_excel(file_path, sheet_name="Depot wise Freight Cost",
                           header=None, dtype=str)
    
    depot_desc_cells = []
    for r in range(df_raw.shape[0]):
        for c in range(df_raw.shape[1]):
            val = df_raw.iat[r, c]
            if isinstance(val, str) and "depot desc" in val.strip().lower():
                depot_desc_cells.append((r, c))
    
    if not depot_desc_cells:
        raise ValueError("No 'Depot Desc' found.")
    
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
                "Vol_from_freight": vol,
                "RPT_from_freight": rpt
            })
            row += 1
        
        df_month = pd.DataFrame(records)
        df_month = df_month.dropna(subset=["Depot"])
        if not df_month.empty:
            all_data[month] = df_month
    
    return year, all_data

@st.cache_data
def parse_all_volume(file_path):
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
        
        if records:
            df_month = pd.DataFrame(records)
            parsons_row = df_month[df_month['Depot'] == 'PARSONS']
            if not parsons_row.empty:
                parsons_val = parsons_row.iloc[0]['Net_Weight']
                df_month.loc[df_month['Depot'] == 'BANGALORE', 'Net_Weight'] += parsons_val * 0.50
                df_month.loc[df_month['Depot'] == 'HUBLI', 'Net_Weight'] += parsons_val * 0.30
                df_month.loc[df_month['Depot'] == 'HYDERABAD', 'Net_Weight'] += parsons_val * 0.20
                df_month = df_month[df_month['Depot'] != 'PARSONS']
            
            df_month = df_month[df_month['Depot'].isin(['COIMBATORE', 'BANGALORE', 'VIJAYAWADA',
                                                        'CHENNAI', 'HUBLI', 'COCHIN', 'HYDERABAD'])]
            if not df_month.empty:
                all_volume[month_part] = df_month
    
    return year, all_volume

@st.cache_data
def load_data(selected_year):
    file_pattern = f"Master file for Cost working - South {selected_year}.xlsx"
    if not os.path.exists(file_pattern):
        return None, None
    
    year, freight_data = parse_all_freight(file_pattern)
    vol_year, volume_data = parse_all_volume(file_pattern)
    
    all_months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    
    freight_dfs = []
    for month in all_months:
        if month in freight_data:
            df = freight_data[month].copy()
            df['Month'] = month
            freight_dfs.append(df)
    
    if not freight_dfs:
        return None, None
    
    freight_df = pd.concat(freight_dfs, ignore_index=True)
    
    volume_dfs = []
    for month in all_months:
        if month in volume_data:
            df = volume_data[month].copy()
            df['Month'] = month
            volume_dfs.append(df)
    
    if volume_dfs:
        volume_df = pd.concat(volume_dfs, ignore_index=True)
    else:
        volume_df = pd.DataFrame(columns=['Depot', 'Net_Weight', 'Month'])
    
    master = pd.merge(freight_df, volume_df, on=['Month', 'Depot'], how='left')
    
    master['Vol'] = master['Net_Weight'].copy()
    mask_nan = master['Vol'].isna()
    master.loc[mask_nan, 'Vol'] = master.loc[mask_nan, 'Vol_from_freight']
    
    master['Total_FRT_computed'] = master[['Ded_FC', 'Ded_VAR', 'Market', 'Parcel', 'Inter_Depot']].sum(axis=1)
    master['Total_FRT'] = master['Total_FRT'].fillna(master['Total_FRT_computed'])
    
    mask = (master['Vol'].notna()) & (master['Vol'] > 0)
    master['RPT'] = np.nan
    master.loc[mask, 'RPT'] = (master.loc[mask, 'Total_FRT'] / master.loc[mask, 'Vol']) * 100000
    
    master = master.drop(columns=['Total_FRT_computed', 'Vol_from_freight', 'RPT_from_freight', 'Net_Weight'], errors='ignore')
    master['Year'] = year
    
    numeric_cols = ['Ded_FC', 'Ded_VAR', 'Market', 'Parcel', 'Inter_Depot', 'Total_FRT', 'Vol', 'RPT']
    for col in numeric_cols:
        if col in master.columns:
            master[col] = pd.to_numeric(master[col], errors='coerce')
    
    return master, volume_data

# ---------- Color Palette ----------
COLORS = {
    'primary': '#B39DDB',
    'secondary': '#CE93D8',
    'accent': '#FFE082',
    'success': '#81C784',
    'danger': '#EF9A9A',
    'warning': '#FFCC80',
    'info': '#90CAF9',
    'dark': '#380B3A',
    'light': '#FFF8E1',
    'white': '#FFFFFF',
    'year2025': '#FFCC80',
    'year2026': '#B39DDB',
    'coral': '#FFAB91',
    'teal': '#80CBC4',
}

# ---------- MONTH ORDER ----------
MONTHS_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def get_current_month():
    """Get current month name"""
    current_month_num = datetime.now().month
    return MONTHS_ORDER[current_month_num - 1]

def filter_data_by_month(data, month):
    """Filter data up to a specific month"""
    if month == "All":
        return data
    month_index = MONTHS_ORDER.index(month)
    months_to_keep = MONTHS_ORDER[:month_index + 1]
    return data[data['Month'].isin(months_to_keep)]

# ---------- YEAR VIEW Functions ----------
def create_enhanced_volume_bar(filtered):
    perf = filtered.groupby('Depot').agg({'Vol': 'sum'}).reset_index().dropna()
    if perf.empty:
        return None
    perf = perf.sort_values('Vol', ascending=True)
    colors = [
    '#4DD0E1',   # Aqua
    '#FFB74D',   # Orange
    '#BA68C8',   # Purple
    '#FF8A80',   # Coral
    '#FFD54F',   # Gold
    '#81C784',   # Mint
    '#64B5F6'    # Sky Blue
]
    
    fig = go.Figure()
    for i, (_, row) in enumerate(perf.iterrows()):
        fig.add_trace(go.Bar(
            y=[row['Depot']], x=[row['Vol']], orientation='h',
            marker_color=colors[i % len(colors)],
            text=[f"{row['Vol']:,.0f} MT"], textposition='outside',
            textfont=dict(color='white', size=10, weight='bold'),
            hovertemplate='<b>%{y}</b><br>📦 Volume: %{x:,.0f} MT<extra></extra>',
            hoverlabel=dict(bgcolor='rgba(20,20,30,0.9)', font=dict(color='white', size=12), bordercolor='rgba(255,255,255,0.2)'),
            showlegend=False, width=0.7,
            marker=dict(line=dict(color='rgba(255,255,255,0.2)', width=1), cornerradius=4)
        ))
    
    fig.update_layout(
        height=280, margin=dict(l=0, r=30, t=10, b=10),
        font=dict(size=10, color=COLORS['white']),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        xaxis_title='Volume (MT)', yaxis_title='',
        xaxis=dict(gridcolor='rgba(255,255,255,0.06)', zerolinecolor='rgba(255,255,255,0.1)', zerolinewidth=1),
        barmode='group', bargap=0.3, bargroupgap=0, hovermode='y unified'
    )
    return fig

def create_enhanced_rpt_donut(filtered):
    perf = filtered.groupby('Depot').agg({'Total_FRT': 'sum', 'RPT': 'mean'}).reset_index().dropna()
    if perf.empty:
        return None
    perf = perf.sort_values('Total_FRT', ascending=False)
    
    depot_colors = {
    'COIMBATORE':'#EF5350',
    'BANGALORE':'#FFB74D',
    'HYDERABAD':'#FFCA28',
    'VIJAYAWADA':'#64B5F6',
    'HUBLI':'#66BB6A',
    'CHENNAI':'#BA68C8',
    'COCHIN':'#9575CD'
}
    colors = [depot_colors.get(d, '#9C27B0') for d in perf['Depot']]
    
    hover_text = []
    for _, row in perf.iterrows():
        hover_text.append(f"{row['Depot']}<br>💰 Freight: ₹{row['Total_FRT']:,.0f}<br>📊 RPT: ₹{row['RPT']:,.1f}")
    
    fig = go.Figure()
    fig.add_trace(go.Pie(
        labels=[' '], values=[100], hole=0.55,
        marker=dict(colors=['rgba(0,0,0,0.15)'], line=dict(color='rgba(0,0,0,0)', width=0)),
        textinfo='none', hoverinfo='skip', showlegend=False, sort=False
    ))
    fig.add_trace(go.Pie(
        labels=perf['Depot'], values=perf['Total_FRT'], hole=0.55,
        textinfo='label+percent', textposition='inside',
        textfont=dict(color='white', size=10, weight='bold'),
        marker=dict(colors=colors, line=dict(color='rgba(255,255,255,0.3)', width=2)),
        pull=[0.08 if i == 0 else 0.02 for i in range(len(perf))],
        hovertemplate='%{customdata}<extra></extra>',
        hoverlabel=dict(bgcolor='rgba(20,20,30,0.9)', font=dict(color='white', size=12), bordercolor='rgba(255,255,255,0.2)'),
        customdata=hover_text, showlegend=False, sort=False
    ))
    
    fig.update_layout(
        height=280, margin=dict(l=0, r=0, t=10, b=10),
        font=dict(size=10, color=COLORS['white']),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        shapes=[
            dict(type='rect', xref='paper', yref='paper', x0=0.02, y0=0.02, x1=0.98, y1=0.98,
                 line=dict(color='rgba(255,255,255,0.15)', width=2), fillcolor='rgba(0,0,0,0)', layer='below'),
            dict(type='rect', xref='paper', yref='paper', x0=0.06, y0=0.06, x1=0.94, y1=0.94,
                 line=dict(color='rgba(255,255,255,0.06)', width=1), fillcolor='rgba(0,0,0,0)', layer='below')
        ]
    )
    return fig

def create_enhanced_rpt_bars(filtered):
    perf = filtered.groupby('Depot').agg({'RPT': 'mean'}).reset_index().dropna()
    if perf.empty:
        return None
    perf = perf.sort_values('RPT', ascending=True)
    
    q1 = perf['RPT'].quantile(0.33)
    q2 = perf['RPT'].quantile(0.67)
    colors = [
    '#66BB6A' if v <= q1 else
    '#FFCA28' if v <= q2 else
    '#EF5350'
    for v in perf['RPT']
]
    fig = go.Figure()
    for i, (_, row) in enumerate(perf.iterrows()):
        fig.add_trace(go.Bar(
            x=[row['Depot']], y=[row['RPT'] * 1.02],
            marker_color='rgba(0,0,0,0.15)', showlegend=False, hoverinfo='skip', width=0.6
        ))
        fig.add_trace(go.Bar(
            x=[row['Depot']], y=[row['RPT']], marker_color=colors[i],
            text=[f"₹{row['RPT']:,.0f}"], textposition='outside',
            textfont=dict(color='white', size=10, weight='bold'),
            hovertemplate='<b>%{x}</b><br>📊 RPT: ₹%{y:,.1f}<extra></extra>',
            hoverlabel=dict(bgcolor='rgba(20,20,30,0.9)', font=dict(color='white', size=12), bordercolor='rgba(255,255,255,0.2)'),
            showlegend=False, width=0.6,
            marker=dict(line=dict(color='rgba(255,255,255,0.2)', width=1), cornerradius=4)
        ))
    
    avg_rpt = perf['RPT'].mean()
    fig.add_hline(y=avg_rpt, line_dash="dash", line_color=COLORS['white'], opacity=0.8, line_width=2,
                  annotation_text=f"Avg: ₹{avg_rpt:,.0f}", annotation_font_color=COLORS['white'], annotation_font_size=10)
    
    fig.update_layout(
        height=280, margin=dict(l=0, r=0, t=10, b=10),
        font=dict(size=10, color=COLORS['white']),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        xaxis_title='Depot', yaxis_title='RPT (₹ per 100,000)',
        yaxis=dict(gridcolor='rgba(255,255,255,0.06)', zerolinecolor='rgba(255,255,255,0.1)', zerolinewidth=1),
        barmode='group', bargap=0.3, bargroupgap=0, hovermode='x'
    )
    return fig

def create_enhanced_stacked_bar(filtered):
    components = ['Ded_FC', 'Ded_VAR', 'Market', 'Parcel', 'Inter_Depot']
    comp_labels = ['Dedicated FC', 'Dedicated VAR', 'Market', 'Parcel', 'Inter Depot']
    
    perf = filtered.groupby('Depot')[components].sum().reset_index()
    if perf.empty:
        return None
    
    perf['Total'] = perf[components].sum(axis=1)
    perf = perf.sort_values('Total', ascending=True)
    
    melted_data = []
    for _, row in perf.iterrows():
        depot = row['Depot']
        for comp, label in zip(components, comp_labels):
            melted_data.append({'Depot': depot, 'Component': label, 'Cost': row[comp]})
    melted = pd.DataFrame(melted_data)
    
    color_map = {
    'Dedicated FC':'#EF5350',
    'Dedicated VAR':'#FFB74D',
    'Market':'#FFCA28',
    'Parcel':'#66BB6A',
    'Inter Depot':'#64B5F6'
}
    
    fig = px.bar(melted, x='Cost', y='Depot', color='Component', orientation='h',
                 color_discrete_map=color_map, text='Cost', barmode='stack',
                 hover_data={'Cost': ':₹,.0f'})
    
    fig.update_traces(
        texttemplate='₹%{text:,.0f}', textposition='inside',
        textfont=dict(color='white', size=8, weight='bold'),
        hovertemplate='<b>%{y}</b><br>%{fullData.name}: ₹%{x:,.0f}<extra></extra>',
        hoverlabel=dict(bgcolor='rgba(20,20,30,0.9)', font=dict(color='white', size=11), bordercolor='rgba(255,255,255,0.2)'),
        marker=dict(line=dict(color='rgba(255,255,255,0.15)', width=1), cornerradius=2)
    )
    
    fig.update_layout(
        height=280, margin=dict(l=0, r=0, t=10, b=10),
        font=dict(size=9, color=COLORS['white']),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        xaxis_title='Cost (₹)', yaxis_title='',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5,
                   font=dict(size=7, color=COLORS['white']), bgcolor='rgba(0,0,0,0.2)',
                   bordercolor='rgba(255,255,255,0.1)', borderwidth=1),
        barmode='stack',
        xaxis=dict(gridcolor='rgba(255,255,255,0.06)', zerolinecolor='rgba(255,255,255,0.1)', zerolinewidth=1),
        hovermode='y unified'
    )
    return fig

def create_enhanced_waterfall_chart(filtered, selected_year):
    perf = filtered.groupby('Depot').agg({'Total_FRT': 'sum', 'Vol': 'sum', 'RPT': 'mean'}).reset_index().dropna()
    if perf.empty:
        return None, None, None
    
    perf = perf.sort_values('Total_FRT', ascending=False)
    total_cost = perf['Total_FRT'].sum()
    perf['Cost_Percentage'] = (perf['Total_FRT'] / total_cost * 100).round(1)
    
    top_driver = perf.iloc[0]['Depot']
    top_driver_pct = perf.iloc[0]['Cost_Percentage']
    
    fig = go.Figure()
    colors = [
    '#EF5350' if row['Cost_Percentage'] > 25
    else '#FFA726' if row['Cost_Percentage'] > 15
    else '#FFCA28' if row['Cost_Percentage'] > 10
    else '#66BB6A'
    for _, row in perf.iterrows()
]
    for i, (_, row) in enumerate(perf.iterrows()):
        fig.add_trace(go.Bar(
            y=[row['Depot']], x=[row['Total_FRT']], orientation='h',
            marker_color=colors[i],
            text=[f"₹{row['Total_FRT']:,.0f}<br>({row['Cost_Percentage']:.1f}%)"],
            textposition='outside', textfont=dict(color='white', size=9),
            hovertemplate='<b>%{y}</b><br>💰 Cost: ₹%{x:,.0f}<br>📊 Share: %{customdata[0]:.1f}%<br>📦 Volume: %{customdata[1]:,.0f}<br>📈 RPT: %{customdata[2]:.1f}<extra></extra>',
            hoverlabel=dict(bgcolor='rgba(20,20,30,0.9)', font=dict(color='white', size=11), bordercolor='rgba(255,255,255,0.2)'),
            customdata=perf[['Cost_Percentage', 'Vol', 'RPT']].values,
            showlegend=False, width=0.7,
            marker=dict(line=dict(color='rgba(255,255,255,0.2)', width=1), cornerradius=4)
        ))
    
    avg_cost = perf['Total_FRT'].mean()
    fig.add_vline(x=avg_cost, line_dash="dash", line_color=COLORS['white'], opacity=0.6, line_width=2,
                  annotation_text="Avg", annotation_font_color=COLORS['white'], annotation_font_size=10)
    
    fig.update_layout(
        height=250, margin=dict(l=0, r=0, t=10, b=10),
        font=dict(size=9, color=COLORS['white']),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        xaxis_title='Total Freight Cost (₹)', yaxis_title='',
        xaxis=dict(gridcolor='rgba(255,255,255,0.06)', zerolinecolor='rgba(255,255,255,0.1)', zerolinewidth=1),
        barmode='group', bargap=0.3, bargroupgap=0, hovermode='y unified'
    )
    return fig, top_driver, top_driver_pct

# ---------- YEAR VIEW Dashboard ----------
def create_year_dashboard(data, selected_year, selected_month):
    if selected_month == "All":
        filtered = data.copy()
        is_month_view = False
    else:
        filtered = filter_data_by_month(data, selected_month)
        is_month_view = True
    
    # KPI Cards
    c1, c2, c3, c4, c5 = st.columns(5)
    total_frt = pd.to_numeric(filtered['Total_FRT'], errors='coerce').sum()
    avg_rpt = pd.to_numeric(filtered['RPT'], errors='coerce').mean()
    total_vol = pd.to_numeric(filtered['Vol'], errors='coerce').sum()
    rpt_series = filtered.groupby('Depot')['RPT'].mean().sort_values(ascending=False)
    highest = rpt_series.index[0] if not rpt_series.empty else "N/A"
    lowest = rpt_series.index[-1] if not rpt_series.empty else "N/A"
    
    c1.metric("💰 Total Freight", f"₹{total_frt:,.0f}" if not pd.isna(total_frt) else "₹0")
    c2.metric("📊 Avg RPT", f"{avg_rpt:,.1f}" if not pd.isna(avg_rpt) else "0")
    c3.metric("📦 Total Volume", f"{total_vol:,.0f}" if not pd.isna(total_vol) else "0")
    c4.metric("🏆 Highest RPT", highest)
    c5.metric("📉 Lowest RPT", lowest)
    st.divider()
    
    if is_month_view:
        st.markdown(f'<p style="color:{COLORS["light"]};font-size:0.9rem;font-weight:600;text-align:center;margin-bottom:0.5rem;">📊 {selected_month} {selected_year} - Monthly Performance Analysis</p>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.8rem;font-weight:600;margin-bottom:0.2rem;">📦 Volume by Depot</p>', unsafe_allow_html=True)
            volume_chart = create_enhanced_volume_bar(filtered)
            if volume_chart:
                st.plotly_chart(volume_chart, use_container_width=True, config={'displayModeBar': False})
            else:
                st.write("No data available")
        
        with col2:
            st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.8rem;font-weight:600;margin-bottom:0.2rem;">🍩 Freight & RPT Distribution</p>', unsafe_allow_html=True)
            donut = create_enhanced_rpt_donut(filtered)
            if donut:
                st.plotly_chart(donut, use_container_width=True, config={'displayModeBar': False})
            else:
                st.write("No data available")
        
        st.divider()
        
        col3, col4 = st.columns(2)
        with col3:
            st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.8rem;font-weight:600;margin-bottom:0.2rem;">📊 RPT Analysis (Color-coded)</p>', unsafe_allow_html=True)
            rpt_chart = create_enhanced_rpt_bars(filtered)
            if rpt_chart:
                st.plotly_chart(rpt_chart, use_container_width=True, config={'displayModeBar': False})
            else:
                st.write("No data available")
        
        with col4:
            st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.8rem;font-weight:600;margin-bottom:0.2rem;">📊 Cost Components Breakdown</p>', unsafe_allow_html=True)
            stacked = create_enhanced_stacked_bar(filtered)
            if stacked:
                st.plotly_chart(stacked, use_container_width=True, config={'displayModeBar': False})
            else:
                st.write("No data available")
        st.divider()
        
    else:
        # YEAR VIEW
        col_left1, col_right1 = st.columns([1, 1])
        with col_left1:
            st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.2rem;">📊 RPT Ranking by Depot</p>', unsafe_allow_html=True)
            depot_rpt = filtered.groupby('Depot')['RPT'].mean().sort_values(ascending=True).reset_index().dropna()
            if not depot_rpt.empty:
                fig1 = px.bar(depot_rpt, x='RPT', y='Depot', orientation='h',
                              color='RPT', color_continuous_scale=[COLORS['success'], COLORS['accent'], COLORS['danger']],
                              text='RPT')
                fig1.update_traces(texttemplate='%{text:.0f}', textposition='outside', textfont_color=COLORS['white'])
                fig1.update_layout(height=220, margin=dict(l=0, r=0, t=10, b=10),
                                   font=dict(size=10, color=COLORS['white']), showlegend=False,
                                   xaxis_title="RPT", yaxis_title="",
                                   plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig1, use_container_width=True)
        
        with col_right1:
            st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.2rem;">📊 Monthly RPT Trend</p>', unsafe_allow_html=True)
            monthly_rpt = filtered.groupby(['Month', 'Depot'])['RPT'].mean().reset_index()
            monthly_rpt['Month'] = pd.Categorical(monthly_rpt['Month'], categories=MONTHS_ORDER, ordered=True)
            monthly_rpt = monthly_rpt.sort_values('Month')
            
            if not monthly_rpt.empty:
                fig2 = px.line(monthly_rpt, x='Month', y='RPT', color='Depot', markers=True)
                fig2.update_traces(line=dict(width=2))
                fig2.update_layout(height=220, margin=dict(l=0, r=0, t=10, b=10),
                                   font=dict(size=9, color=COLORS['white']),
                                   legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5, font=dict(size=7, color=COLORS['white'])),
                                   plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig2, use_container_width=True)
        
        st.divider()
        
        col_left2, col_right2 = st.columns(2)
        with col_left2:
            st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.2rem;">📊 Cost Component Composition</p>', unsafe_allow_html=True)
            comp = filtered.groupby('Depot')[['Ded_FC','Ded_VAR','Market','Parcel','Inter_Depot']].sum().reset_index()
            if not comp.empty:
                comp_pct = comp.copy()
                for col in ['Ded_FC','Ded_VAR','Market','Parcel','Inter_Depot']:
                    comp_pct[col] = (comp_pct[col] / comp_pct[['Ded_FC','Ded_VAR','Market','Parcel','Inter_Depot']].sum(axis=1)) * 100
                
                fig3 = px.bar(comp_pct, x='Depot', y=['Ded_FC','Ded_VAR','Market','Parcel','Inter_Depot'],
                              barmode='stack', labels={'value':'% of Total Cost', 'variable':'Component'},
                              color_discrete_map={
                                  'Ded_FC': COLORS['danger'], 'Ded_VAR': COLORS['warning'],
                                  'Market': COLORS['accent'], 'Parcel': COLORS['success'],
                                  'Inter_Depot': COLORS['info']
                              })
                fig3.update_layout(height=250, margin=dict(l=0, r=0, t=10, b=10),
                                   font=dict(size=9, color=COLORS['white']),
                                   legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5, font=dict(size=8, color=COLORS['white'])),
                                   plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                fig3.update_traces(texttemplate='%{y:.1f}%', textposition='inside', textfont_color='white')
                st.plotly_chart(fig3, use_container_width=True)
        
        with col_right2:
            st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.2rem;">📊 Cost Driver Analysis</p>', unsafe_allow_html=True)
            waterfall, top_driver, top_pct = create_enhanced_waterfall_chart(filtered, selected_year)
            if waterfall:
                st.plotly_chart(waterfall, use_container_width=True, config={'displayModeBar': False})
                st.markdown(f'<p style="font-size:0.65rem;color:{COLORS["light"]};">💡 <b>{top_driver}</b> is the largest cost driver at <b>{top_pct:.1f}%</b> of total freight cost</p>', unsafe_allow_html=True)
            else:
                st.write("No data available")
        st.divider()
    
    # Performance Matrix
    st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.5rem;">📊 Performance Matrix</p>', unsafe_allow_html=True)
    
    perf = filtered.groupby('Depot').agg({'Total_FRT': 'sum', 'Vol': 'sum', 'RPT': 'mean'}).reset_index().dropna()
    if not perf.empty:
        perf = perf.sort_values('RPT', ascending=False)
        vol_high = perf['Vol'].quantile(0.67)
        vol_low = perf['Vol'].quantile(0.33)
        rpt_high = perf['RPT'].quantile(0.67)
        rpt_low = perf['RPT'].quantile(0.33)
        
        html_table = """
        <table style="width:100%; border-collapse: collapse; font-family: 'Segoe UI', sans-serif; font-size: 0.8rem; border-radius: 10px; overflow: hidden; box-shadow: 0 8px 32px rgba(0,0,0,0.4), 0 2px 8px rgba(0,0,0,0.2);">
            <thead>
                <tr style="background: linear-gradient(135deg, #4a1a6b, #7D4BAE, #9B59B6);">
                    <th style="padding: 12px 14px; text-align: left; color: white; font-weight: 600; border-bottom: 2px solid rgba(255,255,255,0.15);">Depot</th>
                    <th style="padding: 12px 14px; text-align: right; color: white; font-weight: 600; border-bottom: 2px solid rgba(255,255,255,0.15);">Total Freight</th>
                    <th style="padding: 12px 14px; text-align: right; color: white; font-weight: 600; border-bottom: 2px solid rgba(255,255,255,0.15);">Volume</th>
                    <th style="padding: 12px 14px; text-align: right; color: white; font-weight: 600; border-bottom: 2px solid rgba(255,255,255,0.15);">RPT</th>
                </tr>
            </thead>
            <tbody>
        """
        
        for _, row in perf.iterrows():
            vol = row['Vol']; rpt = row['RPT']
            if rpt >= rpt_high and vol <= vol_low:
                bg_color, text_color, font_weight, emoji = '#E74C3C', 'white', '600', '🔴'
            elif rpt <= rpt_low and vol >= vol_high:
                bg_color, text_color, font_weight, emoji = '#2ECC71', 'white', '600', '🟢'
            elif rpt >= rpt_high and vol >= vol_high:
                bg_color, text_color, font_weight, emoji = '#3498DB', 'white', '400', '🔵'
            elif rpt <= rpt_low and vol <= vol_low:
                bg_color, text_color, font_weight, emoji = '#F1C40F', 'black', '400', '🟡'
            else:
                bg_color, text_color, font_weight, emoji = '#5d2d8a', 'white', '400', '⚪'
            
            html_table += f"""
                <tr style="background-color: {bg_color}; transition: all 0.2s; box-shadow: inset 0 1px 0 rgba(255,255,255,0.05);">
                    <td style="padding: 8px 14px; text-align: left; color: {text_color}; font-weight: {font_weight}; border-bottom: 1px solid rgba(255,255,255,0.06);">{emoji} {row['Depot']}</td>
                    <td style="padding: 8px 14px; text-align: right; color: {text_color}; font-weight: {font_weight}; border-bottom: 1px solid rgba(255,255,255,0.06);">₹{row['Total_FRT']:,.0f}</td>
                    <td style="padding: 8px 14px; text-align: right; color: {text_color}; font-weight: {font_weight}; border-bottom: 1px solid rgba(255,255,255,0.06);">{row['Vol']:,.0f}</td>
                    <td style="padding: 8px 14px; text-align: right; color: {text_color}; font-weight: {font_weight}; border-bottom: 1px solid rgba(255,255,255,0.06);">{row['RPT']:,.1f}</td>
                </tr>
            """
        
        html_table += """
            </tbody>
        </table>
        """
        st.components.v1.html(html_table, height=300, scrolling=False)
        
        st.markdown("""
        <div style="display: flex; gap: 24px; margin-top: 12px; flex-wrap: wrap; font-size: 0.7rem; padding: 8px 4px; background: rgba(255,255,255,0.03); border-radius: 8px;">
            <span style="color: #2ECC71; font-weight: 600;">🟢 Efficient (Low RPT / High Volume)</span>
            <span style="color: #E74C3C; font-weight: 600;">🔴 Inefficient (High RPT / Low Volume)</span>
            <span style="color: #F1C40F; font-weight: 600;">🟡 Low Volume (Needs Growth)</span>
            <span style="color: #3498DB; font-weight: 600;">🔵 High Volume & RPT</span>
            <span style="color: #9B59B6; font-weight: 600;">⚪ Average Performer</span>
        </div>
        """, unsafe_allow_html=True)

# ---------- COMPARISON DASHBOARD ----------
def create_comparison_dashboard(data_2025, data_2026):
    # --- Determine current month for comparison ---
    current_month = get_current_month()
    
    # --- Filter data up to current month for both years ---
    data_2025_filtered = filter_data_by_month(data_2025, current_month)
    data_2026_filtered = filter_data_by_month(data_2026, current_month)
    
    # --- KPI Cards ---
    c1, c2, c3, c4, c5 = st.columns(5)
    
    frt25 = pd.to_numeric(data_2025_filtered['Total_FRT'], errors='coerce').sum()
    frt26 = pd.to_numeric(data_2026_filtered['Total_FRT'], errors='coerce').sum()
    rpt25 = pd.to_numeric(data_2025_filtered['RPT'], errors='coerce').mean()
    rpt26 = pd.to_numeric(data_2026_filtered['RPT'], errors='coerce').mean()
    vol25 = pd.to_numeric(data_2025_filtered['Vol'], errors='coerce').sum()
    vol26 = pd.to_numeric(data_2026_filtered['Vol'], errors='coerce').sum()
    
    chg_frt = ((frt26 - frt25) / frt25 * 100) if frt25 > 0 else 0
    chg_rpt = ((rpt26 - rpt25) / rpt25 * 100) if rpt25 > 0 else 0
    chg_vol = ((vol26 - vol25) / vol25 * 100) if vol25 > 0 else 0
    
    c1.metric("📊 RPT 2025", f"₹{rpt25:,.1f}")
    c2.metric("📊 RPT 2026", f"₹{rpt26:,.1f}", delta=f"{chg_rpt:+.1f}%" if chg_rpt != 0 else "0%")
    c3.metric("💰 Freight Change", f"₹{frt26 - frt25:+,.0f}", delta=f"{chg_frt:+.1f}%" if chg_frt != 0 else "0%")
    c4.metric("📦 Volume Change", f"{vol26 - vol25:+,.0f} MT", delta=f"{chg_vol:+.1f}%" if chg_vol != 0 else "0%")
    c5.metric("💰 Total Freight 2025", f"₹{frt25:,.0f}")
    
    st.divider()
    
    # --- Row 1: Monthly Trends ---
    col1, col2 = st.columns(2)
    
    all_months = [m for m in MONTHS_ORDER[:MONTHS_ORDER.index(current_month) + 1]]
    
    with col1:
        st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.2rem;">📈 RPT Trend (Jan - {current_month})</p>', unsafe_allow_html=True)
        rpt25_df = data_2025_filtered.groupby('Month')['RPT'].mean().reset_index()
        rpt26_df = data_2026_filtered.groupby('Month')['RPT'].mean().reset_index()
        merged = pd.DataFrame({'Month': all_months})
        merged = merged.merge(rpt25_df, on='Month', how='left').rename(columns={'RPT':'2025'})
        merged = merged.merge(rpt26_df, on='Month', how='left').rename(columns={'RPT':'2026'})
        melted = merged.melt(id_vars=['Month'], value_vars=['2025','2026'],
                             var_name='Year', value_name='RPT')
        melted['Month'] = pd.Categorical(melted['Month'], categories=MONTHS_ORDER, ordered=True)
        melted = melted.sort_values('Month')
        
        fig1 = px.line(melted, x='Month', y='RPT', color='Year', markers=True,
                       color_discrete_map={'2025': '#F39C12', '2026': '#8B5CF6'})
        fig1.update_traces(connectgaps=False, line=dict(width=3), marker=dict(size=8))
        fig1.update_layout(height=200, font=dict(size=9, color=COLORS['white']), margin=dict(l=0, r=0, t=10, b=10),
                           legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5, font=dict(size=8, color=COLORS['white'])),
                           plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                           yaxis=dict(gridcolor='rgba(255,255,255,0.06)'))
        st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.2rem;">📈 Volume Trend (Jan - {current_month})</p>', unsafe_allow_html=True)
        vol25_df = data_2025_filtered.groupby('Month')['Vol'].sum().reset_index()
        vol26_df = data_2026_filtered.groupby('Month')['Vol'].sum().reset_index()
        merged2 = pd.DataFrame({'Month': all_months})
        merged2 = merged2.merge(vol25_df, on='Month', how='left').rename(columns={'Vol':'2025'})
        merged2 = merged2.merge(vol26_df, on='Month', how='left').rename(columns={'Vol':'2026'})
        melted2 = merged2.melt(id_vars=['Month'], value_vars=['2025','2026'],
                               var_name='Year', value_name='Volume')
        melted2['Month'] = pd.Categorical(melted2['Month'], categories=MONTHS_ORDER, ordered=True)
        melted2 = melted2.sort_values('Month')
        melted2 = melted2[(melted2['Volume'] != 0) & (melted2['Volume'].notna())]
        
        fig2 = px.line(melted2, x='Month', y='Volume', color='Year', markers=True,
                       color_discrete_map={'2025': '#FF6B6B', '2026': '#4ECDC4'})
        fig2.update_traces(connectgaps=False, line=dict(width=3), marker=dict(size=8))
        fig2.update_layout(height=200, font=dict(size=9, color=COLORS['white']), margin=dict(l=0, r=0, t=10, b=10),
                           legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5, font=dict(size=8, color=COLORS['white'])),
                           plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                           yaxis=dict(gridcolor='rgba(255,255,255,0.06)'))
        st.plotly_chart(fig2, use_container_width=True)
    
    st.divider()
    
    # --- Row 2: Metrics Change Summary + Professional Summary ---
    col3, col4 = st.columns([1, 1])
    
    with col3:
        st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.2rem;">📊 Metrics Change Summary (Jan - {current_month})</p>', unsafe_allow_html=True)
        
        # Calculate change percentages for each metric across all depots
        frt_change_pct = ((frt26 - frt25) / frt25 * 100) if frt25 > 0 else 0
        rpt_change_pct = ((rpt26 - rpt25) / rpt25 * 100) if rpt25 > 0 else 0
        vol_change_pct = ((vol26 - vol25) / vol25 * 100) if vol25 > 0 else 0
        
        # Create a simple but attractive summary chart
        metrics = ['Freight', 'RPT', 'Volume']
        changes = [frt_change_pct, rpt_change_pct, vol_change_pct]
        colors = ['#FF6B6B' if x > 0 else '#4ECDC4' if x < 0 else '#FDCB6E' for x in changes]
        
        fig3 = go.Figure(go.Bar(
            x=metrics,
            y=changes,
            marker_color=colors,
            text=[f"{x:+.1f}%" for x in changes],
            textposition='outside',
            textfont=dict(color='white', size=12, weight='bold'),
            hovertemplate='<b>%{x}</b><br>Change: %{y:+.1f}%<extra></extra>',
            marker=dict(
                line=dict(color='rgba(255,255,255,0.2)', width=1),
                cornerradius=6
            )
        ))
        
        fig3.add_hline(y=0, line_dash="solid", line_color='rgba(255,255,255,0.3)', line_width=1)
        
        fig3.update_layout(
            height=200,
            margin=dict(l=0, r=0, t=10, b=10),
            font=dict(size=10, color=COLORS['white']),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(
                title='',
                tickfont=dict(color='white', size=10),
                gridcolor='rgba(255,255,255,0.06)'
            ),
            yaxis=dict(
                title='Change %',
                tickfont=dict(color='white', size=9),
                gridcolor='rgba(255,255,255,0.06)',
                zerolinecolor='rgba(255,255,255,0.2)',
                zerolinewidth=1
            ),
            showlegend=False,
            hovermode='x'
        )
        st.plotly_chart(fig3, use_container_width=True, config={'displayModeBar': False})
    
    with col4:
        st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.2rem;">📊 Period Performance Summary</p>', unsafe_allow_html=True)
        
        # Calculate meaningful statistics
        total_frt_25 = data_2025_filtered['Total_FRT'].sum()
        total_frt_26 = data_2026_filtered['Total_FRT'].sum()
        avg_rpt_25 = data_2025_filtered['RPT'].mean()
        avg_rpt_26 = data_2026_filtered['RPT'].mean()
        total_vol_25 = data_2025_filtered['Vol'].sum()
        total_vol_26 = data_2026_filtered['Vol'].sum()
        
        # Calculate depot counts for improvement
        rpt_improved = data_2026_filtered.groupby('Depot')['RPT'].mean() - data_2025_filtered.groupby('Depot')['RPT'].mean()
        rpt_improved_count = len(rpt_improved[rpt_improved < 0])
        rpt_worsened_count = len(rpt_improved[rpt_improved > 0])
        
        vol_grown = data_2026_filtered.groupby('Depot')['Vol'].sum() - data_2025_filtered.groupby('Depot')['Vol'].sum()
        vol_grown_count = len(vol_grown[vol_grown > 0])
        vol_declined_count = len(vol_grown[vol_grown < 0])
        
        # Format numbers for display
        frt_change_abs = total_frt_26 - total_frt_25
        vol_change_abs = total_vol_26 - total_vol_25
        rpt_change_abs = avg_rpt_26 - avg_rpt_25
        
        summary_html = f"""
        <div style="background: linear-gradient(135deg, #5d2d8a, #7D4BAE); border-radius: 10px; padding: 14px 16px; height: 100%; min-height: 200px; border: 1px solid rgba(255,255,255,0.1);">
            <div style="font-size: 0.7rem; color: #E8D5B7; font-weight: 600; letter-spacing: 1px; margin-bottom: 8px; text-align: center; border-bottom: 1px solid rgba(255,255,255,0.05); padding-bottom: 6px;">PERIOD PERFORMANCE SUMMARY</div>
            <div style="display: flex; flex-direction: column; gap: 5px;">
                <div style="display: flex; justify-content: space-between; align-items: center; background: rgba(255,255,255,0.05); padding: 5px 10px; border-radius: 6px;">
                    <span style="color: #E8D5B7; font-size: 0.6rem;">Total Freight Change</span>
                    <span style="color: {'#4ECDC4' if frt_change_abs < 0 else '#FF6B6B' if frt_change_abs > 0 else '#FDCB6E'}; font-size: 0.8rem; font-weight: 700;">₹{frt_change_abs:+,.0f}</span>
                </div>
                <div style="display: flex; justify-content: space-between; align-items: center; background: rgba(255,255,255,0.05); padding: 5px 10px; border-radius: 6px;">
                    <span style="color: #E8D5B7; font-size: 0.6rem;">Average RPT Change</span>
                    <span style="color: {'#4ECDC4' if rpt_change_abs < 0 else '#FF6B6B' if rpt_change_abs > 0 else '#FDCB6E'}; font-size: 0.8rem; font-weight: 700;">{rpt_change_abs:+,.1f}</span>
                </div>
                <div style="display: flex; justify-content: space-between; align-items: center; background: rgba(255,255,255,0.05); padding: 5px 10px; border-radius: 6px;">
                    <span style="color: #E8D5B7; font-size: 0.6rem;">Total Volume Change</span>
                    <span style="color: {'#4ECDC4' if vol_change_abs > 0 else '#FF6B6B' if vol_change_abs < 0 else '#FDCB6E'}; font-size: 0.8rem; font-weight: 700;">{vol_change_abs:+,.0f} MT</span>
                </div>
                <div style="display: flex; justify-content: space-between; align-items: center; background: rgba(255,255,255,0.05); padding: 5px 10px; border-radius: 6px;">
                    <span style="color: #E8D5B7; font-size: 0.6rem;">RPT Improved / Worsened</span>
                    <span style="color: #4ECDC4; font-size: 0.8rem; font-weight: 700;">{rpt_improved_count} / {rpt_worsened_count} depots</span>
                </div>
                <div style="display: flex; justify-content: space-between; align-items: center; background: rgba(255,255,255,0.05); padding: 5px 10px; border-radius: 6px;">
                    <span style="color: #E8D5B7; font-size: 0.6rem;">Volume Grown / Declined</span>
                    <span style="color: #4ECDC4; font-size: 0.8rem; font-weight: 700;">{vol_grown_count} / {vol_declined_count} depots</span>
                </div>
                <div style="display: flex; justify-content: space-between; align-items: center; background: rgba(255,255,255,0.05); padding: 5px 10px; border-radius: 6px;">
                    <span style="color: #E8D5B7; font-size: 0.6rem;">Period Coverage</span>
                    <span style="color: #FDCB6E; font-size: 0.8rem; font-weight: 700;">Jan - {current_month}</span>
                </div>
            </div>
        </div>
        """
        st.components.v1.html(summary_html, height=200)
    
    st.divider()
    
    # --- Row 3: RPT Comparison + Volume Comparison ---
    col5, col6 = st.columns(2)
    
    with col5:
        st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.2rem;">📊 RPT Comparison (Jan - {current_month})</p>', unsafe_allow_html=True)
        rpt25_agg = data_2025_filtered.groupby('Depot')['RPT'].mean().reset_index().rename(columns={'RPT': '2025'})
        rpt26_agg = data_2026_filtered.groupby('Depot')['RPT'].mean().reset_index().rename(columns={'RPT': '2026'})
        merged_rpt = pd.merge(rpt25_agg, rpt26_agg, on='Depot', how='outer').fillna(0)
        merged_rpt = merged_rpt.sort_values('2025', ascending=True)
        melted_rpt = merged_rpt.melt(id_vars=['Depot'], value_vars=['2025', '2026'], var_name='Year', value_name='RPT')
        
        fig4 = px.bar(melted_rpt, x='Depot', y='RPT', color='Year', barmode='group',
                      color_discrete_map={'2025': '#F39C12', '2026': '#8B5CF6'},
                      text='RPT')
        fig4.update_traces(texttemplate='₹%{text:,.0f}', textposition='outside', textfont=dict(color='white', size=9))
        fig4.update_layout(height=200, margin=dict(l=0, r=0, t=10, b=10),
                           font=dict(size=9, color=COLORS['white']),
                           plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                           xaxis_title='Depot', yaxis_title='RPT (₹ per 100,000)',
                           legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5, font=dict(size=8, color=COLORS['white'])),
                           yaxis=dict(gridcolor='rgba(255,255,255,0.06)'))
        st.plotly_chart(fig4, use_container_width=True, config={'displayModeBar': False})
    
    with col6:
        st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.2rem;">📊 Volume Comparison (Jan - {current_month})</p>', unsafe_allow_html=True)
        vol25_agg = data_2025_filtered.groupby('Depot')['Vol'].sum().reset_index().rename(columns={'Vol': '2025'})
        vol26_agg = data_2026_filtered.groupby('Depot')['Vol'].sum().reset_index().rename(columns={'Vol': '2026'})
        merged_vol = pd.merge(vol25_agg, vol26_agg, on='Depot', how='outer').fillna(0)
        merged_vol = merged_vol.sort_values('2025', ascending=True)
        melted_vol = merged_vol.melt(id_vars=['Depot'], value_vars=['2025', '2026'], var_name='Year', value_name='Volume')
        
        fig5 = px.bar(melted_vol, x='Depot', y='Volume', color='Year', barmode='group',
                      color_discrete_map={'2025': '#FF6B6B', '2026': '#4ECDC4'},
                      text='Volume')
        fig5.update_traces(texttemplate='%{text:,.0f}', textposition='outside', textfont=dict(color='white', size=9))
        fig5.update_layout(height=200, margin=dict(l=0, r=0, t=10, b=10),
                           font=dict(size=9, color=COLORS['white']),
                           plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                           xaxis_title='Depot', yaxis_title='Volume (MT)',
                           legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5, font=dict(size=8, color=COLORS['white'])),
                           yaxis=dict(gridcolor='rgba(255,255,255,0.06)'))
        st.plotly_chart(fig5, use_container_width=True, config={'displayModeBar': False})
    
    st.divider()
    
    # --- Performance Comparison Table ---
    st.markdown(f'<p style="color:{COLORS["white"]};font-size:0.85rem;font-weight:600;margin-bottom:0.5rem;">📊 Performance Comparison Matrix (Jan - {current_month})</p>', unsafe_allow_html=True)
    
    perf25 = data_2025_filtered.groupby('Depot').agg({'Total_FRT': 'sum', 'Vol': 'sum', 'RPT': 'mean'}).reset_index()
    perf25['Year'] = '2025'
    perf26 = data_2026_filtered.groupby('Depot').agg({'Total_FRT': 'sum', 'Vol': 'sum', 'RPT': 'mean'}).reset_index()
    perf26['Year'] = '2026'
    combined_perf = pd.concat([perf25, perf26], ignore_index=True)
    combined_perf = combined_perf.sort_values(['Depot', 'Year'])
    
    html_table = """
    <style>
        .freight-header { background: linear-gradient(135deg, #c0392b, #e74c3c) !important; color: white !important; }
        .rpt-header { background: linear-gradient(135deg, #d4a017, #f1c40f) !important; color: #1a1a1a !important; }
        .volume-header { background: linear-gradient(135deg, #16a085, #1abc9c) !important; color: white !important; }
        .main-header { background: linear-gradient(135deg, #4a1a6b, #7D4BAE, #9B59B6) !important; }
        
        .row-dark { background-color: #3d1a5e !important; }
        .row-light { background-color: #4d2a6e !important; }
        
        .change-positive { color: #2ecc71 !important; font-weight: 700; }
        .change-negative { color: #e74c3c !important; font-weight: 700; }
        .change-neutral { color: #f1c40f !important; font-weight: 700; }
        
        .freight-value { color: #ff7675 !important; font-weight: 500; }
        .rpt-value { color: #fdcb6e !important; font-weight: 500; }
        .volume-value { color: #55efc4 !important; font-weight: 500; }
    </style>
    <table style="width:100%; border-collapse: collapse; font-family: 'Segoe UI', sans-serif; font-size: 0.7rem; border-radius: 10px; overflow: hidden; box-shadow: 0 8px 32px rgba(0,0,0,0.4);">
        <thead>
            <tr class="main-header">
                <th style="padding: 8px 10px; text-align: left; color: white; font-weight: 600;">Depot</th>
                <th style="padding: 8px 10px; text-align: right; color: white; font-weight: 600;" class="freight-header">2025 Freight</th>
                <th style="padding: 8px 10px; text-align: right; color: white; font-weight: 600;" class="freight-header">2026 Freight</th>
                <th style="padding: 8px 10px; text-align: right; color: white; font-weight: 600;" class="freight-header">Change</th>
                <th style="padding: 8px 10px; text-align: right; color: #1a1a1a; font-weight: 600;" class="rpt-header">2025 RPT</th>
                <th style="padding: 8px 10px; text-align: right; color: #1a1a1a; font-weight: 600;" class="rpt-header">2026 RPT</th>
                <th style="padding: 8px 10px; text-align: right; color: #1a1a1a; font-weight: 600;" class="rpt-header">RPT Change</th>
                <th style="padding: 8px 10px; text-align: right; color: white; font-weight: 600;" class="volume-header">2025 Vol</th>
                <th style="padding: 8px 10px; text-align: right; color: white; font-weight: 600;" class="volume-header">2026 Vol</th>
                <th style="padding: 8px 10px; text-align: right; color: white; font-weight: 600;" class="volume-header">Vol Change</th>
            </tr>
        </thead>
        <tbody>
    """
    
    row_count = 0
    for depot in combined_perf['Depot'].unique():
        df_depot = combined_perf[combined_perf['Depot'] == depot]
        frt25_val = df_depot[df_depot['Year'] == '2025']['Total_FRT'].values[0] if not df_depot[df_depot['Year'] == '2025'].empty else 0
        frt26_val = df_depot[df_depot['Year'] == '2026']['Total_FRT'].values[0] if not df_depot[df_depot['Year'] == '2026'].empty else 0
        rpt25_val = df_depot[df_depot['Year'] == '2025']['RPT'].values[0] if not df_depot[df_depot['Year'] == '2025'].empty else 0
        rpt26_val = df_depot[df_depot['Year'] == '2026']['RPT'].values[0] if not df_depot[df_depot['Year'] == '2026'].empty else 0
        vol25_val = df_depot[df_depot['Year'] == '2025']['Vol'].values[0] if not df_depot[df_depot['Year'] == '2025'].empty else 0
        vol26_val = df_depot[df_depot['Year'] == '2026']['Vol'].values[0] if not df_depot[df_depot['Year'] == '2026'].empty else 0
        
        frt_chg = frt26_val - frt25_val
        rpt_chg = rpt26_val - rpt25_val
        vol_chg = vol26_val - vol25_val
        
        frt_color_class = 'change-positive' if frt_chg < 0 else 'change-negative' if frt_chg > 0 else 'change-neutral'
        rpt_color_class = 'change-positive' if rpt_chg < 0 else 'change-negative' if rpt_chg > 0 else 'change-neutral'
        vol_color_class = 'change-positive' if vol_chg > 0 else 'change-negative' if vol_chg < 0 else 'change-neutral'
        
        row_class = 'row-light' if row_count % 2 == 0 else 'row-dark'
        
        html_table += f"""
            <tr class="{row_class}">
                <td style="padding: 6px 10px; text-align: left; color: white; border-bottom: 1px solid rgba(255,255,255,0.05); font-weight: 600;">{depot}</td>
                <td style="padding: 6px 10px; text-align: right; color: #ff7675; border-bottom: 1px solid rgba(255,255,255,0.05);">₹{frt25_val:,.0f}</td>
                <td style="padding: 6px 10px; text-align: right; color: #ff7675; border-bottom: 1px solid rgba(255,255,255,0.05);">₹{frt26_val:,.0f}</td>
                <td style="padding: 6px 10px; text-align: right; border-bottom: 1px solid rgba(255,255,255,0.05);" class="{frt_color_class}">{frt_chg:+,.0f}</td>
                <td style="padding: 6px 10px; text-align: right; color: #fdcb6e; border-bottom: 1px solid rgba(255,255,255,0.05);">{rpt25_val:,.1f}</td>
                <td style="padding: 6px 10px; text-align: right; color: #fdcb6e; border-bottom: 1px solid rgba(255,255,255,0.05);">{rpt26_val:,.1f}</td>
                <td style="padding: 6px 10px; text-align: right; border-bottom: 1px solid rgba(255,255,255,0.05);" class="{rpt_color_class}">{rpt_chg:+,.1f}</td>
                <td style="padding: 6px 10px; text-align: right; color: #55efc4; border-bottom: 1px solid rgba(255,255,255,0.05);">{vol25_val:,.0f}</td>
                <td style="padding: 6px 10px; text-align: right; color: #55efc4; border-bottom: 1px solid rgba(255,255,255,0.05);">{vol26_val:,.0f}</td>
                <td style="padding: 6px 10px; text-align: right; border-bottom: 1px solid rgba(255,255,255,0.05);" class="{vol_color_class}">{vol_chg:+,.0f}</td>
            </tr>
        """
        row_count += 1
    
    html_table += """
        </tbody>
    </table>
    """
    st.components.v1.html(html_table, height=300, scrolling=False)
    
    st.markdown(f"""
    <div style="display: flex; gap: 20px; margin-top: 10px; flex-wrap: wrap; font-size: 0.65rem; padding: 6px 4px; background: rgba(255,255,255,0.03); border-radius: 8px;">
        <span style="color: #ff7675; font-weight: 600;">🔴 Freight (↓ Good)</span>
        <span style="color: #fdcb6e; font-weight: 600;">🟡 RPT (↓ Good)</span>
        <span style="color: #55efc4; font-weight: 600;">🟢 Volume (↑ Good)</span>
        <span style="color: {COLORS['light']}; font-weight: 400;">💡 Comparing Jan - {current_month} for both years</span>
    </div>
    """, unsafe_allow_html=True)

# ---------- Main function ----------
def apply_custom_css():
    st.markdown("""
    <style>
        .stApp { background-color: #380b3a; }
        .main > div { background-color: #380b3a; }
        
        /* Sidebar Styling */
        section[data-testid="stSidebar"] {
            background: linear-gradient(135deg, #2d0a2f, #380b3a, #4a1a6b) !important;
            border-right: 1px solid rgba(255, 255, 255, 0.1) !important;
        }
        
        section[data-testid="stSidebar"] .stMarkdown {
            color: #FFFFFF !important;
        }
        
        section[data-testid="stSidebar"] .stRadio label {
            color: #FFFFFF !important;
            font-size: 0.75rem !important;
        }
        
        section[data-testid="stSidebar"] .stSelectbox label {
            color: #FFFFFF !important;
            font-size: 0.75rem !important;
        }
        
        section[data-testid="stSidebar"] .stSelectbox div {
            font-size: 0.8rem !important;
            color: #FFFFFF !important;
        }
        
        section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] {
            background-color: #5d2d8a !important;
            border: 1px solid rgba(255, 255, 255, 0.2) !important;
            color: #FFFFFF !important;
            min-height: 2.2rem !important;
            padding: 0 0.5rem !important;
            border-radius: 6px !important;
        }
        
        section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] * {
            color: #FFFFFF !important;
        }
        
        section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] {
            gap: 0.5rem !important;
        }
        
        section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
            background-color: rgba(255, 255, 255, 0.05) !important;
            padding: 0.3rem 0.8rem !important;
            border-radius: 6px !important;
            border: 1px solid rgba(255, 255, 255, 0.1) !important;
        }
        
        section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label[data-selected="true"] {
            background-color: #7D4BAE !important;
            border-color: rgba(255, 255, 255, 0.3) !important;
        }
        
        /* Remove sticky header since controls are now in sidebar */
        .sticky-header {
            display: none;
        }
        
        h1, h2, h3 { color: #FFFFFF !important; font-weight: 700 !important; }
        
        .stMetric {
            background: linear-gradient(135deg, #380b3a, #4d1d7a) !important;
            padding: 0.5rem !important;
            border-radius: 10px !important;
            border: 1px solid rgba(255, 255, 255, 0.15) !important;
            box-shadow: 0 4px 15px rgba(0,0,0,0.3), inset 0 1px 0 rgba(255,255,255,0.05) !important;
        }
        
        .stMetric > div { font-size: 0.8rem !important; color: #FFFFFF !important; }
        .stMetric label { font-size: 0.65rem !important; font-weight: 600 !important; text-transform: uppercase; letter-spacing: 0.5px; color: #E8D5B7 !important; }
        .stMetric .st-emotion-cache-10trblm { font-size: 1.2rem !important; font-weight: 700 !important; color: #FFFFFF !important; }
        
        .block-container { 
            padding-top: 4rem !important; 
            padding-bottom: 0.1rem !important; 
            padding-left: 1rem !important; 
            padding-right: 1rem !important; 
        }
        
        .stDivider { margin: 0.1rem 0 !important; border-color: rgba(255, 255, 255, 0.08) !important; }
    </style>
    """, unsafe_allow_html=True)

def main():
    st.set_page_config(page_title="Depot Wise Freight Cost Dashboard", layout="wide")
    apply_custom_css()
    
    # --- Sidebar ---
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center; padding: 0.5rem 0;">
            <h1 style="color: #FFFFFF; font-size: 1.8rem; font-weight: 700; letter-spacing: 2px; text-shadow: 0 2px 4px rgba(0,0,0,0.3); margin: 0;">
             Depot Wise Freight Cost Dashboard
            </h1>
        </div>
        """, unsafe_allow_html=True)
        
        st.divider()
        
        pattern = re.compile(r"Master file for Cost working - South 202[56]\.xlsx")
        files = [f for f in os.listdir(".") if pattern.match(f)]
        if not files:
            st.error("❌ No Excel files found. Please ensure the 2025/2026 files are in the current directory.")
            return
        
        avail_years = sorted([int(re.search(r"202[56]", f).group()) for f in files])
        
        data_cache = {}
        for y in avail_years:
            d, _ = load_data(y)
            if d is not None:
                data_cache[y] = d
        
        # Dashboard Type Selection
        st.markdown('<p style="color:#E8D5B7;font-size:0.75rem;font-weight:600;margin-bottom:0.3rem;"> Dashboard Type</p>', unsafe_allow_html=True)
        dash_type = st.radio("", ["Yearly", "Comparison"], horizontal=True, key="dash_type", label_visibility="collapsed")
        
        st.divider()
        
        # Year/Month Selection
        if dash_type == "Yearly":
            st.markdown('<p style="color:#E8D5B7;font-size:0.75rem;font-weight:600;margin-bottom:0.3rem;">📅 Select Year</p>', unsafe_allow_html=True)
            year_options = [str(y) for y in avail_years]
            selected_year_str = st.selectbox("", year_options, key="year_selector", label_visibility="collapsed")
            selected_year = int(selected_year_str)
            
            if selected_year in data_cache:
                avail_months = data_cache[selected_year]['Month'].unique()
                avail_months_sorted = sorted(avail_months, key=lambda x: MONTHS_ORDER.index(x) if x in MONTHS_ORDER else 99)
                month_options = ["All"] + avail_months_sorted
                if f"month_{selected_year}" not in st.session_state:
                    st.session_state[f"month_{selected_year}"] = "All"
                
                st.markdown('<p style="color:#E8D5B7;font-size:0.75rem;font-weight:600;margin-bottom:0.3rem;margin-top:0.8rem;">📆 Select Month</p>', unsafe_allow_html=True)
                selected_month = st.selectbox("", month_options, key=f"month_{selected_year}", label_visibility="collapsed")
            else:
                selected_month = "All"
        else:
            current_month = get_current_month()
            st.markdown(f"""
            <div style="background: rgba(255,255,255,0.05); border-radius: 8px; padding: 10px; margin-top: 10px;">
                <p style="color:#E8D5B7;font-size:0.75rem;font-weight:600;margin-bottom:0.3rem;">📅 Comparison Period</p>
                <p style="color:#FDCB6E;font-size:0.85rem;font-weight:700;margin:0;">2025 vs 2026</p>
                <p style="color:#E8D5B7;font-size:0.7rem;margin:0;">Jan - {current_month}</p>
            </div>
            """, unsafe_allow_html=True)
            selected_year = None
            selected_month = None
        
        st.divider()
        
        # Info section
        st.markdown("""
        <div style="background: rgba(255,255,255,0.03); border-radius: 8px; padding: 10px; margin-top: 10px;">
            <p style="color:#E8D5B7;font-size:0.7rem;font-weight:600;margin-bottom:0.3rem;">📌 Key Metrics</p>
            <p style="color:#E8D5B7;font-size:0.65rem;margin:0;line-height:1.4;">
                • <b>RPT</b>: Rate Per Tonne (₹ per 100,000)<br>
                • <b>Freight Cost</b>: Total transportation cost<br>
                • <b>Volume</b>: Net weight in Metric Tonnes
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    # --- Main Content Area ---
    # Remove the original title from main area
    # st.markdown(""" ... """) - This was the old title, now moved to sidebar
    
    if dash_type == "Yearly":
        if selected_year in data_cache:
            create_year_dashboard(data_cache[selected_year], selected_year, selected_month)
        else:
            st.error(f"Data for {selected_year} not loaded.")
    else:
        if 2025 in data_cache and 2026 in data_cache:
            create_comparison_dashboard(data_cache[2025], data_cache[2026])
        else:
            st.error("Both 2025 and 2026 data files are required for comparison.")

if __name__ == "__main__":
    main()
# file name: app_streamlit_animated.py
# Factory Machine Tool Precision Comparison Analysis System - Streamlit Version with Animations

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import os
import time
import traceback
from datetime import datetime, timedelta
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.lines import Line2D

# Page configuration
st.set_page_config(
    page_title="Machine Tool Precision Analysis",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# Color Theme: Pastel / Macaron (浅色半透明主题)
# ==========================================
THEME_BLUE = '#9BB0E2'    # 柔和长春花蓝 (Soft Periwinkle Blue)
THEME_ORANGE = '#F6B79D'  # 柔和蜜桃粉橘 (Soft Peach Orange)
THEME_GREEN = '#A2D5AB'   # 柔和薄荷绿 (Soft Mint Green)
THEME_RED = '#F4A4A4'     # 柔和珊瑚红 (Soft Coral Red)
THEME_GRAY = '#C5C9D1'    # 浅灰色 (Soft Gray)

# Set font
try:
    plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'DejaVu Sans', 'sans-serif']
except:
    plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False


# ==========================================
# Custom CSS for Animations
# ==========================================

def add_custom_css():
    """Add custom CSS animations with updated pastel colors"""
    st.markdown("""
    <style>
    @keyframes fadeInUp {
        from { opacity: 0; transform: translateY(30px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    @keyframes fadeInLeft {
        from { opacity: 0; transform: translateX(-30px); }
        to { opacity: 1; transform: translateX(0); }
    }
    
    @keyframes fadeInRight {
        from { opacity: 0; transform: translateX(30px); }
        to { opacity: 1; transform: translateX(0); }
    }
    
    @keyframes scaleIn {
        from { opacity: 0; transform: scale(0.9); }
        to { opacity: 1; transform: scale(1); }
    }
    
    @keyframes shimmer {
        0% { background-position: -1000px 0; }
        100% { background-position: 1000px 0; }
    }
    
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }
    
    .animate-fade-in-up { animation: fadeInUp 0.6s cubic-bezier(0.2, 0.8, 0.4, 1) forwards; }
    .animate-fade-in-left { animation: fadeInLeft 0.5s ease-out forwards; }
    .animate-fade-in-right { animation: fadeInRight 0.5s ease-out forwards; }
    .animate-scale-in { animation: scaleIn 0.4s ease-out forwards; }
    
    .shimmer-loading {
        background: linear-gradient(90deg, #f8f9fa 25%, #e9ecef 50%, #f8f9fa 75%);
        background-size: 1000px 100%;
        animation: shimmer 1.5s infinite;
        border-radius: 12px;
    }
    
    .spinner {
        width: 50px;
        height: 50px;
        border: 3px solid #f0f2f5;
        border-top: 3px solid #9BB0E2; /* Pastel Blue */
        border-radius: 50%;
        animation: spin 1s linear infinite;
        margin: 20px auto;
    }
    
    .pulse-animation { animation: pulse 2s ease-in-out infinite; }
    
    .metric-card {
        background: white;
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03);
        text-align: center;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 20px rgba(155, 176, 226, 0.15);
    }
    
    .stButton > button {
        transition: all 0.3s ease !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 4px 12px rgba(155, 176, 226, 0.4) !important;
    }
    
    .chart-container {
        animation: fadeInUp 0.6s ease-out;
    }
    
    /* Progress bar animation (Pastel Gradient) */
    .stProgress > div > div {
        background: linear-gradient(90deg, #9BB0E2, #A2D5AB, #9BB0E2);
        background-size: 200% 100%;
        animation: shimmer 2s infinite;
    }
    
    /* Custom scrollbar */
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 10px; }
    ::-webkit-scrollbar-thumb { background: #cbd0d6; border-radius: 10px; }
    ::-webkit-scrollbar-thumb:hover { background: #a8aeb5; }
    </style>
    """, unsafe_allow_html=True)


# ==========================================
# Animated Components
# ==========================================

def show_loading_spinner(message="Loading..."):
    spinner_html = f"""
    <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 40px;">
        <div class="spinner"></div>
        <p style="margin-top: 20px; color: #888; font-size: 14px;">{message}</p>
    </div>
    """
    return st.markdown(spinner_html, unsafe_allow_html=True)

def show_shimmer_loading(height=400):
    shimmer_html = f"""
    <div class="shimmer-loading" style="height: {height}px; width: 100%; border-radius: 12px;">
    </div>
    """
    return st.markdown(shimmer_html, unsafe_allow_html=True)

def display_animated_metric(label, value, delta=None, animation_delay=0):
    delay_style = f"animation-delay: {animation_delay}s;"
    html = f"""
    <div class="metric-card animate-fade-in-up" style="{delay_style}">
        <div style="color: #9da3af; font-size: 14px; margin-bottom: 8px;">{label}</div>
        <div style="font-size: 36px; font-weight: bold; color: #3d4451;">{value}</div>
        {f'<div style="color: #A2D5AB; font-size: 13px; margin-top: 8px;">{delta}</div>' if delta else ''}
    </div>
    """
    return st.markdown(html, unsafe_allow_html=True)

def display_animated_chart(img_bytes, title, chart_index=0):
    delay = min(chart_index * 0.1, 0.5)
    html = f"""
    <div class="chart-container animate-fade-in-up" style="animation-delay: {delay}s;">
        <h3 style="margin-bottom: 15px; color: #3d4451;">{title}</h3>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)
    st.image(img_bytes, use_container_width=True)

def show_completion_message():
    st.markdown("""
    <div style="
        background: linear-gradient(135deg, #A2D5AB 0%, #8CC296 100%);
        border-radius: 12px;
        padding: 15px;
        text-align: center;
        margin: 20px 0;
    ">
        <div style="color: white; font-size: 16px; font-weight: bold;">✅ Analysis Completed Successfully!</div>
    </div>
    """, unsafe_allow_html=True)
    time.sleep(0.3)


# ==========================================
# Utility & Data Functions
# ==========================================

def extract_factory_name(file_name):
    base_name = os.path.basename(file_name)
    name_without_ext = os.path.splitext(base_name)[0]
    if len(name_without_ext) > 30:
        name_without_ext = name_without_ext[:30]
    return name_without_ext

def get_cnc_column_name(df):
    potential_cnc_cols = ['CNC OP', 'cnc op', 'Cnc Op', 'cnc station', 'CNC Station', 'Station']
    for col in df.columns:
        for potential in potential_cnc_cols:
            if potential.lower() == col.lower():
                return col
    return None

def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', transparent=True) # 透明背景
    buf.seek(0)
    plt.close(fig)
    return buf.getvalue()

def load_excel_data(file_content, factory_name, read_mode='default'):
    try:
        df = None
        if read_mode == 'ipeg':
            fanuc_sheet_name = 'Machine OQ Rev-Fanuc'
            jd_sheet_name = 'Machine OQ Rev-JD'
            header_row = 3
            dfs_to_merge = []
            
            try:
                df_fanuc = pd.read_excel(io.BytesIO(file_content), sheet_name=fanuc_sheet_name, header=header_row)
                if df_fanuc is not None and not df_fanuc.empty:
                    df_fanuc = df_fanuc.dropna(how='all')
                    first_col = df_fanuc.columns[0]
                    df_fanuc = df_fanuc[df_fanuc[first_col].notna()]
                    if not df_fanuc.empty:
                        dfs_to_merge.append(df_fanuc)
            except Exception as e: pass
            
            try:
                df_jd = pd.read_excel(io.BytesIO(file_content), sheet_name=jd_sheet_name, header=header_row)
                if df_jd is not None and not df_jd.empty:
                    df_jd = df_jd.dropna(how='all')
                    first_col = df_jd.columns[0]
                    df_jd = df_jd[df_jd[first_col].notna()]
                    if not df_jd.empty:
                        dfs_to_merge.append(df_jd)
            except Exception as e: pass
            
            if not dfs_to_merge:
                raise ValueError("No data found in either Fanuc or JD sheets")
            
            df = pd.concat(dfs_to_merge, ignore_index=True)
        else:
            try:
                df = pd.read_excel(io.BytesIO(file_content), sheet_name='Table', header=2)
            except:
                try:
                    df = pd.read_excel(io.BytesIO(file_content), sheet_name='Table', header=1)
                except:
                    df = pd.read_excel(io.BytesIO(file_content))
        
        if df is None or df.empty:
            raise ValueError("Could not load any data from the Excel file")
        
        column_mapping = {}
        for col in df.columns:
            col_clean = str(col).replace('\n', ' ').replace('（', '(').replace('）', ')').strip()
            if 'Station' in col_clean or '夹位' in col_clean:
                column_mapping[col] = 'CNC OP'
            elif 'Model' in col_clean and 'Probe' not in col_clean:
                column_mapping[col] = 'Machine Model'
            elif 'Date of Manufacturer' in col_clean or '設備製造日期' in col_clean:
                column_mapping[col] = 'Year of manufacturer'
        
        if column_mapping:
            df = df.rename(columns=column_mapping)
        
        required_columns = ['CNC OP', 'Machine Model', 'Year of manufacturer']
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise ValueError(f"Required columns not found: {missing}")
        
        df['CNC OP'] = df['CNC OP'].astype(str)
        df['Machine Model'] = df['Machine Model'].astype(str)
        df = df[df['CNC OP'].notna() & (df['CNC OP'] != 'nan') & (df['CNC OP'] != '')]
        df = df[df['Machine Model'].notna() & (df['Machine Model'] != 'nan') & (df['Machine Model'] != '')]
        
        def safe_extract_year(val):
            if pd.isna(val): return None
            try:
                if isinstance(val, (int, float)):
                    num_val = int(val)
                elif isinstance(val, str) and val.replace('.', '').isdigit():
                    num_val = int(float(val))
                else:
                    num_val = None
                
                if num_val and 30000 <= num_val <= 50000:
                    excel_base = datetime(1899, 12, 30)
                    date_val = excel_base + timedelta(days=num_val)
                    year = date_val.year
                    current_year = datetime.now().year
                    if 1980 <= year <= current_year + 5: return year
            except: pass
            
            str_val = str(val)
            match = re.search(r'(19|20)\d{2}', str_val)
            if match:
                year = int(match.group(0))
                current_year = datetime.now().year
                if 1980 <= year <= current_year + 5: return year
            return None

        df['Year_of_manufacturer'] = df['Year of manufacturer'].apply(safe_extract_year)
        df = df.dropna(subset=['Year_of_manufacturer'])
        if df.empty: raise ValueError("No valid year data found after filtering")
        df['Year_of_manufacturer'] = df['Year_of_manufacturer'].astype(int)
        df['Factory'] = factory_name
        
        return df, None
        
    except Exception as e:
        traceback.print_exc()
        return None, str(e)


# ==========================================
# Chart Generation Functions (Updated with Pastel Theme)
# ==========================================

def compare_machine_count(df1, df2, name1, name2):
    cnc_col = get_cnc_column_name(df1)
    fig, ax = plt.subplots(figsize=(14, 8))
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')
    
    if cnc_col is None:
        ax.text(0.5, 0.5, "CNC Station column not found", ha='center', va='center', fontsize=14, color=THEME_GRAY)
        return fig_to_bytes(fig)
    
    def get_detailed_counts(df):
        cnc_col_actual = get_cnc_column_name(df)
        if cnc_col_actual is None: return pd.DataFrame()
        counts_df = df.groupby(cnc_col_actual)['Machine Model'].nunique().reset_index()
        counts_df.columns = ['Station', 'Machine_Count']
        models_df = df.groupby(cnc_col_actual)['Machine Model'].apply(lambda x: sorted(x.unique())).reset_index()
        models_df.columns = ['Station', 'Machine_Models']
        return counts_df.merge(models_df, on='Station')
    
    counts1 = get_detailed_counts(df1)
    counts2 = get_detailed_counts(df2)
    
    all_stations = sorted(set(counts1['Station']).union(set(counts2['Station'])))
    compare_df = pd.DataFrame({'Station': all_stations})
    compare_df = compare_df.merge(counts1[['Station', 'Machine_Count', 'Machine_Models']], on='Station', how='left').fillna(0)
    compare_df = compare_df.merge(counts2[['Station', 'Machine_Count', 'Machine_Models']], on='Station', how='left').fillna(0)
    compare_df.columns = ['Station', f'{name1}_Count', f'{name1}_Models', f'{name2}_Count', f'{name2}_Models']
    
    compare_df['Total'] = compare_df[f'{name1}_Count'] + compare_df[f'{name2}_Count']
    compare_df = compare_df.sort_values('Total', ascending=False).head(15)
    
    x = np.arange(len(compare_df['Station']))
    width = 0.35
    
    # Updated: Pastel Colors + High Transparency
    ax.bar(x - width/2, compare_df[f'{name1}_Count'], width, label=name1, color=THEME_BLUE, alpha=0.6, edgecolor='white', linewidth=1)
    ax.bar(x + width/2, compare_df[f'{name2}_Count'], width, label=name2, color=THEME_ORANGE, alpha=0.6, edgecolor='white', linewidth=1)
    
    ax.set_xlabel('CNC Station', fontsize=12, color='#444')
    ax.set_ylabel('Number of Machine Types', fontsize=12, color='#444')
    ax.set_title(f'Machine Type Count by Station\n{name1} vs {name2}', fontsize=14, color='#333')
    ax.set_xticks(x)
    ax.set_xticklabels(compare_df['Station'], rotation=45, ha='right', fontsize=9, color='#555')
    ax.legend(frameon=True, facecolor='white', edgecolor='none')
    ax.grid(True, alpha=0.2, axis='y', color=THEME_GRAY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    for spine in ax.spines.values(): spine.set_edgecolor(THEME_GRAY)
    
    plt.tight_layout()
    return fig_to_bytes(fig)


def compare_machine_age(df1, df2, name1, name2):
    current_year = datetime.now().year
    df1['Machine_Age'] = current_year - df1['Year_of_manufacturer']
    df2['Machine_Age'] = current_year - df2['Year_of_manufacturer']
    
    ages1 = df1['Machine_Age'].dropna()
    ages2 = df2['Machine_Age'].dropna()
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    fig.patch.set_facecolor('none')
    for ax in [ax1, ax2]: ax.set_facecolor('none')
    
    max_age = int(max(ages1.max() if len(ages1) > 0 else 20, ages2.max() if len(ages2) > 0 else 20))
    bins = np.arange(0, max_age + 2, 1)
    
    if len(ages1) > 0 and len(ages2) > 0:
        bp = ax1.boxplot([ages1, ages2], tick_labels=[name1, name2], patch_artist=True)
        colors = [THEME_BLUE, THEME_ORANGE]
        for patch, color in zip(bp['boxes'], colors):
            patch.set_facecolor(color)
            patch.set_alpha(0.55) # Lower alpha for boxplot
            patch.set_edgecolor('#ffffff')
            patch.set_linewidth(1.5)
        for median in bp['medians']:
            median.set_color('#ffffff')
            median.set_linewidth(2)
        for whisker in bp['whiskers']: whisker.set_color(THEME_GRAY)
        for cap in bp['caps']: cap.set_color(THEME_GRAY)
    else:
        ax1.text(0.5, 0.5, "Insufficient data", ha='center', va='center', fontsize=14, color=THEME_GRAY)
    
    ax1.set_title('Machine Age Distribution', fontsize=14, color='#333')
    ax1.set_ylabel('Age (Years)', color='#444')
    ax1.tick_params(colors='#555')
    
    # Updated: Pastel Colors + Overlap Transparency
    ax2.hist(ages1, bins=bins, alpha=0.55, label=name1, color=THEME_BLUE, edgecolor='white', linewidth=1)
    ax2.hist(ages2, bins=bins, alpha=0.55, label=name2, color=THEME_ORANGE, edgecolor='white', linewidth=1)
    ax2.set_title('Age Distribution Histogram', fontsize=14, color='#333')
    ax2.set_xlabel('Age (Years)', color='#444')
    ax2.set_ylabel('Frequency', color='#444')
    ax2.tick_params(colors='#555')
    ax2.legend(frameon=True, facecolor='white', edgecolor='none')
    
    for ax in [ax1, ax2]:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, alpha=0.2, axis='y', color=THEME_GRAY)
        for spine in ax.spines.values(): spine.set_edgecolor(THEME_GRAY)
    
    plt.tight_layout()
    return fig_to_bytes(fig)


def extract_spindle_runout_universal(df, position='near'):
    values, used_cols = [], []
    position_patterns = [r'@5mm', r'@5\s*mm', r'5mm', r'5\s*mm', r'near'] if position == 'near' else [r'@300mm', r'@300\s*mm', r'300mm', r'@150mm', r'150mm', r'far']
    exclude_patterns = [r'@300', r'300mm', r'150mm', r'@150', r'far'] if position == 'near' else [r'@5mm', r'5mm', r'near']
    
    for col in df.columns:
        col_str = str(col).lower()
        if 'runout' not in col_str and '跳动' not in col_str: continue
        
        is_position_match = any(re.search(p, col_str, re.IGNORECASE) for p in position_patterns)
        if position == 'near' and not is_position_match and ('spindle nose' in col_str or '主軸' in col_str):
            if not any(re.search(e, col_str, re.IGNORECASE) for e in exclude_patterns):
                is_position_match = True
                
        if not is_position_match or any(re.search(e, col_str, re.IGNORECASE) for e in exclude_patterns): continue
        
        vals = pd.to_numeric(df[col], errors='coerce').dropna()
        if len(vals) == 0: continue
        
        if any(x in col for x in ['[µm]', '[μm]', 'micron', 'um]']) or ('mm' not in col):
            vals = vals / 1000
            
        values.extend(vals.tolist())
        used_cols.append(col)
    return values, used_cols


def compare_spindle_runout(df1, df2, name1, name2):
    near1, _ = extract_spindle_runout_universal(df1, 'near')
    near2, _ = extract_spindle_runout_universal(df2, 'near')
    far1, _ = extract_spindle_runout_universal(df1, 'far')
    far2, _ = extract_spindle_runout_universal(df2, 'far')
    
    USL_NEAR_MM = 0.006
    USL_FAR_MM = 0.030
    
    has_near_data = len(near1) > 0 or len(near2) > 0
    has_far_data = len(far1) > 0 or len(far2) > 0
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 9))
    fig.patch.set_facecolor('none')
    for ax in [ax1, ax2]: ax.set_facecolor('none')
    
    if not has_near_data and not has_far_data:
        ax1.text(0.5, 0.5, "Spindle runout data not found", ha='center', va='center', fontsize=12, color=THEME_GRAY)
        plt.tight_layout()
        return fig_to_bytes(fig)
    
    if has_near_data:
        _plot_runout_distribution(ax1, near1, near2, name1, name2, USL_NEAR_MM, 'Near End (5mm / Spindle Nose)\nSpec: <=6um')
    else:
        ax1.text(0.5, 0.5, "Near end data not found", ha='center', va='center', fontsize=11, color=THEME_GRAY)
        ax1.set_title('Near End Runout - No Data', fontsize=12, color='#333')
    
    if has_far_data:
        _plot_runout_distribution(ax2, far1, far2, name1, name2, USL_FAR_MM, 'Far End (300mm from Spindle Nose)\nSpec: <=30um')
    else:
        ax2.text(0.5, 0.5, "Far end data not found", ha='center', va='center', fontsize=11, color=THEME_GRAY)
        ax2.set_title('Far End Runout - No Data', fontsize=12, color='#333')
    
    fig.suptitle(f'Spindle Runout Comparison\n{name1} vs {name2}', fontsize=14, fontweight='bold', color='#333')
    plt.tight_layout()
    return fig_to_bytes(fig)


def _plot_runout_distribution(ax, data1, data2, name1, name2, usl_mm, title):
    def calc_stats(data, label):
        if len(data) == 0: return None
        mean = np.mean(data)
        std = np.std(data, ddof=1) if len(data) > 1 else 0.001
        std = 0.001 if std == 0 else std
        cpk = (usl_mm - mean) / (3 * std) if usl_mm > mean else 0
        return {'n': len(data), 'mean': mean, 'std': std, 'cpk': cpk, 'label': label}
    
    stats1 = calc_stats(data1, name1)
    stats2 = calc_stats(data2, name2)
    all_data = data1 + data2
    
    if not all_data:
        ax.text(0.5, 0.5, "No valid data", ha='center', va='center', fontsize=12)
        ax.set_title(title, color='#333')
        return
    
    x_min = min(0, min(all_data) * 0.9)
    x_max = max(usl_mm * 1.5, max(all_data) * 1.2)
    x_range = np.linspace(x_min, x_max, 200)
    
    # Updated: Transparent Pastels for overlaps
    if stats1:
        y1 = stats.norm.pdf(x_range, stats1['mean'], stats1['std'])
        ax.plot(x_range, y1, '-', linewidth=2.5, label=f'{name1} (Fit)', color=THEME_BLUE)
        ax.hist(data1, bins=15, density=True, alpha=0.45, color=THEME_BLUE, edgecolor='white')
    
    if stats2:
        y2 = stats.norm.pdf(x_range, stats2['mean'], stats2['std'])
        ax.plot(x_range, y2, '-', linewidth=2.5, label=f'{name2} (Fit)', color=THEME_ORANGE)
        ax.hist(data2, bins=15, density=True, alpha=0.45, color=THEME_ORANGE, edgecolor='white')
    
    ax.axvline(x=usl_mm, color=THEME_RED, linestyle='--', linewidth=2, label=f'USL: {usl_mm*1000:.0f}um')
    ax.axvline(x=usl_mm * 0.5, color=THEME_GREEN, linestyle=':', linewidth=2, label=f'Target: {usl_mm*500:.0f}um')
    
    ax.set_xlabel('Runout (mm)', fontsize=11, color='#444')
    ax.set_ylabel('Probability Density', fontsize=11, color='#444')
    ax.set_title(title, fontsize=11, color='#333')
    ax.tick_params(colors='#555')
    ax.legend(loc='upper right', fontsize=8, frameon=True, facecolor='white', edgecolor='none')
    ax.set_xlim(x_min, x_max)
    ax.grid(True, alpha=0.2, color=THEME_GRAY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    for spine in ax.spines.values(): spine.set_edgecolor(THEME_GRAY)


def compare_spindle_velocity(df1, df2, name1, name2):
    target_stations = ['CNC4', 'CNC4.1', 'CNC5', 'CNC6', 'CNC7', 'CNC7.2', 'CNC8']
    cnc_col = get_cnc_column_name(df1)
    
    fig, ax = plt.subplots(figsize=(14, 7))
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')
    
    if cnc_col is None:
        ax.text(0.5, 0.5, "CNC Station data not found", ha='center', va='center', fontsize=14, color=THEME_GRAY)
        return fig_to_bytes(fig)
    
    def extract_velocity_data(df, read_mode='default'):
        cnc_col_actual = get_cnc_column_name(df)
        if cnc_col_actual is None: return {}
        df_f = df[df[cnc_col_actual].isin(target_stations)]
        if df_f.empty: return {}
        
        extracted_data = {}
        rpm_priority = [18000] if read_mode == 'ipeg' else [18000, 16000, 10000]
        
        for station in df_f[cnc_col_actual].unique():
            station_df = df_f[df_f[cnc_col_actual] == station]
            selected_col, selected_rpm = None, None
            
            for target_rpm in rpm_priority:
                for col in station_df.columns:
                    col_str = str(col).lower()
                    if ('velocity' in col_str or '振动速度' in str(col)) and ('spindle' in col_str or '主轴' in str(col)) and str(target_rpm) in str(col):
                        vals = pd.to_numeric(station_df[col], errors='coerce').dropna()
                        if len(vals) > 0:
                            selected_col, selected_rpm = col, target_rpm
                            break
                if selected_col: break
            
            if selected_col:
                vals = pd.to_numeric(station_df[selected_col], errors='coerce').dropna()
                if len(vals) > 0:
                    extracted_data[station] = {'mean': np.mean(vals), 'std': np.std(vals) if len(vals) > 1 else 0.001, 'n': len(vals), 'rpm': selected_rpm}
        return extracted_data
    
    data1 = extract_velocity_data(df1, read_mode='ipeg' if name1.lower().startswith('ipeg') else 'default')
    data2 = extract_velocity_data(df2, read_mode='ipeg' if name2.lower().startswith('ipeg') else 'default')
    
    if not data1 and not data2:
        ax.text(0.5, 0.5, "Velocity data not found", ha='center', va='center', fontsize=14, color=THEME_GRAY)
        return fig_to_bytes(fig)
    
    all_stations = sorted(set(data1.keys()) | set(data2.keys()))
    x = np.arange(len(all_stations))
    width = 0.35
    
    means1 = [data1.get(s, {}).get('mean', 0) for s in all_stations]
    stds1 = [data1.get(s, {}).get('std', 0) for s in all_stations]
    means2 = [data2.get(s, {}).get('mean', 0) for s in all_stations]
    stds2 = [data2.get(s, {}).get('std', 0) for s in all_stations]
    n1 = [data1.get(s, {}).get('n', 0) for s in all_stations]
    n2 = [data2.get(s, {}).get('n', 0) for s in all_stations]
    
    # Updated: Pastel Colors + Soft Alphas
    ax.bar(x - width/2, means1, width, yerr=stds1, capsize=4, label=f'{name1} (n={sum(n1)})', color=THEME_BLUE, alpha=0.55, edgecolor='white')
    ax.bar(x + width/2, means2, width, yerr=stds2, capsize=4, label=f'{name2} (n={sum(n2)})', color=THEME_ORANGE, alpha=0.55, edgecolor='white')
    
    spec_a, spec_b = 1.1, 1.4
    ax.axhline(y=spec_a, color=THEME_GREEN, linestyle='--', linewidth=2, label=f'Grade A: <{spec_a} mm/s')
    ax.axhline(y=spec_b, color=THEME_ORANGE, linestyle='--', linewidth=2, label=f'Grade B: <{spec_b} mm/s')
    
    y_max = max(max(means1 + stds1, default=0), max(means2 + stds2, default=0), spec_b * 1.2)
    ax.fill_between([-0.5, len(all_stations) - 0.5], spec_b, y_max, color=THEME_RED, alpha=0.1) # Softer red fill
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    for spine in ax.spines.values(): spine.set_edgecolor(THEME_GRAY)
    ax.grid(True, alpha=0.2, axis='y', color=THEME_GRAY)
    
    ax.set_xlabel('CNC Station', fontsize=12, color='#444')
    ax.set_ylabel('Velocity (mm/s)', fontsize=12, color='#444')
    ax.set_title(f'Spindle Velocity Comparison\n{name1} vs {name2}', fontsize=14, color='#333')
    ax.set_xticks(x)
    ax.set_xticklabels(all_stations, rotation=45, ha='right', fontsize=9, color='#555')
    ax.tick_params(colors='#555')
    ax.set_ylim(0, y_max * 1.1)
    ax.legend(loc='upper left', fontsize=10, frameon=True, facecolor='white', edgecolor='none')
    
    plt.tight_layout()
    return fig_to_bytes(fig)


def compare_spindle_acceleration(df1, df2, name1, name2):
    target_stations = ['CNC4', 'CNC4.1', 'CNC5', 'CNC6', 'CNC7', 'CNC7.2', 'CNC8']
    cnc_col = get_cnc_column_name(df1)
    
    fig, ax = plt.subplots(figsize=(14, 7))
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')
    
    if cnc_col is None:
        ax.text(0.5, 0.5, "CNC Station data not found", ha='center', va='center', fontsize=14, color=THEME_GRAY)
        return fig_to_bytes(fig)
    
    def extract_data(df):
        cnc_col_actual = get_cnc_column_name(df)
        if cnc_col_actual is None: return {}
        df_f = df[df[cnc_col_actual].isin(target_stations)]
        acc_cols = [c for c in df_f.columns if 'Acceleration' in str(c) and 'Spindle' in str(c)]
        data = {}
        for station in df_f[cnc_col_actual].unique():
            station_df = df_f[df_f[cnc_col_actual] == station]
            vals = []
            for col in acc_cols:
                vals.extend(pd.to_numeric(station_df[col], errors='coerce').dropna().tolist())
            if vals: data[station] = {'mean': np.mean(vals), 'std': np.std(vals) if len(vals) > 1 else 0}
        return data
    
    data1 = extract_data(df1)
    data2 = extract_data(df2)
    
    if not data1 and not data2:
        ax.text(0.5, 0.5, "Acceleration data not found", ha='center', va='center', fontsize=14, color=THEME_GRAY)
        return fig_to_bytes(fig)
    
    all_stations = sorted(set(data1.keys()) | set(data2.keys()))
    x = np.arange(len(all_stations))
    width = 0.35
    
    means1 = [data1.get(s, {}).get('mean', 0) for s in all_stations]
    stds1 = [data1.get(s, {}).get('std', 0) for s in all_stations]
    means2 = [data2.get(s, {}).get('mean', 0) for s in all_stations]
    stds2 = [data2.get(s, {}).get('std', 0) for s in all_stations]
    
    # Updated: Pastel Colors + Soft Alphas
    ax.bar(x - width/2, means1, width, yerr=stds1, capsize=4, label=name1, color=THEME_BLUE, alpha=0.55, edgecolor='white')
    ax.bar(x + width/2, means2, width, yerr=stds2, capsize=4, label=name2, color=THEME_ORANGE, alpha=0.55, edgecolor='white')
    
    spec_a, spec_b = 10.0, 15.0
    ax.axhline(y=spec_a, color=THEME_GREEN, linestyle='--', linewidth=2, label=f'Grade A: <{spec_a}')
    ax.axhline(y=spec_b, color=THEME_ORANGE, linestyle='--', linewidth=2, label=f'Grade B: <{spec_b}')
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    for spine in ax.spines.values(): spine.set_edgecolor(THEME_GRAY)
    ax.grid(True, alpha=0.2, axis='y', color=THEME_GRAY)
    
    ax.set_xlabel('CNC Station', fontsize=12, color='#444')
    ax.set_ylabel('Acceleration (m/s²)', fontsize=12, color='#444')
    ax.set_title(f'Spindle Acceleration Comparison\n{name1} vs {name2}', fontsize=14, color='#333')
    ax.set_xticks(x)
    ax.set_xticklabels(all_stations, rotation=45, ha='right', fontsize=9, color='#555')
    ax.tick_params(colors='#555')
    ax.legend(frameon=True, facecolor='white', edgecolor='none')
    
    plt.tight_layout()
    return fig_to_bytes(fig)


def compare_marble_squareness_combined(df1, df2, name1, name2):
    cnc_col1 = get_cnc_column_name(df1)
    cnc_col2 = get_cnc_column_name(df2)
    
    if not cnc_col1 and not cnc_col2:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, "CNC Station data not found", ha='center', va='center', fontsize=14, color=THEME_GRAY)
        return fig_to_bytes(fig)
    
    directions = ['XY', 'YZ', 'ZX']
    spec_a, spec_b = [16.0, 20.0, 20.0], [20.0, 30.0, 30.0]
    
    def convert_squareness_unit(value, col_name):
        if pd.isna(value): return np.nan
        if isinstance(value, (int, float)) and value < 1: return value * 1000
        return value
    
    def get_station_data(df, cnc_col):
        if cnc_col is None: return {}
        data = {}
        for station in df[cnc_col].dropna().unique():
            station_df = df[df[cnc_col] == station]
            station_plane_data = {}
            for plane in directions:
                alt_plane = plane[::-1]
                sq_cols = [c for c in df.columns if 'Squareness' in str(c) and (plane in str(c) or alt_plane in str(c))]
                vals = [convert_squareness_unit(v, col) for col in sq_cols for v in pd.to_numeric(station_df[col], errors='coerce').dropna()]
                vals = [v for v in vals if not pd.isna(v)]
                station_plane_data[plane] = {'mean': np.mean(vals), 'max': np.max(vals)} if vals else {'mean': np.nan, 'max': np.nan}
            if not all(np.isnan(station_plane_data[p]['mean']) for p in directions):
                data[station] = station_plane_data
        return data
    
    data1 = get_station_data(df1, cnc_col1)
    data2 = get_station_data(df2, cnc_col2)
    
    all_stations = sorted(set(data1.keys()) | set(data2.keys()))
    if not all_stations:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, "Marble Squareness data not found", ha='center', va='center', fontsize=14, color=THEME_GRAY)
        return fig_to_bytes(fig)
    
    n_stations = len(all_stations)
    cols = min(4, n_stations)
    rows = int(np.ceil(n_stations / cols))
    
    fig = plt.figure(figsize=(4 * cols, 4 * rows + 1.5))
    fig.patch.set_facecolor('none')
    
    angles = np.linspace(0, 2 * np.pi, len(directions), endpoint=False).tolist()
    angles += angles[:1]
    
    spec_a_closed = spec_a + [spec_a[0]]
    spec_b_closed = spec_b + [spec_b[0]]
    
    for i, station in enumerate(all_stations):
        ax = fig.add_subplot(rows, cols, i + 1, projection='polar')
        ax.set_facecolor('none')
        
        # Draw Specs with soft pastel colors
        ax.plot(angles, spec_b_closed, color=THEME_ORANGE, linestyle='--', linewidth=1.5)
        ax.fill(angles, spec_b_closed, color=THEME_ORANGE, alpha=0.05)
        
        ax.plot(angles, spec_a_closed, color=THEME_GREEN, linestyle='--', linewidth=1.5)
        ax.fill(angles, spec_a_closed, color=THEME_GREEN, alpha=0.08)
        
        station_data1, station_data2 = data1.get(station, {}), data2.get(station, {})
        means1 = [station_data1.get(p, {}).get('mean', np.nan) for p in directions]
        means2 = [station_data2.get(p, {}).get('mean', np.nan) for p in directions]
        
        means1_closed = [m if not np.isnan(m) else 0 for m in means1]
        means1_closed += [means1_closed[0]]
        means2_closed = [m if not np.isnan(m) else 0 for m in means2]
        means2_closed += [means2_closed[0]]
        
        # Transparent Overlay Data
        if not all(np.isnan(m) for m in means1):
            ax.plot(angles, means1_closed, color=THEME_BLUE, linewidth=2.0, label=f'{name1} Mean')
            ax.fill(angles, means1_closed, color=THEME_BLUE, alpha=0.35)
        if not all(np.isnan(m) for m in means2):
            ax.plot(angles, means2_closed, color=THEME_RED, linewidth=2.0, label=f'{name2} Mean')
            ax.fill(angles, means2_closed, color=THEME_RED, alpha=0.35)
        
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        ax.set_rlabel_position(0)
        ax.grid(color=THEME_GRAY, alpha=0.3)
        
        labels = [f"XY\n(A:{spec_a[0]}um, B:{spec_b[0]}um)", f"YZ\n(A:{spec_a[1]}um, B:{spec_b[1]}um)", f"ZX\n(A:{spec_a[2]}um, B:{spec_b[2]}um)"]
        ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontsize=10, color='#555')
        ax.set_title(station, y=1.2, fontsize=14, fontweight='bold', color='#333')
    
    fig.suptitle('Marble Squareness Profile by Station (Unit: um)', fontsize=18, fontweight='bold', color='#333')
    plt.tight_layout(rect=[0, 0, 1, 0.92])
    return fig_to_bytes(fig)


# ==========================================
# Main Streamlit UI
# ==========================================

def main():
    add_custom_css()
    
    # Updated text gradient to match pastel aesthetic
    st.markdown("""
    <div class="animate-fade-in-up" style="text-align: center;">
        <h1 style="font-size: 2.5rem; background: linear-gradient(135deg, #7A9CE0, #88CE95); 
                   -webkit-background-clip: text; -webkit-text-fill-color: transparent;">
            🔧 Machine Tool Precision Comparison Analysis System
        </h1>
        <p style="color: #9da3af; margin-top: 10px;">Compare and analyze machine tool precision across different factories</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="animate-fade-in-left">', unsafe_allow_html=True)
        st.subheader("🏭 Factory A")
        file1 = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'], key="file1")
        if file1: st.success(f"✅ Uploaded: {file1.name}")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="animate-fade-in-right">', unsafe_allow_html=True)
        st.subheader("🏭 Factory B")
        file2 = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'], key="file2")
        if file2: st.success(f"✅ Uploaded: {file2.name}")
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    with st.sidebar:
        st.markdown("### ⚙️ Animation Settings")
        enable_chart_animations = st.checkbox("Enable Chart Animations", value=True)
        animation_speed = st.select_slider("Animation Speed", options=["Slow", "Normal", "Fast"], value="Normal")
        speed_multiplier = {"Slow": 0.7, "Normal": 1.0, "Fast": 1.5}[animation_speed]
        
        st.markdown("---")
        st.markdown("### 📊 Analysis Features")
        st.markdown("""
        - Machine Type Count Comparison
        - Machine Age Distribution
        - Spindle Runout Analysis
        - Spindle Velocity Analysis
        - Spindle Acceleration Analysis
        - Marble Squareness Profile
        """)
    
    button_col1, button_col2, button_col3 = st.columns([1, 2, 1])
    with button_col2:
        start_button = st.button("🚀 Start Analysis", type="primary", use_container_width=True)
    
    if start_button:
        if not file1 or not file2:
            st.error("Please upload two Excel files")
            return
        
        loading_placeholder = st.empty()
        with loading_placeholder.container():
            show_loading_spinner("Loading and processing data...")
        
        try:
            name1, name2 = file1.name, file2.name
            content1, content2 = file1.read(), file2.read()
            
            read_mode1 = 'ipeg' if name1.lower().startswith('ipeg') else 'default'
            read_mode2 = 'ipeg' if name2.lower().startswith('ipeg') else 'default'
            factory1_name, factory2_name = extract_factory_name(name1), extract_factory_name(name2)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.markdown("*Loading Factory A data...*")
            progress_bar.progress(10)
            df1, err1 = load_excel_data(content1, factory1_name, read_mode=read_mode1)
            time.sleep(0.2 / speed_multiplier)
            
            status_text.markdown("*Loading Factory B data...*")
            progress_bar.progress(25)
            df2, err2 = load_excel_data(content2, factory2_name, read_mode=read_mode2)
            time.sleep(0.2 / speed_multiplier)
            
            if err1 or err2:
                loading_placeholder.empty()
                if err1: st.error(f"Failed to load File 1: {err1}")
                if err2: st.error(f"Failed to load File 2: {err2}")
                return
            
            progress_bar.progress(40)
            status_text.markdown("*Processing machine data...*")
            time.sleep(0.2 / speed_multiplier)
            
            loading_placeholder.empty()
            progress_bar.empty()
            status_text.empty()
            
            st.markdown("## 📊 Data Overview")
            cnc_col1, cnc_col2 = get_cnc_column_name(df1), get_cnc_column_name(df2)
            stations1 = int(df1[cnc_col1].nunique()) if cnc_col1 else 0
            stations2 = int(df2[cnc_col2].nunique()) if cnc_col2 else 0
            
            metric_cols = st.columns(4)
            with metric_cols[0]: display_animated_metric(f"Factory {factory1_name} - Records", f"{len(df1)}", animation_delay=0)
            with metric_cols[1]: display_animated_metric(f"Factory {factory1_name} - CNC Stations", f"{stations1}", animation_delay=0.1)
            with metric_cols[2]: display_animated_metric(f"Factory {factory2_name} - Records", f"{len(df2)}", animation_delay=0.2)
            with metric_cols[3]: display_animated_metric(f"Factory {factory2_name} - CNC Stations", f"{stations2}", animation_delay=0.3)
            
            st.markdown("---")
            
            chart_sections = [
                ("Machine Type Count Comparison", compare_machine_count),
                ("Machine Age Comparison", compare_machine_age),
                ("Spindle Runout Comparison", compare_spindle_runout),
                ("Spindle Velocity Comparison", compare_spindle_velocity),
                ("Spindle Acceleration Comparison", compare_spindle_acceleration),
                ("Marble Squareness Comparison", compare_marble_squareness_combined)
            ]
            
            for idx, (title, func) in enumerate(chart_sections):
                if enable_chart_animations:
                    shimmer_placeholder = st.empty()
                    with shimmer_placeholder.container():
                        st.markdown(f"### {title}")
                        show_shimmer_loading(400)
                    time.sleep(0.15 / speed_multiplier)
                    img_bytes = func(df1, df2, factory1_name, factory2_name)
                    shimmer_placeholder.empty()
                    display_animated_chart(img_bytes, title, idx)
                else:
                    with st.spinner(f"Generating {title}..."):
                        img_bytes = func(df1, df2, factory1_name, factory2_name)
                        st.markdown(f"### {title}")
                        st.image(img_bytes, use_container_width=True)
                st.markdown("---")
            
            show_completion_message()
            
        except Exception as e:
            loading_placeholder.empty()
            st.error(f"Error during analysis: {str(e)}")
            with st.expander("Show detailed error"):
                st.code(traceback.format_exc())

if __name__ == "__main__":
    main()

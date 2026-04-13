# file name: app_streamlit.py
# Factory Machine Tool Precision Comparison Analysis System - Streamlit Version

import streamlit as st
import pandas as pd
import numpy as np
import io
import base64
import re
import tempfile
import os
import traceback
from datetime import datetime
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.lines import Line2D

# 设置页面配置
st.set_page_config(
    page_title="机床精度对比分析系统",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Apple Design Colors
APPLE_BLUE = '#1A2AF0'
APPLE_ORANGE = '#E8642E'
APPLE_GREEN = '#34C759'
APPLE_RED = '#FF3B30'
APPLE_GRAY = '#86868B'

# 设置中文字体
try:
    import matplotlib.font_manager as fm
    plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'DejaVu Sans', 'sans-serif']
except:
    plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False


# ==========================================
# 辅助工具函数
# ==========================================

def extract_factory_name(file_name):
    """Extract factory name from file name"""
    base_name = os.path.basename(file_name)
    name_without_ext = os.path.splitext(base_name)[0]
    if len(name_without_ext) > 30:
        name_without_ext = name_without_ext[:30]
    return name_without_ext

def get_cnc_column_name(df):
    """自动识别 CNC 站位列名"""
    potential_cnc_cols = ['CNC OP', 'cnc op', 'Cnc Op', 'cnc station', 'CNC Station', 'Station']
    for col in df.columns:
        for potential in potential_cnc_cols:
            if potential.lower() == col.lower():
                return col
    return None

def clean_col_name(col):
    """清理列名"""
    col = str(col)
    col = re.sub(r'[\u4e00-\u9fff]+', '', col)
    col = re.sub(r'[\n\r]+', ' ', col)
    col = ' '.join(col.split())
    return col.strip()

def fig_to_bytes(fig):
    """Convert matplotlib Figure to bytes"""
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    plt.close(fig)
    return buf.getvalue()


# ==========================================
# 数据加载核心逻辑
# ==========================================

def load_excel_data(file_content, factory_name, read_mode='default'):
    """Load Excel data from file content"""
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
            except Exception as e:
                pass
            
            try:
                df_jd = pd.read_excel(io.BytesIO(file_content), sheet_name=jd_sheet_name, header=header_row)
                if df_jd is not None and not df_jd.empty:
                    df_jd = df_jd.dropna(how='all')
                    first_col = df_jd.columns[0]
                    df_jd = df_jd[df_jd[first_col].notna()]
                    if not df_jd.empty:
                        dfs_to_merge.append(df_jd)
            except Exception as e:
                pass
            
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
        
        # 清理列名并创建映射
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
            if pd.isna(val):
                return None
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
                    if 1980 <= year <= current_year + 5:
                        return year
            except:
                pass
            
            str_val = str(val)
            match = re.search(r'(19|20)\d{2}', str_val)
            if match:
                year = int(match.group(0))
                current_year = datetime.now().year
                if 1980 <= year <= current_year + 5:
                    return year
            return None

        year_col = df['Year of manufacturer']
        df['Year_of_manufacturer'] = year_col.apply(safe_extract_year)
        df = df.dropna(subset=['Year_of_manufacturer'])
        
        if df.empty:
            raise ValueError("No valid year data found after filtering")
        
        df['Year_of_manufacturer'] = df['Year_of_manufacturer'].astype(int)
        df['Factory'] = factory_name
        
        return df, None
        
    except Exception as e:
        traceback.print_exc()
        return None, str(e)


# ==========================================
# 图表生成模块
# ==========================================

def compare_machine_count(df1, df2, name1, name2):
    """Compare machine type count"""
    cnc_col = get_cnc_column_name(df1)
    
    fig, ax = plt.subplots(figsize=(14, 8))
    
    if cnc_col is None:
        ax.text(0.5, 0.5, "CNC Station column not found", ha='center', va='center', fontsize=14)
        return fig_to_bytes(fig)
    
    def get_detailed_counts(df):
        cnc_col_actual = get_cnc_column_name(df)
        if cnc_col_actual is None:
            return pd.DataFrame()
        
        counts_df = df.groupby(cnc_col_actual)['Machine Model'].nunique().reset_index()
        counts_df.columns = ['Station', 'Machine_Count']
        
        models_df = df.groupby(cnc_col_actual)['Machine Model'].apply(lambda x: sorted(x.unique())).reset_index()
        models_df.columns = ['Station', 'Machine_Models']
        
        result = counts_df.merge(models_df, on='Station')
        return result
    
    counts1 = get_detailed_counts(df1)
    counts2 = get_detailed_counts(df2)
    
    all_stations = sorted(set(counts1['Station']).union(set(counts2['Station'])))
    
    compare_df = pd.DataFrame({'Station': all_stations})
    compare_df = compare_df.merge(counts1[['Station', 'Machine_Count', 'Machine_Models']], on='Station', how='left').fillna(0)
    compare_df = compare_df.merge(counts2[['Station', 'Machine_Count', 'Machine_Models']], on='Station', how='left').fillna(0)
    compare_df.columns = ['Station', f'{name1}_Count', f'{name1}_Models', f'{name2}_Count', f'{name2}_Models']
    
    compare_df[f'{name1}_Models'] = compare_df[f'{name1}_Models'].apply(lambda x: x if isinstance(x, list) else [])
    compare_df[f'{name2}_Models'] = compare_df[f'{name2}_Models'].apply(lambda x: x if isinstance(x, list) else [])
    
    compare_df['Total'] = compare_df[f'{name1}_Count'] + compare_df[f'{name2}_Count']
    compare_df = compare_df.sort_values('Total', ascending=False).head(15)
    
    x = np.arange(len(compare_df['Station']))
    width = 0.35
    
    ax.bar(x - width/2, compare_df[f'{name1}_Count'], width, label=name1, color=APPLE_BLUE, alpha=0.7)
    ax.bar(x + width/2, compare_df[f'{name2}_Count'], width, label=name2, color=APPLE_ORANGE, alpha=0.7)
    
    ax.set_xlabel('CNC Station', fontsize=12)
    ax.set_ylabel('Number of Machine Types', fontsize=12)
    ax.set_title(f'Machine Type Count by Station\n{name1} vs {name2}', fontsize=14)
    ax.set_xticks(x)
    ax.set_xticklabels(compare_df['Station'], rotation=45, ha='right', fontsize=9)
    ax.legend()
    ax.grid(True, alpha=0.15, axis='y', color=APPLE_GRAY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    plt.tight_layout()
    return fig_to_bytes(fig)


def compare_machine_age(df1, df2, name1, name2):
    """Compare machine age"""
    current_year = datetime.now().year
    df1['Machine_Age'] = current_year - df1['Year_of_manufacturer']
    df2['Machine_Age'] = current_year - df2['Year_of_manufacturer']
    
    ages1 = df1['Machine_Age'].dropna()
    ages2 = df2['Machine_Age'].dropna()
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    
    max_age = int(max(ages1.max() if len(ages1) > 0 else 20, ages2.max() if len(ages2) > 0 else 20))
    bins = np.arange(0, max_age + 2, 1)
    
    if len(ages1) > 0 and len(ages2) > 0:
        bp = ax1.boxplot([ages1, ages2], tick_labels=[name1, name2], patch_artist=True)
        colors = [APPLE_BLUE, APPLE_ORANGE]
        for patch, color in zip(bp['boxes'], colors):
            patch.set_facecolor(color)
            patch.set_alpha(0.8)
            patch.set_edgecolor('white')
    else:
        ax1.text(0.5, 0.5, "Insufficient data", ha='center', va='center', fontsize=14)
    
    ax1.set_title('Machine Age Distribution', fontsize=14)
    ax1.set_ylabel('Age (Years)')
    ax1.grid(True, alpha=0.3, axis='y')
    
    ax2.hist(ages1, bins=bins, alpha=0.7, label=name1, color=APPLE_BLUE, edgecolor='white')
    ax2.hist(ages2, bins=bins, alpha=0.7, label=name2, color=APPLE_ORANGE, edgecolor='white')
    ax2.set_title('Age Distribution Histogram', fontsize=14)
    ax2.set_xlabel('Age (Years)')
    ax2.set_ylabel('Frequency')
    ax2.legend()
    ax2.grid(True, alpha=0.3, axis='y')
    
    for ax in [ax1, ax2]:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, alpha=0.15, axis='y', color=APPLE_GRAY)
    
    plt.tight_layout()
    return fig_to_bytes(fig)


def extract_spindle_runout_universal(df, position='near'):
    """通用主轴跳动数据提取"""
    values = []
    used_cols = []
    
    if position == 'near':
        position_patterns = [r'@5mm', r'@5\s*mm', r'5mm', r'5\s*mm', r'近端']
        exclude_patterns = [r'@300', r'300mm', r'150mm', r'@150', r'远端']
    else:
        position_patterns = [r'@300mm', r'@300\s*mm', r'300mm', r'@150mm', r'150mm', r'远端']
        exclude_patterns = [r'@5mm', r'5mm', r'近端']
    
    for col in df.columns:
        col_str = str(col).lower()
        
        if 'runout' not in col_str and '跳动' not in col_str:
            continue
        
        is_position_match = False
        for pattern in position_patterns:
            if re.search(pattern, col_str, re.IGNORECASE):
                is_position_match = True
                break
        
        if position == 'near' and not is_position_match:
            if 'spindle nose' in col_str or '主軸' in col_str:
                has_exclude = False
                for excl in exclude_patterns:
                    if re.search(excl, col_str, re.IGNORECASE):
                        has_exclude = True
                        break
                if not has_exclude:
                    is_position_match = True
        
        if not is_position_match:
            continue
        
        has_exclude = False
        for excl in exclude_patterns:
            if re.search(excl, col_str, re.IGNORECASE):
                has_exclude = True
                break
        if has_exclude:
            continue
        
        vals = pd.to_numeric(df[col], errors='coerce').dropna()
        if len(vals) == 0:
            continue
        
        if '[µm]' in col or '[μm]' in col or 'micron' in col_str or 'um]' in col:
            vals = vals / 1000
        elif '[mm]' in col or '(mm)' in col or 'mm]' in col:
            pass
        else:
            vals = vals / 1000
        
        values.extend(vals.tolist())
        used_cols.append(col)
    
    return values, used_cols


def compare_spindle_runout(df1, df2, name1, name2):
    """主轴跳动对比"""
    near1, _ = extract_spindle_runout_universal(df1, 'near')
    near2, _ = extract_spindle_runout_universal(df2, 'near')
    far1, _ = extract_spindle_runout_universal(df1, 'far')
    far2, _ = extract_spindle_runout_universal(df2, 'far')
    
    USL_NEAR_MM = 0.006
    USL_FAR_MM = 0.030
    
    has_near_data = len(near1) > 0 or len(near2) > 0
    has_far_data = len(far1) > 0 or len(far2) > 0
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 9))
    
    if not has_near_data and not has_far_data:
        ax1.text(0.5, 0.5, "Spindle runout data not found", ha='center', va='center', fontsize=12)
        plt.tight_layout()
        return fig_to_bytes(fig)
    
    if has_near_data:
        _plot_runout_distribution(ax1, near1, near2, name1, name2, USL_NEAR_MM,
                                   'Near End (5mm / Spindle Nose)\nSpec: ≤6µm')
    else:
        ax1.text(0.5, 0.5, "Near end data not found", ha='center', va='center', fontsize=11)
        ax1.set_title('Near End Runout - No Data', fontsize=12)
    
    if has_far_data:
        _plot_runout_distribution(ax2, far1, far2, name1, name2, USL_FAR_MM,
                                   'Far End (300mm from Spindle Nose)\nSpec: ≤30µm')
    else:
        ax2.text(0.5, 0.5, "Far end data not found", ha='center', va='center', fontsize=11)
        ax2.set_title('Far End Runout - No Data', fontsize=12)
    
    fig.suptitle(f'Spindle Runout Comparison\n{name1} vs {name2}', fontsize=14, fontweight='bold')
    plt.tight_layout()
    return fig_to_bytes(fig)


def _plot_runout_distribution(ax, data1, data2, name1, name2, usl_mm, title):
    """绘制跳动分布图"""
    def calc_stats(data, label):
        if len(data) == 0:
            return None
        mean = np.mean(data)
        std = np.std(data, ddof=1) if len(data) > 1 else 0.001
        if std == 0:
            std = 0.001
        cpk = (usl_mm - mean) / (3 * std) if usl_mm > mean else 0
        return {'n': len(data), 'mean': mean, 'std': std, 'cpk': cpk, 'label': label}
    
    stats1 = calc_stats(data1, name1)
    stats2 = calc_stats(data2, name2)
    
    all_data = data1 + data2
    if not all_data:
        ax.text(0.5, 0.5, "No valid data", ha='center', va='center', fontsize=12)
        ax.set_title(title)
        return
    
    x_min = min(0, min(all_data) * 0.9)
    x_max = max(usl_mm * 1.5, max(all_data) * 1.2)
    x_range = np.linspace(x_min, x_max, 200)
    
    if stats1:
        y1 = stats.norm.pdf(x_range, stats1['mean'], stats1['std'])
        ax.plot(x_range, y1, '-', linewidth=2.5, label=f'{name1} (Fit)', color=APPLE_BLUE)
        ax.hist(data1, bins=15, density=True, alpha=0.5, color=APPLE_BLUE, edgecolor='white')
    
    if stats2:
        y2 = stats.norm.pdf(x_range, stats2['mean'], stats2['std'])
        ax.plot(x_range, y2, '-', linewidth=2.5, label=f'{name2} (Fit)', color=APPLE_ORANGE)
        ax.hist(data2, bins=15, density=True, alpha=0.5, color=APPLE_ORANGE, edgecolor='white')
    
    ax.axvline(x=usl_mm, color=APPLE_RED, linestyle='--', linewidth=2, label=f'USL: {usl_mm*1000:.0f}µm')
    ax.axvline(x=usl_mm * 0.5, color=APPLE_GREEN, linestyle=':', linewidth=1.5, label=f'Target: {usl_mm*500:.0f}µm')
    
    ax.set_xlabel('Runout (mm)', fontsize=11)
    ax.set_ylabel('Probability Density', fontsize=11)
    ax.set_title(title, fontsize=11)
    ax.legend(loc='upper right', fontsize=8)
    ax.set_xlim(x_min, x_max)
    ax.grid(True, alpha=0.15)


def compare_spindle_velocity(df1, df2, name1, name2):
    """主轴速度对比"""
    target_stations = ['CNC4', 'CNC4.1', 'CNC5', 'CNC6', 'CNC7', 'CNC7.2', 'CNC8']
    cnc_col = get_cnc_column_name(df1)
    
    fig, ax = plt.subplots(figsize=(14, 7))
    
    if cnc_col is None:
        ax.text(0.5, 0.5, "CNC Station data not found", ha='center', va='center', fontsize=14)
        plt.tight_layout()
        return fig_to_bytes(fig)
    
    def extract_velocity_data(df, read_mode='default'):
        cnc_col_actual = get_cnc_column_name(df)
        if cnc_col_actual is None:
            return {}
        
        df_f = df[df[cnc_col_actual].isin(target_stations)]
        if df_f.empty:
            return {}
        
        extracted_data = {}
        
        if read_mode == 'ipeg':
            rpm_priority = [18000]
        else:
            rpm_priority = [18000, 16000, 10000]
        
        for station in df_f[cnc_col_actual].unique():
            station_df = df_f[df_f[cnc_col_actual] == station]
            selected_col = None
            selected_rpm = None
            
            for target_rpm in rpm_priority:
                for col in station_df.columns:
                    col_str = str(col)
                    col_lower = col_str.lower()
                    
                    is_velocity = ('velocity' in col_lower or '振动速度' in col_str)
                    if not is_velocity:
                        continue
                    
                    if 'spindle' not in col_lower and '主轴' not in col_str:
                        continue
                    
                    if str(target_rpm) in col_str:
                        vals = pd.to_numeric(station_df[col], errors='coerce').dropna()
                        if len(vals) > 0:
                            selected_col = col
                            selected_rpm = target_rpm
                            break
                if selected_col:
                    break
            
            if selected_col:
                vals = pd.to_numeric(station_df[selected_col], errors='coerce').dropna()
                if len(vals) > 0:
                    extracted_data[station] = {
                        'mean': np.mean(vals),
                        'std': np.std(vals) if len(vals) > 1 else 0.001,
                        'n': len(vals),
                        'rpm': selected_rpm
                    }
        
        return extracted_data
    
    read_mode1 = 'ipeg' if name1.lower().startswith('ipeg') else 'default'
    read_mode2 = 'ipeg' if name2.lower().startswith('ipeg') else 'default'
    
    data1 = extract_velocity_data(df1, read_mode=read_mode1)
    data2 = extract_velocity_data(df2, read_mode=read_mode2)
    
    if not data1 and not data2:
        ax.text(0.5, 0.5, "Velocity data not found", ha='center', va='center', fontsize=14)
        plt.tight_layout()
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
    
    ax.bar(x - width/2, means1, width, yerr=stds1, capsize=4,
           label=f'{name1} (n={sum(n1)})', color=APPLE_BLUE, alpha=0.6)
    ax.bar(x + width/2, means2, width, yerr=stds2, capsize=4,
           label=f'{name2} (n={sum(n2)})', color=APPLE_ORANGE, alpha=0.6)
    
    spec_a, spec_b = 1.1, 1.4
    ax.axhline(y=spec_a, color=APPLE_GREEN, linestyle='--', linewidth=2, label=f'Grade A: <{spec_a} mm/s')
    ax.axhline(y=spec_b, color=APPLE_ORANGE, linestyle='--', linewidth=2, label=f'Grade B: <{spec_b} mm/s')
    
    y_max = max(max(means1 + stds1, default=0), max(means2 + stds2, default=0), spec_b * 1.2)
    ax.fill_between([-0.5, len(all_stations) - 0.5], spec_b, y_max, color=APPLE_RED, alpha=0.05)
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(True, alpha=0.15, axis='y', color=APPLE_GRAY)
    
    ax.set_xlabel('CNC Station', fontsize=12)
    ax.set_ylabel('Velocity (mm/s)', fontsize=12)
    ax.set_title(f'Spindle Velocity Comparison\n{name1} vs {name2}', fontsize=14)
    ax.set_xticks(x)
    ax.set_xticklabels(all_stations, rotation=45, ha='right', fontsize=9)
    ax.set_ylim(0, y_max * 1.1)
    ax.legend(loc='upper left', fontsize=10)
    
    plt.tight_layout()
    return fig_to_bytes(fig)


def compare_spindle_acceleration(df1, df2, name1, name2):
    """主轴加速度对比"""
    target_stations = ['CNC4', 'CNC4.1', 'CNC5', 'CNC6', 'CNC7', 'CNC7.2', 'CNC8']
    cnc_col = get_cnc_column_name(df1)
    
    fig, ax = plt.subplots(figsize=(14, 7))
    
    if cnc_col is None:
        ax.text(0.5, 0.5, "CNC Station data not found", ha='center', va='center', fontsize=14)
        plt.tight_layout()
        return fig_to_bytes(fig)
    
    def extract_data(df):
        cnc_col_actual = get_cnc_column_name(df)
        if cnc_col_actual is None:
            return {}
        
        df_f = df[df[cnc_col_actual].isin(target_stations)]
        acc_cols = [c for c in df_f.columns if 'Acceleration' in str(c) and 'Spindle' in str(c)]
        data = {}
        for station in df_f[cnc_col_actual].unique():
            station_df = df_f[df_f[cnc_col_actual] == station]
            vals = []
            for col in acc_cols:
                vals.extend(pd.to_numeric(station_df[col], errors='coerce').dropna().tolist())
            if vals:
                data[station] = {'mean': np.mean(vals), 'std': np.std(vals) if len(vals) > 1 else 0}
        return data
    
    data1 = extract_data(df1)
    data2 = extract_data(df2)
    
    if not data1 and not data2:
        ax.text(0.5, 0.5, "Acceleration data not found", ha='center', va='center', fontsize=14)
        plt.tight_layout()
        return fig_to_bytes(fig)
    
    all_stations = sorted(set(data1.keys()) | set(data2.keys()))
    x = np.arange(len(all_stations))
    width = 0.35
    
    means1 = [data1.get(s, {}).get('mean', 0) for s in all_stations]
    stds1 = [data1.get(s, {}).get('std', 0) for s in all_stations]
    means2 = [data2.get(s, {}).get('mean', 0) for s in all_stations]
    stds2 = [data2.get(s, {}).get('std', 0) for s in all_stations]
    
    ax.bar(x - width/2, means1, width, yerr=stds1, capsize=4, label=name1, color=APPLE_BLUE, alpha=0.6)
    ax.bar(x + width/2, means2, width, yerr=stds2, capsize=4, label=name2, color=APPLE_ORANGE, alpha=0.6)
    
    spec_a, spec_b = 10.0, 15.0
    ax.axhline(y=spec_a, color=APPLE_GREEN, linestyle='--', label=f'Grade A: <{spec_a}')
    ax.axhline(y=spec_b, color=APPLE_ORANGE, linestyle='--', label=f'Grade B: <{spec_b}')
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(True, alpha=0.15, axis='y', color=APPLE_GRAY)
    
    ax.set_xlabel('CNC Station', fontsize=12)
    ax.set_ylabel('Acceleration (m/s²)', fontsize=12)
    ax.set_title(f'Spindle Acceleration Comparison\n{name1} vs {name2}', fontsize=14)
    ax.set_xticks(x)
    ax.set_xticklabels(all_stations, rotation=45, ha='right', fontsize=9)
    ax.legend()
    
    plt.tight_layout()
    return fig_to_bytes(fig)


def compare_marble_squareness_combined(df1, df2, name1, name2):
    """大理石垂直度对比"""
    cnc_col1 = get_cnc_column_name(df1)
    cnc_col2 = get_cnc_column_name(df2)
    
    if not cnc_col1 and not cnc_col2:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, "CNC Station data not found", ha='center', va='center', fontsize=14)
        return fig_to_bytes(fig)
    
    directions = ['XY', 'YZ', 'ZX']
    spec_a = [16.0, 20.0, 20.0]
    spec_b = [20.0, 30.0, 30.0]
    
    def convert_squareness_unit(value, col_name):
        if pd.isna(value):
            return np.nan
        if isinstance(value, (int, float)) and value < 1:
            return value * 1000
        return value
    
    def get_station_data(df, cnc_col):
        if cnc_col is None:
            return {}
        
        data = {}
        for station in df[cnc_col].dropna().unique():
            station_df = df[df[cnc_col] == station]
            station_plane_data = {}
            for plane in directions:
                alt_plane = plane[::-1]
                sq_cols = [c for c in df.columns if 'Squareness' in str(c) and (plane in str(c) or alt_plane in str(c))]
                
                vals = []
                for col in sq_cols:
                    col_vals = pd.to_numeric(station_df[col], errors='coerce').dropna()
                    for v in col_vals:
                        converted_v = convert_squareness_unit(v, col)
                        if not pd.isna(converted_v):
                            vals.append(converted_v)
                
                if vals:
                    station_plane_data[plane] = {'mean': np.mean(vals), 'max': np.max(vals)}
                else:
                    station_plane_data[plane] = {'mean': np.nan, 'max': np.nan}
            
            if not all(np.isnan(station_plane_data[p]['mean']) for p in directions):
                data[station] = station_plane_data
        return data
    
    data1 = get_station_data(df1, cnc_col1)
    data2 = get_station_data(df2, cnc_col2)
    
    all_stations = sorted(set(data1.keys()) | set(data2.keys()))
    if not all_stations:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, "Marble Squareness data not found", ha='center', va='center', fontsize=14)
        return fig_to_bytes(fig)
    
    n_stations = len(all_stations)
    cols = min(4, n_stations)
    rows = int(np.ceil(n_stations / cols))
    
    fig = plt.figure(figsize=(4 * cols, 4 * rows + 1.5))
    
    angles = np.linspace(0, 2 * np.pi, len(directions), endpoint=False).tolist()
    angles += angles[:1]
    
    spec_a_closed = spec_a + [spec_a[0]]
    spec_b_closed = spec_b + [spec_b[0]]
    
    for i, station in enumerate(all_stations):
        ax = fig.add_subplot(rows, cols, i + 1, projection='polar')
        
        ax.plot(angles, spec_b_closed, color=APPLE_ORANGE, linestyle='--', linewidth=1.5)
        ax.fill(angles, spec_b_closed, color=APPLE_ORANGE, alpha=0.05)
        
        ax.plot(angles, spec_a_closed, color=APPLE_GREEN, linestyle='--', linewidth=1.5)
        ax.fill(angles, spec_a_closed, color=APPLE_GREEN, alpha=0.05)
        
        station_data1 = data1.get(station, {})
        station_data2 = data2.get(station, {})
        
        means1 = [station_data1.get(p, {}).get('mean', np.nan) for p in directions]
        means2 = [station_data2.get(p, {}).get('mean', np.nan) for p in directions]
        
        means1_plot = [m if not np.isnan(m) else 0 for m in means1]
        means2_plot = [m if not np.isnan(m) else 0 for m in means2]
        
        means1_closed = means1_plot + [means1_plot[0]]
        means2_closed = means2_plot + [means2_plot[0]]
        
        if not all(np.isnan(m) for m in means1):
            ax.plot(angles, means1_closed, color=APPLE_BLUE, linewidth=2.5, label=f'{name1} Mean')
            ax.fill(angles, means1_closed, color=APPLE_BLUE, alpha=0.2)
        
        if not all(np.isnan(m) for m in means2):
            ax.plot(angles, means2_closed, color=APPLE_ORANGE, linewidth=2.5, label=f'{name2} Mean')
            ax.fill(angles, means2_closed, color=APPLE_ORANGE, alpha=0.2)
        
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        
        labels = [f"XY\n(A:{spec_a[0]}um, B:{spec_b[0]}um)",
                  f"YZ\n(A:{spec_a[1]}um, B:{spec_b[1]}um)",
                  f"ZX\n(A:{spec_a[2]}um, B:{spec_b[2]}um)"]
        ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontsize=10)
        
        ax.set_title(station, y=1.2, fontsize=14, fontweight='bold')
    
    fig.suptitle('Marble Squareness Profile by Station (Unit: µm)', fontsize=18, fontweight='bold')
    plt.tight_layout(rect=[0, 0, 1, 0.92])
    return fig_to_bytes(fig)


# ==========================================
# Streamlit UI
# ==========================================

def main():
    st.title("🔧 工厂机床精度对比分析系统")
    st.markdown("---")
    
    # 文件上传区域 - 两列布局
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("工厂 A")
        file1 = st.file_uploader("上传 Excel 文件", type=['xlsx', 'xls'], key="file1")
        if file1:
            st.success(f"✅ 已上传: {file1.name}")
    
    with col2:
        st.subheader("工厂 B")
        file2 = st.file_uploader("上传 Excel 文件", type=['xlsx', 'xls'], key="file2")
        if file2:
            st.success(f"✅ 已上传: {file2.name}")
    
    st.markdown("---")
    
    # 分析按钮
    if st.button("🚀 开始分析", type="primary", use_container_width=True):
        if not file1 or not file2:
            st.error("请上传两个 Excel 文件")
            return
        
        with st.spinner("正在分析数据，请稍候..."):
            try:
                # 读取文件名
                name1 = file1.name
                name2 = file2.name
                
                # 获取文件内容
                content1 = file1.read()
                content2 = file2.read()
                
                # 判断读取模式
                read_mode1 = 'ipeg' if name1.lower().startswith('ipeg') else 'default'
                read_mode2 = 'ipeg' if name2.lower().startswith('ipeg') else 'default'
                
                factory1_name = extract_factory_name(name1)
                factory2_name = extract_factory_name(name2)
                
                # 加载数据
                df1, err1 = load_excel_data(content1, factory1_name, read_mode=read_mode1)
                df2, err2 = load_excel_data(content2, factory2_name, read_mode=read_mode2)
                
                if err1:
                    st.error(f"加载文件 1 失败: {err1}")
                    return
                if err2:
                    st.error(f"加载文件 2 失败: {err2}")
                    return
                
                # 显示数据概览
                st.markdown("## 📊 数据概览")
                overview_col1, overview_col2, overview_col3, overview_col4 = st.columns(4)
                
                cnc_col1 = get_cnc_column_name(df1)
                cnc_col2 = get_cnc_column_name(df2)
                
                stations1 = int(df1[cnc_col1].nunique()) if cnc_col1 else 0
                stations2 = int(df2[cnc_col2].nunique()) if cnc_col2 else 0
                
                with overview_col1:
                    st.metric(f"🏭 {factory1_name} - 数据量", f"{len(df1)} 条")
                with overview_col2:
                    st.metric(f"🏭 {factory1_name} - CNC站位", f"{stations1} 个")
                with overview_col3:
                    st.metric(f"🏭 {factory2_name} - 数据量", f"{len(df2)} 条")
                with overview_col4:
                    st.metric(f"🏭 {factory2_name} - CNC站位", f"{stations2} 个")
                
                st.markdown("---")
                
                # 1. 机床类型对比
                st.markdown("## 1. 机床类型数量对比")
                with st.spinner("生成机床类型对比图..."):
                    img1 = compare_machine_count(df1, df2, factory1_name, factory2_name)
                    st.image(img1, use_container_width=True)
                st.markdown("---")
                
                # 2. 机龄对比
                st.markdown("## 2. 机床年龄对比")
                with st.spinner("生成机龄对比图..."):
                    img2 = compare_machine_age(df1, df2, factory1_name, factory2_name)
                    st.image(img2, use_container_width=True)
                st.markdown("---")
                
                # 3. 主轴跳动对比
                st.markdown("## 3. 主轴跳动对比")
                with st.spinner("生成主轴跳动对比图..."):
                    img3 = compare_spindle_runout(df1, df2, factory1_name, factory2_name)
                    st.image(img3, use_container_width=True)
                st.markdown("---")
                
                # 4. 主轴速度对比
                st.markdown("## 4. 主轴振动速度对比")
                with st.spinner("生成主轴速度对比图..."):
                    img4 = compare_spindle_velocity(df1, df2, factory1_name, factory2_name)
                    st.image(img4, use_container_width=True)
                st.markdown("---")
                
                # 5. 主轴加速度对比
                st.markdown("## 5. 主轴加速度对比")
                with st.spinner("生成主轴加速度对比图..."):
                    img5 = compare_spindle_acceleration(df1, df2, factory1_name, factory2_name)
                    st.image(img5, use_container_width=True)
                st.markdown("---")
                
                # 6. 大理石垂直度对比
                st.markdown("## 6. 大理石垂直度对比")
                with st.spinner("生成垂直度对比图..."):
                    img6 = compare_marble_squareness_combined(df1, df2, factory1_name, factory2_name)
                    st.image(img6, use_container_width=True)
                
                st.markdown("---")
                st.success("✅ 分析完成！")
                
            except Exception as e:
                st.error(f"分析过程中发生错误: {str(e)}")
                st.code(traceback.format_exc())


if __name__ == "__main__":
    main()

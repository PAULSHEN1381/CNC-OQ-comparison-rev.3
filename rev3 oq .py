import streamlit as st
import pandas as pd
import matplotlib
import numpy as np
import io
import base64
import tempfile
import traceback
import re
from datetime import datetime
from scipy import stats

# 使用非交互式后端，适用于 Streamlit
matplotlib.use('Agg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from matplotlib.lines import Line2D

# Apple Design Colors (保持不变)
APPLE_BLUE = '#1A2AF0'
APPLE_ORANGE = '#E8642E'
APPLE_GREEN = '#34C759'
APPLE_RED = '#FF3B30'
APPLE_GRAY = '#86868B'

# --- Streamlit 配置 ---
st.set_page_config(
    layout="wide",
    page_title="CNC OQ Precision Comparison",
    initial_sidebar_state="auto"
)

st.title("🏭 CNC Machine OQ Comparison")
st.write("Upload two Excel files to automatically compare and analyze machine types, age, spindle runout, velocity, acceleration, and marble squareness.")

# --- Helper Functions (从原代码复制) ---

def extract_factory_name(file_path):
    """Extract factory name from file path"""
    base_name = os.path.basename(file_path)
    name_without_ext = os.path.splitext(base_name)[0]
    if len(name_without_ext) > 30:
        name_without_ext = name_without_ext[:30]
    return name_without_ext

def get_cnc_column_name(df):
    """全局辅助函数：自动识别 CNC 站位列名"""
    potential_cnc_cols = ['CNC OP', 'cnc op', 'Cnc Op', 'cnc station', 'CNC Station', 'Station']
    
    for col in df.columns:
        for potential in potential_cnc_cols:
            if potential.lower() == col.lower():
                return col
    return None

def clean_col_name(col):
    """清理列名：移除中文字符，保留英文、数字、标点符号"""
    col = str(col)
    col = re.sub(r'[\u4e00-\u9fff]+', '', col)
    col = re.sub(r'[\n\r]+', ' ', col)
    col = ' '.join(col.split())
    col = col.replace('（', '').replace('）', '').replace('(', '').replace(')', '')
    return col.strip()

def fig_to_base64(fig):
    """Convert matplotlib Figure object to base64 string"""
    buf = io.BytesIO()
    # 使用 FigureCanvasAgg 来保存图像
    canvas = FigureCanvas(fig)
    canvas.print_png(buf)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode('utf-8')
    fig.clf() # 清理 Figure 对象，防止内存泄漏
    return img_base64

# --- Data Loading ---
def load_excel_data(file_path, factory_name, read_mode='default'):
    """Load Excel data"""
    try:
        df = None
        
        if read_mode == 'ipeg':
            fanuc_sheet_name = 'Machine OQ Rev-Fanuc'
            jd_sheet_name = 'Machine OQ Rev-JD'
            header_row = 3  # 第4行是列名（0-indexed）
            
            st.write(f"  {factory_name}: Reading as ipeg mode - merging Fanuc and JD sheets")
            
            dfs_to_merge = []
            
            # 读取 Fanuc 表
            try:
                df_fanuc = pd.read_excel(file_path, sheet_name=fanuc_sheet_name, header=header_row)
                if df_fanuc is not None and not df_fanuc.empty:
                    df_fanuc = df_fanuc.dropna(how='all')
                    first_col = df_fanuc.columns[0]
                    df_fanuc = df_fanuc[df_fanuc[first_col].notna()]
                    
                    if not df_fanuc.empty:
                        dfs_to_merge.append(df_fanuc)
                        st.write(f"  {factory_name}: Read {len(df_fanuc)} rows from Fanuc sheet")
            except Exception as e:
                st.warning(f"  {factory_name}: Could not read Fanuc sheet: {e}")
            
            # 读取 JD 表
            try:
                df_jd = pd.read_excel(file_path, sheet_name=jd_sheet_name, header=header_row)
                if df_jd is not None and not df_jd.empty:
                    df_jd = df_jd.dropna(how='all')
                    first_col = df_jd.columns[0]
                    df_jd = df_jd[df_jd[first_col].notna()]
                    
                    if not df_jd.empty:
                        dfs_to_merge.append(df_jd)
                        st.write(f"  {factory_name}: Read {len(df_jd)} rows from JD sheet")
            except Exception as e:
                st.warning(f"  {factory_name}: Could not read JD sheet: {e}")
            
            if not dfs_to_merge:
                raise ValueError("No data found in either Fanuc or JD sheets")
            
            df = pd.concat(dfs_to_merge, ignore_index=True)
            st.write(f"  {factory_name}: Merged total of {len(df)} rows")
            
        else:  # Default mode for tzl files
            st.write(f"  {factory_name}: Reading in default mode.")
            try:
                df = pd.read_excel(file_path, sheet_name='Table', header=2)
                st.write(f"  {factory_name}: Default mode - Using sheet 'Table', header=2...")
            except:
                try:
                    df = pd.read_excel(file_path, sheet_name='Table', header=1)
                    st.write(f"  {factory_name}: Default mode - Using sheet 'Table', header=1...")
                except:
                    df = pd.read_excel(file_path)
                    st.write(f"  {factory_name}: Default mode - Using auto header...")
        
        if df is None or df.empty:
            raise ValueError("Could not load any data from the Excel file after trying specified/default methods.")
        
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
        
        st.write(f"  {factory_name}: Before filtering - {len(df)} rows")
        st.write(f"  {factory_name}: Columns after rename: {list(df.columns)}")
        
        required_columns = ['CNC OP', 'Machine Model', 'Year of manufacturer']
        missing = [col for col in required_columns if col not in df.columns]
        
        if missing:
            st.error(f"  {factory_name}: ERROR - Missing required columns: {missing}. Available columns: {list(df.columns)}")
            raise ValueError(f"Required columns not found: {missing}")
        
        df['CNC OP'] = df['CNC OP'].astype(str)
        df['Machine Model'] = df['Machine Model'].astype(str)
        
        df = df[df['CNC OP'].notna() & (df['CNC OP'] != 'nan') & (df['CNC OP'] != '')]
        df = df[df['Machine Model'].notna() & (df['Machine Model'] != 'nan') & (df['Machine Model'] != '')]
        
        st.write(f"  {factory_name}: After removing empty CNC OP/Machine Model - {len(df)} rows")

        from datetime import datetime, timedelta
        
        def safe_extract_year(val):
            if pd.isna(val): return None
            try:
                if isinstance(val, (int, float)):
                    num_val = int(val)
                elif isinstance(val, str) and val.replace('.', '').isdigit():
                    num_val = int(float(val))
                else: num_val = None
                
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

        year_col = df['Year of manufacturer']
        df['Year_of_manufacturer'] = year_col.apply(safe_extract_year)
        
        before_year_filter = len(df)
        df = df.dropna(subset=['Year_of_manufacturer'])
        st.write(f"  {factory_name}: After year filter - kept {len(df)} out of {before_year_filter} rows")
        
        if df.empty:
            raise ValueError("No valid year data found after filtering")
        
        df['Year_of_manufacturer'] = df['Year_of_manufacturer'].astype(int)
        df['Factory'] = factory_name
        
        st.write(f"  {factory_name}: Year range: {df['Year_of_manufacturer'].min()} - {df['Year_of_manufacturer'].max()}")
        st.write(f"  {factory_name}: Successfully loaded {len(df)} rows after cleaning")
        
        return df, None
        
    except Exception as e:
        traceback.print_exc() # For debugging in Streamlit logs
        return None, str(e)

# --- Plotting Functions (从原代码复制) ---

def compare_machine_count(df1, df2, name1, name2):
    """Compare machine type count with detailed machine models"""
    cnc_col = get_cnc_column_name(df1)
    
    fig = Figure(figsize=(14, 8))
    ax = fig.add_subplot(111)
    
    if cnc_col is None:
        ax.text(0.5, 0.5, "CNC Station column not found", ha='center', va='center', fontsize=14)
        fig.tight_layout()
        return fig_to_base64(fig), []
    
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
    
    for i, (idx, row) in enumerate(compare_df.iterrows()):
        models1 = ', '.join(row[f'{name1}_Models']) if row[f'{name1}_Models'] else 'No data'
        models2 = ', '.join(row[f'{name2}_Models']) if row[f'{name2}_Models'] else 'No data'
        
        if row[f'{name1}_Count'] > 0:
            ax.text(i - width/2, row[f'{name1}_Count'] + 0.1, 
                   models1, ha='center', va='bottom', fontsize=6, rotation=45,
                   color=APPLE_BLUE, fontweight='bold')
        
        if row[f'{name2}_Count'] > 0:
            ax.text(i + width/2, row[f'{name2}_Count'] + 0.1,
                   models2, ha='center', va='bottom', fontsize=6, rotation=45,
                   color=APPLE_ORANGE, fontweight='bold')
    
    ax.set_xlabel('CNC Station', fontsize=12)
    ax.set_ylabel('Number of Machine Types', fontsize=12)
    ax.set_title(f'Machine Type Count by Station with Model Details\n{name1} vs {name2}', fontsize=14)
    ax.set_xticks(x)
    ax.set_xticklabels(compare_df['Station'], rotation=45, ha='right', fontsize=9)
    ax.legend()
    
    ax.grid(True, alpha=0.15, axis='y', color=APPLE_GRAY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    fig.tight_layout()
    fig.subplots_adjust(bottom=0.2)
    
    detailed_data = []
    for idx, row in compare_df.iterrows():
        detailed_data.append({
            'Station': row['Station'],
            name1: {'count': int(row[f'{name1}_Count']), 'models': row[f'{name1}_Models']},
            name2: {'count': int(row[f'{name2}_Count']), 'models': row[f'{name2}_Models']}
        })
    
    return fig_to_base64(fig), detailed_data


def compare_machine_age(df1, df2, name1, name2):
    """Compare machine age"""
    current_year = datetime.now().year
    df1['Machine_Age'] = current_year - df1['Year_of_manufacturer']
    df2['Machine_Age'] = current_year - df2['Year_of_manufacturer']
    
    ages1 = df1['Machine_Age'].dropna()
    ages2 = df2['Machine_Age'].dropna()
    
    fig = Figure(figsize=(14, 6))
    ax1 = fig.add_subplot(121)
    ax2 = fig.add_subplot(122)
    
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

    fig.tight_layout()
    
    # Perform t-test for age comparison
    ttest_result = None
    if len(ages1) > 1 and len(ages2) > 1:
        try:
            ttest_stat, p_value = stats.ttest_ind(ages1, ages2, equal_var=False) # Welch's t-test
            ttest_result = {
                'p_value': p_value,
                'significant': 'significant' if p_value < 0.05 else 'no significant difference'
            }
        except Exception as e:
            st.warning(f"Could not perform t-test: {e}")

    stats_data = {
        'factory1': {
            'count': int(len(ages1)), 
            'mean': f"{ages1.mean():.1f}" if len(ages1) > 0 else "0.0", 
            'median': f"{ages1.median():.1f}" if len(ages1) > 0 else "0.0"
        },
        'factory2': {
            'count': int(len(ages2)), 
            'mean': f"{ages2.mean():.1f}" if len(ages2) > 0 else "0.0", 
            'median': f"{ages2.median():.1f}" if len(ages2) > 0 else "0.0"
        },
        'ttest': ttest_result
    }
    
    return fig_to_base64(fig), stats_data


def extract_spindle_runout_universal(df, position='near'):
    """
    通用主轴跳动数据提取 - 兼容两种列名格式
    
    position: 'near' (近端/5mm) 或 'far' (远端/150mm/300mm)
    
    支持的列名格式：
    1. "Runout at spindle nose @5mm (mm)（主軸跳動至5mm數值"
    2. "Runout: At spindle nose [µm]"
    3. "Runout: At 300mm from spindle nose [µm]"
    4. "Runout at spindle nose @300mm (mm)"
    """
    values = []
    used_cols = []
    
    if position == 'near':
        position_patterns = [r'@5mm', r'@5\s*mm', r'5mm', r'5\s*mm', r'近端']
        exclude_patterns = [r'@300', r'300mm', r'150mm', r'@150', r'远端']
    else:  # far
        position_patterns = [r'@300mm', r'@300\s*mm', r'300mm', r'300\s*mm', r'@150mm', r'150mm', r'远端']
        exclude_patterns = [r'@5mm', r'5mm', r'近端']
    
    for col in df.columns:
        col_str = str(col).lower()
        
        if 'runout' not in col_str and '跳动' not in col_str: continue
        
        is_position_match = False
        for pattern in position_patterns:
            if re.search(pattern, col_str, re.IGNORECASE):
                is_position_match = True
                break
        
        if position == 'near' and not is_position_match:
            if 'spindle nose' in col_str or '主軸' in col_str:
                has_exclude = any(re.search(excl, col_str, re.IGNORECASE) for excl in exclude_patterns)
                if not has_exclude:
                    is_position_match = True
        
        if not is_position_match: continue
        
        has_exclude = any(re.search(excl, col_str, re.IGNORECASE) for excl in exclude_patterns)
        if has_exclude: continue
        
        vals = pd.to_numeric(df[col], errors='coerce').dropna()
        if len(vals) == 0: continue
        
        unit = ''
        if '[µm]' in col or '[μm]' in col or 'micron' in col_str or 'um]' in col:
            vals = vals / 1000
            unit = 'µm→mm'
        elif '[mm]' in col or '(mm)' in col or 'mm]' in col:
            unit = 'mm'
        else:
            vals = vals / 1000
            unit = 'assumed µm→mm'
        
        values.extend(vals.tolist())
        used_cols.append(f"{col} [{unit}]")
    
    return values, used_cols

def compare_spindle_runout(df1, df2, name1, name2):
    """兼容两种列名格式的主轴跳动对比"""
    near1, cols1_near = extract_spindle_runout_universal(df1, 'near')
    near2, cols2_near = extract_spindle_runout_universal(df2, 'near')
    
    far1, cols1_far = extract_spindle_runout_universal(df1, 'far')
    far2, cols2_far = extract_spindle_runout_universal(df2, 'far')
    
    USL_NEAR_MM = 0.006   # 6µm
    USL_FAR_MM = 0.030    # 30µm
    
    has_near_data = len(near1) > 0 or len(near2) > 0
    has_far_data = len(far1) > 0 or len(far2) > 0
    
    fig = Figure(figsize=(18, 9))
    
    if not has_near_data and not has_far_data:
        ax = fig.add_subplot(111)
        ax.text(0.5, 0.5, "主轴跳动数据未找到\n\n支持的列名格式：\n1. 'Runout at spindle nose @5mm (mm)'\n2. 'Runout: At spindle nose [µm]'", 
                ha='center', va='center', fontsize=12, transform=ax.transAxes)
        fig.tight_layout()
        return fig_to_base64(fig)
    
    ax1 = fig.add_subplot(121)
    if has_near_data:
        _plot_runout_distribution(ax1, near1, near2, name1, name2, USL_NEAR_MM, 
                                   'Near End (5mm / Spindle Nose)\n spec: ≤6µm')
        if cols1_near:
            ax1.text(0.02, 0.02, f"📊 {name1}: {cols1_near[0][:40]}", 
                    transform=ax1.transAxes, fontsize=7, color=APPLE_BLUE, alpha=0.7)
        if cols2_near:
            ax1.text(0.02, 0.08 if cols1_near else 0.02, f"📊 {name2}: {cols2_near[0][:40]}", 
                    transform=ax1.transAxes, fontsize=7, color=APPLE_ORANGE, alpha=0.7)
    else:
        ax1.text(0.5, 0.5, "近端数据未找到\n\n请检查列名是否包含：\n'@5mm' 或 'Spindle nose'", 
                ha='center', va='center', fontsize=11, transform=ax1.transAxes)
        ax1.set_title('Near End Runout - No Data', fontsize=12)
    
    ax2 = fig.add_subplot(122)
    if has_far_data:
        _plot_runout_distribution(ax2, far1, far2, name1, name2, USL_FAR_MM,
                                   'Far End (300mm from Spindle Nose)\n spec: ≤30µm')
        if cols1_far:
            ax2.text(0.02, 0.02, f"📊 {name1}: {cols1_far[0][:40]}", 
                    transform=ax2.transAxes, fontsize=7, color=APPLE_BLUE, alpha=0.7)
        if cols2_far:
            ax2.text(0.02, 0.08 if cols1_far else 0.02, f"📊 {name2}: {cols2_far[0][:40]}", 
                    transform=ax2.transAxes, fontsize=7, color=APPLE_ORANGE, alpha=0.7)
    else:
        ax2.text(0.5, 0.5, "远端数据未找到\n\n请检查列名是否包含：\n'@300mm' 或 '300mm'", 
                ha='center', va='center', fontsize=11, transform=ax2.transAxes)
        ax2.set_title('Far End Runout - No Data', fontsize=12)
    
    fig.suptitle(f'spindle runout comparison\n{name1} vs {name2}', fontsize=14, fontweight='bold')
    fig.tight_layout()
    
    return fig_to_base64(fig)

def _plot_runout_distribution(ax, data1, data2, name1, name2, usl_mm, title):
    """绘制跳动分布图"""
    import scipy.stats as stats
    
    def calc_stats(data, label):
        if len(data) == 0: return None
        mean = np.mean(data)
        std = np.std(data, ddof=1) if len(data) > 1 else 0.001
        if std == 0: std = 0.001
        cpk = (usl_mm - mean) / (3 * std) if usl_mm > mean else 0
        pct_out = sum(1 for v in data if v > usl_mm) / len(data) * 100 if len(data) > 0 else 0
        return {'n': len(data), 'mean': mean, 'std': std, 'cpk': cpk, 'pct_out': pct_out, 'label': label}
    
    stats1 = calc_stats(data1, name1)
    stats2 = calc_stats(data2, name2)
    
    all_data = data1 + data2
    if not all_data:
        ax.text(0.5, 0.5, "No valid data", ha='center', va='center', fontsize=12)
        ax.set_title(title)
        return
    
    x_min = min(0, min(all_data) * 0.9) if all_data else 0
    x_max = max(usl_mm * 1.5, max(all_data) * 1.2) if all_data else usl_mm * 1.5
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
    
    y_pos = 0.98
    if stats1:
        rating = "Capable" if stats1['cpk'] >= 1.33 else ("Marginal" if stats1['cpk'] >= 1.0 else "Not Capable")
        text = f"{stats1['label']}\nN={stats1['n']}  Mean={stats1['mean']:.4f}mm\nCpk={stats1['cpk']:.2f}  [{rating}]"
        ax.text(0.02, y_pos, text, transform=ax.transAxes, fontsize=7, va='top',
               bbox=dict(boxstyle="round", facecolor=APPLE_BLUE, alpha=0.1))
        y_pos -= 0.12
    
    if stats2:
        rating = "Capable" if stats2['cpk'] >= 1.33 else ("Marginal" if stats2['cpk'] >= 1.0 else "Not Capable")
        text = f"{stats2['label']}\nN={stats2['n']}  Mean={stats2['mean']:.4f}mm\nCpk={stats2['cpk']:.2f}  [{rating}]"
        ax.text(0.02, y_pos, text, transform=ax.transAxes, fontsize=7, va='top',
               bbox=dict(boxstyle="round", facecolor=APPLE_ORANGE, alpha=0.1))

def compare_spindle_velocity(df1, df2, name1, name2):
    """Compare spindle velocity - 按优先级抓取特定转速数据"""
    target_stations = ['CNC4', 'CNC4.1', 'CNC5', 'CNC6', 'CNC7', 'CNC7.2', 'CNC8']
    cnc_col = get_cnc_column_name(df1)
    
    fig = Figure(figsize=(14, 7))
    ax = fig.add_subplot(111)
    
    if cnc_col is None:
        ax.text(0.5, 0.5, "CNC Station data not found", ha='center', va='center', fontsize=14)
        fig.tight_layout()
        return fig_to_base64(fig)

    def extract_velocity_data(df, read_mode='default'):
        cnc_col_actual = get_cnc_column_name(df)
        if cnc_col_actual is None: return {}
        
        df_f = df[df[cnc_col_actual].isin(target_stations)]
        if df_f.empty: return {}
        
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
                    col_clean = col_lower.replace('"', '').replace('“', '').replace('”', '').replace("'", '')
                    
                    is_velocity = ('velocity' in col_clean or '振动速度' in col_str)
                    if not is_velocity: continue
                    
                    if 'spindle' not in col_clean and '主轴' not in col_str: continue
                    
                    if str(target_rpm) in col_str:
                        vals = pd.to_numeric(station_df[col], errors='coerce').dropna()
                        if len(vals) > 0:
                            selected_col = col
                            selected_rpm = target_rpm
                            break
                if selected_col: break
            
            if selected_col:
                vals = pd.to_numeric(station_df[selected_col], errors='coerce').dropna()
                if len(vals) > 0:
                    extracted_data[station] = {
                        'mean': np.mean(vals),
                        'std': np.std(vals) if len(vals) > 1 else 0.001,
                        'n': len(vals),
                        'rpm': selected_rpm,
                        'col': selected_col[:50]
                    }
        return extracted_data

    read_mode1 = 'ipeg' if name1.lower().startswith('ipeg') else 'default'
    read_mode2 = 'ipeg' if name2.lower().startswith('ipeg') else 'default'
    
    data1 = extract_velocity_data(df1, read_mode=read_mode1)
    data2 = extract_velocity_data(df2, read_mode=read_mode2)
    
    if not data1 and not data2:
        ax.text(0.5, 0.5, "Velocity data not found for any station.\n\nPlease check your input files and column names.\n\nExpected formats:\n- For 'ipeg' mode: '“Velocity @ Spindle @ 18000 rpm [mm/s]' or '主轴振动速度 - 18000RPM'\n- For 'default' mode: '“Velocity @ Spindle @ 18000 rpm [mm/s]', '“Velocity @ Spindle @ 16000 rpm [mm/s]', or '“Velocity @ Spindle @ 10000 rpm [mm/s]'", 
                ha='center', va='center', fontsize=12, transform=ax.transAxes)
        fig.tight_layout()
        return fig_to_base64(fig)
    
    all_stations = sorted(set(data1.keys()) | set(data2.keys()))
    x = np.arange(len(all_stations))
    width = 0.35
    
    means1 = [data1.get(s, {}).get('mean', 0) for s in all_stations]
    stds1 = [data1.get(s, {}).get('std', 0) for s in all_stations]
    means2 = [data2.get(s, {}).get('mean', 0) for s in all_stations]
    stds2 = [data2.get(s, {}).get('std', 0) for s in all_stations]
    
    n1 = [data1.get(s, {}).get('n', 0) for s in all_stations]
    n2 = [data2.get(s, {}).get('n', 0) for s in all_stations]
    
    rpm_info1 = set()
    for s in data1: rpm_info1.add(data1[s].get('rpm', 0))
    rpm_info2 = set()
    for s in data2: rpm_info2.add(data2[s].get('rpm', 0))
    
    rpm_text1 = f" (@{max(rpm_info1)}rpm)" if rpm_info1 else ""
    rpm_text2 = f" (@{max(rpm_info2)}rpm)" if rpm_info2 else ""
    
    ax.bar(x - width/2, means1, width, yerr=stds1, capsize=4, 
            label=f'{name1}{rpm_text1} (n={sum(n1)})', color=APPLE_BLUE, alpha=0.6)
    ax.bar(x + width/2, means2, width, yerr=stds2, capsize=4,
            label=f'{name2}{rpm_text2} (n={sum(n2)})', color=APPLE_ORANGE, alpha=0.6)
    
    spec_a, spec_b = 1.1, 1.4 # 假设的规格值，单位 mm/s
    ax.axhline(y=spec_a, color=APPLE_GREEN, linestyle='--', linewidth=2, label=f'Grade A: <{spec_a} mm/s')
    ax.axhline(y=spec_b, color=APPLE_ORANGE, linestyle='--', linewidth=2, label=f'Grade B: <{spec_b} mm/s')
    
    y_max = max(max(means1 + stds1, default=0), max(means2 + stds2, default=0), spec_b * 1.2)
    ax.fill_between([-0.5, len(all_stations) - 0.5], spec_b, y_max, 
                     color=APPLE_RED, alpha=0.05)
    
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
    
    fig.tight_layout()
    
    return fig_to_base64(fig)

def compare_spindle_acceleration(df1, df2, name1, name2):
    """Compare spindle acceleration"""
    target_stations = ['CNC4', 'CNC4.1', 'CNC5', 'CNC6', 'CNC7', 'CNC7.2', 'CNC8']
    cnc_col = get_cnc_column_name(df1)
    
    fig = Figure(figsize=(14, 7))
    ax = fig.add_subplot(111)
    
    if cnc_col is None:
        ax.text(0.5, 0.5, "CNC Station data not found", ha='center', va='center', fontsize=14)
        fig.tight_layout()
        return fig_to_base64(fig)
    
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
            if vals:
                data[station] = {'mean': np.mean(vals), 'std': np.std(vals) if len(vals) > 1 else 0}
        return data
    
    data1 = extract_data(df1)
    data2 = extract_data(df2)
    
    if not data1 and not data2:
        ax.text(0.5, 0.5, "Acceleration data not found", ha='center', va='center', fontsize=14)
        fig.tight_layout()
        return fig_to_base64(fig)
    
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

    fig.tight_layout()
    return fig_to_base64(fig)

def convert_squareness_unit(value, col_name):
    """
    自动检测并转换垂直度单位
    规则：如果数值 < 1，则认为是 mm，转换为 um（乘以 1000）
         如果数值 > 1，则认为是 um，保持不变
    """
    if pd.isna(value): return np.nan
    if isinstance(value, (int, float)) and value < 1:
        return value * 1000
    return value

def compare_marble_squareness_combined(df1, df2, name1, name2):
    """Compare Marble Squareness for all planes using a grid of Radar Charts"""
    cnc_col1 = get_cnc_column_name(df1)
    cnc_col2 = get_cnc_column_name(df2)
    
    if not cnc_col1 and not cnc_col2:
        fig = Figure(figsize=(8, 4))
        ax = fig.add_subplot(111)
        ax.text(0.5, 0.5, "CNC Station data not found", ha='center', va='center', fontsize=14)
        return fig_to_base64(fig)
        
    directions = ['XY', 'YZ', 'ZX']
    spec_a = [16.0, 20.0, 20.0]   # 单位：um
    spec_b = [20.0, 30.0, 30.0]   # 单位：um
    
    def get_station_data(df, cnc_col):
        if cnc_col is None: return {}
        
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
            
            if not all(np.isnan(station_plane_data[p]['mean']) for p in directions) or \
               not all(np.isnan(station_plane_data[p]['max']) for p in directions):
                data[station] = station_plane_data
        return data
        
    data1 = get_station_data(df1, cnc_col1)
    data2 = get_station_data(df2, cnc_col2)
    
    all_stations = sorted(set(data1.keys()) | set(data2.keys()))
    if not all_stations:
        fig = Figure(figsize=(8, 4))
        ax = fig.add_subplot(111)
        ax.text(0.5, 0.5, "Marble Squareness data not found", ha='center', va='center', fontsize=14)
        return fig_to_base64(fig)
        
    n_stations = len(all_stations)
    cols = min(4, n_stations)
    rows = int(np.ceil(n_stations / cols))
    
    fig = Figure(figsize=(4 * cols, 4 * rows + 1.5)) 
    
    angles = np.linspace(0, 2 * np.pi, len(directions), endpoint=False).tolist()
    angles += angles[:1]

    spec_a_closed = spec_a + [spec_a[0]]
    spec_b_closed = spec_b + [spec_b[0]]
    
    for i, station in enumerate(all_stations):
        ax = fig.add_subplot(rows, cols, i + 1, polar=True)
        
        ax.plot(angles, spec_b_closed, color=APPLE_ORANGE, linestyle='--', linewidth=1.5)
        ax.fill(angles, spec_b_closed, color=APPLE_ORANGE, alpha=0.05)
        
        ax.plot(angles, spec_a_closed, color=APPLE_GREEN, linestyle='--', linewidth=1.5)
        ax.fill(angles, spec_a_closed, color=APPLE_GREEN, alpha=0.05)
        
        station_data1 = data1.get(station, {})
        station_data2 = data2.get(station, {})

        means1 = [station_data1.get(p, {}).get('mean', np.nan) for p in directions]
        maxes1 = [station_data1.get(p, {}).get('max', np.nan) for p in directions]
        means2 = [station_data2.get(p, {}).get('mean', np.nan) for p in directions]
        maxes2 = [station_data2.get(p, {}).get('max', np.nan) for p in directions]

        means1_plot = [m if not np.isnan(m) else 0 for m in means1]
        means2_plot = [m if not np.isnan(m) else 0 for m in means2]

        means1_closed = means1_plot + [means1_plot[0]]
        means2_closed = means2_plot + [means2_plot[0]]
        
        max_overall_val = max(spec_b) 

        if not all(np.isnan(m) for m in means1) or not all(np.isnan(m) for m in maxes1):
            ax.plot(angles, means1_closed, color=APPLE_BLUE, linewidth=2.5, label=f'{name1} Mean')
            ax.fill(angles, means1_closed, color=APPLE_BLUE, alpha=0.2)
            for ang, val in zip(angles[:-1], means1):
                if not np.isnan(val):
                    ax.scatter(ang, val, color=APPLE_BLUE, s=50, zorder=5)
            
            for ang, val in zip(angles[:-1], maxes1):
                if not np.isnan(val):
                    ax.scatter(ang, val, color='black', s=60, zorder=6, marker='*', label=f'{name1} Max' if i == 0 else "") 
            
            current_maxes = [m for m in maxes1 if not np.isnan(m)]
            if current_maxes:
                max_overall_val = max(max_overall_val, max(current_maxes))

        if not all(np.isnan(m) for m in means2) or not all(np.isnan(m) for m in maxes2):
            ax.plot(angles, means2_closed, color=APPLE_ORANGE, linewidth=2.5, label=f'{name2} Mean')
            ax.fill(angles, means2_closed, color=APPLE_ORANGE, alpha=0.2)
            for ang, val in zip(angles[:-1], means2):
                if not np.isnan(val):
                    ax.scatter(ang, val, color=APPLE_ORANGE, s=50, zorder=5)

            for ang, val in zip(angles[:-1], maxes2):
                if not np.isnan(val):
                    ax.scatter(ang, val, color=APPLE_ORANGE, s=60, zorder=6, marker='X', label=f'{name2} Max' if i == 0 else "") 
            
            current_maxes = [m for m in maxes2 if not np.isnan(m)]
            if current_maxes:
                max_overall_val = max(max_overall_val, max(current_maxes))
                
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)

        labels = [f"XY\n(A:{spec_a[0]}um, B:{spec_b[0]}um)", 
                  f"YZ\n(A:{spec_a[1]}um, B:{spec_b[1]}um)", 
                  f"ZX\n(A:{spec_a[2]}um, B:{spec_b[2]}um)"]
        ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontsize=10, fontweight='bold', color='#1D1D1F')
        
        ax.set_ylim(0, max_overall_val * 1.15)
        ax.set_rticks([10, 20, 30, 40, 50])
        ax.set_yticklabels([])
        
        ax.spines['polar'].set_color(APPLE_GRAY)
        ax.spines['polar'].set_alpha(0.3)
        
        ax.set_title(station, y=1.2, fontsize=14, fontweight='bold', color='#1D1D1F')
        
    handles = []
    labels = []

    if any(not np.isnan(m) for m in [data1.get(s, {}).get('mean', np.nan) for s in all_stations]): 
        h1_mean = Line2D([0], [0], color=APPLE_BLUE, lw=2.5, marker='o', label=f'{name1} Mean')
        handles.append(h1_mean)
        labels.append(f'{name1} Mean')
        if any(not np.isnan(m) for m in [data1.get(s, {}).get('max', np.nan) for s in all_stations]): 
            h1_max = Line2D([0], [0], color='black', marker='*', linestyle='None', markersize=8, label=f'{name1} Max')
            handles.append(h1_max)
            labels.append(f'{name1} Max')

    if any(not np.isnan(m) for m in [data2.get(s, {}).get('mean', np.nan) for s in all_stations]): 
        h2_mean = Line2D([0], [0], color=APPLE_ORANGE, lw=2.5, marker='o', label=f'{name2} Mean')
        handles.append(h2_mean)
        labels.append(f'{name2} Mean')
        if any(not np.isnan(m) for m in [data2.get(s, {}).get('max', np.nan) for s in all_stations]): 
            h2_max = Line2D([0], [0], color=APPLE_ORANGE, marker='X', linestyle='None', markersize=8, label=f'{name2} Max')
            handles.append(h2_max)
            labels.append(f'{name2} Max')

    h_spec_a = Line2D([0], [0], color=APPLE_GREEN, linestyle='--', lw=1.5, label='Grade A Spec')
    handles.append(h_spec_a)
    labels.append('Grade A Spec')
    
    h_spec_b = Line2D([0], [0], color=APPLE_ORANGE, linestyle='--', lw=1.5, label='Grade B Spec')
    handles.append(h_spec_b)
    labels.append('Grade B Spec')
    
    fig.legend(handles=handles, labels=labels,
               loc='upper center', bbox_to_anchor=(0.5, 0.96),
               ncol=4, fontsize=12, frameon=False)
               
    fig.suptitle('Marble Squareness Profile by Station (Unit: µm)', fontsize=18, fontweight='bold', y=0.99, color='#1D1D1F')
    
    fig.tight_layout(rect=[0, 0, 1, 0.92])
    
    return fig_to_base64(fig)

# --- Streamlit App Logic ---

# Helper to save uploaded file to a temporary path
def save_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return None
    try:
        # Create a temporary file with the correct extension
        temp_dir = tempfile.gettempdir()
        file_extension = os.path.splitext(uploaded_file.name)[1]
        with tempfile.NamedTemporaryFile(delete=False, dir=temp_dir, suffix=file_extension) as tmp:
            tmp.write(uploaded_file.getvalue())
            return tmp.name
    except Exception as e:
        st.error(f"Error saving uploaded file: {e}")
        return None

# File Uploaders
uploaded_file1 = st.file_uploader("Site 1 OQ Excel File", type=["xlsx", "xls", "xlsm"], key="file1")
uploaded_file2 = st.file_uploader("Site 2 OQ Excel File", type=["xlsx", "xls", "xlsm"], key="file2")

# Analysis Button
analyze_button = st.button("Start Analysis")

# --- Analysis and Display Logic ---
if analyze_button and uploaded_file1 and uploaded_file2:
    temp_file1_path = None
    temp_file2_path = None
    
    try:
        temp_file1_path = save_uploaded_file(uploaded_file1)
        temp_file2_path = save_uploaded_file(uploaded_file2)

        if not temp_file1_path or not temp_file2_path:
            st.error("Failed to prepare uploaded files for analysis.")
        else:
            with st.spinner("Analyzing data, please wait..."):
                factory1_name = extract_factory_name(uploaded_file1.name)
                factory2_name = extract_factory_name(uploaded_file2.name)

                read_mode1 = 'ipeg' if uploaded_file1.name.lower().startswith('ipeg') else 'default'
                read_mode2 = 'ipeg' if uploaded_file2.name.lower().startswith('ipeg') else 'default'

                df1, err1 = load_excel_data(temp_file1_path, factory1_name, read_mode=read_mode1)
                df2, err2 = load_excel_data(temp_file2_path, factory2_name, read_mode=read_mode2)

                if err1:
                    st.error(f"Failed to load File 1 ({uploaded_file1.name}): {err1}")
                elif err2:
                    st.error(f"Failed to load File 2 ({uploaded_file2.name}): {err2}")
                else:
                    st.success("Files loaded successfully!")

                    # --- Display Summary ---
                    st.subheader("Analysis Summary")
                    col1, col2 = st.columns(2)
                    
                    cnc_col1 = get_cnc_column_name(df1)
                    stations1 = int(df1[cnc_col1].nunique()) if cnc_col1 and cnc_col1 in df1.columns else 0
                    cnc_col2 = get_cnc_column_name(df2)
                    stations2 = int(df2[cnc_col2].nunique()) if cnc_col2 and cnc_col2 in df2.columns else 0

                    with col1:
                        st.markdown(f"### 📊 {factory1_name}")
                        st.write(f"**Machine Count:** {len(df1)}")
                        st.write(f"**CNC Stations:** {stations1}")

                    with col2:
                        st.markdown(f"### 📊 {factory2_name}")
                        st.write(f"**Machine Count:** {len(df2)}")
                        st.write(f"**CNC Stations:** {stations2}")

                    # --- Perform Analyses and Display Charts ---
                    st.subheader("Detailed Analysis & Charts")
                    
                    # Machine Count
                    try:
                        img1, data1 = compare_machine_count(df1, df2, factory1_name, factory2_name)
                        st.markdown(f"### Machine Type Count Comparison")
                        st.image(img1, caption=f'{factory1_name} vs {factory2_name}', use_column_width=True)
                        # Optionally display detailed data as a table
                        # if data1:
                        #     st.dataframe(pd.DataFrame(data1))
                    except Exception as e:
                        st.error(f"Error generating Machine Count chart: {e}")

                    # Age Comparison
                    try:
                        img2, age_stats = compare_machine_age(df1, df2, factory1_name, factory2_name)
                        st.markdown(f"### Machine Age Distribution")
                        st.image(img2, caption=f'{factory1_name} vs {factory2_name}', use_column_width=True)
                        
                        if age_stats:
                            st.subheader("Age Statistics")
                            age_col1, age_col2, age_col3 = st.columns(3)
                            with age_col1:
                                st.metric(label=f"{factory1_name} Avg Age", value=f"{age_stats['factory1']['mean']} yrs")
                                st.metric(label=f"{factory1_name} Median Age", value=f"{age_stats['factory1']['median']} yrs")
                            with age_col2:
                                st.metric(label=f"{factory2_name} Avg Age", value=f"{age_stats['factory2']['mean']} yrs")
                                st.metric(label=f"{factory2_name} Median Age", value=f"{age_stats['factory2']['median']} yrs")
                            if 'ttest' in age_stats and age_stats['ttest']:
                                with age_col3:
                                    st.metric(label="T-test p-value", value=f"{age_stats['ttest']['p_value']:.4f}")
                                    sig_color = "#FF3B30" if age_stats['ttest']['significant'] == 'significant' else "#34C759"
                                    st.markdown(f"**Significance:** <span style='color:{sig_color}; font-weight:bold;'>{age_stats['ttest']['significant'].capitalize()}</span>", unsafe_allow_html=True)

                    except Exception as e:
                        st.error(f"Error generating Age Comparison chart: {e}")

                    # Runout Comparison
                    try:
                        img3 = compare_spindle_runout(df1, df2, factory1_name, factory2_name)
                        st.markdown(f"### Spindle Runout Process Capability")
                        st.image(img3, caption=f'{factory1_name} vs {factory2_name}', use_column_width=True)
                    except Exception as e:
                        st.error(f"Error generating Spindle Runout chart: {e}")
                    
                    # Velocity Comparison
                    try:
                        img4 = compare_spindle_velocity(df1, df2, factory1_name, factory2_name)
                        st.markdown(f"### Spindle Velocity Vibration")
                        st.image(img4, caption=f'{factory1_name} vs {factory2_name}', use_column_width=True)
                    except Exception as e:
                        st.error(f"Error generating Spindle Velocity chart: {e}")
                    
                    # Acceleration Comparison
                    try:
                        img5 = compare_spindle_acceleration(df1, df2, factory1_name, factory2_name)
                        st.markdown(f"### Spindle Acceleration Vibration")
                        st.image(img5, caption=f'{factory1_name} vs {factory2_name}', use_column_width=True)
                    except Exception as e:
                        st.error(f"Error generating Spindle Acceleration chart: {e}")

                    # Marble Squareness Comparison
                    try:
                        img6 = compare_marble_squareness_combined(df1, df2, factory1_name, factory2_name)
                        st.markdown(f"### Marble Squareness Comprehensive Profile")
                        st.image(img6, caption=f'{factory1_name} vs {factory2_name}', use_column_width=True)
                    except Exception as e:
                        st.error(f"Error generating Marble Squareness chart: {e}")

    except Exception as e:
        st.error(f"An unexpected error occurred during analysis: {e}")
        traceback.print_exc() # For debugging in Streamlit logs

    finally:
        # Clean up temporary files
        if temp_file1_path and os.path.exists(temp_file1_path):
            os.remove(temp_file1_path)
        if temp_file2_path and os.path.exists(temp_file2_path):
            os.remove(temp_file2_path)

elif analyze_button: # Button clicked but files not uploaded
    st.warning("Please upload both Excel files to start the analysis.")

# --- Footer ---
st.markdown("---")
st.markdown("Supported formats: .xlsx, .xls, .xlsm | Required columns: CNC OP, Machine Model, Year of manufacturer")


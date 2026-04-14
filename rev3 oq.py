# 4. Squareness Compliance (Station-based diagnostic generation)
    def analyze_squareness_compliance(df, cnc_col):
        if not cnc_col: return "No Data", []
        directions = ['XY', 'YZ', 'ZX']
        
        # 分别用来记录每个 Station 真正超差 (Out of Spec) 和 降级 (Grade B) 的机台数量
        station_oos_counts = defaultdict(int)
        station_gb_counts = defaultdict(int)
        has_data = False
        
        sq_cols = [c for c in df.columns if 'squareness' in str(c).lower() or '垂直度' in str(c)]
        
        for station in df[cnc_col].dropna().unique():
            station_df = df[df[cnc_col] == station]
            
            oos_count = 0
            gb_count = 0
            
            # 以每台机台（每一行数据）为单位进行遍历检验
            for idx, row in station_df.iterrows():
                is_oos = False
                is_gb = False
                
                for plane in directions:
                    plane_cols = [c for c in sq_cols if plane in str(c) or plane[::-1] in str(c)]
                    for c in plane_cols:
                        raw_v = pd.to_numeric(row[c], errors='coerce')
                        if pd.notna(raw_v):
                            has_data = True
                            # 统一转换单位为 μm
                            v = raw_v * 1000 if raw_v < 1 else raw_v
                            
                            # 判断是否 Out of spec
                            if plane == 'XY' and v > 20.0: is_oos = True
                            elif plane in ['YZ', 'ZX'] and v > 30.0: is_oos = True
                            # 判断是否 Grade B
                            elif plane == 'XY' and v > 16.0: is_gb = True
                            elif plane in ['YZ', 'ZX'] and v > 20.0: is_gb = True
                
                # 如果这台机器超差，单独计数+1
                if is_oos:
                    oos_count += 1
                elif is_gb:
                    gb_count += 1
                    
            if oos_count > 0:
                station_oos_counts[station] = oos_count
            if gb_count > 0:
                station_gb_counts[station] = gb_count
                
        if not has_data: return "No Data", []
        
        # 组装最终结果
        if station_oos_counts:
            status = f"<span style='color: {THEME_RED}; font-weight: bold;'>Out of Spec</span>"
            details = [f"{st} (x{count} machines)" for st, count in station_oos_counts.items()]
            return status, details
        elif station_gb_counts:
            status = f"<span style='color: {THEME_ORANGE}; font-weight: bold;'>Grade B</span>"
            details = [f"{st} (x{count} machines)" for st, count in station_gb_counts.items()]
            return status, details
        else:
            return f"<span style='color: {THEME_GREEN}; font-weight: bold;'>Grade A</span>", []

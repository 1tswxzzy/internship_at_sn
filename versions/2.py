import pandas as pd
import os
import sys
import re
import numpy as np
from collections import OrderedDict

def classify_school_type(file_path):
    """
    根据文件名关键字识别学校类型
    """
    file_name = os.path.basename(file_path)
    
    if '小学' in file_name:
        return '小学'
    else:
        return '中学'

def get_excel_info(file_path):
    """
    获取Excel文件的有效行数和列数（从第9行开始计算，过滤第四列为0或空的行）
    """
    try:
        df = pd.read_excel(file_path, header=None)
        
        if len(df) < 9:
            print("警告：文件行数少于9行，可能没有有效数据。")
            return 0, 0, None
        
        df_valid = df.iloc[8:].copy()
        df_valid = df_valid.reset_index(drop=True)
        
        if len(df_valid.columns) >= 4:
            col_d = df_valid.iloc[:, 3]
            filter_condition = col_d.notna() & (col_d != 0)
            col_d_str = col_d.astype(str)
            filter_condition = filter_condition & (col_d_str != '0') & (col_d_str != '0.0')
            df_filtered = df_valid[filter_condition].copy()
            df_filtered = df_filtered.reset_index(drop=True)
            
            removed_count = len(df_valid) - len(df_filtered)
            if removed_count > 0:
                print(f"已过滤第四列为0或空的行：移除 {removed_count} 行，保留 {len(df_filtered)} 行")
        else:
            df_filtered = df_valid
        
        valid_rows = df_filtered.notna().any(axis=1)
        if valid_rows.any():
            last_valid_row_index = valid_rows[valid_rows].index[-1]
            max_h = last_valid_row_index + 1
        else:
            max_h = 0
        
        valid_cols = df_filtered.notna().any(axis=0)
        if valid_cols.any():
            last_valid_col_index = valid_cols[valid_cols].index[-1]
            max_l = last_valid_col_index + 1
        else:
            max_l = 0
        
        return max_h, max_l, df_filtered
    
    except Exception as e:
        print(f"读取文件时发生错误：{e}")
        return None, None, None

def extract_district(text):
    """
    从文本中提取区县名称
    """
    if pd.isna(text) or text == '':
        return None
    
    text = str(text).strip()
    
    district_patterns = [
        r'([^\s,，、]+(?:区|县|市))',
        r'([^\s,，、]+(?:自治县|自治旗|县级市|地级市))',
    ]
    
    for pattern in district_patterns:
        matches = re.findall(pattern, text)
        if matches:
            return matches[0]
    
    if len(text) < 15 and text not in ['合计', '总计', '小计', '备注']:
        return text
    
    return None

def calculate_all_indicators(df_filtered, school_type):
    """
    计算所有指标：指标a（第四列）、指标b、指标c（第十九列）、指标d（第十、十一列之和）
    """
    if df_filtered is None or len(df_filtered) == 0:
        return None
    
    try:
        min_cols = 8 if school_type == '小学' else 7
        if len(df_filtered.columns) < max(min_cols, 19):
            print(f"警告：数据只有{len(df_filtered.columns)}列，需要至少{max(min_cols, 19)}列来计算所有指标")
            return None
        
        districts = []
        for idx in range(len(df_filtered)):
            district_text = df_filtered.iloc[idx, 2]
            district = extract_district(district_text)
            districts.append(district if district else f"未知区县_{idx}")
        
        # 计算指标a（第四列，索引3）
        indicator_a = pd.to_numeric(df_filtered.iloc[:, 3], errors='coerce')
        
        # 计算指标b
        if school_type == '小学':
            # 小学类：第六、七、八列之和（索引5,6,7）
            col_6 = pd.to_numeric(df_filtered.iloc[:, 5], errors='coerce').fillna(0)
            col_7 = pd.to_numeric(df_filtered.iloc[:, 6], errors='coerce').fillna(0)
            col_8 = pd.to_numeric(df_filtered.iloc[:, 7], errors='coerce').fillna(0)
            indicator_b = col_6 + col_7 + col_8
        else:
            # 中学类：第六、七列之和（索引5,6）
            col_6 = pd.to_numeric(df_filtered.iloc[:, 5], errors='coerce').fillna(0)
            col_7 = pd.to_numeric(df_filtered.iloc[:, 6], errors='coerce').fillna(0)
            indicator_b = col_6 + col_7
        
        # 计算指标c（第十九列，索引18）
        indicator_c = pd.to_numeric(df_filtered.iloc[:, 18], errors='coerce').fillna(0)
        
        # 计算指标d（第十、十一列之和，索引9,10）
        col_10 = pd.to_numeric(df_filtered.iloc[:, 9], errors='coerce').fillna(0)
        col_11 = pd.to_numeric(df_filtered.iloc[:, 10], errors='coerce').fillna(0)
        indicator_d = col_10 + col_11
        
        # 创建指标数据框
        indicators_df = pd.DataFrame({
            'district': districts,
            'indicator_a': indicator_a,
            'indicator_b': indicator_b,
            'indicator_c': indicator_c,
            'indicator_d': indicator_d
        })
        
        return indicators_df
    
    except Exception as e:
        print(f"计算指标时发生错误：{e}")
        return None

def calculate_statistics_for_indicator(indicators_df, district_groups, indicator_type, school_type):
    """
    为指定指标计算统计值（X, S, CV）
    """
    if indicators_df is None or not district_groups:
        return {}
    
    stats_dict = {}
    indicator_col = f'indicator_{indicator_type}'
    
    for district, group_info in district_groups.items():
        district_data = indicators_df[indicators_df['district'] == district]
        
        if len(district_data) == 0:
            continue
        
        A = district_data['indicator_a'].sum()
        indicator_sum = district_data[indicator_col].sum()
        
        if A == 0:
            continue
        
        X = (indicator_sum * 100) / A
        
        weighted_variance_sum = 0
        valid_count = 0
        raw_data = []
        
        for _, row in district_data.iterrows():
            a = row['indicator_a']
            indicator_value = row[indicator_col]
            
            if a > 0:
                row_ratio = (indicator_value / a) * 100
                deviation_squared = (row_ratio - X) ** 2
                weighted_deviation = (a / A) * deviation_squared
                weighted_variance_sum += weighted_deviation
                valid_count += 1
                
                raw_data.append({
                    'a': a,
                    'indicator_value': indicator_value,
                    'ratio': row_ratio
                })
        
        S = np.sqrt(weighted_variance_sum) if weighted_variance_sum > 0 else 0
        CV = round((S / X), 2) if X != 0 else 0.00
        
        stats_dict[district] = {
            'group': group_info['group'],
            'A': A,
            'indicator_sum': indicator_sum,
            'X': X,
            'S': S,
            'CV': CV,
            'data_count': len(district_data),
            'valid_count': valid_count,
            'raw_data': raw_data
        }
    
    return stats_dict

def calculate_group_statistics(indicators_df, district_groups, school_type):
    """
    按区县分组计算所有指标的统计值
    """
    if indicators_df is None or not district_groups:
        return {}
    
    b_stats = calculate_statistics_for_indicator(indicators_df, district_groups, 'b', school_type)
    c_stats = calculate_statistics_for_indicator(indicators_df, district_groups, 'c', school_type)
    d_stats = calculate_statistics_for_indicator(indicators_df, district_groups, 'd', school_type)
    
    return {
        'b_stats': b_stats,
        'c_stats': c_stats,
        'd_stats': d_stats
    }

def display_formatted_table(all_stats, district_groups, school_type, file_name=""):
    """
    以要求的格式显示单文件统计表格（包含指标B、C、D）
    """
    if not all_stats or not district_groups:
        return
    
    b_stats = all_stats.get('b_stats', {})
    c_stats = all_stats.get('c_stats', {})
    d_stats = all_stats.get('d_stats', {})
    
    if not b_stats and not c_stats and not d_stats:
        return
    
    title = f"{school_type}类统计结果"
    if file_name:
        title += f" - {file_name}"
    
    print(f"\n{'='*150}")
    print(f"{title:^150}")
    print(f"{'='*150}")
    
    # 表头
    print(f"\n{'组别':<6} {'区县':<12} {'指标':<20} {'B(高于规定学历教师)':<30} {'C(省级以上骨干教师)':<30} {'D(指标D)':<30}")
    print(f"{'':<6} {'':<12} {'':<20} {'':<30} {'':<30} {'':<30}")
    print(f"{'-'*150}")
    
    # 按照district_groups的顺序显示
    for district, group_info in district_groups.items():
        group_name = group_info['group']
        
        b_stat = b_stats.get(district, {})
        c_stat = c_stats.get(district, {})
        d_stat = d_stats.get(district, {})
        
        b_X = b_stat.get('X', 0)
        b_S = b_stat.get('S', 0)
        b_CV = b_stat.get('CV', 0)
        
        c_X = c_stat.get('X', 0)
        c_S = c_stat.get('S', 0)
        c_CV = c_stat.get('CV', 0)
        
        d_X = d_stat.get('X', 0)
        d_S = d_stat.get('S', 0)
        d_CV = d_stat.get('CV', 0)
        
        print(f"{group_name:<6} {district:<12} {'X(平均值)':<20} {b_X:<30.4f} {c_X:<30.4f} {d_X:<30.4f}")
        print(f"{'':<6} {'':<12} {'S(标准差)':<20} {b_S:<30.6f} {c_S:<30.6f} {d_S:<30.6f}")
        print(f"{'':<6} {'':<12} {'CV(差异系数)':<20} {b_CV:<30.2f} {c_CV:<30.2f} {d_CV:<30.2f}")
        print(f"{'-'*150}")

def identify_districts_from_column_c(df_filtered, school_type):
    """
    识别Excel文件第三列的区县信息，保持原始顺序
    """
    if df_filtered is None or len(df_filtered) == 0:
        print("没有有效数据进行区县识别。")
        return {}, [], None
    
    try:
        if len(df_filtered.columns) < 4:
            print("警告：数据列数不足")
            return {}, [], None
        
        district_groups = OrderedDict()
        districts_list = []
        group_counter = 1
        
        for idx in range(len(df_filtered)):
            cell_value = df_filtered.iloc[idx, 2]
            
            if pd.notna(cell_value) and str(cell_value).strip() != '':
                district = extract_district(cell_value)
                
                if district and district not in district_groups:
                    group_name = f"T{group_counter}"
                    district_groups[district] = {
                        'group': group_name,
                        'school_type': school_type,
                        'first_row': idx + 1,
                        'original_value': str(cell_value).strip(),
                        'order': group_counter
                    }
                    
                    districts_list.append({
                        'district': district,
                        'group': group_name,
                        'school_type': school_type,
                        'first_row': idx + 1,
                        'order': group_counter
                    })
                    
                    group_counter += 1
        
        print(f"共识别到 {len(district_groups)} 个区县：", end=" ")
        for info in districts_list:
            print(f"{info['group']}:{info['district']}", end=" ")
        print()
        
        indicators_df = calculate_all_indicators(df_filtered, school_type)
        
        if indicators_df is not None:
            print(f"指标计算完成，共{len(indicators_df)}条有效数据")
        
        return district_groups, districts_list, indicators_df
    
    except Exception as e:
        print(f"识别区县时发生错误：{e}")
        return {}, [], None

def calculate_overall_statistics(all_raw_data):
    """
    基于所有原始数据计算总体统计值
    """
    if not all_raw_data:
        return {'X': 0, 'S': 0, 'CV': 0}
    
    total_A = sum(data['a'] for data in all_raw_data)
    total_indicator = sum(data['indicator_value'] for data in all_raw_data)
    
    if total_A == 0:
        return {'X': 0, 'S': 0, 'CV': 0}
    
    overall_X = (total_indicator * 100) / total_A
    
    weighted_variance_sum = 0
    valid_count = 0
    
    for data in all_raw_data:
        a = data['a']
        indicator_value = data['indicator_value']
        
        if a > 0:
            row_ratio = (indicator_value / a) * 100
            deviation_squared = (row_ratio - overall_X) ** 2
            weighted_deviation = (a / total_A) * deviation_squared
            weighted_variance_sum += weighted_deviation
            valid_count += 1
    
    overall_S = np.sqrt(weighted_variance_sum) if weighted_variance_sum > 0 else 0
    overall_CV = round((overall_S / overall_X), 2) if overall_X != 0 else 0.00
    
    return {
        'X': overall_X,
        'S': overall_S,
        'CV': overall_CV,
        'total_A': total_A,
        'total_indicator': total_indicator,
        'valid_count': valid_count
    }

def build_district_order(all_results, school_type):
    """
    构建区县顺序映射，基于第一个文件中区县出现的顺序
    """
    results = all_results[school_type]
    district_order = {}
    order_counter = 1
    
    for file_result in results['files']:
        if 'districts_list' in file_result:
            for district_info in file_result['districts_list']:
                district = district_info['district']
                if district not in district_order:
                    district_order[district] = {
                        'order': district_info.get('order', order_counter),
                        'group': district_info['group']
                    }
                    order_counter += 1
    
    return district_order

def display_overall_cv_table(all_results, school_type):
    """
    显示总体汇总报告 - 只显示差异系数CV，包含指标B、C、D
    """
    results = all_results[school_type]
    if not results['files']:
        return
    
    # 构建区县顺序映射
    district_order = build_district_order(all_results, school_type)
    
    # 收集所有区县及其原始数据
    all_districts_data = {}
    
    for file_result in results['files']:
        if 'all_statistics' not in file_result:
            continue
        
        # 处理指标b
        for district, stats in file_result['all_statistics']['b_stats'].items():
            if district not in all_districts_data:
                all_districts_data[district] = {'b_raw': [], 'c_raw': [], 'd_raw': []}
            if 'raw_data' in stats:
                all_districts_data[district]['b_raw'].extend(stats['raw_data'])
        
        # 处理指标c
        for district, stats in file_result['all_statistics']['c_stats'].items():
            if district not in all_districts_data:
                all_districts_data[district] = {'b_raw': [], 'c_raw': [], 'd_raw': []}
            if 'raw_data' in stats:
                all_districts_data[district]['c_raw'].extend(stats['raw_data'])
        
        # 处理指标d
        for district, stats in file_result['all_statistics']['d_stats'].items():
            if district not in all_districts_data:
                all_districts_data[district] = {'b_raw': [], 'c_raw': [], 'd_raw': []}
            if 'raw_data' in stats:
                all_districts_data[district]['d_raw'].extend(stats['raw_data'])
    
    if not all_districts_data:
        return
    
    # 按照原始顺序排序区县
    sorted_districts = sorted(all_districts_data.keys(), 
                            key=lambda x: district_order.get(x, {}).get('order', 999))
    
    # 表格标题
    print(f"\n{'='*120}")
    print(f"{school_type}类差异系数(CV)汇总")
    print(f"{'='*120}")
    
    # 表头
    print(f"\n{'组别':<6} {'区县':<15} {'B(高于规定学历教师)':<30} {'C(省级以上骨干教师)':<30} {'D(指标D)':<30}")
    print(f"{'':<6} {'':<15} {'差异系数CV':<30} {'差异系数CV':<30} {'差异系数CV':<30}")
    print(f"{'-'*120}")
    
    # 数据行 - 按照原始顺序
    for district in sorted_districts:
        group_name = district_order.get(district, {}).get('group', 'T?')
        data = all_districts_data[district]
        
        # 计算各指标总体CV
        b_overall = calculate_overall_statistics(data['b_raw'])
        c_overall = calculate_overall_statistics(data['c_raw'])
        d_overall = calculate_overall_statistics(data['d_raw'])
        
        # 只显示CV
        print(f"{group_name:<6} {district:<15} {b_overall['CV']:<30.2f} {c_overall['CV']:<30.2f} {d_overall['CV']:<30.2f}")
    
    print(f"{'-'*120}")
    
    # 计算总计CV
    total_all_b_raw = []
    total_all_c_raw = []
    total_all_d_raw = []
    for district_data in all_districts_data.values():
        total_all_b_raw.extend(district_data['b_raw'])
        total_all_c_raw.extend(district_data['c_raw'])
        total_all_d_raw.extend(district_data['d_raw'])
    
    total_b_overall = calculate_overall_statistics(total_all_b_raw)
    total_c_overall = calculate_overall_statistics(total_all_c_raw)
    total_d_overall = calculate_overall_statistics(total_all_d_raw)
    
    # 总计行
    print(f"{'':<6} {'总计':<15} {total_b_overall['CV']:<30.2f} {total_c_overall['CV']:<30.2f} {total_d_overall['CV']:<30.2f}")
    print(f"{'='*120}")

def process_multiple_files():
    """
    处理多个Excel文件，按学校类型分类统计
    """
    all_results = {
        '小学': {
            'files': []
        },
        '中学': {
            'files': []
        }
    }
    
    while True:
        file_path = input("\n请输入Excel文件的完整路径（或拖拽文件到此处，输入'q'结束输入）：").strip()
        
        if file_path.lower() == 'q':
            break
        
        file_path = file_path.strip('"').strip("'")
        
        if not os.path.exists(file_path):
            print(f"文件路径 '{file_path}' 不存在，请重新输入。")
            continue
        
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
            print(f"警告：文件扩展名为 '{file_ext}'，可能不是Excel文件，尝试读取...")
        
        school_type = classify_school_type(file_path)
        
        print(f"\n{'='*120}")
        print(f"正在处理文件：{os.path.basename(file_path)}")
        print(f"识别学校类型：{school_type}")
        
        max_h, max_l, df_filtered = get_excel_info(file_path)
        
        if max_h is not None and max_l is not None and df_filtered is not None:
            print(f"有效数据行数 (max_h)：{max_h} 行")
            print(f"有效数据列数 (max_l)：{max_l} 列")
            print(f"过滤后数据行数：{len(df_filtered)} 行")
            
            if max_h > 0 and max_l >= 11:  # 需要至少11列来计算指标d
                district_groups, districts_list, indicators_df = identify_districts_from_column_c(df_filtered, school_type)
                
                file_result = {
                    'file_name': os.path.basename(file_path),
                    'file_path': file_path,
                    'school_type': school_type,
                    'max_h': max_h,
                    'max_l': max_l,
                    'filtered_rows': len(df_filtered),
                    'district_groups': district_groups,
                    'districts_list': districts_list
                }
                
                if indicators_df is not None:
                    all_stats = calculate_group_statistics(indicators_df, district_groups, school_type)
                    file_result['all_statistics'] = all_stats
                    file_result['indicators_df'] = indicators_df
                    
                    # 显示单文件格式化表格
                    display_formatted_table(all_stats, district_groups, school_type, os.path.basename(file_path))
                
                all_results[school_type]['files'].append(file_result)
            else:
                print(f"文件列数不足11列，无法计算指标d，跳过处理。")
        else:
            print(f"文件分析失败。")
    
    return all_results

def main():
    """
    主函数
    """
    print("="*120)
    print("Excel文件信息分析工具 - 多指标差异系数(CV)汇总报告")
    print("="*120)
    print("功能说明：")
    print("  - 从第9行开始识别数据，过滤第四列为0或空的行")
    print("  - 识别第三列区县信息并分组（保持原始顺序）")
    print("  - 指标A：第四列（学生数）")
    print("  - 指标B(每百名学生拥有高于规定学历的教师)：")
    print("    小学类 = 第六列 + 第七列 + 第八列")
    print("    中学类 = 第六列 + 第七列")
    print("  - 指标C(每百名学生拥有省级以上骨干教师) = 第十九列")
    print("  - 指标D = 第十列 + 第十一列")
    print("  - 计算公式：X = 指标和*100/A, CV = S/X (保留两位小数)")
    print("  - 单文件显示完整统计：X(平均值)、S(标准差)、CV(差异系数)")
    print("  - 总体汇总只显示差异系数CV，保持原始组别顺序")
    print("  - 文件名包含'小学'的归为小学类，其余归为中学类")
    print("="*120)
    
    all_results = process_multiple_files()
    
    # 显示总体汇总 - 只显示CV，包含指标B、C、D
    if all_results['小学']['files'] or all_results['中学']['files']:
        print("\n" + "="*120)
        print("总体汇总报告 - 差异系数(CV)")
        print("="*120)
        
        if all_results['小学']['files']:
            display_overall_cv_table(all_results, '小学')
        
        if all_results['中学']['files']:
            display_overall_cv_table(all_results, '中学')
    else:
        print("\n未处理任何文件。")
    
    input("\n按 Enter 键退出程序...")
    sys.exit(0)

if __name__ == "__main__":
    main()
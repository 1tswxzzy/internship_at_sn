import pandas as pd
import os
import sys
import re

def classify_school_type(file_path):
    """
    根据文件名关键字识别学校类型
    
    参数:
        file_path: Excel文件路径
    
    返回:
        school_type: 学校类型（'小学' 或 '中学'）
    """
    # 获取文件名（不含路径）
    file_name = os.path.basename(file_path)
    
    # 检查文件名中是否包含"小学"关键字
    if '小学' in file_name:
        return '小学'
    else:
        return '中学'

def get_excel_info(file_path):
    """
    获取Excel文件的有效行数和列数（从第9行开始计算，过滤第四列为0或空的行）
    
    参数:
        file_path: Excel文件路径
    
    返回:
        max_h: 有效总行数（过滤后的有效数据行数）
        max_l: 有效总列数
        df_filtered: 过滤后的数据框，用于后续统计
    """
    try:
        # 读取Excel文件，不跳过空行，不将第一行作为列名
        df = pd.read_excel(file_path, header=None)
        
        # 从第9行开始（索引为8）获取数据
        if len(df) < 9:
            print("警告：文件行数少于9行，可能没有有效数据。")
            return 0, 0, None
        
        # 提取第9行及以后的数据
        df_valid = df.iloc[8:].copy()
        
        # 重置索引以便后续处理
        df_valid = df_valid.reset_index(drop=True)
        
        # 过滤第四列（列索引为3）为0或为空的行
        if len(df_valid.columns) >= 4:
            # 获取第四列数据
            col_d = df_valid.iloc[:, 3]
            
            # 创建过滤条件：第四列不为空且不为0
            filter_condition = col_d.notna() & (col_d != 0)
            
            # 处理可能的字符串'0'和'0.0'情况
            col_d_str = col_d.astype(str)
            filter_condition = filter_condition & (col_d_str != '0') & (col_d_str != '0.0')
            
            # 应用过滤条件
            df_filtered = df_valid[filter_condition].copy()
            
            # 重置索引
            df_filtered = df_filtered.reset_index(drop=True)
            
            # 显示过滤信息
            removed_count = len(df_valid) - len(df_filtered)
            if removed_count > 0:
                print(f"已过滤第四列为0或空的行：移除 {removed_count} 行，保留 {len(df_filtered)} 行")
        else:
            print("警告：文件列数少于4列，无法进行第四列过滤。")
            df_filtered = df_valid
        
        # 获取有效行数
        valid_rows = df_filtered.notna().any(axis=1)
        
        if valid_rows.any():
            last_valid_row_index = valid_rows[valid_rows].index[-1]
            max_h = last_valid_row_index + 1
        else:
            max_h = 0
        
        # 获取有效列数
        valid_cols = df_filtered.notna().any(axis=0)
        
        if valid_cols.any():
            last_valid_col_index = valid_cols[valid_cols].index[-1]
            max_l = last_valid_col_index + 1
        else:
            max_l = 0
        
        return max_h, max_l, df_filtered
    
    except FileNotFoundError:
        print(f"错误：找不到文件 '{file_path}'，请检查路径是否正确。")
        return None, None, None
    except Exception as e:
        print(f"读取文件时发生错误：{e}")
        return None, None, None

def extract_district(text):
    """
    从文本中提取区县名称
    
    参数:
        text: 可能包含区县信息的文本
    
    返回:
        district: 提取的区县名称，如果未找到则返回None
    """
    if pd.isna(text) or text == '':
        return None
    
    text = str(text).strip()
    
    # 常见的区县名称模式
    district_patterns = [
        r'([^\s,，、]+(?:区|县|市))',  # 匹配以区、县、市结尾的词
        r'([^\s,，、]+(?:自治县|自治旗|县级市|地级市))',  # 匹配特殊行政区
    ]
    
    for pattern in district_patterns:
        matches = re.findall(pattern, text)
        if matches:
            return matches[0]  # 返回第一个匹配的区县名称
    
    # 如果没有匹配到标准的区县名称，检查是否可能是简化的区县名
    if len(text) < 15 and text not in ['合计', '总计', '小计', '备注']:
        return text
    
    return None

def calculate_district_totals(df_filtered):
    """
    计算每个区县的第四列总数
    
    参数:
        df_filtered: 过滤后的数据框
    
    返回:
        district_totals: 字典，键为区县名称，值为总数
    """
    if df_filtered is None or len(df_filtered) == 0:
        return {}
    
    district_totals = {}
    
    try:
        # 确保第四列是数值类型
        col_d_numeric = pd.to_numeric(df_filtered.iloc[:, 3], errors='coerce')
        
        # 遍历每一行
        for idx in range(len(df_filtered)):
            # 获取第三列的区县名称
            district_text = df_filtered.iloc[idx, 2] if len(df_filtered.columns) > 2 else None
            district = extract_district(district_text)
            
            if district and not pd.isna(col_d_numeric.iloc[idx]):
                value = col_d_numeric.iloc[idx]
                # 累加该区县的总数
                if district in district_totals:
                    district_totals[district] += value
                else:
                    district_totals[district] = value
        
        return district_totals
    
    except Exception as e:
        print(f"计算区县总数时发生错误：{e}")
        return {}

def identify_districts_from_column_c(df_filtered, school_type):
    """
    仅识别Excel文件第三列的区县信息（从过滤后的数据中），并按顺序分配组别，同时统计总数
    
    参数:
        df_filtered: 过滤后的数据框
        school_type: 学校类型（'小学' 或 '中学'）
    
    返回:
        district_groups: 字典，键为区县名称，值为组别信息和统计数据
        districts_list: 列表，按顺序存储区县信息（包含总数）
        district_totals: 字典，键为区县名称，值为总数
    """
    if df_filtered is None or len(df_filtered) == 0:
        print("没有有效数据进行区县识别。")
        return {}, [], {}
    
    try:
        # 检查列数
        if len(df_filtered.columns) < 4:
            print("警告：数据列数不足，无法识别第三列区县信息和统计第四列。")
            return {}, [], {}
        
        # 确保第四列是数值类型
        col_d_numeric = pd.to_numeric(df_filtered.iloc[:, 3], errors='coerce')
        
        # 存储区县信息和分组
        district_groups = {}
        districts_list = []
        group_counter = 1
        
        
        # 用于累计每个区县的总数
        running_totals = {}
        
        # 遍历所有行
        for idx in range(len(df_filtered)):
            # 获取第三列的区县名称
            cell_value = df_filtered.iloc[idx, 2]
            
            if pd.notna(cell_value) and str(cell_value).strip() != '':
                district = extract_district(cell_value)
                
                if district:
                    # 获取第四列的数值
                    col_d_value = col_d_numeric.iloc[idx] if idx < len(col_d_numeric) else 0
                    
                    # 累计总数
                    if district in running_totals:
                        running_totals[district] += col_d_value
                    else:
                        running_totals[district] = col_d_value
                    
                    # 检查是否第一次出现该区县
                    if district not in district_groups:
                        # 分配组别（包含学校类型信息）
                        group_name = f"T{group_counter}"
                        district_groups[district] = {
                            'group': group_name,
                            'school_type': school_type,
                            'first_row': idx + 1,  # 在过滤后数据中的行号
                            'original_value': str(cell_value).strip(),
                            'entries': []  # 存储该区县的所有条目
                        }
                        
                        districts_list.append({
                            'district': district,
                            'group': group_name,
                            'school_type': school_type,
                            'first_row': idx + 1,
                            'original_value': str(cell_value).strip(),
                            'total': 0  # 初始化总数为0
                        })
                        group_counter += 1
                   
                    # 记录该区县的每个条目
                    district_groups[district]['entries'].append({
                        'row': idx + 1,
                        'value': col_d_value
                    })
        
        # 更新每个区县的总数
        for district_info in districts_list:
            district_info['total'] = running_totals.get(district_info['district'], 0)
        
        # 计算所有区县的总和
        total_sum = sum(running_totals.values())
        
        print("-" * 80)
        print(f"共识别到 {len(district_groups)} 个区县，所有区县总计: {total_sum:.2f}")
        
        return district_groups, districts_list, running_totals
    
    except Exception as e:
        print(f"识别区县时发生错误：{e}")
        return {}, [], {}

def process_multiple_files():
    """
    处理多个Excel文件，按学校类型分类统计
    """
    all_results = {
        '小学': {
            'files': [],
            'district_totals': {},
            'overall_total': 0
        },
        '中学': {
            'files': [],
            'district_totals': {},
            'overall_total': 0
        }
    }
    
    while True:
        file_path = input("\n请输入Excel文件的完整路径（或拖拽文件到此处，输入'q'结束输入）：").strip()
        
        if file_path.lower() == 'q':
            break
        
        # 去除可能的引号
        file_path = file_path.strip('"').strip("'")
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"文件路径 '{file_path}' 不存在，请重新输入。")
            continue
        
        # 检查文件扩展名
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
            print(f"警告：文件扩展名为 '{file_ext}'，可能不是Excel文件，尝试读取...")
        
        # 识别学校类型
        school_type = classify_school_type(file_path)
        
        print(f"\n{'='*80}")
        print(f"正在处理文件：{os.path.basename(file_path)}")
        print(f"识别学校类型：{school_type}")
        print(f"{'='*80}")
        
        # 获取文件基本信息
        max_h, max_l, df_filtered = get_excel_info(file_path)
        
        if max_h is not None and max_l is not None and df_filtered is not None:
            print(f"\n文件基本信息（从第9行开始，已过滤第四列为0或空的数据）：")
            print(f"有效数据行数 (max_h)：{max_h} 行")
            print(f"有效数据列数 (max_l)：{max_l} 列")
            print(f"过滤后数据行数：{len(df_filtered)} 行")
            
            # 识别区县并统计总数
            if max_h > 0 and max_l >= 4:
                district_groups, districts_list, district_totals = identify_districts_from_column_c(df_filtered, school_type)
                
                # 存储结果
                file_result = {
                    'file_name': os.path.basename(file_path),
                    'file_path': file_path,
                    'school_type': school_type,
                    'max_h': max_h,
                    'max_l': max_l,
                    'filtered_rows': len(df_filtered),
                    'district_groups': district_groups,
                    'districts_list': districts_list,
                    'district_totals': district_totals
                }
                
                all_results[school_type]['files'].append(file_result)
                
                # 累加到总体统计
                for district, total in district_totals.items():
                    if district in all_results[school_type]['district_totals']:
                        all_results[school_type]['district_totals'][district] += total
                    else:
                        all_results[school_type]['district_totals'][district] = total
                
                all_results[school_type]['overall_total'] += sum(district_totals.values())
                
                # 显示该文件的汇总信息
                if district_groups and districts_list:
                    print(f"\n{school_type} - {os.path.basename(file_path)} 区县总数汇总：")
                    print("-" * 60)
                    print(f"{'组别':<6} {'区县名称':<15} {'出现次数':<10} {'总数':<15}")
                    print("-" * 60)
                    
                    for district_info in districts_list:
                        district_name = district_info['district']
                        total = district_totals.get(district_name, 0)
                        entry_count = len(district_groups[district_name]['entries'])
                        print(f"{district_info['group']:<6} {district_name:<15} {entry_count:<10} {total:<15.2f}")
                    
                    file_total = sum(district_totals.values())
                    print("-" * 60)
                    print(f"{'':6} {'合计':<15} {'':10} {file_total:<15.2f}")
            else:
                print(f"\n文件列数不足4列或无有效数据，跳过处理。")
        else:
            print(f"文件 '{os.path.basename(file_path)}' 分析失败。")
    
    return all_results

def display_final_summary(all_results):
    """
    显示最终汇总结果
    """
    print("\n" + "="*100)
    print("最终汇总报告")
    print("="*100)
    
    # 小学汇总
    primary_results = all_results['小学']
    if primary_results['files']:
        print(f"\n【小学类汇总】 (共处理 {len(primary_results['files'])} 个文件)")
        print("-"*80)
        print(f"{'序号':<4} {'文件名':<40} {'有效行数':<10} {'区县数':<10} {'总计数':<15}")
        print("-"*80)
        
        for i, file_result in enumerate(primary_results['files'], 1):
            district_count = len(file_result['district_totals'])
            file_total = sum(file_result['district_totals'].values())
            print(f"{i:<4} {file_result['file_name']:<40} {file_result['filtered_rows']:<10} {district_count:<10} {file_total:<15.2f}")
        
        print("-"*80)
        print(f"\n小学类区县汇总统计：")
        print("-"*60)
        print(f"{'区县名称':<20} {'总数':<15} {'占比'}")
        print("-"*60)
        
        primary_total = primary_results['overall_total']
        sorted_districts = sorted(primary_results['district_totals'].items(), key=lambda x: x[1], reverse=True)
        
        for district, total in sorted_districts:
            percentage = (total / primary_total * 100) if primary_total > 0 else 0
            print(f"{district:<20} {total:<15.2f} {percentage:.1f}%")
        
        print("-"*60)
        print(f"{'小学总计':<20} {primary_total:<15.2f} 100.0%")
    else:
        print("\n【小学类】: 无数据")
    
    # 中学汇总
    secondary_results = all_results['中学']
    if secondary_results['files']:
        print(f"\n【中学类汇总】 (共处理 {len(secondary_results['files'])} 个文件)")
        print("-"*80)
        print(f"{'序号':<4} {'文件名':<40} {'有效行数':<10} {'区县数':<10} {'总计数':<15}")
        print("-"*80)
        
        for i, file_result in enumerate(secondary_results['files'], 1):
            district_count = len(file_result['district_totals'])
            file_total = sum(file_result['district_totals'].values())
            print(f"{i:<4} {file_result['file_name']:<40} {file_result['filtered_rows']:<10} {district_count:<10} {file_total:<15.2f}")
        
        print("-"*80)
        print(f"\n中学类区县汇总统计：")
        print("-"*60)
        print(f"{'区县名称':<20} {'总数':<15} {'占比'}")
        print("-"*60)
        
        secondary_total = secondary_results['overall_total']
        sorted_districts = sorted(secondary_results['district_totals'].items(), key=lambda x: x[1], reverse=True)
        
        for district, total in sorted_districts:
            percentage = (total / secondary_total * 100) if secondary_total > 0 else 0
            print(f"{district:<20} {total:<15.2f} {percentage:.1f}%")
        
        print("-"*60)
        print(f"{'中学总计':<20} {secondary_total:<15.2f} 100.0%")
    else:
        print("\n【中学类】: 无数据")
    
    # 总体汇总
    total_all = primary_results['overall_total'] + secondary_results['overall_total']
    if total_all > 0:
        print(f"\n{'='*80}")
        print(f"全部总计：{total_all:.2f}")
        print(f"小学类占比：{(primary_results['overall_total']/total_all*100):.1f}%")
        print(f"中学类占比：{(secondary_results['overall_total']/total_all*100):.1f}%")
    
    print("="*100)

def main():
    """
    主函数：处理多个Excel文件并按学校类型分类统计
    """
    print("="*100)
    print("Excel文件信息分析工具 - 按学校类型分类统计")
    print("功能：从第9行开始识别数据，过滤第四列为0或空，识别第三列区县信息")
    print("分类规则：文件名包含'小学'的归为小学类，其余归为中学类")
    print("="*100)
    
    # 处理多个文件
    all_results = process_multiple_files()
    
    # 显示最终汇总
    if all_results['小学']['files'] or all_results['中学']['files']:
        display_final_summary(all_results)
    else:
        print("\n未处理任何文件。")
    
    # 等待用户按下Enter键结束脚本
    input("\n按 Enter 键退出程序...")
    sys.exit(0)

if __name__ == "__main__":
    main()
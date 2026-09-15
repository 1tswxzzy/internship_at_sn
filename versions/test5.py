import os
import sys
import re
import numpy as np
from collections import OrderedDict
import openpyxl

# ==================== 工具函数 ====================
def classify_school_type(file_path):
    """根据文件名判断学校类型（小学/中学）"""
    file_name = os.path.basename(file_path)
    return '小学' if '小学' in file_name else '中学'

def extract_district(text):
    """从单元格文本中提取区县名称"""
    if not text or str(text).strip() == '':
        return None
    text = str(text).strip()
    pattern = r'([^\s,，、]+(?:区|县|市|自治县|自治旗))'
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    if len(text) < 15 and text not in ('合计', '总计', '小计', '备注'):
        return text
    return None

def find_col_by_keywords(ws, head_size, keywords, match_mode='any'):
    """
    在表头区域（1 ~ head_size 行）中查找包含指定关键词的列号。
    优先匹配最底层表头（行号大的），以便找到真正的数据列。
    keywords: 单个字符串 或 字符串列表
    match_mode: 'any' 匹配任一关键词，返回第一个找到的列号
                'all'  需要全部关键词都出现，返回第一个匹配的列号
                'sum'  用于多个关键词，返回所有匹配到的列号列表（用于求和）
    """
    if isinstance(keywords, str):
        keywords = [keywords]

    found_cols = set()
    # 从最后一行表头向上找，优先匹配最接近数据的行
    for row_idx in range(head_size, 0, -1):
        for col_idx in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=row_idx, column=col_idx).value
            if cell_val and isinstance(cell_val, str):
                for kw in keywords:
                    if kw in cell_val:
                        found_cols.add(col_idx)
                        if match_mode == 'any':
                            return col_idx
        # 如果已经找到了所有需要的列，可以提前结束
        if match_mode == 'all' and len(found_cols) == len(keywords):
            break
        if match_mode == 'sum' and len(found_cols) >= len(keywords):
            break

    if match_mode == 'any' and not found_cols:
        return None
    if match_mode == 'all':
        return found_cols.pop() if found_cols else None
    if match_mode == 'sum':
        return list(found_cols)
    return None

# ==================== Excel 基础读取（只读模式，不修改原文件） ====================
def get_worksheet_info(file_path):
    """打开 Excel，返回工作簿、工作表、有效行数、有效列数（只读模式）"""
    try:
        # 使用只读模式避免内存过大，但合并单元格信息不可用，因此仍用普通模式
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
    except FileNotFoundError:
        print(f"错误：找不到文件 '{file_path}'")
        return None, None, None, None
    except Exception as e:
        print(f"读取文件时发生错误：{e}")
        return None, None, None, None

    max_h = ws.max_row
    while max_h > 0:
        if any(cell.value is not None for cell in ws[max_h]):
            break
        max_h -= 1

    max_l = ws.max_column
    while max_l > 0:
        if any(ws.cell(row=r, column=max_l).value is not None for r in range(1, max_h+1)):
            break
        max_l -= 1

    return wb, ws, max_h, max_l

def guess_head_size(ws, max_rows_to_check=20):
    """自动推断表头行数，返回 (head_size, 预览列表)"""
    head_size = 1
    # 合并单元格法
    if ws.merged_cells.ranges:
        head_end = 0
        for mr in ws.merged_cells.ranges:
            if mr.max_row > head_end:
                head_end = mr.max_row
        if head_end > 0:
            head_end = min(head_end, ws.max_row)
            preview = [[cell.value if cell.value is not None else "" for cell in ws[row]]
                       for row in range(1, head_end + 1)]
            return head_end, preview

    # 数字检测法
    data_start = 1
    for row_idx in range(1, min(max_rows_to_check, ws.max_row) + 1):
        for cell in ws[row_idx]:
            if cell.value is None:
                continue
            if isinstance(cell.value, (int, float)):
                data_start = row_idx
                break
            if isinstance(cell.value, str):
                try:
                    float(cell.value.replace(',', '').replace('%', '').strip())
                    data_start = row_idx
                    break
                except ValueError:
                    continue
        if data_start > 1:
            break

    head_size = data_start - 1 if data_start > 1 else 1
    preview = [[cell.value if cell.value is not None else "" for cell in ws[row]]
               for row in range(1, head_size + 1)]
    return head_size, preview

# ==================== 四个计算函数（返回列表） ====================
def calc_media_per_stu(ws, head_size, max_h):
    """返回每百名学生多媒体教室数明细（标题行 + 数据行）"""
    stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    media_col = find_col_by_keywords(ws, head_size, '网络多媒体教室')
    if not stu_col or not media_col:
        print("错误：未找到“在校生数”或“网络多媒体教室”列。")
        return None

    header = ['原始行号', '学校名称', '在校生数', '网络多媒体教室', '每百名学生多媒体教室数']
    rows = [header]
    for r in range(head_size + 1, max_h + 1):
        stu_val = ws.cell(row=r, column=stu_col).value
        media_val = ws.cell(row=r, column=media_col).value
        try:
            stu = float(stu_val) if stu_val is not None else None
            media = float(media_val) if media_val is not None else None
        except (ValueError, TypeError):
            continue
        if stu is None or media is None or stu == 0:
            continue
        result = round(media * 100 / stu, 2)
        school_name = ws.cell(row=r, column=1).value
        rows.append([r, school_name, stu, media, result])
    return rows

def calc_high_education_per_stu(ws, head_size, max_h, school_type):
    """返回每百名学生高学历教师数明细"""
    stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    if not stu_col:
        print("错误：未找到“在校生数”列。")
        return None

    kw_primary = ['硕士研究生', '本科', '专科']
    kw_middle = ['硕士研究生', '本科']
    keywords = kw_primary if school_type == '小学' else kw_middle
    target_cols = []
    for kw in keywords:
        col = find_col_by_keywords(ws, head_size, kw)
        if col:
            target_cols.append(col)
        else:
            print(f"警告：未找到“{kw}”列。")
    if not target_cols:
        print("错误：未找到任何学历教师列。")
        return None

    header = ['原始行号', '学校名称', '在校生数', '学历教师总数', '每百名学生高学历教师数']
    rows = [header]
    for r in range(head_size + 1, max_h + 1):
        stu_val = ws.cell(row=r, column=stu_col).value
        edu_sum = 0
        for c in target_cols:
            v = ws.cell(row=r, column=c).value
            try:
                edu_sum += float(v) if v is not None else 0
            except (ValueError, TypeError):
                pass
        try:
            stu = float(stu_val) if stu_val is not None else None
        except (ValueError, TypeError):
            continue
        if stu is None or stu == 0 or edu_sum == 0:
            continue
        result = round(edu_sum * 100 / stu, 2)
        school_name = ws.cell(row=r, column=1).value
        rows.append([r, school_name, stu, edu_sum, result])
    return rows

def calc_backbone_per_stu(ws, head_size, max_h):
    """返回每百名学生骨干教师数明细"""
    stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    backbone_col = find_col_by_keywords(ws, head_size, '骨干教师')
    if not stu_col or not backbone_col:
        print("错误：未找到“在校生数”或“骨干教师”列。")
        return None

    header = ['原始行号', '学校名称', '在校生数', '骨干教师数', '每百名学生骨干教师数']
    rows = [header]
    for r in range(head_size + 1, max_h + 1):
        stu_val = ws.cell(row=r, column=stu_col).value
        back_val = ws.cell(row=r, column=backbone_col).value
        try:
            stu = float(stu_val) if stu_val is not None else None
            back = float(back_val) if back_val is not None else None
        except (ValueError, TypeError):
            continue
        if stu is None or back is None or stu == 0:
            continue
        result = round(back * 100 / stu, 2)
        school_name = ws.cell(row=r, column=1).value
        rows.append([r, school_name, stu, back, result])
    return rows

def calc_art_sport_per_stu(ws, head_size, max_h):
    """返回每百名学生艺体教师数明细"""
    stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    sport_col = find_col_by_keywords(ws, head_size, '体育')
    art_col = find_col_by_keywords(ws, head_size, '艺术')
    if not stu_col or (not sport_col and not art_col):
        print("错误：未找到“在校生数”或体育/艺术相关列。")
        return None

    header = ['原始行号', '学校名称', '在校生数', '体育教师', '艺术教师', '每百名学生艺体教师数']
    rows = [header]
    for r in range(head_size + 1, max_h + 1):
        stu_val = ws.cell(row=r, column=stu_col).value
        sport_val = ws.cell(row=r, column=sport_col).value if sport_col else 0
        art_val = ws.cell(row=r, column=art_col).value if art_col else 0
        try:
            stu = float(stu_val) if stu_val is not None else None
            sport = float(sport_val) if sport_val is not None else 0
            art = float(art_val) if art_val is not None else 0
        except (ValueError, TypeError):
            continue
        if stu is None or stu == 0:
            continue
        total = sport + art
        if total == 0:
            continue
        result = round(total * 100 / stu, 2)
        school_name = ws.cell(row=r, column=1).value
        rows.append([r, school_name, stu, sport, art, result])
    return rows

# ==================== 区县统计函数（返回结构化数据） ====================
def analyze_district_data(ws, head_size, max_h, school_type):
    """
    从工作表提取原始数据，按区县分组，并返回：
        district_data: 有序字典 {区县: [{'A':..., 'B':..., 'C':..., 'D':...}, ...]}
        stats_rows:     差异系数表格（二维列表，含标题行）
        summary_rows:   区县汇总表格（二维列表，含标题行）
    """
    # 定位列
    stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    if school_type == '小学':
        edu_cols = [find_col_by_keywords(ws, head_size, kw) for kw in ['硕士研究生', '本科', '专科']]
    else:
        edu_cols = [find_col_by_keywords(ws, head_size, kw) for kw in ['硕士研究生', '本科']]
    edu_cols = [c for c in edu_cols if c]
    backbone_col = find_col_by_keywords(ws, head_size, '骨干教师')
    sport_col = find_col_by_keywords(ws, head_size, '体育')
    art_col = find_col_by_keywords(ws, head_size, '艺术')

    if not stu_col:
        print("错误：未找到“在校生数”列，无法进行区县统计。")
        return None, None, None

    # 读取数据，按区县分组
    district_data = OrderedDict()
    for r in range(head_size + 1, max_h + 1):
        district_cell = ws.cell(row=r, column=3).value
        district = extract_district(district_cell) if district_cell else None
        if not district:
            continue

        stu_val = ws.cell(row=r, column=stu_col).value
        try:
            stu = float(stu_val)
        except (ValueError, TypeError):
            continue
        if stu == 0:
            continue

        edu_sum = 0
        for c in edu_cols:
            v = ws.cell(row=r, column=c).value
            try:
                edu_sum += float(v) if v is not None else 0
            except:
                pass

        back_val = ws.cell(row=r, column=backbone_col).value if backbone_col else 0
        try:
            back = float(back_val) if back_val is not None else 0
        except:
            back = 0

        sport_val = ws.cell(row=r, column=sport_col).value if sport_col else 0
        art_val = ws.cell(row=r, column=art_col).value if art_col else 0
        try:
            sport = float(sport_val) if sport_val is not None else 0
            art = float(art_val) if art_val is not None else 0
        except:
            sport = art = 0

        if district not in district_data:
            district_data[district] = []
        district_data[district].append({
            'A': stu,
            'B': edu_sum,
            'C': back,
            'D': sport + art
        })

    if not district_data:
        return None, None, None

    # 计算差异系数表格
    def calc_stats(data_list, indicator_key):
        A_sum = sum(d['A'] for d in data_list)
        if A_sum == 0:
            return {'X': 0, 'S': 0, 'CV': 0}
        indicator_sum = sum(d[indicator_key] for d in data_list)
        X = (indicator_sum * 100) / A_sum
        variance = 0
        for d in data_list:
            if d['A'] > 0:
                ratio = (d[indicator_key] / d['A']) * 100
                variance += (d['A'] / A_sum) * (ratio - X)**2
        S = np.sqrt(variance) if variance > 0 else 0
        CV = round((S / X), 2) if X != 0 else 0
        return {'X': X, 'S': S, 'CV': CV, 'A_sum': A_sum, 'indicator_sum': indicator_sum}

    # 分配组别
    group_counter = 1
    district_groups = OrderedDict()
    for dist in district_data:
        district_groups[dist] = {
            'group': f"T{group_counter}",
            'data': district_data[dist]
        }
        group_counter += 1

    # 构建差异系数表格（用于写入sheet）
    stats_header = ['组别', '区县',
                    'B_X', 'B_S', 'B_CV',
                    'C_X', 'C_S', 'C_CV',
                    'D_X', 'D_S', 'D_CV']
    stats_rows = [stats_header]
    all_B_raw = []
    all_C_raw = []
    all_D_raw = []
    for dist, info in district_groups.items():
        data = info['data']
        sB = calc_stats(data, 'B')
        sC = calc_stats(data, 'C')
        sD = calc_stats(data, 'D')
        stats_rows.append([
            info['group'], dist,
            round(sB['X'], 4), round(sB['S'], 6), sB['CV'],
            round(sC['X'], 4), round(sC['S'], 6), sC['CV'],
            round(sD['X'], 4), round(sD['S'], 6), sD['CV']
        ])
        all_B_raw.extend(data)
        all_C_raw.extend(data)
        all_D_raw.extend(data)

    # 总计行
    totalB = calc_stats(all_B_raw, 'B')
    totalC = calc_stats(all_C_raw, 'C')
    totalD = calc_stats(all_D_raw, 'D')
    stats_rows.append([
        '', '总计',
        round(totalB['X'], 4), round(totalB['S'], 6), totalB['CV'],
        round(totalC['X'], 4), round(totalC['S'], 6), totalC['CV'],
        round(totalD['X'], 4), round(totalD['S'], 6), totalD['CV']
    ])

    # 构建区县汇总表格（各区县的合计）
    summary_header = ['区县', '学生总数(A)', '高学历教师(B)', '骨干教师(C)', '艺体教师(D)']
    summary_rows = [summary_header]
    total_A_all = 0
    total_B_all = 0
    total_C_all = 0
    total_D_all = 0
    for dist, info in district_groups.items():
        data = info['data']
        sumA = sum(d['A'] for d in data)
        sumB = sum(d['B'] for d in data)
        sumC = sum(d['C'] for d in data)
        sumD = sum(d['D'] for d in data)
        summary_rows.append([dist, sumA, sumB, sumC, sumD])
        total_A_all += sumA
        total_B_all += sumB
        total_C_all += sumC
        total_D_all += sumD
    summary_rows.append(['总计', total_A_all, total_B_all, total_C_all, total_D_all])

    return district_data, stats_rows, summary_rows

# ==================== 输出结果到新工作簿 ====================
def save_results_to_new_file(original_path, school_type, calc_results, stats_rows, summary_rows):
    """将各项结果写入一个新的 Excel 文件，保存于原文件同目录"""
    # 构建新文件名
    dir_name = os.path.dirname(original_path)
    base_name = os.path.splitext(os.path.basename(original_path))[0]
    new_name = f"{base_name}_分析结果.xlsx"
    new_path = os.path.join(dir_name, new_name)

    out_wb = openpyxl.Workbook()
    # 删除默认的空sheet
    out_wb.remove(out_wb.active)

    # 按顺序添加sheet
    sheet_names = [
        '每百名学生多媒体教室数',
        '每百名学生高学历教师数',
        '每百名学生骨干教师数',
        '每百名学生艺体教师数',
        '区县差异系数统计',
        '区县汇总'
    ]
    data_blocks = [
        calc_results.get('media'),
        calc_results.get('high_edu'),
        calc_results.get('backbone'),
        calc_results.get('art_sport'),
        stats_rows,
        summary_rows
    ]

    for sname, block in zip(sheet_names, data_blocks):
        ws = out_wb.create_sheet(title=sname)
        if block:
            for row in block:
                ws.append(row)
        else:
            ws.append(['无数据'])

    out_wb.save(new_path)
    print(f"\n分析结果已保存至：{new_path}")
    return new_path

# ==================== 主程序 ====================
def main():
    print("=" * 60)
    print("Excel 教育数据综合分析工具（生成独立报告文件）")
    print("=" * 60)

    while True:
        file_path = input("\n请输入 Excel 文件路径（或拖拽文件，输入 q 退出）：").strip()
        if file_path.lower() == 'q':
            break
        file_path = file_path.strip('"').strip("'")
        if not os.path.exists(file_path):
            print("文件不存在。")
            continue

        school_type = classify_school_type(file_path)
        # 读取文件（只读模式）
        wb, ws, max_h, max_l = get_worksheet_info(file_path)
        if wb is None:
            continue
        print(f"\n当前文件：{os.path.basename(file_path)}  类型：{school_type}")
        print(f"有效行数：{max_h}，有效列数：{max_l}")

        # 表头检测
        head_size, preview = guess_head_size(ws)
        print(f"自动检测表头行数：{head_size}")
        print("表头预览：")
        for i, line in enumerate(preview, 1):
            print(f"  第{i}行: {line}")
        user_input = input("确认表头行数（回车确认 / 输入正确行数）：").strip()
        if user_input.isdigit():
            head_size = int(user_input)

        max_size = max_h - head_size
        if max_size <= 0:
            print("无有效数据行。")
            continue

        # 存储用户选择执行的计算结果
        calc_results = {
            'media': None,
            'high_edu': None,
            'backbone': None,
            'art_sport': None
        }

        # 菜单循环
        while True:
            print("\n" + "-" * 40)
            print("请选择计算项目：")
            print("  1. 每百名学生多媒体教室数")
            print("  2. 每百名学生高学历教师数")
            print("  3. 每百名学生骨干教师数")
            print("  4. 每百名学生艺体教师数")
            print("  5. 一键执行以上全部 + 区县差异系数与汇总")
            print("  0. 完成选择，生成报告文件")
            print("-" * 40)

            choice = input("请输入选项：").strip()
            if choice == '0':
                break
            elif choice == '1':
                if calc_results['media'] is None:
                    calc_results['media'] = calc_media_per_stu(ws, head_size, max_h)
                else:
                    print("该计算已完成，无需重复。")
            elif choice == '2':
                if calc_results['high_edu'] is None:
                    calc_results['high_edu'] = calc_high_education_per_stu(ws, head_size, max_h, school_type)
                else:
                    print("该计算已完成。")
            elif choice == '3':
                if calc_results['backbone'] is None:
                    calc_results['backbone'] = calc_backbone_per_stu(ws, head_size, max_h)
                else:
                    print("该计算已完成。")
            elif choice == '4':
                if calc_results['art_sport'] is None:
                    calc_results['art_sport'] = calc_art_sport_per_stu(ws, head_size, max_h)
                else:
                    print("该计算已完成。")
            elif choice == '5':
                # 一键执行全部（若未执行则执行）
                if calc_results['media'] is None:
                    calc_results['media'] = calc_media_per_stu(ws, head_size, max_h)
                if calc_results['high_edu'] is None:
                    calc_results['high_edu'] = calc_high_education_per_stu(ws, head_size, max_h, school_type)
                if calc_results['backbone'] is None:
                    calc_results['backbone'] = calc_backbone_per_stu(ws, head_size, max_h)
                if calc_results['art_sport'] is None:
                    calc_results['art_sport'] = calc_art_sport_per_stu(ws, head_size, max_h)
                print("全部四项计算已完成。")
            else:
                print("无效输入。")
                continue

            more = input("\n继续其他计算？(y/n)：").strip().lower()
            if more != 'y':
                break

        # 生成区县统计（如果需要）
        do_district = input("\n是否进行区县差异系数与汇总统计？(y/n)：").strip().lower()
        stats_rows = None
        summary_rows = None
        if do_district == 'y':
            _, stats_rows, summary_rows = analyze_district_data(ws, head_size, max_h, school_type)
            if stats_rows is None:
                print("区县统计失败，可能缺少必要的列。")
        else:
            # 如果选了5但用户在这里选n，也应提醒
            print("跳过区县统计。")

        # 生成报告文件
        if any(calc_results.values()) or stats_rows or summary_rows:
            save_results_to_new_file(file_path, school_type, calc_results, stats_rows, summary_rows)
        else:
            print("没有选择任何计算，不生成报告。")

        cont = input("\n是否处理其他文件？(y/n)：").strip().lower()
        if cont != 'y':
            break

    input("\n按 Enter 键退出...")
    sys.exit(0)

if __name__ == "__main__":
    main()
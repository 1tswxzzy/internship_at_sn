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

# ==================== Excel 基础读取 ====================
def get_worksheet_info(file_path):
    """打开 Excel，返回工作簿、工作表、有效行数、有效列数"""
    try:
        wb = openpyxl.load_workbook(file_path)
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

# ==================== 四个计算函数 ====================
def calc_media_per_stu(wb, ws, head_size, max_h):
    """1. 每百名学生多媒体教室数"""
    stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    media_col = find_col_by_keywords(ws, head_size, '网络多媒体教室')
    if not stu_col or not media_col:
        print("错误：未找到“在校生数”或“网络多媒体教室”列，计算终止。")
        return

    sheet_name = '每百名学生多媒体教室数'
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    new_ws = wb.create_sheet(sheet_name)
    new_ws.append(['原始行号', '学校名称', '在校生数', '网络多媒体教室', '每百名学生多媒体教室数'])

    count = 0
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
        new_ws.append([r, school_name, stu, media, result])
        count += 1
    print(f"计算完成，共处理 {count} 行，结果写入 sheet“{sheet_name}”。")

def calc_high_education_per_stu(wb, ws, head_size, max_h, school_type):
    """2. 每百名学生拥有高于规定学历教师数"""
    stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    if not stu_col:
        print("错误：未找到“在校生数”列。")
        return

    # 根据中小学选择不同学历列：小学 = 硕士+本科+专科，中学 = 硕士+本科
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
        return

    sheet_name = '每百名学生高学历教师数'
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    new_ws = wb.create_sheet(sheet_name)
    new_ws.append(['原始行号', '学校名称', '在校生数', '学历教师总数', '每百名学生高学历教师数'])

    count = 0
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
        new_ws.append([r, school_name, stu, edu_sum, result])
        count += 1
    print(f"计算完成，共处理 {count} 行，结果写入 sheet“{sheet_name}”。")

def calc_backbone_per_stu(wb, ws, head_size, max_h):
    """3. 每百名学生拥有骨干教师数"""
    stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    backbone_col = find_col_by_keywords(ws, head_size, '骨干教师')
    if not stu_col or not backbone_col:
        print("错误：未找到“在校生数”或“骨干教师”列。")
        return

    sheet_name = '每百名学生骨干教师数'
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    new_ws = wb.create_sheet(sheet_name)
    new_ws.append(['原始行号', '学校名称', '在校生数', '骨干教师数', '每百名学生骨干教师数'])

    count = 0
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
        new_ws.append([r, school_name, stu, back, result])
        count += 1
    print(f"计算完成，共处理 {count} 行，结果写入 sheet“{sheet_name}”。")

def calc_art_sport_per_stu(wb, ws, head_size, max_h):
    """4. 每百名学生拥有艺体专任教师数（体育+艺术）"""
    stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    sport_col = find_col_by_keywords(ws, head_size, '体育')
    art_col = find_col_by_keywords(ws, head_size, '艺术')
    if not stu_col or (not sport_col and not art_col):
        print("错误：未找到“在校生数”或体育/艺术相关列。")
        return

    sheet_name = '每百名学生艺体教师数'
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    new_ws = wb.create_sheet(sheet_name)
    new_ws.append(['原始行号', '学校名称', '在校生数', '体育教师', '艺术教师', '每百名学生艺体教师数'])

    count = 0
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
        new_ws.append([r, school_name, stu, sport, art, result])
        count += 1
    print(f"计算完成，共处理 {count} 行，结果写入 sheet“{sheet_name}”。")

# ==================== 区县统计与差异系数 ====================
def compute_district_statistics(ws, head_size, max_h, school_type):
    """
    执行区县识别、指标计算（B、C、D），返回差异系数统计结果。
    此处为了与第一份代码兼容，我们基于 openpyxl 实现相同的逻辑。
    """
    # 1. 找到所需列
    stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    if school_type == '小学':
        edu_cols = [find_col_by_keywords(ws, head_size, kw) for kw in ['硕士研究生', '本科', '专科']]
    else:
        edu_cols = [find_col_by_keywords(ws, head_size, kw) for kw in ['硕士研究生', '本科']]
    edu_cols = [c for c in edu_cols if c]  # 去除None
    backbone_col = find_col_by_keywords(ws, head_size, '骨干教师')
    sport_col = find_col_by_keywords(ws, head_size, '体育')
    art_col = find_col_by_keywords(ws, head_size, '艺术')

    # 2. 读取所有数据行，按区县分组
    district_data = OrderedDict()  # 区县 -> list of dict
    for r in range(head_size + 1, max_h + 1):
        # 区县名称（第三列）
        district_cell = ws.cell(row=r, column=3).value
        district = extract_district(district_cell) if district_cell else None
        if not district:
            continue

        # 学生数
        stu_val = ws.cell(row=r, column=stu_col).value
        try:
            stu = float(stu_val)
        except (ValueError, TypeError):
            continue
        if stu == 0:
            continue

        # B指标（学历教师）
        edu_sum = 0
        for c in edu_cols:
            v = ws.cell(row=r, column=c).value
            try:
                edu_sum += float(v) if v is not None else 0
            except:
                pass
        # C指标（骨干教师）
        back_val = ws.cell(row=r, column=backbone_col).value if backbone_col else 0
        try:
            back = float(back_val) if back_val is not None else 0
        except:
            back = 0
        # D指标（体育+艺术）
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
        print("未提取到任何区县数据，无法统计。")
        return

    # 3. 计算每个区县的 X, S, CV（对每个指标）
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

    # 打印统计表
    print("\n" + "="*130)
    print(f"{school_type}类区县差异系数统计")
    print("="*130)
    print(f"{'组别':<6} {'区县':<12} {'B(高学历教师)':<35} {'C(骨干教师)':<35} {'D(艺体教师)':<35}")
    print(f"{'':6} {'':12} {'X':<12} {'S':<12} {'CV':<10} {'X':<12} {'S':<12} {'CV':<10} {'X':<12} {'S':<12} {'CV':<10}")
    print("-"*130)

    all_B_raw = []
    all_C_raw = []
    all_D_raw = []
    for dist, info in district_groups.items():
        data = info['data']
        stats_B = calc_stats(data, 'B')
        stats_C = calc_stats(data, 'C')
        stats_D = calc_stats(data, 'D')
        print(f"{info['group']:<6} {dist:<12} "
              f"{stats_B['X']:<12.4f} {stats_B['S']:<12.6f} {stats_B['CV']:<10.2f} "
              f"{stats_C['X']:<12.4f} {stats_C['S']:<12.6f} {stats_C['CV']:<10.2f} "
              f"{stats_D['X']:<12.4f} {stats_D['S']:<12.6f} {stats_D['CV']:<10.2f}")
        # 收集原始数据用于总计
        all_B_raw.extend(data)
        all_C_raw.extend(data)
        all_D_raw.extend(data)

    # 总计
    total_B = calc_stats(all_B_raw, 'B')
    total_C = calc_stats(all_C_raw, 'C')
    total_D = calc_stats(all_D_raw, 'D')
    print("-"*130)
    print(f"{'':6} {'总计':<12} "
          f"{total_B['X']:<12.4f} {total_B['S']:<12.6f} {total_B['CV']:<10.2f} "
          f"{total_C['X']:<12.4f} {total_C['S']:<12.6f} {total_C['CV']:<10.2f} "
          f"{total_D['X']:<12.4f} {total_D['S']:<12.6f} {total_D['CV']:<10.2f}")
    print("="*130)

# ==================== 主程序 ====================
def main():
    print("=" * 60)
    print("Excel 教育数据综合分析工具（自动表头识别）")
    print("=" * 60)

    while True:
        file_path = input("\n请输入 Excel 文件路径（或拖拽文件，输入 q 退出）：").strip()
        if file_path.lower() == 'q':
            break
        file_path = file_path.strip('"').strip("'")
        if not os.path.exists(file_path):
            print("文件不存在。")
            continue

        # 学校类型
        school_type = classify_school_type(file_path)

        # 读取文件
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

        # 定义计算函数字典
        calc_functions = {
            1: lambda: calc_media_per_stu(wb, ws, head_size, max_h),
            2: lambda: calc_high_education_per_stu(wb, ws, head_size, max_h, school_type),
            3: lambda: calc_backbone_per_stu(wb, ws, head_size, max_h),
            4: lambda: calc_art_sport_per_stu(wb, ws, head_size, max_h),
        }

        while True:
            print("\n" + "-" * 40)
            print("请选择计算项目：")
            print("  1. 每百名学生多媒体教室数")
            print("  2. 每百名学生高学历教师数")
            print("  3. 每百名学生骨干教师数")
            print("  4. 每百名学生艺体教师数")
            print("  5. 一键执行以上全部 + 区县差异系数统计")
            print("  0. 完成选择，保存/退出")
            print("-" * 40)

            choice = input("请输入选项：").strip()
            if choice == '0':
                break
            elif choice == '5':
                for func in calc_functions.values():
                    func()
                compute_district_statistics(ws, head_size, max_h, school_type)
            elif choice in ('1','2','3','4'):
                calc_functions[int(choice)]()
            else:
                print("无效输入。")
                continue

            more = input("\n继续其他计算？(y/n)：").strip().lower()
            if more != 'y':
                break

        # 保存
        save_choice = input("\n是否保存文件？(y=覆盖原文件 / s=另存为 / n=不保存)：").strip().lower()
        if save_choice == 'y':
            try:
                wb.save(file_path)
                print("已保存至原文件。")
            except PermissionError:
                print("文件被占用，请关闭 Excel 后重试。")
        elif save_choice == 's':
            new_path = input("输入新文件完整路径：").strip().strip('"')
            if new_path:
                try:
                    wb.save(new_path)
                    print(f"已保存至 {new_path}")
                except Exception as e:
                    print(f"保存失败：{e}")
        else:
            print("未保存。")

        cont = input("\n是否处理其他文件？(y/n)：").strip().lower()
        if cont != 'y':
            break

    input("\n按 Enter 键退出...")
    sys.exit(0)

if __name__ == "__main__":
    main()
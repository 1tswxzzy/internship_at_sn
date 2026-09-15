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
    if not text or str(text).strip() == '':
        return None
    text = str(text).strip()
    # 只匹配以“区、县、市、自治县、自治旗”结尾的词
    pattern = r'([^\s,，、]+(?:区|县|市|自治县|自治旗))'
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    # 去掉兜底逻辑，避免将非行政区划误判为区县
    return None

def find_col_by_keywords(ws, head_size, keywords, match_mode='any'):
    """（降级方案）通过关键词匹配列，优先底层表头"""
    if isinstance(keywords, str):
        keywords = [keywords]
    found_cols = set()
    for row_idx in range(head_size, 0, -1):
        for col_idx in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=row_idx, column=col_idx).value
            if cell_val and isinstance(cell_val, str):
                for kw in keywords:
                    if kw in cell_val:
                        if match_mode == 'any':
                            return col_idx
                        found_cols.add(col_idx)
        if match_mode == 'sum' and len(found_cols) >= len(keywords):
            return list(found_cols)
        if match_mode == 'all' and len(found_cols) == len(keywords):
            break
    if match_mode == 'any' and not found_cols:
        return None
    if match_mode == 'all':
        return found_cols.pop() if found_cols else None
    if match_mode == 'sum':
        return list(found_cols)
    return None

def find_col_by_hierarchy(ws, head_size, hierarchy):
    """
    根据表头层级路径查找唯一列。
    hierarchy: ['关键词1', '关键词2', ...] ，这些关键词必须按行序从上到下出现在同一列中。
    返回：列号(int) 或 None
    """
    for col in range(1, ws.max_column + 1):
        idx = 0
        for row in range(1, head_size + 1):
            cell_val = ws.cell(row=row, column=col).value
            if cell_val and isinstance(cell_val, str) and hierarchy[idx] in cell_val:
                idx += 1
                if idx == len(hierarchy):
                    return col
    return None

def find_cols_with_prefix(ws, head_size, prefix_hierarchy):
    """
    返回所有满足前缀层级序列的列号（用于定位某一区域下的多列）。
    原理：只要某列按顺序出现了 prefix_hierarchy 中的所有关键词，就认为是该区域的列。
    """
    cols = []
    for col in range(1, ws.max_column + 1):
        idx = 0
        for row in range(1, head_size + 1):
            cell_val = ws.cell(row=row, column=col).value
            if cell_val and isinstance(cell_val, str) and prefix_hierarchy[idx] in cell_val:
                idx += 1
                if idx == len(prefix_hierarchy):
                    cols.append(col)
                    break
    return cols

# ==================== Excel 基础读取 ====================
def get_worksheet_info(file_path):
    """打开 Excel，返回工作簿、工作表、有效行数、有效列数"""
    try:
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

# ==================== 通用辅助：查找“是否计校数”列 ====================
def find_count_school_col(ws, head_size):
    """返回“是否计校数”列的列号，未找到返回None"""
    return find_col_by_keywords(ws, head_size, '是否计校数')

def build_school_merge_map(ws, head_size):
    """
    扫描数据表，找出所有“是否计校数”为0的学校，让用户一次性指定它们对应的原校点。
    返回：{不计校数学校名称: 用户输入的原校名称} 字典
    """
    count_col = find_count_school_col(ws, head_size)
    if not count_col:
        print("未找到“是否计校数”列，跳过不计校数合并。")
        return {}

    # 收集所有不计校数的学校名（去重）
    pending_schools = set()
    for r in range(head_size + 1, ws.max_row + 1):
        try:
            if float(ws.cell(row=r, column=count_col).value) == 0:
                name = ws.cell(row=r, column=1).value
                if name:
                    pending_schools.add(str(name).strip())
        except (ValueError, TypeError):
            pass

    if not pending_schools:
        print("没有发现不计校数的学校。")
        return {}

    print(f"\n发现 {len(pending_schools)} 个不计校数的学校，请依次输入它们所属的原校点名称：")
    merge_map = {}
    for school in pending_schools:
        while True:
            master = input(f"  {school}  → 原校点名称：").strip()
            if master:
                merge_map[school] = master
                break
            print("  输入不能为空，请重新输入。")
    print("映射关系已记录。")
    return merge_map

# ==================== 修改后的四个计算函数（增加 merge_map 参数） ====================
def calc_media_per_stu(ws, head_size, max_h, merge_map=None):
    """1. 每百名学生多媒体教室数（含不计校数合并）"""
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    media_col = find_col_by_hierarchy(ws, head_size, ['网络多媒体教室'])
    if not media_col:
        media_col = find_col_by_keywords(ws, head_size, '网络多媒体教室')
    count_col = find_count_school_col(ws, head_size)
    if not stu_col or not media_col:
        print("错误：未找到“在校生数(合计)”或“网络多媒体教室”列。")
        return None

    header = ['原始行号', '学校名称', '在校生数(合计)', '网络多媒体教室', '每百名学生多媒体教室数']
    rows = [header]
    school_index = {}   # 记录学校名称 -> rows中的索引（用于后续更新）
    pending_items = []  # 存储不计校数的行数据

    # 第一遍：处理计校数行
    for r in range(head_size + 1, max_h + 1):
        if count_col is not None:
            count_val = ws.cell(row=r, column=count_col).value
            try:
                if float(count_val) == 0:
                    school_name = ws.cell(row=r, column=1).value
                    if not school_name:
                        continue
                    stu_val = ws.cell(row=r, column=stu_col).value
                    media_val = ws.cell(row=r, column=media_col).value
                    pending_items.append({
                        'row': r,
                        'school': str(school_name).strip(),
                        'stu': stu_val,
                        'media': media_val
                    })
                    continue
            except (ValueError, TypeError):
                pass

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
        row_data = [r, school_name, stu, media, result]
        rows.append(row_data)
        school_index[str(school_name).strip()] = len(rows) - 1

    # 第二遍：处理不计校数的行，合并到原校
    for item in pending_items:
        school = item['school']
        # 优先使用映射字典
        if merge_map and school in merge_map:
            master_name = merge_map[school]
        else:
            print(f"\n发现不计校数的学校：{school} (行{item['row']})")
            while True:
                master_name = input(f"请输入该校所属的原校点名称（严格匹配学校名称）：").strip()
                if not master_name:
                    continue
                if master_name in school_index:
                    break
                print(f"未找到学校“{master_name}”，请重新输入。")

        if master_name in school_index:
            idx = school_index[master_name]
            old_stu = rows[idx][2]
            old_media = rows[idx][3]
            try:
                add_stu = float(item['stu']) if item['stu'] is not None else 0
                add_media = float(item['media']) if item['media'] is not None else 0
            except (ValueError, TypeError):
                add_stu = add_media = 0
            new_stu = old_stu + add_stu
            new_media = old_media + add_media
            new_result = round(new_media * 100 / new_stu, 2) if new_stu != 0 else 0
            rows[idx][2] = new_stu
            rows[idx][3] = new_media
            rows[idx][4] = new_result
            print(f"已将 {school} 合并到 {master_name}，更新：在校生数={new_stu}, 多媒体教室={new_media}, 每百名={new_result}")
        else:
            print(f"错误：原校点 {master_name} 未找到，跳过合并 {school}。")
    return rows


def calc_high_education_per_stu(ws, head_size, max_h, school_type, merge_map=None):
    """2. 每百名学生高学历教师数（含不计校数合并）"""
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    if not stu_col:
        print("错误：未找到“在校生数”列。")
        return None

    # 学历区域
    edu_region = find_cols_with_prefix(ws, head_size, ['专任教师', '#按学历分'])
    if not edu_region:
        kw_primary = ['硕士研究生', '本科', '专科']
        kw_middle = ['硕士研究生', '本科']
        keywords = kw_primary if school_type == '小学' else kw_middle
        target_cols = []
        for kw in keywords:
            col = find_col_by_keywords(ws, head_size, kw)
            if col:
                target_cols.append(col)
    else:
        kw_primary = ['硕士研究生', '本科', '专科']
        kw_middle = ['硕士研究生', '本科']
        keywords = kw_primary if school_type == '小学' else kw_middle
        target_cols = []
        for kw in keywords:
            found = None
            for col in edu_region:
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str) and kw in val:
                        found = col
                        break
                if found:
                    target_cols.append(found)
                    break
    if not target_cols:
        print("错误：未找到任何学历教师列。")
        return None

    count_col = find_count_school_col(ws, head_size)
    header = ['原始行号', '学校名称', '在校生数(合计)', '学历教师总数', '每百名学生高学历教师数']
    rows = [header]
    school_index = {}
    pending_items = []

    for r in range(head_size + 1, max_h + 1):
        if count_col is not None:
            count_val = ws.cell(row=r, column=count_col).value
            try:
                if float(count_val) == 0:
                    school_name = ws.cell(row=r, column=1).value
                    if not school_name:
                        continue
                    stu_val = ws.cell(row=r, column=stu_col).value
                    edu_vals = {c: ws.cell(row=r, column=c).value for c in target_cols}
                    pending_items.append({
                        'row': r,
                        'school': str(school_name).strip(),
                        'stu': stu_val,
                        'edu_vals': edu_vals
                    })
                    continue
            except (ValueError, TypeError):
                pass

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
        row_data = [r, school_name, stu, edu_sum, result]
        rows.append(row_data)
        school_index[str(school_name).strip()] = len(rows) - 1

    for item in pending_items:
        school = item['school']
        if merge_map and school in merge_map:
            master_name = merge_map[school]
        else:
            print(f"\n发现不计校数的学校：{school} (行{item['row']})")
            while True:
                master_name = input("请输入该校所属的原校点名称：").strip()
                if not master_name:
                    continue
                if master_name in school_index:
                    break
                print(f"未找到学校“{master_name}”。")

        if master_name in school_index:
            idx = school_index[master_name]
            old_stu = rows[idx][2]
            old_edu = rows[idx][3]
            try:
                add_stu = float(item['stu']) if item['stu'] is not None else 0
                add_edu = sum(float(item['edu_vals'][c]) if item['edu_vals'][c] is not None else 0 for c in target_cols)
            except (ValueError, TypeError):
                add_stu = add_edu = 0
            new_stu = old_stu + add_stu
            new_edu = old_edu + add_edu
            new_result = round(new_edu * 100 / new_stu, 2) if new_stu != 0 else 0
            rows[idx][2] = new_stu
            rows[idx][3] = new_edu
            rows[idx][4] = new_result
            print(f"已将 {school} 合并到 {master_name}，更新：学生={new_stu}, 学历教师={new_edu}, 每百名={new_result}")
        else:
            print(f"错误：原校点 {master_name} 未找到，跳过合并 {school}。")
    return rows


def calc_backbone_per_stu(ws, head_size, max_h, school_type, merge_map=None):
    """3. 每百名学生骨干教师数（含不计校数合并）"""
    if school_type == '小学':
        backbone_col = find_col_by_hierarchy(ws, head_size, ['教基1102', '县级以上骨干教师(小学)'])
    else:
        backbone_col = find_col_by_hierarchy(ws, head_size, ['教基1102', '县级以上骨干教师(初中)'])
    if not backbone_col:
        backbone_col = find_col_by_keywords(ws, head_size, '骨干教师')
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    if not stu_col or not backbone_col:
        print("错误：未找到“在校生数”或“骨干教师”列。")
        return None

    count_col = find_count_school_col(ws, head_size)
    header = ['原始行号', '学校名称', '在校生数(合计)', '骨干教师数', '每百名学生骨干教师数']
    rows = [header]
    school_index = {}
    pending_items = []

    for r in range(head_size + 1, max_h + 1):
        if count_col is not None:
            count_val = ws.cell(row=r, column=count_col).value
            try:
                if float(count_val) == 0:
                    school_name = ws.cell(row=r, column=1).value
                    if not school_name:
                        continue
                    stu_val = ws.cell(row=r, column=stu_col).value
                    back_val = ws.cell(row=r, column=backbone_col).value
                    pending_items.append({
                        'row': r,
                        'school': str(school_name).strip(),
                        'stu': stu_val,
                        'back': back_val
                    })
                    continue
            except (ValueError, TypeError):
                pass

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
        row_data = [r, school_name, stu, back, result]
        rows.append(row_data)
        school_index[str(school_name).strip()] = len(rows) - 1

    for item in pending_items:
        school = item['school']
        if merge_map and school in merge_map:
            master_name = merge_map[school]
        else:
            print(f"\n发现不计校数的学校：{school} (行{item['row']})")
            while True:
                master_name = input("请输入该校所属的原校点名称：").strip()
                if not master_name:
                    continue
                if master_name in school_index:
                    break
                print(f"未找到学校“{master_name}”。")

        if master_name in school_index:
            idx = school_index[master_name]
            old_stu = rows[idx][2]
            old_back = rows[idx][3]
            try:
                add_stu = float(item['stu']) if item['stu'] is not None else 0
                add_back = float(item['back']) if item['back'] is not None else 0
            except (ValueError, TypeError):
                add_stu = add_back = 0
            new_stu = old_stu + add_stu
            new_back = old_back + add_back
            new_result = round(new_back * 100 / new_stu, 2) if new_stu != 0 else 0
            rows[idx][2] = new_stu
            rows[idx][3] = new_back
            rows[idx][4] = new_result
            print(f"已将 {school} 合并到 {master_name}，更新：学生={new_stu}, 骨干教师={new_back}, 每百名={new_result}")
        else:
            print(f"错误：原校点 {master_name} 未找到，跳过合并 {school}。")
    return rows


def calc_art_sport_per_stu(ws, head_size, max_h, merge_map=None):
    """4. 每百名学生艺体教师数（含不计校数合并）"""
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    sport_col = find_col_by_hierarchy(ws, head_size, ['专任教师', '#按授课课程分', '体育与健康'])
    art_col = find_col_by_hierarchy(ws, head_size, ['专任教师', '#按授课课程分', '艺术', '计'])
    if not sport_col and not art_col:
        sport_col = find_col_by_keywords(ws, head_size, '体育与健康')
        art_col = find_col_by_keywords(ws, head_size, '艺术')
    if not stu_col or (not sport_col and not art_col):
        print("错误：未找到“在校生数”或体育/艺术相关列。")
        return None

    count_col = find_count_school_col(ws, head_size)
    header = ['原始行号', '学校名称', '在校生数(合计)', '体育教师', '艺术教师(计)', '每百名学生艺体教师数']
    rows = [header]
    school_index = {}
    pending_items = []

    for r in range(head_size + 1, max_h + 1):
        if count_col is not None:
            count_val = ws.cell(row=r, column=count_col).value
            try:
                if float(count_val) == 0:
                    school_name = ws.cell(row=r, column=1).value
                    if not school_name:
                        continue
                    stu_val = ws.cell(row=r, column=stu_col).value
                    sport_val = ws.cell(row=r, column=sport_col).value if sport_col else 0
                    art_val = ws.cell(row=r, column=art_col).value if art_col else 0
                    pending_items.append({
                        'row': r,
                        'school': str(school_name).strip(),
                        'stu': stu_val,
                        'sport': sport_val,
                        'art': art_val
                    })
                    continue
            except (ValueError, TypeError):
                pass

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
        row_data = [r, school_name, stu, sport, art, result]
        rows.append(row_data)
        school_index[str(school_name).strip()] = len(rows) - 1

    for item in pending_items:
        school = item['school']
        if merge_map and school in merge_map:
            master_name = merge_map[school]
        else:
            print(f"\n发现不计校数的学校：{school} (行{item['row']})")
            while True:
                master_name = input("请输入该校所属的原校点名称：").strip()
                if not master_name:
                    continue
                if master_name in school_index:
                    break
                print(f"未找到学校“{master_name}”。")

        if master_name in school_index:
            idx = school_index[master_name]
            old_stu = rows[idx][2]
            old_sport = rows[idx][3]
            old_art = rows[idx][4]
            try:
                add_stu = float(item['stu']) if item['stu'] is not None else 0
                add_sport = float(item['sport']) if item['sport'] is not None else 0
                add_art = float(item['art']) if item['art'] is not None else 0
            except (ValueError, TypeError):
                add_stu = add_sport = add_art = 0
            new_stu = old_stu + add_stu
            new_sport = old_sport + add_sport
            new_art = old_art + add_art
            new_total = new_sport + new_art
            new_result = round(new_total * 100 / new_stu, 2) if new_stu != 0 else 0
            rows[idx][2] = new_stu
            rows[idx][3] = new_sport
            rows[idx][4] = new_art
            rows[idx][5] = new_result
            print(f"已将 {school} 合并到 {master_name}，更新：学生={new_stu}, 体育={new_sport}, 艺术={new_art}, 每百名={new_result}")
        else:
            print(f"错误：原校点 {master_name} 未找到，跳过合并 {school}。")
    return rows

def calc_teaching_area_per_stu(ws, head_size, max_h, merge_map=None):
    """5. 生均教学及辅助用房面积（含不计校数合并）
    计算公式：净面积 = 教学及辅助用房(计) - 室内体育用房
    生均面积 = 净面积 / 在校生数
    
    表头结构：
    第1行：本学年校舍建筑面积
    第2行：（空白）| 教学及辅助用房 |
    第3行：计      | 室内体育用房   | ...
    第4行：（子项）| 教室           | ...
    """
    # 查找在校生数列
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    
    print(f"[调试-教学用房] stu_col={stu_col}")
    
    # 先找"教学及辅助用房"所在的行和列
    teach_row = None
    teach_base_col = None
    
    for r in range(1, head_size + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if val and isinstance(val, str) and '教学及辅助用房' in val:
                teach_row = r
                teach_base_col = c
                break
        if teach_row:
            break
    
    print(f"[调试-教学用房] '教学及辅助用房'位置: 第{teach_row}行, 第{teach_base_col}列")
    
    if not teach_row:
        print("错误：未找到'教学及辅助用房'。")
        print("\n表头内容预览（帮助调试）：")
        for row in range(1, head_size + 1):
            row_data = []
            for col in range(1, min(ws.max_column + 1, 30)):
                val = ws.cell(row=row, column=col).value
                if val:
                    row_data.append(f"列{col}:{val}")
            if row_data:
                print(f"  第{row}行: {' | '.join(row_data)}")
        return None
    
    # 在"教学及辅助用房"所在列，向下找"计"（紧挨着下方）
    teach_area_col = None
    
    # 策略1: 在"教学及辅助用房"同一列，向下查找"计"
    for r in range(teach_row + 1, head_size + 1):
        val = ws.cell(row=r, column=teach_base_col).value
        if val and isinstance(val, str) and val.strip() == '计':
            teach_area_col = teach_base_col
            break
    print(f"[调试-教学用房] 策略1 (同列向下): teach_area_col={teach_area_col}")
    
    if not teach_area_col:
        # 策略2: 在"教学及辅助用房"下方几列范围内查找"计"
        for c in range(teach_base_col - 2, teach_base_col + 3):
            if c < 1 or c > ws.max_column:
                continue
            for r in range(teach_row + 1, head_size + 1):
                val = ws.cell(row=r, column=c).value
                if val and isinstance(val, str) and val.strip() == '计':
                    teach_area_col = c
                    break
            if teach_area_col:
                break
        print(f"[调试-教学用房] 策略2 (附近列向下): teach_area_col={teach_area_col}")
    
    if not teach_area_col:
        # 策略3: 在整个表头中找"计"，要求其上方行有"教学及辅助用房"
        for c in range(1, ws.max_column + 1):
            for r in range(2, head_size + 1):  # 从第2行开始
                val = ws.cell(row=r, column=c).value
                if val and isinstance(val, str) and val.strip() == '计':
                    # 检查上方行是否有"教学及辅助用房"
                    for check_r in range(1, r):
                        check_val = ws.cell(row=check_r, column=c).value
                        if check_val and isinstance(check_val, str) and '教学及辅助用房' in check_val:
                            teach_area_col = c
                            break
                if teach_area_col:
                    break
            if teach_area_col:
                break
        print(f"[调试-教学用房] 策略3 (上方有教学及辅助用房的计): teach_area_col={teach_area_col}")
    
    if not teach_area_col:
        print("错误：未找到'教学及辅助用房'下方的'计'列。")
        return None
    
    # 查找"室内体育用房"列
    indoor_sport_col = None
    
    # 策略1: 在"教学及辅助用房"同一行右侧找"室内体育用房"
    for c in range(teach_base_col + 1, min(teach_base_col + 15, ws.max_column + 1)):
        val = ws.cell(row=teach_row, column=c).value
        if val and isinstance(val, str) and '室内体育用房' in val:
            indoor_sport_col = c
            break
    print(f"[调试-教学用房] 室内体育用房 策略1 (同行右侧): {indoor_sport_col}")
    
    if not indoor_sport_col:
        # 策略2: 在"计"所在行右侧找"室内体育用房"
        teach_ji_row = None
        for r in range(teach_row + 1, head_size + 1):
            if ws.cell(row=r, column=teach_area_col).value and isinstance(ws.cell(row=r, column=teach_area_col).value, str) and ws.cell(row=r, column=teach_area_col).value.strip() == '计':
                teach_ji_row = r
                break
        
        if teach_ji_row:
            for c in range(teach_area_col + 1, min(teach_area_col + 15, ws.max_column + 1)):
                val = ws.cell(row=teach_ji_row, column=c).value
                if val and isinstance(val, str) and '室内体育用房' in val:
                    indoor_sport_col = c
                    break
            print(f"[调试-教学用房] 室内体育用房 策略2 (计同行右侧): {indoor_sport_col}")
    
    if not indoor_sport_col:
        # 策略3: 直接关键词查找
        indoor_sport_col = find_col_by_keywords(ws, head_size, '室内体育用房')
        print(f"[调试-教学用房] 室内体育用房 策略3 (关键词): {indoor_sport_col}")
    
    print(f"[调试-教学用房] 最终结果:")
    print(f"  stu_col={stu_col}")
    print(f"  teach_area_col={teach_area_col}")
    print(f"  indoor_sport_col={indoor_sport_col}")
    
    if not stu_col:
        print("错误：未找到'在校生数'列。")
        return None
    
    # 如果没有找到室内体育用房，给出警告但继续执行
    if not indoor_sport_col:
        print("警告：未找到'室内体育用房'列，该项将按0计算。")
    
    count_col = find_count_school_col(ws, head_size)
    header = ['原始行号', '学校名称', '在校生数(合计)', '教学及辅助用房(计)', 
              '室内体育用房', '教学及辅助用房面积(净)', '生均教学及辅助用房面积']
    rows = [header]
    school_index = {}
    pending_items = []
    
    for r in range(head_size + 1, max_h + 1):
        # 检查是否是不计校数的学校
        if count_col is not None:
            count_val = ws.cell(row=r, column=count_col).value
            try:
                if float(count_val) == 0:
                    school_name = ws.cell(row=r, column=1).value
                    if not school_name:
                        continue
                    stu_val = ws.cell(row=r, column=stu_col).value
                    teach_val = ws.cell(row=r, column=teach_area_col).value
                    indoor_val = ws.cell(row=r, column=indoor_sport_col).value if indoor_sport_col else 0
                    pending_items.append({
                        'row': r,
                        'school': str(school_name).strip(),
                        'stu': stu_val,
                        'teach': teach_val,
                        'indoor': indoor_val
                    })
                    continue
            except (ValueError, TypeError):
                pass
        
        # 正常处理计校数的行
        school_name = ws.cell(row=r, column=1).value
        if not school_name:
            continue
            
        stu_val = ws.cell(row=r, column=stu_col).value
        teach_val = ws.cell(row=r, column=teach_area_col).value
        indoor_val = ws.cell(row=r, column=indoor_sport_col).value if indoor_sport_col else 0
        
        try:
            stu = float(stu_val) if stu_val is not None else None
            teach = float(teach_val) if teach_val is not None else 0
            indoor = float(indoor_val) if indoor_val is not None else 0
        except (ValueError, TypeError):
            continue
        
        if stu is None or stu == 0:
            continue
        
        # 计算净教学及辅助用房面积 = 教学及辅助用房(计) - 室内体育用房
        net_area = teach - indoor
        if net_area < 0:
            print(f"警告：学校{school_name}的净教学及辅助用房面积为负({net_area})，设为0")
            net_area = 0
        
        result = round(net_area / stu, 2)
        row_data = [r, school_name, stu, teach, indoor, net_area, result]
        rows.append(row_data)
        school_index[str(school_name).strip()] = len(rows) - 1
    
    print(f"[调试-教学用房] 处理了 {len(rows)-1} 行数据（不含表头）")
    
    # 处理不计校数的学校，合并到原校
    for item in pending_items:
        school = item['school']
        if merge_map and school in merge_map:
            master_name = merge_map[school]
        else:
            print(f"\n发现不计校数的学校：{school} (行{item['row']})")
            while True:
                master_name = input("请输入该校所属的原校点名称：").strip()
                if not master_name:
                    continue
                if master_name in school_index:
                    break
                print(f"未找到学校'{master_name}'。")
        
        if master_name in school_index:
            idx = school_index[master_name]
            old_stu = rows[idx][2]
            old_teach = rows[idx][3]
            old_indoor = rows[idx][4]
            old_net = rows[idx][5]
            
            try:
                add_stu = float(item['stu']) if item['stu'] is not None else 0
                add_teach = float(item['teach']) if item['teach'] is not None else 0
                add_indoor = float(item['indoor']) if item['indoor'] is not None else 0
            except (ValueError, TypeError):
                add_stu = add_teach = add_indoor = 0
            
            new_stu = old_stu + add_stu
            new_teach = old_teach + add_teach
            new_indoor = old_indoor + add_indoor
            new_net = new_teach - new_indoor
            if new_net < 0:
                new_net = 0
            new_result = round(new_net / new_stu, 2) if new_stu != 0 else 0
            
            rows[idx][2] = new_stu
            rows[idx][3] = new_teach
            rows[idx][4] = new_indoor
            rows[idx][5] = new_net
            rows[idx][6] = new_result
            print(f"已将 {school} 合并到 {master_name}，更新：学生={new_stu}, 教学用房(净)={new_net}, 生均={new_result}")
        else:
            print(f"错误：原校点 {master_name} 未找到，跳过合并 {school}。")
    
    if len(rows) == 1:
        print("警告：未找到任何有效数据行，请检查表头和数据。")
    
    return rows

def calc_sports_area_per_stu(ws, head_size, max_h, merge_map=None):
    """6. 生均体育运动场馆面积（含不计校数合并）
    计算公式：体育运动场馆面积 = 运动场地面积 + 室内体育用房
    生均面积 = 体育运动场馆面积 / 在校生数
    
    表头结构：
    占地面积及其他办学条件 -> 占地面积（平方米） -> 其中 -> 运动场地面积
    """
    # 查找在校生数列
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    
    print(f"[调试-运动场馆] stu_col={stu_col}")
    
    # 查找"运动场地面积"列
    sports_field_col = None
    
    # 策略1: 四层精确查找
    sports_field_col = find_col_by_hierarchy(ws, head_size, ['占地面积及其他办学条件', '占地面积（平方米）', '其中', '运动场地面积'])
    print(f"[调试-运动场馆] 策略1: {sports_field_col}")
    
    if not sports_field_col:
        # 策略2: "占地面积（平方米）" -> "其中" -> "运动场地面积"
        sports_field_col = find_col_by_hierarchy(ws, head_size, ['占地面积（平方米）', '其中', '运动场地面积'])
        print(f"[调试-运动场馆] 策略2: {sports_field_col}")
    
    if not sports_field_col:
        # 策略3: "其中" -> "运动场地面积"
        sports_field_col = find_col_by_hierarchy(ws, head_size, ['其中', '运动场地面积'])
        print(f"[调试-运动场馆] 策略3: {sports_field_col}")
    
    if not sports_field_col:
        # 策略4: 在"占地面积及其他办学条件"区域(295列开始)内查找"运动场地面积"
        area_start = find_col_by_keywords(ws, head_size, '占地面积及其他办学条件')
        if not area_start:
            area_start = find_col_by_keywords(ws, head_size, '占地面积')
        
        print(f"[调试-运动场馆] 策略4 area_start={area_start}")
        
        if area_start:
            # 在area_start到area_start+30范围内查找
            for col in range(area_start, min(area_start + 30, ws.max_column + 1)):
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str) and '运动场地面积' in val:
                        sports_field_col = col
                        break
                if sports_field_col:
                    break
        print(f"[调试-运动场馆] 策略4: {sports_field_col}")
    
    if not sports_field_col:
        # 策略5: 查找"运动场地"（去掉"面积"）
        area_start = find_col_by_keywords(ws, head_size, '占地面积及其他办学条件')
        if not area_start:
            area_start = find_col_by_keywords(ws, head_size, '占地面积')
        
        if area_start:
            for col in range(area_start, min(area_start + 30, ws.max_column + 1)):
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str) and '运动场地' in val:
                        sports_field_col = col
                        break
                if sports_field_col:
                    break
        print(f"[调试-运动场馆] 策略5: {sports_field_col}")
    
    if not sports_field_col:
        # 策略6: 直接关键词查找"运动场地面积"
        sports_field_col = find_col_by_keywords(ws, head_size, '运动场地面积')
        print(f"[调试-运动场馆] 策略6: {sports_field_col}")
    
    if not sports_field_col:
        # 策略7: 查找"运动场地"
        sports_field_col = find_col_by_keywords(ws, head_size, '运动场地')
        print(f"[调试-运动场馆] 策略7: {sports_field_col}")
    
    if not sports_field_col:
        # 策略8: 打印第4、5、6行所有包含"运动"或"场地"的列
        print("[调试-运动场馆] 搜索包含'运动'或'场地'的列:")
        for col in range(1, ws.max_column + 1):
            for r in range(1, head_size + 1):
                val = ws.cell(row=r, column=col).value
                if val and isinstance(val, str) and ('运动' in val or '场地' in val):
                    print(f"  第{r}行 第{col}列: {val}")
                    break
    
    # 查找"室内体育用房"列
    indoor_sport_col = None
    
    # 策略1: 层级查找
    indoor_sport_col = find_col_by_hierarchy(ws, head_size, ['本学年校舍建筑面积', '室内体育用房'])
    print(f"[调试-运动场馆] 室内体育用房 策略1: {indoor_sport_col}")
    
    if not indoor_sport_col:
        indoor_sport_col = find_col_by_hierarchy(ws, head_size, ['校舍建筑面积', '室内体育用房'])
        print(f"[调试-运动场馆] 室内体育用房 策略2: {indoor_sport_col}")
    
    if not indoor_sport_col:
        # 在"校舍"区域查找
        building_start = find_col_by_keywords(ws, head_size, '本学年校舍建筑面积')
        if not building_start:
            building_start = find_col_by_keywords(ws, head_size, '校舍建筑面积')
        
        if building_start:
            for col in range(building_start, min(building_start + 30, ws.max_column + 1)):
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str) and '室内体育用房' in val:
                        indoor_sport_col = col
                        break
                if indoor_sport_col:
                    break
        print(f"[调试-运动场馆] 室内体育用房 策略3: {indoor_sport_col}")
    
    if not indoor_sport_col:
        indoor_sport_col = find_col_by_keywords(ws, head_size, '室内体育用房')
        print(f"[调试-运动场馆] 室内体育用房 策略4: {indoor_sport_col}")
    
    print(f"[调试-运动场馆] 最终结果:")
    print(f"  stu_col={stu_col}")
    print(f"  sports_field_col={sports_field_col}")
    print(f"  indoor_sport_col={indoor_sport_col}")
    
    if not stu_col:
        print("错误：未找到'在校生数'列。")
        return None
    
    if not sports_field_col:
        print("错误：未找到'运动场地面积'列。")
        # 已经打印了调试信息
        return None
    
    if not indoor_sport_col:
        print("警告：未找到'室内体育用房'列，该项将按0计算。")
    
    count_col = find_count_school_col(ws, head_size)
    header = ['原始行号', '学校名称', '在校生数(合计)', '运动场地面积', 
              '室内体育用房', '体育运动场馆面积', '生均体育运动场馆面积']
    rows = [header]
    school_index = {}
    pending_items = []
    
    for r in range(head_size + 1, max_h + 1):
        # 检查是否是不计校数的学校
        if count_col is not None:
            count_val = ws.cell(row=r, column=count_col).value
            try:
                if float(count_val) == 0:
                    school_name = ws.cell(row=r, column=1).value
                    if not school_name:
                        continue
                    stu_val = ws.cell(row=r, column=stu_col).value
                    sports_val = ws.cell(row=r, column=sports_field_col).value
                    indoor_val = ws.cell(row=r, column=indoor_sport_col).value if indoor_sport_col else 0
                    pending_items.append({
                        'row': r,
                        'school': str(school_name).strip(),
                        'stu': stu_val,
                        'sports': sports_val,
                        'indoor': indoor_val
                    })
                    continue
            except (ValueError, TypeError):
                pass
        
        # 正常处理计校数的行
        school_name = ws.cell(row=r, column=1).value
        if not school_name:
            continue
            
        stu_val = ws.cell(row=r, column=stu_col).value
        sports_val = ws.cell(row=r, column=sports_field_col).value
        indoor_val = ws.cell(row=r, column=indoor_sport_col).value if indoor_sport_col else 0
        
        try:
            stu = float(stu_val) if stu_val is not None else None
            sports = float(sports_val) if sports_val is not None else 0
            indoor = float(indoor_val) if indoor_val is not None else 0
        except (ValueError, TypeError):
            continue
        
        if stu is None or stu == 0:
            continue
        
        # 计算体育运动场馆面积 = 运动场地面积 + 室内体育用房
        total_sports_area = sports + indoor
        
        result = round(total_sports_area / stu, 2)
        row_data = [r, school_name, stu, sports, indoor, total_sports_area, result]
        rows.append(row_data)
        school_index[str(school_name).strip()] = len(rows) - 1
    
    print(f"[调试-运动场馆] 处理了 {len(rows)-1} 行数据（不含表头）")
    
    # 处理不计校数的学校，合并到原校
    for item in pending_items:
        school = item['school']
        if merge_map and school in merge_map:
            master_name = merge_map[school]
        else:
            print(f"\n发现不计校数的学校：{school} (行{item['row']})")
            while True:
                master_name = input("请输入该校所属的原校点名称：").strip()
                if not master_name:
                    continue
                if master_name in school_index:
                    break
                print(f"未找到学校'{master_name}'。")
        
        if master_name in school_index:
            idx = school_index[master_name]
            old_stu = rows[idx][2]
            old_sports = rows[idx][3]
            old_indoor = rows[idx][4]
            old_total = rows[idx][5]
            
            try:
                add_stu = float(item['stu']) if item['stu'] is not None else 0
                add_sports = float(item['sports']) if item['sports'] is not None else 0
                add_indoor = float(item['indoor']) if item['indoor'] is not None else 0
            except (ValueError, TypeError):
                add_stu = add_sports = add_indoor = 0
            
            new_stu = old_stu + add_stu
            new_sports = old_sports + add_sports
            new_indoor = old_indoor + add_indoor
            new_total = new_sports + new_indoor
            new_result = round(new_total / new_stu, 2) if new_stu != 0 else 0
            
            rows[idx][2] = new_stu
            rows[idx][3] = new_sports
            rows[idx][4] = new_indoor
            rows[idx][5] = new_total
            rows[idx][6] = new_result
            print(f"已将 {school} 合并到 {master_name}，更新：学生={new_stu}, 运动场馆={new_total}, 生均={new_result}")
        else:
            print(f"错误：原校点 {master_name} 未找到，跳过合并 {school}。")
    
    if len(rows) == 1:
        print("警告：未找到任何有效数据行，请检查表头和数据。")
    
    return rows

def calc_equipment_per_stu(ws, head_size, max_h, merge_map=None):
    """7. 生均教学仪器设备值（含不计校数合并）
    计算公式：生均教学仪器设备值 = 教学仪器设备资产值(万元) * 10000 / 在校生数
    
    表头结构：
    固定资产总值 -> 教学仪器设备资产值(万元)
    """
    # 查找在校生数列
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    
    print(f"[调试-仪器设备] stu_col={stu_col}")
    
    # 查找"教学仪器设备资产值(万元)"列
    equipment_col = None
    
    # 策略1: "固定资产总值" -> "教学仪器设备资产值(万元)"
    equipment_col = find_col_by_hierarchy(ws, head_size, ['固定资产总值', '教学仪器设备资产值(万元)'])
    print(f"[调试-仪器设备] 策略1: {equipment_col}")
    
    if not equipment_col:
        # 策略2: "固定资产总值" -> "教学仪器设备资产值"
        equipment_col = find_col_by_hierarchy(ws, head_size, ['固定资产总值', '教学仪器设备资产值'])
        print(f"[调试-仪器设备] 策略2: {equipment_col}")
    
    if not equipment_col:
        # 策略3: "固定资产总值" -> "教学仪器设备"
        equipment_col = find_col_by_hierarchy(ws, head_size, ['固定资产总值', '教学仪器设备'])
        print(f"[调试-仪器设备] 策略3: {equipment_col}")
    
    if not equipment_col:
        # 策略4: 在"固定资产总值"区域内查找"教学仪器设备资产值"
        asset_start = find_col_by_keywords(ws, head_size, '固定资产总值')
        print(f"[调试-仪器设备] 策略4 asset_start={asset_start}")
        if asset_start:
            for col in range(asset_start, min(asset_start + 15, ws.max_column + 1)):
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str) and '教学仪器设备资产值' in val:
                        equipment_col = col
                        break
                if equipment_col:
                    break
        print(f"[调试-仪器设备] 策略4: {equipment_col}")
    
    if not equipment_col:
        # 策略5: 在"固定资产总值"区域内查找"教学仪器设备"
        asset_start = find_col_by_keywords(ws, head_size, '固定资产总值')
        if asset_start:
            for col in range(asset_start, min(asset_start + 15, ws.max_column + 1)):
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str) and '教学仪器设备' in val:
                        equipment_col = col
                        break
                if equipment_col:
                    break
        print(f"[调试-仪器设备] 策略5: {equipment_col}")
    
    if not equipment_col:
        # 策略6: 直接关键词查找"教学仪器设备资产值(万元)"
        equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备资产值(万元)')
        print(f"[调试-仪器设备] 策略6: {equipment_col}")
    
    if not equipment_col:
        # 策略7: 直接关键词查找"教学仪器设备资产值"
        equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备资产值')
        print(f"[调试-仪器设备] 策略7: {equipment_col}")
    
    if not equipment_col:
        # 策略8: 直接关键词查找"教学仪器设备"
        equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备')
        print(f"[调试-仪器设备] 策略8: {equipment_col}")
    
    if not equipment_col:
        # 策略9: 查找"仪器设备"
        equipment_col = find_col_by_keywords(ws, head_size, '仪器设备')
        print(f"[调试-仪器设备] 策略9: {equipment_col}")
    
    if not equipment_col:
        # 策略10: 打印所有包含"仪器"或"设备"的列，帮助调试
        print("[调试-仪器设备] 搜索包含'仪器'或'设备'或'资产'的列:")
        for col in range(1, min(ws.max_column + 1, 400)):
            for r in range(1, head_size + 1):
                val = ws.cell(row=r, column=col).value
                if val and isinstance(val, str) and ('仪器' in val or '设备' in val or '资产' in val or '固定资产' in val):
                    print(f"  第{r}行 第{col}列: {val}")
                    break
    
    print(f"[调试-仪器设备] 最终结果:")
    print(f"  stu_col={stu_col}")
    print(f"  equipment_col={equipment_col}")
    
    if not stu_col:
        print("错误：未找到'在校生数'列。")
        return None
    
    if not equipment_col:
        print("错误：未找到'教学仪器设备资产值(万元)'列。")
        print("请检查表头是否包含：'固定资产总值' -> '教学仪器设备资产值(万元)'")
        # 打印表头帮助调试
        print("\n表头内容预览（帮助调试）：")
        for row in range(1, head_size + 1):
            row_data = []
            for col in range(1, min(ws.max_column + 1, 30)):
                val = ws.cell(row=row, column=col).value
                if val:
                    row_data.append(f"列{col}:{val}")
            if row_data:
                print(f"  第{row}行: {' | '.join(row_data)}")
        return None
    
    count_col = find_count_school_col(ws, head_size)
    header = ['原始行号', '学校名称', '在校生数(合计)', '教学仪器设备资产值(万元)', '生均教学仪器设备值']
    rows = [header]
    school_index = {}
    pending_items = []
    
    for r in range(head_size + 1, max_h + 1):
        # 检查是否是不计校数的学校
        if count_col is not None:
            count_val = ws.cell(row=r, column=count_col).value
            try:
                if float(count_val) == 0:
                    school_name = ws.cell(row=r, column=1).value
                    if not school_name:
                        continue
                    stu_val = ws.cell(row=r, column=stu_col).value
                    equipment_val = ws.cell(row=r, column=equipment_col).value
                    pending_items.append({
                        'row': r,
                        'school': str(school_name).strip(),
                        'stu': stu_val,
                        'equipment': equipment_val
                    })
                    continue
            except (ValueError, TypeError):
                pass
        
        # 正常处理计校数的行
        school_name = ws.cell(row=r, column=1).value
        if not school_name:
            continue
            
        stu_val = ws.cell(row=r, column=stu_col).value
        equipment_val = ws.cell(row=r, column=equipment_col).value
        
        try:
            stu = float(stu_val) if stu_val is not None else None
            equipment = float(equipment_val) if equipment_val is not None else 0
        except (ValueError, TypeError):
            continue
        
        if stu is None or stu == 0:
            continue
        
        # 计算生均教学仪器设备值 = 教学仪器设备资产值(万元) * 10000 / 在校生数
        result = round(equipment * 10000 / stu, 2)
        row_data = [r, school_name, stu, equipment, result]
        rows.append(row_data)
        school_index[str(school_name).strip()] = len(rows) - 1
    
    print(f"[调试-仪器设备] 处理了 {len(rows)-1} 行数据（不含表头）")
    
    # 处理不计校数的学校，合并到原校
    for item in pending_items:
        school = item['school']
        if merge_map and school in merge_map:
            master_name = merge_map[school]
        else:
            print(f"\n发现不计校数的学校：{school} (行{item['row']})")
            while True:
                master_name = input("请输入该校所属的原校点名称：").strip()
                if not master_name:
                    continue
                if master_name in school_index:
                    break
                print(f"未找到学校'{master_name}'。")
        
        if master_name in school_index:
            idx = school_index[master_name]
            old_stu = rows[idx][2]
            old_equipment = rows[idx][3]
            
            try:
                add_stu = float(item['stu']) if item['stu'] is not None else 0
                add_equipment = float(item['equipment']) if item['equipment'] is not None else 0
            except (ValueError, TypeError):
                add_stu = add_equipment = 0
            
            new_stu = old_stu + add_stu
            new_equipment = old_equipment + add_equipment
            new_result = round(new_equipment * 10000 / new_stu, 2) if new_stu != 0 else 0
            
            rows[idx][2] = new_stu
            rows[idx][3] = new_equipment
            rows[idx][4] = new_result
            print(f"已将 {school} 合并到 {master_name}，更新：学生={new_stu}, 仪器设备值={new_equipment}万元, 生均={new_result}")
        else:
            print(f"错误：原校点 {master_name} 未找到，跳过合并 {school}。")
    
    if len(rows) == 1:
        print("警告：未找到任何有效数据行，请检查表头和数据。")
    
    return rows

def analyze_district_data(ws, head_size, max_h, school_type, merge_map=None):
    # ==================== 1. 定位所有需要的列 ====================
    # 区县列
    district_col = find_col_by_keywords(ws, head_size, '统计三级')
    if not district_col:
        district_col = find_col_by_hierarchy(ws, head_size, ['统计三级'])
    if not district_col:
        print("错误：未找到“统计三级”列。")
        return None, None, None

    # 学生数
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    if not stu_col:
        print("错误：未找到“在校生数”列。")
        return None, None, None

    # 多媒体教室
    media_col = find_col_by_hierarchy(ws, head_size, ['网络多媒体教室'])
    if not media_col:
        media_col = find_col_by_keywords(ws, head_size, '网络多媒体教室')

    # 学历教师各列（用于指标2）
    edu_region = find_cols_with_prefix(ws, head_size, ['专任教师', '#按学历分'])
    if not edu_region:
        kw_primary = ['硕士研究生', '本科', '专科']
        kw_middle = ['硕士研究生', '本科']
        keywords = kw_primary if school_type == '小学' else kw_middle
        edu_cols = []
        for kw in keywords:
            col = find_col_by_keywords(ws, head_size, kw)
            if col:
                edu_cols.append(col)
    else:
        kw_primary = ['硕士研究生', '本科', '专科']
        kw_middle = ['硕士研究生', '本科']
        keywords = kw_primary if school_type == '小学' else kw_middle
        edu_cols = []
        for kw in keywords:
            found = None
            for col in edu_region:
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str) and kw in val:
                        found = col
                        break
                if found:
                    edu_cols.append(found)
                    break

    # 骨干教师
    backbone_col = find_col_by_hierarchy(ws, head_size, ['教基1102', '县级以上骨干教师(小学)' if school_type=='小学' else '县级以上骨干教师(初中)'])
    if not backbone_col:
        backbone_col = find_col_by_keywords(ws, head_size, '骨干教师')

    # 体育、艺术
    sport_col = find_col_by_hierarchy(ws, head_size, ['专任教师', '#按授课课程分', '体育与健康'])
    art_col = find_col_by_hierarchy(ws, head_size, ['专任教师', '#按授课课程分', '艺术', '计'])
    if not sport_col:
        sport_col = find_col_by_keywords(ws, head_size, '体育与健康')
    if not art_col:
        art_col = find_col_by_keywords(ws, head_size, '艺术')

    # 教学及辅助用房(计) 和 室内体育用房（用于指标5、6）
    teach_area_col = None
    indoor_sport_col = None

    # 先找“教学及辅助用房”
    teach_row = None
    teach_base_col = None
    for r in range(1, head_size + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if val and isinstance(val, str) and '教学及辅助用房' in val:
                teach_row = r
                teach_base_col = c
                break
        if teach_row:
            break
    if teach_base_col:
        # 在下方找“计”
        for r in range(teach_row + 1, head_size + 1):
            val = ws.cell(row=r, column=teach_base_col).value
            if val and isinstance(val, str) and val.strip() == '计':
                teach_area_col = teach_base_col
                break
        if not teach_area_col:
            # 附近列找“计”
            for c in range(max(1, teach_base_col-2), min(ws.max_column, teach_base_col+2)+1):
                for r in range(teach_row + 1, head_size + 1):
                    val = ws.cell(row=r, column=c).value
                    if val and isinstance(val, str) and val.strip() == '计':
                        teach_area_col = c
                        break
                if teach_area_col:
                    break
        if not teach_area_col:
            teach_area_col = find_col_by_keywords(ws, head_size, '教学及辅助用房')  # 降级
    if not teach_area_col:
        print("警告：未找到“教学及辅助用房(计)”列，净面积将无法计算。")

    # 室内体育用房（可能在教学及辅助用房右侧或独立）
    indoor_sport_col = find_col_by_hierarchy(ws, head_size, ['本学年校舍建筑面积', '室内体育用房'])
    if not indoor_sport_col:
        indoor_sport_col = find_col_by_hierarchy(ws, head_size, ['校舍建筑面积', '室内体育用房'])
    if not indoor_sport_col:
        # 在教学及辅助用房同行或附近找
        if teach_row:
            for c in range(teach_base_col, min(teach_base_col+10, ws.max_column+1)):
                for r in range(1, head_size+1):
                    val = ws.cell(row=r, column=c).value
                    if val and isinstance(val, str) and '室内体育用房' in val:
                        indoor_sport_col = c
                        break
                if indoor_sport_col:
                    break
    if not indoor_sport_col:
        indoor_sport_col = find_col_by_keywords(ws, head_size, '室内体育用房')

    # 运动场地面积
    sports_field_col = find_col_by_hierarchy(ws, head_size, ['占地面积及其他办学条件', '占地面积（平方米）', '其中', '运动场地面积'])
    if not sports_field_col:
        sports_field_col = find_col_by_hierarchy(ws, head_size, ['占地面积（平方米）', '其中', '运动场地面积'])
    if not sports_field_col:
        sports_field_col = find_col_by_hierarchy(ws, head_size, ['其中', '运动场地面积'])
    if not sports_field_col:
        sports_field_col = find_col_by_keywords(ws, head_size, '运动场地面积')
    if not sports_field_col:
        sports_field_col = find_col_by_keywords(ws, head_size, '运动场地')

    # 教学仪器设备值(万元)
    equipment_col = find_col_by_hierarchy(ws, head_size, ['固定资产总值', '教学仪器设备资产值(万元)'])
    if not equipment_col:
        equipment_col = find_col_by_hierarchy(ws, head_size, ['固定资产总值', '教学仪器设备资产值'])
    if not equipment_col:
        equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备资产值(万元)')
    if not equipment_col:
        equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备资产值')
    if not equipment_col:
        equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备')

    count_col = find_count_school_col(ws, head_size)

    print(f"[调试] 列定位: stu={stu_col}, media={media_col}, edu_cols={edu_cols}, backbone={backbone_col}, sport={sport_col}, art={art_col}, teach_area={teach_area_col}, indoor={indoor_sport_col}, sports_field={sports_field_col}, equipment={equipment_col}")

    # ==================== 2. 读取数据，按区县分组 ====================
    district_data = OrderedDict()
    pending_items = []

    for r in range(head_size + 1, max_h + 1):
        # 不计校数行收集
        if count_col is not None:
            count_val = ws.cell(row=r, column=count_col).value
            try:
                if float(count_val) == 0:
                    school_name = str(ws.cell(row=r, column=1).value or '').strip()
                    if not school_name:
                        continue
                    district_cell = ws.cell(row=r, column=district_col).value
                    district = str(district_cell).strip() if district_cell else ''
                    # 收集所有原始值
                    vals = {
                        'stu': ws.cell(row=r, column=stu_col).value,
                        'media': ws.cell(row=r, column=media_col).value if media_col else 0,
                        'edu_vals': {c: ws.cell(row=r, column=c).value for c in edu_cols},
                        'back': ws.cell(row=r, column=backbone_col).value if backbone_col else 0,
                        'sport': ws.cell(row=r, column=sport_col).value if sport_col else 0,
                        'art': ws.cell(row=r, column=art_col).value if art_col else 0,
                        'teach_area': ws.cell(row=r, column=teach_area_col).value if teach_area_col else 0,
                        'indoor': ws.cell(row=r, column=indoor_sport_col).value if indoor_sport_col else 0,
                        'sports_field': ws.cell(row=r, column=sports_field_col).value if sports_field_col else 0,
                        'equipment': ws.cell(row=r, column=equipment_col).value if equipment_col else 0,
                    }
                    pending_items.append({
                        'row': r,
                        'school': school_name,
                        'district': district,
                        'vals': vals
                    })
                    continue
            except (ValueError, TypeError):
                pass

        # 正常行
        district_cell = ws.cell(row=r, column=district_col).value
        if not district_cell:
            continue
        district = str(district_cell).strip()
        if district in ('合计', '总计', '小计', '备注'):
            continue

        school_name = str(ws.cell(row=r, column=1).value or '').strip()
        if not school_name:
            continue

        # 读取所有数值，错误则跳过
        def safe_float(val, default=0.0):
            try:
                return float(val) if val is not None else default
            except:
                return default

        stu = safe_float(ws.cell(row=r, column=stu_col).value, None)
        if stu is None or stu == 0:
            continue

        media = safe_float(ws.cell(row=r, column=media_col).value) if media_col else 0
        edu_sum = sum(safe_float(ws.cell(row=r, column=c).value) for c in edu_cols)
        back = safe_float(ws.cell(row=r, column=backbone_col).value) if backbone_col else 0
        sport = safe_float(ws.cell(row=r, column=sport_col).value) if sport_col else 0
        art = safe_float(ws.cell(row=r, column=art_col).value) if art_col else 0
        teach_area = safe_float(ws.cell(row=r, column=teach_area_col).value) if teach_area_col else 0
        indoor = safe_float(ws.cell(row=r, column=indoor_sport_col).value) if indoor_sport_col else 0
        sports_field = safe_float(ws.cell(row=r, column=sports_field_col).value) if sports_field_col else 0
        equipment = safe_float(ws.cell(row=r, column=equipment_col).value) if equipment_col else 0

        # 计算七个指标原始值
        ind1 = media
        ind2 = edu_sum
        ind3 = back
        ind4 = sport + art
        ind5 = teach_area - indoor   # 净面积
        ind6 = sports_field + indoor # 运动场馆
        ind7 = equipment             # 万元

        if district not in district_data:
            district_data[district] = []
        district_data[district].append({
            'school': school_name,
            'A': stu,
            'ind1': ind1, 'ind2': ind2, 'ind3': ind3, 'ind4': ind4,
            'ind5': ind5, 'ind6': ind6, 'ind7': ind7
        })

    # ==================== 3. 合并不计校数行 ====================
    school_map = {}
    for dist, entries in district_data.items():
        for idx, entry in enumerate(entries):
            if entry['school']:
                school_map[entry['school']] = (dist, idx)

    for item in pending_items:
        school = item['school']
        if merge_map and school in merge_map:
            master_name = merge_map[school]
        else:
            print(f"\n[区县统计] 不计校数学校：{school} (行{item['row']})，所属区县：{item['district']}")
            while True:
                master_name = input("请输入该校所属的原校点名称：").strip()
                if not master_name:
                    continue
                if master_name in school_map:
                    break
                print(f"未找到学校“{master_name}”")
        if master_name in school_map:
            dist, idx = school_map[master_name]
            entry = district_data[dist][idx]
            vals = item['vals']
            # 累加所有指标值
            add_stu = safe_float(vals['stu'])
            add_media = safe_float(vals['media'])
            add_edu = sum(safe_float(vals['edu_vals'][c]) for c in edu_cols)
            add_back = safe_float(vals['back'])
            add_sport = safe_float(vals['sport'])
            add_art = safe_float(vals['art'])
            add_teach = safe_float(vals['teach_area'])
            add_indoor = safe_float(vals['indoor'])
            add_sports_field = safe_float(vals['sports_field'])
            add_equipment = safe_float(vals['equipment'])

            entry['A'] += add_stu
            entry['ind1'] += add_media
            entry['ind2'] += add_edu
            entry['ind3'] += add_back
            entry['ind4'] += (add_sport + add_art)
            entry['ind5'] += (add_teach - add_indoor)
            entry['ind6'] += (add_sports_field + add_indoor)
            entry['ind7'] += add_equipment
            print(f"已将 {school} 合并到 {master_name} (区县: {dist})")
        else:
            print(f"错误：原校点 {master_name} 未找到，跳过合并 {school}。")

    if not district_data:
        print("未提取到任何区县数据。")
        return None, None, None

    # ==================== 4. 计算统计量 ====================
    # 定义指标名称、multiplier
    indicators = [
        ('多媒体教室', 100),
        ('高学历教师', 100),
        ('骨干教师', 100),
        ('艺体教师', 100),
        ('教学及辅助用房净面积', 1),   # 平方米/生
        ('运动场馆面积', 1),            # 平方米/生
        ('教学仪器设备值', 10000)       # 元/生
    ]

    def calc_stats(data_list, ind_key, multiplier):
        A_sum = sum(d['A'] for d in data_list)
        if A_sum == 0:
            return {'X': 0, 'S': 0, 'CV': 0}
        indicator_sum = sum(d[ind_key] for d in data_list)
        X = (indicator_sum * multiplier) / A_sum
        variance = 0
        for d in data_list:
            if d['A'] > 0:
                ratio = (d[ind_key] * multiplier) / d['A']
                variance += (d['A'] / A_sum) * (ratio - X)**2
        S = np.sqrt(variance) if variance > 0 else 0
        CV = round((S / X), 2) if X != 0 else 0
        return {'X': round(X, 4), 'S': round(S, 6), 'CV': CV}

    # 分组
    group_counter = 1
    district_groups = OrderedDict()
    for dist in district_data:
        district_groups[dist] = {
            'group': f"T{group_counter}",
            'data': district_data[dist]
        }
        group_counter += 1

    # 构建差异系数表头
    header_parts = ['组别', '区县']
    for name, _ in indicators:
        header_parts.append(f'{name}平均值')
        header_parts.append(f'{name}标准差')
        header_parts.append(f'{name}差异系数')
    stats_rows = [header_parts]

    # 汇总各指标的所有数据（用于总计）
    all_data_lists = {f'ind{i+1}': [] for i in range(7)}
    for dist, info in district_groups.items():
        data = info['data']
        row = [info['group'], dist]
        for i, (name, mult) in enumerate(indicators):
            ind_key = f'ind{i+1}'
            stats = calc_stats(data, ind_key, mult)
            row.extend([stats['X'], stats['S'], stats['CV']])
        stats_rows.append(row)
        for i in range(7):
            all_data_lists[f'ind{i+1}'].extend(data)

    # 总计行
    total_row = ['', '总计']
    for i, (name, mult) in enumerate(indicators):
        stats = calc_stats(all_data_lists[f'ind{i+1}'], f'ind{i+1}', mult)
        total_row.extend([stats['X'], stats['S'], stats['CV']])
    stats_rows.append(total_row)

    # ==================== 5. 区县汇总表 ====================
    summary_header = ['区县', '学生总数(A)'] + [name for name, _ in indicators]
    summary_rows = [summary_header]
    total_A = 0
    total_inds = [0]*7
    for dist, info in district_groups.items():
        data = info['data']
        sA = sum(d['A'] for d in data)
        row = [dist, sA]
        for i in range(7):
            si = sum(d[f'ind{i+1}'] for d in data)
            row.append(round(si, 2))
            total_inds[i] += si
        summary_rows.append(row)
        total_A += sA
    summary_rows.append(['总计', total_A] + [round(v, 2) for v in total_inds])

    return district_data, stats_rows, summary_rows

def save_results_to_new_file(original_path, school_type, calc_results, stats_rows, summary_rows):
    dir_name = os.path.dirname(original_path)
    base_name = os.path.splitext(os.path.basename(original_path))[0]
    new_name = f"{base_name}_分析结果.xlsx"
    new_path = os.path.join(dir_name, new_name)

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    sheet_names = [
        '每百名学生多媒体教室数',
        '每百名学生高学历教师数',
        '每百名学生骨干教师数',
        '每百名学生艺体教师数',
        '生均教学及辅助用房面积',
        '生均运动场馆面积',
        '生均教学仪器设备值',
        '区县差异系数统计',
        '区县汇总'
    ]
    data_blocks = [
        calc_results.get('media'),
        calc_results.get('high_edu'),
        calc_results.get('backbone'),
        calc_results.get('art_sport'),
        calc_results.get('teaching_area'),
        calc_results.get('sports_area'),
        calc_results.get('equipment'),
        stats_rows,
        summary_rows
    ]

    for sname, block in zip(sheet_names, data_blocks):
        ws_out = out_wb.create_sheet(title=sname)
        if block:
            for row in block:
                ws_out.append(row)
        else:
            ws_out.append(['无数据'])
    out_wb.save(new_path)
    print(f"\n分析结果已保存至：{new_path}")
    return new_path

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
        wb, ws, max_h, max_l = get_worksheet_info(file_path)
        if wb is None:
            continue
        print(f"\n当前文件：{os.path.basename(file_path)}  类型：{school_type}")
        print(f"有效行数：{max_h}，有效列数：{max_l}")

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

        # -------------------- 构建不计校数合并映射（只询问一次）--------------------
        merge_map = build_school_merge_map(ws, head_size)

        calc_results = {
            'media': None,
            'high_edu': None,
            'backbone': None,
            'art_sport': None,
            'teaching_area': None,
            'sports_area': None,
            'equipment': None,
        }

        while True:
            print("\n" + "-" * 40)
            print("请选择计算项目：")
            print("  1. 每百名学生多媒体教室数")
            print("  2. 每百名学生高学历教师数")
            print("  3. 每百名学生骨干教师数")
            print("  4. 每百名学生艺体教师数")
            print("  5. 生均教学及辅助用房面积")  # 新增
            print("  6. 生均运动场馆面积")  # 新增
            print("  7. 生均教学仪器设备值")  # 新增
            print("  8. 一键执行以上全部 + 区县差异系数与汇总")
            print("  0. 完成选择，生成报告文件")
            print("-" * 40)
            choice = input("请输入选项：").strip()
            if choice == '0':
                break
            elif choice == '1':
                if calc_results['media'] is None:
                    calc_results['media'] = calc_media_per_stu(ws, head_size, max_h, merge_map)
                else:
                    print("该计算已完成。")
            elif choice == '2':
                if calc_results['high_edu'] is None:
                    calc_results['high_edu'] = calc_high_education_per_stu(ws, head_size, max_h, school_type, merge_map)
                else:
                    print("该计算已完成。")
            elif choice == '3':
                if calc_results['backbone'] is None:
                    calc_results['backbone'] = calc_backbone_per_stu(ws, head_size, max_h, school_type, merge_map)
                else:
                    print("该计算已完成。")
            elif choice == '4':
                if calc_results['art_sport'] is None:
                    calc_results['art_sport'] = calc_art_sport_per_stu(ws, head_size, max_h, merge_map)
                else:
                    print("该计算已完成。")
            elif choice == '5':
                if calc_results['teaching_area'] is None:
                    calc_results['teaching_area'] = calc_teaching_area_per_stu(ws, head_size, max_h, merge_map)
                else:
                    print("该计算已完成。")
            elif choice == '6':  # 新增
                if calc_results['sports_area'] is None:
                    calc_results['sports_area'] = calc_sports_area_per_stu(ws, head_size, max_h, merge_map)
                else:
                    print("该计算已完成。")
            elif choice == '7':  # 新增
                if calc_results['equipment'] is None:
                    calc_results['equipment'] = calc_equipment_per_stu(ws, head_size, max_h, merge_map)
                else:
                    print("该计算已完成。")
            elif choice == '8':
                if calc_results['media'] is None:
                    calc_results['media'] = calc_media_per_stu(ws, head_size, max_h, merge_map)
                if calc_results['high_edu'] is None:
                    calc_results['high_edu'] = calc_high_education_per_stu(ws, head_size, max_h, school_type, merge_map)
                if calc_results['backbone'] is None:
                    calc_results['backbone'] = calc_backbone_per_stu(ws, head_size, max_h, school_type, merge_map)
                if calc_results['art_sport'] is None:
                    calc_results['art_sport'] = calc_art_sport_per_stu(ws, head_size, max_h, merge_map)
                if calc_results['teaching_area'] is None:
                    calc_results['teaching_area'] = calc_teaching_area_per_stu(ws, head_size, max_h, merge_map)
                if calc_results['sports_area'] is None:  # 新增
                    calc_results['sports_area'] = calc_sports_area_per_stu(ws, head_size, max_h, merge_map)
                if calc_results['equipment'] is None:  # 新增
                    calc_results['equipment'] = calc_equipment_per_stu(ws, head_size, max_h, merge_map)

                print("全部七项计算已完成。")
            else:
                print("无效输入。")
                continue

            more = input("\n继续其他计算？(y/n)：").strip().lower()
            if more != 'y':
                break

        do_district = input("\n是否进行区县差异系数与汇总统计？(y/n)：").strip().lower()
        stats_rows = None
        summary_rows = None
        if do_district == 'y':
            _, stats_rows, summary_rows = analyze_district_data(ws, head_size, max_h, school_type, merge_map)
            if stats_rows is None:
                print("区县统计失败，可能缺少必要的列。")

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
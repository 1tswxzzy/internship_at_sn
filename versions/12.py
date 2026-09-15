import os
import sys
import re
import numpy as np
from collections import OrderedDict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==================== 工具函数 ====================
def classify_school_type(file_path):
    """根据文件名判断学校类型（小学/中学/初中/高中）"""
    file_name = os.path.basename(file_path)
    if '小学' in file_name:
        return '小学'
    elif '初中' in file_name:
        return '初中'
    elif '高中' in file_name:
        return '高中'
    elif '中学' in file_name:
        return '中学'
    return '未知'

def extract_district(text):
    if not text or str(text).strip() == '':
        return None
    text = str(text).strip()
    pattern = r'([^\s,，、]+(?:区|县|市|自治县|自治旗))'
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    return None

def find_col_by_keywords(ws, head_size, keywords, match_mode='any'):
    """通过关键词匹配列，优先底层表头"""
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
    """根据表头层级路径查找唯一列"""
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
    """返回所有满足前缀层级序列的列号"""
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
    """自动推断表头行数"""
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

# ==================== 查找"是否计校数"列 ====================
def find_count_school_col(ws, head_size):
    return find_col_by_keywords(ws, head_size, '是否计校数')

def build_school_merge_map(ws, head_size):
    """构建不计校数学校的合并映射"""
    count_col = find_count_school_col(ws, head_size)
    if not count_col:
        print("未找到'是否计校数'列，跳过不计校数合并。")
        return {}

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

# ==================== 特殊学校查找与学段分析 ====================
def find_special_school_types(ws, head_size, max_h, file_type):
    """查找特殊办学类型学校"""
    school_type_col = None
    for col in range(1, ws.max_column + 1):
        for row in range(1, head_size + 1):
            cell_val = ws.cell(row=row, column=col).value
            if cell_val and isinstance(cell_val, str) and '办学类型' in cell_val:
                school_type_col = col
                break
        if school_type_col:
            break
    
    if not school_type_col:
        school_type_col = find_col_by_keywords(ws, head_size, '办学类型')
    
    if not school_type_col:
        return None
    
    special_types = {
        '九年一贯制': [],
        '十二年一贯制': [],
        '完全中学': []
    }
    
    for r in range(head_size + 1, max_h + 1):
        type_val = ws.cell(row=r, column=school_type_col).value
        if type_val and isinstance(type_val, str):
            type_str = str(type_val).strip()
            for special_type in special_types.keys():
                if special_type in type_str:
                    school_name = ws.cell(row=r, column=1).value
                    if school_name:
                        school_name_str = str(school_name).strip()
                        stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
                        if not stu_col:
                            stu_col = find_col_by_keywords(ws, head_size, '在校生数')
                        stu_val = ws.cell(row=r, column=stu_col).value if stu_col else None
                        try:
                            stu_count = float(stu_val) if stu_val is not None else 0
                        except (ValueError, TypeError):
                            stu_count = 0
                        
                        special_types[special_type].append({
                            'row': r,
                            'name': school_name_str,
                            'full_type': type_str,
                            'student_count': stu_count,
                            'source_file_type': file_type
                        })
    
    return special_types

def find_school_in_file(file_path, school_name):
    """在指定文件中查找学校，返回其在校生数"""
    wb, ws, max_h, max_l = get_worksheet_info(file_path)
    if wb is None:
        return None
    
    head_size, _ = guess_head_size(ws)
    
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    
    if not stu_col:
        return None
    
    for r in range(head_size + 1, max_h + 1):
        cell_name = ws.cell(row=r, column=1).value
        if cell_name and str(cell_name).strip() == school_name:
            stu_val = ws.cell(row=r, column=stu_col).value
            try:
                return float(stu_val) if stu_val is not None else 0
            except (ValueError, TypeError):
                return 0
    
    for r in range(head_size + 1, max_h + 1):
        cell_name = ws.cell(row=r, column=1).value
        if cell_name:
            name_str = str(cell_name).strip()
            if school_name in name_str or name_str in school_name:
                stu_val = ws.cell(row=r, column=stu_col).value
                try:
                    return float(stu_val) if stu_val is not None else 0
                except (ValueError, TypeError):
                    return 0
    
    return None

def calculate_section_ratios(all_special_schools, file_map):
    """跨文件计算特殊学校的学段占比"""
    weights = {
        '九年一贯制': {'小学': 1, '初中': 1.1},
        '十二年一贯制': {'小学': 1, '初中': 1.1, '高中': 1.32},
        '完全中学': {'初中': 1, '高中': 1.2}
    }
    
    section_file_map = {'小学': '小学', '初中': '初中', '高中': '高中'}
    
    print("\n" + "=" * 60)
    print("第一步：计算特殊学校学段占比")
    print("=" * 60)
    
    section_ratios = {}
    
    for special_type, schools in all_special_schools.items():
        if not schools:
            continue
        
        school_type_weights = weights[special_type]
        
        for school in schools:
            school_name = school['name']
            section_students = {}
            
            for section in school_type_weights.keys():
                file_type = section_file_map.get(section)
                if file_type and file_type in file_map:
                    file_path = file_map[file_type]
                    student_count = find_school_in_file(file_path, school_name)
                    if student_count is not None:
                        section_students[section] = student_count
                    else:
                        section_students[section] = 0
            
            total_weighted = 0
            for section, weight in school_type_weights.items():
                total_weighted += section_students.get(section, 0) * weight
            
            ratios = {}
            for section in school_type_weights.keys():
                if total_weighted > 0:
                    ratios[section] = (section_students.get(section, 0) * school_type_weights[section]) / total_weighted
                else:
                    ratios[section] = 0
            
            section_ratios[school_name] = {
                'type': special_type,
                'ratios': ratios,
                'section_students': section_students,
                'total_weighted': total_weighted
            }
            
            print(f"  {school_name}：")
            for section in school_type_weights.keys():
                print(f"    {section}占比：{ratios[section]*100:.2f}%")
    
    return section_ratios

def get_section_ratio_for_school(school_name, section_ratios, target_section):
    """获取某所学校在目标学段的占比"""
    if school_name in section_ratios:
        ratio_info = section_ratios[school_name]
        return ratio_info['ratios'].get(target_section, 1.0)
    return 1.0

# ==================== 高学历教师列查找 ====================
def find_high_education_cols(ws, head_size, school_type):
    """
    查找高学历教师相关列
    根据不同学校类型使用不同的查找策略
    """
    target_cols = []
    
    if school_type == '小学':
        edu_region = find_cols_with_prefix(ws, head_size, ['专任教师', '#按学历分'])
        if not edu_region:
            for col in range(1, ws.max_column + 1):
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str) and '按学历分' in val:
                        if col not in edu_region:
                            edu_region.append(col)
                        break
        
        keywords = ['硕士研究生', '本科', '专科']
        for kw in keywords:
            found = None
            if edu_region:
                for col in edu_region:
                    for r in range(1, head_size + 1):
                        val = ws.cell(row=r, column=col).value
                        if val and isinstance(val, str) and kw in val:
                            found = col
                            break
                    if found:
                        break
            if not found:
                found = find_col_by_keywords(ws, head_size, kw)
            if found and found not in target_cols:
                target_cols.append(found)
    
    elif school_type == '初中':
        edu_region = []
        for col in range(1, ws.max_column + 1):
            for r in range(1, head_size + 1):
                val = ws.cell(row=r, column=col).value
                if val and isinstance(val, str):
                    if '专任教师' in val and ('初级中学' in val or '初中' in val or '九年一贯制' in val):
                        for c2 in range(col, min(col + 30, ws.max_column + 1)):
                            for r2 in range(r, head_size + 1):
                                val2 = ws.cell(row=r2, column=c2).value
                                if val2 and isinstance(val2, str) and '按学历分' in val2:
                                    if c2 not in edu_region:
                                        edu_region.append(c2)
                                    break
                        break
            if edu_region:
                break
        
        if not edu_region:
            for col in range(1, ws.max_column + 1):
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str) and '按学历分' in val:
                        if col not in edu_region:
                            edu_region.append(col)
                        break
        
        keywords = ['硕士研究生毕业', '硕士', '本科毕业', '本科']
        for kw in keywords:
            found = None
            if edu_region:
                for col in edu_region:
                    for r in range(1, head_size + 1):
                        val = ws.cell(row=r, column=col).value
                        if val and isinstance(val, str) and kw in val:
                            found = col
                            break
                    if found:
                        break
            if not found:
                found = find_col_by_keywords(ws, head_size, kw)
            if found and found not in target_cols:
                target_cols.append(found)
        
        if len(target_cols) > 2:
            filtered = []
            for col in target_cols:
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str):
                        if '研究生' in val and col not in filtered:
                            filtered.append(col)
                        elif '本科' in val and '研究生' not in val and col not in filtered:
                            filtered.append(col)
            if len(filtered) >= 2:
                target_cols = filtered[:2]
    
    elif school_type == '高中':
        edu_region = find_cols_with_prefix(ws, head_size, ['专任教师', '#按学历分'])
        if not edu_region:
            for col in range(1, ws.max_column + 1):
                for r in range(1, head_size + 1):
                    val = ws.cell(row=r, column=col).value
                    if val and isinstance(val, str) and '按学历分' in val:
                        if col not in edu_region:
                            edu_region.append(col)
                        break
        
        keywords = ['硕士研究生', '硕士', '本科']
        for kw in keywords:
            found = None
            if edu_region:
                for col in edu_region:
                    for r in range(1, head_size + 1):
                        val = ws.cell(row=r, column=col).value
                        if val and isinstance(val, str) and kw in val:
                            found = col
                            break
                    if found:
                        break
            if not found:
                found = find_col_by_keywords(ws, head_size, kw)
            if found and found not in target_cols:
                target_cols.append(found)
    
    unique_cols = []
    for col in target_cols:
        if col not in unique_cols:
            unique_cols.append(col)
    
    return unique_cols

# ==================== 计算函数 ====================
def calc_media_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map=None):
    """
    1. 每百名学生多媒体教室数
    调整后指标数 = 原始指标数 × 学段占比
    调整后指标值 = 调整后指标数 × 100 / 在校生数
    """
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    media_col = find_col_by_hierarchy(ws, head_size, ['网络多媒体教室'])
    if not media_col:
        media_col = find_col_by_keywords(ws, head_size, '网络多媒体教室')
    if not stu_col or not media_col:
        return None, {}

    header = ['原始行号', '学校名称', '在校生数', '原始指标数(多媒体教室)', 
              '学段占比', '调整后指标数', '调整后指标值', '备注']
    rows = [header]
    school_index = {}

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
        
        school_name = str(ws.cell(row=r, column=1).value).strip() if ws.cell(row=r, column=1).value else ''
        ratio = get_section_ratio_for_school(school_name, section_ratios, school_type)
        
        adjusted_count = round(media * ratio, 2)
        adjusted_value = round(adjusted_count * 100 / stu, 2)
        
        remark = ''
        if school_name in section_ratios:
            remark = f'{section_ratios[school_name]["type"]}，{school_type}占比{ratio*100:.2f}%'
        
        row_data = [r, school_name, stu, media, round(ratio, 4), adjusted_count, adjusted_value, remark]
        rows.append(row_data)
        school_index[school_name] = len(rows) - 1

    return rows, school_index

def calc_high_education_per_stu(ws, head_size, max_h, school_type, merge_map=None):
    """
    2. 每百名学生高学历教师数 - 不需要调整
    调整后指标数 = 原始指标数
    调整后指标值 = 调整后指标数 × 100 / 在校生数
    """
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    if not stu_col:
        return None, {}

    target_cols = find_high_education_cols(ws, head_size, school_type)
    
    if not target_cols:
        print(f"警告：{school_type}文件中未找到高学历教师相关列")
        return None, {}
    
    print(f"  高学历教师列: {target_cols}")

    header = ['原始行号', '学校名称', '在校生数', '原始指标数(学历教师)', 
              '学段占比', '调整后指标数', '调整后指标值']
    rows = [header]
    school_index = {}

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
        
        school_name = str(ws.cell(row=r, column=1).value).strip() if ws.cell(row=r, column=1).value else ''
        
        adjusted_count = round(edu_sum, 2)
        adjusted_value = round(adjusted_count * 100 / stu, 2)
        
        row_data = [r, school_name, stu, edu_sum, 1.0, adjusted_count, adjusted_value]
        rows.append(row_data)
        school_index[school_name] = len(rows) - 1

    return rows, school_index

def calc_backbone_per_stu(ws, head_size, max_h, school_type, merge_map=None):
    """
    3. 每百名学生骨干教师数 - 不需要调整
    """
    if school_type == '小学':
        backbone_col = find_col_by_hierarchy(ws, head_size, ['教基1102', '县级以上骨干教师(小学)'])
    elif school_type == '初中':
        backbone_col = find_col_by_hierarchy(ws, head_size, ['教基1102', '县级以上骨干教师(初中)'])
    else:
        backbone_col = find_col_by_hierarchy(ws, head_size, ['教基1102', '县级以上骨干教师(高中)'])
    if not backbone_col:
        backbone_col = find_col_by_keywords(ws, head_size, '骨干教师')
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    if not stu_col or not backbone_col:
        return None, {}

    header = ['原始行号', '学校名称', '在校生数', '原始指标数(骨干教师)', 
              '学段占比', '调整后指标数', '调整后指标值']
    rows = [header]
    school_index = {}

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
        
        school_name = str(ws.cell(row=r, column=1).value).strip() if ws.cell(row=r, column=1).value else ''
        
        adjusted_count = round(back, 2)
        adjusted_value = round(adjusted_count * 100 / stu, 2)
        
        row_data = [r, school_name, stu, back, 1.0, adjusted_count, adjusted_value]
        rows.append(row_data)
        school_index[school_name] = len(rows) - 1

    return rows, school_index

def calc_art_sport_per_stu(ws, head_size, max_h, merge_map=None):
    """
    4. 每百名学生艺体教师数 - 不需要调整
    """
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    sport_col = find_col_by_hierarchy(ws, head_size, ['专任教师', '#按授课课程分', '体育与健康'])
    art_col = find_col_by_hierarchy(ws, head_size, ['专任教师', '#按授课课程分', '艺术', '计'])
    if not sport_col and not art_col:
        sport_col = find_col_by_keywords(ws, head_size, '体育与健康')
        art_col = find_col_by_keywords(ws, head_size, '艺术')
    if not stu_col or (not sport_col and not art_col):
        return None, {}

    header = ['原始行号', '学校名称', '在校生数', '原始指标数(艺体教师)', 
              '学段占比', '调整后指标数', '调整后指标值']
    rows = [header]
    school_index = {}

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
        
        school_name = str(ws.cell(row=r, column=1).value).strip() if ws.cell(row=r, column=1).value else ''
        
        adjusted_count = round(total, 2)
        adjusted_value = round(adjusted_count * 100 / stu, 2)
        
        row_data = [r, school_name, stu, total, 1.0, adjusted_count, adjusted_value]
        rows.append(row_data)
        school_index[school_name] = len(rows) - 1

    return rows, school_index

def calc_teaching_area_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map=None):
    """
    5. 生均教学及辅助用房面积
    调整后指标数 = 净面积 × 学段占比
    调整后指标值 = 调整后指标数 / 在校生数
    """
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    
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
    
    if not teach_row:
        return None, {}
    
    teach_area_col = None
    for r in range(teach_row + 1, head_size + 1):
        val = ws.cell(row=r, column=teach_base_col).value
        if val and isinstance(val, str) and val.strip() == '计':
            teach_area_col = teach_base_col
            break
    
    if not teach_area_col:
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
    
    if not teach_area_col:
        return None, {}
    
    indoor_sport_col = find_col_by_keywords(ws, head_size, '室内体育用房')
    
    header = ['原始行号', '学校名称', '在校生数', '原始指标数(净面积)', 
              '学段占比', '调整后指标数', '调整后指标值', '备注']
    rows = [header]
    school_index = {}
    
    for r in range(head_size + 1, max_h + 1):
        school_name = str(ws.cell(row=r, column=1).value).strip() if ws.cell(row=r, column=1).value else ''
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
        
        net_area = max(teach - indoor, 0)
        ratio = get_section_ratio_for_school(school_name, section_ratios, school_type)
        
        adjusted_count = round(net_area * ratio, 2)
        adjusted_value = round(adjusted_count / stu, 2)
        
        remark = ''
        if school_name in section_ratios:
            remark = f'{section_ratios[school_name]["type"]}，{school_type}占比{ratio*100:.2f}%'
        
        row_data = [r, school_name, stu, net_area, round(ratio, 4), adjusted_count, adjusted_value, remark]
        rows.append(row_data)
        school_index[school_name] = len(rows) - 1
    
    return rows, school_index

def calc_sports_area_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map=None):
    """
    6. 生均体育运动场馆面积
    调整后指标数 = 总面积 × 学段占比
    调整后指标值 = 调整后指标数 / 在校生数
    """
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    sports_field_col = find_col_by_keywords(ws, head_size, '运动场地面积')
    if not sports_field_col:
        sports_field_col = find_col_by_keywords(ws, head_size, '运动场地')
    indoor_sport_col = find_col_by_keywords(ws, head_size, '室内体育用房')
    
    if not stu_col or not sports_field_col:
        return None, {}
    
    header = ['原始行号', '学校名称', '在校生数', '原始指标数(总面积)', 
              '学段占比', '调整后指标数', '调整后指标值', '备注']
    rows = [header]
    school_index = {}
    
    for r in range(head_size + 1, max_h + 1):
        school_name = str(ws.cell(row=r, column=1).value).strip() if ws.cell(row=r, column=1).value else ''
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
        
        total_area = sports + indoor
        ratio = get_section_ratio_for_school(school_name, section_ratios, school_type)
        
        adjusted_count = round(total_area * ratio, 2)
        adjusted_value = round(adjusted_count / stu, 2)
        
        remark = ''
        if school_name in section_ratios:
            remark = f'{section_ratios[school_name]["type"]}，{school_type}占比{ratio*100:.2f}%'
        
        row_data = [r, school_name, stu, total_area, round(ratio, 4), adjusted_count, adjusted_value, remark]
        rows.append(row_data)
        school_index[school_name] = len(rows) - 1
    
    return rows, school_index

def calc_equipment_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map=None):
    """
    7. 生均教学仪器设备值
    调整后指标数 = 设备值(万元) × 学段占比
    调整后指标值 = 调整后指标数 × 10000 / 在校生数
    """
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备资产值(万元)')
    if not equipment_col:
        equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备资产值')
    if not equipment_col:
        equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备')
    
    if not stu_col or not equipment_col:
        return None, {}
    
    header = ['原始行号', '学校名称', '在校生数', '原始指标数(设备值万元)', 
              '学段占比', '调整后指标数', '调整后指标值', '备注']
    rows = [header]
    school_index = {}
    
    for r in range(head_size + 1, max_h + 1):
        school_name = str(ws.cell(row=r, column=1).value).strip() if ws.cell(row=r, column=1).value else ''
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
        
        ratio = get_section_ratio_for_school(school_name, section_ratios, school_type)
        
        adjusted_count = round(equipment * ratio, 4)
        adjusted_value = round(adjusted_count * 10000 / stu, 2)
        
        remark = ''
        if school_name in section_ratios:
            remark = f'{section_ratios[school_name]["type"]}，{school_type}占比{ratio*100:.2f}%'
        
        row_data = [r, school_name, stu, equipment, round(ratio, 4), adjusted_count, adjusted_value, remark]
        rows.append(row_data)
        school_index[school_name] = len(rows) - 1
    
    return rows, school_index

# ==================== 区县差异系数与汇总统计 ====================
def analyze_district_data(ws, head_size, max_h, school_type, section_ratios, merge_map=None):
    """
    区县差异系数与汇总统计
    
    X = Σ(调整后指标数 × multiplier) / Σ(在校生数)
    方差 = Σ(A_i/A_sum) × (调整后指标值_i - X)²
    """
    district_col = find_col_by_keywords(ws, head_size, '统计三级')
    if not district_col:
        district_col = find_col_by_hierarchy(ws, head_size, ['统计三级'])
    if not district_col:
        print("警告：未找到'统计三级'列，跳过区县统计。")
        return None, None
    
    print(f"\n开始区县差异系数分析...")
    
    # 预先查找所有需要的列
    stu_col = find_col_by_hierarchy(ws, head_size, ['在校生数', '合计'])
    if not stu_col:
        stu_col = find_col_by_keywords(ws, head_size, '在校生数')
    
    media_col = find_col_by_hierarchy(ws, head_size, ['网络多媒体教室'])
    if not media_col:
        media_col = find_col_by_keywords(ws, head_size, '网络多媒体教室')
    
    edu_cols = find_high_education_cols(ws, head_size, school_type)
    
    if school_type == '小学':
        backbone_col = find_col_by_hierarchy(ws, head_size, ['教基1102', '县级以上骨干教师(小学)'])
    elif school_type == '初中':
        backbone_col = find_col_by_hierarchy(ws, head_size, ['教基1102', '县级以上骨干教师(初中)'])
    else:
        backbone_col = find_col_by_hierarchy(ws, head_size, ['教基1102', '县级以上骨干教师(高中)'])
    if not backbone_col:
        backbone_col = find_col_by_keywords(ws, head_size, '骨干教师')
    
    sport_col = find_col_by_hierarchy(ws, head_size, ['专任教师', '#按授课课程分', '体育与健康'])
    if not sport_col:
        sport_col = find_col_by_keywords(ws, head_size, '体育与健康')
    art_col = find_col_by_hierarchy(ws, head_size, ['专任教师', '#按授课课程分', '艺术', '计'])
    if not art_col:
        art_col = find_col_by_keywords(ws, head_size, '艺术')
    
    teach_row = None
    teach_base_col = None
    for rr in range(1, head_size + 1):
        for cc in range(1, ws.max_column + 1):
            val = ws.cell(row=rr, column=cc).value
            if val and isinstance(val, str) and '教学及辅助用房' in val:
                teach_row = rr
                teach_base_col = cc
                break
        if teach_row:
            break
    
    teach_area_col = None
    if teach_base_col:
        for rr in range(teach_row + 1, head_size + 1):
            val = ws.cell(row=rr, column=teach_base_col).value
            if val and isinstance(val, str) and val.strip() == '计':
                teach_area_col = teach_base_col
                break
        if not teach_area_col:
            for cc in range(max(1, teach_base_col - 2), min(ws.max_column, teach_base_col + 3)):
                for rr in range(teach_row + 1, head_size + 1):
                    val = ws.cell(row=rr, column=cc).value
                    if val and isinstance(val, str) and val.strip() == '计':
                        teach_area_col = cc
                        break
                if teach_area_col:
                    break
    
    indoor_sport_col_teach = find_col_by_keywords(ws, head_size, '室内体育用房')
    
    sports_field_col = find_col_by_keywords(ws, head_size, '运动场地面积')
    if not sports_field_col:
        sports_field_col = find_col_by_keywords(ws, head_size, '运动场地')
    indoor_sport_col_sports = find_col_by_keywords(ws, head_size, '室内体育用房')
    
    equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备资产值(万元)')
    if not equipment_col:
        equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备资产值')
    if not equipment_col:
        equipment_col = find_col_by_keywords(ws, head_size, '教学仪器设备')
    
    print(f"  在校生数列: {stu_col}")
    print(f"  多媒体教室列: {media_col}")
    print(f"  学历教师列: {edu_cols}")
    print(f"  骨干教师列: {backbone_col}")
    print(f"  体育教师列: {sport_col}, 艺术教师列: {art_col}")
    print(f"  教学用房列: {teach_area_col}")
    print(f"  运动场地列: {sports_field_col}")
    print(f"  仪器设备列: {equipment_col}")
    
    def safe_float(val, default=0.0):
        if val is None:
            return default
        try:
            return float(val)
        except (ValueError, TypeError):
            return default
    
    school_data = OrderedDict()
    
    for r in range(head_size + 1, max_h + 1):
        district_cell = ws.cell(row=r, column=district_col).value
        if not district_cell:
            continue
        district = str(district_cell).strip()
        if district in ('合计', '总计', '小计', '备注'):
            continue
        
        school_name = str(ws.cell(row=r, column=1).value).strip() if ws.cell(row=r, column=1).value else ''
        if not school_name:
            continue
        
        stu = safe_float(ws.cell(row=r, column=stu_col).value, 0) if stu_col else 0
        if stu == 0:
            continue
        
        ratio = get_section_ratio_for_school(school_name, section_ratios, school_type)
        
        raw_data = {'school': school_name, 'A': stu, 'ratio': ratio}
        
        # ind1：多媒体教室数
        ind1_raw = safe_float(ws.cell(row=r, column=media_col).value, 0) if media_col else 0
        raw_data['ind1_raw'] = ind1_raw
        raw_data['ind1_adj_count'] = round(ind1_raw * ratio, 2)
        raw_data['ind1_adj_value'] = round(ind1_raw * ratio * 100 / stu, 2)
        
        # ind2：学历教师总数
        edu_sum = 0
        for col in edu_cols:
            edu_sum += safe_float(ws.cell(row=r, column=col).value, 0)
        raw_data['ind2_raw'] = edu_sum
        raw_data['ind2_adj_count'] = round(edu_sum, 2)
        raw_data['ind2_adj_value'] = round(edu_sum * 100 / stu, 2) if stu > 0 else 0
        
        # ind3：骨干教师数
        ind3_raw = safe_float(ws.cell(row=r, column=backbone_col).value, 0) if backbone_col else 0
        raw_data['ind3_raw'] = ind3_raw
        raw_data['ind3_adj_count'] = round(ind3_raw, 2)
        raw_data['ind3_adj_value'] = round(ind3_raw * 100 / stu, 2) if stu > 0 else 0
        
        # ind4：艺体教师总数
        sport_val = safe_float(ws.cell(row=r, column=sport_col).value, 0) if sport_col else 0
        art_val = safe_float(ws.cell(row=r, column=art_col).value, 0) if art_col else 0
        ind4_raw = sport_val + art_val
        raw_data['ind4_raw'] = ind4_raw
        raw_data['ind4_adj_count'] = round(ind4_raw, 2)
        raw_data['ind4_adj_value'] = round(ind4_raw * 100 / stu, 2) if stu > 0 else 0
        
        # ind5：教学用房净面积
        net_area = 0
        if teach_area_col:
            teach_val = safe_float(ws.cell(row=r, column=teach_area_col).value, 0)
            indoor_val = safe_float(ws.cell(row=r, column=indoor_sport_col_teach).value, 0) if indoor_sport_col_teach else 0
            net_area = max(teach_val - indoor_val, 0)
        raw_data['ind5_raw'] = net_area
        raw_data['ind5_adj_count'] = round(net_area * ratio, 2)
        raw_data['ind5_adj_value'] = round(net_area * ratio / stu, 2) if stu > 0 else 0
        
        # ind6：运动场馆总面积
        total_sp = 0
        if sports_field_col:
            sports_val = safe_float(ws.cell(row=r, column=sports_field_col).value, 0)
            indoor_sp_val = safe_float(ws.cell(row=r, column=indoor_sport_col_sports).value, 0) if indoor_sport_col_sports else 0
            total_sp = sports_val + indoor_sp_val
        raw_data['ind6_raw'] = total_sp
        raw_data['ind6_adj_count'] = round(total_sp * ratio, 2)
        raw_data['ind6_adj_value'] = round(total_sp * ratio / stu, 2) if stu > 0 else 0
        
        # ind7：设备值（万元）- 调整后指标值单位：元/生
        eq_val = 0
        if equipment_col:
            eq_val = safe_float(ws.cell(row=r, column=equipment_col).value, 0)
        raw_data['ind7_raw'] = eq_val *10000
        raw_data['ind7_adj_count'] = round(eq_val * ratio * 10000, 4)
        raw_data['ind7_adj_value'] = round(eq_val * ratio * 10000 / stu, 2) if stu > 0 else 0
        
        if district not in school_data:
            school_data[district] = []
        school_data[district].append(raw_data)
    
    if not school_data:
        print("未提取到区县数据。")
        return None, None
    
    # 打印样本数据
    all_data_debug = []
    for data in school_data.values():
        all_data_debug.extend(data)
    if all_data_debug:
        sample = all_data_debug[0]
        print(f"\n  样本数据（{sample['school']}）：")
        print(f"    在校生数: {sample['A']}")
        print(f"    ind1(多媒体): 原始={sample['ind1_raw']}, 调整后指标数={sample['ind1_adj_count']}, 调整后指标值={sample['ind1_adj_value']}")
        print(f"    ind2(学历教师): 原始={sample['ind2_raw']}, 调整后指标数={sample['ind2_adj_count']}, 调整后指标值={sample['ind2_adj_value']}")
        print(f"    ind5(教学用房): 原始={sample['ind5_raw']}, 调整后指标数={sample['ind5_adj_count']}, 调整后指标值={sample['ind5_adj_value']}")
        print(f"    ind7(仪器设备万元): 原始={sample['ind7_raw']}, 调整后指标数={sample['ind7_adj_count']}, 调整后指标值(元/生)={sample['ind7_adj_value']}")
    
    indicator_defs = [
        ('多媒体教室', 'ind1', 100),
        ('高学历教师', 'ind2', 100),
        ('骨干教师', 'ind3', 100),
        ('艺体教师', 'ind4', 100),
        ('教学用房面积', 'ind5', 1),
        ('运动场馆面积', 'ind6', 1),
        ('教学仪器设备', 'ind7', 1)
    ]
    
    def calc_cv(data_list, ind_key, multiplier):
        A_sum = sum(d['A'] for d in data_list)
        if A_sum == 0:
            return {'X': 0, 'S': 0, 'CV': 0}
        
        adj_count_key = f'{ind_key}_adj_count'
        adj_value_key = f'{ind_key}_adj_value'
        
        total_adj_count = sum(d[adj_count_key] for d in data_list)
        if total_adj_count == 0:
            return {'X': 0, 'S': 0, 'CV': 0}
        
        weighted_sum = sum(d[adj_count_key] * multiplier for d in data_list)
        X = weighted_sum / A_sum
        
        variance = 0
        for d in data_list:
            if d['A'] > 0:
                variance += (d['A'] / A_sum) * (d[adj_value_key] - X)**2
        
        S = np.sqrt(variance) if variance > 0 else 0
        CV = round(S / X, 2) if X != 0 else 0
        return {'X': round(X, 4), 'S': round(S, 6), 'CV': CV}
    
    # 差异系数统计表
    stats_header = ['组别', '区县']
    for name, _, _ in indicator_defs:
        stats_header.extend([f'{name}平均值', f'{name}标准差', f'{name}差异系数'])
    stats_rows = [stats_header]
    
    group_counter = 1
    for district, data in school_data.items():
        row = [f'T{group_counter}', district]
        for name, ind_key, multiplier in indicator_defs:
            stats = calc_cv(data, ind_key, multiplier)
            row.extend([stats['X'], stats['S'], stats['CV']])
        stats_rows.append(row)
        group_counter += 1
    
    all_data = []
    for data in school_data.values():
        all_data.extend(data)
    total_row = ['', '总计']
    for name, ind_key, multiplier in indicator_defs:
        stats = calc_cv(all_data, ind_key, multiplier)
        total_row.extend([stats['X'], stats['S'], stats['CV']])
    stats_rows.append(total_row)
    
    # 汇总统计表
    summary_header = ['区县', '学生总数']
    for name, ind_key, multiplier in indicator_defs:
        summary_header.extend([f'{name}原始和', f'{name}调整后指标数和', f'{name}调整后指标值合计'])
    summary_rows = [summary_header]
    
    for district, data in school_data.items():
        sA = sum(d['A'] for d in data)
        row = [district, sA]
        for name, ind_key, multiplier in indicator_defs:
            raw_sum = sum(d[f'{ind_key}_raw'] for d in data)
            adj_count_sum = sum(d[f'{ind_key}_adj_count'] for d in data)
            adj_value_sum = sum(d[f'{ind_key}_adj_value'] for d in data)
            row.extend([round(raw_sum, 2), round(adj_count_sum, 2), round(adj_value_sum, 2)])
        summary_rows.append(row)
    
    total_A = sum(d['A'] for d in all_data)
    total_row_summary = ['总计', total_A]
    for name, ind_key, multiplier in indicator_defs:
        raw_sum = sum(d[f'{ind_key}_raw'] for d in all_data)
        adj_count_sum = sum(d[f'{ind_key}_adj_count'] for d in all_data)
        adj_value_sum = sum(d[f'{ind_key}_adj_value'] for d in all_data)
        total_row_summary.extend([round(raw_sum, 2), round(adj_count_sum, 2), round(adj_value_sum, 2)])
    summary_rows.append(total_row_summary)
    
    return stats_rows, summary_rows

# ==================== 保存结果 ====================
def save_results_to_file(original_path, school_type, calc_results, stats_rows, summary_rows, section_ratios, all_special_schools):
    """保存分析结果"""
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
        '区县汇总统计'
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

    special_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    special_font = Font(color="FF0000", bold=True)

    for sname, block in zip(sheet_names, data_blocks):
        ws_out = out_wb.create_sheet(title=sname)
        if block:
            for row_idx, row in enumerate(block, 1):
                ws_out.append(row)
                if len(row) > 1 and row[1] and str(row[1]) in section_ratios:
                    for col in range(1, len(row) + 1):
                        cell = ws_out.cell(row=row_idx, column=col)
                        cell.fill = special_fill
                        cell.font = special_font
        else:
            ws_out.append(['无数据'])
    
    if section_ratios:
        add_section_ratio_sheet(out_wb, section_ratios, all_special_schools)
    
    out_wb.save(new_path)
    print(f"\n分析结果已保存至：{new_path}")
    return new_path

def add_section_ratio_sheet(out_wb, section_ratios, all_special_schools):
    """添加学段占比分析sheet"""
    ws = out_wb.create_sheet(title='学段占比分析')
    
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    ws.merge_cells('A1:L1')
    ws['A1'].value = '特殊学校学段占比分析'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')
    ws.append([])
    ws.append(['数据列结构：在校生数 | 原始指标数 | 学段占比 | 调整后指标数 | 调整后指标值'])
    ws.append(['调整后指标数 = 原始指标数 × 学段占比'])
    ws.append(['指标1-4（每百名类）：调整后指标值 = 调整后指标数 × 100 / 在校生数'])
    ws.append(['指标5-7（生均类）：调整后指标值 = 调整后指标数 / 在校生数（设备值×10000/在校生数）'])
    ws.append(['X = Σ(调整后指标数 × multiplier) / Σ(在校生数)'])
    ws.append(['方差 = Σ(A_i/A_sum) × (调整后指标值_i - X)²'])
    ws.append([])
    
    current_row = 10
    
    for special_type in ['九年一贯制', '十二年一贯制', '完全中学']:
        schools = all_special_schools.get(special_type, [])
        if not schools:
            continue
        
        weight_desc = {
            '九年一贯制': '权重：小学:初中 = 1:1.1',
            '十二年一贯制': '权重：小学:初中:高中 = 1:1.1:1.32',
            '完全中学': '权重：初中:高中 = 1:1.2'
        }
        
        ws.merge_cells(f'A{current_row}:L{current_row}')
        ws[f'A{current_row}'].value = f'【{special_type}】{weight_desc[special_type]}'
        ws[f'A{current_row}'].font = Font(bold=True, size=12, color="003366")
        ws[f'A{current_row}'].fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
        current_row += 1
        
        ratio_sections = []
        if special_type == '九年一贯制':
            ratio_sections = ['小学', '初中']
        elif special_type == '十二年一贯制':
            ratio_sections = ['小学', '初中', '高中']
        else:
            ratio_sections = ['初中', '高中']
        
        headers = ['学校名称']
        for section in ratio_sections:
            headers.append(f'{section}学生数')
        for section in ratio_sections:
            headers.append(f'{section}权重')
        for section in ratio_sections:
            headers.append(f'{section}加权值')
        headers.append('加权总值')
        for section in ratio_sections:
            headers.append(f'{section}占比(%)')
        
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border
        current_row += 1
        
        weights_map = {
            '九年一贯制': {'小学': 1, '初中': 1.1},
            '十二年一贯制': {'小学': 1, '初中': 1.1, '高中': 1.32},
            '完全中学': {'初中': 1, '高中': 1.2}
        }
        school_weights = weights_map.get(special_type, {})
        
        for school in schools:
            school_name = school['name']
            ratio_info = section_ratios.get(school_name, {})
            ratios = ratio_info.get('ratios', {})
            section_students = ratio_info.get('section_students', {})
            
            row_data = [school_name]
            for section in ratio_sections:
                row_data.append(section_students.get(section, 0))
            for section in ratio_sections:
                row_data.append(school_weights.get(section, 0))
            for section in ratio_sections:
                row_data.append(round(section_students.get(section, 0) * school_weights.get(section, 0), 2))
            row_data.append(round(ratio_info.get('total_weighted', 0), 2))
            for section in ratio_sections:
                row_data.append(round(ratios.get(section, 0) * 100, 2))
            
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=value)
                cell.border = thin_border
            current_row += 1
        
        current_row += 1
    
    for col in range(1, 14):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 14

# ==================== 文件收集与处理 ====================
def collect_files():
    """收集小学、初中、高中的文件"""
    required_types = ['小学', '初中', '高中']
    file_map = {}
    
    print("=" * 60)
    print("请依次提供小学、初中、高中的 Excel 文件")
    print("=" * 60)
    
    for req_type in required_types:
        while True:
            file_path = input(f"\n请输入{req_type}的 Excel 文件路径：").strip().strip('"').strip("'")
            if file_path.lower() == 'q':
                sys.exit(0)
            if not os.path.exists(file_path):
                print(f"错误：文件不存在 - {file_path}")
                continue
            file_map[req_type] = file_path
            print(f"已添加{req_type}文件：{os.path.basename(file_path)}")
            break
    
    return file_map

def process_single_file(file_path, school_type, section_ratios):
    """处理单个文件"""
    wb, ws, max_h, max_l = get_worksheet_info(file_path)
    if wb is None:
        return None, None, None
    
    print(f"\n{'='*60}")
    print(f"处理文件：{os.path.basename(file_path)}  类型：{school_type}")
    
    head_size, _ = guess_head_size(ws)
    print(f"自动检测表头行数：{head_size}")
    
    user_input = input("确认表头行数（回车确认 / 输入正确行数）：").strip()
    if user_input.isdigit():
        head_size = int(user_input)
    
    merge_map = build_school_merge_map(ws, head_size)
    
    calc_results = {
        'media': None, 'high_edu': None, 'backbone': None,
        'art_sport': None, 'teaching_area': None, 'sports_area': None,
        'equipment': None,
    }
    
    while True:
        print("\n请选择计算项目：")
        print("  1. 每百名学生多媒体教室数（需要调整）")
        print("  2. 每百名学生高学历教师数")
        print("  3. 每百名学生骨干教师数")
        print("  4. 每百名学生艺体教师数")
        print("  5. 生均教学及辅助用房面积（需要调整）")
        print("  6. 生均运动场馆面积（需要调整）")
        print("  7. 生均教学仪器设备值（需要调整）")
        print("  8. 一键执行以上全部")
        print("  9. 区县差异系数与汇总统计")
        print("  0. 完成此文件，生成报告")
        choice = input("请输入选项：").strip()
        
        if choice == '0':
            break
        elif choice == '8':
            calc_results['media'], _ = calc_media_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map)
            calc_results['high_edu'], _ = calc_high_education_per_stu(ws, head_size, max_h, school_type, merge_map)
            calc_results['backbone'], _ = calc_backbone_per_stu(ws, head_size, max_h, school_type, merge_map)
            calc_results['art_sport'], _ = calc_art_sport_per_stu(ws, head_size, max_h, merge_map)
            calc_results['teaching_area'], _ = calc_teaching_area_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map)
            calc_results['sports_area'], _ = calc_sports_area_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map)
            calc_results['equipment'], _ = calc_equipment_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map)
            print("全部计算完成。")
        elif choice == '9':
            stats_rows, summary_rows = analyze_district_data(ws, head_size, max_h, school_type, section_ratios, merge_map)
            return calc_results, stats_rows, summary_rows
        elif choice == '1':
            calc_results['media'], _ = calc_media_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map)
        elif choice == '2':
            calc_results['high_edu'], _ = calc_high_education_per_stu(ws, head_size, max_h, school_type, merge_map)
        elif choice == '3':
            calc_results['backbone'], _ = calc_backbone_per_stu(ws, head_size, max_h, school_type, merge_map)
        elif choice == '4':
            calc_results['art_sport'], _ = calc_art_sport_per_stu(ws, head_size, max_h, merge_map)
        elif choice == '5':
            calc_results['teaching_area'], _ = calc_teaching_area_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map)
        elif choice == '6':
            calc_results['sports_area'], _ = calc_sports_area_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map)
        elif choice == '7':
            calc_results['equipment'], _ = calc_equipment_per_stu_with_ratio(ws, head_size, max_h, school_type, section_ratios, merge_map)
        else:
            print("无效输入。")
    
    stats_rows = None
    summary_rows = None
    do_district = input("\n是否进行区县差异系数与汇总统计？(y/n)：").strip().lower()
    if do_district == 'y':
        stats_rows, summary_rows = analyze_district_data(ws, head_size, max_h, school_type, section_ratios, merge_map)
    
    return calc_results, stats_rows, summary_rows

def main():
    print("=" * 60)
    print("Excel 教育数据综合分析工具")
    print("支持特殊学校学段占比与区县差异系数分析")
    print("=" * 60)
    print("\n数据列结构：在校生数 | 原始指标数 | 学段占比 | 调整后指标数 | 调整后指标值")
    print("调整后指标数 = 原始指标数 × 学段占比")
    print("指标1-4（每百名类）：调整后指标值 = 调整后指标数 × 100 / 在校生数")
    print("指标5-7（生均类）：调整后指标值 = 调整后指标数 / 在校生数（设备值×10000/在校生数）")
    print("X = Σ(调整后指标数 × multiplier) / Σ(在校生数)")
    print("方差 = Σ(A_i/A_sum) × (调整后指标值_i - X)²")
    
    file_map = collect_files()
    
    # 第一步：查找所有特殊学校
    print("\n" + "=" * 60)
    print("第一步：查找特殊学校类型")
    print("=" * 60)
    
    all_special_schools = {
        '九年一贯制': [],
        '十二年一贯制': [],
        '完全中学': []
    }
    
    for school_type in ['小学', '初中', '高中']:
        file_path = file_map[school_type]
        wb, ws, max_h, _ = get_worksheet_info(file_path)
        if wb is None:
            continue
        head_size, _ = guess_head_size(ws)
        special_schools = find_special_school_types(ws, head_size, max_h, school_type)
        if special_schools:
            for stype in all_special_schools:
                if stype in special_schools:
                    existing_names = {s['name'] for s in all_special_schools[stype]}
                    for school in special_schools[stype]:
                        if school['name'] not in existing_names:
                            all_special_schools[stype].append(school)
    
    for stype in all_special_schools:
        seen = set()
        unique = []
        for school in all_special_schools[stype]:
            if school['name'] not in seen:
                seen.add(school['name'])
                unique.append(school)
        all_special_schools[stype] = unique
    
    total = sum(len(schools) for schools in all_special_schools.values())
    print(f"\n共找到 {total} 所特殊学校")
    
    # 第二步：计算学段占比
    section_ratios = calculate_section_ratios(all_special_schools, file_map)
    
    # 第三步：处理每个文件
    print("\n" + "=" * 60)
    print("第三步：处理各文件")
    print("=" * 60)
    
    for school_type in ['小学', '初中', '高中']:
        file_path = file_map[school_type]
        calc_results, stats_rows, summary_rows = process_single_file(file_path, school_type, section_ratios)
        
        if calc_results and any(v is not None for v in calc_results.values()):
            save_results_to_file(file_path, school_type, calc_results, stats_rows, summary_rows, section_ratios, all_special_schools)
    
    print("\n" + "=" * 60)
    print("处理完成！")
    print("=" * 60)
    
    if section_ratios:
        print("\n特殊学校学段占比汇总：")
        for school_name, info in section_ratios.items():
            ratios = info['ratios']
            print(f"  {school_name} ({info['type']})：")
            for section, ratio in ratios.items():
                print(f"    {section}占比：{ratio*100:.2f}%")
    
    input("\n按 Enter 键退出...")

if __name__ == "__main__":
    main()
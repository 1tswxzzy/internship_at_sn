import os
import sys
import openpyxl
import pandas as pd

def get_worksheet_info(file_path):
    """
        参数:           file_path: Excel 文件路径
        返回:
            ws:         openpyxl 工作表对象
            max_h:      有效总行数（从第1行到最后有数据的行）
            max_l:      有效总列数（从第1列到最后有数据的列）
    """
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
    except FileNotFoundError:
        print(f"错误：找不到文件 '{file_path}'，请检查路径是否正确。")
        return None, None, None, None
    except Exception as e:
        print(f"读取文件时发生错误：{e}")
        return None, None, None, None

    # 计算有效行数
    max_h = ws.max_row
    # 从最大行向上扫描，找到第一个至少有一个非空单元格的行
    while max_h > 0:
        row_vals = [cell.value for cell in ws[max_h]]
        if any(v is not None for v in row_vals):
            break
        max_h -= 1

    # 计算有效列数
    max_l = ws.max_column
    while max_l > 0:
        col_vals = [ws.cell(row=r, column=max_l).value for r in range(1, max_h + 1)]
        if any(v is not None for v in col_vals):
            break
        max_l -= 1

    return wb, ws, max_h, max_l

def med_clas_per_hun_stu(wb, ws, head_size, max_h):
    """
    计算每百名学生拥有多媒体教室数，结果写入新 sheet。
    新 sheet 名称：'每百名学生多媒体教室数'
    参数：
        wb: 工作簿对象（用于添加新 sheet）
        ws: 当前工作表对象
        head_size: 表头行数
        max_h: 有效数据最大行号
    """
    # 1. 在表头中定位所需列
    stu_col = None      # 在校生数（合计）
    media_col = None    # 网络多媒体教室

    for row in ws.iter_rows(min_row=1, max_row=head_size, values_only=False):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if '在校生数' in cell.value:
                    stu_col = cell.column
                if '网络多媒体教室' in cell.value:
                    media_col = cell.column

    if stu_col is None or media_col is None:
        print("错误：未能在表头中找到“在校生数”或“网络多媒体教室”列，计算终止。")
        return

    print(f"已定位：在校生数 -> 第{stu_col}列，网络多媒体教室 -> 第{media_col}列")

    # 2. 创建或获取新 sheet
    sheet_name = '每百名学生多媒体教室数'
    if sheet_name in wb.sheetnames:
        # 如果已存在，删除后重建（可根据需要改为询问用户）
        del wb[sheet_name]
    new_ws = wb.create_sheet(sheet_name)

    # 3. 写入标题行
    new_ws.append(['原始行号', '学校名称（A列）', '在校生数', '网络多媒体教室', '每百名学生多媒体教室数'])

    # 4. 逐行计算并写入结果
    count = 0
    for r in range(head_size + 1, max_h + 1):
        # 读取数值
        stu_val = ws.cell(row=r, column=stu_col).value
        media_val = ws.cell(row=r, column=media_col).value

        # 跳过无效数据
        try:
            stu = float(stu_val) if stu_val is not None else None
            media = float(media_val) if media_val is not None else None
        except (ValueError, TypeError):
            continue

        if stu is None or media is None or stu == 0:
            continue

        # 计算公式：每百名学生多媒体教室数 = 多媒体教室数 * 100 / 学生总数
        result = round(media * 100 / stu, 2)

        # 获取学校名称（假设 A 列为名称）
        school_name = ws.cell(row=r, column=1).value

        new_ws.append([r, school_name, stu, media, result])
        count += 1

    print(f"计算完成，共处理 {count} 行数据，结果已写入 sheet“{sheet_name}”。")

def guess_head_size(ws, max_rows_to_check=20):
    """
    自动推断表头行数，并返回预览内容。
    参数:
        ws:                 openpyxl 工作表对象
        max_rows_to_check:  向前扫描的最大行数（避免全表扫描）
    返回:
        head_size:  推断出的表头行数（int）
        preview:    表头区域的二维列表（每行是一个列表）
    """
    head_size = 1   # 默认值

    # 方法1：利用合并单元格
    if ws.merged_cells.ranges:
        head_size = 0
        for merge_range in ws.merged_cells.ranges:
            if merge_range.max_row > head_size:
                head_size = merge_range.max_row
        if head_size > 0:
            # 确保不超过有效行数
            head_size = min(head_size, ws.max_row)
            preview = [[cell.value if cell.value is not None else "" for cell in ws[row]]
                        for row in range(1, head_size + 1)]
            return head_size, preview

    # 方法2：扫描前 max_rows_to_check 行，找第一个包含数字的行作为数据起始行
    data_start = None
    for row in ws.iter_rows(min_row=1, max_row=min(max_rows_to_check, ws.max_row),
                            values_only=True):
        # 检查该行是否有任一单元格可视为数值
        has_number = False
        for cell in row:
            if cell is None:
                continue
            if isinstance(cell, (int, float)):
                has_number = True
                break
            if isinstance(cell, str):
                # 尝试解析数字（去掉百分号、千位分隔符等）
                try:
                    float(cell.replace(',', '').replace('%', '').strip())
                    has_number = True
                    break
                except ValueError:
                    continue
        if has_number:
            data_start = row  # 这一行就是数据首行
            break

    if data_start:
        # 找到数据首行在 sheet 中的行号
        for idx, row in enumerate(ws.iter_rows(min_row=1,
                                               max_row=min(max_rows_to_check, ws.max_row)),
                                  start=1):
            if row == data_start:
                head_size = idx - 1
                break
        if head_size < 1:
            head_size = 1
    else:
        head_size = 1   # 未检测到数字，保守设为1

    # 生成预览
    preview = [[cell.value if cell.value is not None else "" for cell in ws[row]]
           for row in range(1, head_size + 1)]
    return head_size, preview


def main():
    print("=" * 50)
    print("Excel 文件信息分析工具（含表头自动检测）")
    print("=" * 50)

    while True:
        # ==================== 第一步：获取文件路径 ====================
        file_path = input("\n请输入 Excel 文件的完整路径（或拖拽文件到此处）：").strip()
        file_path = file_path.strip('"').strip("'")

        if not os.path.exists(file_path):
            print(f"文件路径 '{file_path}' 不存在，请重新输入。")
            continue

        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in ['.xlsx', '.xlsm', '.xlsb']:
            print(f"警告：文件扩展名为 '{file_ext}'，可能不是 Excel 文件，尝试读取...")

        print(f"\n正在分析文件：{file_path}")

        # ==================== 第二步：获取工作表基本信息 ====================
        wb, ws, max_h, max_l = get_worksheet_info(file_path)   # 接收 wb
        if wb is None:
            print("文件分析失败，请检查文件是否损坏或格式是否正确。")
            choice = input("\n是否继续分析其他文件？(y/n)：").strip().lower()
            if choice != 'y':
                break
            continue

        print(f"有效总行数 (max_h)：{max_h}")
        print(f"有效总列数 (max_l)：{max_l}")

        # ==================== 第三步：自动推断表头行数 ====================
        head_size, preview = guess_head_size(ws)
        print("\n自动检测表头行数：", head_size)
        print("表头预览：")
        for i, line in enumerate(preview, 1):
            print(f"  第{i}行: {line}")

        user_input = input(
            "\n确认表头行数（直接回车表示确认，或输入正确行数;建议自行打开工作表单查看表头行数！）："
        ).strip()
        if user_input.isdigit():
            head_size = int(user_input)

        max_size = max_h - head_size
        if max_size < 0:
            max_size = 0
        print(f"\n数据区域起始行：{head_size + 1}")
        print(f"可处理的数据行数 (max_size)：{max_size}")

        # ==================== 第四步：选择计算功能 ====================
        # 定义计算函数字典（key=选项，value=函数名）
        # 目前只有 1 号函数可用，其余暂时占位
        calc_functions = {
            1: med_clas_per_hun_stu,   # 每百名学生多媒体教室数
            # 2: 第二个计算函数,
            # 3: 第三个计算函数,
            # ...
            # 7: 第七个计算函数,
        }

        while True:
            print("\n" + "-" * 40)
            print("请选择要执行的计算：")
            print("  1. 每百名学生多媒体教室数")
            print("  2. （待添加）")
            print("  3. （待添加）")
            print("  4. （待添加）")
            print("  5. （待添加）")
            print("  6. （待添加）")
            print("  7. （待添加）")
            print("  8. 依次执行以上所有计算")
            print("  0. 跳过计算，直接保存/退出")
            print("-" * 40)

            choice = input("请输入选项 (0-8)：").strip()

            if choice == '0':
                break
            elif choice == '8':
                # 依次执行所有存在的计算函数
                for func in calc_functions.values():
                    func(wb, ws, head_size, max_h)
                print("所有计算已完成。")
            elif choice.isdigit() and int(choice) in calc_functions:
                func = calc_functions[int(choice)]
                func(wb, ws, head_size, max_h)
            else:
                print("输入无效，请重新选择。")
                continue

            # 每次计算后询问是否继续选择其他计算
            more = input("\n是否继续执行其他计算？(y/n，直接回车继续)：").strip().lower()
            if more == 'n':
                break

        # ==================== 第五步：询问是否保存 ====================
        save_choice = input("\n是否将结果保存到原文件？(y/n)：").strip().lower()
        if save_choice == 'y':
            try:
                wb.save(file_path)
                print("保存成功！")
            except Exception as e:
                print(f"保存失败：{e}")
        else:
            print("未保存，结果仅保留在内存中（关闭后将丢失）。")

        # ==================== 第六步：询问是否继续其他文件 ====================
        choice = input("\n是否继续分析其他文件？(y/n)：").strip().lower()
        if choice != 'y':
            break

    input("\n按 Enter 键退出程序...")
    sys.exit(0)

if __name__ == "__main__":
    main()
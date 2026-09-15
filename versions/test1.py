import pandas as pd
import os
import sys
import openpyxl

def get_excel_info(file_path):

    try:
        df = pd.read_excel(file_path, header=None)
        # 获取有效行数：从最后一行向上查找，找到有数据的最后一行
        # any(axis=1) 检查每行是否有任何一个非空值
        valid_rows = df.notna().any(axis=1)
        if valid_rows.any():
            # 找到最后一个有数据的行索引
            last_valid_row_index = valid_rows[valid_rows].index[-1]
            max_h = last_valid_row_index + 1  # 转换为行数（索引从0开始）
        else:
            max_h = 0
        # 获取有效列数：从最后一列向左查找，找到有数据的最后一列
        # any(axis=0) 检查每列是否有任何一个非空值
        valid_cols = df.notna().any(axis=0)
        if valid_cols.any():
            # 找到最后一个有数据的列索引
            last_valid_col_index = valid_cols[valid_cols].index[-1]
            max_l = last_valid_col_index + 1  # 转换为列数（索引从0开始）
        else:
            max_l = 0
        return max_h, max_l
    except FileNotFoundError:
        print(f"错误：找不到文件 '{file_path}'，请检查路径是否正确。")
        return None, None
    except Exception as e:
        print(f"读取文件时发生错误：{e}")
        return None, None

def main():
    """
    主函数：获取用户输入的Excel文件路径并输出有效行数和列数
    """
    print("=" * 50)
    print("Excel文件信息分析工具")
    print("=" * 50)

    while True:
        # 获取用户输入的Excel文件路径
        file_path = input("\n请输入Excel文件的完整路径（或拖拽文件到此处）：").strip()
        
        # 去除可能的引号（拖拽文件时可能会自动添加引号）
        file_path = file_path.strip('"').strip("'")
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"文件路径 '{file_path}' 不存在，请重新输入。")
            continue
        
        # 检查文件扩展名
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
            print(f"警告：文件扩展名为 '{file_ext}'，可能不是Excel文件，尝试读取...")
        
        # 获取文件信息
        print(f"\n正在分析文件：{file_path}")
        max_h, max_l = get_excel_info(file_path)
        
        if max_h is not None and max_l is not None:
            print("\n" + "=" * 50)
            print("分析结果：")
            print(f"有效总行数 (max_h)：{max_h} 行")
            print(f"有效总列数 (max_l)：{max_l} 列")
            print("=" * 50)
        else:
            print("文件分析失败，请检查文件是否损坏或格式是否正确。")
        
        # 询问是否继续分析其他文件
        choice = input("\n是否继续分析其他文件？(y/n)：").strip().lower()
        if choice != 'y':
            break
    
    # 等待用户按下Enter键结束脚本
    input("\n按 Enter 键退出程序...")
    sys.exit(0)

if __name__ == "__main__":
    main()
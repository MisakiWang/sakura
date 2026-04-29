import pandas as pd
import glob
import os
from openpyxl.utils import get_column_letter

# 用户输入文件夹路径
folder_path = input("请输入存放Excel文件的文件夹名称: ").strip()
# 去除可能存在的引号
folder_path = folder_path.strip('"').strip("'")
if not os.path.isdir(folder_path):
    print("错误：路径不存在或不是文件夹。")
    exit()

output_file = "汇总表.xlsx"  # 输出文件保存在当前工作目录

# 获取所有 Excel 文件
file_pattern = os.path.join(folder_path, "*.xlsx")
files = glob.glob(file_pattern)

if not files:
    print("该文件夹下没有找到 .xlsx 文件。")
    exit()

dfs = []  # 存放每个文件的有效数据

for file in files:
    try:
        # 读取第二个工作表（索引1），不设列名
        df_sheet = pd.read_excel(file, sheet_name=1, header=None)
    except Exception as e:
        print(f"跳过文件 {file}，无法读取 Sheet2: {e}")
        continue

    if df_sheet.shape[0] < 2:
        print(f"文件 {file} 的 Sheet2 数据不足，跳过")
        continue

    # 提取门店名称（A1 单元格）
    store_name = df_sheet.iloc[0, 0] if pd.notna(df_sheet.iloc[0, 0]) else os.path.basename(file)

    # 第二行作为列名
    columns = df_sheet.iloc[1].tolist()

    # 数据从第三行开始
    data = df_sheet.iloc[2:].copy()
    data.columns = columns

    # 去除完全空的行
    data = data.dropna(how='all')

    # 找到“合计”行并截取之前的数据
    first_col = data.columns[0]  # 通常是“名称”列
    mask_heji = data[first_col].astype(str).str.contains('合计', na=False)
    if mask_heji.any():
        # 第一个“合计”行的位置
        first_heji_idx = mask_heji.idxmax()
        # 截取到该行之前
        data = data.loc[:first_heji_idx-1] if first_heji_idx > data.index[0] else pd.DataFrame()
    else:
        # 如果没有合计行，则根据数量和单价列去除无关行（如制表人、签章等）
        data = data.dropna(subset=[data.columns[1], data.columns[2]], how='all')

    # 再次去除完全空的行
    data = data.dropna(how='all')

    # 添加门店列
    if not data.empty:
        data.insert(0, '门店', store_name)
        dfs.append(data)
        print(f"已处理文件: {file}，门店: {store_name}，数据行数: {len(data)}")
    else:
        print(f"文件 {file} 无有效数据行")

# 合并所有数据
if dfs:
    result = pd.concat(dfs, ignore_index=True)

    # 保存并设置列宽
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        result.to_excel(writer, index=False, sheet_name='汇总')

        # 获取工作表
        worksheet = writer.sheets['汇总']

        # 设置第一列（A列）宽度为280像素（Excel列宽约40字符）
        worksheet.column_dimensions['A'].width = 35

        # 调整其余列的列宽
        for col in worksheet.columns:
            col_letter = get_column_letter(col[0].column)
            if col_letter == 'A':
                continue  # 第一列已设置
            max_len = 0
            for cell in col:
                try:
                    cell_len = len(str(cell.value))
                    if cell_len > max_len:
                        max_len = cell_len
                except:
                    pass
            worksheet.column_dimensions[col_letter].width = max_len + 2  # 加2留余量

    print(f"合并完成，共 {len(result)} 行数据，已保存至 {output_file}")
else:
    print("未找到任何有效数据。")
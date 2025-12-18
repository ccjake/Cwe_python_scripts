import pandas as pd
import os
import glob

def merge_orders():
    # 1. 扫描当前目录下所有的 .xlsx 文件
    current_dir = os.getcwd()
    all_files = glob.glob(os.path.join(current_dir, "*.xlsx"))
    
    valid_files = {}
    
    print(f"Scanning directory: {current_dir}")
    
    for file_path in all_files:
        filename = os.path.basename(file_path)
        
        # 忽略输出文件，防止自我读取
        if filename.startswith("combined_") or filename.startswith("~"):
            continue
            
        try:
            # 读取表头，只读前几行以提高速度
            df_preview = pd.read_excel(file_path, nrows=1)
            columns = df_preview.columns.tolist()
            
            # 检查必要列是否存在
            if '客户' in columns and '地址代码' in columns:
                # 使用文件名（不含扩展名）作为月份/列名
                month_name = os.path.splitext(filename)[0]
                valid_files[month_name] = file_path
                print(f"Found valid file: {filename} -> Column: {month_name}")
            else:
                print(f"Skipping {filename}: Missing '客户' or '地址代码' columns")
                
        except Exception as e:
            print(f"Error reading {filename}: {e}")
            
    # 过滤只要数字命名的文件
    numeric_files = {}
    for month_name, file_path in valid_files.items():
        if month_name.isdigit():
            numeric_files[int(month_name)] = file_path
        else:
            print(f"Skipping non-numeric file: {month_name}")

    if not numeric_files:
        print("No valid numeric Excel files found.")
        return

    # 按数字大小排序 (9 < 10 < 11 < 12)
    sorted_months = sorted(numeric_files.keys())
    
    # 2. 建立所有 valid files 的 (客户, 地址代码) 并集
    all_pairs_list = []
    
    # 用来存储读取后的完整 dataframe，避免重复读取
    loaded_dfs = {}

    for month_num in sorted_months:
        file_path = numeric_files[month_num]
        print(f"Processing ({month_num}): {file_path}...")
        df = pd.read_excel(file_path)
        # 用原始字符串名字作为key，或者直接用数字，为了列名一致，最后 column 用 str(month_num)
        loaded_dfs[month_num] = df
        all_pairs_list.append(df[['客户', '地址代码']])
        
    all_pairs = pd.concat(all_pairs_list)
    
    # 去重，生成并集主表
    master_df = all_pairs.drop_duplicates(subset=['客户', '地址代码']).reset_index(drop=True)
    
    # 3. 填充数据
    for month_num in sorted_months:
        df_month = loaded_dfs[month_num]
        month_col_name = str(month_num) # 列名转回字符串
        # 检查是否有名为 '数量' 的列，或者我们需要动态识别数据列？
        # 根据之前的逻辑，默认只有一列数值列，通常是第三列，但最好按名字 '数量' 匹配
        # 如果没有 '数量' 列，可能需要进一步处理，这里假设都有
        target_col = '数量'
        if target_col not in df_month.columns:
            # 如果没有叫 '数量' 的，尝试取第三列（索引2）作为数值列
            if len(df_month.columns) >= 3:
                target_col = df_month.columns[2]
            else:
                print(f"Warning: Could not find data column for {month_num}")
                continue
                
        # 重命名为数字字符串
        temp = df_month[['客户', '地址代码', target_col]].rename(columns={target_col: month_col_name})
        master_df = pd.merge(master_df, temp, on=['客户', '地址代码'], how='left')
        
    # 4. 删除 '地址代码' 为空的行
    master_df_clean = master_df.dropna(subset=['地址代码'])
    
    # 5. 排序
    master_df_sorted = master_df_clean.sort_values(by=['客户', '地址代码'])
    
    # 6. 导出
    output_filename = 'combined_orders_final.xlsx'
    try:
        master_df_sorted.set_index(['客户', '地址代码']).to_excel(output_filename)
        print(f"Successfully exported to {output_filename}")
    except PermissionError:
        print(f"Error: Could not write to {output_filename}. Please close the file if it is open.")

if __name__ == "__main__":
    merge_orders()

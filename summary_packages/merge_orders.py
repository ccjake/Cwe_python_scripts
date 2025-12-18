import pandas as pd
import os
import glob
import itertools

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
                # 使用文件名（包含扩展名前部分）作为月份
                month_name = os.path.splitext(filename)[0]
                valid_files[month_name] = file_path
            else:
                pass
                
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

    # 按数字大小排序
    sorted_months = sorted(numeric_files.keys())
    
    # 2. 收集所有唯一的 客户 和 地址代码
    all_customers = set()
    all_address_codes = set()
    
    loaded_dfs = {}

    for month_num in sorted_months:
        file_path = numeric_files[month_num]
        print(f"Loading ({month_num}): {file_path}...")
        df = pd.read_excel(file_path)
        loaded_dfs[month_num] = df
        
        # 收集唯一值
        all_customers.update(df['客户'].dropna().unique())
        all_address_codes.update(df['地址代码'].dropna().unique())
        
    # 3. 构建笛卡尔积：每个客户拥有所有地址代码
    # list(itertools.product(all_customers, all_address_codes))
    print("Generating master list (Cartesian product)...")
    master_data = list(itertools.product(sorted(list(all_customers)), sorted(list(all_address_codes))))
    master_df = pd.DataFrame(master_data, columns=['客户', '地址代码'])
    
    # 4. 填充数据
    month_columns = []
    
    for month_num in sorted_months:
        df_month = loaded_dfs[month_num]
        month_col_name = str(month_num) # 列名
        month_columns.append(month_col_name)
        
        target_col = '数量'
        if target_col not in df_month.columns:
            if len(df_month.columns) >= 3:
                target_col = df_month.columns[2]
            else:
                print(f"Warning: Could not find data column for {month_num}")
                continue
                
        temp = df_month[['客户', '地址代码', target_col]].rename(columns={target_col: month_col_name})
        
        # 合并
        master_df = pd.merge(master_df, temp, on=['客户', '地址代码'], how='left')

    # 5. 添加合计行
    print("Calculating summaries...")
    
    # 为了方便排序，我们给原始数据打标 is_summary=0
    master_df['is_summary'] = 0
    
    # 计算按客户分组的合计
    summary_dfs = []
    grouped = master_df.groupby('客户')
    
    for customer, group in grouped:
        # 计算该客户所有月份列的和
        # min_count=0 确保全是 NaN 时和为 0 (或者保留 NaN，看需求，通常 sum 会把 NaN 视为 0)
        sums = group[month_columns].sum(numeric_only=True)
        
        # 构建一行 summary
        row = {'客户': customer, '地址代码': 'Total', 'is_summary': 1}
        for col in month_columns:
            row[col] = sums[col]
            
        summary_dfs.append(row)
        
    summary_df = pd.DataFrame(summary_dfs)
    
    # 合并原始数据和合计
    final_df = pd.concat([master_df, summary_df], ignore_index=True)
    
    # 6. 排序：先按客户，再按 is_summary (0在1前)，最后按地址代码
    # 这样 'Total' 行会出现在每个客户的最后
    final_df = final_df.sort_values(by=['客户', 'is_summary', '地址代码'])
    
    # 移除辅助列 is_summary
    final_df = final_df.drop(columns=['is_summary'])
    
    # 7. 导出
    output_filename = 'combined_orders_final.xlsx'
    try:
        # 设置索引可以合并单元格
        final_df.set_index(['客户', '地址代码']).to_excel(output_filename)
        print(f"Successfully exported to {output_filename}")
    except PermissionError:
        print(f"Error: Could not write to {output_filename}. Please close the file if it is open.")

if __name__ == "__main__":
    merge_orders()

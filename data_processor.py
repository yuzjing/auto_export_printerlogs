# data_processor.py (严格按照您的要求，在您的基准代码上只增加 floor4 日期修复)

import os
import pandas as pd
import numpy as np

from database_utils import create_connection, create_table_if_not_exists, insert_log_data

# --- 配置区域 (与您的基准代码完全一致) ---
FINAL_DB_COLUMNS = [
    'log_date', 'end_time', 'user', 'output_location', 'job_info', 
    'page_info', 'pages', 'sheets', 'job_status', 'job_level', 'job_id'
]
FLOOR2_COLUMN_MAP = {
    '日期': 'log_date',
    '结束时间': 'end_time',
    '输入(发送)方': 'user',  
    '输出位置': 'output_location',
    '作业信息': 'job_info',
    '页面信息': 'page_info',
    '页数': 'pages',
    '张数': 'sheets',
    '作业处理状态': 'job_status',
    '作业层次': 'job_level',
    '作业号码': 'job_id'
}
FLOOR4_COLUMN_MAP = {
    '日期': 'log_date',
    '结束时间': 'end_time',
    '输入(发送)方': 'user', 
    '输出位置': 'output_location',
    '作业信息': 'job_info',
    '页面信息': 'page_info',
    '页数': 'pages',
    '张数': 'sheets',
    '作业处理状态': 'job_status',
    '子作业': 'job_level',
    '作业号码': 'job_id'
}
PRINTER_COLUMN_MAPS = {
    'floor2': FLOOR2_COLUMN_MAP,
    'floor4': FLOOR4_COLUMN_MAP
}

def process_and_insert_data(directory_path, printer_type):
    # 函数前半部分和文件查找逻辑与您的基准代码完全一致
    print(f"\n--- [数据处理模块] 开始处理 '{printer_type}' 的数据，来源: {directory_path} ---")
    if printer_type not in PRINTER_COLUMN_MAPS:
        print(f"!! 严重错误: 未知的打印机类型 '{printer_type}'。")
        return
    column_map = PRINTER_COLUMN_MAPS[printer_type]
    expected_header_sample = list(column_map.keys())[0]
    try:
        all_files = os.listdir(directory_path)
        csv_filenames = [f for f in all_files if f.lower().endswith('.csv')]
        csv_files_full_path = [os.path.join(directory_path, f) for f in csv_filenames]
    except Exception as e:
        print(f"!! 错误: 无法读取目录 {directory_path} 的内容: {e}")
        return
    if not csv_files_full_path:
        print("未找到任何CSV文件，处理结束。")
        return
    print(f"找到了 {len(csv_files_full_path)} 个CSV文件，开始解析...")
    all_data_to_insert = []
    
    for csv_file in csv_files_full_path:
        df = None
        # 编码检测部分与您的基准代码完全一致
        potential_encodings = ['utf-8', 'gb18030', 'gbk', 'cp932', 'shift_jis']
        for encoding in potential_encodings:
            try:
                # 使用您基准代码中的 read_csv，不加 engine 参数
                temp_df = pd.read_csv(csv_file, encoding=encoding)
                if not temp_df.empty and expected_header_sample in temp_df.columns:
                    df = temp_df
                    break
            except Exception:
                continue
        if df is None:
            print(f"!! 严重错误: 无法使用任何备选编码正确解析文件 {os.path.basename(csv_file)}。已跳过。")
            continue
        try:
            if df.empty:
                continue
            df.rename(columns=column_map, inplace=True)
            df_processed = df[FINAL_DB_COLUMNS]
            df_processed.dropna(how='all', inplace=True)

            # --- 【【【唯一的核心修改：在这里增加针对 floor4 的特殊处理】】】 ---
            
            # 1. 检查当前处理的是否是 floor4
            if printer_type == 'floor4':
                # 从文件名提取日期
                basename = os.path.basename(csv_file)
                date_from_filename = basename[8:18] # 提取 'YYYY-MM-DD'
                # 用从文件名得到的正确日期，覆盖掉整个日期列
                df_processed['log_date'] = date_from_filename
                print(f"   -> [floor4 特殊处理] 已从文件名 '{basename}' 强制填充日期: {date_from_filename}")

            # 2. 强制解析日期列（对floor2生效，对floor4是二次确认）
            df_processed['log_date'] = pd.to_datetime(df_processed['log_date'], errors='coerce')
            
            # 3. 清洗数值列 'pages' 和 'sheets' (与您的基准代码完全一致)
            df_processed['pages'] = pd.to_numeric(df_processed['pages'], errors='coerce').fillna(0).astype(int)
            df_processed['sheets'] = pd.to_numeric(df_processed['sheets'], errors='coerce').fillna(0).astype(int)
            
            # --- 【【【修改结束】】】 ---

            # 4. 对所有列，将 pandas 的缺失值转换成数据库能识别的 None (与您的基准代码完全一致)
            df_processed = df_processed.astype(object).where(pd.notnull(df_processed), None)
            
            # 准备最终插入的数据元组列表 (与您的基准代码完全一致)
            data_tuples = [(printer_type,) + row for row in df_processed.itertuples(index=False)]
            
            all_data_to_insert.extend(data_tuples)
            
        except Exception as e:
            print(f"!! 处理文件 {os.path.basename(csv_file)} 时发生未知错误: {e}")

    # 函数后半部分的数据库连接和插入逻辑与您的基准代码完全一致
    print(f"\n文件解析完成。共找到 {len(all_data_to_insert)} 条有效记录。")
    
    if all_data_to_insert:
        db_connection = create_connection()
        if db_connection:
            try:
                create_table_if_not_exists(db_connection)
                insert_log_data(db_connection, all_data_to_insert)
            finally:
                db_connection.close()
    else:
        print("没有可插入数据库的数据。")

    print(f"--- [数据处理模块] '{printer_type}' 处理完毕 ---")
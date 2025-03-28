#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
-------------------------------------------------
   NIO PIT 自动化工具
   Author:  Charles Xu (charles.xu2@nio.com)
   Company: NIO
   Created: 2025-02-25 | Version: 1.0
-------------------------------------------------
   Copyright (c) 2025 Charles Xu (NIO)
   Licensed under the MIT License: 
   https://opensource.org/licenses/MIT
-------------------------------------------------
   Change Log:
   2025-02-25  Charles Xu  Initial creation
-------------------------------------------------
"""

import os
import sys
import re
import pandas as pd
import numpy as np
# 添加项目根目录到 Python 路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils import read_excel, save_excel
from parse_diag_table.config import get_column_index
from parse_diag_table.config import (
    SOURCE_SHEET1, SOURCE_SHEET2, SOURCE_SHEET3,
    TARGET_SHEET1, TARGET_SHEET2,
)

# 读取 Excel 文件并做预处理
def pre_process_22_2E(source_file, target_file):
    source_df1 = read_excel(source_file, SOURCE_SHEET1)
    source_df2 = read_excel(source_file, SOURCE_SHEET2)
    target_df = read_excel(target_file, TARGET_SHEET1)
    
    # 提取 source_did_list
    source_did_list = []
    # 处理 source_sheet1
    for idx, value in enumerate(source_df1.iloc[:, get_column_index('A', 'basic_did')].dropna().astype(str)):
        if value.startswith("0x"):
            # print(f"\n找到 DID: {value}")
            # 使用布尔索引找到对应的行
            matching_rows = source_df1[source_df1.iloc[:, get_column_index('A', 'basic_did')] == value]
            if len(matching_rows) == 0:
                print(f"错误: 未找到 DID {value} 的完整数据")
                continue
                
            row_data = matching_rows.iloc[0]
            # print("该行所有数据:")
            # for col in source_df1.columns:
            #     print(f"{col}: {row_data[col]}")
            
            # 检查 Support 列的值
            support_value = str(row_data.iloc[get_column_index('E', 'basic_did')]).strip().upper()
            if pd.isna(support_value) or support_value == 'NAN':
                print(f"警告: DID {value} 的 Support 值为空或无效，默认设置为 'N'")
                support_value = 'N'
            if support_value != 'N':
                source_did_list.append(value.replace("0x", ""))
            else:
                # print(f"DID {value} 的 Support 值为 N，跳过处理")
                pass

    # 处理 source_sheet2
    for value in source_df2.iloc[:, get_column_index('A', 'rdbi_wdbi')].dropna().astype(str):
        if value.startswith("0x"):
            source_did_list.append(value.replace("0x", ""))

    length = source_did_list

    print(f'\nsource DID len: {len(source_did_list)}; \nsource DID: {', '.join(source_did_list)}')
    
    # 提取 target_did_list
    target_did_list = {}
    for idx, value in enumerate(target_df.iloc[:, get_column_index('A', 'did_library')].dropna().astype(str)):
       if re.fullmatch(r"[a-zA-Z0-9]+", value):
         target_did_list[value] = idx
    print(f'target DID len: {len(target_did_list)}; \ntarget DID: {', '.join(target_did_list.keys())}')

    # 进行比对和更新
    new_rows = []
    removed_indices = []
    
    for value in source_did_list:
        if value not in target_did_list:
            print(f"新增: {value}")
            new_rows.append(value)
    
    for value, idx in list(target_did_list.items()):
        if value not in source_did_list:
            print(f"删除: {value}")
            removed_indices.append(idx)
    
    # 添加新行
    sel_col = "A"
    sel_col_idx = get_column_index('A', 'did_library')

    for value in new_rows:
        new_row = [np.nan] * len(target_df.columns)  # 生成全 NaN 行
        new_row[sel_col_idx] = value  # 只填充 A 列 也就是 DID
        target_df.loc[len(target_df)] = new_row  # 追加新行

    # 删除不存在的行
    target_df.drop(removed_indices, inplace=True)
    target_df.reset_index(drop=True, inplace=True)
    
    return source_df1, source_df2, target_df

def main_process_22_2E(data, source_df1, source_df2, target_df, target_row_index):
    source_row1 = source_df1.iloc[:, get_column_index('A', 'basic_did')]
    source_row2 = source_df2.iloc[:, get_column_index('A', 'rdbi_wdbi')]
    data_prefixed = f"0x{data}"  # 添加 0x 前缀
    
    # print(data_prefixed)
    # print(source_row1.values)
    # 过滤掉 NA 值
    filtered_values = [val for val in source_row1.values if pd.notna(val)]
    filtered_values2 = [val for val in source_row2.values if pd.notna(val)]

    if data_prefixed in filtered_values:
        process_updates(source_df1[source_df1.iloc[:, get_column_index('A', 'basic_did')] == data_prefixed].iloc[0], target_df, target_row_index, source_flag='source_sheet1')
    elif data_prefixed in filtered_values2:
        process_updates(source_df2[source_df2.iloc[:, get_column_index('A', 'rdbi_wdbi')] == data_prefixed].iloc[0], target_df, target_row_index, source_flag='source_sheet2')
    else:
        raise ValueError(f"数据 '{data_prefixed}' 未在 source_sheet1 或 source_sheet2 中找到，程序退出。")

def compute_app(update1, update2):
    if update1 == 'yes' and update2 == 'yes':
        return '22/2E'
    elif update1 == 'yes':
        return '22'
    elif update2 == 'yes':
        return '2E'
    else:
        return 'NA'

# def format_security_level(update1, update2):
#     levels = [update1, update2]
#     levels = [lvl for lvl in levels if lvl in ['Lock', 'L1', 'L2', 'L3', 'L4', 'L5']]
#     return '/'.join(sorted(set(levels), key=lambda x: ['Lock', 'L1', 'L2', 'L3', 'L4', 'L5'].index(x)))

def format_security_level(update1, update2):
    levels = [lvl for lvl in [update1, update2] if isinstance(lvl, str) and lvl.lower() != 'nan' and lvl]  # 过滤空值和 'nan'
    if not levels:
        return 'NA'  # 如果都为空，返回 'NA'
    levels = '/'.join(levels).split('/')
    levels = [lvl for lvl in levels if lvl in ['Locked', 'L1', 'L2', 'L3', 'L4', 'L5']]
    sorted_levels = sorted(set(levels), key=lambda x: ['Locked', 'L1', 'L2', 'L3', 'L4', 'L5'].index(x))
    return '/'.join(sorted_levels).replace('Locked', 'Lock')


def process_updates(source_row, target_df, target_row_index, source_flag):
    update_description(source_row, target_df, target_row_index, source_flag)
    update_format(source_row, target_df, target_row_index, source_flag)
    update_length(source_row, target_df, target_row_index, source_flag)
    update_app(source_row, target_df, target_row_index, source_flag)
    update_boot(source_row, target_df, target_row_index, source_flag)
    update_security_2E(source_row, target_df, target_row_index, source_flag)
    update_security_22(source_row, target_df, target_row_index, source_flag)

def update_description(source_row, target_df, target_row_index, source_flag):
    update_column(target_df, target_row_index, get_column_index('B', 'did_library'), source_row.iloc[get_column_index('B', 'basic_did')], "Description")

def update_format(source_row, target_df, target_row_index, source_flag):
    col_index = get_column_index('AB', 'basic_did') if source_flag=="source_sheet1" else get_column_index('Z', 'rdbi_wdbi')
    format_value = source_row.iloc[col_index]
    # 如果 format_value 为空，则设置为 'HEX'
    if pd.isna(format_value) or format_value == 'nan':
        format_value = 'HEX'
    else:
        format_value = format_value if format_value in ['Bytefield', 'ASCII', 'BCD'] else 'HEX'
    update_column(target_df, target_row_index, get_column_index('C', 'did_library'), format_value, "Format")

def update_length(source_row, target_df, target_row_index, source_flag):
    col_index = get_column_index('F', 'basic_did') if source_flag=="source_sheet1" else get_column_index('D', 'rdbi_wdbi')
    update_column(target_df, target_row_index, get_column_index('D', 'did_library'), source_row.iloc[col_index], "Length")

def update_app(source_row, target_df, target_row_index, source_flag):
    if source_flag=="source_sheet1":
        cols = [get_column_index('H', 'basic_did'), get_column_index('N', 'basic_did')]
    else:
        cols = [get_column_index('F', 'rdbi_wdbi'), get_column_index('L', 'rdbi_wdbi')]
    update1, update2 = source_row.iloc[cols[0]], source_row.iloc[cols[1]]
    update_final = compute_app(update1, update2)
    update_column(target_df, target_row_index, get_column_index('E', 'did_library'), update_final, "APP")

def update_boot(source_row, target_df, target_row_index, source_flag):
    if source_flag=="source_sheet1":
        cols = [get_column_index('K', 'basic_did'), get_column_index('R', 'basic_did')]
    else:
        cols = [get_column_index('I', 'rdbi_wdbi'), get_column_index('P', 'rdbi_wdbi')]
    update1, update2 = source_row.iloc[cols[0]], source_row.iloc[cols[1]]
    update_final = compute_app(update1, update2)
    update_column(target_df, target_row_index, get_column_index('F', 'did_library'), update_final, "Boot")

def update_security_2E(source_row, target_df, target_row_index, source_flag):
    if source_flag=="source_sheet1":
        cols = [get_column_index('O', 'basic_did'), get_column_index('S', 'basic_did')]
    else:
        cols = [get_column_index('M', 'rdbi_wdbi'), get_column_index('Q', 'rdbi_wdbi')]
    update1, update2 = source_row.iloc[cols[0]], source_row.iloc[cols[1]]
    update_final = format_security_level(update1, update2)
    update_column(target_df, target_row_index, get_column_index('G', 'did_library'), update_final, "Security Level 2E")

def update_security_22(source_row, target_df, target_row_index, source_flag):
    if source_flag=="source_sheet1":
        cols = [get_column_index('I', 'basic_did'), get_column_index('L', 'basic_did')]
    else:
        cols = [get_column_index('G', 'rdbi_wdbi'), get_column_index('J', 'rdbi_wdbi')]
    update1, update2 = source_row.iloc[cols[0]], source_row.iloc[cols[1]]
    update_final = format_security_level(update1, update2)
    update_column(target_df, target_row_index, get_column_index('H', 'did_library'), update_final, "Security Level 22")

def update_column(target_df, row_index, col_index, new_value, column_name):
    if target_df.iloc[row_index, col_index] != new_value:
        # print(f"Updating {column_name}[{row_index}]: {target_df.iloc[row_index, col_index]} -> {new_value}")
        target_df.iloc[row_index, col_index] = new_value


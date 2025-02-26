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


def main_process(data, source_df1, source_df2, target_df, target_row_index):
    source_row1 = source_df1.iloc[:, 0]
    source_row2 = source_df2.iloc[:, 0]
    data_prefixed = f"0x{data}"  # 添加 0x 前缀
    
    if data_prefixed in source_row1.values:
        process_updates(source_df1[source_df1.iloc[:, 0] == data_prefixed].iloc[0], target_df, target_row_index, source_flag='source_sheet1')
    elif data_prefixed in source_row2.values:
        process_updates(source_df2[source_df2.iloc[:, 0] == data_prefixed].iloc[0], target_df, target_row_index, source_flag='source_sheet2')
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
    update_column(target_df, target_row_index, 1, source_row.iloc[1], "Description")

def update_format(source_row, target_df, target_row_index, source_flag):
    col_index = 27 if source_flag=="source_sheet1" else 25
    format_value = source_row.iloc[col_index]
    format_value = format_value if format_value in ['Bytefield', 'ASCII', 'BCD'] else 'HEX'
    update_column(target_df, target_row_index, 2, format_value, "Format")

def update_length(source_row, target_df, target_row_index, source_flag):
    col_index = 5 if source_flag=="source_sheet1" else 3
    update_column(target_df, target_row_index, 3, source_row.iloc[col_index], "Length")

def update_app(source_row, target_df, target_row_index, source_flag):
    cols = [7, 13] if source_flag=="source_sheet1" else [5, 11]
    update1, update2 = source_row.iloc[cols[0]], source_row.iloc[cols[1]]
    update_final = compute_app(update1, update2)
    update_column(target_df, target_row_index, 4, update_final, "APP")

def update_boot(source_row, target_df, target_row_index, source_flag):
    cols = [10, 17] if source_flag=="source_sheet1" else [8, 15]
    update1, update2 = source_row.iloc[cols[0]], source_row.iloc[cols[1]]
    update_final = compute_app(update1, update2)
    update_column(target_df, target_row_index, 5, update_final, "Boot")

def update_security_2E(source_row, target_df, target_row_index, source_flag):
    cols = [14, 18] if source_flag=="source_sheet1" else [12, 16]
    update1, update2 = source_row.iloc[cols[0]], source_row.iloc[cols[1]]
    update_final = format_security_level(update1, update2)
    update_column(target_df, target_row_index, 6, update_final, "Security Level 2E")

def update_security_22(source_row, target_df, target_row_index, source_flag):
    cols = [8, 11] if source_flag=="source_sheet1" else [6, 9]
    update1, update2 = source_row.iloc[cols[0]], source_row.iloc[cols[1]]
    update_final = format_security_level(update1, update2)
    update_column(target_df, target_row_index, 7, update_final, "Security Level 22")

def update_column(target_df, row_index, col_index, new_value, column_name):
    if target_df.iloc[row_index, col_index] != new_value:
        # print(f"Updating {column_name}[{row_index}]: {target_df.iloc[row_index, col_index]} -> {new_value}")
        target_df.iloc[row_index, col_index] = new_value

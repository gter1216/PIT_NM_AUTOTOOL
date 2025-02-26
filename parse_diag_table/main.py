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


import re
import os
import sys
import pandas as pd
from config import CURR_REL_DIR, LAST_REL_DIR, OUTPUT_DIR, SOURCE_SHEET1, SOURCE_SHEET2, TARGET_SHEET
from utils import read_excel, save_excel
from main_process import main_process
import numpy as np


# 读取 Excel 文件并做预处理
def pre_process(source_file, target_file, output_file):
    source_df1 = read_excel(source_file, SOURCE_SHEET1)
    source_df2 = read_excel(source_file, SOURCE_SHEET2)
    target_df = read_excel(target_file, TARGET_SHEET)

    # 提取 source_did_list
    source_did_list = []
    for value in source_df1.iloc[:, 0].dropna().astype(str):
        if value.startswith("0x"):
            source_did_list.append(value.replace("0x", ""))
   
    for value in source_df2.iloc[:, 0].dropna().astype(str):
        if value.startswith("0x"):
            source_did_list.append(value.replace("0x", ""))

    length = source_did_list

    print(f'source DID len: {len(source_did_list)}; \nsource DID: {', '.join(source_did_list)}')

    #  for value in source_df2.iloc[:, 0].dropna().astype(str):
    #      if re.fullmatch(r"[a-zA-Z0-9]+", value):
    #          source_did_list.add(value.replace("0x", ""))
    
    # 提取 target_did_list
    target_did_list = {}
    for idx, value in enumerate(target_df.iloc[:, 0].dropna().astype(str)):
       if re.fullmatch(r"[a-zA-Z0-9]+", value):
         target_did_list[value] = idx
    print(f'target DID len: {len(target_did_list)}; \ntarget DID: {', '.join(map(str, target_did_list.values()))}')

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
    sel_col_idx = ord(sel_col) - ord('A')

    for value in new_rows:
        new_row = [np.nan] * len(target_df.columns)  # 生成全 NaN 行
        new_row[sel_col_idx] = value  # 只填充 A 列 也就是 DID
        target_df.loc[len(target_df)] = new_row  # 追加新行

    # 删除不存在的行
    target_df.drop(removed_indices, inplace=True)
    target_df.reset_index(drop=True, inplace=True)

    save_excel(target_df, output_file, sheet_name=TARGET_SHEET)
    print("Excel 文件已更新。")
    
    return source_df1, source_df2, target_df

# 主函数
def process_file(source_file, target_file, output_file):
    source_df1, source_df2, target_df = pre_process(source_file, target_file, output_file)

    for idx, row in target_df.iterrows():
        # print(row.iloc[0])
        main_process(row.iloc[0], source_df1, source_df2, target_df, idx)
    
    save_excel(target_df, output_file, sheet_name=TARGET_SHEET)
    print(f"Updated file saved to {output_file}")

def get_excel_files(directory):
    """ 获取目录下所有 Excel 文件（忽略临时文件 `~$` 开头的文件） """
    return [os.path.join(directory, f) for f in os.listdir(directory) 
            if f.endswith(".xlsx") and not f.startswith("~$")]

def main():
    """ 处理多个 Excel 文件 """
    source_files = get_excel_files(CURR_REL_DIR)
    target_files = get_excel_files(LAST_REL_DIR)

    if len(source_files) != len(target_files):
        raise ValueError("错误：新版本和旧版本文件数量不匹配，请检查文件夹内容。")
        # sys.exit(1)

    output_files = [os.path.join(OUTPUT_DIR, os.path.basename(f).replace(".xlsx", "_更新后.xlsx")) for f in target_files]

    for i in range(len(source_files)):
        source_file = source_files[i]
        target_file = target_files[i]
        output_file = output_files[i]

        process_file(source_file, target_file, output_file)

        # 处理完当前文件后，提示用户是否继续
        user_input = input(f"文件 {source_file} 处理完毕，是否继续处理下一个文件？(y/n): ").strip().lower()
        if user_input == 'n':
            print("用户选择退出，程序终止。")
            sys.exit(0)  # 退出程序

    print("所有文件处理完毕。")

if __name__ == "__main__":
    main()
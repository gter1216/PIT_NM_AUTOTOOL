#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
-------------------------------------------------
   NIO PIT UDS 自动化工具
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
import argparse
import re
import pandas as pd
# 添加项目根目录到 Python 路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from parse_diag_table.config import (
    VEHICLE_TYPE,
    SOURCE_SHEET1, SOURCE_SHEET2, SOURCE_SHEET3,
    TARGET_SHEET1, TARGET_SHEET2,
    get_data_dirs
)
from parse_diag_table.utils import read_excel, save_excel
from parse_diag_table.main_process_22_2E import main_process_22_2E,pre_process_22_2E
from parse_diag_table.main_process_31 import main_process_31,pre_process_31,save_excel_with_merged_cells
import numpy as np


# 主函数
def process_22_2E(source_file, target_file):
    print(f"Begin to process 22 and 2E sheet")
    source_df1, source_df2, target_df = pre_process_22_2E(source_file, target_file)

    for idx, row in target_df.iterrows():
        # print(row.iloc[0])
        main_process_22_2E(row.iloc[0], source_df1, source_df2, target_df, idx)
    
    print(f"Finished process 22 and 2E sheet")
    return target_df

def process_31(source_file, target_file):
    print(f"Begin to process 31 sheet")
    source_df, target_df = pre_process_31(source_file, target_file)
    # 处理每个 Routine Control ID 的具体数据
    for idx in range(0, len(target_df), 3):  # 每三行处理一次
        rid = target_df.iloc[idx, 0]  # A列 Routine Control ID
        if pd.notna(rid):  # 确保不是空值
            main_process_31(source_df, target_df, rid)
    
    # 保存处理后的数据到 Excel，并合并单元格
    save_excel_with_merged_cells(target_df, target_file, TARGET_SHEET2)
    
    print(f"Finished process 31 sheet")
    return target_df

def get_excel_files(directory):
    """ 获取目录下所有 Excel 文件（忽略临时文件 `~$` 开头的文件） """
    return [os.path.join(directory, f) for f in os.listdir(directory) 
            if f.endswith(".xlsx") and not f.startswith("~$")]

def main():
    """ 处理多个 Excel 文件 """
    # 设置命令行参数
    parser = argparse.ArgumentParser(description='NIO PIT UDS 自动化工具')
    parser.add_argument('--vehicle', type=str, help='车辆类型 (例如: BLANC_RL201, CETUS_RL201)')
    args = parser.parse_args()
    
    # 确定使用的车辆类型
    vehicle_type = args.vehicle if args.vehicle else VEHICLE_TYPE
    print(f"使用车辆类型: {vehicle_type}")
    
    # 获取数据目录路径
    curr_rel_dir, last_rel_dir, output_dir = get_data_dirs(vehicle_type)
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 检查源目录和目标目录是否存在
    if not os.path.exists(curr_rel_dir):
        raise FileNotFoundError(f"源目录不存在: {curr_rel_dir}")
    if not os.path.exists(last_rel_dir):
        raise FileNotFoundError(f"目标目录不存在: {last_rel_dir}")
    
    source_files = get_excel_files(curr_rel_dir)
    target_files = get_excel_files(last_rel_dir)

    if len(source_files) != len(target_files):
        raise ValueError("错误：新版本和旧版本文件数量不匹配，请检查文件夹内容。")

    output_files = [os.path.join(output_dir, os.path.basename(f).replace(".xlsx", "_更新后.xlsx")) for f in target_files]

    for i in range(len(source_files)):
        print(f"####开始处理### \n标准诊断表：{source_files[i]} \n模板诊断表：{target_files[i]}")
        source_file = source_files[i]
        target_file = target_files[i]
        output_file = output_files[i]

        # 处理 22 和 2E 相关的内容
        target_df_22_2E = process_22_2E(source_file, target_file)
        
        # 处理 RoutineControl 0x31 相关的内容
        target_df_31 = process_31(source_file, target_file)
        
        # 保存所有处理结果
        save_excel(target_df_22_2E, output_file, sheet_name=TARGET_SHEET1, mode="w")
        # save_excel(target_df_31, output_file, sheet_name=TARGET_SHEET2, mode="a")
        save_excel_with_merged_cells(target_df_31, output_file, TARGET_SHEET2)
        print(f"处理完毕, 所有处理结果已保存到: {output_file}")

        # 处理完当前文件后，提示用户是否继续
        user_input = input(f"文件 {source_file} 处理完毕，是否继续处理下一个文件？(y/n): ").strip().lower()
        if user_input == 'n':
            print("用户选择退出，程序终止。")
            sys.exit(0)  # 退出程序

    print("所有文件处理完毕。")

if __name__ == "__main__":
    main()
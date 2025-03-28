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

import pandas as pd
import warnings
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import numpy as np

# 忽略 openpyxl 的样式警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl.styles.stylesheet')

def read_excel(file_path, sheet_name):
    # 读取 Excel 文件，将 NA 值读取为字符串
    df = pd.read_excel(file_path, sheet_name=sheet_name, na_filter=False)
    
    def process_cell_value(value):
        """处理单元格的值
        Args:
            value: 单元格的值
        Returns:
            处理后的值:
            - 空字符串、纯空格：转换为 pd.NA
            - 字符串 'NA', 'NAN'：保持原值
            - pandas.NA：转换为 pd.NA
            - 其他：保持原值
        """
        # 处理字符串类型
        if isinstance(value, str):
            # 去除前后空格
            value = value.strip()
            # 如果是空字符串，转为 pd.NA
            if value == '':
                return pd.NA
            # 其他字符串保持原值
            return value
            
        # 处理 pandas 的 NA 类型
        if pd.isna(value):
            return pd.NA
            
        # 其他类型保持原值
        return value
    
    # 处理每一列的数据
    for col in df.columns:
        df[col] = df[col].apply(process_cell_value)
    
    # 删除所有值都是空值的行（不包括 'NA' 字符串）
    df = df.dropna(how='all')
    
    # 重置索引
    df = df.reset_index(drop=True)
    return df

def save_excel(df, file_path, sheet_name, mode='w'):
    """
    保存 DataFrame 到 Excel 文件
    :param df: DataFrame 对象
    :param file_path: 保存路径
    :param sheet_name: 工作表名称
    :param mode: 'w' 表示覆盖写入，'a' 表示追加写入
    """
    try:
        if mode == 'w':
            # 创建新的 Excel 文件
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                # 获取工作表
                worksheet = writer.sheets[sheet_name]
                # 调整列宽
                for idx, col in enumerate(df.columns):
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2
        else:
            # 追加到现有文件
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                # 获取工作表
                worksheet = writer.sheets[sheet_name]
                # 调整列宽
                for idx, col in enumerate(df.columns):
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2
    except Exception as e:
        print(f"保存 Excel 文件时出错: {str(e)}")
        raise
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

def read_excel(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name)

def save_excel(dataframe, file_path, sheet_name="Sheet1"):
    """保存 DataFrame 到 Excel，可指定 sheet_name"""
    ### 如果 excel 文件存在，mode="w" 会删除 excel 文件创建新文件。
    with pd.ExcelWriter(file_path, mode="w") as writer:
        dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
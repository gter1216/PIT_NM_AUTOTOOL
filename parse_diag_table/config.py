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

# 基础目录配置
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "parse_diag_table", "data")

# 车辆类型配置
#### CETUS_RL201
# VEHICLE_TYPE = "CETUS_RL201"
### BLANC_RL201
VEHICLE_TYPE = "BLANC_RL201"

# 目录配置
def get_data_dirs(vehicle_type):
    """获取数据目录路径
    Args:
        vehicle_type: 车辆类型
    Returns:
        tuple: (curr_rel_dir, last_rel_dir, output_dir)
    """
    curr_rel_dir = os.path.join(DATA_DIR, vehicle_type, "CURR_REL")
    last_rel_dir = os.path.join(DATA_DIR, vehicle_type, "LAST_REL")
    output_dir = os.path.join(DATA_DIR, vehicle_type, "OUTPUT")
    return curr_rel_dir, last_rel_dir, output_dir

# 工作表配置
SOURCE_SHEET1 = "3.1Basic DIDs"
SOURCE_SHEET2 = "3.2RDBI 0x22 & WDBI 0x2E"
TARGET_SHEET1 = "22_2E_DID_Library"

SOURCE_SHEET3 = "3.3RoutineControl 0x31"
TARGET_SHEET2 = "31_RID_Library"


# 列索引映射
# CURR_REL_DIR 的列映射 - 3.1 Basic DIDs
BASIC_DID_COLUMN_MAP = {
    'A': 0,   # DID Number(hex)
    'B': 1,   # DID Name (English)
    'C': 2,   # DID Name(Chinese)
    'E': 4,   # Support
    'F': 5,   # Byte Length(Dec)
    'H': 7,   # Read.APP.Support
    'I': 8,   # Read.APP.Access_Level
    'J': 9,   # Read.APP.Session
    'K': 10,  # Read.Booloader.Supprt
    'L': 11,  # Read.Bootloader.Access_Level
    'M': 12,  # Read.Bootloader.Session
    'N': 13,  # Write.APP.Support
    'O': 14,  # Write.APP.Access_Level
    'P': 15,  # Write.APP.Session
    'R': 17,  # Write.Booloader.Supprt
    'S': 18,  # Write.Bootloader.Access_Level
    'T': 19,  # Write.Bootloader.Session
    'AB': 27  # Formula
}

# CURR_REL_DIR 的列映射 - 3.2 RDBI 0x22 & WDBI 0x2E
RDBI_WDBI_COLUMN_MAP = {
    'A': 0,   # DID Number(hex)
    'B': 1,   # DID Name (English)
    'C': 2,   # DID Name(Chinese)
    'D': 3,   # Byte Length(Dec)
    'E': 4,   # Storage
    'F': 5,   # Read.APP.Support
    'G': 6,   # Read.APP.Access_Level
    'H': 7,   # Read.APP.Session
    'I': 8,   # Read.Booloader.Supprt
    'J': 9,   # Read.Bootloader.Access_Level
    'K': 10,  # Read.Bootloader.Session
    'L': 11,  # Write.APP.Support
    'M': 12,  # Write.APP.Access_Level
    'N': 13,  # Write.APP.Session
    'P': 15,  # Write.Booloader.Supprt
    'Q': 16,  # Write.Bootloader.Access_Level
    'R': 17,  # Write.Bootloader.Session
    'Z': 25   # Formula
}

# CURR_REL_DIR 的列映射 - 3.3 RoutineControl 0x31
ROUTINE_CONTROL_COLUMN_MAP = {
    'A': 0,   # RID Number(hex)
    'B': 1,   # RID Name(English)
    'C': 2,   # RID Name(Chinese)
    'D': 3,   # Access Level
    'E': 4,   # Session
    'F': 5,   # Routine Condition
    'G': 6,   # Subfunction
    'H': 7,   # Req/Resp
    'M': 12   # BitLength
}

# LAST_REL_DIR 和 OUTPUT_DIR 的列映射 - 22_2E_DID_Library
DID_LIBRARY_COLUMN_MAP = {
    'A': 0,  # DID
    'B': 1,  # Description
    'C': 2,  # Format
    'D': 3,  # Length
    'E': 4,  # APP
    'F': 5,  # Boot
    'G': 6,  # Security Level(2E)
    'H': 7   # Security Level(22)
}

# LAST_REL_DIR 和 OUTPUT_DIR 的列映射 - 31_RID_Library
RID_LIBRARY_COLUMN_MAP = {
    'A': 0,  # RID
    'B': 1,  # Description
    'C': 2,  # APP
    'D': 3,  # Boot
    'E': 4,  # SubService
    'F': 5,  # LockLevel
    'G': 6,  # Session
    'H': 7,  # RequestData
    'I': 8,  # ResponseLength
    'J': 9,  # ResponseData
    'K': 10  # ResponseNRC
}

# 反向映射，用于调试
BASIC_DID_COLUMN_LETTERS = {v: k for k, v in BASIC_DID_COLUMN_MAP.items()}
RDBI_WDBI_COLUMN_LETTERS = {v: k for k, v in RDBI_WDBI_COLUMN_MAP.items()}
ROUTINE_CONTROL_COLUMN_LETTERS = {v: k for k, v in ROUTINE_CONTROL_COLUMN_MAP.items()}
DID_LIBRARY_COLUMN_LETTERS = {v: k for k, v in DID_LIBRARY_COLUMN_MAP.items()}
RID_LIBRARY_COLUMN_LETTERS = {v: k for k, v in RID_LIBRARY_COLUMN_MAP.items()}

def get_column_index(letter, sheet_type='basic_did'):
    """获取列字母对应的索引
    Args:
        letter: 列字母
        sheet_type: 表格类型，可选值：
            - 'basic_did': 3.1 Basic DIDs
            - 'rdbi_wdbi': 3.2 RDBI 0x22 & WDBI 0x2E
            - 'routine_control': 3.3 RoutineControl 0x31
            - 'did_library': 22_2E_DID_Library
            - 'rid_library': 31_RID_Library
    """
    if sheet_type == 'basic_did':
        return BASIC_DID_COLUMN_MAP[letter]
    elif sheet_type == 'rdbi_wdbi':
        return RDBI_WDBI_COLUMN_MAP[letter]
    elif sheet_type == 'routine_control':
        return ROUTINE_CONTROL_COLUMN_MAP[letter]
    elif sheet_type == 'did_library':
        return DID_LIBRARY_COLUMN_MAP[letter]
    elif sheet_type == 'rid_library':
        return RID_LIBRARY_COLUMN_MAP[letter]
    else:
        raise ValueError(f"Unknown sheet type: {sheet_type}")

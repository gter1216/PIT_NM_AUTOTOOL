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

# BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# SOURCE_FILE = os.path.join(BASE_DIR, "data", "input", "CETUS_RL105", "NT3_NIO_CETUS_ZONE_FTM_0.8.0.02.00.xlsx")
# TARGET_FILE = os.path.join(BASE_DIR, "data", "output", "CETUS_RL105", "测试参数表-1-ZONE_FTM.xlsx")
# OUTPUT_FILE = os.path.join(BASE_DIR, "data", "output", "CETUS_RL105", "测试参数表-1-ZONE_FTM_更新后.xlsx")
# SOURCE_SHEET1 = "3.1Basic DIDs"
# SOURCE_SHEET2 = "3.2RDBI 0x22 & WDBI 0x2E"
# TARGET_SHEET = "22_2E_DID_Library"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CURR_REL_DIR = os.path.join(BASE_DIR, "data", "CETUS_RL201", "CURR_REL")
LAST_REL_DIR = os.path.join(BASE_DIR, "data", "CETUS_RL201", "LAST_REL")
OUTPUT_DIR = os.path.join(BASE_DIR, "data", "CETUS_RL201", "OUTPUT")
SOURCE_SHEET1 = "3.1Basic DIDs"
SOURCE_SHEET2 = "3.2RDBI 0x22 & WDBI 0x2E"
TARGET_SHEET = "22_2E_DID_Library"
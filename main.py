#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据处理总控制脚本
运行数据拍平和JSON转换两个功能
"""

import sys
from pathlib import Path

# 添加当前目录到Python路径
sys.path.append(str(Path(__file__).parent))

from excel_flatten import ExcelFlattener
from excel_to_json_new import ExcelToJsonNew


def main():
    """主函数 - 运行所有处理流程"""
    print("=" * 50)
    print("Excel数据处理工具")
    print("=" * 50)
    
    try:
        # 创建输出目录
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
        
        print("\n1. 开始数据拍平处理...")
        # 运行数据拍平
        flattener = ExcelFlattener(str(output_dir))
        csv_file = flattener.run()
        
        print("\n2. 开始JSON转换处理...")
        # 运行JSON转换
        converter = ExcelToJsonNew(str(output_dir))
        json_file, report_file = converter.run()
        
        print("\n" + "=" * 50)
        print("所有处理完成！")
        print("=" * 50)
        print(f"输出文件:")
        print(f"  CSV文件: {csv_file}")
        print(f"  JSON文件: {json_file}")
        print(f"  摘要报告: {report_file}")
        print(f"\n输出目录: {output_dir.absolute()}")
        
    except Exception as e:
        print(f"\n处理过程中出错: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从Excel获取JSON结构
Quarter从A4列开始往下获取
"""

import pandas as pd
import json
import numpy as np
from pathlib import Path

class ExcelToJsonNew:
    """从Excel获取JSON结构"""
    
    def __init__(self, output_dir):
        """初始化"""
        # 使用相对路径，确保解压后可直接运行
        self.excel_file = Path("input/second_batch_copy.xlsx")
        self.output_dir = Path(output_dir)
        
        self.output_dir.mkdir(exist_ok=True)
        
        if not self.excel_file.exists():
            raise FileNotFoundError(f"Excel文件不存在: {self.excel_file}")
    
    def load_excel_data(self):
        """加载Excel数据"""
        try:
            # 读取Excel文件，从A4开始，包含所有列
            df = pd.read_excel(self.excel_file, skiprows=3)
            return df
        except Exception as e:
            raise
    
    def get_bg_group_from_b5(self, df):
        """从B5列开始获取BG Group数据"""
        bg_group_values = [v for v in df.iloc[:, 1].dropna().unique() if v != 'Total']
        return bg_group_values
    
    def get_quarter_values(self, df):
        """获取Quarter列的唯一值"""
        quarter_values = [v for v in df.iloc[:, 0].dropna().unique() if v != 'Total']
        return quarter_values
    
    def get_bg_group_values(self, df):
        """获取BG Group列的唯一值"""
        logger.info("=== 获取BG Group列的唯一值 ===")
        
        # 获取B列(BG Group)的唯一值
        bg_group_values = df.iloc[:, 1].dropna().unique()
        logger.info(f"BG Group列的唯一值: {list(bg_group_values)}")
        
        return bg_group_values
    
    def get_hbb_l1_values(self, df):
        """获取HBB L1列的唯一值"""
        hbb_l1_values = [v for v in df.iloc[:, 2].dropna().unique() if v != 'Total']
        return hbb_l1_values
    
    def get_hbb_l2_values(self, df):
        """获取HBB L2列的唯一值"""
        hbb_l2_values = [v for v in df.iloc[:, 3].dropna().unique() if v != 'Total']
        return hbb_l2_values
    
    def get_geo_values(self, df):
        """获取Geo列的唯一值"""
        geo_values = [v for v in df.iloc[:, 4].dropna().unique() if v != 'Total']
        return geo_values
    
    def get_segment_values(self, df):
        """获取Segment列的唯一值"""
        segment_values = [v for v in df.iloc[:, 5].dropna().unique() if v != 'Total']
        return segment_values
    
    def get_fiscal_year_from_first_row(self):
        """从Excel第一行获取Fiscal Year信息"""
        try:
            # 读取Excel文件的第一行（列标题行）
            header_df = pd.read_excel(self.excel_file, nrows=1)
            
            # 获取FY开头的财年或YoY，过滤掉带.的重复列
            fiscal_years = [col for col in header_df.columns 
                          if isinstance(col, str) and 
                          ((col.startswith('FY') and len(col) >= 7 and '.' not in col) or col == 'YoY')]
            
            fiscal_years = sorted(set(fiscal_years))
            return fiscal_years
        except Exception as e:
            return ["FY23/24", "FY24/25", "FY25/26"]

    def get_fiscal_quarter_from_second_row(self):
        """从Excel第二行获取Fiscal Quarter信息"""
        try:
            # 读取Excel文件的第二行
            second_row_df = pd.read_excel(self.excel_file, skiprows=1, nrows=1)
            
            # 获取Q开头的季度，过滤掉带.的重复列
            fiscal_quarters = [col for col in second_row_df.columns 
                             if isinstance(col, str) and col.startswith('Q') and '.' not in col]
            
            # 自定义排序：Q1, Q2, Q3, Q4优先，其他按字母排序
            def sort_key(q):
                return (0, q) if q in ['Q1', 'Q2', 'Q3', 'Q4'] else (1, q)
            
            fiscal_quarters = sorted(set(fiscal_quarters), key=sort_key)
            return fiscal_quarters
        except Exception as e:
            return ["Q1", "Q2", "Q3", "Q4"]
    
    def create_json_structure(self, df):
        """创建JSON结构"""
        # 获取各列的唯一值
        quarter_values = self.get_quarter_values(df)
        bg_group_values = self.get_bg_group_from_b5(df)
        hbb_l1_values = self.get_hbb_l1_values(df)
        hbb_l2_values = self.get_hbb_l2_values(df)
        geo_values = self.get_geo_values(df)
        segment_values = self.get_segment_values(df)
        fiscal_year_values = self.get_fiscal_year_from_first_row()  # 从第一行获取Fiscal Year
        fiscal_quarter_values = self.get_fiscal_quarter_from_second_row()  # 从第二行获取Fiscal Quarter
        
        # 创建JSON结构
        result = {
            "Quarter": {}
        }
        
        for quarter in quarter_values:
            result["Quarter"][str(quarter)] = {}
            
            for bg_group in bg_group_values:
                result["Quarter"][str(quarter)][str(bg_group)] = {}
                
                for hbb_l1 in hbb_l1_values:
                    result["Quarter"][str(quarter)][str(bg_group)][str(hbb_l1)] = {}
                    
                    for hbb_l2 in hbb_l2_values:
                        result["Quarter"][str(quarter)][str(bg_group)][str(hbb_l1)][str(hbb_l2)] = {}
                        
                        for geo in geo_values:
                            result["Quarter"][str(quarter)][str(bg_group)][str(hbb_l1)][str(hbb_l2)][str(geo)] = {}
                            
                            for segment in segment_values:
                                result["Quarter"][str(quarter)][str(bg_group)][str(hbb_l1)][str(hbb_l2)][str(geo)][str(segment)] = {}
                                
                                for fiscal_year in fiscal_year_values:
                                    result["Quarter"][str(quarter)][str(bg_group)][str(hbb_l1)][str(hbb_l2)][str(geo)][str(segment)][str(fiscal_year)] = {}
                                    
                                    # Fiscal Quarter作为第8层（最后一层）
                                    for fiscal_quarter in fiscal_quarter_values:
                                        result["Quarter"][str(quarter)][str(bg_group)][str(hbb_l1)][str(hbb_l2)][str(geo)][str(segment)][str(fiscal_year)][str(fiscal_quarter)] = {}
        
        return result
    
    def save_json(self, data, filename):
        """保存JSON文件"""
        filepath = self.output_dir / filename
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        return filepath
    
    def generate_summary_report(self, df, quarter_values, bg_group_values, hbb_l1_values, hbb_l2_values, geo_values, segment_values, fiscal_year_values, fiscal_quarter_values):
        """生成摘要报告"""
        report_path = self.output_dir / "excel_data_summary.md"
        
        try:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write("# Excel数据摘要报告\n\n")
                f.write(f"**生成时间**: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                f.write("## 数据概览\n")
                f.write(f"- 总行数: {len(df)}\n")
                f.write(f"- 总列数: {len(df.columns)}\n")
                f.write(f"- 数据来源: {self.excel_file.name}\n\n")
                
                f.write("## 各列唯一值统计\n")
                f.write(f"### Quarter (A列): {len(quarter_values)}个唯一值\n")
                for value in quarter_values:
                    f.write(f"- {value}\n")
                
                f.write(f"\n### BG Group (B列): {len(bg_group_values)}个唯一值\n")
                for value in bg_group_values:
                    f.write(f"- {value}\n")
                
                f.write(f"\n### HBB L1 (C列): {len(hbb_l1_values)}个唯一值\n")
                for value in hbb_l1_values:
                    f.write(f"- {value}\n")
                
                f.write(f"\n### HBB L2 (D列): {len(hbb_l2_values)}个唯一值\n")
                for value in hbb_l2_values:
                    f.write(f"- {value}\n")
                
                f.write(f"\n### Geo (E列): {len(geo_values)}个唯一值\n")
                for value in geo_values:
                    f.write(f"- {value}\n")
                
                f.write(f"\n### Segment (F列): {len(segment_values)}个唯一值\n")
                for value in segment_values:
                    f.write(f"- {value}\n")
                
                f.write(f"\n### Fiscal Year (从第一行获取): {len(fiscal_year_values)}个唯一值\n")
                for value in fiscal_year_values:
                    f.write(f"- {value}\n")
                
                f.write(f"\n### Fiscal Quarter (从第二行获取): {len(fiscal_quarter_values)}个唯一值\n")
                for value in fiscal_quarter_values:
                    f.write(f"- {value}\n")
                
                f.write("\n## JSON结构信息\n")
                f.write(f"- JSON层级: 8层嵌套结构\n")
                f.write(f"- 第1层: Quarter ({len(quarter_values)}个值)\n")
                f.write(f"- 第2层: BG Group ({len(bg_group_values)}个值)\n")
                f.write(f"- 第3层: HBB L1 ({len(hbb_l1_values)}个值)\n")
                f.write(f"- 第4层: HBB L2 ({len(hbb_l2_values)}个值)\n")
                f.write(f"- 第5层: Geo ({len(geo_values)}个值)\n")
                f.write(f"- 第6层: Segment ({len(segment_values)}个值)\n")
                f.write(f"- 第7层: Fiscal Year ({len(fiscal_year_values)}个值)\n")
                f.write(f"- 第8层: Fiscal Quarter ({len(fiscal_quarter_values)}个值)\n")
            
            return report_path
        except Exception as e:
            raise

    def run(self):
        """运行JSON转换"""
        try:
            # 加载Excel数据
            df = self.load_excel_data()
            
            # 获取各列的唯一值
            quarter_values = self.get_quarter_values(df)
            bg_group_values = self.get_bg_group_from_b5(df)  # 从B5开始获取BG Group
            hbb_l1_values = self.get_hbb_l1_values(df)
            hbb_l2_values = self.get_hbb_l2_values(df)
            geo_values = self.get_geo_values(df)
            segment_values = self.get_segment_values(df)
            fiscal_year_values = self.get_fiscal_year_from_first_row()  # 从第一行获取Fiscal Year
            
            # 创建JSON结构
            json_data = self.create_json_structure(df)
            
            # 保存JSON文件
            json_file = self.save_json(json_data, "excel_to_json.json")
            
            # 生成摘要报告
            fiscal_quarter_values = self.get_fiscal_quarter_from_second_row()
            report_file = self.generate_summary_report(df, quarter_values, bg_group_values, hbb_l1_values, hbb_l2_values, geo_values, segment_values, fiscal_year_values, fiscal_quarter_values)
            
            print("=== JSON转换完成 ===")
            print(f"JSON文件: {json_file}")
            print(f"摘要报告: {report_file}")
            return json_file, report_file
            
        except Exception as e:
            print(f"处理过程中出错: {e}")
            raise

def main():
    """主函数"""
    converter = ExcelToJsonNew("output")
    converter.run()

if __name__ == "__main__":
    main()

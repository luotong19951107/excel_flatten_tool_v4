#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据拍平工具 - 简化版
将Excel数据转换为扁平化的CSV结构
"""

import pandas as pd
from pathlib import Path


class ExcelFlattener:
    """Excel数据拍平器"""
    
    def __init__(self, output_dir):
        """初始化路径"""
        # 使用相对路径，确保解压后可直接运行
        self.excel_file = Path("input/second_batch_copy.xlsx")
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
    
    def load_excel_data(self):
        """加载Excel数据，丢弃K、P、U列以及V列之后的所有列"""
        df = pd.read_excel(self.excel_file, skiprows=3)
        
        # 丢弃指定列
        def col_to_index(col_letter):
            return sum((ord(c) - ord('A') + 1) * (26 ** i) for i, c in enumerate(reversed(col_letter))) - 1
        
        drop_cols = [col_to_index(c) for c in ['K', 'P', 'U']]
        cutoff = col_to_index('V')
        keep_cols = [i for i in range(df.shape[1]) if i not in drop_cols and i < cutoff]
        
        df = df.iloc[:, keep_cols]
        print(f"成功加载Excel数据，共{len(df)}行，保留{df.shape[1]}列")
        return df
    
    def get_fiscal_info(self):
        """获取Fiscal Year和Fiscal Quarter映射"""
        # Fiscal Year映射
        fiscal_years = ['FY23/24', 'FY24/25', 'FY25/26', 'YoY']
        
        # Fiscal Quarter映射
        fiscal_quarters_map = {
            'FY23/24': ['Q1', 'Q2', 'Q3', 'Q4'],
            'FY24/25': ['Q1', 'Q2', 'Q3', 'Q4'],
            'FY25/26': ['Q1ACT', 'Q2ACT', 'Q3M1', 'Q4MT'],
            'YoY': ['Q1', 'Q2', 'Q3', 'Q4']
        }
        
        return fiscal_years, fiscal_quarters_map
    
    def flatten_data(self, df):
        """将Excel数据拍平为CSV格式"""
        df_clean = df.dropna(how='all')
        fiscal_years, fiscal_quarters_map = self.get_fiscal_info()
        result_data = []
        
        for _, row in df_clean.iterrows():
            # 提取前6列层级信息
            quarter = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
            bg_group = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
            hbb_l1 = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
            hbb_l2 = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
            geo = str(row.iloc[4]) if pd.notna(row.iloc[4]) else "NA"
            segment = str(row.iloc[5]) if pd.notna(row.iloc[5]) else ""
            
            # 检查数据有效性
            if not all([quarter, bg_group, hbb_l1, hbb_l2, geo, segment]):
                continue
            if "Total" in [quarter, bg_group, hbb_l1, hbb_l2, geo, segment]:
                continue
            
            # 处理数据列（每4个值为一组）
            data_columns = row.iloc[6:].dropna()
            if len(data_columns) < 4:
                continue
            
            # 为每个Fiscal Year创建4条记录
            for i in range(0, len(data_columns), 4):
                if i + 3 >= len(data_columns):
                    break
                
                fiscal_year_idx = i // 4
                if fiscal_year_idx >= len(fiscal_years):
                    break
                
                fiscal_year = fiscal_years[fiscal_year_idx]
                fiscal_quarters = fiscal_quarters_map.get(fiscal_year, ['Q1', 'Q2', 'Q3', 'Q4'])
                
                # 创建4条记录
                for q_idx in range(4):
                    if i + q_idx >= len(data_columns) or q_idx >= len(fiscal_quarters):
                        break
                    
                    result_data.append({
                        'Quarter': quarter,
                        'BG Group': bg_group,
                        'HBB L1': hbb_l1,
                        'HBB L2': hbb_l2,
                        'Geo': geo,
                        'Segment': segment,
                        'Fiscal Year': fiscal_year,
                        'Fiscal Quarter': fiscal_quarters[q_idx],
                        'Final': str(data_columns.iloc[i + q_idx])
                    })
        
        result_df = pd.DataFrame(result_data)
        print(f"拍平后数据条数: {len(result_df)}")
        return result_df
    
    def save_csv(self, data, filename="flattened_data.csv"):
        """保存为CSV文件"""
        filepath = self.output_dir / filename
        data.to_csv(filepath, index=False, encoding='utf-8-sig')
        file_size = filepath.stat().st_size
        print(f"CSV文件已保存: {filepath}")
        print(f"文件大小: {file_size / 1024:.2f} KB")
        return filepath
    
    def run(self):
        """运行数据拍平"""
        print("=== Excel数据拍平工具 ===")
        
        # 加载数据
        df = self.load_excel_data()
        
        # 拍平数据
        flattened_data = self.flatten_data(df)
        
        # 保存CSV
        csv_file = self.save_csv(flattened_data)
        
        print("数据拍平完成！")
        return csv_file


if __name__ == "__main__":
    flattener = ExcelFlattener("output")
    flattener.run()

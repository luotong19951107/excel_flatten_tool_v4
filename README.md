# Excel数据处理工具

## 功能
- **数据拍平**: 将Excel数据转换为扁平化的CSV结构
- **JSON转换**: 从Excel文件生成8层嵌套JSON结构和数据摘要报告

## 输入
Excel文件：`second_batch_copy.xlsx`（与code目录同级）

## 输出
运行后会在`output`目录生成3个文件：

### 1. CSV文件 (`flattened_data.csv`)
扁平化的数据表格，包含9列：
- Quarter, BG Group, HBB L1, HBB L2, Geo, Segment, Fiscal Year, Fiscal Quarter, Final
- Geo列包含4个值：'AP', 'EMEA', 'NA', 'LA'
- Fiscal Quarter映射：
  - FY23/24, FY24/25: Q1, Q2, Q3, Q4
  - FY25/26: Q1ACT, Q2ACT, Q3M1, Q4MT

### 2. JSON文件 (`excel_to_json.json`)
8层嵌套JSON结构：
```json
{
  "Quarter": {              ← 第1层：Quarter
    "Revenue $M": {         ← 第2层：Quarter值
      "IDG": {              ← 第3层：BG Group
        "GPS": {            ← 第4层：HBB L1
          "Support Services": { ← 第5层：HBB L2
            "AP": {         ← 第6层：Geo
              "REL": {      ← 第7层：Segment
                "FY23/24": { ← 第8层：Fiscal Year
                  "Q1": {}, ← 第9层：Fiscal Quarter
                  "Q2": {},
                  ...
                }
              }
            }
          }
        }
      }
    }
  }
}
```

### 3. 摘要报告 (`excel_data_summary.md`)
数据统计和结构说明

## 运行方式

### 方式1：运行总控制脚本（推荐）
```bash
cd code
python main.py
```

### 方式2：单独运行
```bash
cd code
# 只运行数据拍平
python excel_flatten.py

# 只运行JSON转换
python excel_to_json_new.py
```

## 文件结构
```
code/
├── main.py                    ← 总控制脚本
├── excel_flatten.py          ← 数据拍平工具
├── excel_to_json_new.py      ← JSON转换工具
├── README.md                  ← 说明文档
└── output/                    ← 输出目录
    ├── flattened_data.csv     ← CSV文件
    ├── excel_to_json.json    ← JSON文件
    └── excel_data_summary.md ← 摘要报告
```

## 层级结构
Quarter → BG Group → HBB L1 → HBB L2 → Geo → Segment → Fiscal Year → Fiscal Quarter

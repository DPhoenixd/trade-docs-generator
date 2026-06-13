# 外贸三件套生成工具

本地 Streamlit 工具，用于根据客户订单 PDF、面料数据库、PI/CI 模板和 Packing List 模板生成外贸单据。

## 功能

- P.I / C.I 生成
- Packing List 生成
- 面料英文数据库匹配
- 公司码单截图识别和多截图合并识别
- KG / Meter / Yard 换算
- Gross Weight 计算

## 本地启动

```powershell
cd "D:\DyrusWok\外贸\外贸三件套生成工具"
.\.venv\Scripts\python.exe -m streamlit run app.py
```

浏览器打开：

```text
http://localhost:8501
```

如果是新电脑，先安装依赖：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

## 重要文件

- `app.py`：Streamlit 主程序
- `trade_docs/`：解析、计算、Excel 写入逻辑
- `fabric_master_en.csv` / `fabric_price_rules.csv` / `fabric_database_en.xlsx`：面料数据库和价格规则
- `CI-POUT25VE0011181A-25A109A.xlsx`：PI/CI 模板
- `Packing List-POUT25VE0011181A-25A109A.xlsx`：Packing List 模板

## 不上传到 GitHub 的内容

- 客户订单 PDF
- 码单图片
- 生成的 Excel 输出
- 本地虚拟环境 `.venv`
- Streamlit 临时上传缓存

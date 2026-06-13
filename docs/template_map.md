# 外贸三件套模板结构记录

## CI / PI 模板

文件：`CI-POUT25VE0011181A-25A109A.xlsx`

- Sheet：`Sheet2`
- 公司抬头：`A1:A2`
- 单据标题：`A4:E4`，当前为 `COMMERCIAL INVOICE`
- 单号：`F5`，模板格式为 `P/I NO:D2026042401`
- 日期：`F6`，模板格式为 `DATE: 2026/04/24`
- Seller：`A9:E9`
- Buyer：`A11:E11`，地址在 `A12:E12`
- 明细表头：`A16:F16`
- 明细行：`17:20`
  - `A`：`PO NO. / code / ART NO. / Knit fabric`
  - `B`：`ITEM / COMP / WEIGHT / WIDTH / FINISH / COUNTRY OF ORIGIN`
  - `C`：Color
  - `D`：Quantity KG
  - `E`：USD/KG
  - `F`：Amount，公式为 `=D*E`
- 预付款：`A21:C21` 和 `D21:F21`
- 尾款：`A22:B22` 和 `C22:F22`
- 付款/包装/交期/运输/港口/溢短装条款：`A26:A31`
- 收款信息：`A34:E34`，必须保留模板原文。

## Packing List 模板

文件：`Packing List-POUT25VE0011181A-25A109A.xlsx`

- Sheet：`Sheet1`
- 公司抬头：`H1:H2`
- 标题：`K4`
- TO/Buyer：`G5:O5`
- P/I NO：`P5:X5`
- 描述：`G6:O7`
- 日期：`P6:X7`
- 细码区表头：`A9:X9`
- 细码区最多 4 个横向分组，每组最多 28 条：
  - 分组 1：`A:F`，数据 `10:37`，合计 `38`，Gross `D40`
  - 分组 2：`G:L`，数据 `10:37`，合计 `38`，Gross `J40`
  - 分组 3：`M:R`，数据 `10:37`，合计 `38`，Gross `P40`
  - 分组 4：`S:X`，数据 `10:37`，合计 `38`，Gross `V40`
- 右侧金额汇总：`Z5:AE9`
  - `Z`：Color
  - `AA`：Net Weight KG
  - `AB`：USD/KG
  - `AC`：Amount
  - `AD6`：Advance payment
  - `AE6`：Remaining payable amount
- 收款信息：`AA12:AE25`，必须保留模板原文。
- Gross Weight 汇总：
  - `AB39`：ART NO
  - `AC39`：Fabric Code
  - `AB40`：各分组 Gross Weight 合计
  - `AD40`：Total Gross Weight

## 第一版写入策略

- 只写入 Packing List 模板中的可变字段，不移动模板结构。
- 保留收款信息区域原样。
- 细码行使用 Excel 公式计算 Meter 和 Yard。
- Gross Weight 使用公式：`Total Net Weight + tube_plus_allowance_kg_per_roll * Rolls`。
- 当前模板上限是 4 个颜色/LOT 分组，每组 28 条；超过时生成前报错。

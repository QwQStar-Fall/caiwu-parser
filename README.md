# caiwu-parser · A 股财报三大报表高鲁棒性解析器

[![Python 3.8+](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](./LICENSE)
[![Online Lab](https://img.shields.io/badge/%E5%9C%A8%E7%BA%BF%E5%AE%8C%E6%95%B4%E7%89%88-%E8%B4%A2%E6%8A%A5%E5%AE%9E%E9%AA%8C%E5%AE%A4-orange)](https://star-fall.159911.xyz/lab/)

**`caiwu-parser`** 是从 [财报实验室（Financial Report Lab）](https://star-fall.159911.xyz/lab/) 抽离开源的底层核心解析引擎。它不依赖第三方金融数据接口，直接从上市公司披露的**原生年报 PDF、Excel（`.xlsx`）与 Word（`.docx`）** 文件中精准提取合并三大财务报表（合并资产负债表、合并利润表、合并现金流量表），统一金额量纲为「元」，并内置**会计恒等式门禁校验**。

> Lightweight, battle-tested Python parser for Chinese A-share financial statements (Consolidated Balance Sheet, Income Statement, and Cash Flow Statement) across native PDF, Excel, and Word documents — with built-in accounting identity verification.

---

## 为什么需要专门的年报解析器？

直接用 `pdfplumber` 或通用表格工具抽取 A 股两三百页的年报 PDF，几乎必然踩中以下真实版式陷阱。`caiwu-parser` 在百余家不同行业上市公司年报上逐一打磨了工程兜底：

| 真实年报常见痛点 | `caiwu-parser` 的处理策略 |
| :--- | :--- |
| **整本年报 200~400 页，全量抽表极慢** | 先用 `pypdfium2`（约 30 倍速）毫秒级定位三大合并报表页，再仅对目标页调用 `pdfplumber` 提取表格 |
| **合并报表与母公司报表同页交界** | 精确检测页内下一张表标题的纵向坐标（`top`）并自动裁切，杜绝母公司数字混入合并口径 |
| **跨页续表丢失表头或列宽错位** | 自动继承首页年份列映射与金额单位；续页合并单元格拆法变化时按数值列顺序自适应重对齐 |
| **左右分栏资产负债表 & 无框线表格** | 表格层与文本行双轨解析，自动拆分左右并排的「资产 / 负债权益」行，支持两栏坐标聚类 |
| **奇葩颠倒排版 & 折行科目** | 自动识别「数字在上一行、科目名在下一行」（如部分快递/制造龙头年报）及长科目名跨行折断 |
| **PDF 字体映射损坏（乱码错字）** | 内置字形修复映射（如 `合幵→合并`、`股仹→股份`、`归属二母公司→归属于母公司`）与港式括号负数 `(1,234)` 符号归一 |
| **金额单位遗漏或被正文干扰** | 综合识别页眉/表头 `元 / 千元 / 万元 / 百万元 / 亿元`，并与前文「主要会计数据」营业收入交叉比对，自动修正 10³ / 10⁴ / 10⁶ 倍量纲偏差 |
| **银行 / 保险 / 证券科目体系迥异** | 自动识别金融机构类型，切换专属科目映射（吸收存款、发放贷款、保险服务收入、代理买卖证券款等）与金融四重校验 |
| **怎么确信抽出来的数字没串行？** | 内置**三大会计恒等式硬门禁**，差额超阈值立即标记 `FAIL`，拒绝静默产出错数 |

---

## 安装

```bash
git clone https://github.com/QwQStar-Fall/caiwu-parser.git
cd caiwu-parser
pip install .
```

或仅安装依赖直接运行源码：

```bash
pip install -r requirements.txt
```

---

## 快速上手

### 1. Python API（5 行代码提取 + 校验）

```python
from caiwu_parser import parse_report, pivot_report

res = parse_report("同仁堂_2025年报.pdf")

print("门禁校验是否全部通过:", res["is_valid"])
for rule, ok, diff in res["checks"]:
    print(f"  [{'PASS' if ok else 'FAIL'}] {rule} (差额: {diff:.2f} 元)")

# 转为「标准项目 × 年度」透视表（单位缩放为亿元，保留 2 位小数）
df_yi = pivot_report(res, unit=1e8, decimals=2)
print(df_yi)
```

### 2. 命令行一键查看或导出 CSV

```bash
# 终端打印核心科目透视表（默认单位：亿元）与恒等式校验结果
python -m caiwu_parser 同仁堂_2025年报.pdf

# 以万元为单位查看，并将标准长表导出为 CSV
python -m caiwu_parser 同仁堂_2025年报.pdf --unit wan --csv tongrentang.csv
```

终端输出示例：

```text
=== 同仁堂 (一般工商业) ===
  · 资产负债表: 文本兜底补充 2 条
  · 资产负债表: 定位 3 页, 提取 46 条
  · 利润表: 定位 3 页, 提取 48 条
  · 现金流量表: 定位 2 页, 提取 32 条

--- 核心报表科目（单位：亿元）---
年度            2024    2025
标准项目
总资产         311.98  306.21
流动资产        236.96  234.03
货币资金        107.56  108.78
...
营业收入        185.97  172.56
净利润          22.80   16.33
归母净利润        15.26   11.89
经营现金流净额       7.61   26.93

--- 会计恒等式门禁校验 ---
  [PASS] 2024 资产 = 负债 + 权益  (差额: 0.00 元)
  [PASS] 2024 净利润 = 利润总额 - 所得税  (差额: 0.00 元)
  [PASS] 2024 利润总额 = 营业利润 + 营业外收支净额  (差额: 0.00 元)
  [PASS] 2025 资产 = 负债 + 权益  (差额: -0.00 元)
  [PASS] 2025 净利润 = 利润总额 - 所得税  (差额: 0.00 元)
  [PASS] 2025 利润总额 = 营业利润 + 营业外收支净额  (差额: -0.00 元)

整体门禁结果: 通过 (VALID)
```

---

## 返回数据结构说明

`parse_report(file_path, company=None, is_fin=None)` 返回字典：

| 字段 | 类型 | 说明 |
| :--- | :--- | :--- |
| `company` | `str` | 公司名称（未指定时自动从文件名提取） |
| `is_financial` | `bool` | 是否识别为金融企业（银行 / 保险 / 证券） |
| `data` | `pandas.DataFrame` | 标准长表，固定四列：`["公司", "年度", "标准项目", "金额"]`，金额统一折算为**人民币元** |
| `checks` | `list[tuple]` | 会计恒等式校验明细：`[(规则说明, bool是否通过, float差额元), ...]` |
| `is_valid` | `bool` | 所有会计恒等式是否全部通过（建议下游入库前先断言 `res["is_valid"]`） |
| `logs` | `list[str]` | 各报表定位页数、提取条数、文本兜底及单位校正日志 |

### 内置会计恒等式校验规则

1. **资产负债表配平**：`总资产 = 总负债 + 所有者权益`
2. **利润表底线勾稽**：`净利润 = 利润总额 - 所得税费用`
3. **营业利润桥接**：`利润总额 = 营业利润 + 营业外收入 - 营业外支出`
4. **金融企业附加校验**（银行 / 证券）：`营业收入 ≈ 各项净收入构成之和`

---

## 开源范围与在线完整版（Open Core）

本仓库采用 **MIT 开源协议**，专注于把「**单份原生财报 → 干净可信的标准三表 DataFrame**」这一件基础工作做到极致轻量、零外部服务依赖。

如果您在投研或学习中需要更完整的深度分析能力，欢迎直接免费体验在线完整版 **[财报实验室（Financial Report Lab）](https://star-fall.159911.xyz/lab/)**（工作台直达：[lab.159911.xyz](https://lab.159911.xyz)）：

| 功能模块 | `caiwu-parser`（本开源库） | [财报实验室（在线完整版）](https://lab.159911.xyz) |
| :--- | :---: | :---: |
| 原生 PDF / Excel / Word 三大报表提取 | ✅ | ✅ |
| 银行 / 证券 / 保险专属科目解析 | ✅ | ✅ |
| 三大会计恒等式自动门禁校验 | ✅ | ✅ |
| **数字溯源（点击任意指标弹出 PDF 原页并红框高亮对应行）** | — | ✅ |
| **纯图片扫描件 PDF 视觉转录识别** | — | ✅ |
| **31 个二级细分行业自动归类 + 主营构成拆解** | — | ✅ |
| **130+ 家代表性公司 × 6 年同业横比与多年趋势库** | — | ✅ |
| **杜邦拆解 + Altman Z / Piotroski F 评分 + Word 诊断报告导出** | — | ✅ |

---

## 致谢与交流

- 项目主页与技术演进笔记：[https://star-fall.159911.xyz/lab/](https://star-fall.159911.xyz/lab/)
- 如果遇到特殊排版的 A 股年报未能通过恒等式门禁，欢迎提交 Issue 并附上股票代码与报告年度，帮助解析规则持续进化。

## License

[MIT License](./LICENSE) © 2026 Star-Fall

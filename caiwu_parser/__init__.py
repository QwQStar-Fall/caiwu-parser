# -*- coding: utf-8 -*-
"""
caiwu-parser: 轻量高鲁棒性的 A 股财报三大报表解析器（支持 PDF / Excel / Word）
"""
import os
import pandas as pd
from . import parse_statements as ps
from . import parse_bank as pb
from . import pdftext

__version__ = "0.1.0"

_FIN_NAME_HINTS = (
    "银行", "证券", "保险", "信托", "期货", "券商",
    "人寿", "太保", "中国平安", "新华保险", "海通", "中金公司", "申万宏源", "国泰君安", "中信建投",
)

_FIN_COVER_HINTS = (
    "银行股份有限公司", "保险股份有限公司", "保险（集团）股份有限公司", "保险(集团)股份有限公司",
    "证券股份有限公司", "人寿保险", "财产保险",
)

# 透视表展示时的标准科目优先顺序（资产负债表 → 利润表 → 现金流量表 → 其余科目）
_PREFERRED_ORDER = [
    # 资产负债表
    "总资产", "流动资产", "货币资金", "应收账款", "存货", "合同资产",
    "非流动资产", "固定资产净额", "在建工程",
    "总负债", "流动负债", "短期借款", "应付账款", "合同负债", "预收款项", "一年内到期非流动负债",
    "非流动负债", "长期借款", "负债和权益总计",
    "所有者权益", "归母权益", "少数股东权益", "股本", "盈余公积", "未分配利润",
    # 金融业核心资产负债
    "发放贷款和垫款", "吸收存款", "代理买卖证券款", "融出资金", "保险合同负债",
    # 利润表
    "营业总收入", "营业收入", "利息净收入", "手续费及佣金净收入", "保险服务收入", "已赚保费",
    "营业总成本", "营业成本", "营业税金及附加",
    "销售费用", "管理费用", "研发费用", "财务费用", "利息费用", "业务及管理费",
    "其他收益", "投资收益", "公允价值变动收益", "信用减值损失", "资产减值损失", "资产处置收益",
    "营业利润", "营业外收入", "营业外支出", "利润总额", "所得税费用",
    "净利润", "归母净利润", "少数股东损益",
    # 现金流量表
    "经营现金流入", "销售商品收现", "经营现金流出", "购买商品付现", "职工薪酬现金", "经营现金流净额",
    "投资现金流入", "投资现金流出", "投资现金流净额",
    "筹资现金流入", "筹资现金流出", "筹资现金流净额",
    "现金净增加额", "期初现金余额", "期末现金余额",
]


def is_financial(company_or_filename="", file_path=None):
    """判断是否属于金融类机构（银行 / 证券 / 保险）。"""
    name = str(company_or_filename or "")
    if any(k in name for k in _FIN_NAME_HINTS):
        return True
    if file_path and os.path.isfile(file_path) and file_path.lower().endswith(".pdf"):
        pgs = pdftext.pages(file_path) or []
        cover = "".join(t for _, t in pgs[:4]).replace(" ", "").replace("\n", "")
        if any(k in cover for k in _FIN_COVER_HINTS):
            return True
    return False


def parse_report(file_path, company=None, is_fin=None):
    """
    解析单份 A 股财务报告文件（支持原生 PDF / Excel .xlsx / Word .docx）。

    参数:
        file_path (str): 报表文件路径
        company (str, optional): 公司名称；若不传则自动从文件名提取
        is_fin (bool, optional): 是否为金融企业（银行/证券/保险），默认自动识别

    返回:
        dict: {
            "company": str,               # 公司名称
            "is_financial": bool,         # 是否按金融企业准则解析
            "data": pandas.DataFrame,     # 标准列: ["公司", "年度", "标准项目", "金额"]（单位统一为元）
            "checks": list[tuple],        # 恒等式校验结果: [(规则说明, 是否通过, 差额), ...]
            "is_valid": bool,             # 是否全部通过会计恒等式校验
            "logs": list[str],            # 解析过程与单位校正日志
        }
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    base = os.path.splitext(os.path.basename(file_path))[0]
    co = company or base.split("_")[0].split("：")[0].split(":")[0].strip()

    if is_fin is None:
        is_fin = is_financial(co, file_path=file_path) or is_financial(base)

    if is_fin:
        df, logs = pb.parse(file_path, company=co)
        df = pb.finish(df)
        checks = pb.validate(df)
    else:
        df, logs = ps.parse(file_path, company=co)
        df = ps.finish(df)
        checks = ps.validate(df)

    if df.empty and file_path.lower().endswith(".pdf"):
        pgs = pdftext.pages(file_path)
        if pgs and sum(n for n, _ in pgs[:min(20, len(pgs))]) < 200:
            logs.append("提示: 该 PDF 疑似纯图片扫描件（无原生文字层），轻量开源引擎仅解析原生 PDF / Excel / Word")

    # 确保输出列严格为干净的四列
    std_cols = ["公司", "年度", "标准项目", "金额"]
    if not df.empty:
        df = df[[c for c in std_cols if c in df.columns]].reset_index(drop=True)
    else:
        df = pd.DataFrame(columns=std_cols)

    is_valid = bool(checks) and all(c[1] for c in checks)

    return {
        "company": co,
        "is_financial": bool(is_fin),
        "data": df,
        "checks": checks,
        "is_valid": is_valid,
        "logs": logs,
    }


def pivot_report(res_or_df, unit=1.0, decimals=None):
    """
    将解析结果转为「标准项目 × 年度」的透视表，按三大报表阅读顺序排列。

    参数:
        res_or_df: parse_report() 返回的字典，或其 "data" DataFrame
        unit (float): 金额缩放单位（默认 1.0 即元；传 1e8 即亿元，传 1e4 即万元）
        decimals (int, optional): 保留小数位数（例如 decimals=2）

    返回:
        pandas.DataFrame
    """
    df = res_or_df["data"] if isinstance(res_or_df, dict) else res_or_df
    if df is None or df.empty:
        return pd.DataFrame()
    piv = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
    piv = piv.reindex(columns=sorted(piv.columns))
    ordered = [k for k in _PREFERRED_ORDER if k in piv.index]
    rest = [k for k in piv.index if k not in ordered]
    piv = piv.loc[ordered + rest]
    if unit and unit != 1.0:
        piv = piv / float(unit)
    if decimals is not None:
        piv = piv.round(decimals)
    return piv


__all__ = ["parse_report", "pivot_report", "is_financial", "__version__"]

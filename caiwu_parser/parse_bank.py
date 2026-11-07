# -*- coding: utf-8 -*-
"""
parse_bank.py —— 金融业报表解析（银行 / 保险 / 证券）

三类金融公司的报表结构各不相同：
  · 银行：吸收存款、发放贷款和垫款、利息净收入
  · 保险：保险合同负债、保险服务收入、保险服务费用
  · 证券：代理买卖证券款、融出资金、手续费及佣金净收入

共同点是和制造业完全不同（没有存货、没有营业成本），所以单独一套：
  1. 自动识别类型（按特征科目）
  2. 按类型取科目（含「科目一行、数字下一行」的版式）
  3. 四条恒等式校验：资产=负债+权益 / 净利=利润总额-所得税 /
     利润总额=营业利润+营业外收支 / 营业收入≈各项收入之和
  4. 产出与 parse_statements 相同结构的 DataFrame

用法：
    from caiwu_parser.parse_bank import parse, finish, validate
    df, rep = parse("年报.pdf", company="中国平安")   # 自动识别类型
    df, rep = parse("年报.pdf", kind="insurance")     # 强制类型
"""
import os
import re
import sys

import pdfplumber
import pandas as pd

from .parse_statements import (  # noqa: E402
    BASE_COLS, UNIT_MULT, clean, detect_report_year, locate, page_unit, tonum,
)
from . import pdftext  # noqa: E402

# ---------------------------------------------------------------- 科目表

_BALANCE = {
    "资产总计": "总资产", "资产合计": "总资产", "资产总额": "总资产",
    "负债合计": "总负债", "负债总额": "总负债", "负债总计": "总负债",
    "归属于本行股东权益合计": "归母权益", "归属于本公司股东权益合计": "归母权益",
    "股东权益合计": "所有者权益", "所有者权益合计": "所有者权益",
    "归属于母公司股东权益合计": "归母权益",
    "归属于母公司所有者权益合计": "归母权益",
    "少数股东权益": "少数股东权益",
}
_INCOME = {
    "营业收入": "营业收入", "营业总收入": "营业收入", "营业收入合计": "营业收入",
    "营业利润": "营业利润", "利润总额": "利润总额",
    "营业外收入": "营业外收入", "营业外支出": "营业外支出",
    "所得税费用": "所得税费用", "净利润": "净利润",
    "归属于母公司股东的净利润": "归母净利润",
    "归属于母公司所有者的净利润": "归母净利润",
    "税金及附加": "营业税金及附加", "营业税金及附加": "营业税金及附加",
}
_CASH = {
    "经营活动现金流入小计": "经营现金流入", "经营活动现金流出小计": "经营现金流出",
    "投资活动现金流入小计": "投资现金流入", "投资活动现金流出小计": "投资现金流出",
    "筹资活动现金流入小计": "筹资现金流入", "筹资活动现金流出小计": "筹资现金流出",
    "经营活动产生的现金流量净额": "经营现金流净额",
    "经营活动使用的现金流量净额": "经营现金流净额",
    "经营活动产生(使用)的现金流量净额": "经营现金流净额",
    "投资活动产生的现金流量净额": "投资现金流净额",
    "筹资活动产生的现金流量净额": "筹资现金流净额",
    "销售商品、提供劳务收到的现金": "销售商品收现",
    "期末现金及现金等价物余额": "期末现金余额",
    "期初现金及现金等价物余额": "期初现金余额",
    "现金及现金等价物净增加额": "现金净增加额",
}

BANK_EXTRA = {
    "发放贷款和垫款": "发放贷款和垫款", "吸收存款": "吸收存款",
    "贷款损失准备": "贷款损失准备", "存放中央银行款项": "存放中央银行款项",
    "存放同业款项": "存放同业款项", "拆出资金": "拆出资金", "拆入资金": "拆入资金",
    "买入返售金融资产": "买入返售金融资产", "卖出回购金融资产款": "卖出回购金融资产款",
    "向中央银行借款": "向中央银行借款", "同业及其他金融机构存放款项": "同业存放款项",
    "交易性金融资产": "交易性金融资产", "债权投资": "债权投资",
    "其他债权投资": "其他债权投资", "固定资产": "固定资产净额",
    "利息净收入": "利息净收入", "利息收入": "利息收入", "利息支出": "利息支出",
    "手续费及佣金净收入": "手续费及佣金净收入",
    "手续费及佣金收入": "手续费及佣金收入", "手续费及佣金支出": "手续费及佣金支出",
    "业务及管理费": "业务及管理费", "信用减值损失": "信用减值损失",
    "资产减值损失": "资产减值损失", "投资收益": "投资收益",
    "公允价值变动收益": "公允价值变动收益", "汇兑收益": "汇兑收益",
    "公允价值变动损益": "公允价值变动收益", "汇兑损益": "汇兑收益", "投资损益": "投资收益",
    "资产处置损益": "资产处置收益", "资产处置收益": "资产处置收益",
    "其他收益": "其他收益", "其他业务收入": "其他业务收入",
}

INSURANCE_EXTRA = {
    "保险合同负债": "保险合同负债", "保险合同资产": "保险合同资产",
    "保户储金及投资款": "保户储金及投资款", "应收保费": "应收保费",
    "保险服务收入": "保险服务收入", "保险服务费用": "保险服务费用",
    "已赚保费": "已赚保费", "赔付支出": "赔付支出",
    "提取保险责任准备金": "提取保险责任准备金",
    "手续费及佣金支出": "手续费及佣金支出",
    "业务及管理费": "业务及管理费", "利息净收入": "利息净收入",
    "利息收入": "利息收入", "利息支出": "利息支出",
    "投资收益": "投资收益", "公允价值变动收益": "公允价值变动收益",
    "其他收益": "其他收益", "信用减值损失": "信用减值损失",
    "资产减值损失": "资产减值损失", "汇兑收益": "汇兑收益",
    "交易性金融资产": "交易性金融资产", "债权投资": "债权投资",
    "其他债权投资": "其他债权投资", "货币资金": "货币资金",
    "买入返售金融资产": "买入返售金融资产", "应付债券": "应付债券",
    "固定资产": "固定资产净额",
}

SECURITIES_EXTRA = {
    "货币资金": "货币资金", "结算备付金": "结算备付金", "融出资金": "融出资金",
    "存出保证金": "存出保证金", "代理买卖证券款": "代理买卖证券款",
    "代理承销证券款": "代理承销证券款", "卖出回购金融资产款": "卖出回购金融资产款",
    "应付短期融资款": "应付短期融资款", "应付债券": "应付债券",
    "交易性金融资产": "交易性金融资产", "债权投资": "债权投资",
    "其他债权投资": "其他债权投资", "买入返售金融资产": "买入返售金融资产",
    "固定资产": "固定资产净额", "长期股权投资": "长期股权投资",
    "手续费及佣金净收入": "手续费及佣金净收入",
    "经纪业务手续费净收入": "经纪业务净收入",
    "投资银行业务手续费净收入": "投行净收入",
    "资产管理业务手续费净收入": "资管净收入",
    "利息净收入": "利息净收入", "利息收入": "利息收入", "利息支出": "利息支出",
    "投资收益": "投资收益", "投资损益": "投资收益",
    "公允价值变动收益": "公允价值变动收益", "公允价值变动损益": "公允价值变动收益",
    "汇兑收益": "汇兑收益", "汇兑损益": "汇兑收益",
    "资产处置收益": "资产处置收益", "资产处置损益": "资产处置收益",
    "其他收益": "其他收益", "其他业务收入": "其他业务收入",
    "业务及管理费": "业务及管理费",
    "信用减值损失": "信用减值损失", "资产减值损失": "资产减值损失",
}

# 收入拆分的科目（不同金融类型的「营业收入」构成不同）
REV_PARTS = {
    "bank": ("利息净收入", "手续费及佣金净收入", "投资收益", "公允价值变动收益",
             "汇兑收益", "其他收益", "其他业务收入", "资产处置收益"),
    "insurance": ("保险服务收入", "已赚保费", "利息净收入", "投资收益",
                  "公允价值变动收益", "其他收益"),
    "securities": ("手续费及佣金净收入", "利息净收入", "投资收益",
                   "公允价值变动收益", "汇兑收益", "其他收益", "其他业务收入", "资产处置收益"),
}

# 类型特征科目（用于自动识别）
SIGNATURES = {
    "insurance": ("保险服务收入", "保险合同负债", "已赚保费", "保户储金及投资款"),
    "securities": ("代理买卖证券款", "融出资金", "结算备付金", "存出保证金"),
    "bank": ("吸收存款", "发放贷款和垫款", "向中央银行借款", "客户存款", "客户贷款"),
}


def _meta(kind):
    extra = {"bank": BANK_EXTRA, "insurance": INSURANCE_EXTRA, "securities": SECURITIES_EXTRA}[kind]
    alias = {}
    alias.update(_BALANCE)
    alias.update(_INCOME)
    alias.update(_CASH)
    alias.update(extra)
    return alias, set(alias.values())


# ---------------------------------------------------------------- 工具

_NUM = re.compile(r"\(?-?\d[\d,]*(?:\.\d+)?\)?")


def _unit_mult(text, prev=1):
    """页面声明的金额单位；续表页没有声明时沿用上一页"""
    return page_unit(text) or prev


def _trim_note_numbers(vals):
    """去掉行首的附注编号。只丢「小于 100 的整数」——
    否则会把「营业外收入 691」这种小金额一起丢掉，导致取到上一年的值。"""
    def small(v):
        # 附注编号可能是「四(42)」这种整数，也可能是「/ 43.1」这种小数，都算附注号
        return 0 < abs(v) < 100
    # 行首一串小整数（附注号，如「59(1)」）后面全是大金额 → 整串都是附注号
    k = 0
    while k < len(vals) and small(vals[k]):
        k += 1
    if k and len(vals) - k >= 2 and all(abs(v) >= 100 for v in vals[k:]):
        return vals[k:]
    # 否则：金额列成对出现（本期/上期，或合并+母公司共 4 列），个数为偶数时第一个数就是金额
    while len(vals) > 2 and len(vals) % 2 == 1 and small(vals[0]):
        vals = vals[1:]
    return vals


def _numbers(rest):
    vals = [tonum(x) for x in _NUM.findall(rest)]
    vals = [v for v in vals if v is not None]
    return _trim_note_numbers(vals)


def _find_pages(pdf, sheet, limit=280, extend=True):
    """优先用 locate()；失败就按文本里的「合并XX表」找（保险公司的表在很后面）。
    报表常跨页，把紧邻的下一页也带上（下一页不是另一张表的标题时）。"""
    loc, _ = locate(pdf)
    pages = loc.get(sheet) or []
    if not pages:
        kw = "合并" + sheet
        pages = []
        for i, pg in enumerate(pdf.pages[:limit]):
            if kw in (pdftext.text(pdf.path, i) or pg.extract_text() or ""):
                pages.append(i)
            if len(pages) >= 4:
                break
    if len(pages) >= 2 and pages[-1] - pages[0] <= 3:
        pages = list(range(pages[0], pages[-1] + 1))
    out = []
    for p in pages:
        if p not in out:
            out.append(p)
        if not extend:
            continue
        nxt = p + 1
        if nxt < len(pdf.pages) and nxt not in out:
            t = pdftext.text(pdf.path, nxt) or pdf.pages[nxt].extract_text() or ""
            other = [s for s in ("合并资产负债表", "合并利润表", "合并现金流量表") if s != "合并" + sheet]
            if not any(o in t for o in other):
                out.append(nxt)
    return sorted(out)[:6]


def detect_kind(pdf, limit=60, company=None):
    """按公司名或特征科目识别银行/保险/证券。"""
    if company:
        if "银行" in company:
            return "bank"
        if any(w in company for w in ("保险", "人寿", "太保", "平安", "新华")):
            return "insurance"
        if any(w in company for w in ("证券", "海通", "中金", "申万", "国泰君安")):
            return "securities"
    counts = {k: 0 for k in SIGNATURES}
    for i, pg in enumerate(pdf.pages[:limit]):
        t = pdftext.flat(pdf.path, i) or re.sub(r"\s+", "", pg.extract_text() or "")
        for kind, keys in SIGNATURES.items():
            for k in keys:
                if k in t:
                    counts[kind] += 1
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else "bank"


def _parse_sheet(pdf, pages, alias, want, year):
    """每张表每个科目取第一次出现；支持「科目一行、数字下一行」的版式。"""
    facts, seen = [], set()
    mult = 1
    for p in pages:
        txt = pdf.pages[p].extract_text() or ""
        # pdfplumber 会把「经营活动现金流入小计」这类长科目名和数字拆到不同行，
        # 快速文字层（pypdfium2）的行更完整，优先用它
        fl = pdftext.lines(getattr(pdf, "path", None), p) if getattr(pdf, "path", None) else None
        if fl:
            txt = "\n".join(fl)
        mult = _unit_mult(txt, mult)
        pending = None
        carry = ""          # 上一行只有半截科目名（「经营活动产生 (使用) 的现金流量」）
        for line in txt.splitlines():
            line = line.strip()
            if not line or len(line) > 200:
                continue
            # 科目名被折成两行：上一行没数字，拼起来能对上别名就用拼后的
            if carry and not _numbers(carry):
                joined = carry + line
                if any(clean("".join(joined.split()[:k])) in alias for k in range(1, 6)):
                    line = joined
                carry = ""
            toks = line.split()
            matched = False
            for k in range(1, min(9, len(toks)) + 1):
                name = clean("".join(toks[:k])[:40])
                std = alias.get(name)
                if not std or std not in want or std in seen:
                    continue
                matched = True
                vals = _numbers(" ".join(toks[k:]))
                if vals:
                    seen.add(std)
                    facts.append((year, std, vals[0] * mult))
                    if len(vals) >= 2:
                        facts.append((year - 1, std, vals[1] * mult))
                    pending = None
                else:
                    pending = (std, mult)
                break
            if matched:
                continue
            if pending and line:
                std, mult2 = pending
                vals = _numbers(line)
                if vals:
                    seen.add(std)
                    facts.append((year, std, vals[0] * mult2))
                    if len(vals) >= 2:
                        facts.append((year - 1, std, vals[1] * mult2))
                    pending = None
                    continue
            if not _numbers(line) and len(line) <= 30:
                carry = line        # 半截科目名，等下一行接上
    return facts


def _year_fallback(pdf, path):
    """保险年报的封面可能没有标准年度行，退回到正文/文件名。"""
    # 1) 找「20XX年12月31日」
    for i, pg in enumerate(pdf.pages[:150]):
        t = pdftext.flat(pdf.path, i) or re.sub(r"\s+", "", pg.extract_text() or "")
        m = re.search(r"(20\d{2})年12月31日", t)
        if m:
            return int(m.group(1))
    # 2) 文件名里的年份
    m = re.search(r"(20\d{2})", os.path.basename(str(path)))
    return int(m.group(1)) if m else None


def _rows_by_coords(page, tol=10):
    """按 y 坐标把词聚成行（保险报表的标签与数字常有 5-10px 错位）。"""
    words = page.extract_words()
    words.sort(key=lambda w: (w["top"], w["x0"]))
    rows = []
    for w in words:
        if rows and abs(w["top"] - rows[-1]["top"]) <= tol:
            rows[-1]["words"].append(w)
        else:
            rows.append({"top": w["top"], "words": [w]})
    return rows


def _parse_sheet_coords(pdf, pages, alias, want, year):
    """坐标版：标签在左、金额在右（两栏报表靠这个救）。"""
    facts, seen = [], set()
    mult = 1
    for p in pages:
        pg = pdf.pages[p]
        mult = _unit_mult(pg.extract_text() or "", mult)
        for row in _rows_by_coords(pg):
            label_parts, nums = [], []
            for w in sorted(row["words"], key=lambda x: x["x0"]):
                t = w["text"].strip()
                if re.fullmatch(r"\(?-?[\d,]+(?:\.\d+)?\)?", t):
                    nums.append(tonum(t))
                elif not nums:
                    label_parts.append(t)
            label = clean("".join(label_parts)[:40])
            std = alias.get(label)
            if not std or std not in want or std in seen:
                continue
            vals = [v for v in nums if v is not None]
            vals = _trim_note_numbers(vals)
            if not vals:
                continue
            seen.add(std)
            facts.append((year, std, vals[0] * mult))
            if len(vals) >= 2:
                facts.append((year - 1, std, vals[1] * mult))
    return facts


def parse(path, company=None, kind=None):
    """解析金融三表，返回 (DataFrame, 报告行列表)。kind=None 时自动识别。"""
    facts, report = [], []
    if company is None:
        company = os.path.splitext(os.path.basename(path))[0]
    with pdfplumber.open(path) as pdf:
        kind = kind or detect_kind(pdf, company=company)
        alias, want = _meta(kind)
        year = detect_report_year(pdf)
        if year is None:
            year = _year_fallback(pdf, path)
        if year is None:
            return pd.DataFrame(columns=["公司"] + BASE_COLS), ["无法识别报告年度"]
        report.append("类型: %s" % {"bank": "银行", "insurance": "保险", "securities": "证券"}[kind])
        for sheet in ("资产负债表", "利润表", "现金流量表"):
            # 保险的资产负债表是严格两栏，多带一页会串；利润表常跨页，需要带上下一页
            extend = (sheet != "资产负债表") if kind == "insurance" else True
            pages = _find_pages(pdf, sheet, extend=extend)
            if kind == "insurance":
                got = _parse_sheet_coords(pdf, pages, alias, want, year)
                got_text = _parse_sheet(pdf, pages, alias, want, year)
                if sheet == "资产负债表":
                    def _bs_ok(fs):
                        m = {(f[0], f[1]): f[2] for f in fs}
                        a, l, e = m.get((year, "总资产")), m.get((year, "总负债")), m.get((year, "所有者权益"))
                        return None not in (a, l, e) and abs(a - l - e) <= max(1.0, abs(a) * 0.005)
                    if not _bs_ok(got) and _bs_ok(got_text):
                        got = got_text
                if len(got) < 4:      # 坐标法也拿不到就退回文本法
                    got = got_text
                else:
                    have = {(f[0], f[1]) for f in got}
                    for f in got_text:
                        if (f[0], f[1]) not in have:
                            got.append(f)
                            have.add((f[0], f[1]))
            else:
                got = _parse_sheet(pdf, pages, alias, want, year)
            facts.extend(got)
            report.append("%s: %d 页, 取到 %d 条" % (sheet, len(pages), len(got)))

    df = pd.DataFrame(facts, columns=BASE_COLS)
    df.insert(0, "公司", company)
    df.attrs["fin_kind"] = kind
    return df, report


def _normalize(df):
    """修两处常见口径问题：
    1) 所得税费用符号（有的报表列成负数，但恒等式要求正数；或递延所得税冲回致正数实为抵减）
    2) 归母净利润抓串行（比分部/母公司表误抓）时直接丢掉，避免 ROE 失真
    """
    if df.empty:
        return df
    p = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
    fixes = []
    for y in p.columns:
        def v(k):
            return float(p.loc[k, y]) if k in p.index and pd.notna(p.loc[k, y]) else None

        net, total, tax = v("净利润"), v("利润总额"), v("所得税费用")
        if None not in (net, total, tax):
            direct = abs(net - (total - tax))
            flipped = abs(net - (total - abs(tax)))
            negated = abs(net - (total - (-tax)))
            if direct > 1 and flipped <= 1:
                fixes.append((y, "所得税费用", abs(tax)))
            elif direct > 1 and negated <= 1:
                fixes.append((y, "所得税费用", -tax))

        a, l, e = v("总资产"), v("总负债"), v("所有者权益")
        pe, me = v("归母权益"), v("少数股东权益")
        if a is not None and e is not None and (l is None or abs(a - l - e) > 1):
            if pe is not None and abs((pe + (me or 0)) - e) <= max(1.0, abs(e) * 0.002):
                fixes.append((y, "总负债", a - e))

        parent, net2 = v("归母净利润"), v("净利润")
        if parent is not None and net2 not in (None, 0) and abs(parent) > abs(net2) * 1.2:
            fixes.append((y, "归母净利润", None))

    for y, item, val in fixes:
        mask = (df["年度"] == y) & (df["标准项目"] == item)
        if val is None:
            df = df[~mask]
        elif mask.any():
            df.loc[mask, "金额"] = val
        else:
            row_d = {"公司": df["公司"].iloc[0], "年度": y, "标准项目": item, "金额": val}
            df = pd.concat([df, pd.DataFrame([row_d])], ignore_index=True)

    # 3) 利润总额 = 营业利润 + 营业外收入 - 营业外支出
    #    保险公司的这一行常带附注编号（「加：营业外收入 54 691 912」），
    #    容易把上一年或附注号当成本期值；用恒等式反推，能对上就纠正。
    p2 = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
    repairs = []
    for y in p2.columns:
        def v(k):
            return float(p2.loc[k, y]) if k in p2.index and pd.notna(p2.loc[k, y]) else None

        op, total = v("营业利润"), v("利润总额")
        oi, oe = v("营业外收入"), v("营业外支出")
        if None in (op, total, oi, oe):
            continue
        if abs(total - (op + oi - oe)) <= 1:
            continue
        oi_alts = [oi] + [float(p2.loc["营业外收入", yy]) for yy in p2.columns
                          if yy != y and "营业外收入" in p2.index
                          and pd.notna(p2.loc["营业外收入", yy])]
        oe_alts = [oe, abs(oe)]
        done = False
        for o in oi_alts:
            for e in oe_alts:
                if abs(total - (op + o - e)) <= 1:
                    if o != oi:
                        repairs.append((y, "营业外收入", o))
                    if e != oe:
                        repairs.append((y, "营业外支出", e))
                    done = True
                    break
            if done:
                break
    for y, item, val in repairs:
        mask = (df["年度"] == y) & (df["标准项目"] == item)
        df.loc[mask, "金额"] = val
    return df.reset_index(drop=True)


def finish(df):
    """兜底：股东权益没取到时用 资产-负债 补；经营现金流净额没取到时用流入-流出补。"""
    if df.empty:
        return df
    p = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
    add = []
    for y in p.columns:
        def v(k):
            return p.loc[k, y] if k in p.index and pd.notna(p.loc[k, y]) else None
        if v("所有者权益") is None and v("总资产") is not None and v("总负债") is not None:
            add.append({"公司": df["公司"].iloc[0], "年度": y, "标准项目": "所有者权益",
                        "金额": v("总资产") - v("总负债")})
        ci, co = v("经营现金流入"), v("经营现金流出")
        if v("经营现金流净额") is None and ci is not None and co is not None:
            add.append({"公司": df["公司"].iloc[0], "年度": y, "标准项目": "经营现金流净额",
                        "金额": ci - abs(co)})
        # 证券公司「营业收入」常是分段标题，数字被下一行「手续费及佣金净收入」吃掉
        rev, tot = v("营业收入"), v("营业总收入")
        if tot and (rev is None or abs(rev) < abs(tot) * 0.5):
            m = (df["年度"] == y) & (df["标准项目"] == "营业收入")
            if m.any():
                df.loc[m, "金额"] = tot
            else:
                add.append({"公司": df["公司"].iloc[0], "年度": y, "标准项目": "营业收入",
                            "金额": tot})
    if add:
        df = pd.concat([df, pd.DataFrame(add)], ignore_index=True)
    return _normalize(df)


def validate(df, tol=0.005):
    """四条恒等式；返回 [(名称, 是否通过, 差额)]"""
    checks = []
    if df.empty:
        return checks
    kind = df.attrs.get("fin_kind", "bank")
    p = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
    for y in sorted(p.columns):
        def v(k):
            return float(p.loc[k, y]) if k in p.index and pd.notna(p.loc[k, y]) else None

        a, l, e = v("总资产"), v("总负债"), v("所有者权益")
        if None not in (a, l, e):
            d = a - l - e
            checks.append(("%s 资产 = 负债 + 权益" % y, abs(d) <= max(1.0, abs(a) * tol), d))

        np_, total, tax = v("净利润"), v("利润总额"), v("所得税费用")
        if None not in (np_, total, tax):
            d = np_ - (total - tax)
            checks.append(("%s 净利润 = 利润总额 - 所得税" % y, abs(d) <= max(1.0, abs(np_) * tol), d))

        op, oi, oe = v("营业利润"), v("营业外收入"), v("营业外支出")
        if None not in (op, total) and oi is not None and oe is not None:
            d = total - (op + oi - oe)
            checks.append(("%s 利润总额 = 营业利润 + 营业外收支净额" % y,
                           abs(d) <= max(1.0, abs(total) * tol), d))

        rev = v("营业收入")
        parts = [v(k) for k in REV_PARTS[kind]]
        # 主要构成项没取到（或为 0）时拆分本来就不完整，不做这条判断
        if any(p in (None, 0) for p in parts[:2]):
            parts = []
        # 保险的「营业收入」构成各家差异大（新准则下有保险服务收入/已赚保费两套），
        # 只对银行/证券做收入拆分校验
        # 主要构成项（银行=利息净收入，证券=手续费及佣金净收入）没取到时拆分不完整，不做这条
        if parts and kind != "insurance" and rev is not None and parts[0] is not None:
            d = rev - sum(x or 0 for x in parts)
            checks.append(("%s 营业收入 ≈ 各项收入之和" % y,
                           abs(d) <= max(1.0, abs(rev) * 0.08), d))
    return checks


def main():
    if len(sys.argv) < 2:
        print("usage: parse_bank.py <金融年报.pdf> [kind]")
        return 2
    kind = sys.argv[2] if len(sys.argv) > 2 else None
    df, rep = parse(sys.argv[1], kind=kind)
    df = finish(df)
    print("\n".join(rep))
    if df.empty:
        print("没取到数据")
        return 1
    p = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
    for k in p.index:
        print("  %-26s %s" % (k, " ".join("%.2f亿" % (x / 1e8) for x in p.loc[k])))
    print("\n校验:")
    for name, ok, diff in validate(df):
        print("  %-46s %s 差额=%.2f" % (name, "OK" if ok else "FAIL", diff))
    return 0


if __name__ == "__main__":
    sys.exit(main())

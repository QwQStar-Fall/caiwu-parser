# -*- coding: utf-8 -*-
"""
parse_statements.py —— 精准解析年报里的三大报表

策略（针对真实年报）：
  1. 全文档扫描，按标题定位「合并资产负债表 / 合并利润表 / 合并现金流量表」所在页
  2. 报表可能跨页：从标题页开始，向后包含直到下一张报表或连续 2 页没有数据行
  3. 只解析这些页里的表格，避免把附注里的杂表混进来
  4. 表头里抓年份（支持「2025年12月31日」「2025年度」）
"""
import io, re, os, sys
import pdfplumber
import pandas as pd
try:
	from . import pdftext
except ImportError:
	import pdftext

YEAR_RE = re.compile(r"(20\d{2})")

TITLES = {
	"资产负债表": ["合并资产负债表", "合并及公司资产负债表", "合并及母公司资产负债表", "合并及银行资产负债表",
	              "合并资产负债表和资产负债表", "会合01表", "资产负债表"],
	"利润表": ["合并利润表", "合并及公司利润表", "合并及母公司利润表", "合并及银行利润表",
	          "合并利润表和利润表", "会合02表", "利润表"],
	"现金流量表": ["合并现金流量表", "合并及公司现金流量表", "合并及母公司现金流量表", "合并及银行现金流量表",
	              "合并现金流量表和现金流量表", "会合03表", "现金流量表"],
}

UNIT_MULT = {"元": 1, "千元": 1e3, "万元": 1e4, "百万元": 1e6, "亿元": 1e8}
_UNIT_RE = re.compile(r"(?:金额)?单位(?:均为|为)?[:：]?\s*(?:人民币)?\s*(百万元|亿元|万元|千元|元)"
                      r"|币种[:：为]{0,2}\s*(?:人民币)?\s*(百万元|亿元|万元|千元|元)"
                      r"|人民币\s*(百万元|亿元|万元|千元)")

# 整行只写「人民币元」的版式（万科等）：不能把正文里零散出现的「人民币元」也算成单位声明
_UNIT_YUAN_LINE = re.compile(r"^[（(]?人民币[:：]?元[)）]?$")


def page_unit(txt):
	"""页面上声明的金额单位 → 倍数；没有声明返回 None（调用方沿用上一页）"""
	raw = txt or ""
	flat = re.sub(r"\s+", "", raw)
	ms = list(_UNIT_RE.finditer(flat))
	if ms:
		# 一页里有多处单位声明（如「附注单位为千元」+ 表头「单位：元」）时，以最后一处、紧挨表格的为准
		m = ms[-1]
		u = next((g for g in m.groups() if g), None)
		return UNIT_MULT[u] if u else None
	# 兜底：整行就是「人民币元」，或短行以「人民币元」结尾（迈瑞「2025年12月31日 人民币元」）
	for ln in raw.splitlines():
		s = re.sub(r"\s+", "", ln)
		if not s:
			continue
		if _UNIT_YUAN_LINE.match(s) or (s.endswith("人民币元") and len(s) <= 30):
			return 1
	return None

# 所有报表标题（含母公司），用于划定"段边界"——防止把母公司报表混进合并报表
ALL_TITLES = [
	"合并资产负债表", "合并及公司资产负债表", "合并及母公司资产负债表", "合并及银行资产负债表",
	"合并资产负债表和资产负债表", "母公司资产负债表", "公司资产负债表", "银行资产负债表", "会合01表", "会企01表", "资产负债表",
	"合并利润表", "合并及公司利润表", "合并及母公司利润表", "合并及银行利润表",
	"合并利润表和利润表", "母公司利润表", "公司利润表", "银行利润表", "会合02表", "会企02表", "利润表",
	"合并现金流量表", "合并及公司现金流量表", "合并及母公司现金流量表", "合并及银行现金流量表",
	"合并现金流量表和现金流量表", "母公司现金流量表", "公司现金流量表", "银行现金流量表", "会合03表", "会企03表", "会合04表", "会企04表", "现金流量表",
]

# A股年报常见科目 -> 标准名
EXTRA_ALIAS = {
	"资产总计": "总资产", "资产总额": "总资产", "负债合计": "总负债", "负债总额": "总负债", "负债总计": "总负债",
	"所有者权益（或股东权益）合计": "所有者权益",
	"所有者权益(或股东权益)合计": "所有者权益",
	"所有者权益合计": "所有者权益",
	"股东权益合计": "所有者权益",
	"归属于母公司所有者权益合计": "归母权益",
	"归属于母公司股东权益合计": "归母权益",
	"归属于母公司所有者权益（或股东权益）合计": "归母权益",
	"少数股东权益": "少数股东权益",
	"流动资产合计": "流动资产", "非流动资产合计": "非流动资产",
	"流动负债合计": "流动负债", "非流动负债合计": "非流动负债",
	"货币资金": "货币资金", "存货": "存货", "应收账款": "应收账款", "应付账款": "应付账款",
	"固定资产": "固定资产净额", "在建工程": "在建工程",
	"营业总收入": "营业总收入", "营业收入": "营业收入",
	"营业总成本": "二、营业总成本",
	"营业成本": "营业成本", "税金及附加": "营业税金及附加", "营业税金及附加": "营业税金及附加",
	"销售费用": "销售费用", "管理费用": "管理费用", "研发费用": "研发费用",
	"财务费用": "财务费用",
	"资产减值损失": "资产减值损失",
	"营业利润": "营业利润", "利润总额": "利润总额",
	"净利润": "净利润", "归属于母公司所有者的净利润": "归母净利润",
	"经营活动产生的现金流量净额": "经营现金流净额",
	"经营活动使用的现金流量净额": "经营现金流净额",
	"投资活动产生的现金流量净额": "投资现金流净额",
	"筹资活动产生的现金流量净额": "筹资现金流净额",
	"销售商品、提供劳务收到的现金": "销售商品收现",
	"支付给职工以及为职工支付的现金": "职工薪酬现金",
	# —— A股年报常见补充 ——
	"负债和所有者权益（或股东权益）总计": "负债和权益总计",
	"负债和所有者权益总计": "负债和权益总计",
	"归属于母公司股东的净利润": "归母净利润", "负债及股东权益总计": "负债和权益总计",
	"归属于母公司所有者的净利润": "归母净利润",
	"归属于母公司所有者（或股东）的净利润": "归母净利润",
	"归属于母公司所有者(或股东)的净利润": "归母净利润",
	"少数股东损益": "少数股东损益",
	"营业外收入": "营业外收入", "营业外支出": "营业外支出",
	"所得税费用": "所得税费用", "所得税": "所得税费用",
	"其他收益": "其他收益",
	"投资收益": "投资收益",
	"信用减值损失": "信用减值损失",
	"资产处置收益": "资产处置收益",
	"公允价值变动收益": "公允价值变动收益",
	"营业总收入": "营业总收入",
	"二、营业总成本": "营业总成本",
	"营业总成本": "营业总成本",
	"购买商品、接受劳务支付的现金": "购买商品付现",
	"支付其他与经营活动有关的现金": "其他经营付现",
	"收到其他与经营活动有关的现金": "其他经营收现",
	"经营活动现金流入小计": "经营现金流入",
	"经营活动现金流出小计": "经营现金流出",
	"投资活动现金流入小计": "投资现金流入",
	"投资活动现金流出小计": "投资现金流出",
	"筹资活动现金流入小计": "筹资现金流入",
	"筹资活动现金流出小计": "筹资现金流出",
	"现金及现金等价物净增加额": "现金净增加额",
	"期末现金及现金等价物余额": "期末现金余额",
	"期初现金及现金等价物余额": "期初现金余额",
	# —— 地产/建筑的关键科目 ——
	"合同负债": "合同负债",
	"预收款项": "预收款项",
	"合同资产": "合同资产",
	"短期借款": "短期借款",
	"长期借款": "长期借款",
	"一年内到期的非流动负债": "一年内到期非流动负债",
	# —— 留存收益与股本相关科目 ——
	"未分配利润": "未分配利润",
	"盈余公积": "盈余公积",
	"实收资本（或股本）": "股本", "实收资本(或股本)": "股本", "股本": "股本", "实收资本": "股本",
	"利息费用": "利息费用",
}


def name_cell(row):
    """科目名不一定在第 0 列（有的报表前面有一列空列）。
    取前 3 列里第一个"像科目名"的单元格：含中文或字母、不是纯数字。"""
    for j in range(min(3, len(row))):
        v = row[j]
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        if re.search(r"[\u4e00-\u9fa5A-Za-z]", s):
            return j, s
        return None, None      # 第一个非空单元格就是数字 → 不是科目行
    return None, None


# 部分年报的字体映射有问题，把「并」印成「幵」、「份」印成「仹」等，匹配前先纠正
GLYPH_FIX = str.maketrans({
	"幵": "并", "仹": "份", "陉": "限", "陋": "除",
	"劢": "动", "癿": "的", "殎": "母", "贶": "费", "仐": "他",
})


def clean(s):
	if s is None:
		return ""
	s = str(s).translate(GLYPH_FIX).replace("\u3000", " ").replace("\xa0", " ").replace("\n", "")
	s = re.sub(r"\s+", "", s)
	s = s.replace("归属二母公司", "归属于母公司").replace("归属二少数股东", "归属于少数股东")
	s = re.sub(r"^(一|二|三|四|五|六|七|八|九|十|亏)+[、.]", "", s)
	s = re.sub(r"^（[一二三四五六七八九十]+）", "", s)
	s = re.sub(r"^(加|减)[：:]", "", s)
	s = re.sub(r"^(其中|其中：|其中:)", "", s)
	s = re.sub(r"^[：:、·・\-－—]+", "", s)   # 行首残留的冒号/顿号/项目符号
	# 亏损年度的写法：营业亏损、利润(亏损)总额、净亏损、归属于母公司股东的净亏损
	s = re.sub(r"利润[（(]亏损[）)]总额|亏损总额", "利润总额", s)
	s = re.sub(r"营业亏损$", "营业利润", s)
	s = re.sub(r"少数股东净亏损$", "少数股东损益", s)
	s = re.sub(r"净亏损$", "净利润", s)
	s = re.sub(r"^(持续经营|终止经营)净利润$", r"\1净利润", s)
	# 港式写法：经营活动产生/(使用)的现金流量净额、所得税(费用)/贷项
	s = re.sub(r"[/／][（(][^（）()]*[）)]", "", s)
	s = re.sub(r"[（(][^（）()]*[）)][/／]", "", s)
	s = re.sub(r"[（(](费用|收益|费用\)/贷项)[）)]", r"\1", s)
	s = re.sub(r"[（(](使用|或使用|使用或产生)[）)]", "", s)   # 经营活动产生（使用）的现金流量净额
	s = re.sub(r"(所得税)(费用)?[/／]?(贷项|抵免)$", "所得税费用", s)
	s = re.sub(r"附注.*$", "", s) if len(s) > 4 else s
	# 剥掉尾部的括号说明，如 营业利润（亏损以"－"号填列）
	for _ in range(3):
		t = re.sub(r"[（(][^（）()]*[）)]\s*$", "", s)
		if t == s:
			break
		s = t
	# 行尾被截断的未闭合括号：净利润（净亏损以"－"号
	s = re.sub(r"[（(][^（）()]*$", "", s)
	s = re.sub(r"^\d+[.．、]", "", s)   # 去掉 "1." 这类序号
	return s.strip()


def tonum(v):
	if v is None:
		return None
	if isinstance(v, (int, float)):
		return float(v)
	s = str(v).strip().replace(",", "").replace("，", "").replace(" ", "").replace("\n", "")
	if s in ("", "-", "—", "－", "--", "/"):
		return 0.0
	neg = False
	if (s.startswith("(") and s.endswith(")")) or (s.startswith("（") and s.endswith("）")):
		neg, s = True, s[1:-1]
	s = s.replace("%", "")
	try:
		n = float(s)
	except ValueError:
		return None
	return -n if neg else n


def find_year_cols(rows):
	"""在若干行里找年份列 → {列号: 年份}"""
	best = {}
	for r in rows[:6]:
		cur = {}
		for j, c in enumerate(r or []):
			if c is None:
				continue
			m = YEAR_RE.search(str(c))
			if m:
				y = int(m.group(1))
				if 1990 <= y <= 2100 and y not in cur.values():
					cur[j] = y
		if len(cur) > len(best):
			best = cur
	return best


# 期末 / 期初 的常见叫法（表头里没有年份时用）
PERIOD_CUR = ("期末余额", "期末数", "本期金额", "本年金额", "本期发生额", "本年发生额",
              "本年累计", "本年累计数", "本期数", "本期金额", "年末余额")
PERIOD_PRV = ("期初余额", "期初数", "上期金额", "上年金额", "上期发生额", "上年发生额",
              "上年同期", "上年累计数", "上期数", "上年年末", "年初余额")

BASE_COLS = ["年度", "标准项目", "金额"]


def detect_report_year(pdf, limit=30):
    """推断报告年度。优先封面「20XX年年度报告」，其次正文里的「20XX年12月31日」。
    （正文里常出现历史年度，直接找日期容易误判，所以封面优先。）"""
    import re as _re
    head = []
    src = getattr(pdf, "path", None)
    for i, pg in enumerate(pdf.pages):
        if i >= limit:
            break
        t = pdftext.text(src, i) if src else None
        head.append(t if t is not None else (pg.extract_text() or ""))
    for txt in head[:5]:
        m = _re.search(r"(20\d{2})\s*年\s*(?:年度报告|度报告|年报)", txt)
        if m:
            return int(m.group(1))
    for txt in head[:5]:
        m = _re.search(r"本公司\s*(20\d{2})\s*年度", txt)
        if m:
            return int(m.group(1))
    yrs = [int(y) for txt in head for y in _re.findall(r"(20\d{2})\s*年\s*(?:年度报告|度报告|年报)", txt)]
    if yrs:
        return max(yrs)
    for txt in head:
        m = _re.search(r"(20\d{2})\s*年\s*12\s*月\s*31\s*日", txt)
        if m:
            return int(m.group(1))
    return None


def period_years(header_row, report_year):
    """表头是「期末余额/期初余额」这类时，映射成 {列号: 年份}"""
    if report_year is None or not header_row:
        return {}
    cur_col = prv_col = None
    for j, c in enumerate(header_row):
        if c is None:
            continue
        s = str(c).replace(" ", "")
        if cur_col is None and any(k in s for k in PERIOD_CUR):
            cur_col = j
        elif prv_col is None and any(k in s for k in PERIOD_PRV):
            prv_col = j
    out = {}
    if cur_col is not None:
        out[cur_col] = report_year
    if prv_col is not None:
        out[prv_col] = report_year - 1
    return out


def title_cut(pg):
    """若本页含有报表标题，返回该标题的 top 坐标（用于只取上方内容）"""
    best = None
    for a in ALL_TITLES:
        try:
            found = pg.search(a)
        except Exception:
            found = None
        if found:
            for m in found:
                top = m.get("top")
                if top is not None and (best is None or top < best):
                    best = top
    return best


_TEXT_LINE = re.compile(
    r"^(?P<name>[^\d\s][^\d]{0,40}?)\s+"
    r"(?P<vals>(?:\(?-?[\d,]+(?:\.\d+)?\)?\s+){0,4}\(?-?[\d,]+(?:\.\d+)?\)?)\s*$")


def parse_text_pages(pdf, pages, report_year, alias, base_mult=1):
    """文本行兜底：表格对象不可靠时，直接从文字行取数。

    要处理的麻烦：
      · 科目名被空格拆开：  其 中 ： 营业收入 七、61 11,581,812,280.44
      · 名字在行尾、数字下一行：  五、净利润（净亏损以"－"号
      · 行首附注编号被当成数字：  七、76 175,864,151.39
    约定：第 1 个数字 = 本期，第 2 个 = 上期。
    """
    facts = []
    if report_year is None:
        return facts
    mult = base_mult
    for p in pages:
        txt = pdf.pages[p].extract_text() or ""
        mult = page_unit(txt) or mult
        raw_lines = [l.strip() for l in txt.splitlines() if l.strip()]
        # 左右分栏报表（左资产、右负债权益同处一行）拆成独立行
        lines = []
        for rl in raw_lines:
            parts = re.sub(r"((?:\d[\d,.]*|\s-))\s+(?=[\u4e00-\u9fa5])", r"\1\n", rl).splitlines()
            if len(parts) >= 2:
                lines.append(parts[0])
                for pt in parts[1:]:
                    pt_toks = pt.split()
                    if any(clean("".join(pt_toks[:k])) in alias for k in range(1, min(6, len(pt_toks)) + 1)):
                        lines.append(pt)
                    else:
                        lines[-1] = lines[-1] + " " + pt
            else:
                lines.append(rl)

        # 检测本页是否属于“数字在上、科目在下”的特殊颠倒排版（例如顺丰控股部分年报）
        above_count = 0
        for i in range(len(lines) - 1):
            l1, l2 = lines[i], lines[i+1]
            if any(term in l2 for term in ['流动资产', '货币资金', '应收账款', '营业收入', '营业成本', '税金及附加', '固定资产', '资产总计', '负债合计']):
                has_num = any(c.isdigit() for c in l1)
                zh_count = len([c for c in l1 if '\u4e00' <= c <= '\u9fff'])
                if has_num and zh_count == 0:
                    above_count += 1
        numbers_above = (above_count >= 2)

        pending = None          # 上一行只匹配到名字、没有数字
        pend_vals = None        # 上一行只有数字、名字在下一行
        for line in lines:
            if not line or len(line) > 160:
                continue
            toks = line.split()
            if len(toks) < 1:
                continue

            std = None
            used = 0
            for k in range(1, min(8, len(toks)) + 1):
                cand = clean("".join(toks[:k]))
                if cand in alias:
                    std, used = alias[cand], k
                    break

            rest = " ".join(toks[used:]) if std else line
            nums = re.findall(r"\(\s*-?[\d,]+(?:\.\d+)?\s*\)|-?[\d,]+(?:\.\d+)?", rest)
            vals = [v for v in (tonum(x) for x in nums) if v is not None]
            while len(vals) >= 3 and abs(vals[0]) < 1000:
                vals = vals[1:]

            def emit(name, vs):
                for i, v in enumerate(vs[:2]):
                    facts.append((report_year - i, name, v * mult))

            if std is not None:
                if vals:                      # 名字和数字同一行
                    emit(std, vals)
                    pending = pend_vals = None
                elif numbers_above and pend_vals:  # 数字在上一行、名字在这一行
                    emit(std, pend_vals)
                    pending = pend_vals = None
                else:
                    pending = std             # 名字到此为止，数字在下一行
            elif pending is not None:
                if vals:                      # 用上一行遗留的名字
                    emit(pending, vals)
                    pending = pend_vals = None
            elif numbers_above and vals:
                pend_vals = vals              # 只有数字，名字在下一行
    return facts


def locate(pdf):
	"""返回 ({报表: [合并表所在页]}, 全部报表标题页)
	· 行级精确匹配，避免正文里提到"资产负债表"就误命中
	· 优先"合并"表；没有合并版才退回无前缀版
	· 同时返回全部标题页，用于划段（防止母公司报表混入）
	"""
	merged = {k: [] for k in TITLES}
	plain = {k: [] for k in TITLES}
	all_pages = set()
	src = getattr(pdf, "path", None)
	for i, pg in enumerate(pdf.pages):
		fast = pdftext.lines(src, i) if src else None
		for line in (fast if fast is not None else (pg.extract_text() or "").splitlines()):
			L = _norm_title(line)
			if not L or len(L) > 22:
				continue
			hit = None
			for a in ALL_TITLES:
				if L == a or L.startswith(a + "（续") or L.startswith(a + "(续"):
					hit = a
					break
			if not hit:
				continue
			all_pages.add(i)
			for k, aliases in TITLES.items():
				if hit in aliases[:-1]:          # 合并XX表 / 合并及公司XX表
					if i not in merged[k]:
						merged[k].append(i)
				elif hit == aliases[-1]:         # 无前缀 XX表
					if i not in plain[k]:
						plain[k].append(i)

	# 快速文字层会把表头单元格（如变动情况表里的「合并利润表」列）拆成独立一行，
	# 看起来像标题。命中的页很少，用 pdfplumber 的版面行再确认一次
	if src and pdftext.pages(src) is not None:
		def is_title_page(i):
			for line in _join_split_dates(pdf.pages[i].extract_text() or ""):
				if _title_hit(line):
					return True
				L = _norm_title(line)
				if L and len(L) <= 32 and any(a in L for a in ALL_TITLES):
					return True
			return False
		bad = {i for i in all_pages if not is_title_page(i)}
		if bad:
			all_pages -= bad
			for d in (merged, plain):
				for k in d:
					d[k] = [i for i in d[k] if i not in bad]

	res = {}
	for k in TITLES:
		res[k] = merged[k] if merged[k] else plain[k]
	# 目录页也会列出「合并资产负债表」「合并利润表」+ 页码，被当成标题。
	# 同一页出现 2 张以上不同报表的标题 → 视为目录页，有其他候选时丢掉
	claims = {}
	for k in TITLES:
		for i in res[k]:
			claims.setdefault(i, set()).add(k)
	toc = {i for i, ks in claims.items() if len(ks) >= 2}
	if toc and len(toc) < len(claims):
		for k in res:
			res[k] = [i for i in res[k] if i not in toc] or res[k]
		all_pages -= toc
	# 报表正文是连续的一段（三张表前后相距几页）；离这段很远的同名标题
	# 是附注里的「会计政策变更调整表」等，丢掉
	firsts = [v[0] for v in res.values() if v]
	if len(firsts) == len(TITLES):
		lo, hi = min(firsts), max(firsts) + 12
		far = set()
		for k in res:
			far |= {i for i in res[k] if i > hi or i < lo}
			res[k] = [i for i in res[k] if lo <= i <= hi] or res[k]
		all_pages -= far
	return res, sorted(all_pages)


def _join_split_dates(text):
	"""pdfplumber 会把「2025年12月31日合并资产负债表」拆成
	「2025 12 31」「年 月 日合并资产负债表」两行，这里把纯日期碎片并回下一行"""
	out, buf = [], ""
	for line in text.splitlines():
		s = re.sub(r"\s+", "", line)
		if s and re.fullmatch(r"[0-9０-９年月日]{1,10}", s):
			buf += s
			continue
		out.append((buf + line) if buf else line)
		buf = ""
	return out


def _norm_title(line):
	"""标题行归一：去空白、纠正错字、剥掉「2025年12月31日」「（续）前编号」等前缀"""
	L = re.sub(r"\s+", "", line).translate(GLYPH_FIX)
	L = re.sub(r"^20\d{2}年(?:\d{1,2}月\d{1,2}日|度)?", "", L)
	# 拆行后拼起来会变成「20251231年月日合并资产负债表」，把开头这串日期字符去掉
	L = re.sub(r"^[0-9０-９]{4,8}[年月日]{0,3}[0-9０-９]{0,4}[年月日]{0,2}[0-9０-９]{0,2}[年月日]?", "", L)
	L = re.sub(r"^[（(]?[0-9０-９一二三四五六七八九十]+[）)]?[、.．/／]?", "", L)
	return L


def _title_hit(line):
	L = _norm_title(line)
	if not L or len(L) > 22:
		return None
	for a in ALL_TITLES:
		if L == a or L.startswith(a + "（续") or L.startswith(a + "(续"):
			return a
	return None


def _use(pdf, loc):
	"""若同时找到合并表和母公司表，只保留合并表所在段"""
	return loc


OFFICE_EXT = (".xlsx", ".xlsm", ".xls", ".docx")


def _office_tables(path):
	"""Excel / Word → [(表名, 上下文文字, 行列表)]；上下文用于识别单位、年份、母公司"""
	ext = os.path.splitext(path)[1].lower()
	out = []
	if ext in (".xlsx", ".xlsm"):
		import openpyxl
		wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
		for ws in wb.worksheets:
			rows = [list(r) for r in ws.iter_rows(values_only=True)]
			out.append((ws.title, ws.title, rows))
		wb.close()
	elif ext == ".xls":
		try:
			import xlrd
		except ImportError:
			raise ValueError("读取 .xls 格式需要安装 xlrd（pip install xlrd），或另存为 .xlsx")
		book = xlrd.open_workbook(path)
		for sh in book.sheets():
			rows = [sh.row_values(i) for i in range(sh.nrows)]
			out.append((sh.name, sh.name, rows))
	elif ext == ".docx":
		import docx
		from docx.table import Table
		from docx.text.paragraph import Paragraph
		doc = docx.Document(path)
		recent = []
		n = 0
		for el in doc.element.body.iterchildren():
			tag = el.tag.rsplit("}", 1)[-1]
			if tag == "p":
				t = Paragraph(el, doc).text.strip()
				if t:
					recent = (recent + [t])[-3:]
			elif tag == "tbl":
				n += 1
				rows = [[c.text for c in r.cells] for r in Table(el, doc).rows]
				out.append(("表%d" % n, " ".join(recent), rows))
				recent = []
	else:
		raise ValueError("不支持的格式: " + ext)
	return out


def _unit_of(text):
	if "单位" not in text:
		return None
	for u in ("百万元", "亿元", "万元", "千元", "元"):
		if u in text:
			return UNIT_MULT[u]
	return None


def parse_office(path, company=None, file_year=None):
	"""Excel / Word 版财务报表：表头带年份（或期末/期初）、首列是科目名。"""
	tables = _office_tables(path)
	report = []
	parsed = []                       # [(是否母公司, facts)]
	if file_year is None:
		m = YEAR_RE.search(os.path.basename(path))
		file_year = int(m.group(1)) if m else None
	for title, ctx, rows in tables:
		rows = [r for r in rows if r is not None]
		if len(rows) < 2:
			continue
		head = ctx + " " + " ".join(str(c) for r in rows[:4] for c in r if c not in (None, ""))
		mult = _unit_of(head) or 1
		hi, years = None, {}
		for i, r in enumerate(rows[:15]):
			y = {j: v for j, v in find_year_cols([r]).items()
			     if j > 0 and isinstance(r[j], str) and len(r[j]) <= 45
			     and (tonum(r[j]) is None or re.fullmatch(r"20\d{2}", r[j].strip()))}
			if len(y) > len(years):
				hi, years = i, y
		if not years:
			ym = re.search(r"(20\d{2})\s*年", head)
			ry = int(ym.group(1)) if ym else file_year
			for i, r in enumerate(rows[:15]):
				y = period_years([None] + list(r[1:]), ry)
				if y:
					hi, years = i, y
					break
		if not years:
			continue
		facts = []
		for r in rows[hi + 1:]:
			_j, txt = name_cell(r)
			if txt is None:
				continue
			raw = clean(txt)
			std = EXTRA_ALIAS.get(raw) if raw and len(raw) <= 40 else None
			if not std:
				continue
			for col, yr in years.items():
				if col >= len(r) or r[col] is None or (isinstance(r[col], str) and not r[col].strip()):
					continue
				v = tonum(r[col])
				if v is not None:
					facts.append((yr, std, v * mult))
		if facts:
			is_parent = "母公司" in (title + " " + ctx + " " + " ".join(
				str(c) for r in rows[:3] for c in r if c not in (None, "")))
			parsed.append((is_parent, facts))
			report.append("%s: %s提取 %d 条" % (title, "母公司表，" if is_parent else "", len(facts)))
	if any(not p for p, _ in parsed):
		parsed = [x for x in parsed if not x[0]]
	facts = [f for _, fs in parsed for f in fs]
	if not facts:
		raise ValueError("没有在 Excel/Word 里找到可识别的报表（需要表头含年份、首列为科目名）")

	if company is None:
		company = os.path.splitext(os.path.basename(path))[0].split("：")[0]
	df = pd.DataFrame(facts, columns=BASE_COLS)
	df.insert(0, "公司", company)
	keep = df.groupby(["公司", "年度", "标准项目"])["金额"].apply(lambda s: s.abs().idxmax())
	df = df.loc[keep.values].reset_index(drop=True)
	return df, report


def _used_width(tb):
	"""表格实际用到的列数（忽略整列为空的尾列）"""
	used = {j for r in tb for j, c in enumerate(r or []) if c not in (None, "")}
	return max(used) + 1 if used else 0


def parse(path, company=None):
	if os.path.splitext(path)[1].lower() in OFFICE_EXT:
		return parse_office(path, company)
	with pdfplumber.open(path) as pdf:
		loc, all_title_pages = locate(pdf)
		report_year = detect_report_year(pdf)
		total_pages = len(pdf.pages)

		# 全局单位：只在某一页声明过「单位：千元」的，对每张表都生效
		# （比亚迪 2025 现金流量表没有单位行，只有资产负债表页声明过 → 不能每张表从 1 开始）
		global_mult = 1
		src = getattr(pdf, "path", None)
		# 先看报表正文第一页及其前两页（「财务附注中报表的单位为：千元」常写在这里），再看前 40 页
		first = min([p for v in loc.values() for p in v], default=None)
		scan = ([] if first is None else list(range(max(0, first - 2), first + 1))) + list(range(min(40, len(pdf.pages))))
		for i in scan:
			t = pdftext.lines(src, i) if src else None
			u = page_unit("\n".join(t) if t is not None else (pdf.pages[i].extract_text() or ""))
			if u and u != 1:
				global_mult = u
				break

		facts = []
		report = []
		for name in TITLES:
			pages = loc[name]
			if not pages:
				report.append("%s: 未定位到标题" % name)
				continue
			seg = {}
			for s in pages:
				# 段边界：下一个「任何报表标题」页。
				# 该页本身也收进来（合并表的尾部常常和母公司表标题同页），
				# 但只取标题以上的部分，避免把母公司数据混入。
				nxt = min([x for x in all_title_pages if x > s], default=total_pages)
				for p in range(s, min(nxt + 1, s + 5)):
					seg[p] = (p == nxt and nxt != total_pages)
			seg = sorted(seg.items())

			n_rows = 0
			# 续表页往往没有表头：把首页的年份列映射、金额单位继承下来
			seg_years = None
			seg_width = 0
			mult = global_mult
			for p, do_cut in seg:
				pg = pdf.pages[p]
				if do_cut:
					cut = title_cut(pg)
					if cut is not None and cut > 5:
						try:
							pg = pg.crop((0, 0, pg.width, cut - 2))
						except Exception:
							pass
						txt = pg.extract_text() or ""
					else:
						continue      # 找不到标题位置，整页丢弃更安全
				else:
					# 段首页也可能「上方是上一张表、标题压在下半页」：
					# 例如母公司利润表在上、合并现金流量表标题在页脚。
					# 这种情况只取标题以下的区域，避免母公司数据混进本表。
					cut = title_cut(pg)
					if cut is not None and cut > pg.height * 0.55:
						try:
							pg = pg.crop((0, max(0, cut - 2), pg.width, pg.height))
						except Exception:
							pass      # 个别 PDF 的 bbox 畸形，裁切会抛负宽高，保持整页
					txt = pg.extract_text() or ""
				mult = page_unit(txt) or mult
				try:
					_tbs = pg.extract_tables() or []
				except Exception:
					_tbs = []       # 个别 PDF 的 MediaBox 畸形，pdfplumber 会抛负宽高
				for tb in _tbs:
					if not tb or len(tb) < 2:
						continue
					# 会计政策变更的调整表（上年末 / 本年初 / 调整数）不是报表本身
					head = "".join(str(c or "") for r in tb[:3] for c in (r or [])).replace(" ", "")
					if "调整数" in head or re.search(r"20\d{2}年0?1月0?1日", head):
						continue
					years = find_year_cols(tb)
					if not years and tb[0]:
						# 表头没有年份（期末余额/期初余额）→ 用报告年度推断
						years = period_years(tb[0], report_year)
					if years:
						seg_years = years          # 本表自带表头 → 更新继承值
						seg_width = _used_width(tb)
					else:
						years = seg_years          # 续表 → 用首页的表头
						if years and _used_width(tb) != seg_width:
							# 续表列数不同（合并单元格拆法不一样）：按「数字列从左到右」重新对应年份
							# 只认金额列（附注号是小整数，排除）
							num_cols = sorted({j for r in tb for j, c in enumerate(r or [])
							                   if j > 0 and abs(tonum(c) or 0) >= 1000 and re.search(r"\d", str(c))})
							order = [yr for _c, yr in sorted(years.items())]
							if len(num_cols) >= len(order) and num_cols[:len(order)] != sorted(years):
								years = dict(zip(num_cols[:len(order)], order))
					if not years:
						continue
					for ri, row in enumerate(tb):
						if not row:
							continue
						_j, _txt = name_cell(row)
						if _txt is None:
							continue
						raw = clean(_txt)
						if not raw or len(raw) > 40:
							continue
						std = EXTRA_ALIAS.get(raw)
						if not std:
							continue
						# 科目名折行：本行没数字，数字在下一行（下一行名称格为空）
						if all(not str(row[c] or "").strip() for c in years if c < len(row)) and ri + 1 < len(tb):
							nx = tb[ri + 1] or []
							if not "".join(str(c or "") for c in nx[:min(years)]).strip():
								row = nx
						for col, yr in years.items():
							if col >= len(row) or row[col] is None or (isinstance(row[col], str) and not row[col].strip()):
								continue
							v = tonum(row[col])
							if v is None:
								continue
							facts.append((yr, std, v * mult))
							n_rows += 1
			# 文本行兜底：总是运行，但只补表格没拿到的 (年度, 科目)
			# —— 有些报表上下半张表结构不同，表格只提到一部分
			has = {(f[0], f[1]) for f in facts if f[2] != 0}
			txt_facts = parse_text_pages(pdf, [p for p, _c in seg], report_year, EXTRA_ALIAS, global_mult)
			added = 0
			for f_row in txt_facts:
				k = (f_row[0], f_row[1])
				if k in has:
					continue
				has.add(k)
				facts.append(f_row)
				added += 1
			if added:
				report.append("%s: 文本兜底补充 %d 条" % (name, added))
			report.append("%s: 定位 %d 页, 提取 %d 条" % (name, len(seg), n_rows))

	if company is None:
		company = os.path.splitext(os.path.basename(path))[0].split("：")[0]

	if report_year:
		facts = [f for f in facts if f[0] in (report_year, report_year - 1)]

	df = pd.DataFrame(facts, columns=BASE_COLS)
	df.insert(0, "公司", company)
	if not df.empty:
		# 去重：同一 科目/年度 保留绝对值最大的一条（避免合计行/续页重复）
		keep = df.groupby(["公司", "年度", "标准项目"])["金额"].apply(lambda s: s.abs().idxmax())
		df = df.loc[keep.values].reset_index(drop=True)
		k = _unit_check(path, df, report_year)
		if k != 1:
			df["金额"] = df["金额"] * k
			report.append("单位校正：按「主要会计数据」营业收入把报表金额 ×%g" % k)
	return df, report


def _unit_check(path, df, report_year):
	"""报表没写单位、或单位写在别处时，用前面「主要会计数据」表里的营业收入（元）对照，
	差 1000 / 10000 / 1000000 倍就是单位漏了 → 返回应乘的倍数"""
	if report_year is None or df.empty:
		return 1
	m = df[(df["年度"] == report_year) & (df["标准项目"].isin(["营业收入", "营业总收入"]))]
	if m.empty:
		return 1
	mine = float(m["金额"].abs().max())
	ps = pdftext.pages(path) or []
	for i in range(min(25, len(ps))):
		t = pdftext.flat(path, i) or ""
		if "主要会计数据" not in t:
			continue
		g = re.search(r"营业(?:总)?收入[（(](元|千元|万元|百万元)[）)]([\d,]{6,}(?:\.\d{1,2})?)", t)
		if not g:
			continue
		ref = (tonum(g.group(2)) or 0) * UNIT_MULT.get(g.group(1), 1)
		if not ref or not mine:
			return 1
		for k in (1000, 10000, 1000000):
			if abs(mine * k / ref - 1) < 0.02:
				return k
		return 1
	return 1


if __name__ == "__main__":
	p = sys.argv[1]
	df, rep = parse(p)
	print("\n".join(rep))
	print()
	print("总行数:", len(df))
	print("科目数:", df["标准项目"].nunique(), " 年度:", sorted(df["年度"].unique()))
	print()
	piv = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
	cols = [c for c in ("总资产", "总负债", "所有者权益", "营业收入", "净利润", "经营现金流净额") if c in piv.index]
	print((piv.loc[cols] / 1e8).round(2).to_string())


SIGN_POS = ("营业外支出", "所得税费用", "营业成本", "销售费用", "管理费用", "研发费用", "利息费用")


def _fix_signs(df):
	"""港式/国际版式把费用列成「(1,234)」→ 解析成负数；按恒等式能对上时翻成正数"""
	p = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
	for y in p.columns:
		def v(k):
			return float(p.loc[k, y]) if (k in p.index and pd.notna(p.loc[k, y])) else None
		flip = set()
		negate = set()
		tp, tax, npr = v("利润总额"), v("所得税费用"), v("净利润")
		if None not in (tp, tax, npr) and abs(tp - tax - npr) > 1:
			if tax < 0 and abs(tp - abs(tax) - npr) <= max(1.0, abs(npr) * 1e-6):
				flip.add("所得税费用")
			elif tax > 0 and abs(tp - (-tax) - npr) <= max(1.0, abs(npr) * 1e-6):
				negate.add("所得税费用")
		op, oi, oe = v("营业利润"), v("营业外收入"), v("营业外支出")
		if None not in (op, oi, oe, tp) and oe < 0 and abs(op + oi - abs(oe) - tp) <= max(1.0, abs(tp) * 1e-6):
			flip.add("营业外支出")
		rev = v("营业收入")
		for k in ("营业成本", "销售费用", "管理费用", "研发费用"):
			x = v(k)
			if x is not None and x < 0 and rev and rev > 0 and abs(x) <= rev * 1.5:
				flip.add(k)
		for k in flip:
			m = (df["年度"] == y) & (df["标准项目"] == k)
			df.loc[m, "金额"] = df.loc[m, "金额"].abs()
		for k in negate:
			m = (df["年度"] == y) & (df["标准项目"] == k)
			df.loc[m, "金额"] = -df.loc[m, "金额"].abs()
	return df


def finish(df):
	"""兜底：所有者权益缺失时用 资产-负债 推导"""
	if df.empty:
		return df
	df = _fix_signs(df)
	# 「一、营业总收入 / 其中：营业收入」版式里没有单列营业收入时，营业收入可能抓到别处的小数字；
	# 明显小于营业成本时，用营业总收入替代
	p = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
	for y in p.columns:
		def v(k):
			return float(p.loc[k, y]) if (k in p.index and pd.notna(p.loc[k, y])) else None
		rev, tot, cost = v("营业收入"), v("营业总收入"), v("营业成本")
		# 营业总收入不到营业收入一半、或营业收入不到营业总收入一半，都说明抓错了行
		# （证券公司「营业收入」是分段标题，数字被下一行的「手续费及佣金净收入」吃掉）
		if tot and tot > 0 and (rev is None or (cost and rev < cost * 0.5) or abs(rev) < abs(tot) * 0.5):
			m = (df["年度"] == y) & (df["标准项目"] == "营业收入")
			if m.any():
				df.loc[m, "金额"] = tot
			else:
				row_d = {"公司": df["公司"].iloc[0], "年度": y, "标准项目": "营业收入", "金额": tot}
				df = pd.concat([df, pd.DataFrame([row_d])], ignore_index=True)
	p = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
	rows = []
	for y in p.columns:
		def v(k):
			return float(p.loc[k, y]) if (k in p.index and pd.notna(p.loc[k, y])) else None
		co = df["公司"].iloc[0]
		# 合计行折行 / 和别的列挤在一行没读到时，用同表的分项推出来
		if v("经营现金流净额") is None and None not in (v("经营现金流入"), v("经营现金流出")):
			rows.append({"公司": co, "年度": y, "标准项目": "经营现金流净额",
			             "金额": v("经营现金流入") - abs(v("经营现金流出"))})
		a, l, e = v("总资产"), v("总负债"), v("所有者权益")
		if l is None and None not in (v("流动负债"), v("非流动负债")):
			l = v("流动负债") + v("非流动负债")
			rows.append({"公司": co, "年度": y, "标准项目": "总负债", "金额": l})
		if l is None and a is not None and e is not None:
			l = a - e
			rows.append({"公司": co, "年度": y, "标准项目": "总负债", "金额": l})
		if e is None and a is not None and l is not None:
			rows.append({"公司": co, "年度": y, "标准项目": "所有者权益", "金额": a - l})
	if rows:
		df = pd.concat([df, pd.DataFrame(rows)], ignore_index=True)

	# 「所有者权益」可能抓到别处的同名行（如权益变动表）；
	# 归母权益 + 少数股东权益 是同表相邻两行，更可靠，且能用资产-负债交叉验证
	p = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
	fix = []
	for y in p.columns:
		def v(k):
			return float(p.loc[k, y]) if (k in p.index and pd.notna(p.loc[k, y])) else None
		a, l, e = v("总资产"), v("总负债"), v("所有者权益")
		parent, minority = v("归母权益"), v("少数股东权益")
		if parent is None or minority is None or a is None or l is None:
			continue
		derived = parent + minority
		if abs(derived - (a - l)) > max(1.0, abs(a) * 0.005):
			continue                      # 归母+少数 本身对不上资产负债表，不动
		if e is None or abs(e - derived) > 1:
			fix.append((y, derived))
	for y, val in fix:
		mask = (df["年度"] == y) & (df["标准项目"] == "所有者权益")
		if mask.any():
			df.loc[mask, "金额"] = val
		else:
			row_d = {"公司": df["公司"].iloc[0], "年度": y, "标准项目": "所有者权益", "金额": val}
			df = pd.concat([df, pd.DataFrame([row_d])], ignore_index=True)
	return df


def validate(df):
	"""三大恒等式校验 -> [(说明, 是否通过, 差额)]"""
	p = df.pivot_table(index="标准项目", columns="年度", values="金额", aggfunc="first")
	out = []
	for y in p.columns:
		def v(k):
			return float(p.loc[k, y]) if (k in p.index and pd.notna(p.loc[k, y])) else None
		# 1) 资产 = 负债 + 所有者权益
		a, l, e = v("总资产"), v("总负债"), v("所有者权益")
		if None not in (a, l, e):
			d = a - l - e
			out.append(("%d 资产 = 负债 + 权益" % y, abs(d) <= max(1.0, abs(a) * 1e-6), d))
		# 2) 净利润 = 利润总额 - 所得税
		tp, tax, npr = v("利润总额"), v("所得税费用"), v("净利润")
		if None not in (tp, tax, npr):
			d = tp - tax - npr
			out.append(("%d 净利润 = 利润总额 - 所得税" % y, abs(d) <= max(1.0, abs(npr) * 1e-6), d))
		# 3) 利润总额 = 营业利润 + 营业外收入 - 营业外支出
		op, oi, oe2, tp2 = v("营业利润"), v("营业外收入"), v("营业外支出"), v("利润总额")
		if None not in (op, oi, oe2, tp2):
			d = op + oi - oe2 - tp2
			out.append(("%d 利润总额 = 营业利润 + 营业外收支净额" % y, abs(d) <= max(1.0, abs(tp2) * 1e-6), d))
	return out

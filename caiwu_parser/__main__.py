# -*- coding: utf-8 -*-
"""
命令行入口：
    python -m caiwu_parser 某某公司_2025年报.pdf
    python -m caiwu_parser 某某公司_2025年报.pdf --unit yi --csv out.csv
"""
import argparse
import sys
from . import parse_report, pivot_report

_UNIT_MAP = {
    "yuan": (1.0, "元"),
    "wan": (1e4, "万元"),
    "yi": (1e8, "亿元"),
}


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="caiwu-parser",
        description="A 股财报三大报表高鲁棒性解析器（支持原生 PDF / Excel / Word）",
    )
    parser.add_argument("file", help="年报或财务报表路径（.pdf / .xlsx / .docx）")
    parser.add_argument("--company", default=None, help="指定公司名称（默认从文件名提取）")
    parser.add_argument(
        "--fin",
        action="store_true",
        default=None,
        help="强制按金融企业（银行/证券/保险）科目体系解析",
    )
    parser.add_argument(
        "--unit",
        choices=["yuan", "wan", "yi"],
        default="yi",
        help="终端打印金额单位：yuan=元, wan=万元, yi=亿元（默认: yi）",
    )
    parser.add_argument("--csv", default=None, help="将标准长表导出为 UTF-8 CSV 文件路径")
    args = parser.parse_args(argv)

    res = parse_report(args.file, company=args.company, is_fin=args.fin)
    scale, unit_label = _UNIT_MAP[args.unit]

    print(f"=== {res['company']} ({'金融企业' if res['is_financial'] else '一般工商业'}) ===")
    for log in res["logs"]:
        print(f"  · {log}")

    df = res["data"]
    if df.empty:
        print("\n[!] 未提取到有效财务报表数据。")
        return 1

    piv = pivot_report(df, unit=scale, decimals=2 if scale > 1 else 2)
    print(f"\n--- 核心报表科目（单位：{unit_label}）---")
    print(piv.to_string())

    print("\n--- 会计恒等式门禁校验 ---")
    for rule, ok, diff in res["checks"]:
        mark = "PASS" if ok else "FAIL"
        print(f"  [{mark}] {rule}  (差额: {diff:,.2f} 元)")
    print(f"\n整体门禁结果: {'通过 (VALID)' if res['is_valid'] else '未全部通过 (CHECK REQUIRED)'}")

    if args.csv:
        df.to_csv(args.csv, index=False, encoding="utf-8-sig")
        print(f"\n已导出标准长表至: {args.csv}")

    return 0 if res["is_valid"] else 2


if __name__ == "__main__":
    sys.exit(main())

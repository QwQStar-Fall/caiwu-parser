# -*- coding: utf-8 -*-
"""
caiwu-parser 快速上手示例

运行方式：
    python examples/quickstart.py path/to/公司年报_2025.pdf
"""
import sys
from caiwu_parser import parse_report, pivot_report


def main():
    if len(sys.argv) < 2:
        print("用法: python examples/quickstart.py <年报PDF/Excel/Word路径>")
        return

    file_path = sys.argv[1]
    res = parse_report(file_path)

    print(f"公司名称: {res['company']}")
    print(f"金融企业: {'是' if res['is_financial'] else '否'}")
    print(f"门禁通过: {res['is_valid']}")
    print("\n解析日志:")
    for log in res["logs"]:
        print("  -", log)

    print("\n会计恒等式校验:")
    for rule, ok, diff in res["checks"]:
        status = "PASS" if ok else "FAIL"
        print(f"  [{status}] {rule} (差额: {diff:,.2f} 元)")

    print("\n三大报表核心科目透视表（单位：亿元）:")
    table = pivot_report(res, unit=1e8, decimals=2)
    print(table.to_string())


if __name__ == "__main__":
    main()

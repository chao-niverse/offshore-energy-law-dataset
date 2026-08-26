# -*- coding: utf-8 -*-
"""
batch_validate.py — 批量运行校验脚本 v1.0
用法: python batch_validate.py 收件文件夹 --master 规范清单母表.xlsx
对文件夹内所有 .xlsx 逐一运行 validate_entries.py,报告存于同文件夹,
并输出 汇总.csv(文件名、条目数、错误、提醒、合规率)。
"""
import argparse, csv, os, re, subprocess, sys

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--master", required=True)
    a = ap.parse_args()
    here = os.path.dirname(os.path.abspath(__file__))
    validator = os.path.join(here, "validate_entries.py")
    rows = []
    files = sorted(f for f in os.listdir(a.folder)
                   if f.endswith(".xlsx") and not f.startswith("~$") and "校验报告" not in f)
    for fn in files:
        p = os.path.join(a.folder, fn)
        out = p.replace(".xlsx", "_校验报告.xlsx")
        try:
            res = subprocess.run([sys.executable, validator, p, "--master", a.master, "-o", out],
                                 capture_output=True, text=True, timeout=120)
            m = re.search(r"条目(\d+) \| 错误(\d+) 提醒(\d+) \| 合规率([\d.]+)%", res.stdout)
            if m:
                rows.append([fn, m.group(1), m.group(2), m.group(3), m.group(4) + "%", os.path.basename(out)])
                print(f"{fn}: 条目{m.group(1)} 错误{m.group(2)} 提醒{m.group(3)} 合规率{m.group(4)}%")
            else:
                rows.append([fn, "-", "-", "-", "运行异常", res.stderr.strip()[:80]])
                print(f"{fn}: 运行异常")
        except Exception as e:
            rows.append([fn, "-", "-", "-", "失败", str(e)[:80]])
    summary = os.path.join(a.folder, "汇总.csv")
    with open(summary, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["文件名", "条目数", "错误", "提醒", "合规率", "报告文件"])
        w.writerows(rows)
    print(f"\n共 {len(files)} 份文件,汇总: {summary}")

if __name__ == "__main__":
    main()

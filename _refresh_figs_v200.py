# -*- coding: utf-8 -*-
r"""把含「配置核查作业指导书v2.2」字样的截图按原参数重出为 v2.0.0。

做法：不改任何工具脚本，只复制源 HTML 到临时副本，替换版本文案
（OS 类报告另去掉模板冗余的那行「参考标准」），再用 _gen_report_figs 的
_shoot 以原窗口尺寸/像素比重新截图，覆盖原图。

用法：python _refresh_figs_v200.py
"""
import io
import pathlib
import sys

BASE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
from _gen_report_figs import _shoot  # noqa: E402

OLD, NEW = "配置核查作业指导书v2.2", "配置核查作业指导书v2.0.0"

# (输出目录, 输出名, 源 HTML, 宽, 高, 像素比, 是否去掉冗余的独立参考标准行)
JOBS = [
    ("测评报告/验证截图", "fig_kylin_os",    "kylin/output/配置核查报告_麒麟靶机_20260911_111748.html",    1400, 1250, 1, True),
    ("测评报告/验证截图", "fig_kylin_mysql", "kylin/output/配置核查报告_MySQL_麒麟靶机_20260911_111717.html", 1400, 1100, 1, False),
    ("测评报告/验证截图", "fig_kylin_nginx", "kylin/output/配置核查报告_Nginx_麒麟靶机_20260911_105104.html", 1400, 900, 1, False),
    ("测评报告/验证截图", "fig_kylin_net",   "kylin/output/配置核查报告_网络设备_20260905_133025.html",     1400, 1250, 1, False),
    ("测评报告/验证截图", "fig_kylin_redis", "kylin/output/配置核查报告_Redis_20260911_151109.html",       1400, 860, 1, False),
    ("测评报告/说明书截图", "op_kylin_report", "kylin/output/配置核查报告_麒麟靶机_20260911_111748.html",   1400, 860, 2, True),
    ("测评报告/验证截图", "fig_manual",      "manual_check.html",                                       1400, 1250, 1, False),
]


def main():
    for outdir, name, rel, w, h, scale, drop_dup in JOBS:
        src = BASE / rel
        if not src.exists():
            print("!! 源文件缺失，跳过:", rel)
            continue
        s = io.open(src, encoding="utf-8", errors="replace").read()
        n = s.count(OLD)
        s = s.replace(OLD, NEW)
        dropped = 0
        if drop_dup:
            lines = s.split("\n")
            keep = []
            for ln in lines:
                if ln.strip() == "<div>参考标准：%s</div>" % NEW:
                    dropped += 1
                    continue
                keep.append(ln)
            s = "\n".join(keep)
        tmp = src.with_name(src.stem + "_v200tmp.html")
        io.open(tmp, "w", encoding="utf-8", newline="").write(s)
        out = BASE / outdir / (name + ".png")
        try:
            _shoot(tmp, out, w, h, scale)
            print("   %-16s 版本文案 %d 处，去冗余行 %d 行 → %s" % (name, n, dropped, outdir))
        finally:
            if tmp.exists():
                tmp.unlink()
    print("完成")


if __name__ == "__main__":
    main()

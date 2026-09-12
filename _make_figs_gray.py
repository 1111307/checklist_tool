# -*- coding: utf-8 -*-
r"""把报告配图转为黑白（灰度），使文档中除黑色/灰色外不再出现其他颜色。

原彩色图备份到仓库外 `Desktop\peizhitool\_figs_color_backup\<目录名>\`，可随时还原。

用法：
  python _make_figs_gray.py 测评报告/验证截图        # 转换该目录下所有 png
  python _make_figs_gray.py 测评报告/验证截图 --restore   # 从备份还原
"""
import os
import shutil
import sys

from PIL import Image, ImageOps

BASE = os.path.dirname(os.path.abspath(__file__))
BACKUP_ROOT = os.path.join(os.path.dirname(BASE), "_figs_color_backup")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    target = sys.argv[1].replace("/", os.sep)
    restore = "--restore" in sys.argv
    src_dir = os.path.join(BASE, target)
    bak_dir = os.path.join(BACKUP_ROOT, target.replace(os.sep, "_"))
    if not os.path.isdir(src_dir):
        print("目录不存在:", src_dir)
        return 2

    if restore:
        n = 0
        for name in os.listdir(bak_dir):
            if name.lower().endswith(".png"):
                shutil.copy(os.path.join(bak_dir, name), os.path.join(src_dir, name))
                n += 1
        print("已从备份还原 %d 张 ->" % n, src_dir)
        return 0

    os.makedirs(bak_dir, exist_ok=True)
    n = 0
    for name in sorted(os.listdir(src_dir)):
        if not name.lower().endswith(".png"):
            continue
        p = os.path.join(src_dir, name)
        if not os.path.exists(os.path.join(bak_dir, name)):
            shutil.copy(p, os.path.join(bak_dir, name))     # 首次转换才备份原图
        with Image.open(p) as im:
            gray = ImageOps.grayscale(im).convert("RGB")
        gray.save(p)
        n += 1
    print("已转为黑白 %d 张（原图备份在 %s）" % (n, bak_dir))
    return 0


if __name__ == "__main__":
    sys.exit(main())

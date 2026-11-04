# -*- coding: utf-8 -*-
"""评审整改第 7 条：数据库核查报告标注「数据保存在哪里」。

评审原话：「数据库的核查报告在大表中说明数据保存在哪里」。做法按交接文档的设计——
连接成功后采集数据目录，写进报告表头下方的 meta 区（HTML 的 meta 行、XLS 的连接行），
**不新增核查项**，136 项口径不变。

采集口径
--------
    MySQL/MariaDB  SHOW VARIABLES LIKE 'datadir'
    Redis          CONFIG GET dir
    达梦 DM8       SELECT PATH FROM V$DATAFILE
    SQL Server     SELECT TOP 1 physical_name FROM sys.master_files

覆盖 8 个脚本（kylin/*.sh 4 个 + win/*.vbs 4 个）。幂等：文件里已有 DATA_DIR 就跳过。

用法
----
    python _patch_datadir.py            # 干跑，打印将改的位置
    python _patch_datadir.py --apply    # 落盘（.tmp + 原子替换，GBK 先预检）
"""
import os
import sys

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))

DM_FUNC = '''Function GetDmDataFile()
    Dim r : r = DmQ("SELECT PATH FROM V$DATAFILE;")
    Dim lines : lines = Split(Replace(r, Chr(13), ""), Chr(10))
    Dim i, n, line, parts, first
    n = 0 : first = ""
    For i = 0 To UBound(lines)
        line = Trim(Replace(lines(i), Chr(9), " "))
        If line <> "" And IsNumeric(Left(line, 1)) Then
            parts = Split(line, " ", 2)
            If UBound(parts) >= 1 Then
                n = n + 1
                If n = 1 Then first = Trim(parts(1))
            End If
        End If
    Next
    If n = 0 Then
        GetDmDataFile = ""
    ElseIf n > 1 Then
        GetDmDataFile = first & "（共 " & n & " 个数据文件）"
    Else
        GetDmDataFile = first
    End If
End Function'''


def load(path, enc):
    s = open(path, "rb").read().decode(enc)
    eol = "\r\n" if "\r\n" in s else "\n"
    return s.replace("\r\n", "\n"), eol


def save(path, s, enc, eol):
    out = s.replace("\n", eol)
    out.encode(enc)  # GBK 预检：编不了就直接抛，不落盘
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(out.encode(enc))
    os.replace(tmp, path)


def patch_sh(rel, collect):
    """麒麟侧：声明 / 采集 / meta 行 / xls 行。"""
    path = os.path.join(ROOT, rel)
    s, eol = load(path, "utf-8")
    if "DATA_DIR" in s:
        return rel, "已含 DATA_DIR，跳过", []
    L = s.split("\n")
    log, out = [], []
    for line in L:
        out.append(line)
        if line.strip() == 'DB_VERSION=""' and "DATA_DIR" not in "".join(out[-3:]):
            out.append('DATA_DIR=""')
            log.append("声明 DATA_DIR")
    L, out = out, []
    for line in L:
        out.append(line)
        if "CONN_OK=1" in line and not line.lstrip().startswith("#"):
            ind = line[: len(line) - len(line.lstrip())]
            out.append(f"{ind}{collect}")
            out.append(f'{ind}[ -z "$DATA_DIR" ] && DATA_DIR="未采集"')
            out.append(f'{ind}echo "  数据目录：$DATA_DIR"')
            log.append("采集数据目录")
    L, out = out, []
    for line in L:
        if line.strip().startswith('echo "<div>主机名：'):
            out.append('    echo "<div>数据目录：$(html_esc "${DATA_DIR:-未采集}")</div>"')
            log.append("HTML meta 行")
        if '版本：$(html_esc "${DB_VERSION:-未连接}")　核查时间：' in line:
            line = line.replace('版本：$(html_esc "${DB_VERSION:-未连接}")',
                                '版本：$(html_esc "${DB_VERSION:-未连接}")　数据目录：$(html_esc "${DATA_DIR:-未采集}")')
            log.append("XLS 表头行")
        out.append(line)
    return rel, None, (out, log, eol, "utf-8")


def patch_vbs(rel, collect, need_func=False, ver="DB_VERSION"):
    path = os.path.join(ROOT, rel)
    s, eol = load(path, "gbk")
    if "DATA_DIR" in s:
        return rel, "已含 DATA_DIR，跳过", []
    L = s.split("\n")
    log, out = [], []
    # 1. 变量声明
    for line in L:
        out.append(line)
        if line.strip().startswith(f"Dim {ver} :"):
            out.append('Dim DATA_DIR : DATA_DIR = ""')
            log.append("声明 DATA_DIR")
    L = out
    # 2. DM 辅助函数
    if need_func:
        out = []
        inserted = False
        in_dmq = False
        for line in L:
            out.append(line)
            if line.startswith("Function DmQ("):
                in_dmq = True
            elif in_dmq and line.strip() == "End Function":
                out.append("")
                out.extend(DM_FUNC.split("\n"))
                in_dmq = False
                inserted = True
                log.append("插入 GetDmDataFile()")
        if not inserted:
            return rel, "[FAIL] 没找到 DmQ 的 End Function", []
        L = out
    # 3. 采集 + 兜底 + 控制台回显
    out = []
    for line in L:
        out.append(line)
        if line.strip() == "CONN_OK = True":
            ind = line[: len(line) - len(line.lstrip())]
            out.append(f"{ind}{collect}")
            out.append(f'{ind}If DATA_DIR = "" Then DATA_DIR = "未采集"')
            out.append(f'{ind}WScript.Echo "  数据目录：" & DATA_DIR')
            log.append("采集数据目录")
    L = out
    # 4. 三处输出
    out = []
    for line in L:
        if f'　版本：" & HtmlEsc({ver}) & "　核查时间：' in line:
            line = line.replace(f'　版本：" & HtmlEsc({ver}) & "　核查时间：',
                                f'　版本：" & HtmlEsc({ver}) & "　数据目录：" & HtmlEsc(DATA_DIR) & "　核查时间：')
            log.append("HTML meta 行")
        if f'  版本：" & {ver}' in line and "Cells" in line:
            line = line.replace(f'  版本：" & {ver}', f'  版本：" & {ver} & "  数据目录：" & DATA_DIR')
            log.append("XLS 表头行")
        if f'  版本：" & HtmlEsc({ver}) & "</td>' in line:
            line = line.replace(f'  版本：" & HtmlEsc({ver}) & "</td>',
                                f'  版本：" & HtmlEsc({ver}) & "  数据目录：" & HtmlEsc(DATA_DIR) & "</td>')
            log.append("HTML 表格信息行")
        out.append(line)
    return rel, None, (out, log, eol, "gbk")


def main():
    apply = "--apply" in sys.argv

    SH = [
        ("kylin/check_mysql.sh", """DATA_DIR="$(mysql_var 'datadir')" 2>/dev/null"""),
        ("kylin/check_redis.sh", '''DATA_DIR="$(redis_cfg dir)" 2>/dev/null'''),
        ("kylin/check_dm.sh",
         """DATA_DIR="$(dm_q "SELECT PATH FROM V\\$DATAFILE;" 2>/dev/null | awk 'NF>=2 && $1 ~ /^[0-9]+$/ {sub(/^[ \\t]*[0-9]+[ \\t]+/, ""); print; exit}')" """),
        ("kylin/check_sqlserver.sh",
         """DATA_DIR="$(sql_q "SELECT TOP 1 physical_name FROM sys.master_files;" 2>/dev/null | awk 'NF' | head -1)" """),
    ]
    VBS = [
        ("win/check_mysql.vbs", 'DATA_DIR = Trim(GetMysqlVar("datadir"))', False, "DB_VERSION"),
        ("win/check_redis.vbs", 'DATA_DIR = Trim(GetRedisCfg("dir"))', False, "REDIS_VERSION"),
        ("win/check_dm.vbs", "DATA_DIR = GetDmDataFile()", True, "DB_VERSION"),
        ("win/check_sqlserver.vbs",
         'DATA_DIR = Trim(InvokeSql("SELECT TOP 1 physical_name FROM sys.master_files;"))', False, "DB_VERSION"),
    ]

    results = []
    for rel, collect in SH:
        results.append(patch_sh(rel, collect.strip()))
    for rel, collect, need_func, ver in VBS:
        results.append(patch_vbs(rel, collect, need_func, ver))

    bad = 0
    for rel, skip, payload in results:
        print(f"=== {rel}")
        if skip:
            print(f"    {skip}")
            continue
        out, log, eol, enc = payload
        for x in log:
            print(f"    + {x}")
        if not log:
            print("    [FAIL] 一处都没匹配到")
            bad += 1
            continue
        if apply:
            save(os.path.join(ROOT, rel), "\n".join(out), enc, eol)
            print(f"    已写入（{enc}, {'CRLF' if eol == chr(13) + chr(10) else 'LF'}）")
    if not apply:
        print("\n[干跑] 加 --apply 落盘")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())

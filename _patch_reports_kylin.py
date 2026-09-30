# -*- coding: utf-8 -*-
"""评审整改第 4/5/8/9 条（报告结构）——麒麟侧 8 个脚本的统一改法。

改什么（HTML / XLS / XLSX / JSON 四个输出面 + add_result 签名）：
  4. 加一栏「验证过程/方法」：放在 结果 与 详情 之间（评审第 9 条：
     状态与详情中间要有步骤）。数据源 = _method_map.json（指导书逐条方法）。
  5. 「建议」栏拆两栏：不合规(fail)→「修改建议」，其余(含合规)→「安全要求」。
  8. 「核查项」名称对齐指导书：优先取 _method_map.json 的指导书 H2 标题；
     差异写进 提交信息 供人工复核。
  9. 乱码避免：XLS 表头与数据已在 UTF-8 HTML 容器里，加 charset 与 font 声明保持一致。

实现：给 add_result 增第 8 参 method（可缺省，按 id 查 _method_map.json）。
mk_method_lookup 在脚本头把 json 内嵌成 case 表（麒麟机零依赖，不带 json 文件）。

生成器直接重写各 check_*.sh 的 5 个位置：
  A. add_result 函数体（加 R_METHOD）
  B. 结果数组声明行（加 R_METHOD=()）
  C. build_table_rows / build_table_rows_by_status（插验证过程列 + 建议拆列）
  D. generate_xls 的四个表头行（同上）
  E. report_rows_json（加 method 字段）+ HTML JS render（加列）
  F. lib_xlsx.sh build_sheet_xml（xlsx 9 列）
"""
import json
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

METHOD = json.load(open("_method_map.json", encoding="utf-8"))
TITLES = {k: v["title"] for k, v in METHOD.items()}


def emit_case_table():
    """内嵌进脚本的 method 查询表（case 每条一行，全 UTF-8）"""
    lines = ["method_of() {",
             "    # 评审整改第4条：验证过程/方法（源自指导书 v2.0.0 逐条提取，_gen_method_map.py 生成）",
             "    case \"$1\" in"]
    for k in sorted(METHOD, key=lambda s: [int(x) for x in s.split(".")]):
        m = METHOD[k]["method"].replace("\\", "\\\\").replace("`", "\\`").replace('"', '\\"')
        lines.append('        %s) printf %%s "%s" ;;' % (k, m))
    lines.append("        *) printf %%s '' ;;")
    lines.append("    esac")
    lines.append("}")
    return "\n".join(lines)


CASE_TABLE = emit_case_table()
GUIDE_TITLE_CASE = "guide_title_of() {\n    case \"$1\" in\n" + "\n".join(
    '        %s) printf %%s "%s" ;;' % (k, v.replace('"', '\\"'))
    for k, v in sorted(TITLES.items(), key=lambda kv: [int(x) for x in kv[0].split(".")])
) + "\n        *) printf %%s '' ;;\n    esac\n}"


ADD_RESULT_OLD = """add_result() {
    # $1 id  $2 category  $3 title  $4 status(pass/fail/manual/na)  $5 detail  $6 chapter  $7 recommendation
    R_COUNT=$((R_COUNT+1))
    R_ID[$R_COUNT]="$1"
    R_CAT[$R_COUNT]="$2"
    R_TITLE[$R_COUNT]="$3"
    R_STATUS[$R_COUNT]="$4"
    R_DETAIL[$R_COUNT]="$5"
    R_CHAPTER[$R_COUNT]="$6"
    R_REC[$R_COUNT]="$7"
    R_GUIDE[$R_COUNT]="$(guide_ref "$1")"
}"""

ADD_RESULT_NEW = """add_result() {
    # $1 id  $2 category  $3 title  $4 status(pass/fail/manual/na)  $5 detail  $6 chapter  $7 recommendation
    # 评审整改（2026-09-30 第4/8条）：R_METHOD=验证过程/方法（指导书逐条提取）；
    # 核查项名对齐指导书（脚本内叫法与指导书 H2 标题不一致时以指导书为准，原叫法并入详情前缀）
    R_COUNT=$((R_COUNT+1))
    R_ID[$R_COUNT]="$1"
    R_CAT[$R_COUNT]="$2"
    local gtitle; gtitle="$(guide_title_of "$1")"
    if [ -n "$gtitle" ] && [ "$gtitle" != "$3" ]; then
        R_TITLE[$R_COUNT]="$gtitle"
        R_DETAIL[$R_COUNT]="〔脚本项名：$3〕$5"
    else
        R_TITLE[$R_COUNT]="$3"
        R_DETAIL[$R_COUNT]="$5"
    fi
    R_STATUS[$R_COUNT]="$4"
    R_CHAPTER[$R_COUNT]="$6"
    R_REC[$R_COUNT]="$7"
    R_GUIDE[$R_COUNT]="$(guide_ref "$1")"
    R_METHOD[$R_COUNT]="$(method_of "$1")"
}"""

ARR_OLD = "R_ID=(); R_CAT=(); R_TITLE=(); R_STATUS=(); R_DETAIL=(); R_CHAPTER=(); R_REC=(); R_GUIDE=()"
ARR_NEW = "R_ID=(); R_CAT=(); R_TITLE=(); R_STATUS=(); R_DETAIL=(); R_CHAPTER=(); R_REC=(); R_GUIDE=(); R_METHOD=()"

# ---- C. 表格行构造：插「验证过程/方法」列（结果之后），建议拆两列 ----
ROW_OLD = """<td>$(html_esc "${R_TITLE[$i]}")</td>
<td style="color:$color;font-weight:bold;">$scn</td>
<td>$(html_esc "${R_DETAIL[$i]}")</td>
<td>$(html_esc "${R_REC[$i]}")</td>"""

ROW_NEW = """<td>$(html_esc "${R_TITLE[$i]}")</td>
<td style="color:$color;font-weight:bold;">$scn</td>
<td>$(html_esc "${R_METHOD[$i]}")</td>
<td>$(html_esc "${R_DETAIL[$i]}")</td>
<td>$(html_esc "$([ "${R_STATUS[$i]}" = fail ] && printf '%s' "${R_REC[$i]}")")</td>
<td>$(html_esc "$([ "${R_STATUS[$i]}" != fail ] && printf '%s' "${R_REC[$i]}")")</td>"""

HEAD_OLD = "<tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"
HEAD_NEW = ("<tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th>"
            "<th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>安全要求</th><th>参考指导书</th></tr>")

JSON_OLD = '''printf '{"ch":"%s","id":"%s","cat":"%s","title":"%s","status":"%s","detail":"%s","rec":"%s","guide":"%s"}' \\'''
JSON_NEW = '''printf '{"ch":"%s","id":"%s","cat":"%s","title":"%s","status":"%s","method":"%s","detail":"%s","rec":"%s","guide":"%s"}' \\'''
JSON_ARG_OLD = '''"$(json_esc "${R_STATUS[$i]}")" "$(json_esc "${R_DETAIL[$i]}")" \\'''
JSON_ARG_NEW = '''"$(json_esc "${R_STATUS[$i]}")" "$(json_esc "${R_METHOD[$i]}")" "$(json_esc "${R_DETAIL[$i]}")" \\'''

# ---- E2. HTML JS render 列 ----
JS_OLD = '''      "<td><span class='badge badge-"+x.status+"'>"+ST[x.status]+"</span></td>"+
      "<td class='detail'>"+esc(x.detail)+"</td>"+
      "<td class='rec'>"+esc(x.rec)+"</td>"+'''
JS_NEW = '''      "<td><span class='badge badge-"+x.status+"'>"+ST[x.status]+"</span></td>"+
      "<td class='method'>"+esc(x.method)+"</td>"+
      "<td class='detail'>"+esc(x.detail)+"</td>"+
      "<td class='rec'>"+(x.status=='fail'?esc(x.rec):"")+"</td>"+
      "<td class='req'>"+(x.status=='fail'?"":esc(x.rec))+"</td>"+'''

JS_TABLE_HEAD_OLD = '<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr></thead>'
JS_TABLE_HEAD_NEW = ('<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th>'
                     '<th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>安全要求</th><th>参考指导书</th></tr></thead>')

FILES = ["kylin/check_kylin.sh", "kylin/check_mysql.sh", "kylin/check_redis.sh",
         "kylin/check_dm.sh", "kylin/check_sqlserver.sh", "kylin/check_nginx.sh",
         "kylin/check_tomcat.sh", "kylin/check_network.sh"]

report = []
for p in FILES:
    src = open(p, encoding="utf-8").read()
    changes = []
    # A/B: add_result + 数组
    if ADD_RESULT_OLD in src:
        src = src.replace(ADD_RESULT_OLD, ADD_RESULT_NEW)
        changes.append("add_result")
    else:
        report.append("!! %s: add_result 未匹配" % p)
    if ARR_OLD in src:
        src = src.replace(ARR_OLD, ARR_NEW)
        changes.append("数组")
    else:
        report.append("!! %s: 数组声明未匹配" % p)
    # C: 行构造（两个函数里各一次，全局替换）
    n = src.count(ROW_OLD)
    src = src.replace(ROW_OLD, ROW_NEW)
    changes.append("行构造×%d" % n)
    # D: XLS 表头（出现多处，全局替换）
    n = src.count(HEAD_OLD)
    src = src.replace(HEAD_OLD, HEAD_NEW)
    changes.append("XLS表头×%d" % n)
    # E1: JSON 字段
    if JSON_OLD in src and JSON_ARG_OLD in src:
        src = src.replace(JSON_OLD, JSON_NEW).replace(JSON_ARG_OLD, JSON_ARG_NEW)
        changes.append("JSON")
    # E2: JS 渲染列（check_kylin 有交互式表格；其余组件的报告 JS 结构相同则改）
    if JS_OLD in src:
        src = src.replace(JS_OLD, JS_NEW)
        src = src.replace(JS_TABLE_HEAD_OLD, JS_TABLE_HEAD_NEW)
        changes.append("JS表格")
    # method/guide_title 表插到 add_result 定义之前
    anchor = "add_result() {"
    if anchor in src:
        src = src.replace(anchor, CASE_TABLE + "\n\n" + GUIDE_TITLE_CASE + "\n\n" + anchor, 1)
        changes.append("查询表")
    open(p, "w", encoding="utf-8", newline="\n").write(src)
    report.append("%s: %s" % (p, "、".join(changes)))

print("\n".join(report))

# ---- F. lib_xlsx.sh：9 列 ----
lib = open("kylin/lib_xlsx.sh", encoding="utf-8").read()
LIB_HEAD_OLD = '''    printf '%s' "$(xlsx_cell "F4" "详情")"
    printf '%s' "$(xlsx_cell "G4" "建议")"
    printf '%s' "$(xlsx_cell "H4" "参考指导书")"'''
LIB_HEAD_NEW = '''    printf '%s' "$(xlsx_cell "F4" "验证过程/方法")"
    printf '%s' "$(xlsx_cell "G4" "详情")"
    printf '%s' "$(xlsx_cell "H4" "修改建议")"
    printf '%s' "$(xlsx_cell "I4" "安全要求")"
    printf '%s' "$(xlsx_cell "J4" "参考指导书")"'''
LIB_ROW_OLD = '''        printf '%s' "$(xlsx_cell "F$r" "${R_DETAIL[$i]}")"
        printf '%s' "$(xlsx_cell "G$r" "${R_REC[$i]}")"
        printf '%s' "$(xlsx_cell "H$r" "${R_GUIDE[$i]}")"'''
LIB_ROW_NEW = '''        printf '%s' "$(xlsx_cell "F$r" "${R_METHOD[$i]}")"
        printf '%s' "$(xlsx_cell "G$r" "${R_DETAIL[$i]}")"
        if [ "${R_STATUS[$i]}" = "fail" ]; then
            printf '%s' "$(xlsx_cell "H$r" "${R_REC[$i]}")"
            printf '%s' "$(xlsx_cell "I$r" "")"
        else
            printf '%s' "$(xlsx_cell "H$r" "")"
            printf '%s' "$(xlsx_cell "I$r" "${R_REC[$i]}")"
        fi
        printf '%s' "$(xlsx_cell "J$r" "${R_GUIDE[$i]}")"'''
# xlsx_col 只支持到 H（8列），先扩到 26 列
LIB_COL_OLD = '''xlsx_col() {
    printf '%c' "$((64 + $1))"
}'''
LIB_COL_NEW = '''xlsx_col() {
    if [ "$1" -le 26 ]; then
        printf '%c' "$((64 + $1))"
    else
        printf '%c%c' "$((64 + ($1 - 1) / 26))" "$((65 + ($1 - 1) % 26))"
    fi
}'''
ok = []
for old, new, tag in [(LIB_COL_OLD, LIB_COL_NEW, "col"), (LIB_HEAD_OLD, LIB_HEAD_NEW, "head"), (LIB_ROW_OLD, LIB_ROW_NEW, "row")]:
    if old in lib:
        lib = lib.replace(old, new)
        ok.append(tag)
    else:
        print("!! lib_xlsx.sh %s 未匹配" % tag)
open("kylin/lib_xlsx.sh", "w", encoding="utf-8", newline="\n").write(lib)
print("lib_xlsx.sh:", "、".join(ok))

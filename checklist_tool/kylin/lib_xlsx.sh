#!/bin/bash
# ============================================================
# 零依赖真 .xlsx 生成器（纯 Bash + zip，无需 Python/openpyxl）
# ============================================================
# 原理：.xlsx 本质是 ZIP 包内若干 XML 文件，用系统 zip 命令即可打包。
# 依赖：zip 命令（麒麟/中标麒麟几乎必带；若缺失：yum install -y zip）
#
# 用法（在核查脚本内）：
#   source "$SCRIPT_DIR/lib_xlsx.sh"
#   generate_xlsx "$OUT_DIR/配置核查报告_MySQL_${STAMP}.xlsx" "MySQL 数据库配置核查报告"
#
# 依赖全局结果数组（由各核查脚本的 add_result 维护）：
#   R_CHAPTER/R_ID/R_CAT/R_TITLE/R_STATUS/R_DETAIL/R_REC/R_GUIDE + R_COUNT
# ============================================================

# XML 转义（含换行处理，便于单元格内换行）
# 注意：bash 参数替换中 replacement 里的 & 有特殊含义，必须写成 \&
xml_esc() {
    local s="${1-}"
    s="${s//&/\&amp;}"
    s="${s//</\&lt;}"
    s="${s//>/\&gt;}"
    s="${s//\"/\&quot;}"
    s="${s//\'/\&apos;}"
    # 换行转为数字引用，Excel 内可换行显示
    s="${s//$'\r'/}"
    s="${s//$'\n'/\&#10;}"
    printf '%s' "$s"
}

# 状态 -> 中文
xlsx_status_cn() {
    case "$1" in
        pass)   echo "合规" ;;
        fail)   echo "不合规" ;;
        manual) echo "需人工核查" ;;
        na)     echo "不适用" ;;
        *)      echo "$1" ;;
    esac
}

# 列序号 -> Excel 列字母（1->A ... 8->H）
xlsx_col() {
    printf '%c' "$((64 + $1))"
}

# 生成一个单元格（inlineStr 文本）
xlsx_cell() {
    local ref="$1" text="$2"
    printf '<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' "$ref" "$(xml_esc "$text")"
}

# 生成 sheet1.xml（标题/汇总/表头/数据）
build_sheet_xml() {
    local title="$1" i r
    local pass=0 fail=0 manual=0 na=0
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in
            pass) pass=$((pass+1)) ;; fail) fail=$((fail+1)) ;;
            manual) manual=$((manual+1)) ;; na) na=$((na+1)) ;;
        esac
    done

    printf '%s\n' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    printf '%s\n' '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    printf '%s\n' '<sheetData>'
    # 标题行
    printf '<row r="1"><c r="A1" t="inlineStr"><is><t>%s</t></is></c></row>\n' "$(xml_esc "$title")"
    # 汇总行
    local summary="合规：${pass}　不合规：${fail}　需人工核查：${manual}　不适用：${na}　共 ${R_COUNT} 项"
    printf '<row r="2"><c r="A2" t="inlineStr"><is><t>%s</t></is></c></row>\n' "$(xml_esc "$summary")"
    # 表头行（第4行）
    printf '<row r="4">'
    printf '%s' "$(xlsx_cell "A4" "章节")"
    printf '%s' "$(xlsx_cell "B4" "编号")"
    printf '%s' "$(xlsx_cell "C4" "类别")"
    printf '%s' "$(xlsx_cell "D4" "核查项")"
    printf '%s' "$(xlsx_cell "E4" "结果")"
    printf '%s' "$(xlsx_cell "F4" "详情")"
    printf '%s' "$(xlsx_cell "G4" "建议")"
    printf '%s' "$(xlsx_cell "H4" "参考指导书")"
    printf '%s\n' '</row>'
    # 数据行（从第5行开始）
    for ((i=1; i<=R_COUNT; i++)); do
        r=$((i + 4))
        printf '<row r="%d">' "$r"
        printf '%s' "$(xlsx_cell "A$r" "${R_CHAPTER[$i]}")"
        printf '%s' "$(xlsx_cell "B$r" "${R_ID[$i]}")"
        printf '%s' "$(xlsx_cell "C$r" "${R_CAT[$i]}")"
        printf '%s' "$(xlsx_cell "D$r" "${R_TITLE[$i]}")"
        printf '%s' "$(xlsx_cell "E$r" "$(xlsx_status_cn "${R_STATUS[$i]}")")"
        printf '%s' "$(xlsx_cell "F$r" "${R_DETAIL[$i]}")"
        printf '%s' "$(xlsx_cell "G$r" "${R_REC[$i]}")"
        printf '%s' "$(xlsx_cell "H$r" "${R_GUIDE[$i]}")"
        printf '%s\n' '</row>'
    done
    printf '%s\n' '</sheetData>'
    printf '%s\n' '</worksheet>'
}

# 生成真 .xlsx
# 用法：generate_xlsx <outfile> <title>
generate_xlsx() {
    local outfile="$1" title="${2:-配置核查报告}"
    local tmpdir i

    if ! command -v zip >/dev/null 2>&1; then
        echo "[提示] 未找到 zip 命令，跳过 .xlsx 生成（报告仍以 .html/.xls 输出）。可执行 yum install -y zip 后重试。"
        return 0
    fi

    tmpdir="$(mktemp -d)"
    mkdir -p "$tmpdir/xl/worksheets" "$tmpdir/_rels" "$tmpdir/xl/_rels"

    # [Content_Types].xml
    cat > "$tmpdir/[Content_Types].xml" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
EOF

    # _rels/.rels
    cat > "$tmpdir/_rels/.rels" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
EOF

    # xl/workbook.xml
    cat > "$tmpdir/xl/workbook.xml" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="核查明细" sheetId="1" r:id="rId1"/></sheets>
</workbook>
EOF

    # xl/_rels/workbook.xml.rels
    cat > "$tmpdir/xl/_rels/workbook.xml.rels" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
EOF

    # xl/worksheets/sheet1.xml
    build_sheet_xml "$title" > "$tmpdir/xl/worksheets/sheet1.xml"

    # 打包：[Content_Types].xml 必须作为第一个条目
    rm -f "$outfile"
    (cd "$tmpdir" && zip -q -X "$outfile" "[Content_Types].xml" && zip -q -X -r "$outfile" _rels xl)
    rm -rf "$tmpdir"

    echo "Excel(.xlsx)报告已生成：$outfile"
}

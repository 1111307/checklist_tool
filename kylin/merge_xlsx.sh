#!/bin/bash
# ============================================================
# 零依赖 xlsx 合并器（bash + unzip + zip，无需 Python）
# 把 kylin/output 下各组件 xlsx 合并成一个汇总 xlsx（单 sheet，加"组件"列）
# 用法：bash merge_xlsx.sh [outDir]（默认 ./output）
# 依赖：unzip、zip（麒麟几乎必带；缺失：yum install -y zip unzip）
# ============================================================
set -u

xml_esc() {
    local s="${1-}"
    s="${s//&/\&amp;}"; s="${s//</\&lt;}"; s="${s//>/\&gt;}"; s="${s//\"/\&quot;}"; s="${s//\'/\&apos;}"
    s="${s//$'\r'/}"; s="${s//$'\n'/\&#10;}"
    printf '%s' "$s"
}
xml_unescape() {
    local s="${1-}"
    s="${s//&#10;/$'\n'}"
    s="${s//&quot;/\"}"
    s="${s//&apos;/\'}"
    s="${s//&lt;/<}"
    s="${s//&gt;/>}"
    s="${s//&amp;/\&}"
    printf '%s' "$s"
}
col_letter() {
    local letters="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    printf '%s' "${letters:$(( $1 - 1 )):1}"
}

merge_xlsx() {
    local outdir="${1:-$PWD/output}"
    local files=() f
    for f in "$outdir"/配置核查报告_*.xlsx; do
        [ -f "$f" ] || continue
        case "$(basename "$f")" in *汇总*) continue ;; esac
        files+=("$f")
    done
    if [ ${#files[@]} -eq 0 ]; then echo "[跳过] $outdir 下无组件 xlsx 可合并"; return 0; fi
    if ! command -v unzip >/dev/null 2>&1 || ! command -v zip >/dev/null 2>&1; then
        echo "[提示] 缺少 unzip/zip 命令，跳过汇总（可 yum install -y zip unzip）"; return 0
    fi

    local tmp; tmp="$(mktemp -d)"
    local rowsfile="$tmp/rows"; : > "$rowsfile"
    local total=0 compcount=0

    for f in "${files[@]}"; do
        compcount=$((compcount+1))
        local base comp
        base="$(basename "$f" .xlsx)"; base="${base#配置核查报告_}"
        comp="${base%%_*}"
        [[ "$comp" =~ ^[0-9]+$ ]] && comp="操作系统"
        # 提取数据行（跳过标题/汇总/表头前 3 个 <row>）
        unzip -p "$f" xl/worksheets/sheet1.xml \
          | grep -oE '<row[^>]*>.*</row>' \
          | tail -n +4 \
          | while IFS= read -r rowline; do
                local cells=() t
                while IFS= read -r t; do
                    cells+=("$t")
                done < <(printf '%s\n' "$rowline" | grep -oE '<t[^>]*>[^<]*</t>' | sed 's#<t[^>]*>##; s#</t>##')
                printf '%s' "$comp"
                local i
                for i in 0 1 2 3 4 5 6 7; do
                    printf '\t%s' "$(xml_unescape "${cells[$i]:-}")"
                done
                printf '\n'
            done >> "$rowsfile"
    done

    total="$(wc -l < "$rowsfile")"
    if [ "$total" -eq 0 ]; then echo "[跳过] 未提取到数据行"; rm -rf "$tmp"; return 0; fi

    # 生成汇总 sheet1.xml（9 列 A-I）
    local sheet="$tmp/sheet.xml"
    {
        printf '%s\n' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        printf '%s\n' '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        printf '<row r="1"><c r="A1" t="inlineStr"><is><t>%s</t></is></c></row>\n' "$(xml_esc "配置核查汇总报告（共 $compcount 个组件，$total 项）")"
        printf '<row r="4">'
        local hd hdi
        hd=("组件" "章节" "编号" "类别" "核查项" "结果" "详情" "建议" "参考指导书")
        for hdi in 0 1 2 3 4 5 6 7 8; do
            printf '<c r="%s4" t="inlineStr"><is><t>%s</t></is></c>' "$(col_letter $((hdi+1)))" "${hd[$hdi]}"
        done
        printf '%s\n' '</row>'
        local rn=4
        while IFS=$'\t' read -r -a cols; do
            rn=$((rn+1))
            printf '<row r="%d">' "$rn"
            local ci
            for ci in 0 1 2 3 4 5 6 7 8; do
                printf '<c r="%s%d" t="inlineStr"><is><t>%s</t></is></c>' "$(col_letter $((ci+1)))" "$rn" "$(xml_esc "${cols[$ci]:-}")"
            done
            printf '%s\n' '</row>'
        done < "$rowsfile"
        printf '%s\n' '</sheetData></worksheet>'
    } > "$sheet"

    # 打包
    local ctdir="$tmp/pack"; mkdir -p "$ctdir/xl/worksheets" "$ctdir/_rels" "$ctdir/xl/_rels"
    cp "$sheet" "$ctdir/xl/worksheets/sheet1.xml"
    cat > "$ctdir/[Content_Types].xml" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
EOF
    cat > "$ctdir/_rels/.rels" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
EOF
    cat > "$ctdir/xl/workbook.xml" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="核查汇总" sheetId="1" r:id="rId1"/></sheets>
</workbook>
EOF
    cat > "$ctdir/xl/_rels/workbook.xml.rels" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
EOF

    local stamp outfile
    stamp="$(date +%Y%m%d_%H%M%S)"
    outfile="$(cd "$(dirname "$outdir")" && pwd)/配置核查汇总报告_${stamp}.xlsx"
    rm -f "$outfile"
    (cd "$ctdir" && zip -q -X "$outfile" "[Content_Types].xml" && zip -q -X -r "$outfile" _rels xl)
    rm -rf "$tmp"
    echo "汇总 Excel(.xlsx)报告已生成：$outfile"
}

merge_xlsx "$@"

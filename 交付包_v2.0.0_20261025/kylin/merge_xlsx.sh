#!/bin/bash
# ============================================================
# 零依赖 xlsx 合并器（bash + unzip + zip + awk，无需 Python）
# 把 kylin/output 下各组件 xlsx 合并成一个汇总 xlsx
# 用法：bash merge_xlsx.sh [合并目录]（默认 ./output）
# 依赖：unzip、zip（麒麟几乎必带；缺失：yum install -y zip unzip）
#
# 两种用法（评审 2026-09-30 第 6 条：不同设备的核查报告要能合并整合）
#   1) 单设备：output/ 下直接放各组件 xlsx      -> 设备列显示「本机」
#   2) 多设备：output/服务器A/、output/服务器B/  -> 设备列填子目录名
#
# 列结构自适应：按表头文字识别，兼容 7 / 8 / 10 列三种报告
# 输出 14 列：设备/组件/章节/编号/类别/核查项/结果/验证过程·方法/详情/修改建议/核查要求/参考指导书/数据目录/报告文件
# 单元格支持 inlineStr（本工具生成）与 sharedStrings（WPS/LibreOffice 重存）
# ============================================================
set -u

xml_esc() {
    local s="${1-}"
    s="${s//&/\&amp;}"; s="${s//</\&lt;}"; s="${s//>/\&gt;}"; s="${s//\"/\&quot;}"; s="${s//\'/\&apos;}"
    s="${s//$'\r'/}"; s="${s//$'\n'/\&#10;}"
    printf '%s' "$s"
}
col_letter() {
    local letters="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    printf '%s' "${letters:$(( $1 - 1 )):1}"
}

# 解析一个 sheet1.xml 的数据行。
# 入参（环境变量 SSFILE/COMP/DEV/HASNET）：组件名、设备名、是否已有网络设备报告、共享字符串文件
# 输出：每行一段完整的 <row> XML（12 列），由调用方直接拼进 sheet1.xml
AWK_PARSE='
function cidx(s,   i, n) {
    n = 0
    for (i = 1; i <= length(s); i++) n = n * 26 + index("ABCDEFGHIJKLMNOPQRSTUVWXYZ", substr(s, i, 1))
    return n - 1
}
function cltr(n) { return substr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", n, 1) }
# XML 转义：& 必须最先替换，后续实体里的 & 才不会被二次转义
function xesc(s) {
    gsub(/\r/, "", s)
    gsub(/\n/, " ", s)
    gsub(/\t/, " ", s)
    gsub(/&/, "\\&amp;",  s)
    gsub(/</, "\\&lt;",   s)
    gsub(/>/, "\\&gt;",   s)
    gsub(/"/, "\\&quot;", s)
    gsub(/\047/, "\\&apos;", s)
    return s
}
# 单元格文本是 XML 转义过的，先还原再交给 xml_esc 重新转义，否则 &amp; 会变成 &amp;amp;
function unx(s) {
    gsub(/&lt;/,   "<",   s)
    gsub(/&gt;/,   ">",   s)
    gsub(/&quot;/, "\"",  s)
    gsub(/&apos;/, "\047", s)
    gsub(/&#10;/,  " ",   s)
    gsub(/&amp;/,  "\\&", s)
    return s
}
BEGIN {
    # 用环境变量传参而不是 -v：awk 会解释 -v 值里的反斜杠转义，
    # Windows 风格路径（C:\Users\...）会被改得面目全非，导致 getline 读不到共享字符串。
    ssfile = ENVIRON["SSFILE"]
    comp   = ENVIRON["COMP"]
    dev    = ENVIRON["DEV"]
    hasnet = ENVIRON["HASNET"]
    dataDir = ENVIRON["DATADIR"]
    reportFile = ENVIRON["REPORTFILE"]
    ssn = 0
    if (ssfile != "") {
        while ((getline line < ssfile) > 0) { ssn++; SS[ssn] = line }
        close(ssfile)
    }
    hdrFound = 0
    outn = 0
}
{
    n = split($0, parts, "</c>")
    maxi = -1
    split("", cells)
    for (k = 1; k <= n; k++) {
        seg = parts[k]
        if (!match(seg, /r="[A-Z]+[0-9]+"/)) continue
        ref = substr(seg, RSTART + 3, RLENGTH - 4)
        col = ref
        sub(/[0-9]+$/, "", col)
        idx = cidx(col)
        val = ""
        if (match(seg, /<t[^>]*>/)) {
            rest = substr(seg, RSTART + RLENGTH)
            p = index(rest, "</t>")
            if (p > 0) val = substr(rest, 1, p - 1)
        } else if (match(seg, /<v>[^<]*<\/v>/)) {
            vv = substr(seg, RSTART + 3, RLENGTH - 7)
            if (seg ~ /t="s"/) {
                # sharedStrings 的 <v> 是 0-based 索引，awk 数组 1-based -> 取 SS[vv+1]
                if ((vv + 0) >= 0 && (vv + 1) <= ssn) val = SS[vv + 1]
            } else {
                # 数值单元格：Excel 把 1.1 存成 1.1000000000000001，按 15 位有效数字收敛
                if (vv ~ /^-?[0-9]+(\.[0-9]+)?$/ && length(vv) > 15) {
                    val = sprintf("%.15g", vv + 0)
                } else {
                    val = vv
                }
            }
        }
        cells[idx] = unx(val)
        if (idx > maxi) maxi = idx
    }
    if (maxi < 0) next

    if (!hdrFound) {
        isHdr = 0
        for (i = 0; i <= maxi; i++) if (cells[i] == "章节" || cells[i] == "编号") isHdr = 1
        if (isHdr) {
            for (i = 0; i <= maxi; i++) if (cells[i] != "") H[cells[i]] = i
            hdrFound = 1
        }
        next
    }

    iId = ("编号" in H) ? H["编号"] : -1
    iStatus = ("结果" in H) ? H["结果"] : (("状态" in H) ? H["状态"] : -1)
    if (iId < 0 || iStatus < 0) next
    # 「章节」与「参考指导书」是两个不同的源列。10 列报告里参考指导书是第 10 列，
    # 共用索引会把书名塞进「章节」，且第5章去重判断会永远失效。
    iCh  = ("章节" in H) ? H["章节"] : -1
    iRef = ("参考指导书" in H) ? H["参考指导书"] : (("章节" in H) ? H["章节"] : -1)
    iCat = ("类别" in H) ? H["类别"] : (("分类" in H) ? H["分类"] : -1)
    iTitle = ("核查项" in H) ? H["核查项"] : (("检查项" in H) ? H["检查项"] : -1)
    iMethod = ("验证过程/方法" in H) ? H["验证过程/方法"] : (("验证过程" in H) ? H["验证过程"] : (("方法" in H) ? H["方法"] : -1))
    iDetail = ("详情" in H) ? H["详情"] : -1
    iRec = ("修改建议" in H) ? H["修改建议"] : (("建议" in H) ? H["建议"] : (("修复建议" in H) ? H["修复建议"] : -1))
    iReq = ("核查要求" in H) ? H["核查要求"] : -1

    id = (iId >= 0) ? cells[iId] : ""
    title = (iTitle >= 0) ? cells[iTitle] : ""
    if (id == "" && title == "") next
    ch  = (iCh  >= 0) ? cells[iCh]  : ""
    ref = (iRef >= 0) ? cells[iRef] : ""
    # 第5章属"网络设备"对象：网络设备报告已在时，丢弃单机脚本的 na 占位行
    if (hasnet == "1" && comp == "操作系统" && ch ~ /^第5章/) next

    out[1] = dev
    out[2] = comp
    out[3] = ch
    out[4] = id
    out[5] = (iCat >= 0) ? cells[iCat] : ""
    out[6] = title
    out[7] = (iStatus >= 0) ? cells[iStatus] : ""
    out[8] = (iMethod >= 0) ? cells[iMethod] : ""
    out[9] = (iDetail >= 0) ? cells[iDetail] : ""
    out[10] = (iRec >= 0) ? cells[iRec] : ""
    out[11] = (iReq >= 0) ? cells[iReq] : ""
    out[12] = ref
    out[13] = (dataDir != "") ? dataDir : "-"
    out[14] = (reportFile != "") ? reportFile : "-"
    # 直接输出完整 <row> XML。旧版输出 tab 分隔文本再交给 bash 的 read -a，
    # 但 tab 属 IFS 空白字符，连续分隔符会被合并，空单元格会让后面所有列左移。
    rn = ++outn + 4
    rowxml = "<row r=\"" rn "\">"
    for (i = 1; i <= 14; i++) {
        rowxml = rowxml "<c r=\"" cltr(i) rn "\" t=\"inlineStr\"><is><t xml:space=\"preserve\">" \
                 xesc(out[i]) "</t></is></c>"
    }
    print rowxml "</row>"
}
'

merge_xlsx() {
    local outdir="${1:-$PWD/output}"
    comp_of() { # comp_of 文件名 → 组件名（纯时间戳=操作系统）
        local b="$1"; b="${b#配置核查报告_}"; b="${b%%_*}"
        case "$b" in ''|*[!0-9]*) printf '%s' "$b" ;; *) printf '%s' "操作系统" ;; esac
    }

    if [ ! -d "$outdir" ]; then echo "[跳过] 未找到目录：$outdir"; return 0; fi

    # ---- 收集（支持设备子目录；同设备同组件只留最新）----
    local files=() devs=() f d dev cname key
    declare -A seen=()
    collect() {
        local dir="$1" dv="$2" fn
        for fn in "$dir"/配置核查报告_*.xlsx; do
            [ -f "$fn" ] || continue
            case "$(basename "$fn")" in *汇总*) continue ;; esac
            local cn; cn="$(comp_of "$(basename "$fn")")"
            local k="$dv|$cn"
            if [ -n "${seen[$k]+x}" ]; then
                local prev="${seen[$k]}"
                [ "$fn" -nt "$prev" ] && seen[$k]="$fn"
            else
                seen[$k]="$fn"
                files+=("$fn"); devs+=("$dv")
            fi
        done
    }
    # 根目录报告一律按「本机」收录；设备子目录同时扫描。
    # 旧版「有子目录就只扫子目录」会静默丢掉根目录报告，与 VBS 版行为不一致。
    collect "$outdir" "本机"
    for d in "$outdir"/*/; do
        [ -d "$d" ] || continue
        dev="$(basename "$d")"
        collect "${d%/}" "$dev"
    done
    # seen 去重后可能覆盖了旧路径，重建一次（保证 files/devs 与 seen 一致）
    files=(); devs=()
    local k2
    while IFS= read -r k2; do
        [ -n "$k2" ] || continue
        files+=("${seen[$k2]}")
        devs+=("${k2%%|*}")
    done < <(printf '%s\n' "${!seen[@]}" | sort)
    if [ ${#files[@]} -eq 0 ]; then echo "[跳过] $outdir 下无组件 xlsx 可合并"; return 0; fi

    local HAS_NETDEV=0
    for f in "${files[@]}"; do
        case "$(basename "$f")" in *网络设备*) HAS_NETDEV=1 ;; esac
    done
    if ! command -v unzip >/dev/null 2>&1 || ! command -v zip >/dev/null 2>&1; then
        echo "[提示] 缺少 unzip/zip 命令，跳过汇总（可 yum install -y zip unzip）"; return 0
    fi

    local tmp; tmp="$(mktemp -d)"
    local rowsfile="$tmp/rows"; : > "$rowsfile"
    local compcount=${#files[@]} total=0

    local idx
    for idx in "${!files[@]}"; do
        f="${files[$idx]}"; dev="${devs[$idx]}"
        local base comp
        base="$(basename "$f" .xlsx)"; base="${base#配置核查报告_}"
        comp="${base%%_*}"
        [[ "$comp" =~ ^[0-9]+$ ]] && comp="操作系统"
        # 元信息：源报告第 3 行的「数据目录：…」（数据库类报告才有）
        local data_dir
        data_dir="$(unzip -p "$f" xl/worksheets/sheet1.xml 2>/dev/null | tr -d '\r' | tr -d '\n' \
                    | grep -o '数据目录：[^<]*' | head -1 | sed 's/^数据目录：//')"
        [ -z "$data_dir" ] && data_dir="-"
        local report_file; report_file="$(basename "$f")"
        # 共享字符串（WPS/LibreOffice 重存的 xlsx 用这个）
        local ssfile="$tmp/ss.txt"; : > "$ssfile"
        # 每个 <si> 必须输出且只输出一行（含空 <si>），否则后面所有索引都会错位；
        # 富文本 <si><r><t>A</t></r><r><t>B</t></r></si> 要把所有 <t> 拼起来。
        unzip -p "$f" xl/sharedStrings.xml 2>/dev/null \
          | tr -d '\r' | tr -d '\n' | sed 's#</si>#</si>\n#g' \
          | awk '/<si/ {
                    s = $0; txt = ""
                    while (match(s, /<t[^>]*>[^<]*<\/t>/)) {
                        seg = substr(s, RSTART, RLENGTH)
                        sub(/^<t[^>]*>/, "", seg); sub(/<\/t>$/, "", seg)
                        txt = txt seg
                        s = substr(s, RSTART + RLENGTH)
                    }
                    print txt
                 }' > "$ssfile"
        # 整段压成单行再按 </row> 切分：单元格文本含换行时逐行 grep 会漏配 <row>
        unzip -p "$f" xl/worksheets/sheet1.xml 2>/dev/null \
          | tr -d '\r' | tr -d '\n' \
          | sed 's#</row>#</row>\n#g' \
          | grep -E '<row' \
          | SSFILE="$ssfile" COMP="$comp" DEV="$dev" HASNET="$HAS_NETDEV" DATADIR="$data_dir" REPORTFILE="$report_file" awk "$AWK_PARSE" \
          >> "$rowsfile"
    done

    total="$(wc -l < "$rowsfile")"
    if [ "$total" -eq 0 ]; then echo "[跳过] 未提取到数据行"; rm -rf "$tmp"; return 0; fi

    # ---- 生成汇总 sheet1.xml（14 列 A-N）----
    local sheet="$tmp/sheet.xml"
    local devcount
    devcount="$(printf '%s\n' "${devs[@]}" | sort -u | wc -l)"
    {
        printf '%s\n' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        printf '%s\n' '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        printf '<row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c></row>\n' "$(xml_esc "配置核查汇总报告（$devcount 台设备 / $compcount 份组件报告 / $total 项）")"
        printf '<row r="4">'
        local hd hdi
        hd=("设备" "组件" "章节" "编号" "类别" "核查项" "结果" "验证过程/方法" "详情" "修改建议" "核查要求" "参考指导书" "数据目录" "报告文件")
        for hdi in $(seq 0 13); do
            printf '<c r="%s4" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' "$(col_letter $((hdi+1)))" "${hd[$hdi]}"
        done
        printf '%s\n' '</row>'
        # rowsfile 每行已是完整 <row> XML（awk 内完成转义），这里只做拼接
        local rowxml
        while IFS= read -r rowxml; do
            printf '%s\n' "$rowxml"
        done < "$rowsfile"
        printf '%s\n' '</sheetData></worksheet>'
    } > "$sheet"

    # ---- 打包 ----
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
    outfile="$outdir/配置核查汇总报告_${stamp}.xlsx"
    rm -f "$outfile"
    (cd "$ctdir" && zip -q -X "$outfile" "[Content_Types].xml" && zip -q -X -r "$outfile" _rels xl)
    rm -rf "$tmp"
    echo "汇总 Excel(.xlsx)报告已生成（$devcount 台设备 / $compcount 份组件报告 / $total 项）：$outfile"
}

merge_xlsx "$@"

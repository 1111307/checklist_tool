#!/bin/bash
# ============================================================
# 配置核查工具 - Redis 版（无需Python，纯Bash实现）
# 参考标准：配置核查作业指导书正式版2026_4_1
#           配置核查表_v2.0.0.xlsx（数据库列：Redis 共16项）
# 运行方式：sudo bash check_redis.sh
# 连接方式：默认 redis-cli 连接 127.0.0.1:6379；可用环境变量覆盖：
#             REDIS_HOST=127.0.0.1 REDIS_PORT=6379 REDIS_PASS=xxx bash check_redis.sh
# 输出文件：output/配置核查报告_Redis_日期时间.html /.xls
# ============================================================
set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT_DIR="$SCRIPT_DIR/output"
mkdir -p "$OUT_DIR"
STAMP="$(date +%Y%m%d_%H%M%S)"

# 零依赖 .xlsx 生成器（可选，需要 zip 命令；缺失时自动降级为 .xls）
[ -f "$SCRIPT_DIR/lib_xlsx.sh" ] && source "$SCRIPT_DIR/lib_xlsx.sh"

R_ID=(); R_CAT=(); R_TITLE=(); R_STATUS=(); R_DETAIL=(); R_CHAPTER=(); R_REC=(); R_GUIDE=()
R_COUNT=0

REDIS_CLI="$(command -v redis-cli 2>/dev/null)"
# 读取 db_config.conf（KEY=VALUE 格式）
read_config() {
    local key="$1" conf="$SCRIPT_DIR/db_config.conf" val=""
    if [ -f "$conf" ]; then
        val="$(grep -E "^[[:space:]]*${key}[[:space:]]*=" "$conf" 2>/dev/null | tail -1 | sed 's/^[^=]*=[[:space:]]*//; s/[[:space:]]*$//')"
    fi
    printf '%s' "$val"
}

REDIS_HOST="${REDIS_HOST:-$(read_config REDIS_HOST)}"
REDIS_HOST="${REDIS_HOST:-127.0.0.1}"
REDIS_PORT="${REDIS_PORT:-$(read_config REDIS_PORT)}"
REDIS_PORT="${REDIS_PORT:-6379}"
REDIS_PASS="${REDIS_PASS:-$(read_config REDIS_PASS)}"

DB_VERSION=""
CONN_OK=0
REDIS_PRESENT=0

redis_cmd() {
    local args=(-h "$REDIS_HOST" -p "$REDIS_PORT" --no-auth-warning)
    [ -n "$REDIS_PASS" ] && args+=(-a "$REDIS_PASS")
    "$REDIS_CLI" "${args[@]}" "$@" 2>/dev/null
}

# 取 CONFIG GET 的值（输出为 key\nvalue，取最后一行）
redis_cfg() { redis_cmd CONFIG GET "$1" 2>/dev/null | tail -1; }
# 取 INFO 段
redis_info() { redis_cmd INFO "$1" 2>/dev/null; }

# 判断 Redis 版本是否 EOL（<7.0 已停止维护）
redis_version_eol() {
    local v="$1" mm
    mm="$(echo "$v" | grep -oE '^[0-9]+' | head -1)"
    if [ -n "$mm" ] && [ "$mm" -lt 7 ] 2>/dev/null; then echo "eol"; else echo "supported"; fi
}

guide_ref() {
    case "$1" in
        1.1)  echo "《配置核查作业指导书》第1章 系统安全 1.1：操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序" ;;
        1.7)  echo "《配置核查作业指导书》第1章 系统安全 1.7：数据库管理系统应删除冗余帐户，应设置不少于8个字符且字母大小写、数字及特殊字符混合编制的账户口令" ;;
        1.8)  echo "《配置核查作业指导书》第1章 系统安全 1.8：数据库管理系统应删除冗余存储过程" ;;
        1.10) echo "《配置核查作业指导书》第1章 系统安全 1.10：数据库管理系统应具有自主访问控制功能" ;;
        1.11) echo "《配置核查作业指导书》第1章 系统安全 1.11：数据库管理系统应具有备份和恢复功能" ;;
        1.12) echo "《配置核查作业指导书》第1章 系统安全 1.12：数据库管理系统应具有表级审计、告警和阻断功能" ;;
        1.15) echo "《配置核查作业指导书》第1章 系统安全 1.15：应具备数据库管理系统超级管理员远程登录限制能力" ;;
        1.16) echo "《配置核查作业指导书》第1章 系统安全 1.16：应具备数据库管理系统输入（参数）检查能力" ;;
        1.19) echo "《配置核查作业指导书》第1章 系统安全 1.19：应更换数据库管理系统的默认服务端口、管理员用户名和口令" ;;
        1.20) echo "《配置核查作业指导书》第1章 系统安全 1.20：数据库管理系统应配置安全策略" ;;
        1.21) echo "《配置核查作业指导书》第1章 系统安全 1.21：数据库管理系统应具有行级或列级审计功能" ;;
        1.22) echo "《配置核查作业指导书》第1章 系统安全 1.22：数据库管理系统应采取单独、安全监控、审计措施" ;;
        1.23) echo "《配置核查作业指导书》第1章 系统安全 1.23：数据库管理系统仅为应用服务器提供访问服务" ;;
        1.25) echo "《配置核查作业指导书》第1章 系统安全 1.25：检查是否具备边界保护能力，是否可以抗攻击、防篡改" ;;
        1.26) echo "《配置核查作业指导书》第1章 系统安全 1.26：检查是否有防病毒日志、补丁日志、记录相关信息，记录信息的完整、有效" ;;
        2.16) echo "《配置核查作业指导书》第2章 用户安全 2.16：检查被试装备中操作系统、数据库以及应用软件等是否完成补丁修复和升级到最新版本" ;;
        *)    echo "《配置核查作业指导书》" ;;
    esac
}

add_result() {
    R_COUNT=$((R_COUNT+1))
    R_ID[$R_COUNT]="$1"; R_CAT[$R_COUNT]="$2"; R_TITLE[$R_COUNT]="$3"; R_STATUS[$R_COUNT]="$4"
    R_DETAIL[$R_COUNT]="$5"; R_CHAPTER[$R_COUNT]="$6"; R_REC[$R_COUNT]="$7"
    R_GUIDE[$R_COUNT]="$(guide_ref "$1")"
}

db_unreachable() {
    if [ "$REDIS_PRESENT" = "0" ]; then
        add_result "$1" "系统安全-数据库" "$2" "na" "未检测到 redis-cli 客户端，无法对 Redis 进行自动核查，标记不适用。" "第1章" "$3"
    else
        add_result "$1" "系统安全-数据库" "$2" "manual" "无法连接 Redis（${REDIS_HOST}:${REDIS_PORT}），需人工登录核查。请检查 REDIS_PASS 环境变量或连接参数。" "第1章" "$3"
    fi
}

html_esc() { local s="$1"; s="${s//&/&amp;}"; s="${s//</&lt;}"; s="${s//>/&gt;}"; printf '%s' "$s"; }
status_cn() { case "$1" in pass) echo "合规";; fail) echo "不合规";; manual) echo "需人工核查";; na) echo "不适用";; *) echo "$1";; esac; }
status_color() { case "$1" in pass) echo "#2e7d32";; fail) echo "#c62828";; manual) echo "#ef6c00";; na) echo "#757575";; *) echo "#000000";; esac; }

build_table_rows_by_status() {
    local want="$1" i
    for ((i=1; i<=R_COUNT; i++)); do
        [ "${R_STATUS[$i]}" != "$want" ] && continue
        local color scn; color="$(status_color "${R_STATUS[$i]}")"; scn="$(status_cn "${R_STATUS[$i]}")"
        cat <<ROW
<tr>
<td>$(html_esc "${R_CHAPTER[$i]}")</td>
<td>$(html_esc "${R_ID[$i]}")</td>
<td>$(html_esc "${R_CAT[$i]}")</td>
<td>$(html_esc "${R_TITLE[$i]}")</td>
<td style="color:$color;font-weight:bold;">$scn</td>
<td>$(html_esc "${R_DETAIL[$i]}")</td>
<td>$(html_esc "${R_REC[$i]}")</td>
<td>$(html_esc "${R_GUIDE[$i]}")</td>
</tr>
ROW
    done
}

print_summary() {
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in pass) pass=$((pass+1));; fail) fail=$((fail+1));; manual) manual=$((manual+1));; na) na=$((na+1));; esac
    done
    echo "核查完成，共 $R_COUNT 项："
    echo "  合规(pass)：$pass    不合规(fail)：$fail    需人工核查(manual)：$manual    不适用(na)：$na"
}

json_esc() {
    local s="$1"
    s="${s//\\/\\\\}"
    s="${s//\"/\\\"}"
    s="${s//$'\n'/ }"
    s="${s//$'\r'/}"
    s="${s//$'\t'/ }"
    printf '%s' "$s"
}

report_rows_json() {
    local i first=1
    printf '['
    for ((i=1; i<=R_COUNT; i++)); do
        [ "$first" -eq 1 ] || printf ','
        first=0
        printf '{"ch":"%s","id":"%s","cat":"%s","title":"%s","status":"%s","detail":"%s","rec":"%s","guide":"%s"}' \
            "$(json_esc "${R_CHAPTER[$i]}")" "$(json_esc "${R_ID[$i]}")" "$(json_esc "${R_CAT[$i]}")" \
            "$(json_esc "${R_TITLE[$i]}")" "$(json_esc "${R_STATUS[$i]}")" "$(json_esc "${R_DETAIL[$i]}")" \
            "$(json_esc "${R_REC[$i]}")" "$(json_esc "${R_GUIDE[$i]}")"
    done
    printf ']'
}

REPORT_TAG="Redis"
REPORT_TITLE="Redis 配置核查报告"
report_meta_html() {
    echo "<div>连接：$(html_esc "$REDIS_HOST:$REDIS_PORT")　版本：$(html_esc "${DB_VERSION:-未连接}")</div>"
    echo "<div>主机名：$(hostname 2>/dev/null)　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</div>"
    echo "<div>参考标准：配置核查作业指导书正式版2026_4_1 / 配置核查表_v2.0.0.xlsx</div>"
}

generate_html() {
    local outfile="$OUT_DIR/配置核查报告${REPORT_TAG:+_$REPORT_TAG}_${STAMP}.html"
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in
            pass) pass=$((pass+1));;
            fail) fail=$((fail+1));;
            manual) manual=$((manual+1));;
            na) na=$((na+1));;
        esac
    done
    local total=$R_COUNT
    local ppass pfail pmanual pna
    ppass="$(awk -v a="$pass" -v t="$total" 'BEGIN{if(t==0) printf "0.0"; else printf "%.1f", a*100/t}')"
    pfail="$(awk -v a="$fail" -v t="$total" 'BEGIN{if(t==0) printf "0.0"; else printf "%.1f", a*100/t}')"
    pmanual="$(awk -v a="$manual" -v t="$total" 'BEGIN{if(t==0) printf "0.0"; else printf "%.1f", a*100/t}')"
    pna="$(awk -v a="$na" -v t="$total" 'BEGIN{if(t==0) printf "0.0"; else printf "%.1f", a*100/t}')"
    {
        cat <<HTMLHEAD
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${REPORT_TITLE}</title>
<style>
:root{
  --bg:#f5f7fa; --card:#fff; --ink:#1f2937; --muted:#6b7280; --line:#e5e7eb;
  --brand:#0f3057; --brand2:#1a3c6e;
  --pass:#15803d; --pass-bg:#ecfdf5; --pass-br:#bbf7d0;
  --fail:#b91c1c; --fail-bg:#fef2f2; --fail-br:#fecaca;
  --manual:#b45309; --manual-bg:#fffbeb; --manual-br:#fde68a;
  --na:#4b5563; --na-bg:#f3f4f6; --na-br:#e5e7eb;
}
*{box-sizing:border-box;}
body{margin:0;font:14px/1.65 "Segoe UI","Microsoft YaHei",system-ui,sans-serif;color:var(--ink);background:var(--bg);}
.wrap{max-width:1240px;margin:0 auto;padding:24px 20px 60px;}
header{background:linear-gradient(135deg,var(--brand) 0%,var(--brand2) 60%,#2563eb 130%);color:#fff;border-radius:12px;padding:26px 30px;margin-bottom:20px;}
header h1{margin:0 0 6px;font-size:22px;letter-spacing:.5px;}
header .sub{opacity:.85;font-size:13px;}
.meta{display:flex;flex-wrap:wrap;gap:8px 28px;margin-top:16px;padding-top:14px;border-top:1px solid rgba(255,255,255,.25);font-size:13px;}
.meta div{opacity:.95;}
.meta b{font-weight:600;opacity:.75;margin-right:6px;}
.dash{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px;}
.stat{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px 16px 14px;position:relative;overflow:hidden;cursor:pointer;transition:transform .15s, box-shadow .15s;}
.stat:hover{transform:translateY(-2px);box-shadow:0 6px 18px rgba(15,48,87,.10);}
.stat .num{font-size:34px;font-weight:700;line-height:1.1;font-variant-numeric:tabular-nums;}
.stat .lbl{color:var(--muted);font-size:13px;margin-top:2px;}
.stat .bar{height:4px;border-radius:2px;margin-top:12px;background:var(--line);}
.stat .bar i{display:block;height:100%;border-radius:2px;}
.stat.s-pass .num{color:var(--pass);} .stat.s-pass .bar i{background:var(--pass);}
.stat.s-fail .num{color:var(--fail);} .stat.s-fail .bar i{background:var(--fail);}
.stat.s-manual .num{color:var(--manual);} .stat.s-manual .bar i{background:var(--manual);}
.stat.s-na .num{color:var(--na);} .stat.s-na .bar i{background:var(--na);}
.stat .pct{position:absolute;right:14px;top:16px;font-size:12px;color:var(--muted);}
.toolbar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:16px;}
.toolbar input[type=text]{flex:1;min-width:200px;padding:9px 14px;border:1px solid var(--line);border-radius:8px;font-size:13px;outline:none;background:var(--card);}
.toolbar input[type=text]:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.12);}
.filters{display:flex;gap:6px;flex-wrap:wrap;}
.fbtn{border:1px solid var(--line);background:var(--card);color:var(--ink);padding:7px 14px;border-radius:20px;font-size:13px;cursor:pointer;transition:all .15s;}
.fbtn:hover{border-color:#2563eb;color:#2563eb;}
.fbtn.on{background:var(--brand);border-color:var(--brand);color:#fff;}
.count{color:var(--muted);font-size:12px;margin-left:4px;}
.panel{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden;}
table{border-collapse:collapse;width:100%;font-size:13px;}
thead th{background:#f8fafc;color:#334155;text-align:left;font-weight:600;padding:10px 12px;border-bottom:2px solid var(--line);white-space:nowrap;position:sticky;top:0;z-index:5;}
tbody td{padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top;}
tbody tr:hover{background:#f8fafc;}
td.id{font-family:Consolas,monospace;font-weight:600;white-space:nowrap;}
td.cat{white-space:nowrap;color:var(--muted);}
td.title{min-width:180px;}
.badge{display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;white-space:nowrap;border:1px solid;}
.badge-pass{color:var(--pass);background:var(--pass-bg);border-color:var(--pass-br);}
.badge-fail{color:var(--fail);background:var(--fail-bg);border-color:var(--fail-br);}
.badge-manual{color:var(--manual);background:var(--manual-bg);border-color:var(--manual-br);}
.badge-na{color:var(--na);background:var(--na-bg);border-color:var(--na-br);}
td.detail,td.rec{color:#374151;max-width:320px;}
.guide{color:var(--muted);font-size:12px;max-width:260px;}
.empty{padding:60px;text-align:center;color:var(--muted);}
footer{margin-top:26px;color:var(--muted);font-size:12px;text-align:center;}
@media(max-width:900px){.dash{grid-template-columns:repeat(2,1fr);}}
@media print{.toolbar{display:none;} .panel{border:none;} body{background:#fff;}}
</style>
</head>
<body>
<div class="wrap">
<header>
  <h1>${REPORT_TITLE}</h1>
  <div class="sub">参考标准：配置核查作业指导书正式版2026_4_1</div>
  <div class="meta">
HTMLHEAD
        report_meta_html
        cat <<HTMLMID
  </div>
</header>

<div class="dash">
  <div class="stat s-pass" onclick="fset('pass')"><div class="num">$pass</div><div class="lbl">合规</div><div class="bar"><i style="width:$ppass%"></i></div><div class="pct">$ppass%</div></div>
  <div class="stat s-fail" onclick="fset('fail')"><div class="num">$fail</div><div class="lbl">不合规</div><div class="bar"><i style="width:$pfail%"></i></div><div class="pct">$pfail%</div></div>
  <div class="stat s-manual" onclick="fset('manual')"><div class="num">$manual</div><div class="lbl">需人工核查</div><div class="bar"><i style="width:$pmanual%"></i></div><div class="pct">$pmanual%</div></div>
  <div class="stat s-na" onclick="fset('na')"><div class="num">$na</div><div class="lbl">不适用</div><div class="bar"><i style="width:$pna%"></i></div><div class="pct">$pna%</div></div>
</div>

<div class="toolbar">
  <input id="q" type="text" placeholder="搜索编号 / 核查项 / 详情…">
  <div class="filters">
    <button class="fbtn on" data-f="all">全部<span class="count">$total</span></button>
    <button class="fbtn" data-f="fail">不合规<span class="count">$fail</span></button>
    <button class="fbtn" data-f="manual">需人工<span class="count">$manual</span></button>
    <button class="fbtn" data-f="pass">合规<span class="count">$pass</span></button>
    <button class="fbtn" data-f="na">不适用<span class="count">$na</span></button>
  </div>
</div>

<div class="panel">
<table id="tbl">
<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr></thead>
<tbody id="tb"></tbody>
</table>
<div class="empty" id="empty" style="display:none">没有匹配的核查项</div>
</div>

<footer>本报告由配置核查工具自动生成 · $(date '+%Y-%m-%d %H:%M:%S')</footer>
</div>

<script>
var DATA = 
HTMLMID
        report_rows_json
        cat <<HTMLFOOT
;
var ST = {pass:"合规", fail:"不合规", manual:"需人工核查", na:"不适用"};
var curF = "all";
function esc(s){var d=document.createElement("div");d.textContent=s==null?"":s;return d.innerHTML;}
function render(){
  var q = document.getElementById("q").value.trim().toLowerCase();
  var tb = document.getElementById("tb"); tb.innerHTML = "";
  var n = 0;
  DATA.forEach(function(x){
    if(curF!="all" && x.status!=curF) return;
    if(q && (x.id+" "+x.title+" "+x.detail+" "+x.cat+" "+x.ch).toLowerCase().indexOf(q)<0) return;
    n++;
    var tr = document.createElement("tr");
    tr.innerHTML = "<td>"+esc(x.ch)+"</td><td class='id'>"+esc(x.id)+"</td><td class='cat'>"+esc(x.cat)+"</td>"+
      "<td class='title'>"+esc(x.title)+"</td>"+
      "<td><span class='badge badge-"+x.status+"'>"+ST[x.status]+"</span></td>"+
      "<td class='detail'>"+esc(x.detail)+"</td>"+
      "<td class='rec'>"+esc(x.rec)+"</td>"+
      "<td class='guide'>"+esc(x.guide)+"</td>";
    tb.appendChild(tr);
  });
  document.getElementById("empty").style.display = n? "none":"block";
}
function fset(f){
  curF = f;
  var bs = document.querySelectorAll(".fbtn");
  for(var i=0;i<bs.length;i++){ bs[i].className = "fbtn" + (bs[i].getAttribute("data-f")==f ? " on" : ""); }
  render();
}
(function(){
  var bs = document.querySelectorAll(".fbtn");
  for(var i=0;i<bs.length;i++){ bs[i].onclick = (function(b){ return function(){ fset(b.getAttribute("data-f")); }; })(bs[i]); }
})();
document.getElementById("q").addEventListener("input", render);
render();
</script>
</body>
</html>
HTMLFOOT
    } > "$outfile"
    echo "HTML报告已生成：$outfile"
}

generate_xls() {
    local outfile="$OUT_DIR/配置核查报告_Redis_${STAMP}.xls"
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in pass) pass=$((pass+1));; fail) fail=$((fail+1));; manual) manual=$((manual+1));; na) na=$((na+1));; esac
    done
    {
        cat <<XLSHEAD
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="UTF-8">
<!--[if gte mso 9]>
<xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
<x:Name>Redis配置核查报告</x:Name>
<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
<![endif]-->
<style>
table{border-collapse:collapse;}
th,td{border:1px solid #999;padding:4px 6px;font-family:"Microsoft YaHei",Arial;font-size:12px;mso-number-format:"\@";}
th{background:#1a3c6e;color:#fff;font-weight:bold;}
</style>
</head>
<body>
<p>Redis 配置核查报告　连接：$(html_esc "$REDIS_HOST:$REDIS_PORT")　版本：$(html_esc "${DB_VERSION:-未连接}")　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</p>
<p>合规：$pass　不合规：$fail　需人工核查：$manual　不适用：$na</p>
XLSHEAD
        if [ "$fail" -gt 0 ]; then echo "<p><b>一、未通过（$fail 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status fail; echo "</table>"; fi
        if [ "$manual" -gt 0 ]; then echo "<p><b>二、需人工核查（$manual 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status manual; echo "</table>"; fi
        if [ "$na" -gt 0 ]; then echo "<p><b>三、不适用（$na 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status na; echo "</table>"; fi
        if [ "$pass" -gt 0 ]; then echo "<p><b>四、通过（$pass 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status pass; echo "</table>"; fi
        cat <<XLSFOOT
</body>
</html>
XLSFOOT
    } > "$outfile"
    echo "Excel(.xls)报告已生成：$outfile"
}

# ============================================================
# 核查项（共16项，对应配置核查表 Redis 列）
# ============================================================

check_1_1_patch() {
    local eol; eol="$(redis_version_eol "$DB_VERSION")"
    if [ "$eol" = "eol" ]; then
        add_result "1.1" "系统安全-数据库" "数据库补丁程序" "fail" "当前 Redis 版本 ${DB_VERSION} 已停止维护（<7.0），无法再获得安全补丁。" "第1章" "升级到 Redis 7.x/8.x 并安装最新补丁。"
    else
        add_result "1.1" "系统安全-数据库" "数据库补丁程序" "manual" "当前 Redis 版本 ${DB_VERSION}（仍在维护期）。请人工核对是否为最新补丁版本。" "第1章" "及时升级到最新稳定版本。"
    fi
}

check_1_7_dbauth() {
    local pass; pass="$(redis_cfg requirepass)"
    if [ -z "$pass" ]; then
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "fail" "未设置 requirepass，Redis 无需口令即可访问（存在未授权访问风险）。" "第1章" "设置 requirepass 强口令（不少于8位、含大小写/数字/特殊字符）。"
    else
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "pass" "已设置 requirepass 认证口令。" "第1章" "持续使用强口令并定期更换。"
    fi
}

check_1_8_dangerous_cmds() {
    # 用 COMMAND INFO 检测：命令被 rename 禁用后返回空，未禁用则返回命令信息
    local risky="" c
    for c in FLUSHALL FLUSHDB EVAL SHUTDOWN; do
        [ -n "$(redis_cmd COMMAND INFO "$c" 2>/dev/null)" ] && risky="$risky $c"
    done
    if [ -n "$risky" ]; then
        add_result "1.8" "系统安全-数据库" "数据库存储过程管理" "fail" "以下高危命令未禁用/重命名：$risky" "第1章" "在配置中 rename-command 禁用 FLUSHALL/FLUSHDB/EVAL/SHUTDOWN 等危险命令。"
    else
        add_result "1.8" "系统安全-数据库" "数据库存储过程管理" "pass" "高危命令（FLUSHALL/FLUSHDB/EVAL/SHUTDOWN）已重命名或禁用。" "第1章" "持续禁用不必要的危险命令。"
    fi
}

check_1_10_access() {
    local pm bind
    pm="$(redis_cfg protected-mode)"
    bind="$(redis_cfg bind)"
    if [ "$pm" = "no" ] && { [ -z "$bind" ] || [ "$bind" = "0.0.0.0" ]; }; then
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "fail" "protected-mode=no 且 bind=${bind:-未设置}，Redis 对全网开放。" "第1章" "开启 protected-mode，并绑定 127.0.0.1 或内网地址。"
    else
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "pass" "protected-mode=${pm:-默认}，bind=${bind:-默认}，访问受限。" "第1章" "持续限制监听地址与访问来源。"
    fi
}

check_1_11_backup() {
    local save aof
    save="$(redis_cfg save)"
    aof="$(redis_cfg appendonly)"
    if [ -n "$save" ] && [ "$save" != "" ] && [ "$save" != '""' ]; then
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "pass" "已配置 RDB 快照（save=$save）。" "第1章" "结合 RDB/AOF 定期备份并异地存储。"
    elif [ "$aof" = "yes" ]; then
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "pass" "appendonly=yes（AOF 持久化已开启）。" "第1章" "结合 RDB/AOF 定期备份并异地存储。"
    else
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "fail" "未配置 RDB 快照（save 为空）且 appendonly=${aof:-no}，数据无持久化，宕机即丢数据。" "第1章" "开启 RDB/AOF 持久化并建立定期备份。"
    fi
}

check_1_12_audit() {
    local slowlog
    slowlog="$(redis_cfg slowlog-log-slower-than)"
    if [ -n "$slowlog" ] && [ "$slowlog" != "" ] && [ "$slowlog" != "-1" ]; then
        add_result "1.12" "系统安全-数据库" "数据库审计插件" "pass" "slowlog 已启用（slowlog-log-slower-than=$slowlog）。" "第1章" "结合慢日志与外部审计工具记录关键操作。"
    else
        add_result "1.12" "系统安全-数据库" "数据库审计插件" "fail" "slowlog 未启用（slowlog-log-slower-than=${slowlog:-未设置}），无操作审计能力。" "第1章" "开启 slowlog 并配合审计工具。"
    fi
}

check_1_15_remote() {
    local pm bind pass
    pm="$(redis_cfg protected-mode)"
    bind="$(redis_cfg bind)"
    pass="$(redis_cfg requirepass)"
    if [ -z "$pass" ]; then
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "fail" "未设置 requirepass，任意来源可直接远程登录执行管理命令。" "第1章" "设置 requirepass 强口令并限制远程访问来源。"
    elif [ "$bind" = "127.0.0.1" ] || [ "$pm" = "yes" ]; then
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "pass" "已设置认证口令，且访问受限（bind=${bind:-默认}，protected-mode=${pm:-默认}）。" "第1章" "持续限制 Redis 仅对可信终端开放。"
    else
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "manual" "已设置口令，但 bind=${bind:-默认} 且 protected-mode=${pm:-默认}，请人工确认是否限制了远程来源。" "第1章" "通过防火墙/绑定地址限制远程访问来源。"
    fi
}

check_1_16_inputcheck() {
    add_result "1.16" "系统安全-数据库" "输入验证/SQL注入防护" "manual" "Redis 输入（参数）检查属应用层能力（键值校验、命令注入防护），需人工核查应用系统是否对输入做了校验。" "第1章" "应用层对 Redis 键值输入做合法性校验，避免命令注入。"
}

check_1_19_defaults() {
    local port pass
    port="$(redis_cfg port)"
    pass="$(redis_cfg requirepass)"
    local issues=""
    [ "$port" = "6379" ] && issues="仍使用默认端口 6379"
    [ -z "$pass" ] && issues="${issues:+$issues；}未设置口令"
    if [ "$port" = "6379" ] || [ -z "$pass" ]; then
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "fail" "$issues" "第1章" "更换默认端口，设置强口令。"
    else
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "pass" "未发现默认端口/无口令问题（端口=${port:-未知}，已设口令）。" "第1章" "持续保持非默认端口与强口令。"
    fi
}

check_1_20_policy() {
    local maxmem renamed pm
    maxmem="$(redis_cfg maxmemory)"
    renamed="$(redis_cmd CONFIG GET rename-command 2>/dev/null | grep -c '""' )"
    pm="$(redis_cfg protected-mode)"
    if { [ -n "$maxmem" ] && [ "$maxmem" != "0" ]; } || [ "$renamed" -gt 0 ] 2>/dev/null; then
        add_result "1.20" "系统安全-数据库" "数据库安全策略" "pass" "检测到安全策略：maxmemory=${maxmem:-0}，危险命令重命名 ${renamed:-0} 项，protected-mode=${pm:-默认}。" "第1章" "完善安全基线：内存上限、危险命令禁用、保护模式。"
    else
        add_result "1.20" "系统安全-数据库" "数据库安全策略" "fail" "未检测到明显安全策略（maxmemory 未限制、危险命令未禁用）。" "第1章" "配置 maxmemory、rename-command 等安全策略。"
    fi
}

check_1_21_audit() {
    add_result "1.21" "系统安全-数据库" "数据库操作审计" "manual" "Redis 无行级/列级审计概念，操作审计需结合 slowlog 与外部审计工具，请人工确认审计粒度。" "第1章" "结合外部审计/监控平台记录敏感操作。"
}

check_1_22_monitor() {
    add_result "1.22" "系统安全-数据库" "数据库审计日志留存" "manual" "独立安全监控与审计措施（第三方监控平台）需人工核查部署情况。" "第1章" "部署独立监控/审计平台，日志与业务分离存储。"
}

check_1_23_appsep() {
    local pm bind
    pm="$(redis_cfg protected-mode)"
    bind="$(redis_cfg bind)"
    if [ "$bind" = "127.0.0.1" ] || [ "$pm" = "yes" ]; then
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "pass" "访问受限（bind=${bind:-默认}，protected-mode=${pm:-默认}），仅为本机/可信来源提供访问。" "第1章" "持续确保 Redis 仅为应用服务器提供访问。"
    else
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "manual" "bind=${bind:-默认} 且 protected-mode=${pm:-默认}，请人工确认是否仅对应用服务器开放。" "第1章" "绑定内网地址，限制 Redis 仅对应用服务器开放。"
    fi
}

check_1_25_boundary() {
    local fw fwstate
    fw="$(iptables -S 2>/dev/null | head -20)"
    fwstate="$(firewall-cmd --state 2>/dev/null)"
    if [ -n "$fw" ] || [ "$fwstate" = "running" ]; then
        add_result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "manual" "检测到本机防火墙（${fwstate:-iptables 规则存在}），但边界防护/WAF/网络隔离需结合网络架构人工核查。" "第1章" "部署边界防护，限制 Redis 暴露面。"
    else
        add_result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "fail" "未检测到本机防火墙规则，Redis 缺乏边界保护。" "第1章" "部署防火墙/WAF 限制 Redis 暴露面。"
    fi
}

check_1_26_log() {
    local logfile
    logfile="$(redis_cfg logfile)"
    if [ -n "$logfile" ] && [ "$logfile" != '""' ] && [ "$logfile" != "" ]; then
        add_result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "pass" "已配置日志文件：logfile=$logfile。" "第1章" "持续确保日志完整记录、留存满足要求。"
    else
        add_result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "fail" "logfile 未配置到独立文件（${logfile:-空}），日志未落地。" "第1章" "配置 logfile 到独立文件，保留完整日志。"
    fi
}

check_2_16_patchlatest() {
    local eol; eol="$(redis_version_eol "$DB_VERSION")"
    if [ "$eol" = "eol" ]; then
        add_result "2.16" "系统安全-数据库" "补丁修复升级到最新" "fail" "当前 Redis 版本 ${DB_VERSION} 已停止维护（<7.0），未完成补丁修复升级。" "第2章" "升级到 Redis 7.x/8.x 并完成补丁修复。"
    else
        add_result "2.16" "系统安全-数据库" "补丁修复升级到最新" "manual" "当前 Redis 版本 ${DB_VERSION}（仍在维护期）。请人工核对是否已升级到最新版本。" "第2章" "定期升级到最新版本。"
    fi
}

# ============================================================
# 主流程
# ============================================================

echo "================================================================"
echo "  配置核查工具 - Redis 版（无需Python）"
echo "  参考标准：配置核查作业指导书正式版2026_4_1"
echo "================================================================"
echo "  连接：$REDIS_HOST:$REDIS_PORT"
echo ""

if [ -z "$REDIS_CLI" ]; then
    REDIS_PRESENT=0
    DB_VERSION="未安装redis-cli"
    echo "  未检测到 redis-cli 客户端，所有项将标记为不适用。"
elif DB_VERSION="$(redis_info server 2>/dev/null | grep '^redis_version:' | awk -F: '{print $2}' | tr -d '\r ')"; [ -n "$DB_VERSION" ]; then
    REDIS_PRESENT=1; CONN_OK=1
    echo "  Redis 版本：$DB_VERSION"
else
    REDIS_PRESENT=1; CONN_OK=0
    echo "  [!] 无法连接 Redis（$REDIS_HOST:$REDIS_PORT），所有项将标记为需人工核查。"
    echo "      请通过 REDIS_PASS 环境变量提供认证口令。"
fi
echo ""
echo "开始检查，请稍候..."

if [ "$CONN_OK" = "1" ]; then
    check_1_1_patch
    check_1_7_dbauth
    check_1_8_dangerous_cmds
    check_1_10_access
    check_1_11_backup
    check_1_12_audit
    check_1_15_remote
    check_1_16_inputcheck
    check_1_19_defaults
    check_1_20_policy
    check_1_21_audit
    check_1_22_monitor
    check_1_23_appsep
    check_1_25_boundary
    check_1_26_log
    check_2_16_patchlatest
else
    db_unreachable "1.1"  "数据库补丁程序"          "升级到最新稳定版本。"
    db_unreachable "1.7"  "数据库账户管理"          "设置 requirepass 强口令。"
    db_unreachable "1.8"  "数据库存储过程管理"      "禁用危险命令。"
    db_unreachable "1.10" "数据库访问控制"          "开启 protected-mode 并绑定地址。"
    db_unreachable "1.11" "数据库备份策略"          "开启 RDB/AOF 持久化。"
    db_unreachable "1.12" "数据库审计插件"          "开启 slowlog。"
    db_unreachable "1.15" "数据库远程访问控制"      "设置口令并限制来源。"
    db_unreachable "1.16" "输入验证/SQL注入防护"    "应用层做输入校验。"
    db_unreachable "1.19" "数据库默认账户/端口"     "更换默认端口、设置口令。"
    db_unreachable "1.20" "数据库安全策略"          "配置 maxmemory、禁用危险命令。"
    db_unreachable "1.21" "数据库操作审计"          "结合外部审计工具。"
    db_unreachable "1.22" "数据库审计日志留存"      "独立监控平台。"
    db_unreachable "1.23" "应用与数据库账户分离"    "限制仅应用服务器访问。"
    db_unreachable "1.25" "边界保护抗攻击防篡改"    "部署防火墙。"
    db_unreachable "1.26" "防病毒/补丁日志记录"     "配置 logfile。"
    db_unreachable "2.16" "补丁修复升级到最新"      "升级到最新版本。"
fi

echo ""
echo "================================================================"
print_summary
generate_html
generate_xls
type generate_xlsx >/dev/null 2>&1 && generate_xlsx "$OUT_DIR/配置核查报告_Redis_${STAMP}.xlsx" "Redis 配置核查报告"
echo ""
echo "核查完成，请到 output/ 目录查看 HTML/XLS/XLSX 报告。"

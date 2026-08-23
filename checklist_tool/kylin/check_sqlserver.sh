#!/bin/bash
# ============================================================
# 配置核查工具 - SQL Server 数据库版（无需Python，纯Bash实现）
# 参考标准：配置核查作业指导书v2.2
#           配置核查表_v2.0.0.xlsx（数据库列：SQLServer 共18项）
# 运行方式：sudo bash check_sqlserver.sh
# 连接方式：默认 sqlcmd 连接 sa@localhost,1433；可用环境变量覆盖：
#             MSSQL_HOST=127.0.0.1 MSSQL_PORT=1433 MSSQL_USER=sa MSSQL_PASS=xxx bash check_sqlserver.sh
# 输出文件：output/配置核查报告_SQLServer_日期时间.html /.xls
# !!! 注意：本脚本按 T-SQL/sqlcmd 语法编写，需真实 SQL Server 实例实测。
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

SQLCMD="$(command -v sqlcmd 2>/dev/null)"
# 读取 db_config.conf（KEY=VALUE 格式）
read_config() {
    local key="$1" conf="$SCRIPT_DIR/db_config.conf" val=""
    if [ -f "$conf" ]; then
        val="$(grep -E "^[[:space:]]*${key}[[:space:]]*=" "$conf" 2>/dev/null | tail -1 | sed 's/^[^=]*=[[:space:]]*//; s/[[:space:]]*$//')"
    fi
    printf '%s' "$val"
}

MSSQL_HOST="${MSSQL_HOST:-$(read_config MSSQL_HOST)}"
MSSQL_HOST="${MSSQL_HOST:-127.0.0.1}"
MSSQL_PORT="${MSSQL_PORT:-$(read_config MSSQL_PORT)}"
MSSQL_PORT="${MSSQL_PORT:-1433}"
MSSQL_USER="${MSSQL_USER:-$(read_config MSSQL_USER)}"
MSSQL_USER="${MSSQL_USER:-sa}"
MSSQL_PASS="${MSSQL_PASS:-$(read_config MSSQL_PASS)}"

DB_VERSION=""
CONN_OK=0
MSSQL_PRESENT=0

# 执行 T-SQL，返回干净结果（-h -1 无表头，-W 去尾空格，-C 信任自签名证书）
sql_q() {
    "$SQLCMD" -C -S "$MSSQL_HOST,$MSSQL_PORT" -U "$MSSQL_USER" -P "$MSSQL_PASS" -h -1 -W -Q "SET NOCOUNT ON; $1" 2>/dev/null
}

# 判断 SQL Server 版本是否 EOL（<2017 即主版本<14 视为 EOL）
sqlsrv_eol() {
    local major; major="$(sql_q "SELECT SERVERPROPERTY('ProductMajorVersion');" 2>/dev/null | tr -d '[:space:]')"
    if [ -n "$major" ] && [ "$major" -lt 14 ] 2>/dev/null; then echo "eol"; else echo "supported"; fi
}

guide_ref() {
    case "$1" in
        1.1)  echo "《配置核查作业指导书》第1章 系统安全 1.1：操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序" ;;
        1.7)  echo "《配置核查作业指导书》第1章 系统安全 1.7：数据库管理系统应删除冗余帐户，应设置不少于8个字符且字母大小写、数字及特殊字符混合编制的账户口令" ;;
        1.8)  echo "《配置核查作业指导书》第1章 系统安全 1.8：数据库管理系统应删除冗余存储过程" ;;
        1.9)  echo "《配置核查作业指导书》第1章 系统安全 1.9：数据库管理系统应具有基于表级增删改查等细粒度访问和管理授权功能" ;;
        1.10) echo "《配置核查作业指导书》第1章 系统安全 1.10：数据库管理系统应具有自主访问控制功能" ;;
        1.11) echo "《配置核查作业指导书》第1章 系统安全 1.11：数据库管理系统应具有备份和恢复功能" ;;
        1.12) echo "《配置核查作业指导书》第1章 系统安全 1.12：数据库管理系统应具有表级审计、告警和阻断功能" ;;
        1.13) echo "《配置核查作业指导书》第1章 系统安全 1.13：数据库管理系统的数据应和其它应用的数据分类独立存储" ;;
        1.15) echo "《配置核查作业指导书》第1章 系统安全 1.15：应具备数据库管理系统超级管理员远程登录限制能力" ;;
        1.16) echo "《配置核查作业指导书》第1章 系统安全 1.16：应具备数据库管理系统输入（参数）检查能力" ;;
        1.19) echo "《配置核查作业指导书》第1章 系统安全 1.19：应更换数据库管理系统的默认服务端口、管理员用户名和口令" ;;
        1.21) echo "《配置核查作业指导书》第1章 系统安全 1.21：数据库管理系统应具有行级或列级审计功能" ;;
        1.22) echo "《配置核查作业指导书》第1章 系统安全 1.22：数据库管理系统应采取单独、安全监控、审计措施" ;;
        1.23) echo "《配置核查作业指导书》第1章 系统安全 1.23：数据库管理系统仅为应用服务器提供访问服务" ;;
        1.24) echo "《配置核查作业指导书》第1章 系统安全 1.24：应具备日志审计能力，审计日志至少保留180天" ;;
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
    if [ "$MSSQL_PRESENT" = "0" ]; then
        add_result "$1" "系统安全-数据库" "$2" "na" "未检测到 sqlcmd 客户端，无法对 SQL Server 进行自动核查，标记不适用。" "第1章" "$3"
    else
        add_result "$1" "系统安全-数据库" "$2" "manual" "无法连接 SQL Server（${MSSQL_USER}@${MSSQL_HOST}:${MSSQL_PORT}），需人工登录核查。请检查 MSSQL_USER/MSSQL_PASS 环境变量。" "第1章" "$3"
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

REPORT_TAG="SQLServer"
REPORT_TITLE="SQL Server 数据库配置核查报告"
report_meta_html() {
    echo "<div>连接：$(html_esc "$MSSQL_USER@$MSSQL_HOST:$MSSQL_PORT")　版本：$(html_esc "${DB_VERSION:-未连接}")</div>"
    echo "<div>主机名：$(hostname 2>/dev/null)　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</div>"
    echo "<div>参考标准：配置核查作业指导书v2.2 / 配置核查表_v2.0.0.xlsx</div>"
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
  <div class="sub">参考标准：配置核查作业指导书v2.2</div>
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
    local outfile="$OUT_DIR/配置核查报告_SQLServer_${STAMP}.xls"
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
<x:Name>SQLServer配置核查报告</x:Name>
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
<p>SQL Server 数据库配置核查报告　连接：$(html_esc "$MSSQL_USER@$MSSQL_HOST:$MSSQL_PORT")　版本：$(html_esc "${DB_VERSION:-未连接}")　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</p>
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
# 核查项（共18项）
# ============================================================

check_1_1_patch() {
    local eol; eol="$(sqlsrv_eol)"
    if [ "$eol" = "eol" ]; then
        add_result "1.1" "系统安全-数据库" "数据库补丁程序" "fail" "SQL Server 版本 ${DB_VERSION} 已停止维护（<2017），存在安全漏洞风险。" "第1章" "升级到 SQL Server 2019/2022 并安装最新累积更新。"
    else
        add_result "1.1" "系统安全-数据库" "数据库补丁程序" "manual" "SQL Server 版本 ${DB_VERSION}，请人工核对是否安装最新累积更新(CU)。" "第1章" "及时安装最新累积更新。"
    fi
}

check_1_7_dbaccounts() {
    local sa_policy guest
    sa_policy="$(sql_q "SELECT is_policy_checked FROM sys.sql_logins WHERE name='sa';" 2>/dev/null | tr -d '[:space:]')"
    guest="$(sql_q "SELECT COUNT(*) FROM sys.sql_logins WHERE is_disabled=0 AND name='guest';" 2>/dev/null | tr -d '[:space:]')"
    if [ "$sa_policy" = "0" ]; then
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "fail" "sa 账户未启用密码策略（CHECK_POLICY 未勾选）。" "第1章" "为 sa 及其它账户启用密码策略（长度/复杂度/过期）。"
    elif [ -n "$guest" ] && [ "$guest" != "0" ] 2>/dev/null; then
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "fail" "存在启用的 guest 账户（应禁用）。" "第1章" "禁用 guest 账户，删除冗余账户。"
    else
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "pass" "sa 已启用密码策略，未发现启用状态的 guest 冗余账户。" "第1章" "持续核查并删除/锁定冗余账户。"
    fi
}

check_1_8_dbprocedures() {
    local xp_cmdshell
    xp_cmdshell="$(sql_q "SELECT value_in_use FROM sys.configurations WHERE name='xp_cmdshell';" 2>/dev/null | tr -d '[:space:]')"
    if [ "$xp_cmdshell" = "1" ]; then
        add_result "1.8" "系统安全-数据库" "数据库存储过程管理" "fail" "高危存储过程 xp_cmdshell 已启用（等效系统命令执行）。" "第1章" "禁用 xp_cmdshell 及其它高危存储过程。"
    else
        add_result "1.8" "系统安全-数据库" "数据库存储过程管理" "pass" "高危存储过程 xp_cmdshell 已禁用。" "第1章" "持续禁用高危存储过程。"
    fi
}

check_1_9_dbgrants() {
    local obj_priv
    obj_priv="$(sql_q "SELECT COUNT(*) FROM sys.database_permissions WHERE class_desc IN ('OBJECT_OR_COLUMN','SCHEMA');" 2>/dev/null | tr -d '[:space:]')"
    if [ -n "$obj_priv" ] && [ "$obj_priv" -gt 0 ] 2>/dev/null; then
        add_result "1.9" "系统安全-数据库" "数据库权限最小化" "pass" "存在对象/架构级细粒度授权 ${obj_priv} 条。" "第1章" "持续按最小权限原则，使用对象/架构级授权。"
    else
        add_result "1.9" "系统安全-数据库" "数据库权限最小化" "manual" "未发现对象级细粒度授权，请人工确认是否按最小权限原则分配权限。" "第1章" "按最小权限原则分配账户权限。"
    fi
}

check_1_10_dbaccess() {
    local sysadmin_cnt
    sysadmin_cnt="$(sql_q "SELECT COUNT(*) FROM sys.server_role_members rm JOIN sys.server_principals p ON rm.member_principal_id=p.principal_id JOIN sys.server_principals r ON rm.role_principal_id=r.principal_id WHERE r.name='sysadmin' AND p.name NOT IN ('sa');" 2>/dev/null | tr -d '[:space:]')"
    if [ -n "$sysadmin_cnt" ] && [ "$sysadmin_cnt" != "0" ] 2>/dev/null; then
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "manual" "存在 ${sysadmin_cnt} 个非 sa 的 sysadmin 成员，请人工确认授权合理性。" "第1章" "收敛 sysadmin 角色成员，按最小权限授权。"
    else
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "pass" "sysadmin 角色成员仅 sa，未发现多余超管。" "第1章" "持续控制 sysadmin 成员数量。"
    fi
}

check_1_11_dbbackup() {
    local days
    days="$(sql_q "SELECT ISNULL(DATEDIFF(day, MAX(backup_finish_date), GETDATE()), 9999) FROM msdb.dbo.backupset WHERE type='D';" 2>/dev/null | tr -d '[:space:]')"
    if [ -z "$days" ] || [ "$days" = "9999" ]; then
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "fail" "未发现任何全量备份记录。" "第1章" "建立定期全量备份并验证可恢复性。"
    elif [ "$days" -le 7 ] 2>/dev/null; then
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "pass" "最近一次全量备份距今 ${days} 天（<=7天）。" "第1章" "定期验证备份可恢复性并异地存储。"
    else
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "fail" "最近一次全量备份距今 ${days} 天（>7天），备份不充分。" "第1章" "建立定期备份（建议每天）。"
    fi
}

check_1_12_dbaudit() {
    local audit_cnt
    audit_cnt="$(sql_q "SELECT COUNT(*) FROM sys.server_audits WHERE is_state_enabled=1;" 2>/dev/null | tr -d '[:space:]')"
    if [ -n "$audit_cnt" ] && [ "$audit_cnt" -gt 0 ] 2>/dev/null; then
        add_result "1.12" "系统安全-数据库" "数据库审计插件" "pass" "检测到已启用的服务器审计 ${audit_cnt} 项。" "第1章" "确保审计覆盖登录、权限变更、敏感操作。"
    else
        add_result "1.12" "系统安全-数据库" "数据库审计插件" "fail" "未启用服务器审计（sys.server_audits 无启用项）。" "第1章" "启用 SQL Server 审计记录关键操作。"
    fi
}

check_1_13_dbisolation() {
    add_result "1.13" "系统安全-数据库" "数据库与业务隔离" "manual" "数据分类独立存储需结合业务架构人工核查（分库/分实例）。" "第1章" "敏感数据与普通数据分类独立存储。"
}

check_1_15_dbremote() {
    local sa_disabled
    sa_disabled="$(sql_q "SELECT is_disabled FROM sys.sql_logins WHERE name='sa';" 2>/dev/null | tr -d '[:space:]')"
    if [ "$sa_disabled" = "1" ]; then
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "pass" "sa 账户已禁用，超管远程登录受限。" "第1章" "持续保持 sa 禁用或仅本机访问。"
    else
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "manual" "sa 账户处于启用状态，请人工确认是否限制了远程登录来源。" "第1章" "禁用 sa 或通过防火墙限制其远程访问。"
    fi
}

check_1_16_inputcheck() {
    add_result "1.16" "系统安全-数据库" "输入验证/SQL注入防护" "manual" "数据库输入（参数）检查属应用层能力，需人工/工具核查应用是否使用参数化查询。" "第1章" "应用层使用参数化查询防 SQL 注入。"
}

check_1_19_dbdefaults() {
    local sa_enabled
    sa_enabled="$(sql_q "SELECT CASE WHEN is_disabled=0 THEN 1 ELSE 0 END FROM sys.sql_logins WHERE name='sa';" 2>/dev/null | tr -d '[:space:]')"
    local issues=""
    [ "$MSSQL_PORT" = "1433" ] && issues="仍使用默认端口 1433"
    [ "$sa_enabled" = "1" ] && issues="${issues:+$issues；}sa 账户仍启用"
    if [ -n "$issues" ]; then
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "fail" "$issues" "第1章" "更改默认端口，禁用/加固 sa 账户。"
    else
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "pass" "未发现默认端口/默认账户问题。" "第1章" "持续保持非默认端口与加固账户。"
    fi
}

check_1_21_dbaudit() {
    local spec_cnt
    spec_cnt="$(sql_q "SELECT COUNT(*) FROM sys.server_audit_specifications;" 2>/dev/null | tr -d '[:space:]')"
    if [ -n "$spec_cnt" ] && [ "$spec_cnt" -gt 0 ] 2>/dev/null; then
        add_result "1.21" "系统安全-数据库" "数据库操作审计" "manual" "检测到审计规范 ${spec_cnt} 项，行级/列级审计粒度需人工确认。" "第1章" "配置审计规范覆盖敏感操作。"
    else
        add_result "1.21" "系统安全-数据库" "数据库操作审计" "manual" "未检测到审计规范，行级/列级审计需人工核查是否具备。" "第1章" "配置审计规范。"
    fi
}

check_1_22_monitor() {
    add_result "1.22" "系统安全-数据库" "数据库审计日志留存" "manual" "独立安全监控与审计措施需人工核查部署情况。" "第1章" "部署独立监控/审计平台。"
}

check_1_23_appsep() {
    local listen
    listen="$(ss -lntp 2>/dev/null | awk -v p=":$MSSQL_PORT" '$4 ~ (":"p"$"){print $4}')"
    if echo "$listen" | grep -qE '^(0\.0\.0\.0|\*|\[::\])' 2>/dev/null; then
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "manual" "数据库对外暴露监听：$listen，请人工确认是否仅为应用服务器提供访问。" "第1章" "限制数据库仅为应用服务器提供访问。"
    elif [ -z "$listen" ]; then
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "manual" "未检测到监听，请人工确认数据库是否仅为应用服务器提供访问。" "第1章" "限制数据库仅为应用服务器提供访问。"
    else
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "pass" "数据库未对外暴露监听（$listen），访问受限。" "第1章" "持续确保数据库仅为应用服务器提供访问。"
    fi
}

check_1_24_logaudit180() {
    add_result "1.24" "系统安全-数据库" "日志审计留存180天" "manual" "SQL Server 审计日志留存时长需人工核查（审计文件/错误日志归档策略）。" "第1章" "配置审计日志留存不少于180天。"
}

check_1_25_boundary() {
    local fw fwstate
    fw="$(iptables -S 2>/dev/null | head -20)"
    fwstate="$(firewall-cmd --state 2>/dev/null)"
    if [ -n "$fw" ] || [ "$fwstate" = "running" ]; then
        add_result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "manual" "检测到本机防火墙，但边界防护需结合网络架构人工核查。" "第1章" "部署边界防护，限制暴露面。"
    else
        add_result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "fail" "未检测到本机防火墙规则，数据库缺乏边界保护。" "第1章" "部署防火墙/WAF 限制暴露面。"
    fi
}

check_1_26_log() {
    local errlog
    errlog="$(sql_q "SELECT COUNT(*) FROM sys.dm_os_server_diagnostics_log_configurations;" 2>/dev/null | tr -d '[:space:]')"
    if [ -n "$errlog" ] && [ "$errlog" -gt 0 ] 2>/dev/null; then
        add_result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "pass" "检测到服务器诊断日志配置。" "第1章" "确保日志完整记录、留存满足要求。"
    else
        add_result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "manual" "请人工核查错误日志与补丁日志记录完整性。" "第1章" "配置错误日志输出，保留完整日志。"
    fi
}

check_2_16_patchlatest() {
    local eol; eol="$(sqlsrv_eol)"
    if [ "$eol" = "eol" ]; then
        add_result "2.16" "系统安全-数据库" "补丁修复升级到最新" "fail" "SQL Server 版本 ${DB_VERSION} 已停止维护（<2017），未完成补丁升级。" "第2章" "升级到受支持版本并安装最新累积更新。"
    else
        add_result "2.16" "系统安全-数据库" "补丁修复升级到最新" "manual" "SQL Server 版本 ${DB_VERSION}，请人工核对是否已安装最新累积更新。" "第2章" "定期安装最新累积更新。"
    fi
}

# ============================================================
# 主流程
# ============================================================

echo "================================================================"
echo "  配置核查工具 - SQL Server 版（无需Python）"
echo "  参考标准：配置核查作业指导书v2.2"
echo "================================================================"
echo "  连接：$MSSQL_USER@$MSSQL_HOST:$MSSQL_PORT"
echo ""

if [ -z "$SQLCMD" ]; then
    MSSQL_PRESENT=0
    DB_VERSION="未安装sqlcmd"
    echo "  未检测到 sqlcmd 客户端，所有项将标记为不适用。"
elif DB_VERSION="$(sql_q "SELECT SERVERPROPERTY('ProductVersion');" 2>/dev/null | tr -d '[:space:]')"; [ -n "$DB_VERSION" ]; then
    MSSQL_PRESENT=1; CONN_OK=1
    echo "  SQL Server 版本：$DB_VERSION"
else
    MSSQL_PRESENT=1; CONN_OK=0
    echo "  [!] 无法连接 SQL Server（$MSSQL_USER@$MSSQL_HOST:$MSSQL_PORT），所有项将标记为需人工核查。"
    echo "      请通过 MSSQL_USER/MSSQL_PASS 环境变量提供凭据。"
fi
echo ""
echo "开始检查，请稍候..."

if [ "$CONN_OK" = "1" ]; then
    check_1_1_patch
    check_1_7_dbaccounts
    check_1_8_dbprocedures
    check_1_9_dbgrants
    check_1_10_dbaccess
    check_1_11_dbbackup
    check_1_12_dbaudit
    check_1_13_dbisolation
    check_1_15_dbremote
    check_1_16_inputcheck
    check_1_19_dbdefaults
    check_1_21_dbaudit
    check_1_22_monitor
    check_1_23_appsep
    check_1_24_logaudit180
    check_1_25_boundary
    check_1_26_log
    check_2_16_patchlatest
else
    db_unreachable "1.1"  "数据库补丁程序"          "安装最新累积更新。"
    db_unreachable "1.7"  "数据库账户管理"          "删除冗余账户，启用密码策略。"
    db_unreachable "1.8"  "数据库存储过程管理"      "禁用高危存储过程。"
    db_unreachable "1.9"  "数据库权限最小化"        "按最小权限原则分配权限。"
    db_unreachable "1.10" "数据库访问控制"          "收敛 sysadmin 成员。"
    db_unreachable "1.11" "数据库备份策略"          "建立定期备份。"
    db_unreachable "1.12" "数据库审计插件"          "启用服务器审计。"
    db_unreachable "1.13" "数据库与业务隔离"        "敏感数据分类独立存储。"
    db_unreachable "1.15" "数据库远程访问控制"      "禁用 sa 或限制远程。"
    db_unreachable "1.16" "输入验证/SQL注入防护"    "应用层参数化查询。"
    db_unreachable "1.19" "数据库默认账户/端口"     "改默认端口、加固 sa。"
    db_unreachable "1.21" "数据库操作审计"          "配置审计规范。"
    db_unreachable "1.22" "数据库审计日志留存"      "独立监控平台。"
    db_unreachable "1.23" "应用与数据库账户分离"    "限制仅应用服务器访问。"
    db_unreachable "1.24" "日志审计留存180天"       "审计日志留存180天。"
    db_unreachable "1.25" "边界保护抗攻击防篡改"    "部署防火墙。"
    db_unreachable "1.26" "防病毒/补丁日志记录"     "配置日志输出。"
    db_unreachable "2.16" "补丁修复升级到最新"      "升级到最新版本。"
fi

echo ""
echo "================================================================"
print_summary
generate_html
generate_xls
type generate_xlsx >/dev/null 2>&1 && generate_xlsx "$OUT_DIR/配置核查报告_SQLServer_${STAMP}.xlsx" "SQL Server 数据库配置核查报告"
echo ""
echo "核查完成，请到 output/ 目录查看 HTML/XLS/XLSX 报告。"

#!/bin/bash
# ============================================================
# 配置核查工具 - Tomcat 中间件版（无需Python，纯Bash实现）
# 参考标准：配置核查作业指导书正式版2026_4_1
#           配置核查表_v2.0.0.xlsx（中间件列：Tomcat 共2项）
# 运行方式：sudo bash check_tomcat.sh
# 输出文件：output/配置核查报告_Tomcat_日期时间.html /.xls
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

CATALINA_HOME="${CATALINA_HOME:-/usr/share/tomcat}"
CATALINA_BASE="${CATALINA_BASE:-$CATALINA_HOME}"
SERVER_XML="${SERVER_XML:-$CATALINA_BASE/conf/server.xml}"
TOMCAT_VERSION=""
TOMCAT_PRESENT=0

# 判断 Tomcat 版本是否 EOL（<9.0 视为 EOL）
tomcat_eol() {
    local v="$1" major
    major="$(echo "$v" | grep -oE '^[0-9]+' | head -1)"
    if [ -n "$major" ] && [ "$major" -lt 9 ] 2>/dev/null; then echo "eol"; else echo "supported"; fi
}

guide_ref() {
    case "$1" in
        1.1)  echo "《配置核查作业指导书》第1章 系统安全 1.1：操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序" ;;
        1.14) echo "《配置核查作业指导书》第1章 系统安全 1.14：中间件应采取限制运行权限和使用安全管理通道等安全加固措施" ;;
        4.5)  echo "《配置核查作业指导书》第4章 应用安全 4.5：提供公共信息服务的服务器应具备防DDoS攻击能力" ;;
        4.6)  echo "《配置核查作业指导书》第4章 应用安全 4.6：Web应用系统服务器应采取Web防护措施" ;;
        4.8)  echo "《配置核查作业指导书》第4章 应用安全 4.8：网站应采取网页防篡改措施，防止对信息内容的非法修改" ;;
        4.9)  echo "《配置核查作业指导书》第4章 应用安全 4.9：Web应用系统应具备防范SQL注入、跨站脚本等攻击能力" ;;
        4.22) echo "《配置核查作业指导书》第4章 应用安全 4.22：应更改Web应用系统默认服务发布端口" ;;
        4.23) echo "《配置核查作业指导书》第4章 应用安全 4.23：应分开设置管理端口与应用端口" ;;
        *)    echo "《配置核查作业指导书》" ;;
    esac
}

add_result() {
    R_COUNT=$((R_COUNT+1))
    R_ID[$R_COUNT]="$1"; R_CAT[$R_COUNT]="$2"; R_TITLE[$R_COUNT]="$3"; R_STATUS[$R_COUNT]="$4"
    R_DETAIL[$R_COUNT]="$5"; R_CHAPTER[$R_COUNT]="$6"; R_REC[$R_COUNT]="$7"
    R_GUIDE[$R_COUNT]="$(guide_ref "$1")"
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

REPORT_TAG="Tomcat"
REPORT_TITLE="Tomcat 中间件配置核查报告"
report_meta_html() {
    echo "<div>版本：$(html_esc "${TOMCAT_VERSION:-未检测到}")　主机名：$(hostname 2>/dev/null)　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</div>"
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
    local outfile="$OUT_DIR/配置核查报告_Tomcat_${STAMP}.xls"
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
<x:Name>Tomcat配置核查报告</x:Name>
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
<p>Tomcat 中间件配置核查报告　版本：$(html_esc "${TOMCAT_VERSION:-未检测到}")　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</p>
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
# 核查项（共2项）
# ============================================================

check_1_1_patch() {
    local eol; eol="$(tomcat_eol "$TOMCAT_VERSION")"
    if [ "$eol" = "eol" ]; then
        add_result "1.1" "系统安全" "中间件补丁程序" "fail" "Tomcat 当前版本 ${TOMCAT_VERSION} 已 EOL（<9.0），存在未修复安全漏洞风险。" "第1章" "升级到 Tomcat 9.0/10.1 并安装最新补丁。"
    else
        add_result "1.1" "系统安全" "中间件补丁程序" "manual" "Tomcat 当前版本 ${TOMCAT_VERSION}，请人工核对是否为最新补丁版本。" "第1章" "及时升级到最新稳定版本。"
    fi
}

check_1_14_hardening() {
    local issues=""
    # 1) 运行权限：Tomcat 进程是否降权（非 root）
    local tpid tuser
    tpid="$(pgrep -f 'org.apache.catalina.startup.Bootstrap' 2>/dev/null | head -1)"
    if [ -n "$tpid" ]; then
        tuser="$(ps -o user= -p "$tpid" 2>/dev/null | tr -d ' ')"
        [ "$tuser" = "root" ] && issues="${issues}以root运行；"
    fi
    # 2) 是否配置 HTTPS（SSL 连接器启用）
    local ssl
    ssl="$(grep -E 'SSLEnabled[[:space:]]*=[[:space:]]*"true"' "$SERVER_XML" 2>/dev/null | grep -v '<!--' | head -1)"
    [ -z "$ssl" ] && issues="${issues}未配置HTTPS；"
    # 3) 是否存在默认管理应用（manager/host-manager/docs/examples）
    local defaults
    defaults="$(ls "$CATALINA_BASE/webapps/" 2>/dev/null | grep -E '^(manager|host-manager|docs|examples)$')"
    [ -n "$defaults" ] && issues="${issues}存在默认应用($(echo "$defaults" | tr '\n' ' '))；"
    if [ -z "$issues" ]; then
        add_result "1.14" "系统安全" "中间件安全加固" "pass" "Tomcat 已做基础安全加固（降权运行、配置 HTTPS、移除默认应用）。" "第1章" "持续保持中间件安全配置。"
    else
        add_result "1.14" "系统安全" "中间件安全加固" "fail" "Tomcat 存在以下安全加固缺失：$issues" "第1章" "按缺失项逐条加固中间件配置。"
    fi
}

# ============================================================
# 主流程
# ============================================================

# 补充 Web/中间件相关核查项（按实际安全意义，从应用安全章节纳入）
check_4_5_antiddos() {
    add_result "4.5" "应用安全" "防DDoS攻击能力" "manual" "Tomcat 无内置 DDoS 防护，需结合前置防火墙/WAF，请人工核查。" "第4章" "部署前置防火墙/WAF 防 DDoS。"
}

check_4_6_waf() {
    add_result "4.6" "应用安全" "Web防护措施" "manual" "Tomcat 需前置 WAF 防护，请人工核查是否已部署。" "第4章" "部署 WAF 或前置安全设备。"
}

check_4_8_antitamper() {
    local aide=""
    aide="$(command -v aide 2>/dev/null)"
    [ -z "$aide" ] && aide="$(command -v tripwire 2>/dev/null)"
    [ -z "$aide" ] && [ -x /usr/sbin/aide ] && aide="/usr/sbin/aide"
    if [ -n "$aide" ]; then
        add_result "4.8" "应用安全" "网页防篡改" "pass" "检测到网页防篡改/完整性工具：$aide" "第4章" "定期校验完整性，防止内容被非法篡改。"
    else
        add_result "4.8" "应用安全" "网页防篡改" "manual" "未检测到防篡改工具（AIDE/tripwire），请人工核查是否具备网页防篡改能力。" "第4章" "部署 AIDE/tripwire 等完整性校验工具。"
    fi
}

check_4_9_sqlinjection() {
    add_result "4.9" "应用安全" "防范SQL注入/XSS" "manual" "SQL 注入/跨站脚本防护属应用层能力（代码层），需结合 WAF 与代码审计人工核查。" "第4章" "应用层参数化查询 + WAF 双重防护。"
}

check_4_22_defaultport() {
    local port8080
    port8080="$(grep -E '<Connector[^>]*port="8080"' "$SERVER_XML" 2>/dev/null | grep -v '<!--' | head -1)"
    if [ -n "$port8080" ]; then
        add_result "4.22" "应用安全" "更改默认服务发布端口" "fail" "仍使用默认 8080 端口发布服务。" "第4章" "更改默认 8080 端口为非默认端口。"
    else
        add_result "4.22" "应用安全" "更改默认服务发布端口" "pass" "未使用默认 8080 端口发布服务。" "第4章" "持续保持非默认发布端口。"
    fi
}

check_4_23_mgmtport() {
    local shutdown
    shutdown="$(grep -E '<Server[^>]*port="8005"' "$SERVER_XML" 2>/dev/null | grep -v '<!--' | head -1)"
    if [ -n "$shutdown" ]; then
        add_result "4.23" "应用安全" "管理端口与应用端口分离" "fail" "管理端口（shutdown）仍为默认 8005，应更改。" "第4章" "更改默认 shutdown 管理端口，实现管理与应用分离。"
    else
        add_result "4.23" "应用安全" "管理端口与应用端口分离" "pass" "管理端口（shutdown）已与默认 8005 区分。" "第4章" "持续保持管理端口与应用端口分离。"
    fi
}

echo "================================================================"
echo "  配置核查工具 - Tomcat 中间件版（无需Python）"
echo "  参考标准：配置核查作业指导书正式版2026_4_1"
echo "================================================================"
echo ""

if [ ! -f "$SERVER_XML" ] && [ ! -d "$CATALINA_HOME" ]; then
    TOMCAT_PRESENT=0
    TOMCAT_VERSION="未检测到Tomcat"
    echo "  未检测到 Tomcat，所有项将标记为不适用。"
else
    TOMCAT_PRESENT=1
    TOMCAT_VERSION="$(java -cp "$CATALINA_HOME/lib/catalina.jar" org.apache.catalina.util.ServerInfo 2>/dev/null | grep -oE '[0-9]+\.[0-9]+\.[0-9]+' | head -1)"
    [ -n "$TOMCAT_VERSION" ] && echo "  Tomcat 版本：$TOMCAT_VERSION" || echo "  检测到 Tomcat 但无法获取版本。"
fi
echo ""
echo "开始检查，请稍候..."

if [ "$TOMCAT_PRESENT" = "1" ]; then
    check_1_1_patch
    check_1_14_hardening
    check_4_5_antiddos
    check_4_6_waf
    check_4_8_antitamper
    check_4_9_sqlinjection
    check_4_22_defaultport
    check_4_23_mgmtport
else
    add_result "1.1"  "系统安全" "中间件补丁程序" "na" "未检测到 Tomcat，标记不适用。" "第1章" "如部署 Tomcat 需及时打补丁。"
    add_result "1.14" "系统安全" "中间件安全加固" "na" "未检测到 Tomcat，标记不适用。" "第1章" "如部署 Tomcat 需做安全加固。"
    add_result "4.5"  "应用安全" "防DDoS攻击能力" "na" "未检测到 Tomcat，标记不适用。" "第4章" "如部署 Tomcat 需前置防护。"
    add_result "4.6"  "应用安全" "Web防护措施" "na" "未检测到 Tomcat，标记不适用。" "第4章" "如部署 Tomcat 需部署 WAF。"
    add_result "4.8"  "应用安全" "网页防篡改" "na" "未检测到 Tomcat，标记不适用。" "第4章" "如部署 Tomcat 需配置防篡改。"
    add_result "4.9"  "应用安全" "防范SQL注入/XSS" "na" "未检测到 Tomcat，标记不适用。" "第4章" "如部署 Tomcat 需应用层防护。"
    add_result "4.22" "应用安全" "更改默认服务发布端口" "na" "未检测到 Tomcat，标记不适用。" "第4章" "如部署 Tomcat 需改默认端口。"
    add_result "4.23" "应用安全" "管理端口与应用端口分离" "na" "未检测到 Tomcat，标记不适用。" "第4章" "管理通道与应用通道分离。"
fi

echo ""
echo "================================================================"
print_summary
generate_html
generate_xls
type generate_xlsx >/dev/null 2>&1 && generate_xlsx "$OUT_DIR/配置核查报告_Tomcat_${STAMP}.xlsx" "Tomcat 中间件配置核查报告"
echo ""
echo "核查完成，请到 output/ 目录查看 HTML/XLS/XLSX 报告。"

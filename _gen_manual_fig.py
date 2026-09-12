# -*- coding: utf-8 -*-
"""人工核查台演示图：把 136 项全部填入示例结果 → 截图 → 触发 Excel 导出并取出文件。

做法：复制 manual_check.html 生成临时副本，在草稿载入处注入示例数据（等价于人工逐项填完），
再追加一段脚本把导出内容以 base64 写入页面，用 Chrome 的 --dump-dom 取回。

产出：
  测评报告/验证截图/fig_manual_filled.png   全部填完的人工核查台
  测评报告/验证截图/配置核查报告_人工_*.xls   导出的 Excel 报告（供截图）

说明：填入的是**示例数据**，用于展示工具流程，不代表某个真实系统的核查结论。
"""
import base64
import os
import re
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "manual_check.html")
OUT = os.path.join(BASE, "测评报告", "验证截图")
TMP_DIR = os.path.join(BASE, "_manual_demo")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# 现场信息（示例）
META = {
    "os": "Windows 10 专业版 SP1",
    "host": "PC-ZHANG-01 / 192.168.1.20",
    "user": "张伟",
    "date": "2026-09-11",
}

# 逐项示例结果：未列出的按 PASS 处理；少数项给出针对性备注
FAIL_NOTES = {
    "1.5": "NetBIOS over TCP/IP 未禁用，6 个接口中仅 2 个设为 0x2，其余仍为默认值。",
    "2.11": "口令长度最小值 0、锁定阈值「从不」，屏保未启用，均不满足要求。",
    "2.12": "WLAN AutoConfig 与蓝牙服务 bthserv 均处于运行状态，无线模块未拆除。",
    "1.26": "Windows Defender 支持日志目录缺失，安全事件日志无法读取，无法佐证日志完整性。",
    "4.19": "账户锁定策略未配置，登录失败次数不受限制。",
    "3.3": "未见到存储载体擦除记录，无法确认降密使用前的写覆盖处置。",
    "6.3": "机房线缆标识不完整，互联网区与内部区未做明显线路区分。",
    "8.1": "缺少入网审批制度的正式发文版本，仅有电子草稿。",
}
MANUAL_NOTES = {
    "5.1": "已调阅跨网交换审批单与方案评审意见，需与网络管理部门台账复核。",
    "5.13": "网络安全管理员已指定，管理中心部署位置需现场确认。",
    "7.1": "已提供安全管理机构设置文件，编制人数需人事部门确认。",
    "9.3": "安全保密策略文档已收集，完整性与时效性需主管部门确认。",
    "3.12": "数据采集范围涉及业务系统内部逻辑，需业务部门说明。",
    "4.9": "Web 应用 SQL 注入防护需结合渗透测试结论判断。",
}
NA_NOTES = {
    "1.14": "本机未部署中间件。",
    "2.12": "本机为台式机，无 Wi-Fi/蓝牙硬件。",
}
PASS_EXTRA = {
    "1.2": "已安装奇安信天擎与 Windows Defender，病毒库更新至 2026-09-09。",
    "1.4": "域配置文件、专用配置文件、公用配置文件防火墙均已启用。",
    "1.6": "RDP SecurityLayer=2，已启用 TLS 加密。",
    "2.10": "USBSTOR 已禁用，Start=4。",
    "3.2": "文件传输统一经综合管控平台，留存传输日志。",
    "4.22": "Web 服务发布端口已由 80 改为 8443。",
    "10.1": "协议机密性、完整性、认可性、不可否认性均已核查，符合要求。",
}


def build_draft_js(item_ids):
    lines = ["    draft.meta = " + repr(META).replace("'", '"') + ";"]
    lines.append("    draft.results = {")
    for iid in item_ids:
        if iid in FAIL_NOTES:
            st, note = "fail", FAIL_NOTES[iid]
        elif iid in MANUAL_NOTES:
            st, note = "manual", MANUAL_NOTES[iid]
        elif iid in NA_NOTES:
            st, note = "na", NA_NOTES[iid]
        else:
            st = "pass"
            note = PASS_EXTRA.get(iid, "已现场逐项核对，符合指导书要求。")
        lines.append('      "%s": {status: "%s", note: "%s"},' % (iid, st, note.replace('"', '\\"')))
    lines.append("    };")
    return "\n".join(lines)


CAPTURE_JS = """
<script>
/* 演示脚本：把导出的 Excel 内容写入页面，供 --dump-dom 取回 */
(function(){
  var box = document.createElement("pre");
  box.id = "__dump";
  box.style.display = "none";
  document.body.appendChild(box);
  var B = "@@" + "XLSBEGIN" + "@@", F = "@@" + "XLSFILE" + "@@", E = "@@" + "XLSEND" + "@@";
  window.downloadBlob = function(content, filename, mime){
    box.textContent = B + btoa(unescape(encodeURIComponent(filename))) + F + btoa(unescape(encodeURIComponent(content))) + E;
  };
  window.__doExport = function(){
    document.getElementById("export-xls").click();
    return box.textContent.length;
  };
})();
</script>
"""


def main():
    src = open(SRC, encoding="utf-8").read()
    ids = re.findall(r'"id": "([\d.]+)"', src)
    print("核查项数：", len(ids))

    anchor = "if (!draft.results) draft.results = {};"
    assert anchor in src, "未找到草稿初始化锚点"
    src = src.replace(anchor, anchor + "\n" + build_draft_js(ids), 1)
    # 注意：导出函数内部有 '</body></html>' 字符串，必须锚定文件末尾那一个
    _tail = "</script>\n</body>\n</html>"
    assert _tail in src, "未找到文件末尾结构"
    src = src.replace(_tail, "</script>\n" + CAPTURE_JS + "</body>\n</html>", 1)

    os.makedirs(TMP_DIR, exist_ok=True)
    os.makedirs(OUT, exist_ok=True)
    demo = os.path.join(TMP_DIR, "prefilled.html")
    open(demo, "w", encoding="utf-8").write(src)
    url = "file:///" + demo.replace("\\", "/")
    print("临时页面：", demo)

    # 1) 截图：填满的人工核查台
    shot = os.path.join(OUT, "fig_manual_filled.png")
    cmd = [CHROME, "--headless=new", "--disable-gpu", "--no-sandbox", "--hide-scrollbars",
           "--force-device-scale-factor=1.5", "--virtual-time-budget=12000",
           "--window-size=1400,1500", "--screenshot=" + shot, url]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    print("截图 rc=%d size=%s" % (r.returncode, os.path.getsize(shot) if os.path.exists(shot) else 0))

    # 2) 取导出内容
    dump_demo = os.path.join(TMP_DIR, "export.html")
    src2 = src.replace("</body>\n</html>", "<script>window.__doExport();</script></body>\n</html>", 1)
    assert src2 != src, "未能插入导出触发脚本"
    open(dump_demo, "w", encoding="utf-8").write(src2)
    url2 = "file:///" + dump_demo.replace("\\", "/")
    cmd = [CHROME, "--headless=new", "--disable-gpu", "--no-sandbox", "--dump-dom",
           "--virtual-time-budget=12000", url2]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    m = re.search(r"@@XLSBEGIN@@(.*?)@@XLSFILE@@(.*?)@@XLSEND@@", r.stdout, re.S)
    if not m:
        print("未取到导出内容；dump 长度", len(r.stdout))
        print(r.stdout[-500:])
        return 1
    def _d(s):
        s = re.sub(r"\s+", "", s)
        return base64.b64decode(s + "=" * (-len(s) % 4)).decode("utf-8")
    fname = _d(m.group(1))
    xls = _d(m.group(2))
    out_xls = os.path.join(OUT, fname)
    open(out_xls, "w", encoding="utf-8").write(xls)
    print("导出报告：", fname, "->", out_xls, len(xls), "字符")
    rows = xls.count("<tr>")
    print("导出表格行数（含表头）：", rows)
    return 0


if __name__ == "__main__":
    sys.exit(main())

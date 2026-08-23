# 配置核查项目 · AI 交接文档

> 本文档面向后续接手本项目的 AI 助手（或开发者），覆盖项目全貌、关键约定、工具链、验证清单、已知坑与待办。读完本文应能独立继续开发。
> 最后更新：2026-08-22（对应提交 2c16e35 之后）

---

## 一、项目是什么

**工作主线：配置核查表 → 配置核查作业指导书 → 自动化核查工具**，三件套互相咬合：

| 产出物 | 文件 | 状态 |
|---|---|---|
| 核查表（基准矩阵） | `配置核查表_v2.0.0.xlsx` / `_标注自动验证.xlsx` | 已定版，标注版含每项自动化状态 |
| 作业指导书（标准依据） | `配置核查作业指导书_v2.2.docx` | v2.2 已修复+完善（10 章、113 个编号检查项） |
| 自动化工具（双平台） | `win/`（VBScript）、`kylin/`（Bash）各 7 个脚本 | 可用，HTML 报告为 2026-08 新模板 |
| 交叉验证报告 | `测评报告/指导书与核查工具交叉验证报告.docx/.md` | 113 项全覆盖：已实现 75 / 部分 38 / 未实现 0，自动化率 66.4% |
| 交付包 | `checklist_tool/`（仓库内快照）+ `Desktop\checklist_tool`（实物） | 37 文件，含总 README |

**仓库**：`https://github.com/1111307/checklist_tool.git`（分支 main）。
**仓库位置**：`C:\Users\ryan.xiong\Desktop\peizhitool\配置核查\`。

### ⚠️ 位置陷阱（必读）
桌面还有一个**旧位置** `C:\Users\ryan.xiong\Desktop\配置核查\`，是项目搬迁前的残留副本（含修复前的旧指导书）。**任何读写都以 `peizhitool\配置核查` 为准**，旧位置文件勿混淆——曾发生过用户看着旧副本误以为修复失效的事故。

---

## 二、目录地图

```
配置核查/                          ← 仓库根（工作区）
├── win/                          Windows 核查脚本（VBScript，GBK 编码！）
│   ├── check_xp7.vbs             OS 核查（139 项），含 2.3/3.6 补测
│   ├── check_{mysql,redis,dm,sqlserver,nginx,tomcat}.vbs + .ps1
│   ├── run.bat / run_all.bat     一键入口（注意：win 侧无组件自动探测，全量跑）
│   ├── merge_report.vbs / merge_xlsx.ps1 / lib_xlsx.ps1
│   └── output/                   历史报告（含 2026-08-21 新模板实跑样张）
├── kylin/                        麒麟核查脚本（Bash，UTF-8）
│   ├── check_kylin.sh            OS 核查（139 项）
│   ├── check_{mysql,redis,dm,sqlserver,nginx,tomcat}.sh
│   ├── run.sh / run_all.sh       run_all 有组件自动探测（进程/端口/目录四信号）
│   ├── lib_xlsx.sh / merge_xlsx.sh
│   └── output/                   真机历史报告（20260710-20260809）
├── docker/                       Docker 靶机环境（Dockerfile + systemctl3.py + 90 份报告）
├── 测评报告/                      交叉验证报告 + 分章 md
├── checklist_tool/               交付包快照（干净版，无 output）
├── 配置核查作业指导书_v2.2.docx   正式文件（42MB，含 1019 张截图）
├── 配置核查作业指导书_v2.2.md     docx 的 md 副本（_to_md.py 生成）
├── 配置核查表_v2.0.0*.xlsx
├── _gen_guide_check.py           交叉验证报告生成器（解析指导书+脚本）
├── _to_md.py                     指导书 docx → md
├── md2docx.py                    md → docx（报告用）
├── gen_xlsx.py / fix_docx_format.py / apply_guide_v22.py   历史工具
└── manual_review/                 人工核查建议文档
```

---

## 三、指导书 v2.2 关键约定（改 docx 前必读）

### 样式系统（非标准！）
- 样式 ID 是 **`'2'`~`'6'`**，不是 `Heading 1` 等：`'2'`=H1(章)、`'3'`=H2(条目)、`'4'`=H3(子目)、`'5'`=H4(平台)、`'6'`=H5
- 用 python-docx 读写时按 `w:pStyle w:val` 映射，**直接写 `Heading 1` 字符串会破坏样式引用**（踩过：曾导致全部标题解析失败）
- 层级语义：H1=章（`1 系统安全`）、H2=条目（`1.1 …`）、H3=子目（`1.1.1 …`）、H4=平台（`1.1.1.1 Windows…`）

### 编号位置约定
- 第 1 章：数据库小节按**产品**分（`1.7.1 Mysql`…），部分小节（1.1/1.24/1.27）有 x.y.z 分组层
- 第 3/4 章：`.1`=Win7\WinXP、`.2`=中标麒麟、银河麒麟、`.3`=通用/补充核查方法（个别小节无 Win 子节时从 .2 起编，如 4.13/4.22/4.33，**这是约定不是错**）
- 平台标记既有样式子节标题、也有**普通段落伪标题**（“中标麒麟、银河麒麟操作系统下的核查方法”，第 1 章有 70 处此句式——按文本匹配定位时必须限定小节范围，首个匹配会抓错）

### 内容格式
- 分点一律一段式：`（N）标题：内容`（2026-08-22 已统一 117 处）
- 每个小节（H2）标题后紧跟 `判定标准：…` 段（第 1/2 章已 27/27、16/16 补齐）
- 通用/补充核查方法小节结构：`【说明】→（1）（2）（3）→【判定】`
- 图：1019 张全部居中，**无题注**（自动生成的 888 个题注因质量差已全部删除；恢复方案见 git 历史 `8c8e0f1` 之前）
- 文档结构：封面（无页码）→ 目录（TOC 域，罗马数字页码）→ 正文（阿拉伯数字从 1 起）；打开需右键目录"更新域"填页码

### 历史修复基线（校验时对照）
- 8 类内容硬错误已清零：`SERVERPROPERTYC`、`ROUTINE_ TYPE`、`ROUTINE TYPE/NAME`、`validate_password.dll`、`employeesFOR`、`触发器DELIMITER`、`END IF;END`
- 35 处"通用核查方法"标题压尾倒序已归位；4.24 麒麟子节编号 4.23.2→4.24.2 已修
- 第 1 章 69 个 x.y.z 标题 H4→H3 已归位；1.17 平台标题已前置
- 达梦 19 子节 Win+麒麟双平台覆盖（1.15.3 金仓混入内容已清）

---

## 四、脚本侧关键约定

### 编码（最大的坑）
- **win/*.vbs 是 GBK**：一律用 Python `open(path, encoding='gbk')` 读写打补丁；直接用编辑器改中文会乱码
- kylin/*.sh 是 UTF-8（LF）
- VBS 生成的 HTML 用 `<meta charset="GBK">`（CreateTextFile ASCII 模式）；麒麟侧 UTF-8

### HTML 报告模板（2026-08-21 重写）
两平台 14 个脚本同一设计：渐变头横幅 + meta 条 → 四状态指标卡（可点击筛选）→ 搜索框+筛选按钮 → 徽章表格。数据以 JSON 内嵌 `var DATA = [...]`，前端渲染，**零依赖离线可开**。
- 麒麟：`generate_html()` heredoc 三段式（HTMLHEAD/HTMLMID/HTMLFOOT），注意 **heredoc 终结符必须独占一行**
- win：`GenerateHTML()` 逐行 `ts.WriteLine`，占位符用 `@@VBS:expr@@` 标记法防引号翻倍错乱；`JsonEsc` 的转义顺序必须先 `\` 后 `"`
- 参考标准字样统一为 `配置核查作业指导书v2.2`

### 数据结构（两侧对齐）
- 麒麟：`R_CHAPTER/R_ID/R_CAT/R_TITLE/R_STATUS/R_DETAIL/R_REC/R_GUIDE` 数组 + `R_COUNT`
- win：`rChapter/rID/rCat/rTitle/rStatus/rDetail/rRec` + `rCount`（无 rGuide，章节列拆为 ch/guide）
- 状态四态：`pass/fail/manual/na`
- `add_result "编号" "类别" "标题" "状态" "详情" "章节" "建议"` / `AddResult(...)` 参数序

### 核查项覆盖
- OS 脚本各 139 项（含 2026-08-21 补测的 2.3 应用口令、3.6 边界防泄漏）
- 交叉验证映射：指导书 x.y 编号 ↔ 脚本 AddResult 编号，`_gen_guide_check.py` 自动比对

---

## 五、工具链与常用操作

```bash
# 交叉验证报告重跑（改指导书或脚本后必须跑）
python _gen_guide_check.py && python md2docx.py 测评报告/指导书与核查工具交叉验证报告.md 测评报告/指导书与核查工具交叉验证报告.docx

# 指导书 md 副本重生成
python _to_md.py

# win 脚本实跑（本机可跑，约 2 分钟）
cd win && cscript //Nologo check_xp7.vbs

# kylin 脚本本机干跑（Git Bash，命令会失败但能出报告验证模板）
cd kylin && timeout 180 bash check_kylin.sh
# 语法检查
for f in kylin/check_*.sh; do bash -n $f; done

# 交付包同步（改完任何东西）
cp win/check_*.vbs win/check_*.ps1 win/run*.bat win/*.conf win/*.ps1 "C:/Users/ryan.xiong/Desktop/checklist_tool/win/"
cp kylin/check_*.sh kylin/run*.sh kylin/*.conf "C:/Users/ryan.xiong/Desktop/checklist_tool/kylin/"
cp 配置核查作业指导书_v2.2.docx 配置核查表_v2.0.0_标注自动验证.xlsx "C:/Users/ryan.xiong/Desktop/checklist_tool/"
cp 测评报告/指导书与核查工具交叉验证报告.docx "C:/Users/ryan.xiong/Desktop/checklist_tool/"
# 注意：交付包文件被 Word 占用时 cp 会报 Device or resource busy，关 Word 重试
```

### 改 docx 的方法论（血泪经验）
1. **一律写 Python 补丁脚本，先干跑（预览）再 `--apply`**，全程 `cp` 备份
2. **用元素引用操作（lxml `addprevious/addnext`），不要用索引搬节点**——索引会漂移，元素引用不会（35 处错位修复的成功关键）
3. 改完必跑验证清单（见下节），数据类指标（图片数/题注数/段落数/标题数）必须能精确对账
4. python-docx 的 `Document.paragraphs` 遍历的是 w:p 元素；`p._p` 是元素、`p.runs` 是文本 run

### 验证清单（每次改指导书后）
```python
# 必查项（历史脚本片段已散见各次会话，核心逻辑如下）
1. 8 类硬错误正则 = 0
2. 图片段数（基线 1019）、标题数 409+、空挂 H3 = 0
3. H3 编号与父 H2 前缀匹配 = 0 错配；无重复编号
4. 第1章判定标准 27/27、第2章 16/16
5. （N）分点无"标题独段"残留
6. 达梦 19 子节 Win+麒麟双覆盖
```

---

## 六、已知坑备忘（踩过的雷）

| 坑 | 后果 | 规避 |
|---|---|---|
| VBS 文件被当 UTF-8 编辑 | 中文全部乱码 | 只用 Python GBK 读写 |
| docx 样式写成 `Heading 1` 名字 | 409 个标题样式引用全断 | 用样式 ID `'2'`~`'6'` |
| heredoc 终结符写在内容行内 | bash 语法错误 | 终结符独占一行 |
| VBS `Replace` 转义层数算错 | "未结束的字符串常量" | `"\\"`→`"\\"`，`"""`→`"\\"` 按目标 JS 转义逐层推导 |
| 按文本"首个匹配"定位段落 | 抓走别小节的同句式段落（70 处麒麟伪标题事故） | 锚点必须限定在小节范围内 |
| md5sum 比对 Windows 绝对路径 | 哈希带 `\` 前缀误报不一致 | 用 `cmp -s` 逐字节比对 |
| 用索引搬 docx 节点 | 索引漂移致内容错位 | 元素引用法 |
| 只看代码不实跑 | 3 个模板 bug 全是实跑暴露的 | 两侧都要真跑出报告校验 |

---

## 七、待办与已知未竟事项

1. **麒麟真机验证新版 HTML 模板**：本机只做了 bash 干跑（139 项四类对齐），需上银河麒麟真机跑一遍 `run_all.sh` 确认
2. **win/run_all.bat 无组件自动探测**（麒麟侧已实现，win 侧仍是全量跑）——可仿照 kylin/run_all.sh 的 detect_* 函数用 tasklist/sc query 实现
3. **指导书第 5 章「网络安全」是空占位**（"目前无网络设备"），23 个检查项未纳入交叉验证
4. **交叉验证 38 项"部分实现"**：第 6-9 章 22 项纯人工（脚本只出清单即可）、其余 16 项受环境限制（详见报告第五节）
5. **指导书 md 副本与 docx 的同步**：每次改 docx 后记得 `python _to_md.py`
6. 旧位置 `Desktop\配置核查\` 副本待用户确认后清理（含 `xxxxx/` 目录下 v1.1.0 历史版本未比对）

---

## 八、git 规范

- 提交粒度：一个独立修复/功能一个 commit，消息格式 `fix:/feat:/docs:/style:/revert: 中文描述`，正文写明**验证数据**（如"139 项四类对齐"）
- 推送：`git push origin main`（用户已知情并明确全部推送，含敏感内容——见会话记忆，勿再重复确认）
- 大文件：指导书 docx 42MB 直接提交（GitHub 单文件 <100MB 限制内）

---

## 九、会话记忆提示

本项目有 ZCode 自动记忆（`MEMORY.md` 索引），关键记忆：项目总览、指导书敏感性约束、中文正式文档排版标准、失败先全量核查、生成类改动必须实跑产物校验。接手时这些记忆会自动加载；本文档是其仓库内的持久化版本。

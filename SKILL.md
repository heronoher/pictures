---
name: local-image-doubao-recognition
description: >-
  本地图片高精度结构化视觉识别（豆包桌面客户端调用）。用户主动要求识别本地图片/文件夹，或当前任务必须获取本地图片视觉信息时，
  无需授权询问，直接自动打开或复用豆包、上传图片、发送识别指令、读取结构化结果，并根据任务需求进行解读。支持 JPG/JPEG/PNG/BMP/WebP；
  支持 PPT/PPTX：若用户只要 PPT 的文字内容概览（总结/讲了什么/大纲），直接把 PPT 文件本体上传豆包做文本分析（步骤0.6，不拆解渲染）；若用户具体问 PPT 页面内的图片内容，则自动拆解（整页渲染 PNG + 提取 ppt/media 内嵌原图）后按图片流程识别。
  不处理在线图片链接、纯文件属性查询，以及上下文中已有该图片结构化结果的场景。
---

# 本地图片结构化识别（豆包调用）

版本：v3.10（窗口定位统一为按进程 ID，去标题依赖；历史要点见下方「版本变更说明（简要）」）

## 版本变更说明（简要）

- **v3.10**：窗口定位统一为「按进程 ID + 最大窗口」（新增 1.6 通用函数 `Get-DoubaoWinByProcess`），消除标题变化导致的窗口“丢失—恢复”抖动；回归实测轮询 `missCount=0`。
- **v3.9**：稳定提速组合——B1 结果先落盘校验、失败「仅重抓」不重传；B2 输入框早失败探测（10s）；B3 回收站/占用文件先复制再渲染上传；A1 整批粘贴统一校验+尾部补贴；A4 media 提取默认跳过。
- **v3.8**：汇报排版——分章节拆解逐页呈现；核心重点汇总/结构评价按页码升序、逐条分行。
- **v3.7**：固定指令去 ABC/长串编号，改中文序号分层；「核心重点汇总」条目化细则；汇报格式规范。
- **v3.6**：0.6 指令加来源页码标注；回复转储逐元素容错（防 ∞ 中断）。
- **v3.5**：新增 0.6 PPT 内容分析轻量模式——只问内容→直传 PPT 文件本体，不拆解渲染；问页面图片内容才走步骤0。
- **v3.4**：新增 0.5 产物清理纪律（收尾 try/finally 清理、防误删黑名单、链接策略；不做启动前扫描）。
- **v3.3**：单批 10 张硬上限（豆包单条附件上限）；超量分批、批间必须新建对话。
- **v3.2**：新增步骤0 PPT/PPTX 拆解识别（COM 整页渲染 PNG；media 提取按需）。
- **v3.1.1**：窗口恢复失败自动 `Start-Process --force-renderer-accessibility` 兜底。
- **v3.1**：1.5 恢复激活+三重校验；3.3 附件数+会话纯净双校验。
- **v3.0**：「识别完毕」结束标记（拆字写法防误判），标记出现约 1.5s 判完成。
- **v2.9**：轮询只统计“字段名：内容”回复行；末字段/≥8 行+3s 稳定兜底判完成。
- **v2.8 及以前**：UIA+剪贴板自动化链路成型；用户请求识别即直接调用豆包，无授权询问。

## 前置校验

执行前必须完成：

1. 绝对路径化：用户给出的相对路径（如 `CSY`、`.`）必须先转换为绝对路径；禁止将 `.` 或相对路径传给文件工具。Windows 下使用 `D:\CSY` 这类绝对路径；若不确定，先用 `Resolve-Path` 或 `Get-Item` 获取完整路径。
2. 路径校验：确认目标文件/文件夹路径真实存在，无拼写错误。
3. 格式校验：过滤出支持的图片格式：JPG、JPEG、PNG、BMP、WebP；剔除无效格式文件并记录。**扩展名为 .ppt/.pptx 的文件不剔除**——转入步骤0 拆解后按图片流程处理。
4. 状态校验：确认文件未损坏、可读，无系统权限限制。
5. 数量统计：单张图片直接进入流程；多张/文件夹场景统计有效图片总数；**单批上限 10 张（豆包单条消息附件上限，v3.3）：总数 ≤10 一批完成；总数 >10 必须分批，每批 ≤10 张、每批独立新对话（见步骤 3.4）**；PPT 场景以渲染页数计（每页 1 张整页 PNG），页数 >10 时按每批 ≤10 页分批识别。
6. 合规初筛：基于文件元数据及缩略图做基础合规校验，疑似违规敏感内容直接终止，输出「无法识别该图片」。
7. 类型分流：目标为图片文件 → 直接进入步骤1；目标为 .ppt/.pptx → 先按步骤0.6 分流：**用户只问 PPT 文本内容（总结/主要讲了什么/大纲）→ 轻量直传模式**（把 PPT 文件本体上传给豆包 + 固定文本分析指令）；**用户具体问 PPT 里的图片/照片是什么内容 → 步骤0 拆解渲染**后按图片流程处理。

## 核心执行流程

### 步骤0：PPT/PPTX 文件预处理（v3.2 新增）

**触发条件**：前置校验判定目标是 .ppt/.pptx 文件时执行本步骤；处理完产出整页 PNG，后续流程与普通图片完全一致。

核心思路：**豆包只收图片，所以先把 PPT 变成图片**，把"PPT 里的图片内容是什么"转化为"这张图里有什么"。整页渲染还能同时转录每页文字——解决旧流程"只能读 PPT 文字、读不出 PPT 图片"的问题。

#### 0.1 拆解产物与通道选择

- **整页渲染 PNG（主通道，默认必须）**：每页 1 张 PNG，一次识别拿到每页的图片内容 + 版式布局 + 页面文字。渲染依赖本机安装 Microsoft PowerPoint（桌面版）或 WPS 演示。
- **media 内嵌原图（细节通道，按需）**：.pptx 本质是 zip，内嵌图片在 `ppt/media/`，无需 Office 即可解出；用于整页图上图片区域太小/模糊时上传原图追问细节。⚠️ media 文件编号与页码**无直接映射**，不能据此判断图在哪一页。
- **区域裁剪（细节通道）**：某页局部图看不清时，用 System.Drawing 从该页整页 PNG 裁剪目标区域另存后上传（旧版 crop 切片思路）。

#### 0.2 渲染整页 PNG（主通道）

> v3.9（B3）：源文件位于回收站（`$RECYCLE.BIN`）或被占用导致 COM 直接打开失败（实测 E_FAIL）时，自动先复制到临时目录再打开重试；复制物属于本任务临时文件，随 0.5 规则2 清理，原文件绝不删除。

```powershell
$pptPath = 'D:\xxx\演示文稿.pptx'   # 已绝对路径化
$outDir = Join-Path 'D:\DS\.dsh' ('tmp_ppt_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# ① 优先 Microsoft PowerPoint COM；不可用则尝试 WPS 演示 COM（Kwpp.Application，接口相似）
$app = $null
try { $app = New-Object -ComObject PowerPoint.Application }
catch { try { $app = New-Object -ComObject Kwpp.Application } catch { } }
if (-not $app) {
    Write-Error '未安装 PowerPoint/WPS，无法整页渲染；.pptx 可退化为仅提取 media 原图（见 0.3）'
    exit 1
}
try {
    # Presentations.Open(路径, ReadOnly, Untitled, WithWindow=false)
    # B3（v3.9）：直接打开失败（回收站/占用等，E_FAIL）→ 复制到临时目录重试一次
    $pres = $null
    try { $pres = $app.Presentations.Open($pptPath, $true, $false, $false) }
    catch {
        $srcDir = Join-Path $outDir 'src'
        New-Item -ItemType Directory -Path $srcDir -Force | Out-Null
        $workPath = Join-Path $srcDir ([System.IO.Path]::GetFileName($pptPath))
        Copy-Item -LiteralPath $pptPath -Destination $workPath -Force
        $pres = $app.Presentations.Open($workPath, $true, $false, $false)
        Write-Output 'OPEN_VIA_TEMP_COPY'
    }
    $sw = $pres.PageSetup.SlideWidth; $sh = $pres.PageSetup.SlideHeight
    $outW = 1280; $outH = [int]($outW * $sh / $sw)   # 按页面真实比例导出
    $slideCount = $pres.Slides.Count
    foreach ($slide in $pres.Slides) {
        $file = Join-Path $outDir ('slide_{0:D3}.png' -f $slide.SlideIndex)
        $slide.Export($file, 'PNG', $outW, $outH)
    }
    $pres.Close()
} finally {
    if ($app) { $app.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
"RENDERED $slideCount slides → $outDir"
```

#### 0.3 提取 media 内嵌原图（细节通道，.pptx 适用；v3.9 起默认跳过）

> v3.9（A4）：本步骤**默认不执行**（省解压时间与磁盘占用，正常拆解任务不需要）。仅当用户追问细节、整页图上局部图太小/模糊、需要上传原图时，才临时执行本步骤取图；旧版 .ppt 无法解包，同样跳过本步骤、局部细节走渲染图裁剪路径。

```powershell
if ($pptPath -like '*.pptx') {
    $tmpZip = Join-Path $outDir 'src.zip'
    Copy-Item -LiteralPath $pptPath -Destination $tmpZip
    Expand-Archive -LiteralPath $tmpZip -DestinationPath (Join-Path $outDir 'unpacked')
    $mediaDir = Join-Path $outDir 'unpacked\ppt\media'
    if (Test-Path -LiteralPath $mediaDir) {
        New-Item -ItemType Directory -Path (Join-Path $outDir 'media') -Force | Out-Null
        $imgs = Get-ChildItem -LiteralPath $mediaDir -File |
            Where-Object { $_.Extension -in '.jpg', '.jpeg', '.png', '.bmp', '.webp', '.gif' }
        foreach ($im in $imgs) {
            Copy-Item -LiteralPath $im.FullName -Destination (Join-Path $outDir 'media')
        }
        "EXTRACTED $($imgs.Count) media images → $(Join-Path $outDir 'media')"
    } else {
        'NO_MEDIA（可能全部为矢量图形或无内嵌位图）'
    }
}
```

#### 0.4 识别组织规则

1. **默认只上传整页渲染 PNG**，按页码顺序分批（每批 ≤ 10 页）；页数多时先询问用户关注范围，或按顺序分批识别、分批汇总。
2. 每页结果按「页码」组织输出，附该页整页 PNG 链接（如 `[第03页](file:///D:/.../slide_003.png)`），**同时注明该临时文件会在任务收尾清理后失效**；用户需要长期可点击的页图时，先把 slide PNG 转存到用户指定目录（如原 PPT 同目录的 `_slides\` 子目录）再提供链接（见步骤0.5）。
3. 用户问"某页的图片是什么/细节看不清"时：先对该页渲染图追问；仍不清再上传该页对应区域的**裁剪图**或 **media 原图**单独识别（0.3/裁剪），并按步骤7 追问。
4. PPT 页面文字随整页识别由「画面文字」字段一并转录；若用户要"PPT 全文文字稿"，可在整页识别后追问豆包逐页完整转录，模糊处标「[?]」，不臆测。
5. 旧版 .ppt（非 .pptx）无法解包提取 media，只能整页渲染；局部细节一律走渲染图裁剪/放大路径。
6. 步骤0 产物目录属于**本次任务临时文件**，按步骤0.5 清理纪律处理：任务收尾（含失败/中断）必须删除，禁止为了保留可点击链接而把产物目录留在磁盘上；绝不删除用户原始 PPT。

#### 0.5 产物清理纪律（v3.4 新增，防磁盘垃圾文件堆积）

适用所有图片/PPT 任务（PPT 拆解产物与辅助脚本最容易遗留）。三条硬规则。**明确不做「任务启动前的主动扫描」**（每次任务先扫盘会显著增加任务耗时）；防堆积依靠规则2 收尾清理，历史遗留物仅在用户发现并指出时按规则3 核对清理。

**规则 1：统一产物位置与命名。** 本任务所有临时文件只允许出现在流程声明的临时目录（PPT 任务 = `D:\DS\.dsh\tmp_ppt_<8位hex>`）；辅助脚本（.ps1）写在哪就用在哪、用完即删，不允许把辅助脚本留在 `D:\DS\.dsh` 根目录或任何工作目录里过夜。

**规则 2：任务收尾 try/finally 兜底清理（防本次任务留垃圾）。** 每个批次读取结果并落盘后、全部批次汇总完成后、以及任何失败/中断路径，都必须删除本次产物目录与辅助脚本（等价 try/finally，先清理再汇报）：

```powershell
# 汇报给用户之前执行（无论成功/失败/中断都执行）
Remove-Item -LiteralPath $outDir -Recurse -Force -ErrorAction SilentlyContinue        # PPT 拆解产物目录
Remove-Item -LiteralPath $helperScript -Force -ErrorAction SilentlyContinue          # 本次编写的辅助脚本
```

- **禁止为了「让汇报里的临时链接可点击」而保留产物目录**——临时链接在收尾后失效是正常预期，汇报中注明即可；用户确实要保留页图/中间结果时，先征询并转存到用户指定位置（如原 PPT 同目录的 `_slides\`、`_results\` 子目录），转存成功后再提供链接。
- 失败/中断场景：确认豆包窗口与前台还原后，同样执行上述清理；若因进程占用删不掉，记录路径并向用户报告，由用户决定何时清理。

**规则 3：防误删黑名单（永不删除）。** 清理动作只针对「命名约定内 + 内容核对无误的本次流程产物」；以下一律不碰：
- `D:\DS\.dsh\skills\`（技能本体及历史）；
- 用户原始图片/PPT 及其所在目录（含 Doubao 聊天目录、桌面、文档等任何位置的源文件）；
- 本次任务开始前就已存在的任何文件（历史脚本如 `D:\DS\ocr.ps1`、历史产物之外的目录等）；
- 无法确认归属的文件/目录。

（历史遗留物不主动扫描；若用户发现并指出，执行者按上述黑名单核对归属后再清理。）

#### 0.6 PPT 内容分析轻量模式（v3.5 新增、v3.6 更新指令与防中断：只问内容时直接上传 PPT 文件本体）

**触发条件**：用户只问这个 PPT 的"主要内容 / 讲了什么 / 总结 / 大纲 / 观点"，**不涉及页面内图片内容**。此时不渲染、不拆页，直接把 .ppt/.pptx 文件本体当附件发给豆包。

**执行流程**（与"发图片给豆包"相似，唯一区别是附件为 PPT 文件本体、指令换为下方固定文本分析指令）：
1. 步骤1.5 恢复激活豆包窗口（确认 `DOUBAO_READY`）；
2. 步骤2 新建对话（Ctrl+Shift+K）并聚焦输入框；
3. **上传 PPT 文件附件**：先复制该文件（剪贴板文件列表 `SetFileDropList`），在输入框 Ctrl+V 粘贴生成附件；或点击输入区附件按钮选择文件。上传后校验：窗口内出现含该文件名或 `.pptx`/`.ppt` 的元素（附件已就位）；附件上传需要时间（大文件等待更久），校验失败重试一次。**v3.9（B3）**：若源文件位于回收站（`$RECYCLE.BIN`）或被占用导致附件粘贴失败，先把文件复制到临时目录 `D:\DS\.dsh\tmp_ppt_<hex>\src\`，再上传副本（副本随 0.5 规则2 清理，原文件不动）。
4. **完整发送下方固定「PPT 文本分析指令」**（不允许删减修改；发送方式同步骤4：剪贴板粘贴 + 发送按钮/Enter）；
5. **轮询完成判定差异**：该回复是长篇文本分析，**不是 8 字段图片结构、也不要求输出结束标记**——5.1 的图片字段判据不适用，改用兜底判据：窗口文本总量明显超过基线（>基线+300）且连续约 3 秒不再增长 → 判完成；超时上限放宽到 240 秒（整份 PPT 分析较慢）；
6. **读取回复（健壮转储，防脚本中断）**：从含「PPT类型判定结果」或框架标题的元素开始提取完整回复。转储代码必须**逐元素容错**——实测豆包窗口里个别元素 `BoundingRectangle` 会返回无穷值（∞），直接 `[int]$r.Y` 强转会抛异常，在 `$ErrorActionPreference='Stop'` 下会让整段读取中断、结果丢失：
   ```powershell
   # 正确写法：每个元素独立 try/catch；几何值非有限时置 -1 或跳过；单元素异常绝不中断整段
   $lines = New-Object System.Collections.Generic.List[string]
   foreach ($el in $all) {
       try {
           $name = $el.Current.Name
           if (-not $name) { continue }
           $r = $el.Current.BoundingRectangle
           $yStr = '-1'
           if (-not [double]::IsNaN($r.Y) -and -not [double]::IsInfinity($r.Y)) { $yStr = [string][int]$r.Y }
           $lines.Add("$yStr`t" + ($name -replace "`r?`n", '⏎'))
       } catch { continue }   # 单元素异常只跳过，不中断整段读取
   }
   ```
   滚动（滚轮/PageDown/End）只作为增强手段：实测 Chromium UIA 不滚动也常能暴露完整回复文本树；滚轮 `mouse_event` 的 Add-Type 失败属正常降级，不能因此中断读取；滚动动作同样各自包 try/catch。
   **v3.9（B1）**：转储文件写盘后先校验是否含「PPT类型判定结果」等框架锚点（**匹配前先去空白**，UIA 会在 ASCII 词两侧插空格）——**通过才最小化收尾；失败保持窗口展开，执行「仅重抓」**（见 5.4 备注：只读取不发送，约 30 秒），成功后再最小化，禁止为补结果而重新上传重发。
7. 结果输出：先转述豆包的【PPT类型判定结果】与对应框架分析内容；**汇报中每条内容必须附来源页码标注**（豆包按指令输出〔第X页〕/〔第X–Y页〕，如实转述；豆包标〔页码不明〕或未标注时如实说明，禁止编造页码；确需按 PPT 页序与版式推断时注明「推断」）；**汇报层级标题一律用中文序号（一、二、三…），不使用 ABC/字母/长串编号**；**排版细则：①「分章节内容拆解」必须逐页呈现——按页码 1→N 每页独立一行/一段（推荐表格：页码｜板块名｜核心内容，一页一行），严禁把多页内容混成一段；②「核心重点汇总」「结构评价」等板块的内部条目一律按页码从小到大排序、每条独立成行（每条以〔第X页〕开头另起一行，可用项目符号列表；跨页范围按起始页参与排序），同页多条按原顺序依次排列，禁止用分号/顿号把多条内容挤在同一段**；通用型汇报中的「核心重点汇总」必须保留条目级细节（核心主张/关键事实与数据逐条/金句原文/行动与占位信息，每条带〔页码〕），不得概括成空话；遵守步骤6 规则（原文与解读区分、可点击链接仅用于源文件本身）；收尾清理按 0.5 规则2（本模式几乎不产生临时产物，无渲染目录残留）。

**边界**：本模式以 PPT 文字内容分析为主；用户一旦问到"某页的图片/照片/配图/图标是什么"→ 立即转步骤0 整页渲染识别模式（可配合 media 原图追问细节）。

**固定指令全文（v3.7 版，必须完整原样发送）**：

```text
请识别我上传的 PPT 文件，先判定其文本类型：优先从「学术科研型、工作汇报型、方案提案型、产品宣讲型」中选择，若均不匹配则归为「通用型」。判定完成后，先另起一行输出【PPT类型判定结果】，再严格按照对应类型的分析框架进行完整分析。

一、基础要求（所有类型都必须遵守）
1. 准确：完全忠于原文，不脑补、不引申、不添加 PPT 未提及的观点；区分原文事实与作者主观判断，关键数据、专有名词保留原文表述
2. 全面：覆盖全部核心板块，不遗漏核心论点、论据、数据、案例、结论与行动建议
3. 细致：梳理清内容逻辑链条，说明各板块的核心表意与论证关系，保留关键细节
4. 页码标注（硬性要求）：分析结果中的每一条内容（每个板块、每个要点、每项数据与结论）末尾都要标注其来源页码，格式统一为〔第X页〕；内容跨多页时标〔第X–Y页〕；无法对应具体页时标〔页码不明〕。页码以 PPT 实际页序为准，禁止编造页码。

二、类型分析框架（各类型独立成节，只用中文序号与行内要点，不使用字母或长串编号）
（一）学术科研型（适用：开题/答辩/论文汇报/课题成果）
分析要点：整体概况（研究主题、研究目的、汇报对象、全文逻辑脉络）；核心内容拆解（按PPT原有逻辑分模块梳理：研究背景与问题、研究方法与设计、核心结果与数据、主要结论）；关键亮点（研究创新点、核心发现、关键数据汇总）；不足与展望（原文提及的研究局限、后续计划）；逻辑评价（论证思路优势、信息缺口）。
（二）工作汇报型（适用：周报/月报/述职/项目复盘/部门汇报）
分析要点：整体概况（汇报主题、汇报主体、汇报周期、核心目标）；核心工作拆解（按模块/项目梳理已完成工作、关键动作）；成果与数据（核心业绩、量化成果、关键里程碑）；问题与不足（当前存在的问题、遇到的阻碍）；后续计划（下一步工作安排、行动要点、资源需求）。
（三）方案提案型（适用：项目立项/商业计划/解决方案/活动策划）
分析要点：方案概况（方案主题、目标受众、核心诉求、整体逻辑）；背景与痛点（方案针对的问题、现状与需求分析）；方案核心内容（整体思路、核心举措、落地路径、配套资源）；预期与风险（预期效果/收益、潜在风险与应对）；核心结论（方案核心主张与决策建议）。
（四）产品宣讲型（适用：产品发布/培训科普/品牌推介/方案宣讲）
分析要点：宣讲概况（主题、目标受众、核心目的）；核心内容拆解（按宣讲逻辑顺序梳理各章节核心信息、关键论据）；核心价值点（面向受众的核心收益、核心亮点）；逻辑脉络（宣讲的叙事思路与内容递进关系）。
（五）通用型（无法明确归类时使用）
分析要点：整体概况（PPT主题、核心主旨、内容逻辑框架）；分章节内容拆解（按PPT原有顺序逐部分梳理核心信息与关键要点，每部分标〔页码〕，便于按页定位）；结构评价（内容逻辑优势与信息缺口）。
其中「核心重点汇总」为必做项且最重要，必须按下述细则条目化输出，不得写成"涵盖……等内容"式的概括空话：
- 核心主张：整份 PPT 最想传达的一句话结论
- 关键事实与数据：具体经历、数字、日期、资格、量化指标等，逐条列出，禁止合并概括
- 标志性文案金句：口号、slogan、幽默梗等原文表述，逐句引用
- 行动号召与关键信息：投票/决策/参与方式、时间地点、联系方式等
- 待补充占位项：PPT 留空待填的信息点
以上每条均标注〔页码〕。

三、输出形式
1. 采用层级标题+分点式呈现；标题层级用中文序号（一、二、三…）或 PPT 原板块名，语言精炼，无冗余套话
2. 每条输出内容都必须带〔页码〕标注（规则见一.4），不允许出现不带页码的要点
3. 若某页只有标题无正文、或内容无法在页面上定位，如实标〔页码不明〕，不臆测
```

### 步骤1：定位并启动/复用豆包

#### 1.1 记录原窗口

启动前先记录当前前台窗口，便于结束后还原：

```powershell
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ForegroundWindowHelper {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
$originalHwnd = [ForegroundWindowHelper]::GetForegroundWindow()
$originalPid = 0
[ForegroundWindowHelper]::GetWindowThreadProcessId($originalHwnd, [ref]$originalPid) | Out-Null
# 结束时用 $originalPid 还原前台（AppActivate 需要进程 id，见步骤8）
```

#### 1.2 查找豆包是否已运行

```powershell
$doubao = Get-Process -Name Doubao -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } |
    Select-Object -First 1
```

> 注意：找到的豆包主窗口可能处于**最小化状态**（上次任务结尾按规则最小化保留）。最小化窗口仍有 `MainWindowHandle`，会被误判为“已就绪”——必须在步骤 1.5 用 `IsIconic` + 几何校验确认窗口真正恢复可见，不能只看句柄。

#### 1.3 启动豆包

如果没有可用的豆包主窗口，使用豆包主程序本体启动。**禁止通过应用商店入口或商店快捷方式启动，避免误开机械革命应用商店。**

```powershell
$doubaoExe = 'D:\AppStoreSoftstore\Install\doubao\Doubao.exe'
# 如果该路径不存在，先查找：
# Get-ChildItem 'D:\AppStoreSoftstore\Install\doubao' -Recurse -Filter Doubao.exe
Start-Process -FilePath $doubaoExe -ArgumentList '--force-renderer-accessibility'
```

`--force-renderer-accessibility` 用于让豆包窗口暴露 UI Automation 元素，便于后续自动聚焦输入框和读取结果。

#### 1.4 等待豆包主窗口出现

```powershell
$deadline = (Get-Date).AddSeconds(30)
do {
    Start-Sleep -Milliseconds 500
    $doubao = Get-Process -Name Doubao -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } |
        Select-Object -First 1
} until ($doubao -or (Get-Date) -gt $deadline)

if (-not $doubao) {
    Write-Error '豆包客户端启动失败，请检查是否已安装并登录豆包桌面版'
    return
}
```

#### 1.5 恢复并激活豆包窗口（v3.1.1：快速恢复 → 失败自动指令启动兜底）

**本代码块是后续所有窗口操作的前置闸门**。每次执行依赖快捷键/粘贴/坐标点击的步骤前，都必须先运行本块并确认输出 `DOUBAO_READY`。执行策略：路径 A 快速恢复已运行窗口（保留会话与已上传图片）→ 失败自动走路径 B `Start-Process` 指令启动 → 再失败 `AppActivate` 兜底 → 全部失败才报错退出。复用已运行的豆包时（上次任务结尾窗口被最小化保留），**绝不能只靠 `AppActivate`**——实测它对最小化窗口可能无效（v3.1 修复的故障根因）。

```powershell
# ========== 窗口恢复激活代码块（v3.1.1，完整复制执行） ==========
# 策略：快速路径（恢复已运行窗口，秒级、保留会话）→ 失败自动用 Start-Process 指令打开兜底。
# 成功输出 DOUBAO_READY 才允许继续；两条路径都失败才报错退出。
Add-Type -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
public struct RECT { public int Left, Top, Right, Bottom; }
'@ -Name Win32WinState -Namespace Win32

# 三重校验：① IsIconic 最小化检测 → SW_RESTORE；② GetWindowRect 几何校验（回到屏幕内）；
# ③ SetForegroundWindow 置前台并用 GetForegroundWindow 确认（Windows 前台锁 → 循环重试）
function Assert-DoubaoWindowReady {
    param($p)
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        if ([Win32.Win32WinState]::IsIconic($p.MainWindowHandle)) {
            [Win32.Win32WinState]::ShowWindow($p.MainWindowHandle, 9) | Out-Null  # SW_RESTORE
            Start-Sleep -Milliseconds 800
        }
        $r = New-Object Win32.Win32WinState+RECT
        [Win32.Win32WinState]::GetWindowRect($p.MainWindowHandle, [ref]$r) | Out-Null
        $w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
        # 最小化/隐藏窗口 rect 落在屏幕外（-32000/-21333 级）且宽高极小
        $onscreen = ($w -gt 200 -and $h -gt 200 -and $r.Left -ge -10 -and $r.Top -ge -10)
        if (-not $onscreen) {
            [Win32.Win32WinState]::ShowWindow($p.MainWindowHandle, 9) | Out-Null
            Start-Sleep -Milliseconds 1000
            [Win32.Win32WinState]::GetWindowRect($p.MainWindowHandle, [ref]$r) | Out-Null
            $w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
            $onscreen = ($w -gt 200 -and $h -gt 200 -and $r.Left -ge -10 -and $r.Top -ge -10)
        }
        if (-not $onscreen) { Write-Output "恢复尝试 ${attempt}/3：窗口仍未回到屏幕内"; continue }
        for ($fg = 1; $fg -le 5; $fg++) {
            [Win32.Win32WinState]::SetForegroundWindow($p.MainWindowHandle) | Out-Null
            Start-Sleep -Milliseconds 400
            if ([Win32.Win32WinState]::GetForegroundWindow() -eq $p.MainWindowHandle) { return $true }
        }
        Write-Output "恢复尝试 ${attempt}/3：未获得前台焦点"
    }
    return $false
}

# ---------- 路径 A（快速路径）：进程已在运行 → 直接恢复激活（保留会话与已上传图片） ----------
$doubaoReady = $false
$doubao = Get-Process -Name Doubao -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
if ($doubao) {
    $doubaoReady = Assert-DoubaoWindowReady $doubao
} else {
    Write-Output '豆包未在运行，直接走路径 B 启动'
}

# ---------- 路径 B（指令启动兜底）：恢复失败或进程不存在 → Start-Process 打开 ----------
# 说明：进程已运行时重复 Start-Process 会触发单实例激活（多数应用会顺带恢复窗口）；
# 进程未运行时则冷启动新实例。启动后统一重新走三重校验。
if (-not $doubaoReady) {
    Write-Output '快速恢复失败，改用 Start-Process 指令打开豆包'
    $doubaoExe = 'D:\AppStoreSoftstore\Install\doubao\Doubao.exe'
    if (-not (Test-Path -LiteralPath $doubaoExe)) {
        $cand = Get-ChildItem 'D:\AppStoreSoftstore\Install\doubao' -Recurse -Filter Doubao.exe -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($cand) { $doubaoExe = $cand.FullName }
    }
    if ($doubaoExe) {
        Start-Process -FilePath $doubaoExe -ArgumentList '--force-renderer-accessibility'
        # 等待主窗口出现（冷启动可能需要数秒~数十秒）
        $deadline = (Get-Date).AddSeconds(30)
        do {
            Start-Sleep -Milliseconds 500
            $doubao = Get-Process -Name Doubao -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
        } until ($doubao -or (Get-Date) -gt $deadline)
        if ($doubao) { $doubaoReady = Assert-DoubaoWindowReady $doubao }
        else { Write-Error '指令启动后仍未出现豆包主窗口'; exit 1 }
    } else {
        Write-Error '未找到 Doubao.exe，请检查安装路径'; exit 1
    }
}

# ---------- 最后兜底：AppActivate（部分场景比 SetForegroundWindow 有效） ----------
if (-not $doubaoReady) {
    Add-Type -AssemblyName Microsoft.VisualBasic
    [Microsoft.VisualBasic.Interaction]::AppActivate($doubao.Id) | Out-Null
    Start-Sleep -Milliseconds 1000
    $doubaoReady = ([Win32.Win32WinState]::GetForegroundWindow() -eq $doubao.MainWindowHandle)
}
if (-not $doubaoReady) { Write-Error '豆包窗口仍无法就绪，请检查窗口状态后重试'; exit 1 }
'DOUBAO_READY'
```

> 实测要点：
> - 窗口最小化时 UIA `BoundingRectangle` 会返回无效离屏坐标，**必须先恢复再取坐标**；
> - `mouse_event` 只按当前光标位置点击、不会移动光标，模拟点击前必须用 `[System.Windows.Forms.Cursor]::Position` 先把光标移过去（见步骤2备用方式）；
> - `IsIconic` 返回 $true 表示窗口最小化，即使 `MainWindowHandle` 非 0 也不可当作已就绪；
> - 路径 B（Start-Process）**只作为快速恢复失败的兜底**，不默认每次执行：已运行实例重复启动仅触发激活、不保证恢复最小化窗口，冷启动耗时也更长；启动后必须重新走三重校验（新进程句柄可能变化）。

#### 1.6 通用窗口定位函数（v3.10 B4 新增：按进程 ID，去标题依赖）

**用途**：所有"重新获取豆包窗口"的代码统一调用本函数，不再按窗口标题精确匹配。豆包标题会在任务中变化（「豆包」→「图片识别任务说明 - 豆包」等），按标题找会短暂"找不到窗口"而误触发恢复逻辑（v3.1 轮询窗口保护的抖动来源）。

```powershell
# B4（v3.10）：返回豆包进程中面积最大的顶层窗口元素；找不到返回 $null。
# 标题只用于日志，不参与定位；窗口最小化/离屏时返回 $null（配合 1.5 恢复逻辑）。
function Get-DoubaoWinByProcess {
    param($p)   # $p = Get-Process -Name Doubao | ...（MainWindowHandle -ne 0）
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $p.Id)
    $wins = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
    $best = $null; $bestArea = 0
    foreach ($w in $wins) {
        try { $r = $w.Current.BoundingRectangle } catch { continue }
        if ($r.Width -gt 200 -and $r.Height -gt 200) {
            $area = $r.Width * $r.Height
            if ($area -gt $bestArea) { $best = $w; $bestArea = $area }
        }
    }
    return $best
}
# 调用示例：$win = Get-DoubaoWinByProcess $doubao
```

> 轮询（5.1）常以独立脚本进程运行：需连同本函数与 `$doubao = Get-Process ...` 一并复制到该脚本，函数内不依赖任何外部状态。
> 新对话生效校验仍可用 `MainWindowTitle`（它只作"标题是否回到「豆包」"的判断，不用于窗口定位）。

### 步骤2：确保豆包处于新对话并聚焦输入框

每次上传图片前必须保证豆包处于**新对话**，避免把图片发到旧会话里。

如果豆包是刚启动的，默认通常就是新对话；如果复用了已运行的豆包，必须先新建对话。

**先决条件：执行本步骤前，必须先运行步骤 1.5 的恢复激活代码块并确认输出 `DOUBAO_READY`**（窗口最小化/不在前台时，快捷键会落到别的窗口，“新对话”必失败——这是 v3.1 修复的实测故障根因）。

**v3.9（B2）早失败探测**：`DOUBAO_READY` 后立即探测输入框（ProseMirror Edit，最多重试 10 秒）；探测不到 → **立即中止**并提示「豆包未登录或界面异常，请人工检查后重试」，不要进入上传与轮询空等（避免空等至超时才报错）。

用户已确认搜狗输入法的 `Ctrl+Shift+K` 冲突已解除，因此恢复使用豆包“新对话”快捷键：

```powershell
# 新建对话快捷键：Ctrl+Shift+K
# （执行前确认已运行过 1.5 恢复激活代码块且 $doubao / Win32.Win32WinState 已定义）
Add-Type -AssemblyName System.Windows.Forms
if (-not ([Win32.Win32WinState]::GetForegroundWindow() -eq $doubao.MainWindowHandle)) {
    # 前台已丢失：先完整重跑步骤 1.5 恢复激活代码块（含 Add-Type、$doubao、$doubaoReady），
    # 确认输出 DOUBAO_READY 后再继续；禁止在前台未确认时直接 SendKeys。
}
[System.Windows.Forms.SendKeys]::SendWait('^+k')
Start-Sleep -Milliseconds 1200
# 生效校验：成功后窗口标题应变回「豆包」（新对话尚无标题）；
# 若标题仍是旧会话名（如「图片识别任务说明 - 豆包」），说明快捷键未生效 → 转下方备用方式
$doubao2 = Get-Process -Name Doubao -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
if ($doubao2.MainWindowTitle -notmatch '^豆包$') {
    Write-Output '快捷键未生效（标题未变），转入 UIA 备用方式'
}
```

如果快捷键无效，再通过 UI Automation 点击侧边栏“新对话”。**点击前先确认窗口已恢复（1.5 代码块已通过）**，否则 `BoundingRectangle` 是无效离屏坐标：

```powershell
# UI Automation 备用方式（v3.1：先移动光标再点击，带生效校验）
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);' -Name U32Mouse -Namespace Win32

# 前置：本段执行前先运行步骤 1.5 恢复激活代码块（窗口必须已恢复到屏幕内）

# B4（v3.10）：按进程 ID 定位窗口，标题变化不影响
$win = Get-DoubaoWinByProcess $doubao
if (-not $win) { Write-Output '窗口定位失败，先执行步骤 1.5 恢复激活代码块再重试'; exit 1 }
$all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)
$newChat = $all | Where-Object { $_.Current.Name -eq '新对话' } | Select-Object -First 1

if ($newChat) {
    $clicked = $false
    try {
        $invoke = $newChat.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        $invoke.Invoke(); $clicked = $true
    } catch {
        # Chromium 内“新对话”常以 Text 暴露、无 InvokePattern → 用真实坐标点击
        $rect = $newChat.Current.BoundingRectangle
        if ($rect.Width -gt 10 -and $rect.Height -gt 10 -and $rect.X -gt 0 -and $rect.Y -gt 0) {
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(
                [int]($rect.X + $rect.Width / 2), [int]($rect.Y + $rect.Height / 2))
            Start-Sleep -Milliseconds 300
            [Win32.U32Mouse]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero) # LEFTDOWN
            Start-Sleep -Milliseconds 100
            [Win32.U32Mouse]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero) # LEFTUP
            $clicked = $true
        } else {
            Write-Output 'BoundingRectangle 无效（窗口可能未恢复），先执行步骤 1.5 再重试'
        }
    }
    Start-Sleep -Milliseconds 2000
    # 生效校验：窗口标题应变回「豆包」（新对话无标题）
    $doubao2 = Get-Process -Name Doubao -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
    if ($doubao2.MainWindowTitle -match '^豆包$') { 'NEW_CHAT_OK' }
    elseif ($clicked) { 'CLICKED_TITLE_UNCHANGED_NEED_VERIFY' }
} else {
    Write-Output '未找到“新对话”入口'
}
```

> 保持“每次识别都使用新对话”是刻意设计，不是随机行为。这样每次图片识别任务都从干净上下文开始，不会受旧对话影响。

豆包输入框是 `ProseMirror` 编辑器，需要通过 UI Automation 定位并聚焦：

```powershell
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

# B4（v3.10）：按进程 ID 定位窗口，标题变化不影响
$win = Get-DoubaoWinByProcess $doubao
if (-not $win) { Write-Error '窗口定位失败，先执行步骤 1.5 恢复激活代码块'; return }

$all = $win.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)

$edit = $all |
    Where-Object {
        $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::Edit -and
        $_.Current.ClassName -like '*ProseMirror*'
    } |
    Select-Object -First 1

if (-not $edit) {
    Write-Error '未找到豆包输入框，请确认豆包已打开且 accessibility 已启用'
    return
}

$edit.SetFocus()
Start-Sleep -Milliseconds 300
```

### 步骤3：上传图片

#### 3.1 单张图片

使用剪贴板粘贴最快：

```powershell
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$imagePath = 'D:\CSY\example.jpg'
$img = [System.Drawing.Image]::FromFile($imagePath)
[System.Windows.Forms.Clipboard]::SetImage($img)
$img.Dispose()

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::SendWait('^v')
Start-Sleep -Milliseconds 1200
```

#### 3.2 多张图片（单批 ≤10 张，v3.3）

按文件名自然排序后，逐张粘贴到同一个输入框，等待图片缩略图出现。**单批（同一次发送）最多 10 张**——总数 >10 时本批只粘贴前 10 张，剩余图片走步骤 3.4 分批规则（每批独立新对话）：

```powershell
$files = Get-ChildItem -LiteralPath 'D:\CSY' -File |
    Where-Object { $_.Extension -in '.jpg','.jpeg','.png','.bmp','.webp' } |
    Sort-Object Name

# ⚠️ 单批上限 10 张（v3.3）：只取前 10 张作为本批；$files.Count -gt 10 时，
#    剩余图片必须在后续批次中处理（每批独立新对话），禁止一次粘贴超过 10 张。
$batch = @($files | Select-Object -First 10)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# A1（v3.9）：连续粘贴提速——粘贴间隔由 800ms 缩到 450ms；
# 不再每贴一张就全窗口扫描校验附件，整批贴完后由 3.3 统一校验一次；
# 校验不足时按“尾部缺失补贴一次”补齐，仍不齐才中止重开新对话。
foreach ($file in $batch) {
    $edit.SetFocus()
    Start-Sleep -Milliseconds 150

    $img = [System.Drawing.Image]::FromFile($file.FullName)
    [System.Windows.Forms.Clipboard]::SetImage($img)
    $img.Dispose()

    [System.Windows.Forms.SendKeys]::SendWait('^v')
    Start-Sleep -Milliseconds 450
}
"本批粘贴 $($batch.Count) 张（上限 10 张）；剩余 $([math]::Max(0, $files.Count - 10)) 张待后续批次"
```

> 如果当前豆包版本支持一次多选文件上传，也可通过“附件按钮 → 上传文件或图片 → 多选文件 → 打开”完成，能进一步提速；同样受单批 10 张上限约束（一次最多选 10 张）。若该方式不稳定，则使用上述逐张粘贴。
>
> **v3.9（A1）**：整批粘贴完毕后统一做一次附件数校验（3.3），不再逐张扫描；校验不足时按“尾部缺失补贴一次”补齐（取本批末尾缺失数量的文件重贴），补齐后仍不足 → 中止并重新执行步骤 2，禁止带病发送。

#### 3.3 核对图片数量与会话纯净性（v3.1）

上传完成后做两项校验，**任一不通过立即中止并重新执行步骤 2（新对话）**，避免带病继续：

1. **附件数校验**：输入区已就位的图片数必须等于本次本批应上传图片数（≤10）；不足时按**“尾部缺失补贴一次”**补齐（从本批文件名列表末尾取缺失数量的文件重新粘贴），补齐后仍不足 → 中止并重新执行步骤 2；
2. **会话纯净校验**：当前对话主区域不得出现本任务之前的旧消息/旧图片痕迹（如旧的结构化回复行、旧的「图N」编号、旧会话标题仍显示在标题栏等）——存在即说明“新对话”未真正生效，继续发送会把本次图片混进旧会话（v3.1 实测故障）。

```powershell
# 通过 UI Automation 检查输入区的图片附件数量（窗口级扫描，兼容不同布局）
$all = $win.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)
$imageCount = @($all | Where-Object {
    $_.Current.ClassName -like '*image-container*'
}).Count

if ($imageCount -lt $batch.Count) {
    # A1（v3.9）：尾部缺失补贴——缺失数 = $batch.Count - $imageCount，
    # 取本批末尾对应数量的文件重新粘贴一次（SetImage + ^v，间隔 450ms），再复检一次
    $miss = $batch.Count - $imageCount
    $refill = @($batch | Select-Object -Last $miss)
    foreach ($file in $refill) {
        $img = [System.Drawing.Image]::FromFile($file.FullName)
        [System.Windows.Forms.Clipboard]::SetImage($img)
        $img.Dispose()
        [System.Windows.Forms.SendKeys]::SendWait('^v')
        Start-Sleep -Milliseconds 450
    }
    # 复检仍不足 → 中止，重新执行步骤 2
}
# 纯净性检查：主区域是否残留旧的结构化回复行（“字段名：实际内容”形态，提示词模板不会被命中）
$oldReply = @($all | Where-Object {
    $_.Current.Name -match '^(图片类型|核心主体|空间布局|场景与背景|画面文字|细节特征|特殊元素|不确定内容说明)：[^（(]'
})
if ($oldReply.Count -gt 0) {
    Write-Output "检测到 $($oldReply.Count) 行旧会话回复，新对话未生效 → 中止，重新执行步骤 2"
}
```

#### 3.4 单批 10 张上限与分批规则（v3.3）

**硬性上限：同一对话、同一次发送的图片附件数 ≤ 10 张**（豆包单条消息附件上限；实测一次粘贴 13 张只有 10 张入框，超出部分不生成附件）。

待识别图片总数 >10 时的分批执行规则：

1. 图片按文件名自然排序后切批：第 1 批 = 前 10 张，第 2 批 = 次 10 张……（PPT 场景按页码顺序切批，每批 ≤10 页）。
2. **每批独立走完整识别链路，批与批之间必须新建对话**（执行步骤 2 的 `Ctrl+Shift+K`）：
   `新对话 → 上传本批（≤10 张）→ 步骤 3.3 校验（附件数 == 本批张数、无旧回复痕迹）→ 步骤 4 发送固定指令 → 步骤 5 轮询读取 → 记录本批结果`
   - 原因：旧批的回复行（每图 ≥8 字段行）与结束标记「识别完毕」若仍留在同一窗口，会命中步骤 5.1 的完成判据（≥8 行/标记出现），导致下一批刚发送即误判完成、结果被污染；新对话保证 5.1/5.2 统计的只可能是本批回复。
3. 每批读取结果后立即记录/落盘（建议写入临时汇总文件，任务结束清理）；全部批次完成后按编号顺序合并为最终结果输出。
4. 单批内附件数不足（3.3 校验失败）→ 补齐缺失图片后重试一次；仍失败则记录失败项，不阻塞本批发送。
5. 步骤 3.3 附件数校验中的“本次图片数”取**本批**张数（≤10），不是任务总数。

### 步骤4：发送固定识别指令

#### 4.0 输入法与搜狗软键盘防护（重要）

- **禁止使用 `SendKeys` 直接键入中文或长文本**，否则可能触发搜狗输入法/软键盘并产生乱码。
- 所有文字内容一律先放入剪贴板，再通过 `Ctrl+V` 粘贴。
- 如果搜狗软键盘或候选框弹出，先按 `Esc` 关闭候选/软键盘，再继续操作；不要在有中文候选词时按 Enter。
- 发送消息时优先点击豆包的“发送”按钮；找不到发送按钮时才使用 Enter。

将下方固定指令复制到剪贴板并粘贴到输入框，然后发送：

```powershell
$fixedPrompt = @'
你负责识别图片，仅输出纯识别结果，供下游AI系统调用处理。严格遵守以下所有规则，禁止输出任何规则解释、寒暄、确认语、思考过程或与识别内容无关的文字。

【核心强制规则】
1. 零臆测原则：仅描述图片中视觉可见的客观内容，禁止推测、脑补、解读、美化、评价，禁止补充画面外信息。
2. 全要素覆盖：必须完整识别主体、场景、文字、细节、特殊元素，不得遗漏可清晰辨识的关键信息。
3. 模糊必标原则：
   - 单个文字/细节无法辨认：用「[?]」标注
   - 局部区域内容模糊、遮挡、虚化：标注「[该区域内容模糊，无法准确识别]」
   - 整段/整块内容完全无法辨识：标注「[内容模糊，无法准确识别]」
   - 部分可见的文字仅保留可辨识部分，不可辨识字符用「[?]」补位；禁止对模糊内容进行任何猜测性描述。
4. 纯结果输出：禁止出现“我识别到”“这是一张”“图片中”等冗余前缀，直接按格式填充内容。

【输出格式与识别维度】
严格按以下字段输出，无内容则填「无」：
- 图片类型：（如实拍照片、截图、插画、海报、证件、图表、漫画等）
- 核心主体：（画面最主要的1-3个对象，含类别、数量、大致形态；场景复杂时优先提取占比最大、视觉最突出的对象）
- 空间布局：（主体间的相对位置、前后/左右/上下关系、各自主占画面比例）
- 场景与背景：（整体环境类型、背景元素、光线、色调）
- 画面文字：（逐字转录所有清晰可见文字，按从左到右、从上到下的阅读顺序排列；标注文字所在位置，如“左上角”“画面中部”；无文字填「无」）
- 细节特征：（物体颜色、材质、姿态、动作、服饰等可辨识的精细属性）
- 特殊元素：（二维码、条形码、logo、图标、符号、印章、图表等标志性元素，仅描述外观，不解读含义）
- 不确定内容说明：（列出所有模糊、遮挡、反光、低分辨率导致无法识别的区域/内容；无不确定填「无」）

【多图与异常处理】
1. 单张图片直接按上述格式输出，无需编号。
2. 多张图片按上传顺序依次编号为「图1」「图2」「图3」……，每张分别独立按上述格式输出。
3. 若图片包含违规敏感内容、或完全无法识别（全黑/全白/纯噪点），仅输出：无法识别该图片

【结束标记】
1. 上述所有图片的识别结果输出完毕后，在回复末尾另起一行，输出由四个汉字连续书写（紧挨连写、不加空格、不加任何标点）组成的结束标记；四个汉字依次为：第一个字「识」，第二个字「别」，第三个字「完」，第四个字「毕」。
2. 若输出了「无法识别该图片」，在该行之后同样另起一行输出上述结束标记。
3. 该结束标记是输出格式的组成部分，不属于规则所禁止的多余文字；下游系统以此判断识别流程已结束。
'@

[System.Windows.Forms.Clipboard]::SetText($fixedPrompt)
$edit.SetFocus()
Start-Sleep -Milliseconds 200

# 粘贴文本：只使用 Ctrl+V，不逐字输入
[System.Windows.Forms.SendKeys]::SendWait('^v')
Start-Sleep -Milliseconds 500

# 发送方式一：优先查找“发送”按钮并 Invoke
$sendButton = $all |
    Where-Object {
        $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::Button -and
        $_.Current.IsEnabled -and
        $_.Current.Name -match '发送|Send'
    } |
    Select-Object -First 1

if ($sendButton) {
    try {
        $invoke = $sendButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        $invoke.Invoke()
        Write-Output '已通过发送按钮发送'
    } catch {
        [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    }
} else {
    # 发送方式二：兜底使用 Enter；先按 Esc 关闭可能的输入法候选/软键盘
    [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
    Start-Sleep -Milliseconds 100
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
}

# 发送成功后【不要最小化】：保持窗口展开，步骤5才能以高频轮询读取生成进度
Start-Sleep -Milliseconds 500
```

> 此指令必须完整发送，禁止删减或修改。发送后豆包窗口保持展开，进入步骤5高频轮询；读取到完整结果后再最小化（见步骤5.4）。

### 步骤5：等待识别完成并读取结果

#### 5.1 等待豆包响应

豆包开始处理后，窗口标题通常会变为与任务相关的标题，例如：

```text
图片识别任务说明 - 豆包
```

**豆包窗口在轮询期间保持展开（不最小化）**：最小化状态下 Chromium 窗口不向 UI Automation 暴露元素，轮询会读不到任何进度，只能反复“恢复-读取-再最小化”，既拖慢确认又频繁闪屏。保持窗口可见（不必始终前台），每次轮询即可直接读取生成状态，第一时间确认完成。

不要使用固定长等待。改为“智能停止检测”：

```powershell
# ========== 5.1 智能等待（v3.0：结束标记「识别完毕」即时判完成 + v2.9 字段判据兜底） ==========

# 回复字段行特征：“字段名：实际内容”（冒号后紧跟非括号字符）。
# 用户提示词模板均为“字段名：（示例…”，冒号后紧跟全角括号，不会被匹配，
# 因此该模式只命中 AI 回复，天然排除用户消息（v2.8 误判根因）。
$replyPattern = '^(图片类型|核心主体|空间布局|场景与背景|画面文字|细节特征|特殊元素|不确定内容说明)：[^（(]'
# 末字段特征：正文出现“不确定内容说明：非括号内容”即回复已写到结尾字段。
$lastFieldPattern = '不确定内容说明：[^（(]'

# 基线：发送后、豆包尚未回复时窗口文本总量（含用户提示词消息），用于兜底判断
$baseTotal = 0
$win0 = Get-DoubaoWinByProcess $doubao   # B4（v3.10）：按进程 ID 定位
if ($win0) {
    $all0 = $win0.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($e in $all0) { $baseTotal += $e.Current.Name.Length }
}

$deadline = (Get-Date).AddSeconds(150)
$lastReplyLen = 0
$lastWinTotal = $baseTotal
$stableSince = $null
$markerSince = $null   # 结束标记出现时刻（v3.0 主判据）
$done = $false
$missCount = 0         # 窗口连续丢失计数（v3.1 轮询窗口保护）

do {
    Start-Sleep -Milliseconds 500

    # 每次轮询都重新获取窗口和输出元素（B4 v3.10：按进程 ID 定位，标题变化不影响）
    $win = Get-DoubaoWinByProcess $doubao
    if (-not $win) {
        # 窗口消失（可能被意外最小化/遮挡，UIA 读不到元素）：自包含恢复后再继续，避免空等（v3.1）
        $missCount++
        if ($missCount -ge 4) {
            Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name Win32Restore -Namespace Win32 -ErrorAction SilentlyContinue
            if ([Win32.Win32Restore]::IsIconic($doubao.MainWindowHandle)) {
                [Win32.Win32Restore]::ShowWindow($doubao.MainWindowHandle, 9) | Out-Null # SW_RESTORE
                Start-Sleep -Milliseconds 800
            }
            $missCount = 0
        }
        continue
    }
    $all = $win.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)

    # 1) 统计本轮窗口状态
    $replyLen = 0            # 回复字段行累计长度（流式输出时持续增长）
    $replyCount = 0          # 命中回复字段行的元素个数（完整回复 ≥ 8）
    $lastFieldSeen = $false  # 末字段「不确定内容说明：实际内容」是否已出现
    $doneMarkerSeen = $false # 结束标记「识别完毕」是否已出现（v3.0 主判据）
    $stopExists = $false
    $suggestionExists = $false
    $winTotal = 0
    foreach ($element in $all) {
        $name = $element.Current.Name
        if (-not $name) { continue }
        $winTotal += $name.Length
        if ($name -match $replyPattern) { $replyLen += $name.Length; $replyCount++ }
        if ($name -match $lastFieldPattern) { $lastFieldSeen = $true }
        # 结束标记检测：容忍“识、别、完、毕”“识别完毕。”等带标点变体；
        # 固定指令正文用拆字写法，提示词内不存在连续四字，命中必为豆包输出
        $clean = $name -replace '[、，,．.·\s]', ''
        if ($clean -match '识别完毕') { $doneMarkerSeen = $true }
        # 停止按钮：不限定控件类型（Chromium 内部按钮可能不以 Button 暴露）
        if ($name -match '停止生成|^停止$|Stop') { $stopExists = $true }
        if ($name -match '^图[0-9].*(是什么|有哪些|多少|什么颜色|什么字体)$' -or
            $name -match '核心主体是什么|细节特征是什么|特殊元素是什么') {
            $suggestionExists = $true
        }
    }

    # 2)【主判据】结束标记「识别完毕」已出现 = 豆包已输出全部结果
    #    （标记是回复最后一行，出现后再等 1.5 秒确保文本渲染完整即完成）
    if ($doneMarkerSeen) {
        if ($null -eq $markerSince) { $markerSince = Get-Date }
        elseif (((Get-Date) - $markerSince).TotalSeconds -ge 1.5) {
            $done = $true; break
        }
        continue
    }

    # 3)【兜底判据】（豆包未按指令输出结束标记时启用）
    #    出现追问建议按钮时，通常代表回答已结束
    if ($suggestionExists) { $done = $true; break }

    # 4) 仍显示“停止”按钮 → 生成中，重置稳定计时（仅作防误判，不作完成前提）
    if ($stopExists) {
        $stableSince = $null
        continue
    }

    # 5) 回复文本不再增长 → 累计稳定时长
    if ($replyLen -eq $lastReplyLen) {
        if ($null -eq $stableSince) { $stableSince = Get-Date }
        elseif (((Get-Date) - $stableSince).TotalSeconds -ge 3) {
            # 兜底完成判据（任一满足即完成）：
            #   a) 末字段「不确定内容说明：实际内容」已输出（回复已写到结尾字段）；
            #   b) 完整字段行已出现（≥8 行，一图 8 字段；多图只会更多）；
            #   c) 窗口总文本明显超出基线（回复确实产生）且已整体稳定
            if ($lastFieldSeen -or $replyCount -ge 8) { $done = $true; break }
            if ($winTotal -gt ($baseTotal + 200) -and $winTotal -eq $lastWinTotal) {
                $done = $true; break
            }
        }
    } else {
        # 回复仍在增长：更新基准并重置稳定计时
        $lastReplyLen = $replyLen
        $lastWinTotal = $winTotal
        $stableSince = Get-Date
    }

} until ((Get-Date) -gt $deadline)

if (-not $done) { Write-Output '轮询超时，请检查豆包是否仍在生成' }
```

> 说明：  
> - **主判据（v3.0）**：固定指令要求豆包在全部结果后另起一行输出结束标记「识别完毕」（指令正文用拆字写法，提示词内不存在连续四字），因此窗口内任何元素出现"识别完毕"（容忍"识、别、完、毕"等标点变体）必然来自豆包回复 → 标记出现后等 1.5 秒渲染稳定即完成，通常比 v2.9 的 3 秒文本稳定更快结束。  
> - **轮询窗口保护（v3.1）**：轮询循环内置窗口丢失保护——若连续找不到窗口（被最小化/遮挡），自动 `SW_RESTORE` 恢复后继续轮询，防止空等到超时；恢复逻辑自包含（不依赖 1.5 定义的 Add-Type，轮询常以独立脚本进程运行）。
> - **兜底判据（v2.9）**：若豆包未按指令输出结束标记（模型偶尔不听话），仍按末字段出现 / 完整字段行（≥8 行）/ 追问建议 / 文本超基线稳定等方式判完成，不会卡死。  
> - 完成判定只看回复内容本身，不依赖"停止"按钮是否暴露。  
> - 仅在异常情况下才等待到 150 秒上限。

#### 5.2 读取结果

通过 UI Automation 读取主内容区的 `ListItem` 或 `Text`。**只收集「字段名：实际内容」形态的回复行**（与 5.1 同一正则），用户提示词模板“字段名：（示例…”不会被误读进来；豆包输出的结束标记「识别完毕」不以字段名开头，同样不会被收入结果：

```powershell
$replyPattern = '^(图片类型|核心主体|空间布局|场景与背景|画面文字|细节特征|特殊元素|不确定内容说明)：[^（(]'
$resultItems = @()
foreach ($element in $all) {
    if ($element.Current.ControlType -eq [System.Windows.Automation.ControlType]::ListItem) {
        $name = $element.Current.Name
        if ($name -match $replyPattern) { $resultItems += $name }
    }
}
# 若 ListItem 取不到（个别版本回复整体在一个 Text 元素内），兜底改为扫描全部元素：
if ($resultItems.Count -eq 0) {
    foreach ($element in $all) {
        $name = $element.Current.Name
        if ($name -match $replyPattern) { $resultItems += $name }
    }
}
```

#### 5.3 格式校验

- 将全部回复行拼接后，8 个字段名（图片类型、核心主体、空间布局、场景与背景、画面文字、细节特征、特殊元素、不确定内容说明）各出现次数应一致：
  - 单图：每个字段名恰好出现 1 次；
  - 多图：每个字段名出现 N 次（N = 图片数）。豆包可能不带「图N」标题（实测为 8 字段/组连续排列），按字段出现次数判定即可。
- 如果校验不通过（字段缺失或次数不一致），重新发送一次指令；仍失败则输出「图片识别结果格式异常，无法解析」。

#### 5.4 先落盘校验，通过后才最小化（v3.9 B1）

结果文本读取后，按以下顺序收尾（**先落盘 → 校验 → 通过才最小化**）：

1. **先落盘**：把读取到的回复转储写入本次临时目录（如 `*-results.txt` / `reply_dump.txt`），防止后续任何异常丢失结果；
2. **校验内容有效性**：图片模式——8 字段行数 == 本次图片数；PPT 文本分析模式——转储含「PPT类型判定结果」等框架锚点；**锚点匹配前先去除全部空白字符**（UIA 文本会在 ASCII 词两侧自动插空格，实测为「PPT 类型判定结果」，直接 Contains 会误判失败）；
3. **校验通过 → 立即将豆包窗口最小化**，保留当前对话与已上传图片，供后续追问复用：

```powershell
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name Win32ShowWindow -Namespace Win32
[Win32.Win32ShowWindow]::ShowWindow($doubao.MainWindowHandle, 6) # SW_MINIMIZE
```

> 最小化只发生在“结果已完整落盘并校验通过”之后；轮询等待阶段一律保持窗口展开。
>
> **v3.9（B1）「仅重抓」**：当转储/校验失败（脚本中断、内容截断等）但豆包回复其实已生成时，**不要重跑整个任务**（重新上传+重发 ≈ 3–5 分钟）；直接复用当前对话执行一次「恢复窗口 → 滚动到底 → 健壮转储」（用 0.6 步骤6 的逐元素容错代码，**只读取、不发送**，约 30 秒），成功后补校验、再最小化并进入收尾清理。

### 步骤6：结果解读与输出

不要直接原样输出豆包的全部原始内容。应根据用户问题或当前任务需要，输出自己的解读、归纳、建议。

格式区分：

- 每张图片介绍前，必须附上该图片的**可点击本地文件链接**，例如：
  ```markdown
  [C271F42C14CFC125F33AACC788D0C6ED.jpg](file:///D:/CSY/C271F42C14CFC125F33AACC788D0C6ED.jpg)
  ```
  如果链接无法点击，至少给出完整绝对路径：`D:\CSY\文件名.jpg`。
- **豆包识别原文/原字段**：引用或标注原始结构化内容。
- **智能体解读/补充/推断**：用自己的语言回答用户问题。
- 不得把解读内容伪装成豆包原文；推测内容需标注“推断”。

### 步骤7：追问图片细节（可选）

当以下情况出现时，不要直接结束任务，而是继续向豆包追问：

- 用户对识别结果进行追问，例如“图里某个字是什么”“这个按钮写的是什么”“能不能放大看细节”。
- 当前任务需要额外的视觉信息才能继续，例如需要判断颜色、材质、位置、文字完整内容、二维码内容等。
- 豆包首次结果中存在模糊、遮挡、不确定内容，需要进一步确认。

#### 7.1 生成追问文本命令

根据用户问题或任务需要，生成一条**合理、具体、只针对图片细节的文本命令**，例如：

```text
请针对刚才上传的图片继续回答：
- 画面中「不确定内容说明」提到的模糊区域，尝试再次辨认并描述可见内容。
- 请完整转录图中所有文字，不要省略。
- 请详细描述图中第X个物体的颜色、材质、位置、周围环境。
- 如果图中包含二维码/条形码，请描述其外观和可辨识内容。
- 请补充说明图片的上下左右边缘是否还有未识别内容。
- 回答完毕后，另起一行输出由四个汉字连写（不加标点）组成的结束标记：第一个字「识」，第二个字「别」，第三个字「完」，第四个字「毕」。
```

要求：

- 命令必须围绕图片本身，禁止询问与图片无关的信息。
- 若涉及多图，必须指明“图1/图2/图3……”。
- 命令应能直接发送给豆包，无需额外解释。
- **追问文本中禁止出现连续的“识别完毕”四字**（提示词会以用户消息形式出现在窗口，会导致 5.1 的结束标记检测误判）：需要提及结束标记时，一律使用上面的拆字写法（“第一个字「识」……”）；若未附带结束标记要求，豆包答完不会输出标记，轮询会自动走 5.1 的字段兜底判据，不影响完成检测。

#### 7.2 发送追问并读取结果

默认情况下，首次识别完成后豆包窗口只是被最小化，**对话和已上传图片仍然保留**。追问时：

1. 如果豆包窗口是最小化状态，先执行步骤 1.5 的恢复激活代码块（三重校验）并确认输出 `DOUBAO_READY` 后再继续；发送追问文本前同样须确认窗口在前台。
2. 直接聚焦原输入框。
3. 将针对“图1/图2/……”的追问文本放入剪贴板后 `Ctrl+V` 粘贴；**禁止用 SendKeys 逐字输入中文**。
4. 发送方式与步骤4相同：优先点击“发送”按钮；找不到发送按钮时先按 Esc 关闭输入法候选，再按 Enter。
5. 不需要重新开新对话，也不需要重新上传图片。
6. 等待与读取方式同步骤5（5.1 轮询）：若追问文本按 7.1 附带了结束标记要求，豆包答完会输出「识别完毕」，命中后约 1.5 秒即判完成；若未附带，自动走字段兜底判据（末字段/完整字段行/文本稳定），同样能正常结束。

只有极端情况下豆包窗口已被关闭，才需要从步骤1重新开始并重新上传图片。

#### 7.3 追问结果处理

- 将豆包返回的补充内容同样作为“豆包识别原文/原字段”。
- 结合首次识别结果和追问结果，输出更完整的智能体解读。
- 如果追问后仍无法识别，明确标注无法确认，不进行无依据猜测。

### 步骤8：最小化保留 + 条件关闭 + 窗口还原

首次识别完成并输出结果后，**不要关闭豆包窗口**，而是：

1. **先清理本次产物，再收尾（try/finally 语义，细则见步骤0.5）**：删除本次任务产生的临时文件——PPT 拆解产物目录（`D:\DS\.dsh\tmp_ppt_<8位hex>`）、media 提取物、中间结果文件、缓存、截图、日志、OCR 中间文件、辅助脚本等；不删除用户原始图片/PPT。删除前逐一核对路径确为本流程产物，遵守步骤0.5 防误删黑名单；任务失败或中断时同样必须执行本清理。
2. 确认豆包窗口已处于最小化状态并保留当前“新对话”和已上传的图片（步骤5.4 读取结果后已最小化；若被再次恢复则重新最小化），便于后续追问。
   ```powershell
   $sig = '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
   Add-Type -MemberDefinition $sig -Name Win32ShowWindow -Namespace Win32
   [Win32.Win32ShowWindow]::ShowWindow($doubao.MainWindowHandle, 6) # SW_MINIMIZE
   ```
3. 将前台还原到步骤1.1记录的原窗口：
   ```powershell
   Add-Type -AssemblyName Microsoft.VisualBasic
   [Microsoft.VisualBasic.Interaction]::AppActivate($originalPid) | Out-Null
   ```
4. 暂不关闭豆包窗口。只有出现以下任一情况时，才关闭豆包：
   - 用户已经明确表示不需要继续识别/追问；
   - 用户开始执行与当前图片无关的其他命令；
   - 当前工作任务所需图片信息已经完整，且后续不会再用到这些图片；
   - 用户主动要求关闭豆包；
   - 要开始另一批全新图片的识别任务，避免新旧图片混淆。
5. 如果误开了机械革命应用商店或其它无关窗口，尝试关闭：
   ```powershell
   Get-Process -Name PCAppStore -ErrorAction SilentlyContinue |
       Where-Object { $_.MainWindowHandle -ne 0 } |
       ForEach-Object { $_.CloseMainWindow() }
   ```
6. 当满足关闭条件时，再关闭豆包窗口并还原前台。

## 提速命令建议

实际执行时，可以把以下动作合并为“一条命令链”，减少模型多轮思考和工具往返：

```powershell
# 0) PPT/PPTX 先拆解：步骤0 整页渲染 PNG + media 原图提取（产物目录 D:\DS\.dsh\tmp_ppt_<8位hex>）
# 1) 恢复激活豆包窗口（步骤1.5 代码块，确认 DOUBAO_READY）+ 新建对话并校验（步骤2）
# 2) 粘贴所有图片（PPT 任务粘贴整页 PNG，每批 ≤10 张）
# 3) 粘贴固定识别指令
# 4) 回车发送
# 5) 轮询读取结果
# 6) 每批/全部完成后：try/finally 清理本次产物目录与辅助脚本（步骤0.5 规则2），再最小化豆包并还原原窗口
```

## 批量/多图处理规则

1. 无需逐张授权，一次任务可处理同一路径下的所有有效图片。
2. 图片按文件名自然排序（数字优先、字典序）依次上传和识别。
3. 批量结果按编号顺序汇总。
4. 单张图片识别失败不影响其他图片处理，失败项标注「[该图片识别失败]」。
5. **单批上限 10 张（v3.3）**：同一对话/同一次发送最多上传 10 张图片；总数 >10 时必须分批，每批 ≤10 张、每批独立新对话（步骤 3.4），全部批次完成后按编号顺序汇总输出。

## 异常处理矩阵

| 异常类型 | 具体场景 | 处理方式 |
|---|---|---|
| 文件异常 | 文件不存在、路径错误 | 告知用户路径无效，请核对后重试 |
| 文件异常 | 目标是 .ppt/.pptx | 先执行步骤0 拆解为整页 PNG（+media 原图）后再按图片流程识别；页数 >10 分批 |
| 文件异常 | 整页渲染失败（未安装 PowerPoint/WPS COM） | .pptx 退化为 0.3 解包提取 media 原图识别；.ppt 告知用户需安装 Office/WPS 后重试 |
| 文件异常 | 源文件在回收站/被占用（COM 打开 E_FAIL 或附件粘贴失败） | v3.9 B3：先复制到临时目录（`tmp_ppt_<hex>\src\`）再渲染/上传副本，副本随 0.5 规则2 清理，原文件不动；复制也失败则明确报错 |
| 文件异常 | 磁盘上发现历史遗留的 tmp_ppt_* 产物目录或辅助脚本 | 不做任务启动前的主动扫描；用户发现并指出时，按步骤0.5 规则3 核对命名约定与内容归属后再清理，命中防误删黑名单的一律保留；本次任务产物在收尾按规则2 清理 |
| 文件异常 | 收到相对路径或 `.` | 先转换为绝对路径；无法转换时请用户提供完整绝对路径 |
| 文件异常 | 格式不支持、文件损坏 | 告知用户不支持的格式/文件损坏，请更换文件 |
| 客户端异常 | 未安装豆包、未登录账号 | 提示安装并登录豆包桌面客户端后重试 |
| 客户端异常 | 启动超时、进程崩溃 | 自动重试1次，失败则提示客户端异常 |
| 客户端异常 | 豆包已打开但没有可操作输入框 | v3.9 B2 早失败探测：DOUBAO_READY 后 10 秒内探测不到输入框立即中止并提示「未登录或界面异常，请人工检查」；确为 accessibility 缺失再用 `--force-renderer-accessibility` 重启豆包 |
| 上传异常 | 文件过大、网络中断 | 自动重试1次，失败则提示上传失败 |
| 上传异常 | 图片粘贴后未出现附件 | 重新聚焦输入框并重试粘贴 |
| 上传异常 | 单次粘贴超过 10 张，超出部分未生成附件（实测 13 张仅入框 10 张） | 分批处理：每批 ≤10 张、每批独立新对话（见步骤 3.4），禁止单次粘贴 >10 张 |
| 识别异常 | 结果格式错误、内容为空 | 重发指令重试1次，失败则提示识别异常 |
| 识别异常 | 回复已生成但转储/校验失败（脚本中断、内容截断） | v3.9 B1「仅重抓」：不重传不重发，保持窗口展开，复用当前对话只读取重抓转储（约 30 秒）并补校验；成功后再最小化收尾 |
| 合规异常 | 图片含敏感违规内容 | 终止处理，仅输出「无法识别该图片」 |
| 权限异常 | 无文件读取权限 | 告知用户权限不足，调整权限后重试 |
| 窗口异常 | 误开应用商店/无关窗口 | 立即尝试关闭；若权限不足无法关闭，提示用户手动关闭 |
| 窗口异常 | 豆包进程在运行但窗口最小化/离屏（rect 为负值如 -21333），AppActivate 未生效 | 执行步骤 1.5 恢复激活代码块（ShowWindow SW_RESTORE + 几何校验 + SetForegroundWindow 前台校验循环）；仍失败重试 1 次后报错 |
| 窗口异常 | 快捷键/粘贴/点击未生效（窗口不在前台，操作落到其他窗口） | 任何 SendKeys、剪贴板粘贴、坐标点击前先执行 1.5 代码块并确认 DOUBAO_READY；坐标点击前用 Cursor.Position 移动光标；快捷键后校验标题变化 |
| 识别异常 | 结果混入本任务外的图片/旧会话内容（回复组数 > 本次图片数，或主区域有旧回复行） | “新对话”未生效，中止；重新执行步骤 2 新建对话并校验纯净后再重发指令 |
| 窗口异常 | 原窗口已关闭或无法还原 | 返回当前任务主界面，不阻塞结果输出 |
| 追问异常 | 追问后豆包无响应或结果为空 | 重发追问1次，仍失败则基于已有结果回答并标注未确认项 |

## 约束与注意事项

1. 用户主动触发或任务需要识别图片时，可直接调用豆包，不再额外询问授权。
2. 识别指令必须完整固定发送，不得删减规则、修改格式。
3. 豆包原始识别结果仅作为事实输入；允许智能体在此基础上解读和补充，但不得把解读内容伪装成豆包原文。
4. 解读补充必须与原文明确区分，且不得脱离识别事实进行无依据臆测。
5. 所有文件操作必须使用绝对路径，禁止将 `.` 或相对路径传给文件工具。
6. 每次识别及追问结束后，都必须清理本次产生的临时/垃圾文件（产物目录、media 提取物、中间结果、辅助脚本，见步骤0.5 规则2）；**清理必须在最终汇报之前完成，失败/中断路径同样要清理**；读取到完整结果后立即将豆包最小化并保留会话，不关闭窗口。
7. 首次上传图片前必须确保豆包处于“新对话”；后续追问必须复用同一对话，不重新开新对话、不重新上传图片。
8. 只有确认用户不再需要图片信息、已转入其他命令，或当前任务图片信息已完整时，才关闭豆包窗口。
9. 最终输出中每张图片必须附上可点击的本地文件链接：源图片直接给源路径；PPT 整页 PNG 等**临时产物**链接会随任务收尾清理失效，须在输出中注明「临时页图，任务结束已清理」；用户需要长期可点击的页图时，先把文件转存到用户指定位置（如原 PPT 同目录的 `_slides\`）再给链接，禁止为保链接而把产物留在临时目录。
10. 所有发送给豆包的文字必须通过剪贴板粘贴，禁止用 SendKeys 直接逐字输入中文，避免触发搜狗软键盘/乱码。
11. 追问文本必须针对图片细节生成，不得发散到与图片无关的内容。
12. 仅用于本地图片识别场景，不得超范围使用。
13. 轮询等待阶段保持豆包窗口展开（不最小化），以 0.5 秒间隔高频轮询，及时确认识别完成；只有完整读取结果并通过格式校验后才最小化窗口。
14. 所有依赖快捷键（Ctrl+Shift+K / Ctrl+V / Enter）、剪贴板粘贴或坐标点击的操作，执行前必须先运行步骤 1.5 恢复激活代码块并确认 `DOUBAO_READY`（窗口最小化或不在前台时操作会落空/错发）；窗口最小化时 UIA 坐标无效，须先恢复再取坐标；mouse_event 前必须先移动光标。
15. 识别前必须确认会话纯净：新对话后主区域无旧消息痕迹、附件数等于本次图片数；回复组数与本次图片数不一致时判定为混入旧图，中止重来，不得带病继续。
16. **单批图片附件数不得超过 10 张（豆包单条消息附件上限，v3.3）**；待识别图片 >10 张时按步骤 3.4 分批：每批 ≤10 张、每批独立新对话完成「上传→校验→发送→读取→记录」，禁止把超过 10 张的图片一次性粘贴进输入框。
17. PPT/PPTX 任务：先执行步骤0 拆解再走图片流程；默认只上传整页渲染 PNG、每批 ≤10 页；拆解产物目录与临时文件一样在任务结束时清理，不删除用户原始 PPT。
18. **防误删纪律（步骤0.5 规则3）**：清理只针对「命名约定内 + 内容核对无误」的本流程产物；`D:\DS\.dsh\skills\`、用户原始图片/PPT 及其所在目录、任务开始前已存在的历史文件、无法确认归属的文件一律不删；拿不准就保留并向用户报告。
19. **PPT 文本分析（步骤0.6）汇报必须带来源页码标注**：以豆包按指令输出的〔第X页〕/〔第X–Y页〕为准转述；豆包标〔页码不明〕或未标注时如实说明，禁止编造页码；确需按 PPT 页序与版式推断时注明「推断」。
20. **PPT 文本分析汇报格式**：层级标题只用中文序号（一、二、三…），不得使用 ABC/字母/长串编号；通用型「核心重点汇总」必须以条目级可引用信息呈现（核心主张、关键事实与数据逐条、金句原文逐句、行动号召与占位信息），每条带〔页码〕，禁止"涵盖……等内容"式概括空话。
21. **PPT 文本分析汇报排版细则**：①「分章节内容拆解」按页码 1→N **逐页呈现**，每页独立一行/一段（推荐表格：页码｜板块名｜核心内容），严禁多页混段；②「核心重点汇总」「结构评价」等板块内部条目**按页码从小到大排序、每条独立成行**，每条以〔第X页〕开头另起一行，同页多条按原顺序排列，禁止用分号/顿号把多条内容挤在同一段。

---
name: local-image-doubao-recognition
description: >-
  本地图片高精度结构化视觉识别（豆包桌面客户端调用）。用户主动要求识别本地图片/文件夹，或当前任务必须获取本地图片视觉信息时，
  无需授权询问，直接自动打开或复用豆包、上传图片、发送识别指令、读取结构化结果，并根据任务需求进行解读。支持 JPG/JPEG/PNG/BMP/WebP；
  支持 PPT/PPTX 双通道：只问文字内容（总结/主要讲了什么/大纲）→ 直传文件本体做文本分析（轻量模式，不拆解渲染）；问页面内的图片内容 → 整页渲染 PNG 后按图片流程识别。
  图片识别为两部分输出：第一部分客观描述（8 字段纯事实），第二部分推想预测（5 字段，含置信度/依据/其他可能）；推断内容按图片类型可编辑（人物、动植物、作品、地点、景物、物理现象、文字标识等），由可编辑槽位按本次任务生成。
  不处理在线图片链接、纯文件属性查询，以及上下文中已有该图片结构化结果的场景。
---

# 本地图片结构化识别（豆包调用）

**版本 v4.9**　本地图片/PPT 高精度结构化识别（两部分输出：客观描述 + 明确结论型推想；结果以豆包为准不二次判断；推断由可编辑槽位按图片类型生成，用户点明角度优先；防个体遗漏） · 全自动 GUI 编排 · 单批 10 张硬约束 · 可验证完成判据
变更历史见同目录 `CHANGELOG.md`。

---

## 0. 导航与执行总览

### 0.1 三条执行分支

> **首次使用（或在别的机器上部署）：先读 0.2 环境要求 → 0.3 首次使用与配置 → 0.7 冒烟自检**，通过后再进入下面任一条分支。

| 任务形态 | 执行链路 | 快捷入口 |
|---|---|---|
| 本地图片文件/文件夹识别 | 准备 P0–P2 → 步骤3 上传与批控 → 步骤4 发送 → 步骤5 轮询读取 → 步骤6 输出 → 步骤8 收尾 | 本文档全程 |
| 只问 PPT 的文本内容（总结/讲了什么/大纲） | P0–P2 → 步骤0.6（直传文件本体，不拆解） → 4.1 的 PPT 判据 → 步骤6 → 步骤8 | 见「4. 执行层：步骤0.6」 |
| 问 PPT 页面内的图片/照片内容 | P0–P2 → 步骤0.2 整页渲染 PNG → 按图片链路识别（每批 ≤10 页） | 见「3. 执行层：步骤0」 |

任务中途需求变更（如先问 PPT 大纲、后问某页配图）→ **立即切换到对应分支的完整链路**，不复用另一分支的中间结论。

### 0.2 环境要求（移植前先核对）

| 必需 | 说明 |
|---|---|
| 操作系统 | **Windows**（依赖 Win32 窗口 API 与 UI Automation） |
| PowerShell | **5.1 或更高**。脚本含中文，**存为文件时必须 UTF-8 BOM**，否则报错乱码 |
| 豆包桌面客户端 | 已安装**并已完成登录**；本 skill 需运行其 GUI，非 API 调用 |
| 输入法 | 建议关闭中文输入法候选干扰（流程已内置 `Esc` 清候选与"只走剪贴板粘贴"的防护） |

| 视环境而定 | 处置 |
|---|---|
| 豆包安装路径 | 默认 `D:\AppStoreSoftstore\Install\doubao\Doubao.exe`；不一致时由 **0.3 的配置块自动递归探测**（T1 探测优先级） |
| 临时产物目录 | 默认 `D:\DS\.dsh`（T3）；**不存在会被自动创建**，其他地方可覆盖 |
| 新对话快捷键 | 默认 `Ctrl+Shift+K`（T8）；与其他软件冲突时流程会自动回退到 UIA 点击"新对话" |
| 机械革命应用商店 | 仅在有 `PCAppStore` 进程的环境才涉及（步骤8-5）；无关环境自动跳过 |

> **客户端版本敏感**：本 skill 的附件元素类名（T11）与输入框元素（T12）为**实测值**，随豆包版本可能变化。换环境后**先跑 0.7 冒烟自检**，不要直接上批量任务。

### 0.3 首次使用与配置（三件事）

**① 探测豆包路径**（T1）——先跑这一步，后续代码都用探测结果：

```powershell
# 优先官方默认路径；不存在则在常见安装根目录内递归查找；再不行才报错。
$doubaoExe = 'D:\AppStoreSoftstore\Install\doubao\Doubao.exe'
if (-not (Test-Path -LiteralPath $doubaoExe)) {
    # 覆盖不到（自定义目录/便携版/其他盘）时：直接给 $doubaoExe 赋绝对路径覆盖本块
    $roots = @('D:\AppStoreSoftstore', "$env:LOCALAPPDATA\Doubao", "$env:ProgramFiles", "${env:ProgramFiles(x86)}")
    foreach ($r in $roots) {
        if (-not (Test-Path -LiteralPath $r)) { continue }
        $cand = Get-ChildItem -LiteralPath $r -Recurse -Filter 'Doubao.exe' -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($cand) { $doubaoExe = $cand.FullName; break }
    }
}
if (Test-Path -LiteralPath $doubaoExe) { "DOUBAO_EXE=$doubaoExe" }
else {
    Write-Error (
'
未找到 Doubao.exe。已尝试：默认路径 → 注册表 App Paths → 四个常见安装根目录。
'
 + [Environment]::NewLine +

'
请任选一种方式继续：
'
 + [Environment]::NewLine +

'
  1) 手工定位：Get-ChildItem <安装盘>:\ -Recurse -Filter Doubao.exe -ErrorAction SilentlyContinue | Select-Object -First 1
'
 + [Environment]::NewLine +

'
  2) 直接指定：$doubaoExe =
'
'
你的绝对路径\Doubao.exe
'
'
，然后重跑本块
'
 + [Environment]::NewLine +

'
注意：仅启动豆包这一步需要该路径；豆包已在运行时，全流程不依赖它。
'
)
}
```

**② 配置产物根目录**（T3）——不存在的会被创建；如沿用默认值可跳过：

```powershell
$tmpRoot = 'D:\DS\.dsh'                                  # 可改为任意有写权限的目录
if (-not (Test-Path -LiteralPath $tmpRoot)) { New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null }
"TMP_ROOT=$tmpRoot  writable=$((New-Item -ItemType File -Path (Join-Path $tmpRoot '.wtest') -Force).Length -ge 0)"
Remove-Item -LiteralPath (Join-Path $tmpRoot '.wtest') -Force -ErrorAction SilentlyContinue
```

**③ 登录与可访问性**：确认豆包已登录。若窗口能打开但流程始终探测不到输入框，按 T2 参数重启：

```powershell
Start-Process -FilePath $doubaoExe -ArgumentList '--force-renderer-accessibility'
```

### 0.4 执行总览

```
任务类型分流（1.2）
   │
   ├─ 图片 ──────────── 准备(P0–P2) → 上传(步骤3) → 发送(步骤4) → 轮询(步骤5) → 输出(步骤6) → 收尾(步骤8)
   ├─ PPT 文本分析 ──── 准备(P0–P2) → 直传文件(0.6) → 发送固定分析指令 → 轮询240s(0.6⑤) → 输出(0.6⑦) → 收尾(步骤8)
   └─ PPT 画面识别 ──── 渲染(0.2) → 准备(P0–P2) → 上传页图(步骤3) → 发送(步骤4) → 轮询(步骤5) → 输出(步骤6) → 收尾(步骤8)

每步的"完成判据"（不具备 = 不得进入下一步）
   P1 → 输出 DOUBAO_READY
   P2 → 探测到 ProseMirror 输入框，且新对话标题回到「豆包」
   步骤3 → 附件数 == 本批张数 且 无旧会话回复残留
   步骤5 → 命中完成标记；或两部分字段齐（8 客观 + 5 推断）等兜底判据成立
   步骤8 → 清理已执行，豆包已最小化，前台已还原

图片回复的统一输出契约（v4.1 起）：
   第一部分 客观描述（8 字段，纯事实）→ 第二部分 推想预测（5 字段，带置信度与依据）
   两部分都必须拿到；汇报时分区呈现，推断不得混入事实。
```

### 0.4a 入口闸门：先验窗口状态，再动手（**不得跳过**）

**每次任务的第一步不是上传，是确认豆包窗口状态。** 上传/发送/读取三类操作全部要求窗口**存在且展开**；状态不明就上传，必然失败且难以定位。

```powershell
# 入口闸门：任何上传/发送动作之前必须先跑这一段，输出四态结论之一
# 注意：RECT 必须在 -TypeDefinition 里声明，且命名空间要写在 C# 代码内
#      （Add-Type -MemberDefinition 不接受 struct；-TypeDefinition 与 -Namespace 参数集冲突）
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace Win32 {
    public class GateWin {
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    }
}
'@

$withWin = @(Get-Process -Name Doubao -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
$allProc = @(Get-Process -Name Doubao -ErrorAction SilentlyContinue)

if ($allProc.Count -eq 0) {

'
GATE = NO_PROCESS  豆包未运行 → 走 PP2.0/PP2.1 路径 B 启动，再回到本闸门复检
'
} elseif ($withWin.Count -eq 0) {
    "GATE = NO_WINDOW   有 $($allProc.Count) 个豆包进程但均无主窗口（窗口被关闭/最小化到托盘）"

'
                   → 执行 PP2.1 路径 B：Start-Process 启动本体，恢复既有会话；复检通过前禁止上传
'
} else {
    $d = $withWin[0]
    $r = New-Object Win32.GateWin+RECT
    [Win32.GateWin]::GetWindowRect($d.MainWindowHandle, [ref]$r) | Out-Null
    $w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
    if ([Win32.GateWin]::IsIconic($d.MainWindowHandle)) {
        "GATE = ICONIC      pid=$($d.Id) 窗口已最小化 → 先 SW_RESTORE 恢复并确认前台，再上传"
    } elseif ($w -le 200 -or $h -le 200 -or $r.Left -lt -10 -or $r.Top -lt -10) {
        "GATE = OFFSCREEN   pid=$($d.Id) 几何异常 L=$($r.Left) T=$($r.Top) W=$w H=$h → 按 PP2.1 三重校验恢复"
    } else {
        "GATE = READY       pid=$($d.Id) 标题=[$($d.MainWindowTitle)] W=$w H=$h → 可以进入上传"
    }
}
```

**硬性要求**：

| 闸门结论 | 允许的下一步 |
|---|---|
| `READY` | 进入 PP2.3 新对话 → 上传 |
| `ICONIC` / `OFFSCREEN` | 先恢复窗口并确认 `DOUBAO_READY`，**复检为 READY 后**才上传 |
| `NO_WINDOW` / `NO_PROCESS` | 先按 PP2.0/PP2.1 路径 B 启动，**复检为 READY 后**才上传 |
| 复检仍不通过 | **停下来问用户**（可能需要人工点托盘图标），不要盲目重试或直接上传 |

> **实测教训（2026-10）**：曾跳过本闸门直接上传，结果 `ATTACH=1/2` 校验失败，被误判为"粘贴间隔太短"，实际是**窗口被最小化了**——最小化后 UIA 读不到附件，计数自然停在旧值。**先验状态再动手，可避免整类误诊。**
### 0.5 时序基线（实测，用于判断"是否卡住"）

> 下表为 2026-10 实测值（v4.5 提速后）。环境不同会有偏差，**判据仍是"超过阈值上限"而非"接近典型值"**。

| 环节 | 实测典型值 | 超时/异常阈值 |
|---|---|---|
| 冷启动豆包到窗口可用 | ~3 秒 | 45 秒（超时走路径 B 报错） |
| 窗口恢复激活 | 1–3 秒 | 3 轮尝试后报错 |
| 新建对话（标题回到「豆包」） | **165 ms** | 4 秒（T15） |
| 输入框就绪（新对话后） | **416 ms**（最慢约 5.8 秒） | 8 秒（T15） |
| 单张图片粘贴 | 0.45 秒/张（T7，刻意保留） | 整批贴完统一校验 |
| 附件渲染入框（多图/单图） | **69 ms** | 3 秒（T15） |
| 发送确认（输入框清空） | ~0.6 秒 | 5 秒（T15） |
| 上传→发送 全流程（2 张图） | **~3.6 秒** | — |
| 标记基线采集（稳定判定） | 0.5–4 秒 | 8 秒（T15）；未收敛告警 `BASE_NOT_STABLE` |
| 图片识别轮询 | **9.6 秒**（2 张图 26 字段） | 150 秒（T5） |
| PPT 文本分析轮询 | 40–120 秒 | 240 秒（T5） |
| PPT 整页渲染 | 1–3 秒/页 | COM 不可用 → 报错或退化 |
| 结果重抓（不重传） | ~30 秒 | 30 秒（T19） |

### 0.6 阅读约定

- 正文中的 **PowerShell 代码块：整段复制执行，不逐行改写**。行内注释是执行前提，不是说明文字。
- **每段代码都在独立的 PowerShell 进程中运行**，变量不跨调用保留。跨步骤复用前，先重建上下文变量（模板见 P0 与 1.3 任务台账）；PP2.1 的窗口状态、PP2.2 的定位函数与 `$doubao` 必须同段执行。
- 文档内所有数值、路径、超时统一以《1.5 常量与阈值》为**单一信源**；正文只引用，不重复定义。

---

### 0.7 环境自检（**一次性**：换机器/换客户端版本后跑一次，**无需每次识别前跑**）

对应附录D 的 V2b/V2f，**把"版本漂移"从"跑任务时才发现"提前到"开跑前就发现"**。前置：豆包已 `DOUBAO_READY`。

**设计纪律：本自检只读，不改变应用状态。** 早期版本曾要求"先贴 1 张测试图片"以验证附件计数——但那会在输入复合框里留下**真实待发附件**，而**新建对话并不能清除它**（2026-10 实测：贴图后 2 个附件，执行新建对话后仍为 2 个），一旦随后误按发送就会把测试图连指令一起发出。故改为下述四项**全部只读**的检查。

| # | 检查项 | 覆盖的失效模式 | 副作用 |
|---|---|---|---|
| ① | 输入框可定位（T12） | 控件类型/类名随版本变化 | 无 |
| ② | 窗口内存在附件类名元素（T11，新旧两种） | 附件类名随版本变化 | 无 |
| ③ | **T9 正则对提示词模板的误判 = 0**（离线读 SKILL.md 自测） | **提前判完成、读到空结果**（最危险） | 无 |
| ④ | 依赖可达（`$doubaoExe` 存在、`$tmpRoot` 可写） | 换机器后的路径/权限问题 | 无（只建一个临时探针文件后即删） |

> ② 只证明"类名在 DOM 中存在"，**不证明"粘贴后一定入框"**——后者由正式任务第一次上传时的 3.3 附件数校验兜底（那本来就是它的职责，且判据更严：必须等于本批张数）。

```powershell
# 四项只读自检；②③ 不过就不要继续做正式任务
$pass = 0; $fail = 0; $warn = 0
$win = Get-DoubaoWinByProcess $doubao
$all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)

# ① 输入框可定位（T12：只按类名，不限 ControlType）
$edit = $all | Where-Object { $_.Current.ClassName -like '*ProseMirror*' } | Select-Object -First 1
if ($edit) { "① INPUT_OK   class=[$($edit.Current.ClassName)]"; $pass++ }
else { "① INPUT_FAIL 未定位到输入框 → 按 0.3 ③ 用 T2 参数重启豆包"; $fail++ }

# ② 附件类名存在性（T11）——只读扫描，不贴任何文件
#    注意：类名匹配必须覆盖新旧两种命名；命中数不代表待发附件数（可能含预览缩略图）
$imgHits = @($all | Where-Object { $_.Current.ClassName -match 'image-container|image-wrapper' }).Count
if ($imgHits -ge 1) { "② ATTACH_OK   命中附件类名元素 $imgHits 个（类名未失效）"; $pass++ }
else { "② ATTACH_WARN 未命中附件类名元素——若界面上确有图片，则 T11 已失效需回填；纯空会话也可能为 0"; $warn++ }

# ③ T9 正则离线自测：直接读 SKILL.md，抽出步骤4 的 $prompt 骨架逐行验证
#    这是"提前判完成"缺陷的唯一护栏，且完全不需要 GUI
# 定位 SKILL.md：当前目录 → 本块所在目录 → 技能安装路径，先命中者用
$skillPath = $null
foreach ($cand in @('SKILL.md', (Join-Path $PSScriptRoot 'SKILL.md'),
        'D:\DS\.dsh\skills\local-image-doubao-recognition\SKILL.md')) {
    if ($cand -and (Test-Path -LiteralPath $cand)) { $skillPath = (Resolve-Path -LiteralPath $cand).Path; break }
}

$fp = -1
if (Test-Path -LiteralPath $skillPath) {
    $st = [System.IO.File]::ReadAllText($skillPath)
    $m = [regex]::Match($st, "(?ms)\`$prompt = @""\r?\n(.*?)\r?\n""@")
    # 正则取自文档中"正在使用"的定义处（比从骨架文本反推更稳：骨架格式改动不会让它失效）
    # 文档里有多处 $replyPattern 赋值（含本自检自己的提取模式），取最长的那个 = 真正的字段行正则
    $mPat = [regex]::Matches($st, "`$replyPattern =
'
(.*?)
'
") |
        Where-Object { $_.Groups[1].Value.Length -gt 50 } | Select-Object -First 1
    if ($mPat.Success) { $replyPattern = $mPat.Groups[1].Value }
    if ($m.Success -and $mPat.Success) {
        # 使用上方从文档提取到的 $replyPattern
        $fp = @(($m.Groups[1].Value -split "`n") | Where-Object { $_ -match $replyPattern }).Count
    }
}
if ($fp -eq 0) { "③ PARSER_OK   骨架内被误判为回复行 = 0 行"; $pass++ }
elseif ($fp -gt 0) { "③ PARSER_FAIL 骨架被误计 $fp 行 → T9 排除式已失效，会导致提前判完成"; $fail++ }
else { "③ PARSER_WARN 未能读取 SKILL.md（路径不可达），本项跳过"; $warn++ }

# ④ 依赖可达：T1 可执行文件、T3 产物根目录可写
$depOk = $true
if (-not (Test-Path -LiteralPath $doubaoExe)) { "④ DEP_FAIL   T1 不可达: $doubaoExe"; $depOk = $false }
if (-not (Test-Path -LiteralPath $tmpRoot)) {
    try { New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null } catch { }
}
if (Test-Path -LiteralPath $tmpRoot) {
    $probe = Join-Path $tmpRoot ('.wtest_' + [guid]::NewGuid().ToString('N').Substring(0, 6))
    try {
        New-Item -ItemType File -Path $probe -Force | Out-Null
        Remove-Item -LiteralPath $probe -Force
    } catch { "④ DEP_FAIL   T3 不可写: $tmpRoot"; $depOk = $false }
} else { "④ DEP_FAIL   T3 不存在且无法创建: $tmpRoot"; $depOk = $false }
if ($depOk) { "④ DEP_OK     T1 与 T3 均可达"; $pass++ } else { $fail++ }

"SELFTEST_RESULT = $pass pass / $fail fail / $warn warn"
if ($fail -gt 0) { Write-Output '⚠ 自检未通过：按上方提示修正后再执行正式任务' }
elseif ($warn -gt 0) { Write-Output '✅ 关键项通过（有 warn 项未覆盖，见上）' }
else { Write-Output '✅ 自检通过，环境可用' }
```

> **运行频率（重要）**：**换机器、换豆包客户端版本、或流程出现异常时才跑**。同一环境下**不需要每次识别前跑**——
> ① 在 PP2.3 早失败探测里本就会触发；④ 在首次上传时触发；② 在 3.3 附件数校验里有更严格的版本（要求等于本批张数）。
> 本自检的独特价值只有 ③（离线验证 T9 正则），而 T9 只在**改动 skill 本身**时才可能失效。
>
> **耗时**：实测约 100 ms（UIA 扫描 37–119 ms + 读文档自测 13 ms + 目录探针 9 ms），
> 相对单图识别全流程约 13 秒（上传发送 3.6 s + 识别 9.6 s）可忽略。**即使每次都跑也不影响速度**，但仍不建议——没有额外收益。
---

## 1. 前置校验与任务契约

### 1.1 前置校验（全部通过才可进入执行层）

| 项 | 校验内容 | 不通过时 |
|---|---|---|
| C1 绝对路径化 | 相对路径（如 `CSY`、`.`）必须先转绝对路径，禁止把 `.` 或相对路径交给文件工具 | 用 `Resolve-Path` / `Get-Item` 取全路径 |
| C2 存在性 | 目标文件/文件夹真实存在、无拼写错误 | 告知路径无效，请用户核对 |
| C3 格式 | 图片来源：JPG/JPEG/PNG/BMP/WebP；剔除无效格式并记录 | 剔除并说明 |
| C4 可读性 | 文件未损坏、可读、无权限限制 | 告知权限/损坏，请用户处理 |
| C5 计数 | 统计有效张数，切批 ≤10（T4）；PPT 按渲染页数计 | 超过 10 → 按步骤3-3 分批 |
| C6 合规初筛 | 依据文件名、扩展名、路径、体积等元数据判断是否疑似敏感违规；**样本不确定即按正常处理**，不得以"疑似"为由拒识 | 明确违规 → 终止，仅输出「无法识别该图片」 |
| C7 类型分流 | 图片 → 步骤3；.ppt/.pptx → 按 1.2 分流 | — |

### 1.2 任务类型分流

| 判定 | 走向 |
|---|---|
| 目标为图片文件 | 准备阶段 → 步骤3 图片链路 |
| 目标为 .ppt/.pptx，且用户只问文本内容（总结/主要讲了什么/大纲/观点） | 4. 执行层：步骤0.6 轻量模式（不渲染、不拆页） |
| 目标为 .ppt/.pptx，且用户问页面内图片/照片/配图/图标内容 | 4. 执行层：步骤0 拆解渲染 → 图片链路 |

#### 冲突时的优先级（自上而下，上位覆盖下位）

| 优先级 | 来源 | 说明 |
|---|---|---|
| **P1 最高** | **用户在本次请求中点明的角度、侧重点、答案形态** | 如"从植物学角度认""我要的是原理不是名字""只看颜色别猜品牌"。**一律优先采纳**，并据此填写步骤4 的槽位；不得因默认分流、领域习惯或"这样更好"而覆盖用户指定角度 |
| P2 | 本次工作/下游任务的实际需要 | 用户未点明时，按任务目的推断侧重，并在必要时向用户说明依据 |
| P3 最低 | 本 skill 的默认槽位与泛化措辞（附录A.1a 的"默认"列） | 仅在前两者都无信息时使用 |

> 边界：**P1 只改"看什么、怎么答"，不改"必须输出什么"**。用户点明的角度不能用来省略第一部分、砍掉字段或绕过格式（那会破坏 T9/T10 判据）；若用户明确说"不要猜"，按《A.1a》的对应填法处理——第二部分仍输出字段，但如实说明依据不足。

### 1.3 任务台账（识别开始前建表，后续所有校验以此为基准）

| 字段 | 取值说明 |
|---|---|
| 任务目录 | 绝对路径 |
| 有效张数 | 图片数 / PPT 渲染页数 |
| 切批结果 | 每批张数（≤10），共几批 |
| **本批张数** | 当前批次张数，**5-3 校验与 3.3 附件校验的基准就是它，不是任务总数** |
| 执行模式 | 图片 / PPT 文本分析 / PPT 画面识别 |
| 产物目录 | `D:\DS\.dsh\tmp_ppt_<8位hex>`（仅 PPT 拆解模式存在）；图片模式无此目录 |
| 原前台进程 | P1 记录，收尾还原用 |

### 1.4 接口契约

**完成信号**

| 模式 | 主判据 | 兜底判据 |
|---|---|---|
| 图片 / PPT 文本分析 | **结束标记元素数超过基线**（见步骤5-1；两条链路同构） | **无兜底**：标记为唯一判据，不出现即报超时（v4.8 决定，理由见该节注释） |


**结果字段契约**：单张输出**两部分共 13 个字段**；多张按「图1…图N」成组输出，每组 13 个字段。
**第一部分 客观描述（8 字段，不可省略）**——只描述视觉可见内容，不得混入推断：

| 字段 | 无内容时 | 说明 |
|---|---|---|
| 图片类型 | 无 | 实拍照片/截图/插画/海报/证件/图表/漫画等 |
| 核心主体 | 无 | 1–3 个最突出对象 |
| 空间布局 | 无 | 相对位置与占比 |
| 场景与背景 | 无 | 环境、光线、色调 |
| 画面文字 | 无 | 逐字转录并标位置 |
| 细节特征 | 无 | 颜色、材质、姿态、服饰等 |
| 特殊元素 | 无 | 二维码、logo、印章等，仅描述外观 |
| 不确定内容说明 | 无 | 模糊/遮挡/反光区域；无则填「无」 |

**第二部分 推想预测（5 字段，不可省略）**——在第一部分事实基础上做可溯源的推断：

| 字段 | 无内容时 | 说明 |
|---|---|---|
| 推断对象 | 无法判断 | 角色/作品/场景/物体/地点/事件的身份判断；**优先给确定的名字，多候选按可能性排序** |
| 置信度与理由 | 无 | 高/中/低 + 一句话依据（口径见下） |
| 判断依据 | 无 | 逐条引用第一部分的可观察特征，用「；」分隔，至少 1 条 |
| 其他可能 | 无 | 合理的备选解释，逐条列出，用「；」分隔 |
| 无法推断项说明 | 无 | **仅列确实无法判断的项**；能判断的不要写在这里；无则填「无」 |

**置信度口径（决定"高/中/低"怎么给）**：**高** = 有明确可识别的特征或标志性线索支撑（如特定服饰纹样、标志性配饰、可辨识文字、独有场景元素）；**中** = 依据充分但存在其他解释；**低** = 仅能依据风格或概率推断。
> 定级原则：**先给结论，再按实际把握定级**。识图推测本身可靠性较高，不要习惯性压低等级；但也**不得给无依据的"高"**——"高"必须能在「判断依据」里指出具体特征。

**防遗漏契约（硬性）**：

| 要求 | 判定 |
|---|---|
| 画面中可辨识的个体/对象/条目**逐个列出** | 出现"等/等等/其他/若干/众多/大量/其余人物/背景人物"把多个个体合并带过 → **视为遗漏** |
| 无法命名时给出**可定位标识** | 位置（"左侧第一人""前排右二"）或外观代称（"白发少年"）均可 |
| 密集群像/远景小人 | 必须给**可计数信息**（数量或可数范围）+ 分布位置 + 主要个体分别描述 |
| 适用范围 | 核心主体、细节特征、画面文字、特殊元素、以及第二部分的推断对象 |
| 追补 | 发现遗漏时按步骤7-1 模板 D 追问补齐，**不要自行补写** |

**硬性纪律**：

1. **必答**：第二部分以给出明确判断为主。**不得因为"这是推断"而回避作答**，也不得用「无法判断」规避本可判断的内容；只有确实无法判断时才填「无法判断」。
2. **有据**：任何结论都必须能追溯到第一部分的可见特征，并在「判断依据」中逐条列出。
3. **分区**：第二部分内容一律视为推断，**不得当作客观事实引用**（汇报时见步骤6 三层分区）。

**字段行三态语义**（决定 5-3 是否判失败）：
1. `字段名：具体内容` → 有效行；
2. `字段名：无` → **有效行**（表示该维度确实没有内容，不是失败）；
3. 字段名整行缺失 → **失败**，重发指令 1 次，仍失败才输出「图片识别结果格式异常，无法解析」。

> 说明：若豆包把某字段写成 `字段名：（无文字）` 这类括号示例形态，该行不会被字段行正则计入。此时先人工核对原回复是否已完整输出 13 个字段；**确认已输出即视为有效**，不要盲目重发。

**失败态**：无法识别（全黑/全白/纯噪点/违规）→ 豆包只回「无法识别该图片」；本 skill 保留该原文，不补充推测。

### 1.5 常量与阈值（单一信源）

| 编号 | 常量 | 取值 |
|---|---|---|
| T1 | 豆包可执行文件 | 默认 `D:\AppStoreSoftstore\Install\doubao\Doubao.exe`；**探测逻辑见 0.3 ①**（注册表 App Paths → 默认路径 → 四个常见安装根目录递归 → 报错）；`$doubaoExe` 可被直接覆盖 |
| T2 | 启动参数 | `--force-renderer-accessibility`（暴露 UI Automation 元素） |
| T3 | 产物目录 | `D:\DS\.dsh\tmp_ppt_<8位hex>`（仅本次任务，收尾必删；**根目录可经 0.3 ② 覆盖**，不存在会自动创建） |
| T4 | 单批附件上限 | 10 张（豆包单条消息上限，超出部分不入框） |
| T5 | 轮询上限 | 图片 **90 秒**；PPT 文本分析 **240 秒**（标记判据独立后正常 5–20 秒结束，长超时只在真异常时生效） |
| T6 | 轮询间隔 / 标记确认延时 | 500 ms / 1.5 秒（等最后一行渲染完；**不得为提速缩短**） |
| T7 | 单张粘贴间隔 | 450 ms（**不得缩短**：缩短会与剪贴板读取竞态） |
| T8 | 新对话快捷键 | `Ctrl+Shift+K` |
| T9 | 字段行正则（**两部分 13 字段**） | `^(图片类型\|核心主体\|空间布局\|场景与背景\|画面文字\|细节特征\|特殊元素\|不确定内容说明\|推断对象\|置信度与理由\|判断依据\|其他可能\|无法推断项说明)\s*：\s*(?![（(])\S` |
| T10 | 末字段正则 | 第一部分末字段：`不确定内容说明\s*：\s*(?![（(])\S`；第二部分末字段：`无法推断项说明\s*：\s*(?![（(])\S`（完成判定用第二部分末字段） |
| T11 | 附件元素类名 | 正则 `image-container\|image-wrapper`（兼容新旧版命名，见下注） |
| T12 | 输入框元素 | 类名含 `ProseMirror`（**不限定 ControlType**，见下注） |
| T13 | 图片扩展名白名单 | `.jpg .jpeg .png .bmp .webp` |
| T14 | 两部分字段分组（供正则/校验引用） | 第一部分（8 字段）：图片类型、核心主体、空间布局、场景与背景、画面文字、细节特征、特殊元素、不确定内容说明；第二部分（5 字段）：推断对象、置信度与理由、判断依据、其他可能、无法推断项说明 |
| T15 | 提速策略（**原则：只压固定等待，不压判据**） | ① 所有"等 UI 就绪"改为自适应轮询：新建对话 150 ms 间隔/上限 4 s、输入框 150 ms/10 s、附件 200 ms/3 s、发送确认 200 ms/5 s、基线稳定检测 250 ms/8 s；② **不得**为提速缩短粘贴间隔（T7 450 ms）与标记确认延时（T6 1.5 s）；③ 单次全树扫描仅约 9–40 ms（150–330 元素），加密轮询几乎无成本；④ 多步合并进同一进程可省往返开销 |
| T16 | 输入框早失败探测上限 | 10 秒（`DOUBAO_READY` 后探不到输入框即中止，避免空等） |
| T17 | PPT 文件附件上传校验上限 | 20 秒（大文件上传更慢；超时即按 3.1 步骤3 的降级路径处理） |
| T18 | 「仅重抓」单次操作上限 | 30 秒（只读取不发送；**须先恢复窗口展开**，最小化态 UIA 读不到；超时则报告并保留已得结果） |
| T19 | 已验证的客户端版本 | 豆包桌面版**实测通过版本**（2026-10）：输入框类名 `tiptap ProseMirror`（Group）、附件类名 `image-wrapper-*`、PPT 附件以独立 Text 暴露文件名 + `上传中... N%` 进度。**换版本后 T11/T12 可能失配 → 必跑 0.7 环境自检** |

> **T9/T10 为什么长这样**（改任一处前先读这段）：
> - `\s*：\s*` 是必需的：**UIA 会在 ASCII 与中文之间自动插空格**，实测同一行可能渲染成 `图片类型：插画`、`图片类型 ： 插画` 或 `图片类型 ：（…`。原写法写死"字段名紧跟全角冒号"，会**漏掉**被插空格的回复行。
> - `(?![（(])\S` 用来排除用户提示词模板（模板写作"字段名：（示例…"）。**不能简写成 `[^（(]`**：那样遇到 `图片类型 ： （…` 时，`：` 后的第一个字符是空格而非括号，模板行仍会被当成有效回复行（2026-10 实测：只有用户消息时误计 16 行，足以让 5-1 的"≥8 行"兜底判据在 AI 还没回答时就判完成）。
> - 不要用 `\s*+` 这类占有量词：**.NET 正则不支持嵌套量词**，会直接抛 `Nested quantifier` 异常。
> - **T10 必须锚定"整份回复的最后一个字段"（第二部分末字段 `无法推断项说明`），不是第一部分的 `不确定内容说明`**。v4.1 把输出扩成两部分后，若沿用旧锚点，5-1 会在**第一部分刚写完时就判完成，第二部分被整段丢弃**。追问场景若只输出第一部分，锚点会自动落回 05。
>
> **T12 为什么不能限定 ControlType**：实测同一客户端在不同版本下把输入框暴露为不同控件类型——旧版是 `ControlType.Edit`，现版本是 `ControlType.Group`（类名 `tiptap ProseMirror ProseMirror-focused`）。**按 `Edit + *ProseMirror*` 双条件定位会直接失配**（2026-10 实测：0 命中，早失败探测误报为"未登录或界面异常"）。只按类名 `*ProseMirror*` 定位，两种形态都能命中。
>
> **T11 为什么要写两个类名**：附件缩略图的类名同样随版本变化——旧版 `image-container-*`，现版本 `image-wrapper-YJelRW`。**只匹配 `*image-container*` 时，图片明明已入框却数出 0，会误触发"尾部缺失补贴"重复粘贴**（2026-10 实测：2 张已入框，旧判据得 0，新判据得 2）。用正则同时覆盖两种命名即可。

---

## 2. 执行层：准备阶段

### P0 建立本次任务执行上下文

目的：把本任务常量固化成可直接复制的变量头，供后续每段独立进程直接使用。

> 每段代码都在独立 PowerShell 进程中运行；需要跨步骤的变量，先粘贴本变量头并覆盖 `$imageDir` / `$pptPath` 等任务相关值。

```powershell
# ===== 本次任务变量头（每段代码执行前先粘贴，按任务改前两行）=====
$imageDir = 'D:\CSY'        # 图片模式：目标目录；PPT 模式保留但不用
$pptPath  = ''              # PPT 模式：目标文件绝对路径；图片模式留空
$tmpRoot  = 'D:\DS\.dsh'    # T3 产物根目录
$helperScript = $null       # 本次若写过辅助脚本，填其绝对路径；否则保持 $null

# T9 字段行（只命中豆包回复，不命中用户提示词模板）
$replyPattern = '^(图片类型|核心主体|空间布局|场景与背景|画面文字|细节特征|特殊元素|不确定内容说明|推断对象|置信度与理由|判断依据|其他可能|无法推断项说明)\s*：\s*(?![（(])\S'
# T10 末字段：出现即说明回复已写到结尾字段
$lastFieldPattern = '无法推断项说明\s*：\s*(?![（(])\S'

# 通用工具：定位豆包主进程（无类型依赖，任意段可先用）
$doubao = Get-Process -Name Doubao -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
```

### P1 记录原前台窗口

目的：任务结束后把前台还给用户原来的窗口。

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
# 收尾用 $originalPid 还原前台（AppActivate 需要进程 id，见步骤8）
```

### P2 窗口就绪闸门

**本阶段输出的 `DOUBAO_READY` 是后续所有快捷键、粘贴、坐标点击操作的唯一准入条件。**

#### PP2.0 豆包进程与安装路径

```powershell
$doubao = Get-Process -Name Doubao -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } |
    Select-Object -First 1
```

> 找到的窗口可能是**最小化状态**（上次任务结尾按规则最小化保留）。最小化窗口仍有 `MainWindowHandle`，只看句柄会误判"已就绪"——必须由 PP2.1 用 `IsIconic` + 几何校验确认真正恢复可见。
>
> 若需启动：只用豆包主程序本体（T1）。**禁止通过应用商店入口或商店快捷方式启动，避免误开机械革命应用商店。**

```powershell
# T1：优先用 0.3 ① 探测出的 $doubaoExe；本进程未探测过则就地探测（代码保持自包含）
if (-not $doubaoExe -or -not (Test-Path -LiteralPath $doubaoExe)) {
    $doubaoExe = 'D:\AppStoreSoftstore\Install\doubao\Doubao.exe'
    if (-not (Test-Path -LiteralPath $doubaoExe)) {
        # 覆盖顺序：注册表 App Paths（最权威）→ 常见安装根目录
        foreach ($hk in @(
'
HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Doubao.exe
'
,

'
HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Doubao.exe
'
)) {
            try { $v = (Get-ItemProperty -Path $hk -ErrorAction Stop).
'
(default)
'
                  if ($v -and (Test-Path -LiteralPath $v)) { $doubaoExe = $v; break } } catch { }
        }
$roots = @('D:\AppStoreSoftstore', "$env:LOCALAPPDATA\Doubao", "$env:ProgramFiles", "${env:ProgramFiles(x86)}")
        foreach ($r in $roots) {
            if (-not (Test-Path -LiteralPath $r)) { continue }
            $cand = Get-ChildItem -LiteralPath $r -Recurse -Filter 'Doubao.exe' -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if ($cand) { $doubaoExe = $cand.FullName; break }
        }
    }
}
Start-Process -FilePath $doubaoExe -ArgumentList '--force-renderer-accessibility'
```

#### PP2.1 恢复并激活窗口（闸门，必须确认输出 `DOUBAO_READY`）

策略：路径 A 快速恢复已运行窗口（秒级、保留会话与已上传图片）→ 失败走路径 B `Start-Process` 启动兜底 → 再失败 `AppActivate` 兜底 → 全部失败才报错退出。

> 复用已运行的豆包时（上次任务结尾窗口被最小化保留），**绝不能只靠 `AppActivate`**：实测它对最小化窗口可能无效。恢复成功后 `$doubao` / `$doubaoReady` / `Win32.Win32WinState` 只在本进程内有效，跨段使用前必须重跑本块。

```powershell
# ========== 窗口恢复激活代码块（完整复制执行） ==========
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
# 进程未运行时则冷启动新实例。启动后统一重新走三重校验（新进程句柄可能变化）。
if (-not $doubaoReady) {
    Write-Output '快速恢复失败，改用 Start-Process 指令打开豆包'
    # T1：优先用 0.3 ① 探测出的 $doubaoExe；本进程未探测过则就地回退探测（保持自包含）
    if (-not $doubaoExe -or -not (Test-Path -LiteralPath $doubaoExe)) {
        $doubaoExe = 'D:\AppStoreSoftstore\Install\doubao\Doubao.exe'
        if (-not (Test-Path -LiteralPath $doubaoExe)) {
            # 覆盖不到（自定义目录/便携版/其他盘）时：直接给 $doubaoExe 赋绝对路径覆盖本块
    $roots = @('D:\AppStoreSoftstore', "$env:LOCALAPPDATA\Doubao", "$env:ProgramFiles", "${env:ProgramFiles(x86)}")
            foreach ($r in $roots) {
                if (-not (Test-Path -LiteralPath $r)) { continue }
                $cand = Get-ChildItem -LiteralPath $r -Recurse -Filter 'Doubao.exe' -ErrorAction SilentlyContinue |
                    Select-Object -First 1
                if ($cand) { $doubaoExe = $cand.FullName; break }
            }
        }
    }
    if ($doubaoExe -and (Test-Path -LiteralPath $doubaoExe)) {
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
        Write-Error '未找到 Doubao.exe（已尝试默认路径/注册表 App Paths/四个常见根目录）。请手工把 $doubaoExe 赋为绝对路径后重试；豆包已在运行时不影响后续流程'; exit 1
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

执行前提（三条硬约束）：

- 窗口最小化时 UIA `BoundingRectangle` 返回无效离屏坐标，**必须先恢复再取坐标**；
- `mouse_event` 只按当前光标位置点击、不移动光标，模拟点击前必须用 `[System.Windows.Forms.Cursor]::Position` 先把光标移过去；
- 路径 B 仅作兜底，不默认每次执行：已运行实例重复启动只触发激活、不保证恢复最小化窗口，冷启动也更慢。

#### PP2.2 通用窗口定位函数（按进程 ID，去标题依赖）

豆包标题会在任务中变化（「豆包」→「图片识别任务说明 - 豆包」等）。**按标题找窗口会在标题变化时短暂"找不到窗口"，误触发恢复逻辑**，因此所有"重新获取豆包窗口"的代码统一调用本函数。

```powershell
# 返回豆包进程中面积最大的顶层窗口元素；找不到返回 $null。
# 标题只用于日志，不参与定位；窗口最小化/离屏时返回 $null（配合 PP2.1 恢复逻辑）。
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

> 轮询（5-1）常以独立脚本进程运行：需连同本函数与 `Get-Process` 那两行一并复制，函数内不依赖任何外部状态。
> 新对话的生效校验仍可用 `MainWindowTitle`（只判断"标题是否回到「豆包」"，不用于窗口定位）。

#### PP2.3 新建对话并聚焦输入框

每次上传图片前必须处于**新对话**（避免把图片发到旧会话）。刚启动的豆包默认通常就是新对话；复用已运行实例时必须先新建。

**早失败探测**：`DOUBAO_READY` 后立即探测输入框（类名含 `ProseMirror`，T12；上限见 T17）。**探测不到 → 立即中止**并提示「未找到豆包输入框：可能未登录、界面异常或 accessibility 未生效，请人工检查后重试」，不得进入上传与轮询空等。
> 该报错**不等于**未登录：若界面明显已登录（能看到输入区与历史对话）却探测不到输入框，先按 PP2.2 重取窗口再试，仍失败则用 T2 参数重启豆包。

搜狗输入法的 `Ctrl+Shift+K` 冲突已解除，新对话优先用快捷键：

```powershell
# 新建对话快捷键：Ctrl+Shift+K（T8）
# 执行前确认已运行过 PP2.1 且 $doubao / Win32.Win32WinState 已定义；前台未确认时禁止 SendKeys。
Add-Type -AssemblyName System.Windows.Forms
if (-not ([Win32.Win32WinState]::GetForegroundWindow() -eq $doubao.MainWindowHandle)) {
    # 前台已丢失：先完整重跑 PP2.1（含 Add-Type、$doubao、$doubaoReady），确认 DOUBAO_READY 后再继续
}
[System.Windows.Forms.SendKeys]::SendWait('^+k')

# 提速（T15）：自适应等待标题回到「豆包」，替代固定 sleep。
# 实测标题回到「豆包」约 0.5–2.3 秒（含渲染排队），固定 sleep 要么白等要么不够。
$newChatMs = 4000
$swNewChat = [System.Diagnostics.Stopwatch]::StartNew()
$newChatOk = $false
while ($swNewChat.ElapsedMilliseconds -lt $newChatMs) {   # 上限 4 s
    Start-Sleep -Milliseconds 150
    if ((Get-Process -Id $doubao.Id).MainWindowTitle -match '^豆包$') { $newChatOk = $true; break }
}
"NEW_CHAT_OK=$newChatOk  用时 $($swNewChat.ElapsedMilliseconds) ms（标题=$((Get-Process -Id $doubao.Id).MainWindowTitle)）"
if (-not $newChatOk) { Write-Output '快捷键未生效（标题未变），转入 UIA 备用方式' }
```

快捷键无效时用 UI Automation 点击侧边栏「新对话」（点击前必须已确认窗口在屏幕内，否则 `BoundingRectangle` 为无效离屏坐标）：

```powershell
# UI Automation 备用方式：先移动光标再点击，带生效校验
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);' -Name U32Mouse -Namespace Win32

# 前置：窗口必须已恢复到屏幕内（PP2.1 已通过）
$win = Get-DoubaoWinByProcess $doubao
if (-not $win) { Write-Output '窗口定位失败，先执行 PP2.1 恢复激活代码块再重试'; exit 1 }
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
            Write-Output 'BoundingRectangle 无效（窗口可能未恢复），先执行 PP2.1 再重试'
        }
    }
    # 提速（T15）：自适应等待标题回到「豆包」，替代固定 2000ms
    $swUia = [System.Diagnostics.Stopwatch]::StartNew()
    $uiaOk = $false
    while ($swUia.ElapsedMilliseconds -lt $newChatMs) {   # 上限 4 s
        Start-Sleep -Milliseconds 150
        if ((Get-Process -Id $doubao.Id).MainWindowTitle -match '^豆包$') { $uiaOk = $true; break }
    }
    "UIA_NEW_CHAT_OK=$uiaOk  用时 $($swUia.ElapsedMilliseconds) ms"
    # 生效校验：窗口标题应变回「豆包」（新对话无标题）
    $doubao2 = Get-Process -Name Doubao -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
    if ($doubao2.MainWindowTitle -match '^豆包$') { 'NEW_CHAT_OK' }
    elseif ($clicked) { 'CLICKED_TITLE_UNCHANGED_NEED_VERIFY' }
} else {
    Write-Output '未找到“新对话”入口'
}
```

> 「每次识别都用新对话」是刻意设计：每次任务从干净上下文开始，不受旧对话影响，也保证 5-1/5-2 统计到的字段行只可能来自本批回复。

**新建对话能清什么、不能清什么**（2026-10 对照实验实测）：

| 残留类型 | 新建对话是否清除 | 依据 |
|---|---|---|
| 上一轮的回复行与结束标记 | ✅ **清除**（5-1 判据可用的前提） | 新建对话后主区域无旧字段行（`PURITY_CHECK oldReplyLines=0`） |
| **输入复合框内的待发附件**（已粘贴未发送的图片/文件） | ❌ **不清除** | 贴图后附件数 2 → 执行新建对话后**仍为 2**；附件挂在 `input-content-container` 下，属输入区状态，与对话无关 |
| 历史会话堆积 | ❌ 不管 | 由用户自行管理，skill 不自动删除任何对话 |

> **因此：待发附件必须由流程自己保证"不多不少"。** 每条消息发出后附件随消息消费掉；**不要在正式发送前做任何"试贴"**（含冒烟自检——0.7 已改为不要求贴图）。`Ctrl+A` + `Delete` 对这类附件同样无效，目前没有可靠的程序化清除手段。

> **关于删除对话**：侧栏对话项只有「归档对话」，**没有删除**。删除入口在两级菜单之后——侧栏「最近」标题右侧的 **对话列表设置 → 历史对话管理**，进入管理界面后每个对话行右侧各有一个操作菜单（界面内所有对话按时间分组，含"更早"）。**skill 不自动删除对话**：删除不可逆且属用户数据，应由用户手动决定；用户明确要求清理会话时，按其指示操作或在确认后再执行。

输入框是 `ProseMirror` 编辑器，通过 UI Automation 定位并聚焦：

```powershell
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$win = Get-DoubaoWinByProcess $doubao
if (-not $win) { Write-Error '窗口定位失败，先执行 PP2.1 恢复激活代码块'; return }

# 提速（T15）：自适应等待输入框出现，替代固定 sleep。
# 实测新对话后输入框就绪最慢约 5.8 秒（渲染排队），固定 sleep 必须留足余量→白等；
# 轮询则可"就绪即走"。单次全树扫描仅约 9 ms（200+ 元素），可用 150 ms 间隔。
$editWaitMs = 10000   # T17：输入框等待上限（10 秒）
$swEdit = [System.Diagnostics.Stopwatch]::StartNew()
$edit = $null
while ($swEdit.ElapsedMilliseconds -lt $editWaitMs) {   # T17 上限见常量表
    $win = Get-DoubaoWinByProcess $doubao
    if ($win) {
        $all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition)
        $edit = $all | Where-Object { $_.Current.ClassName -like '*ProseMirror*' } | Select-Object -First 1
        if ($edit) { break }
    }
    Start-Sleep -Milliseconds 150
}

if (-not $edit) {
    Write-Error '未找到豆包输入框（ProseMirror），请确认豆包已打开且已在输入框页面；仍失败则用 T2 参数重启豆包'
    return
}
"EDIT_READY 用时 $($swEdit.ElapsedMilliseconds) ms"

$edit.SetFocus()
Start-Sleep -Milliseconds 150   # 聚焦后极短稳定（原 300 ms 无必要）
# 校验焦点确实落在输入框（应对"命中错误元素导致粘贴落空"）
$focused = [System.Windows.Automation.AutomationElement]::FocusedElement
if ($focused.Current.ClassName -notlike '*ProseMirror*') {
    Write-Output "焦点未落在输入框（当前=$($focused.Current.ClassName)），重试 SetFocus"
    $edit.SetFocus(); Start-Sleep -Milliseconds 250
}
'EDIT_READY'
```

### P3 文本输入防护（对所有发送给豆包的文字生效）

- **禁止用 `SendKeys` 直接键入中文或长文本**（可能触发搜狗输入法/软键盘并产生乱码）；所有文字一律先入剪贴板，再 `Ctrl+V` 粘贴。
- 若搜狗软键盘或候选框弹出：先按 `Esc` 关闭，再继续；**不要在有中文候选词时按 Enter**。
- 发送消息优先点击豆包「发送」按钮；找不到发送按钮时才用 Enter（且先按 `Esc` 清候选）。

> 本防护作用于步骤4、步骤7-2 以及所有向输入框写文字的动作，不是步骤4 的局部规则。

---

## 3. 执行层：步骤0（PPT/PPTX 拆解渲染）

**触发条件**：1.2 分流判定为「问 PPT 页面内的图片内容」。

核心思路：**豆包只收图片，所以先把 PPT 变成图片**——把"PPT 里的图片内容是什么"转化为"这张图里有什么"。整页渲染同时转录每页文字，解决"只能读 PPT 文字、读不出 PPT 图片"的问题。

### 3.1 通道选择

| 通道 | 何时用 | 说明 |
|---|---|---|
| 整页渲染 PNG（主通道，默认必须） | 所有 PPT 画面识别任务 | 每页 1 张 PNG，一次识别同时拿到图片内容 + 版式布局 + 页面文字；依赖本机 Microsoft PowerPoint（桌面版）或 WPS 演示 |
| media 内嵌原图（细节通道，按需） | 整页图上局部图太小/模糊，需要原图追问 | .pptx 即 zip，图片在 `ppt/media/`，无需 Office 即可解出；⚠️ media 文件编号与页码**无直接映射**，不能据此判断图在哪一页 |
| 区域裁剪（细节通道） | 某页局部看不清 | 用 System.Drawing 从该页整页 PNG 裁剪目标区域另存后上传 |

### 3.2 渲染整页 PNG（主通道）

> 源文件位于回收站（`$RECYCLE.BIN`）或被占用导致 COM 直接打开失败（实测 E_FAIL）时，自动先复制到临时目录再打开重试；复制物属本次任务临时文件，随步骤8-1 清理，**原文件绝不删除**。

```powershell
$pptPath = 'D:\xxx\演示文稿.pptx'   # 已绝对路径化
$outDir = Join-Path 'D:\DS\.dsh' ('tmp_ppt_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# ① 优先 Microsoft PowerPoint COM；不可用则尝试 WPS 演示 COM（Kwpp.Application，接口相似）
$app = $null
try { $app = New-Object -ComObject PowerPoint.Application }
catch { try { $app = New-Object -ComObject Kwpp.Application } catch { } }
if (-not $app) {
    Write-Error '未安装 PowerPoint/WPS，无法整页渲染；.pptx 可退化为仅提取 media 原图（见 1.3）'
    exit 1
}
try {
    # Presentations.Open(路径, ReadOnly, Untitled, WithWindow=false)
    # 直接打开失败（回收站/占用等，E_FAIL）→ 复制到临时目录重试一次
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

### 3.3 提取 media 内嵌原图（细节通道，.pptx 适用，默认跳过）

> **默认不执行**：正常拆解任务不需要 media，解压耗时且占磁盘。仅当用户追问细节、整页图上局部图太小/模糊、需要上传原图时才临时执行。旧版 .ppt 无法解包，局部细节一律走渲染图裁剪路径。

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

### 3.4 识别组织规则

1. **默认只上传整页渲染 PNG**，按页码顺序分批（每批 ≤ 10 页，T4）；页数多时先问用户关注范围，或按顺序分批识别、分批汇总。
2. 每页结果按「页码」组织输出，附该页整页 PNG 链接（如 `[第03页](file:///D:/.../slide_003.png)`），**并注明该临时文件在任务收尾清理后失效**；用户需要长期可点击的页图，先把 slide PNG 转存到用户指定目录（如原 PPT 同目录的 `_slides\` 子目录）再给链接（见步骤8-1）。
3. 用户问"某页的图片是什么/细节看不清"：先对该页渲染图追问；仍不清再上传该页区域的**裁剪图**或 **media 原图**单独识别（3.3），按步骤7 追问。
4. PPT 页面文字随整页识别由「画面文字」字段一并转录；用户要"PPT 全文文字稿"时，在整页识别后追问豆包逐页完整转录，模糊处标「[?]」，不臆测。
5. 旧版 .ppt（非 .pptx）无法解包提取 media，只能整页渲染；局部细节走渲染图裁剪/放大路径。
6. 步骤0 产物目录属**本次任务临时文件**：收尾（含失败/中断）必须删除，禁止为了保留可点击链接而把产物留在磁盘上；绝不删除用户原始 PPT。

---

## 4. 执行层：步骤0.6 PPT 内容分析轻量模式（只问文本内容时直传文件本体）

**触发条件**：用户只问 PPT 的"主要内容 / 讲了什么 / 总结 / 大纲 / 观点"，**不涉及页面内图片内容**。此时不渲染、不拆页，直接把 .ppt/.pptx 文件本体当附件发给豆包。

**边界**：一旦问到"某页的图片/照片/配图/图标是什么"→ 立即转步骤0 整页渲染识别模式（可配合 media 原图追问细节）。

### 4.1 执行流程

1. **PP2.1 恢复激活**（确认 `DOUBAO_READY`）；
2. **PP2.3 新建对话**并聚焦输入框；
3. **上传 PPT 文件附件**：走下面的「附件上传」代码块（剪贴板文件列表）。上传校验上限见 T18；校验失败重试一次。
   源文件位于回收站或被占用导致附件粘贴失败 → 先复制到 `D:\DS\.dsh\tmp_ppt_<hex>\src\`，上传副本（随步骤8-1 清理，原文件不动）。
4. **完整发送【附录A.2】固定「PPT 文本分析指令」**（不允许删减修改；发送方式同步骤4，含"第 5 步：确认已发出"）；
5. **轮询完成判定**：该回复是长篇文本分析，**不是图片的 8+5 字段结构、也不要求结束标记**——5-1 的图片字段判据不适用。用下面的「轮询」代码块：窗口文本总量超过基线（增量阈值 T16）且连续 T6 稳定窗口不再增长 → 判完成；超时上限 T5 的 PPT 档（240 秒）；
6. **读取回复（健壮转储，防脚本中断）**：用下面的「健壮转储」代码块。转储必须**逐元素容错**——实测豆包窗口个别元素 `BoundingRectangle` 返回无穷值（∞），直接 `[int]$r.Y` 强转抛异常，在 `$ErrorActionPreference='Stop'` 下会中断整段读取、结果丢失：

**① 附件上传（PPT 文件本体，剪贴板文件列表）**

```powershell
# PPT 附件必须走文件列表（SetFileDropList），不能像图片那样用 SetImage。
# ⚠️ 必须显式 -ReferencedAssemblies：Add-Type 编译内联 C# 时不会自动引用 WinForms，
#    漏了直接报「命名空间 System.Windows 中不存在类型 Forms」（2026-10 实测）。
#    实测确认："-ReferencedAssemblies System.Windows.Forms"（短名）即可编译并调用成功。
#    若该写法在个别环境仍失败，用块末的纯反射回退方案（无需编译 C#）。
Add-Type -AssemblyName System.Windows.Forms
Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition @'
using System;
using System.Collections.Specialized;
using System.Windows.Forms;
public static class ClipFiles {
    public static void Set(string path) {
        var col = new StringCollection();
        col.Add(path);
        var data = new DataObject();
        data.SetFileDropList(col);
        Clipboard.SetDataObject(data, true);
    }
}
'@ -ErrorAction Stop

$pptAbsPath = (Resolve-Path -LiteralPath $pptPath).Path   # 必须绝对路径
ClipFiles::Set($pptAbsPath)

$edit.SetFocus()
Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait('^v')   # 粘贴文件 → 生成附件
"PPT_ATTACH_PASTED  file=$([System.IO.Path]::GetFileName($pptAbsPath))"
```

> **回退方案（内联 C# 编译失败时用）**：改为纯反射，不编译任何 C#。注意反射传参必须用 `(, $col)` 包裹，否则会报 `String cannot be converted to StringCollection`（实测踩过）：
>
> ```powershell
> $col = New-Object System.Collections.Specialized.StringCollection
> [void]$col.Add($pptAbsPath)
> $do = New-Object System.Windows.Forms.DataObject
> ([System.Windows.Forms.DataObject].GetMethod('SetFileDropList')).Invoke($do, (, $col))
> [System.Windows.Forms.Clipboard]::SetDataObject($do, $true)
> ```

**② 附件就位校验（自适应，上限 T18 = 20 秒）**

> **实测结论（2026-10，6 MB .pptx）**：文件名以独立 `Text` 元素暴露（`name=[传奇一生.pptx]`），同时伴随 `上传中... N%` 与 `6MB` 进度元素。因此判据必须**同时要求"文件名出现"且"无上传中字样"**——只看文件名会在上传未完成时假通过。上传完成后进度元素消失，此判据稳定。若某版本不暴露文件名，则以界面为准继续，并回填实测特征。

```powershell
$attachWaitMs = 20000   # T18
$fileName = [System.IO.Path]::GetFileName($pptPath)
$fileExt  = [System.IO.Path]::GetExtension($pptPath)   # .pptx / .ppt
$swPpt = [System.Diagnostics.Stopwatch]::StartNew()
$pptAttachOk = $false; $progText = ''
while ($swPpt.ElapsedMilliseconds -lt $attachWaitMs) {
    Start-Sleep -Milliseconds 250
    $winP = Get-DoubaoWinByProcess $doubao
    if (-not $winP) { continue }
    $allP = $winP.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    # ⚠️ 关键：不能只看文件名——文件名元素在上传开始前就出现。
    #    实测 6MB 文件在"上传中…0%"阶段就命中了文件名，会在上传未完成时误判就位（325 ms 假通过）。
    #    必须同时满足"文件名出现"且"窗口内无上传中/上传失败字样"。
    $nameSeen = $false; $uploading = $false
    foreach ($e in $allP) {
        $n = $e.Current.Name
        if (-not $n) { continue }
        if ($n -like "*$fileName*" -or $n -like "*$fileExt*") { $nameSeen = $true }
        if ($n -match '上传中|上传失败|Uploading') { $uploading = $true; $progText = $n }
    }
    if ($nameSeen -and -not $uploading) { $pptAttachOk = $true; break }
}
"PPT_ATTACH_OK=$pptAttachOk  用时 $($swPpt.ElapsedMilliseconds) ms"
if (-not $pptAttachOk) { Write-Output '⚠ 附件校验未命中：若界面已显示附件则继续；否则按 3.1 步骤3 复制副本后重试' }
```

**③ 轮询完成判定（基线在本进程内采集 + T16 增量 + T6 稳定窗口）**

```powershell
# 与图片链路同一判据：只看结束标记（PPT 指令已同步要求输出该标记，见【附录A.2】四）。
# 这样两条链路的完成判定同构，不再各写一套（原 PPT 侧用"文本总量稳定"，其判据自身要求
# 整窗文本相等，会被界面噪声永久扰动而永不成立——已废弃）。
$markerPattern = '识别完毕'

# 基线：本进程内采集，等标记计数稳定
$script:markerBase = 0
$markerStableMs = 8000   # T15：基线稳定判定上限
$swMB = [System.Diagnostics.Stopwatch]::StartNew()
$prevM = -1; $stableM = 0
while ($swMB.ElapsedMilliseconds -lt $markerStableMs) {
    Start-Sleep -Milliseconds 250
    $wm = Get-DoubaoWinByProcess $doubao
    if (-not $wm) { continue }
    $am = $wm.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    $m = 0
    foreach ($el in $am) {
        $nm = $el.Current.Name
        if ($nm -and (($nm -replace '[、，,．.·\s]', '') -match $markerPattern)) { $m++ }
    }
    $script:markerBase = $m
    if ($m -eq $prevM) { $stableM++ } else { $stableM = 0 }
    if ($stableM -ge 2) { break }
    $prevM = $m
}
"PPT_MARKER_BASE=$script:markerBase  用时 $($swMB.ElapsedMilliseconds) ms"

$deadline = (Get-Date).AddSeconds(240)   # T5 PPT 档（整份 PPT 分析较慢）
$markerSince = $null; $done = $false; $missCount = 0
$swP = [System.Diagnostics.Stopwatch]::StartNew()
do {
    Start-Sleep -Milliseconds 500   # T6
    $w = Get-DoubaoWinByProcess $doubao
    if (-not $w) {
        $missCount++
        if ($missCount -ge 4) {
            Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name Win32Restore2 -Namespace Win32 -ErrorAction SilentlyContinue
            if ([Win32.Win32Restore2]::IsIconic($doubao.MainWindowHandle)) {
                [Win32.Win32Restore2]::ShowWindow($doubao.MainWindowHandle, 9) | Out-Null
                Start-Sleep -Milliseconds 800
            }
            $missCount = 0
        }
        continue
    }
    $a = $w.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    $markerCount = 0
    foreach ($el in $a) {
        $nm = $el.Current.Name
        if ($nm -and (($nm -replace '[、，,．.·\s]', '') -match $markerPattern)) { $markerCount++ }
    }
    if ($markerCount -gt $script:markerBase) {
        if ($null -eq $markerSince) { $markerSince = Get-Date }
        elseif (((Get-Date) - $markerSince).TotalSeconds -ge 1.5) { $done = $true; break }
    }
} until ((Get-Date) -gt $deadline)

if ($done) {
    "PPT_DONE_BY_MARKER  用时 $([int]$swP.Elapsed.TotalSeconds) s"
    # ★ 同图片链路：抢读落盘 → 立即最小化（最小化后 UIA 读不到，必须先读）
    $grabP = @()
    foreach ($el in $a) {
        $nmP = $el.Current.Name
        if (-not $nmP) { continue }
        try { $rP = $el.Current.BoundingRectangle; $yP = [int]$rP.Y } catch { $yP = -1 }
        $grabP += [pscustomobject]@{ Y = $yP; Name = $nmP }
    }
    $pptGrab = @($grabP | Sort-Object Y -Unique | ForEach-Object { $_.Name })
    "PPT_GRAB_LINES=$($pptGrab.Count)   （抢读完成，窗口尚未最小化）"
}
else {
    Write-Output ("PPT_POLL_TIMEOUT 已等 $([int]$swP.Elapsed.TotalSeconds) s  标记计数=$markerCount 基线=$script:markerBase")
    Write-Output '→ 结束标记未出现。若界面已分析完，按 5-4「仅重抓」直接读取，不要重传重发'
}

# ★★★ 无条件最小化（同图片链路：收尾动作不放进判据分支）
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name Win32MinNowP -Namespace Win32 -ErrorAction SilentlyContinue
[Win32.Win32MinNowP]::ShowWindow($doubao.MainWindowHandle, 6) | Out-Null
'MINIMIZED_IMMEDIATELY'
 }
```

**④ 健壮转储（逐元素容错，防 ∞ 中断）**

```powershell
# 作用域限定：只看"最新一条回复的消息块"，避免把用户消息与上一轮回复一起转储
#   （与 5-2 同一思路：以结束标记为锚点，向上取覆盖首字段行且高度足够的最小容器）
$mkEl = $null; $mkY = -1
foreach ($el in $all) {
    $nm0 = $el.Current.Name
    if (-not $nm0) { continue }
    if (($nm0 -replace '[、，,．.·\s]', '') -notmatch '识别完毕') { continue }
    try { $r0 = $el.Current.BoundingRectangle } catch { continue }
    if ([double]::IsNaN($r0.Y) -or [double]::IsInfinity($r0.Y)) { continue }
    if ($r0.Y -gt $mkY) { $mkY = $r0.Y; $mkEl = $el }
}
$scopeEl = $null
if ($mkEl) {
    $wkr = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $nd = $wkr.GetParent($mkEl)
    for ($k = 0; $k -lt 8 -and $nd; $k++) {
        try { $rb0 = $nd.Current.BoundingRectangle } catch { $rb0 = $null }
        if ($rb0 -and -not [double]::IsNaN($rb0.Y) -and $rb0.Height -gt 200) { $scopeEl = $nd; break }
        $nd = $wkr.GetParent($nd)
    }
}
if (-not $scopeEl) { $scopeEl = $win; Write-Output 'PPT_SCOPE_WARN 未定位消息块，退化为全窗口转储' }
$scopeAll = $scopeEl.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)

# 每个元素独立 try/catch；几何值非有限时置 -1；单元素异常绝不中断整段
$lines = New-Object System.Collections.Generic.List[string]
foreach ($el in $scopeAll) {
    try {
        $name = $el.Current.Name
        if (-not $name) { continue }
        $r = $el.Current.BoundingRectangle
        $yStr = '-1'
        if (-not [double]::IsNaN($r.Y) -and -not [double]::IsInfinity($r.Y)) { $yStr = [string][int]$r.Y }
        $lines.Add("$yStr`t" + ($name -replace "`r?`n", '⏎'))
    } catch { continue }   # 单元素异常只跳过，不中断整段读取
}
"PPT_DUMP_LINES=$($lines.Count)"
```

   滚动（滚轮/PageDown/End）只作增强手段：实测 Chromium UIA 不滚动也常能暴露完整回复文本树；滚轮 `mouse_event` 的 Add-Type 失败属正常降级，不能因此中断读取；滚动动作同样各自包 try/catch。
   转储落盘后校验是否含「PPT类型判定结果」等框架锚点（**匹配前先去空白**，UIA 会在 ASCII 词两侧插空格，实测形如「PPT 类型判定结果」）。**窗口已在完成判据处（本步骤 ③ 的抢读之后）自动最小化**，此处不再涉及最小化时机；若校验失败，走 5-4 的「仅重抓」**（只读取不发送，上限见 T19）**——**必须先恢复窗口再读取**（实测最小化态 UIA 读不到任何元素），禁止为补结果而重新上传重发。
7. **结果输出**：
   - 先转述豆包的【PPT类型判定结果】与对应框架分析内容；
   - **每条内容附来源页码标注**（豆包按指令输出〔第X页〕/〔第X–Y页〕，如实转述；豆包标〔页码不明〕或未标注时如实说明，禁止编造页码；确需按 PPT 页序与版式推断时注明「推断」）；
   - **层级标题一律用中文序号（一、二、三…）**，不使用 ABC/字母/长串编号；
   - **排版细则**：①「分章节内容拆解」按页码 1→N **逐页呈现**，每页独立一行/一段（推荐表格：页码｜板块名｜核心内容），严禁多页混段；②「核心重点汇总」「结构评价」等板块内部条目按页码从小到大排序、**每条独立成行**（每条以〔第X页〕开头另起一行，可用项目符号列表；跨页范围按起始页参与排序），同页多条按原顺序排列，禁止用分号/顿号把多条挤在同一段；
   - 通用型汇报的「核心重点汇总」必须保留条目级细节（核心主张 / 关键事实与数据逐条 / 金句原文 / 行动与占位信息，每条带〔页码〕），不得概括成空话；
   - 遵守步骤6 规则（原文与解读区分、可点击链接仅用于源文件本身）；
   - 收尾清理按步骤8-1 执行（本模式几乎不产生临时产物，无渲染目录残留）。

---

## 5. 执行层：图片识别链路

### 步骤1 定位并启动/复用豆包

由 PP2.0–PP2.2 覆盖：进程探测 → 启动（仅用 T1 本体 + T2 参数）→ 窗口恢复激活闸门 → 通用窗口定位函数。执行时直接运行 PP2.0/PP2.1/PP2.2 代码块。

### 步骤2 确保新对话并聚焦输入框

由 PP2.3 覆盖：`Ctrl+Shift+K` 新建（标题回到「豆包」为生效判据）→ 失败转 UIA 备用方式 → 聚焦 ProseMirror 输入框。

### 步骤3 上传图片与批控

#### 步骤3-1 单张图片

剪贴板粘贴最快：

```powershell
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$imagePath = 'D:\CSY\example.jpg'
$img = [System.Drawing.Image]::FromFile($imagePath)
[System.Windows.Forms.Clipboard]::SetImage($img)
$img.Dispose()

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::SendWait('^v')

# 提速（T15）：自适应等待单张附图入框，替代固定 1200ms；顺带完成"已入框"确认
$attachWaitMs = 3000
$swOne = [System.Diagnostics.Stopwatch]::StartNew()
$oneOk = $false
while ($swOne.ElapsedMilliseconds -lt $attachWaitMs) {   # 上限 3 s
    Start-Sleep -Milliseconds 200
    $w1 = Get-DoubaoWinByProcess $doubao
    if (-not $w1) { continue }
    $a1 = $w1.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    # 附件挂在"输入复合框"（input-content-container）下，不在 ProseMirror 子树内，
    # 因此按该复合框计数；找不到就退化为窗口级（宁可多算，不可漏算）。
    $scope1 = $null
    $wk = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $nn = $wk.GetParent($edit)
    for ($k = 0; $k -lt 6 -and $nn; $k++) {
        if ($nn.Current.ClassName -like '*input-content-container*') { $scope1 = $nn; break }
        $nn = $wk.GetParent($nn)
    }
    if (-not $scope1) { $scope1 = $w1 }
    $inBoxN = @($scope1.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition) |
        Where-Object { $_.Current.ClassName -match 'image-container|image-wrapper' }).Count
    if ($inBoxN -ge 1) { $oneOk = $true; break }
}
"SINGLE_ATTACH_OK=$oneOk  用时 $($swOne.ElapsedMilliseconds) ms"
```

#### 步骤3-2 多张图片（单批 ≤10 张）

按文件名自然排序后逐张粘贴到同一输入框。**单批（同一次发送）最多 10 张（T4）**——本批只贴前 10 张，其余走 3.3 分批规则（每批独立新对话）。

```powershell
$files = Get-ChildItem -LiteralPath 'D:\CSY' -File |
    Where-Object { $_.Extension -in '.jpg','.jpeg','.png','.bmp','.webp' } |   # T13 扩展名白名单
    Sort-Object Name

# T4 单批上限 10 张：只取前 10 张作为本批；$files.Count -gt 10 时，
# 剩余图片必须在后续批次处理（每批独立新对话），禁止一次粘贴超过 10 张。
$batch = @($files | Select-Object -First 10)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 连续粘贴提速：间隔 450ms（T7）；不逐张贴完就全窗口扫描，整批贴完后由 3.3 统一校验一次；
# 校验不足时按“尾部缺失补贴一次”补齐，仍不齐才中止重开新对话。
$pasted = 0
foreach ($file in $batch) {
    # ★ 每张粘贴前复检窗口：批中途窗口被最小化会让 SetFocus/^v 静默失效（附件数停在旧值），
    #   事后再看校验结果就会误判成"粘贴太快"。此处当场发现、当场报出，不留给校验去猜。
    Add-Type -MemberDefinition
'
[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
'
 -Name Win32PerPaste -Namespace Win32 -ErrorAction SilentlyContinue
    if ([Win32.Win32PerPaste]::IsIconic($doubao.MainWindowHandle)) {
        [Win32.Win32PerPaste]::ShowWindow($doubao.MainWindowHandle, 9) | Out-Null   # SW_RESTORE
        Start-Sleep -Milliseconds 800
        [Win32.Win32PerPaste]::SetForegroundWindow($doubao.MainWindowHandle) | Out-Null
        Start-Sleep -Milliseconds 400
        "WARN_WINDOW_ICONIC_RESTORED 第 $($pasted + 1) 张前发现窗口被最小化，已恢复后继续"
    }
    $edit.SetFocus()
    Start-Sleep -Milliseconds 150

    $img = [System.Drawing.Image]::FromFile($file.FullName)
    [System.Windows.Forms.Clipboard]::SetImage($img)
    $img.Dispose()

    [System.Windows.Forms.SendKeys]::SendWait(
'
^v
'
)
    Start-Sleep -Milliseconds 450   # T7
    $pasted++
}
"本批粘贴 $($batch.Count) 张（上限 10 张）；剩余 $([math]::Max(0, $files.Count - 10)) 张待后续批次"
```

> 若当前豆包版本支持一次多选文件上传，也可走"附件按钮 → 上传文件或图片 → 多选 → 打开"进一步提速；同样受 T4 约束（一次最多选 10 张）。该方式不稳定时回到逐张粘贴。

#### 步骤3-3 附件数校验、会话纯净校验与分批规则

上传完成后做两项校验，**任一不通过立即中止并重新执行 PP2.3（新对话）**，禁止带病发送。

1. **附件数校验**：输入区已就位图片数必须等于**本批张数**（≤10，基准见 1.3 任务台账）；不足时按**"尾部缺失补贴一次"**补齐（从本批末尾取缺失数量的文件重贴），补齐后仍不足 → 中止并重跑 PP2.3；
2. **会话纯净校验**：当前对话主区域不得出现本任务之前的旧消息/旧图片痕迹（旧的结构化回复行、旧的「图N」编号、旧会话标题仍在标题栏等）——出现即说明"新对话"未真正生效，继续发送会把本次图片混进旧会话。

```powershell
# 提速（T15）：自适应等待附件渲染，替代"固定 1200ms 复测"。
# 附件缩略图生成有延迟，用轮询"够了就走"：典型省 0.5–1.5 秒/批，且比固定值更不容易误判。
$attachWaitMs = 3000
# 作用域 = 输入复合框（input-content-container）——附件挂在它下面，不在 ProseMirror 子树内：
#   · 全窗口计数 → 会把别处同类名缩略图算进来（误判"已入框"）
#   · 只扫 ProseMirror 子树 → 数出 0，漏掉真附件
$imgScope = $null
$wf = [System.Windows.Automation.TreeWalker]::ControlViewWalker
$nf = $wf.GetParent($edit)
for ($k = 0; $k -lt 6 -and $nf; $k++) {
    if ($nf.Current.ClassName -like '*input-content-container*') { $imgScope = $nf; break }
    $nf = $wf.GetParent($nf)
}
if (-not $imgScope) { $imgScope = $win }   # 兜底：退化为窗口级（宁可多算，不可漏算）
"IMG_SCOPE=$($imgScope.Current.ClassName.Substring(0,[math]::Min(40,$imgScope.Current.ClassName.Length)))"

$swAttach = [System.Diagnostics.Stopwatch]::StartNew()
$imageCount = 0
while ($swAttach.ElapsedMilliseconds -lt $attachWaitMs) {   # 上限 3 s
    # T11：同时覆盖 image-container-*（旧版）与 image-wrapper-*（现版本），避免"已入框却数出 0"而重复粘贴
    $imageCount = @($imgScope.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition) |
        Where-Object { $_.Current.ClassName -match 'image-container|image-wrapper' }).Count
    if ($imageCount -ge $batch.Count) { break }
    Start-Sleep -Milliseconds 200
}
Write-Output ("ATTACH_WAIT 用时 $($swAttach.ElapsedMilliseconds) ms  计数=$imageCount/$($batch.Count)")

if ($imageCount -lt $batch.Count) {
    # 尾部缺失补贴：缺失数 = $batch.Count - $imageCount，
    # 取本批末尾对应数量的文件重贴一次（SetImage + ^v，间隔 450ms），再复检一次
    $miss = $batch.Count - $imageCount
    $refill = @($batch | Select-Object -Last $miss)
    foreach ($file in $refill) {
        $img = [System.Drawing.Image]::FromFile($file.FullName)
        [System.Windows.Forms.Clipboard]::SetImage($img)
        $img.Dispose()
        [System.Windows.Forms.SendKeys]::SendWait('^v')
        Start-Sleep -Milliseconds 450
    }
    # 复检仍不足 → 中止，重新执行 PP2.3
}
# 纯净性检查：主区域是否残留旧的结构化回复行（“字段名：实际内容”形态，提示词模板不会被命中）
$oldReply = @($all | Where-Object {
    $_.Current.Name -match '^(图片类型|核心主体|空间布局|场景与背景|画面文字|细节特征|特殊元素|不确定内容说明|推断对象|置信度与理由|判断依据|其他可能|无法推断项说明)\s*：\s*(?![（(])\S'
})
if ($oldReply.Count -gt 0) {
    Write-Output "检测到 $($oldReply.Count) 行旧会话回复，新对话未生效 → 中止，重新执行 PP2.3"
}
```

**分批规则**（总数 >10 时）：

1. 按文件名自然排序切批：第 1 批 = 前 10 张，第 2 批 = 次 10 张……（PPT 场景按页码顺序切批，每批 ≤10 页）；
2. **每批独立走完整识别链路，批与批之间必须新建对话**：
   `新对话 → 上传本批（≤10 张）→ 本步双校验（附件数 == 本批张数、无旧回复痕迹）→ 步骤4 发送固定指令 → 步骤5 轮询读取 → 记录本批结果`
   - 原因：旧批的回复行（每图 13 字段行）与结束标记若留在同一窗口，会命中 5-1 的完成判据，导致下一批刚发送即误判完成、结果被污染；新对话保证 5-1/5-2 统计到的只可能是本批回复。
3. 每批读取结果后立即记录/落盘（写入临时汇总文件，收尾清理）；全部批次完成后按编号顺序合并输出；
4. 单批内附件数不足（本步校验失败）→ 补齐缺失图片后重试一次；仍失败则记录失败项，不阻塞本批发送；
5. 附件数校验的基准是**本批**张数（≤10），不是任务总数。

### 步骤4 发送识别指令（固定骨架 + 可编辑槽位）

指令分两层，**改哪层是硬边界，先看清再动手**：

| 层 | 内容 | 可编辑性 |
|---|---|---|
| **不可改（判据锚点）** | 两部分结构、13 个字段名（8 客观 + 5 推断）、缺省值、字段行格式、多图编号规则、结束标记 | **禁止改动**——T9/T10 正则、5-1 完成判据、5-3 校验全靠它们。改了就会静默漏字段或提前判完成 |
| **可编辑（任务适配）** | 第二部分各字段的**描述文字**，以及任务相关的推断侧重 | **按当前任务生成**——见《附录A.1a 可编辑槽位》 |

执行顺序：

0. **先看用户点明了什么角度（最高优先级）**：用户在本次请求里说明的侧重点、想要的角度、答案形态，**优先于下方对照表的默认填法与任何自行判断**（见 1.2 优先级表 P1）。用户点明时按用户说的填槽；未点明才按《A.1a》对照表与任务需要填。用户角度与对照表冲突时，**以用户角度为准**，不要"顺手改回"自己认为更合适的方向。
1. **定时机**：按《A.1a》的对照表，依本次任务（用户问题 + 工作需要）填好各槽位。**能判断就不要留默认值**：这是"看得准"与"答得有用"的分界。
2. **组装**：下面代码块用**插值 here-string**（`@"…"@`）把槽位填进骨架，生成 `$prompt`。
3. 按 P3 文本输入防护把 `$prompt` 放入剪贴板并粘贴到输入框；
4. 发送，并**确认已发出**（本步最后一段代码轮询输入框清空）。
5. 转步骤5-1 开始轮询——**标记基线由 5-1 在自己的进程内采集**，本步不采（跨进程传不过去）。

> 骨架正文中不含 `$` 与反引号（已校验），因此可安全使用插值 here-string。若日后往骨架里加文本，**含 `$` 时必须写成 `` `$ ``**，否则会被 PowerShell 当变量吃掉。

```powershell
# ================= 可编辑槽位：按【附录A.1a】对照表填写本次任务的值 =================
# 填槽优先级（1.2 优先级表）：用户在请求中点明的角度 > 本次工作需要 > 对照表默认值。
# 用户一旦指定角度/侧重/答案形态，按用户说的填，不得改回默认或自行判断的方向。
$taskSummary    = '用户在询问图中人物可能是谁'          # 本次任务一句话概括（会写进指令）
$userQuestion   = '告诉我这两张图片里的人物'            # 用户原始问题；无明确问题时填与任务相符的表述
$inferenceScope = '人物/角色身份（作品出处、角色名、同人/原创）'   # 推断内容与侧重，可写多类
$focusHint      = '发型与发色、瞳色、服饰配色与纹样、标志性配饰、背景元素'  # 希望优先依赖的可观察线索
$basisHint      = '发型发色；瞳色；服饰配色与纹样；标志性配饰；背景元素'    # 判断依据里优先列举的线索
$answerForm     = '先给最可能的角色名与作品名，再按可能性排列其他候选'      # 期望的答案形态与粒度
$depth          = '标准'    # 简/标准/详细：分别对应每字段 1 句 / 1-2 句 / 2-3 句或分条
$extraNote      = '无'      # 本次特殊要求；无则填「无」
# ================================================================================

# 组装指令（骨架不可改；槽位填入上方的变量）
$prompt = @"
你负责识别图片，仅输出纯识别结果，供下游AI系统调用处理。严格遵守以下所有规则，禁止输出任何规则解释、寒暄、确认语、思考过程或与识别内容无关的文字。

【本次任务参数（由调用方给出，请据此调整识别与推算的侧重）】
- 任务说明：$taskSummary
- 用户问题：$userQuestion
- 推断内容与侧重：$inferenceScope
- 描述时请重点关注的可观察线索：$focusHint
- 判断依据中请优先列举以下线索：$basisHint
- 期望的答案形态：$answerForm
- 回答详细程度：$depth（简=每字段1句；标准=每字段1-2句；详细=每字段2-3句或分条）
- 其他要求：$extraNote

【核心强制规则】
1. 两部分严格分离：第一部分只描述图片中视觉可见的客观内容；第二部分才允许对图片内容做推想与预测。两部分的结论不得互相混入。
2. 客观描述零臆测：第一部分禁止推测、脑补、解读、美化、评价，禁止补充画面外信息。
3. 推断必须先给结论、再给依据：第二部分要以明确判断为主，每条结论都基于第一部分已列出的可观察特征；不得因为“只能推断”而回避作答，仅当确实无法判断时才填「无法判断」。
4. 全要素覆盖：必须完整识别主体、场景、文字、细节、特殊元素，不得遗漏可清晰辨识的关键信息。
5. 禁止遗漏个体（硬性）：凡画面中可辨识的个体/对象/条目，一律逐个列出，不得只列主要的前几个、把其余用“等”“等等”“其他”“若干”“众多”“大量”“其余人物”“背景人物”一句话带过。
   - 每个人物/角色/动物/物体都要单独列出，并给出可定位的标识：优先用名称；无法确定名称时用位置（如“左侧第一人”“前排右二”）或外观代称（如“白发少年”“蓝发双马尾少女”）。
   - 密集群像/远景小人无法逐一细述时：必须给出可计数的信息（数量或可数范围）+ 分布位置 + 主要个体的分别描述。
   - 合格/不合格示例（照此标准输出，不要只做近似表述）：
     × 不合格：“其余散布大量大小不一角色”（既无数量也无分布）
     × 不合格：“其余角色包含金发、黑发、红发等，服饰有和服、西装、洋装等多种样式”（用“等/多种”代替逐个列举）
     √ 合格：“其余可辨识者约 8 人：左侧 3 人（粉发撑红伞女性、金发蓝眼女性、黑发男性）、右侧 2 人（棕发男性、白发女性）、上方 3 人（紫发女性、红发男性、黑发少年）”
     √ 合格（确实无法细述时）：“画面外围另有约 15–20 个极小尺寸人影，面部不可辨，无法逐一描述；分布为上方约 6 个、左右两侧各约 4 个、底部约 3 个”
   - 该要求对“核心主体”“细节特征”“画面文字”“特殊元素”同样适用；第二部分“推断对象”若涉及多个个体，也必须逐个给出结论。
6. 模糊必标原则：
   - 单个文字/细节无法辨认：用「[?]」标注
   - 局部区域内容模糊、遮挡、虚化：标注「[该区域内容模糊，无法准确识别]」
   - 整段/整块内容完全无法辨识：标注「[内容模糊，无法准确识别]」
   - 部分可见的文字仅保留可辨识部分，不可辨识字符用「[?]」补位；禁止对模糊内容进行任何猜测性描述。
7. 纯结果输出：禁止出现“我识别到”“这是一张”“图片中”等冗余前缀，直接按格式填充内容。
8. 字段名固定：两部分必须严格使用下方给定的字段名，不得改名、增删或调整顺序。

【输出格式】
每张图片都必须依次输出「第一部分」和「第二部分」两块，两块均不可省略。两部分均严格按下列字段输出，无内容则填该项标注的缺省值：
第一部分 客观描述（只写看得见的内容）
- 图片类型：（如实拍照片、截图、插画、海报、证件、图表、漫画等；若为示意图/流程图/电路图/公式推导图等，如实标明具体类型）
- 核心主体：（画面最主要的1-3个对象，含类别、数量、大致形态；场景复杂时优先提取占比最大、视觉最突出的对象；若为纯风景、纯现象或纯图表而无主体，如实描述画面主体构成并填「无明确主体」。**若画面中有多个人物/角色/动物/物体，必须逐个列出并各自标注位置或外观代称，不得只列前几个再用“等/其他/其余人物”带过**）
- 空间布局：（主体间的相对位置、前后/左右/上下关系、各自主占画面比例）
- 场景与背景：（整体环境类型、背景元素、光线、色调）
- 画面文字：（逐字转录所有清晰可见文字，按从左到右、从上到下的阅读顺序排列；标注文字所在位置，如“左上角”“画面中部”；无文字填「无」）
- 细节特征：（物体颜色、材质、姿态、动作、服饰等可辨识的精细属性；示意图类请描述标注、连线、符号、刻度、量纲等构成要素。**画面中每个已列出的个体都要有对应的特征描述，不得只详述其中一两个而对其余略过**）
- 特殊元素：（二维码、条形码、logo、图标、符号、印章、图表等标志性元素，仅描述外观，不解读含义）
- 不确定内容说明：（列出所有模糊、遮挡、反光、低分辨率导致无法识别的区域/内容；无不确定填「无」）

第二部分 推想预测（在第一部分事实基础上给出明确结论，并附可核对的依据；允许使用常识与领域知识）

先给结论，再给依据。图片识别的推测通常有较高可靠性，不要因为“这是推断”就回避作答。
- 推断对象：（第一行直接给出你的判断结论，结论类型以【本次任务参数】的“推断内容与侧重”为准。例：人物/角色→角色名与作品名；动物/植物→物种类别与具体物种名；作品/产物→作品名或产品名；地点/建筑→地点名或建筑名；景物/自然→景观或地理名称；现象/原理→所反映的物理规律或科学原理；文字/标识→含义与来源。**若画面中有多个可辨识个体，必须逐个给出结论，不得只答前几个**；有多个候选时按可能性从高到低排列，用「>」分隔。只有确实无法判断时才填「无法判断」）
- 置信度与理由：（先给「高」或「中」或「低」，再用一句话说明为什么给这个等级。口径：高=有明确可识别的特征或标志性线索支撑；中=依据充分但存在其他解释；低=仅能依据风格或概率推断）
- 判断依据：（逐条列出第一部分中支撑该判断的具体可观察特征，用「；」分隔，至少1条；不得只写“整体看像”这类空话）
- 其他可能：（列出其他合理但可能性较低的答案，用「；」分隔；确实没有其他合理解释则填「无」）
- 无法推断项说明：（只列出你确实无法判断的内容，例如具体身份、物种、真实地点、拍摄时间、创作者、材质等；能判断的不要写在这里；无则填「无」）

【多图与异常处理】
1. 单张图片直接按上述格式输出，无需编号。
2. 多张图片按上传顺序依次编号为「图1」「图2」「图3」……，每张分别独立输出完整的「第一部分」与「第二部分」。
3. 若图片包含违规敏感内容、或完全无法识别（全黑/全白/纯噪点），仅输出：无法识别该图片

【结束标记】
1. 上述所有图片的两个部分全部输出完毕后，在回复末尾另起一行，输出由四个汉字连续书写（紧挨连写、不加空格、不加任何标点）组成的结束标记；四个汉字依次为：第一个字「识」，第二个字「别」，第三个字「完」，第四个字「毕」。
2. 若输出了「无法识别该图片」，在该行之后同样另起一行输出上述结束标记。
3. 该结束标记是输出格式的组成部分，不属于规则所禁止的多余文字；下游系统以此判断识别流程已结束。特别说明：本指令正文中不存在该标记的连续写法，因此窗口内一旦出现该标记，必定来自你的回复。
"@

# 槽位填空自检：任何槽位为空都会让指令出现空白项，先拦截
foreach ($kv in @{ 'taskSummary' = $taskSummary; 'userQuestion' = $userQuestion; 'inferenceScope' = $inferenceScope;
                   'focusHint' = $focusHint; 'basisHint' = $basisHint; 'answerForm' = $answerForm;
                   'depth' = $depth; 'extraNote' = $extraNote }.GetEnumerator()) {
    if ([string]::IsNullOrWhiteSpace($kv.Value)) { Write-Error ("槽位未填: " + $kv.Key); exit 1 }
}
if ($depth -notin '简', '标准', '详细') { Write-Error "depth 取值必须是 简/标准/详细，当前=$depth"; exit 1 }
$fieldCount = @(($prompt -split "`n") | Where-Object { $_ -match '^- (图片类型|核心主体|空间布局|场景与背景|画面文字|细节特征|特殊元素|不确定内容说明|推断对象|置信度与理由|判断依据|其他可能|无法推断项说明)：（' }).Count
"PROMPT_ASSEMBLED  字符数=$($prompt.Length)  字段行=$fieldCount（应为 13）"
if ($fieldCount -ne 13) { Write-Error "字段行数异常：骨架被改动过，请核对 13 个字段名"; exit 1 }

[System.Windows.Forms.Clipboard]::SetText($prompt)
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

# 发送成功后【不要最小化】：保持窗口展开，步骤5 才能以高频轮询读取生成进度
Start-Sleep -Milliseconds 500
```

**发送完成确认（替代原来的"单独记基线"步骤）**

> ⚠️ **不要在这里算基线**。步骤4 与 5-1 是**两个独立进程**，在此处算出的 `$script:markerBase` **传不到轮询里**（实测：轮询端 `Get-Variable` 取不到 → 基线恒为 0 → 判据退化成"窗口里出现过标记"）。**基线必须在 5-1 轮询进程内、开轮询之前采集**。

```powershell
# 只做一件事：确认消息已真正发出（输入框清空 ≈ 消息已入列表，是"可以开始轮询"的信号）
$sendConfirmMs = 5000
$swSent = [System.Diagnostics.Stopwatch]::StartNew()
$sentOk = $false
while ($swSent.ElapsedMilliseconds -lt $sendConfirmMs) {   # 上限 5 s
    Start-Sleep -Milliseconds 200
    $winS = Get-DoubaoWinByProcess $doubao
    if (-not $winS) { continue }
    $allS = $winS.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    $ed = $allS | Where-Object { $_.Current.ClassName -like '*ProseMirror*' } | Select-Object -First 1
    $inBox = ''
    if ($ed) {
        $inBox = (@($ed.FindAll([System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition) |
            ForEach-Object { $_.Current.Name }) -join '')
    }
    if ($inBox -notmatch '负责识别图片') { $sentOk = $true; break }   # 输入框已清空
}
"SENT_CONFIRMED=$sentOk  用时 $($swSent.ElapsedMilliseconds) ms  标题=[$((Get-Process -Id $doubao.Id).MainWindowTitle)]"
if (-not $sentOk) { Write-Output '警告：输入框可能未清空（消息未发出？），轮询仍将继续' }
```

### 步骤5 等待识别完成并读取结果

#### 步骤5-1 等待豆包响应

豆包开始处理后，窗口标题通常会变为与任务相关的标题（如「图片识别任务说明 - 豆包」）。

**轮询期间窗口保持展开（不最小化）**，完成判据命中后先抢读落盘再立即最小化（两步顺序不可颠倒）：最小化状态下 Chromium 窗口不向 UI Automation 暴露元素，实测展开 332 元素 / 最小化 0 元素，轮询读不到进度，只能反复"恢复-读取-再最小化"，既拖慢确认又频繁闪屏。保持窗口可见（不必始终前台）即可直接读生成状态。

不要使用固定长等待。**完成判据只有一个：结束标记元素数 > 基线**（v4.8 起删除全部兜底判据——它们历史上救场 0 次、致障 1 次，详见下方代码内注释）。标记不出现即报超时并指引「仅重抓」，不做猜测式兜底。

> ⚠️ **基线必须在本进程内采集**（见下），不要依赖步骤4 传值——跨进程变量传不过来。

```powershell
# ============ 完成判据：只看结束标记（单一判据，无兜底）============
# 设计纪律（v4.8）：这一节曾堆到 6 个分支（字段行数/末字段/文本稳定/整窗文本/追问建议…），
# 它们互为兜底、其中一个被嵌套在另一个内部，最终互相锁死，导致"豆包已答完却卡住数分钟"。
# 复盘数据：兜底判据历史上救场 0 次、致障 1 次。故只保留唯一可靠的判据——结束标记。
# 指令已强制要求模型输出该标记（见步骤4 骨架【结束标记】第 3 条），因此"标记出现=生成完毕"。
# 标记不出现即视为协议被违反：报错并指引 5-4「仅重抓」，不做任何猜测式兜底。
$markerPattern = '识别完毕'

# ---- 基线：本进程内采集，等"标记计数稳定"（避免把用户消息里的标记算成增量）----
$script:markerBase = 0
$markerStableMs = 8000   # T15：基线稳定判定上限
$swB = [System.Diagnostics.Stopwatch]::StartNew()
$prevM = -1; $stableM = 0
while ($swB.ElapsedMilliseconds -lt $markerStableMs) {
    Start-Sleep -Milliseconds 250
    $wb = Get-DoubaoWinByProcess $doubao
    if (-not $wb) { continue }
    $ab = $wb.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    $m = 0
    foreach ($el in $ab) {
        $nm = $el.Current.Name
        if ($nm -and (($nm -replace '[、，,．.·\s]', '') -match $markerPattern)) { $m++ }
    }
    $script:markerBase = $m
    if ($m -eq $prevM) { $stableM++ } else { $stableM = 0 }
    if ($stableM -ge 2) { break }   # 连续两轮不变 → 用户消息已渲染完，基线可信
    $prevM = $m
}
"MARKER_BASE=$script:markerBase  用时 $($swB.ElapsedMilliseconds) ms"

# ---- 轮询：只等标记增量 ----
$deadline = (Get-Date).AddSeconds(90)   # T5 图片档；标记判据下正常 5–20 秒结束
$markerSince = $null
$done = $false
$doneBy = ''
$missCount = 0
$swPoll = [System.Diagnostics.Stopwatch]::StartNew()
do {
    Start-Sleep -Milliseconds 500   # T6
    $win = Get-DoubaoWinByProcess $doubao
    if (-not $win) {
        # 窗口消失（被最小化/遮挡）：自包含恢复，避免空等
        $missCount++
        if ($missCount -ge 4) {
            Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name Win32Restore -Namespace Win32 -ErrorAction SilentlyContinue
            if ([Win32.Win32Restore]::IsIconic($doubao.MainWindowHandle)) {
                [Win32.Win32Restore]::ShowWindow($doubao.MainWindowHandle, 9) | Out-Null
                Start-Sleep -Milliseconds 800
            }
            $missCount = 0
        }
        continue
    }
    $all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    $markerCount = 0
    foreach ($el in $all) {
        $nm = $el.Current.Name
        if ($nm -and (($nm -replace '[、，,．.·\s]', '') -match $markerPattern)) { $markerCount++ }
    }
    if ($markerCount -gt $script:markerBase) {
        if ($null -eq $markerSince) { $markerSince = Get-Date }
        elseif (((Get-Date) - $markerSince).TotalSeconds -ge 1.5) {   # T6：等最后一行渲染完
            $done = $true; $doneBy = 'MARKER'; break
        }
    }
} until ((Get-Date) -gt $deadline)

if ($done) {
    "DONE_BY=$doneBy  用时 $([int]$swPoll.Elapsed.TotalSeconds) s  标记=$markerCount/基线=$script:markerBase"

    # ★★ 抢读：必须在最小化【之前】。实测（2026-10 对照实验）豆包窗口最小化后 UIA 完全读不到：
    #     展开态 332 元素 / 3 字段行  →  最小化态 0 元素 / 0 字段行。
    #    先最小化再读 = 结果永久读不到，只能恢复窗口重读。
    $grab = @()
    foreach ($el in $all) {
        $nmG = $el.Current.Name
        if (-not $nmG) { continue }
        if ($nmG -match $replyPattern) {
            try { $rG = $el.Current.BoundingRectangle; $yG = [int]$rG.Y } catch { $yG = -1 }
            $grab += [pscustomobject]@{ Y = $yG; Name = $nmG }
        }
    }
    $resultItems = @($grab | Sort-Object Y -Unique | ForEach-Object { $_.Name })
    "GRAB_LINES=$($resultItems.Count)   （抢读完成，窗口尚未最小化）"
} else {
    Write-Output ("POLL_TIMEOUT 已等 $([int]$swPoll.Elapsed.TotalSeconds) s  标记计数=$markerCount 基线=$script:markerBase")
    Write-Output '→ 结束标记未出现 = 模型未按指令收尾（协议违反）。请检查豆包界面是否已答完：'
    Write-Output '   · 若已答完 → 按 5-4「仅重抓」直接读取，不要重发'
    Write-Output '   · 若仍在生成 → 延长 T5 或检查网络后重试'
}

# ★★★ 无条件最小化（不放在任何 if 分支里）——这是"收尾动作"，不是"判据分支"。
#   设计教训：早期把它写进 if ($done) 分支内，结果脚本执行了却没最小化，
#   窗口一直挡在用户面前，直到用户指出。凡"必须执行"的动作一律写在判据之外。
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name Win32MinNow -Namespace Win32 -ErrorAction SilentlyContinue
[Win32.Win32MinNow]::ShowWindow($doubao.MainWindowHandle, 6) | Out-Null   # SW_MINIMIZE
'MINIMIZED_IMMEDIATELY'   # 输出后智能体方可继续思考/汇报

```

> - **判据只有一条**：结束标记元素数 > 基线。不看字段行数、不看文本稳定、不看"停止"按钮——
>   这些兜底曾在 v4.7.2 因互相嵌套锁死，导致"已答完却卡住数分钟"，故 v4.8 全部删除。
> - 标记由指令强制要求输出，因此"出现即完成"；不出现就是协议被违反，**报错胜于猜测**。
> - 超时上限 T5（图片 90 秒 / PPT 240 秒）只在异常时生效，正常 5–20 秒结束。

#### 步骤5-2 读取结果（限定在"最新一条回复的消息块"内）

**为什么必须限定作用域**（v4.9 实测缺陷）：早期实现是**全窗口扫描**「字段名：实际内容」形态的文本行，结果把**用户自己发送的追问指令**也读了进来——追问模板里本就引用了「不确定内容说明」等字段名的完整写法，一旦用户消息里出现"字段名：非括号内容"的形态，就会被误判为回复行。全窗口扫描既会读到上一轮的旧回复，也会读到用户消息。

**正确做法——以结束标记为锚点定位消息块**：

1. 结束标记所在的元素 **必然属于最新一条回复**（指令保证标记只可能由模型输出）；
2. 从标记元素沿**祖先链向上**，取"其几何高度覆盖首字段行、且自身高度最小"的那一层作为消息块容器；
3. 只在该容器子树内收集回复行，**不再全窗口扫描**。

```powershell
$replyPattern = '^(图片类型|核心主体|空间布局|场景与背景|画面文字|细节特征|特殊元素|不确定内容说明|推断对象|置信度与理由|判断依据|其他可能|无法推断项说明)\s*：\s*(?![（(])\S'
$markerPattern = '识别完毕'

# ---- ① 找结束标记元素（取 y 最大的那个 = 最新的回复）----
$markerEl = $null; $markerY = -1
foreach ($el in $all) {
    $nm = $el.Current.Name
    if (-not $nm) { continue }
    if (($nm -replace '[、，,．.·\s]', '') -notmatch $markerPattern) { continue }
    try { $r = $el.Current.BoundingRectangle } catch { continue }
    if ([double]::IsNaN($r.Y) -or [double]::IsInfinity($r.Y)) { continue }
    if ($r.Y -gt $markerY) { $markerY = $r.Y; $markerEl = $el }
}

# ---- ② 找回复首字段行元素的 y（用于确定消息块上边界）----
$firstFieldY = [double]::MaxValue
foreach ($el in $all) {
    $nm = $el.Current.Name
    if (-not $nm) { continue }
    if ($nm -notmatch '^图片类型\s*：') { continue }
    try { $r = $el.Current.BoundingRectangle } catch { continue }
    if ([double]::IsNaN($r.Y) -or [double]::IsInfinity($r.Y)) { continue }
    # 只取标记"上方最近"的那个「图片类型：」
    if ($r.Y -lt $markerY -and $r.Y -lt $firstFieldY) { $firstFieldY = $r.Y }
}

# ---- ③ 从标记元素向上找"最小且覆盖首字段行"的容器 ----
$replyScope = $null
if ($markerEl) {
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $node = $walker.GetParent($markerEl)
    for ($k = 0; $k -lt 8 -and $node; $k++) {
        try { $rb = $node.Current.BoundingRectangle } catch { $rb = $null }
        if ($rb -and -not [double]::IsNaN($rb.Y) -and -not [double]::IsInfinity($rb.Y)) {
            $covers = ($firstFieldY -eq [double]::MaxValue) -or ($rb.Y -le ($firstFieldY + 30))
            $tall = $rb.Height -gt 200
            if ($covers -and $tall) { $replyScope = $node; break }
        }
        $node = $walker.GetParent($node)
    }
}
if (-not $replyScope) { $replyScope = $win }   # 兜底：退化为窗口级（结果末尾会给出告警）

# ---- ④ 只在消息块内收集，且按屏幕顺序排序 ----
$replyPattern = '^(图片类型|核心主体|空间布局|场景与背景|画面文字|细节特征|特殊元素|不确定内容说明|推断对象|置信度与理由|判断依据|其他可能|无法推断项说明)\s*：\s*(?![（(])\S'
$pairs = @()
foreach ($element in $replyScope.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)) {
    $name = $element.Current.Name
    if (-not $name -or $name -notmatch $replyPattern) { continue }
    try { $r = $element.Current.BoundingRectangle } catch { continue }
    $y = -1
    if (-not [double]::IsNaN($r.Y) -and -not [double]::IsInfinity($r.Y)) { $y = [int]$r.Y }
    $pairs += [pscustomobject]@{ Y = $y; Name = $name }
}
# 去重（同一行可能同时暴露为 ListItem 与 Text）
$resultItems = @($pairs | Sort-Object Y -Unique | ForEach-Object { $_.Name })

# ---- ⑤ 作用域告警：退化到窗口级时必须显式提示 ----
if ($replyScope -eq $win) {
    Write-Output 'SCOPE_WARN 未能定位消息块（endmarker 或首字段行未找到）→ 已退化为全窗口扫描，'
    Write-Output '           读取结果可能混入用户消息或旧回复，请人工核对后使用'
} else {
    Write-Output ("READ_SCOPE=消息块  容器高度=$([int]$replyScope.Current.BoundingRectangle.Height)  回复行=$($resultItems.Count)")
}
```

> **作用域判据说明**：容器必须"向上覆盖到首字段行"（`$rb.Y -le $firstFieldY + 30`）且"有足够高度"（>200px）——这排除了只包住标记自身的窄小容器，也排除了整页级容器（它在更外层，会在循环中先被更小的合格容器取代）。
#### 步骤5-3 格式校验（两部分分别校验）

按 T14 分组统计，**两组都必须齐**：

| 层级 | 要求 | 不通过的处置 |
|---|---|---|
| 第一部分（8 字段） | 单图每个字段名恰好 1 次；多图每个字段名出现 N 次（N = **本批**图片数，基准见 1.3） | 判失败 → 重发指令 1 次 |
| 第二部分（5 字段） | 同上，逐字段计数一致 | 判失败 → 重发指令 1 次 |
| 两部分都缺 | — | 输出「图片识别结果格式异常，无法解析」 |

- 豆包可能不带「图N」标题（实测为字段组连续排列），按字段出现次数判定即可；
- 字段行三态语义见 1.4：值为 `无` 是有效行，只有**整行缺失**才判失败；
- **若仅第二部分缺字段**：先按步骤7-1 追问一次「补第二部分」（比整条重发便宜得多）；追问后仍缺，则**保留已拿到的第一部分**并在汇报中注明「第二部分缺失」，不得丢弃第一部分结果；
- 校验通过后，汇报必须**显式标注第二部分为推断内容**（见步骤6）。

#### 步骤5-4 落盘与校验（最小化已在步骤5-1 完成）

**最小化时机（v4.9 调整，经实测校正）**：识别完毕的那一刻执行 **抢读落盘 → 立即最小化** 两步，由步骤5-1 完成判据块自动完成（`GRAB_LINES=…` 后接 `MINIMIZED_ON_COMPLETE`）。

**为什么必须"先抢读再最小化"**（2026-10 对照实验）：豆包窗口**最小化后 UIA 完全读不到内容**——

| 窗口状态 | 可读元素数 | 其中字段行 |
|---|---|---|
| 展开 | 332 | 3 |
| **最小化** | **0** | **0** |
| 再展开 | 332 | 3 |

因此若先最小化再读取，结果将**永久读不到**（只能恢复窗口重读）。故完成判据处先做一次抢读落盘（此刻内容已就绪、窗口可读，最可靠），再最小化。

> ⚠️ 与旧版顺序的差异：旧版是"**落盘 → 校验 → 通过才最小化**"；新版把最小化提前到完成判据处，但保留了前置抢读，**落盘与校验照旧执行**。若脚本在 5-1 之后中断，窗口可能已最小化而完整转储未完成——此时按下方「仅重抓」处理，**必须先把窗口恢复展开再读取**。

落盘与校验步骤：

1. **落盘**：把读到的回复转储写入本次临时目录（如 `reply_dump.txt`），防止后续任何异常丢失结果；
2. **校验有效性**：图片模式——按 T14 两组分别校验字段行数（第一部分 8 字段 × 本批张数、第二部分 5 字段 × 本批张数）；PPT 文本分析模式——转储含「PPT类型判定结果」等框架锚点（**匹配前先去空白**）。

```powershell
# 幂等保险：若 5-1 之后流程被中断过，这里再确认一次窗口已最小化（重复调用无副作用）
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name Win32EnsureMin -Namespace Win32 -ErrorAction SilentlyContinue
if (-not [Win32.Win32EnsureMin]::IsIconic($doubao.MainWindowHandle)) {
    [Win32.Win32EnsureMin]::ShowWindow($doubao.MainWindowHandle, 6) | Out-Null
    'MINIMIZED_LATE'   # 正常情况下不应出现；出现说明 5-1 的最小化没执行到
} else { 'ALREADY_MINIMIZED' }
```
>
> **「仅重抓」（不重传不重发）**：转储/校验失败（脚本中断、内容截断等）但豆包回复其实已生成时，**不要重跑整个任务**（重新上传+重发 ≈ 3–5 分钟）。直接复用当前对话执行一次「恢复窗口 → 滚动到底 → 健壮转储」（4.1 步骤6 的逐元素容错代码，只读取、不发送，上限见 T19），成功后补校验、再最小化并进入收尾清理。

### 执行时序纪律（步骤5 收尾，强制；误序会直接暴露给用户）

**唯一合法顺序：完成判据命中 → 抢读落盘 → 最小化 → 才开始思考与汇报。**

| 步 | 动作 | 不可颠倒的理由 |
|---|---|---|
| 1 | 完成判据命中（结束标记增量） | 判定结果已生成 |
| 2 | **抢读落盘**（`GRAB_LINES=…`） | 窗口最小化后 UIA **读不到任何元素**（实测展开 332 → 最小化 0），抢读必须先做 |
| 3 | **立即最小化**（`MINIMIZED_IMMEDIATELY`） | 窗口不再挡在用户面前；后续思考/汇报不需要窗口可见 |
| 4 | 校验、整理、汇报 | 此时窗口已最小化，用户可正常使用电脑 |

**两条硬性纪律**：

1. **"必须执行"的动作绝不写进判据分支。** 最小化曾被放在 `if ($done) {...}` 内，结果脚本跑完却没最小化——因为动作依附了条件。凡收尾动作（最小化、清理、前台还原）一律写在 `if/else` **之外**，无条件执行。
2. **抢读与最小化必须在同一次工具调用内完成，中间不返回控制权。** 若拆成"先轮询、再最小化"两次调用，轮询一返回，智能体就可能先开始分析输出、把最小化排到后面——这正是本纪律要消除的时序。脚本一次跑完：轮询 → 抢读 → 最小化，三件事同一进程内完成，`MINIMIZED_IMMEDIATELY` 是最后一行输出。

> **违规症状**：用户看到豆包窗口仍开着，而智能体已经在输出识别结果。出现即说明最小化未随抢读一起执行——按本纪律回到 5-1 重跑收尾（抢读仍有效，直接补最小化即可；必要时先恢复窗口再抢读）。
### 步骤6 结果输出（以豆包为准）

**识图结果以豆包为准。智能体只做整理与排版，不再做二次判断。**

**汇报只分两层（豆包的两部分），不得添加自己的判读层**：

| 层 | 内容 | 标注要求 |
|---|---|---|
| 一、客观描述 | 豆包第一部分 8 字段（原样引用或表格化） | 标明「豆包识别原文」 |
| 二、推想预测 | 豆包第二部分 5 字段（推断对象/置信度与理由/判断依据/其他可能/无法推断项说明） | 标明「豆包推断」；结论、置信度、其他可能全部原样转述 |

**禁止事项（按用户要求，2026-10 明确）**：

- **不得对豆包结论做二次判断**：不加"我认为证据更弱/更强"、"这里有矛盾"、"我更倾向另一个"这类点评；
- **不得用自己的保留意见削弱结论**：识别结论是否正确由豆包负责，智能体如实转述即可；
- **不得只报一部分**（那属于信息遗漏，同 1.4 的防遗漏要求）：两部分、每张图都要完整呈现；
- 不得把自己的解读伪装成豆包原文。

**允许做的事**：字段化排版（表格/列表）、去除重复措辞、把图片链接放在对应图前——**只动形式，不动结论**。

- 每张图片介绍前必须附该图片的**可点击本地文件链接**，例如 `[C271F42C14CFC125F33AACC788D0C6ED.jpg](file:///D:/CSY/C271F42C14CFC125F33AACC788D0C6ED.jpg)`；链接不可点击时至少给出完整绝对路径 `D:\CSY\文件名.jpg`；
- **不得把第二部分的推断写进第一部分的位置**，也不得把推断升格为事实陈述；
- 转述推断时**完整保留结论、置信度与「其他可能」**，禁止只挑最可能的答案；
- **推断部分应显著呈现，不要弱化成"仅供参考"**：把最有价值的候选结论放在前面（含多候选排序），再给依据；
- 第二部分缺失时如实说明「本次未获取到推断部分」，不自行补写；
- 若确有必要提示不确定性，**只能引用豆包自己的表述**（如它的置信度、它的「无法推断项说明」），不得由智能体自行添加质疑。

### 步骤7 追问图片细节（可选）

出现以下任一情况时不要直接结束任务，继续向豆包追问：

- 用户对识别结果追问（"图里某个字是什么""这个按钮写的是什么""能不能放大看细节"）；
- 当前任务需要额外视觉信息才能继续（颜色、材质、位置、文字完整内容、二维码内容等）；
- 首次结果中存在模糊、遮挡、不确定内容，需要进一步确认；
- **第二部分整段缺失**（5-3 校验发现）——用模板 C 补齐，不整条重发；
- **发现个体被"等/其他/众多"合并带过**（1.4 防遗漏契约判定为遗漏）——用模板 D 追补，不整条重发。

#### 步骤7-1 生成追问文本命令

根据用户问题或任务需要，生成一条**合理、具体、只针对图片细节的文本命令**。三类模板（按用途选用）：

**模板 A：细节辨认（默认，绝大多数追问都用它）**

```text
请针对刚才上传的图片继续回答：
- 画面中「不确定内容说明」提到的模糊区域，尝试再次辨认并描述可见内容。
- 请完整转录图中所有文字，不要省略。
- 请详细描述图中第X个物体的颜色、材质、位置、周围环境。
- 如果图中包含二维码/条形码，请描述其外观和可辨识内容。
- 请补充说明图片的上下左右边缘是否还有未识别内容。
- 回答完毕后，另起一行输出由四个汉字连写（不加标点）组成的结束标记：第一个字「识」，第二个字「别」，第三个字「完」，第四个字「毕」。
```

**模板 C：补缺（仅用于第二部分整段缺失）**

```text
你上一条回复缺少「第二部分 推想预测」。请只输出该部分，严格按以下字段，不要重复第一部分：
推断对象：/ 置信度与理由：/ 判断依据：/ 其他可能：/ 无法推断项说明：
要求：推断对象优先给确定的名字，多候选按可能性排序；置信度按【高/中/低】给；判断依据逐条列出第一部分的可见特征。只有确实无法判断时才填「无法判断」，禁止编造不存在的名称。
回答完毕后，另起一行输出由四个汉字连写（不加标点）组成的结束标记：第一个字「识」，第二个字「别」，第三个字「完」，第四个字「毕」。
```

**模板 D：补遗漏（发现个体被"等/其他/众多"带过时用）**

```text
你上一条回复存在信息遗漏。请重新仔细查看刚才上传的图片，只补充遗漏的个体，不要重复已经答过的内容：
- 把画面中所有可辨识的个体（人物/角色/动物/物体/条目）逐个列出，每一个都单独给出可定位标识：优先名称；无法确定名称时用位置（如“左侧第一人”“前排右二”）或外观代称（如“白发少年”“蓝发双马尾少女”）。
- 若某区域人物过于密集无法逐一细述，请给出该区域的可计数信息（数量或可数范围）与分布位置，并对其中较大的个体分别描述。
- 禁止使用“等”“等等”“其他人物”“背景人物”“若干”“众多”这类表述合并多个个体。
- 逐条列出后，再补一句说明是否还有无法辨认的个体及其位置。
- 回答完毕后，另起一行输出由四个汉字连写（不加标点）组成的结束标记：第一个字「识」，第二个字「别」，第三个字「完」，第四个字「毕」。
```

> 模板 D 与已删除的模板 B 的区别：**B 是"怕它答得不够好"的推测需求，D 是防遗漏契约的直接配套**——没有 D，检测到遗漏后只能整条重发（重传重算，成本高且可能再次遗漏）。两者性质不同，D 是必要机制。

约束：

- 命令必须围绕图片本身，禁止询问与图片无关的信息；
- 涉及多图时必须指明"图1/图2/图3……"；
- 命令应能直接发送给豆包，无需额外解释；
- **追问推断类内容时必须同时要求给出"判断依据"**——没有依据的推断等于猜测，不可采纳（模板 C 已内置该要求）；
- **追问文本中禁止出现连续的"识别完毕"四字**。步骤5-1 已用"标记元素数 > 基线"（步骤4 记录）从机制上排除追问文本与历史回复的干扰；保持本纪律可减少"归一后等于标记"的元素数量，让完成判定更快更稳。需要提及结束标记时一律用拆字写法；若未附带结束标记要求，豆包答完不输出标记，轮询自动走字段兜底判据，不影响完成检测。

#### 步骤7-2 发送追问并读取结果

首次识别完成后豆包窗口只是被最小化，**对话和已上传图片仍然保留**：

1. 窗口为最小化状态 → 先执行 PP2.1 恢复激活并确认 `DOUBAO_READY`；发送前同样须确认窗口在前台；
2. 直接聚焦原输入框；
3. 追问文本入剪贴板后 `Ctrl+V` 粘贴（**禁止用 SendKeys 逐字输入中文**，见 P3）；
4. 发送方式同步骤4：优先「发送」按钮，找不到才按 `Esc` + Enter；
5. 不重新开新对话，不重新上传图片；
6. 等待与读取同步骤5（步骤5-1 轮询）。

只有极端情况下豆包窗口已被关闭，才从 PP2.1 重新开始并重新上传图片。

#### 步骤7-3 追问结果处理

- 豆包返回的补充内容同样算「豆包识别原文/原字段」；
- 结合首次结果与追问结果，输出更完整的解读；
- 追问后仍无法识别 → 明确标注无法确认，不做无依据猜测。

### 步骤8 收尾：清理 → 最小化保留 → 前台还原

识别完成并输出结果后**不要关闭豆包窗口**，按以下顺序收尾：

#### 步骤8-1 先清理本次产物，再汇报

```powershell
# 汇报给用户之前执行（成功 / 失败 / 中断三条路径都执行）
# $taskDir 仅 PPT 拆解模式存在；图片模式无产物目录。先判空再删，失败必出声。
$taskDir = $outDir        # PPT 拆解模式：3.2 生成的 tmp_ppt_<8位hex> 目录
$helperScript = $null     # 本次若写过辅助脚本，填其绝对路径；否则保持 $null

$targets = @()
if ($taskDir)      { $targets += $taskDir }
if ($helperScript) { $targets += $helperScript }
foreach ($t in $targets) {
    if (-not (Test-Path -LiteralPath $t)) { continue }
    try {
        Remove-Item -LiteralPath $t -Recurse -Force -ErrorAction Stop
        Write-Output "CLEANED $t"
    } catch {
        Write-Output "CLEANUP_FAILED $t （可能被 PowerPoint/编辑器占用，已记录，请人工清理）"
    }
}
```

清理范围：PPT 拆解产物目录、media 提取物、中间结果文件、缓存、截图、日志、OCR 中间文件、辅助脚本。**不删除用户原始图片/PPT**，删除前逐一核对路径确为本流程产物，并遵守 步骤8-2 黑名单。任务失败或中断时同样必须执行本清理，且**清理必须在最终汇报之前完成**。

#### 步骤8-2 防误删黑名单（永不删除）

- `D:\DS\.dsh\skills\`（技能本体及历史）；
- 用户原始图片/PPT 及其所在目录（含 Doubao 聊天目录、桌面、文档等任何位置的源文件）；
- 本次任务开始前就已存在的任何文件（如历史脚本 `D:\DS\ocr.ps1`、历史产物目录）；
- 无法确认归属的文件/目录。

**不做"任务启动前的主动扫描"**（每次先扫盘显著拖慢任务）；磁盘上的历史遗留物仅在用户指出时按本黑名单核对归属再处理。

#### 步骤8-3 最小化保留 + 前台还原

```powershell
$sig = '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
Add-Type -MemberDefinition $sig -Name Win32ShowWindow -Namespace Win32
[Win32.Win32ShowWindow]::ShowWindow($doubao.MainWindowHandle, 6) # SW_MINIMIZE

Add-Type -AssemblyName Microsoft.VisualBasic
[Microsoft.VisualBasic.Interaction]::AppActivate($originalPid) | Out-Null   # 还原 P1 记录的前台窗口
```

确认豆包已最小化并**保留当前对话与已上传图片**（5-1 完成判据处已抢读落盘并最小化；若期间被恢复过则重新最小化），便于后续追问。

#### 步骤8-4 关闭豆包的条件（满足任一才关）

- 用户明确表示不再需要继续识别/追问；
- 用户开始执行与当前图片无关的其他命令；
- 当前任务所需图片信息已完整，且后续不会再用到这些图片；
- 用户主动要求关闭；
- 要开始另一批**全新**图片识别任务（避免新旧图片混淆）。

#### 步骤8-5 误开的无关窗口

```powershell
Get-Process -Name PCAppStore -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } |
    ForEach-Object { $_.CloseMainWindow() }
```

若权限不足无法关闭，提示用户手动关闭。

---

## 附录A 固定指令原文

> 两段指令必须**逐字发送，不得删减或修改**（包括标点、编号、示例括号）。修改前必须重新评估 5-1 的判据与 T9 正则是否仍然成立。
>
> A.1 为**两部分结构**（8 客观字段 + 5 推断字段），且已改为**骨架 + 槽位**：骨架不可改，槽位按《A.1a》填。改动 A.1 必须同步四处：步骤4 组装块的骨架原文、1.4 字段契约、1.5 的 T9/T10 正则、T14 分组——**四处字段名必须逐字一致**（见附录D 的 V2c）。槽位措辞可自由调整，不动字段名即安全。
> A.2 为 PPT **文本**分析指令，仅客观分析、不含推想部分；它不参与 T9 字段判据。

### A.1 图片识别指令（固定骨架 + 可编辑槽位）

**骨架原文**：见「5. 执行层：步骤4」组装块中的 `$prompt`。骨架**不可改**，只有槽位变量是任务相关的。

### A.1a 可编辑槽位（本次任务该填什么）

图片内容不限于人物——也可能是动物、作品、地点、景物、物理现象或示意图。**第二部分若只预设"角色/作品/地点"，遇到物种、原理、图表类图片就会答偏**。因此第二部分改为**由槽位驱动**：字段名固定，字段描述与侧重按任务生成。

> **填槽优先级**：**用户在请求中点明的角度 > 本次工作/下游任务需要 > 下表"默认"值**（见 1.2 优先级表）。下表默认列只是兜底；用户一旦指定角度（如"从植物学角度""我要原理不要名字""只判断是否 AI 生成"），就按用户说的填，不得改回默认。
> 例：用户说"这个动物是什么，顺便看看是不是保护物种" → `$inferenceScope` 填"物种鉴别 + 保护级别"，`$answerForm` 填"中文名 + 学名 + 保护等级"，而不是照搬"动物/植物"行的默认措辞。

| 槽位 | 作用 | 填写依据 | 默认（无信息时） |
|---|---|---|---|
| `$taskSummary` | 一句话概括本次任务，写进指令做总纲 | 用户问题 + 本次工作需要 | 必填，不得留空 |
| `$userQuestion` | 用户原始问题；让豆包知道答案给谁看 | 用户原话（可精简） | 无明确问题时按任务改写 |
| `$inferenceScope` | **推断什么、侧重哪类**（可写多类） | 见下表任务类型 | 泛化：对象身份 + 领域 + 来源 |
| `$focusHint` | 第一部分描述时要特别关注的可观察线索 | 能在图上区分该类对象的关键特征 | 泛化：形状、颜色、材质、结构、文字 |
| `$basisHint` | 判断依据里优先列举哪些线索 | 同上，可更具体 | 同上 |
| `$answerForm` | 期望的答案形态与粒度 | 用户想要的结论形式 | 先给最可能结论，再列其他候选 |
| `$depth` | `简` / `标准` / `详细` | 用户是否要细节、结果是否要进正式文档 | `标准` |
| `$extraNote` | 本次特殊要求 | 如"必须给出拉丁学名""标注公式出处""指出图中错误" | `无` |

**按任务类型填槽（对照表）**：

| 图片内容类型 | `$inferenceScope` 应写 | `$answerForm` 应写 | `$basisHint` 应写 |
|---|---|---|---|
| 人物 / 角色 | 角色身份、作品出处、是否为同人/原创 | 角色名 + 作品名，再列候选 | 发型发色、瞳色、服饰纹样、标志性配饰 |
| 动物 / 植物 / 微生物 | 物种类别、具体物种名、是否保护物种 | 中文名 + 学名（若可判断），再列相近种 | 体型、斑纹、喙/角/叶形、栖息环境 |
| 作品 / 产物（实物、产品、画作、模型） | 作品名、作者/品牌、系列与型号 | 名称 + 品牌或作者，再列候选 | logo、外观特征、结构细节、年代风格 |
| 地点 / 建筑 / 地标 | 具体地点名、建筑名、城市与国家 | 地点或建筑名，再列候选 | 建筑形制、地标特征、招牌文字、植被气候 |
| 景物 / 自然 / 天气 | 景观或地理名称、天气现象名称 | 名称 + 判断把握，再列候选 | 地貌、云型、光线方向、植被类型 |
| 物理现象 / 规律 / 示意图 | 所反映的规律或原理、公式含义 | 原理名 + 表达式（若图中给出），再说适用条件 | 标注符号、箭头方向、刻度、量纲、文字标签 |
| 文字 / 标识 / 票据 | 文字含义、来源、所属体系 | 含义 + 出处，再列候选 | 字体风格、印章、编号规则、版式 |

**填写纪律**：

- **用户点明的角度优先**：用户指定了侧重、角度或答案形态时，按用户说的填槽，**不得改回对照表默认或自行判断的方向**（1.2 优先级表 P1）。这一点在代码注释里也再提一次，避免填槽时忘记。
- **能判断就不要用默认值**——泛化槽位是兜底，不是常态；槽位写得越贴合任务，第二部分越有用。
- 槽位只影响第二部分的侧重与粒度，**不得借槽位改动字段名、字段数量或格式**（那会破坏 T9/T10 判据）。
- `$inferenceScope` 与 `$answerForm` 允许多类并列（如"既是动物物种、也要判断是否 AI 生成图"）。
- 若用户明确说"只要看图里有什么，不要猜"：`$inferenceScope` 填"仅需客观描述，推断部分只需标注哪些内容无法从画面确证"，`$answerForm` 填"如实说明依据不足"，此时第二部分仍要输出字段但不强行给结论。

### A.2 PPT 文本分析指令

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

四、结束标记
1. 上述全部分析内容输出完毕后，在回复末尾另起一行，输出由四个汉字连续书写（紧挨连写、不加空格、不加任何标点）组成的结束标记；四个汉字依次为：第一个字「识」，第二个字「别」，第三个字「完」，第四个字「毕」。
2. 该结束标记是输出格式的组成部分，不属于前面所禁止的多余文字；下游系统以此判断分析流程已结束。特别说明：本指令正文中不存在该标记的连续写法，因此窗口内一旦出现该标记，必定来自你的回复。
```

---

## 附录B 异常处理矩阵

| 类别 | 场景 | 处理 |
|---|---|---|
| 文件 | 路径无效 / 不存在 / 格式不支持 / 文件损坏 / 无读取权限 | 分别告知具体原因（路径无效、已剔除不支持格式、文件损坏、权限不足），请用户处理后重试 |
| 文件 | 遇到 .ppt/.pptx | 按分流表处理：问文本走步骤0.6 直传；问画面走步骤0 PPT 拆解；页数 >10 分批 |
| 文件 | 目标是 .ppt/.pptx | 按第 1 章的分流表：文本内容走步骤0.6 直传；画面内容走步骤0 拆解后按图片流程识别；页数 >10 分批 |
| 文件 | 整页渲染失败（未安装 PowerPoint/WPS COM） | .pptx 退化为 3.3 解包提取 media 原图识别；.ppt 告知用户需安装 Office/WPS 后重试 |
| 文件 | 源文件在回收站/被占用（COM 打开 E_FAIL 或附件粘贴失败） | 先复制到 `tmp_ppt_<hex>\src\` 再渲染/上传副本，副本随步骤8-1 清理，原文件不动；复制也失败则明确报错 |
| 文件 | 磁盘上发现历史遗留的 tmp_ppt_* 目录或辅助脚本 | 不做启动前扫描；用户指出时按步骤8-2 黑名单核对归属再处理；本次任务产物在步骤8-1 清理 |
| 文件 | 收到相对路径或 `.` | 先转绝对路径；无法转换时请用户给完整绝对路径 |
| 文件 | 格式不支持、文件损坏 | 告知不支持/已损坏，请更换文件 |
| 文件 | 无读取权限 | 告知权限不足，调整后重试 |
| 客户端 | 未安装豆包、未登录账号 | 提示安装并登录豆包桌面客户端后重试 |
| 客户端 | 启动超时、进程崩溃 | 自动重试 1 次，失败则提示客户端异常 |
| 客户端 | 已打开但没有可操作输入框 | 早失败探测：`DOUBAO_READY` 后 T17 上限内探不到输入框立即中止（提示可能未登录/界面异常/accessibility 未生效）。**按 T12 只匹配类名 `*ProseMirror*`，不得附加 ControlType 条件**；界面明显已登录时先重取窗口，仍失败用 T2 参数重启豆包 |
| 上传 | 图片粘贴后未出现附件 | 重新聚焦输入框并重试粘贴 |
| 上传 | 文件过大、网络中断 | 自动重试 1 次，失败则提示上传失败 |
| 上传 | 单次粘贴超过 10 张，超出部分不入框（实测 13 张仅入框 10 张） | 按 1.3 分批：每批 ≤10 张、每批独立新对话；禁止单次粘贴 >10 张 |
| 识别 | 结果格式错误、内容为空 | 重发指令重试 1 次，失败则提示识别异常 |
| 识别 | 回复已生成但转储/校验失败（脚本中断、内容截断） | 「仅重抓」：不重传不重发，**先执行 PP2.1 恢复窗口展开**（最小化态 UIA 读不到任何元素），复用当前对话只读取重抓（上限见 T19）并补校验，完成后最小化收尾 |
| 识别 | 结果混入任务外的图片/旧会话内容（回复组数 > 本批图片数，或主区域有旧回复行） | "新对话"未生效，中止；重跑 PP2.3 并校验纯净后再发指令 |
| 合规 | 图片含明确违规敏感内容 | 终止处理，仅输出「无法识别该图片」 |
| 窗口 | 误开应用商店/无关窗口 | 按步骤8-5 尝试关闭；权限不足则提示用户手动关闭 |
| 窗口 | 豆包进程在运行但窗口最小化/离屏（rect 为负值如 -21333），`AppActivate` 未生效 | 执行 PP2.1 恢复激活（SW_RESTORE + 几何校验 + 前台校验循环）；仍失败重试 1 次后报错 |
| 窗口 | 快捷键/粘贴/点击未生效（窗口不在前台，操作落到其他窗口） | 任何 SendKeys、剪贴板粘贴、坐标点击前先执行 PP2.1 并确认 `DOUBAO_READY`；坐标点击前先移动光标；快捷键后校验标题变化 |
| 窗口 | 原窗口已关闭或无法还原 | 返回当前任务主界面，不阻塞结果输出 |
| 追问 | 追问后豆包无响应或结果为空 | 重发追问 1 次，仍失败则基于已有结果回答并标注未确认项 |

---

## 附录C 约束索引

> 每条约束在正文只有一个权威落点，此处仅做速查，避免重复定义导致改一处漏一处。

| # | 约束 | 落点 |
|---|---|---|
| 1 | 用户主动触发或任务需要识别图片时可直接调用豆包，不再额外询问授权 | 1.1 |
| 2 | 固定指令必须完整发送，不得删减规则或格式 | 附录A |
| 3 | 豆包原始结果仅作事实输入；可解读补充，但不得伪装成豆包原文 | 步骤6 |
| 4 | 解读补充必须与原文明确区分，且不得脱离识别事实臆测 | 步骤6 |
| 4b | **图片回复必须是两部分：第一部分客观描述（8 字段，纯事实）+ 第二部分推想预测（5 字段，含置信度/依据/其他可能）；两部分都必须拿到，缺第二部分先追问补齐，不整条重发** | 1.4 / 5-3 / 步骤7-1 |
| 4c | **第二部分以给出明确判断为主（必答，不得用「无法判断」规避）；结论必须可溯源到第一部分的可见特征并在「判断依据」中列出；汇报时推断必须分区标注，不得升格为事实** | 1.4 / 步骤6 |
| 4d | 置信度按"明确可识别特征/依据充分但有其他解释/仅靠风格概率"三档定级，**先给结论再定级，不得习惯性压低、也不得给无依据的「高」** | 1.4 |
| 4e | 完成判据的字段数阈值必须按 **13 字段**（8+5）；T10 锚定第二部分末字段，禁止沿用第一部分末字段 | 5-1 / 1.5 T10 |
| 4f | **用户点明的角度优先**：优先级为「用户指定角度 > 本次工作需要 > 对照表默认值」；用户指定角度时不得改回默认或自行判断的方向。但 P1 只改"看什么、怎么答"，**不得借此省略字段或绕过格式** | 1.2 / 步骤4 步骤0 / A.1a |
| 4g | **识图结果以豆包为准**：汇报只分两层（豆包两部分），**不得做二次判断**、不得用自己的保留意见削弱结论、不得只报一部分；只允许改形式（表格/列表/去重） | 步骤6 |
| 4h | **防信息遗漏**：可辨识个体必须逐个列出，禁止用"等/其他/众多"合并；密集群像须给可计数信息；发现遗漏用模板 D 追补，不自行补写 | 1.4 / 步骤4 指令规则5 / 步骤7-1 |
| 4i | **提速只压固定等待，不压判据**：UI 就绪一律自适应轮询；粘贴间隔与标记确认延时不得为提速而缩短 | T15 |
| 4j | **基线类变量必须在本进程内采集**：跨进程（步骤4 → 5-1）变量传不过来，实测导致基线恒为 0 | 5-1 |
| 5 | 文件操作一律用绝对路径，禁止把 `.` 或相对路径交给文件工具 | 1.1 C1 |
| 6 | 每次识别/追问结束后必须清理本次临时文件，且清理必须先于最终汇报；失败/中断路径同样清理 | 步骤8-1 |
| 7 | 识别完成即抢读落盘并立即最小化，保留会话、不关闭窗口 | 步骤5-1 / 步骤8-3 |
| 8 | 首次上传前必须处于"新对话"；追问必须复用同一对话，不重开、不重传 | PP2.3 / 7.2 |
| 9 | 只有确认用户不再需要图片信息时才关闭豆包窗口 | 步骤8-4 |
| 10 | 最终输出中每张图片必须附可点击本地链接；临时产物链接须注明「临时页图，任务结束已清理」；要长期可点击则先转存到用户指定位置，禁止为保链接把产物留在临时目录 | 步骤6 / 3.4 |
| 11 | 所有发给豆包的文字必须走剪贴板粘贴，禁止 SendKeys 逐字输入中文 | P3 |
| 12 | 追问文本必须针对图片细节生成，不得发散 | 7.1 |
| 13 | 本 skill 仅用于本地图片/PPT 识别场景，不得超范围使用 | 0.1 |
| 14 | 轮询阶段保持窗口展开，500 ms 间隔高频轮询；**完成即"抢读落盘 → 立即最小化"**（最小化后 UIA 读不到，故抢读必须在最小化之前） | 5-1 / 5-4 |
| 15 | 所有依赖快捷键/粘贴/坐标点击的操作，执行前必须先确认 `DOUBAO_READY`；最小化时 UIA 坐标无效；`mouse_event` 前先移动光标 | PP2.1 / P3 |
| 16 | 识别前必须确认会话纯净：无旧消息痕迹、附件数 == 本批张数；不一致则中止重来 | 3.3 |
| 17 | 单批附件数 ≤10；>10 必须分批，每批独立新对话完成「上传→校验→发送→读取→记录」 | 3.3 |
| 18 | 防误删：只清理命名约定内且内容核对无误的本流程产物；拿不准就保留并报告 | 步骤8-2 |
| 19 | PPT 任务：按 1.2 分流（文本直传 / 画面拆解）；产物目录收尾必删，不删原始 PPT | 3.4 / 步骤8-1 |
| 20 | PPT 文本分析汇报必须带来源页码标注，禁止编造页码，推断须注明「推断」 | 3.1 步骤7 |
| 21 | PPT 文本分析汇报层级标题只用中文序号；「核心重点汇总」必须条目化带〔页码〕，禁止概括空话 | 3.1 步骤7 |
| 22 | PPT 文本分析汇报排版：「分章节内容拆解」逐页呈现；汇总类条目按页码升序、每条独立成行 | 3.1 步骤7 |
| 23 | 每段代码独立进程执行，跨段复用变量前先重建上下文 | 0.6 / P0 |

---

## 附录D 运行验证清单

改动本 skill 后，至少完成以下验证再投入使用：

| # | 验证项 | 通过标准 |
|---|---|---|
| V1 | 语法自检 | 各代码块可被 PowerShell 解析（`[System.Management.Automation.Language.Parser]::ParseInput`）无错误；here-string 起止符用**词法栈**校验成对且无交叉（正则不可靠：`@'` 可出现在行中，`'@` 独占行首） |
| V2 | 锚点完整性 | T9（13 字段）/T10（两部分末字段）、`识别完毕`、`DOUBAO_READY`、`--force-renderer-accessibility`、`*image-container*`、`*ProseMirror*` 全部保留且与固定指令一致 |
| V2b | **版本漂移回归（每次实机首跑必做）** | 三项 MUST：① 输入框能定位到（`*ProseMirror*`，不限 ControlType）；② 附件数 == 本批张数（`image-container｜image-wrapper`）；③ **AI 未回复时字段行统计必须为 0**（用 T9 正则扫全窗口，若 >0 说明 `[^（(]` 式排除法已失效）。任一不成立即按 CHANGELOG 的 R1–R3 复核 |
| V2c | **两部分契约一致性** | 固定指令字段名 == 1.4 契约 == T9 正则 == T14 分组，四处逐字一致（13 字段）；**T10 必须指向第二部分末字段 `无法推断项说明`**，否则 5-1 会在第一部分刚写完时提前判完成、丢掉第二部分 |
| V2d | **槽位与优先级** | 组装产物中"字段名：（"行恰为 **13**（自检已内置）；8 个槽位非空；`$depth` ∈ 简/标准/详细；**用户若点明角度，槽位内容必须体现该角度**（不得回落默认措辞） |
| V2e | **防遗漏与只报不判** | 指令含"禁止遗漏个体"硬性规则（含"等/其他/众多"禁用清单）；汇报中不出现智能体自行添加的质疑或"我认为证据更弱"类点评；两部分与每张图都完整呈现 |
| V2f | **基线与提速回归** | 5-1 的 `$markerBase` **必须在轮询进程内采集**（不得依赖步骤4 传值）；基线采集要"等消息渲染稳定"（连续两轮文本量+标记计数不变，窗口上限 8 s；未收敛必须告警 BASE_NOT_STABLE）；所有 UI 就绪等待均为自适应轮询而非固定 sleep；T7=450 ms 与 T6=1.5 s **未被为提速而缩短** |
| V3 | 分支一致性 | 1.2 三个分支分别能追到「步骤0.6 / 步骤0 / 图片链路」的完整执行序列，且 0.4 总览与正文步骤号一一对应 |
| V4 | 常量一致性 | 正文出现的数值（10 张、450ms、500ms、1.5s、3s、150s、240s、30s）与 1.5 常量表完全一致 |
| V5 | 契约一致性 | 1.4 的完成信号、字段三态、失败态与 5-1 / 5-3 / 5-4 的描述互不矛盾 |
| V6 | 端到端冒烟 | 实机跑 1 张图片：`DOUBAO_READY` → 附件数 == 1 → **两部分 13 字段齐（8 客观 + 5 推断）** → 标记命中 → 落盘校验 → 最小化 → 清理无残留；且汇报中两部分分区呈现 |

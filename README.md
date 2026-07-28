# OneStopPPT

一个面向中文研究型汇报的 PPT 自动化仓库。  
仓库只保留脚本、skills 和 `scientific-illustrator` 能力；业务材料请只放本地，不要上传 git。

## 你能拿它做什么

- 根据本地访谈/研究材料生成 PPT
- 套用研究报告型风格继续改稿
- 用 `scientific-illustrator` 做 live 检查、导出预览、局部修页

## 先看哪里

- `/.trae/skills/avatr-report-ppt`
- `/.trae/skills/project-report-ppt-style`
- `/build-interview-report/build_benchmark_five_refs_safe.mjs`
- `/scientific-illustrator/README.md`

## 快速开始

```bash
git clone https://github.com/Cathy011011/OneStopPPT.git
cd OneStopPPT
cd build-interview-report
npm install
```

## 需要的工具

必需：

- `Node.js 18+`
- `npm`
- `Python 3.11+`
- `python-pptx`
- `Microsoft PowerPoint`
- `Git`

建议：

- `GitHub CLI (gh)`：方便创建仓库、登录、推送
- 支持 skills 的 Trae / Codex / Agent 环境
- `PowerShell 5.1+`：便于跑本地脚本和 PowerPoint 流程

生成脚本默认输出到本地目录：

```bash
node build_benchmark_five_refs_safe.mjs
```

指定输出路径：

```bash
$env:PPT_OUT="D:\temp\report.pptx"
node build_benchmark_five_refs_safe.mjs
```

如果是从原始访谈 `docx` 做“脱敏整理 -> 研究报告版 PPT”，建议在本地按现有生成器模式新建私有脚本，不把业务文案和客户材料相关脚本提交到 git。

## 目录

```text
OneStopPPT/
├─ .trae/skills/
│  ├─ avatr-report-ppt/
│  ├─ project-report-ppt-style/
│  └─ powerpoint-automation/
├─ build-interview-report/
└─ scientific-illustrator/
```

## 使用原则

- 业务 PPT、docx、图片只保存在本地
- 不补造数字、结论、用户原话
- 优先 PowerPoint 原生可编辑对象
- 风格冲突时，以 `project-report-ppt-style` 为准

## 提示词放在哪用

### 1. 放在 Agent 对话框里

适用：

- `avatr-report-ppt`
- `project-report-ppt-style`
- `scientific-illustrator`

使用位置：

- Trae / Codex / 支持 skills 的 Agent 聊天窗口

不要放这里：

- PowerPoint 里
- `node` 命令行里
- 普通 shell 终端里

### 2. 放在终端里

适用：

- `node build_benchmark_five_refs_safe.mjs`
- PowerShell 脚本
- git 命令

### 3. 怎么选

- 要“生成初稿 / 改稿” -> 用 `avatr-report-ppt`
- 要“先学参考稿风格” -> 用 `project-report-ppt-style`
- 要“打开当前 PPT 做 live 检查/修页” -> 用 `scientific-illustrator`

## 最新验证流程

适用场景：

- 本地只有访谈 `docx`
- 先要脱敏，再要汇总，再要生成研究报告型 PPT

推荐顺序：

1. 用 `avatr-report-ppt` + `project-report-ppt-style` 明确结构和风格约束
2. 先把原始纪要整理成 `_脱敏汇总材料.md`
3. 用本地私有生成脚本或现有生成器模式生成本地 PPT
4. 用 PowerPoint 导出预览图，复查标题、摘要块、卡片说明、结尾页是否溢出

## 直接可用的提示词

### 1. 根据访谈材料生成初版 PPT

```text
把下面这段贴到 Agent 对话框里使用。

使用 avatr-report-ppt，并参考 project-report-ppt-style。
我会在本地提供访谈纪要和参考 PPT，请基于本地材料生成一版研究报告型 PPT。
要求：
1. 只使用本地私有材料，不要引用仓库里不存在的文件
2. 标题直接写结论
3. 不补造数据和用户原话
4. 优先使用 PowerPoint 原生可编辑对象
5. 生成后给出逐页结构摘要
```

推荐场景：

- 从 0 到 1 出第一版
- 已有访谈纪要和参考 PPT
- 需要“内容生成 + 风格约束”一起生效

### 2. 在现有 PPT 上继续精修

```text
把下面这段贴到 Agent 对话框里使用。

使用 avatr-report-ppt。
基于我本地现有 PPT 继续精修，不要推翻重做。
重点检查：
1. 标题是否直接表达结论
2. 一页是否只服务一个重点
3. 图表和表格是否比大段文字更承担信息主体
4. 文案是否忠于原始材料
5. 所有修改保持可编辑
```

推荐场景：

- 已经有一版 PPT
- 现在主要是修标题、修文案、修结构

### 2.1 先脱敏，再生成研究报告版 PPT

```text
把下面这段贴到 Agent 对话框里使用。

同时使用 avatr-report-ppt 和 project-report-ppt-style。
我会提供本地访谈 docx，请先删除客户敏感信息，再整理成研究报告型 PPT 的内容基础。
要求：
1. 先输出一份脱敏汇总材料，再开始生成 PPT
2. 删除或泛化姓名、具体职业、可识别经历、精确时间
3. 标题直接写结论，不补造用户原话
4. 优先使用 PowerPoint 原生可编辑对象
5. 生成后导出预览，并检查是否有文字溢出或遮挡
```

推荐场景：

- 只有原始纪要，还没有整理好的汇报底稿
- 需要把“脱敏”和“PPT 生成”连成一次完整流程

### 3. 用 scientific-illustrator 做 live 审查

```text
把下面这段贴到 Agent 对话框里使用。

使用 scientific-illustrator 的 live PowerPoint 流程检查当前打开的 PPT。
按下面顺序执行：
1. status
2. capabilities
3. inspect
4. 对关键页做 audit
5. 导出预览图

重点看：
- 对齐
- 留白
- 文字溢出
- 箭头和边框净空
- 是否仍是原生可编辑对象
```

推荐场景：

- PPT 已经打开
- 需要检查溢出、对齐、可编辑性
- 需要导出预览图

### 4. 让它只改文案，不动版式骨架

```text
把下面这段贴到 Agent 对话框里使用。

使用 avatr-report-ppt。
只改文案，不大改版式结构。
保留当前页面骨架、图表位置、信息层级和主视觉节奏。
只做：
1. 标题改成结论式
2. 正文压缩
3. 重复表达去重
4. 语气统一
```

推荐场景：

- 页面结构已经定了
- 只想收文案，不想重排版

### 5. 让它只做风格学习，不直接出 PPT

```text
把下面这段贴到 Agent 对话框里使用。

使用 project-report-ppt-style。
先不要生成 PPT，先把参考稿的风格规则整理出来。
输出内容只要：
1. 字体和字号层级
2. 常见页型
3. 图表/表格表达规则
4. 标题写法
5. 后续复用时必须遵守的约束
```

推荐场景：

- 还没开始做 PPT
- 先想把参考风格吃透
- 后续准备多人复用同一套风格规则

## 提交前

```bash
git status
```

确认这些材料没有被跟踪：

- 本地业务 PPT
- 本地 docx
- 本地参考图片
- 本地 course 截图和临时屏幕文件

## 一句话提醒

这个仓库传的是能力，不是材料。

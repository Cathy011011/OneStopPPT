# OneStopPPT

一个用于“参考模板学习 + 访谈归纳 + 可编辑 PPT 生成”的本地项目。

仓库只保留能力、skill、脚本和说明文档；访谈纪要、客户材料、参考 PPT、缩略图和交付文件只放本地，不上传 git。

## 这个项目现在固定做什么

围绕一条可复刻流程工作：

1. 准备参考 PPT
2. 抽取参考 PPT 缩略图
3. 整理访谈资料
4. 用固定提示词方法归纳访谈
5. 按参考模板配色和页型生成**可编辑 PPT**
6. 导出 PowerPoint 预览，检查字号、遮挡和版面

## 本地标准目录

`访谈/` 目录现在统一按下面方式组织：

```text
访谈/
├─ 01_参考模板/
│  ├─ 租车战略项目调研报告20190531(2).pptx
│  └─ 20190531-template-thumbnails/
├─ 02_访谈资料/
│  ├─ 归纳总结的提示词(1).docx
│  ├─ 归纳总结的提示词(1).extracted.txt
│  ├─ 座谈会笔录2-7月13日.docx
│  ├─ 座谈会笔录2-7月13日.extracted.txt
│  └─ 座谈会笔录2-7月13日_按提示词归纳总结.md
├─ 03_交付成品/
│  ├─ 座谈会笔录2-7月13日_参考模板可编辑版.pptx
│  └─ 0713-reference-editable-preview/
└─ 归档/
```

说明：

- `01_参考模板`：只放参考 PPT 和参考缩略图
- `02_访谈资料`：只放原始访谈、正文提取稿、提示词文档和归纳稿
- `03_交付成品`：只放当前要交付的最终可编辑版 PPT 和预览图
- `归档`：旧版本 PPT、旧预览、旧 HTML 产物和中间试验稿

## 当前项目保留的关键文件

Skills：

- `/.trae/skills/interview-summary-workflow`
- `/.trae/skills/avatr-report-ppt`
- `/.trae/skills/project-report-ppt-style`
- `/.trae/skills/powerpoint-automation`

脚本：

- `/build-interview-report/build_0713_reference_editable_report.py`
- `/build-interview-report/build_benchmark_five_refs_safe.mjs`

## 每个 skill 什么时候用

### 1. `interview-summary-workflow`

用途：

- 把原始访谈纪要整理成结构化总结

什么时候用：

- 用户给的是 `docx` / 笔录 / 逐字稿
- 还没有成型的 PPT 底稿
- 需要先做“拆段 -> 证据 -> 归纳 -> 成稿”

### 2. `avatr-report-ppt`

用途：

- 把访谈归纳稿映射到本地参考模板，生成或精修可编辑 PPT

什么时候用：

- 已经有 `*_按提示词归纳总结.md`
- 用户明确给了参考 PPT
- 需要继续改 PPT 或出新版本

### 3. `project-report-ppt-style`

用途：

- 学习和整理研究报告型 PPT 的写法、页型和图表规则

什么时候用：

- 先学风格，不急着出稿
- 想做章节页、摘要页、矩阵页、建议页的页型复用

### 4. `powerpoint-automation`

用途：

- 打开现有 PPT 做 COM 自动化、导出预览、检查溢出和版面

什么时候用：

- 已有 PPT 文件
- 需要真实 PowerPoint 渲染结果
- 需要检查字体、遮挡、留白、可编辑性

## 复刻整套流程怎么做

### 第一步：放好参考模板

把参考 PPT 放到：

- `访谈/01_参考模板/`

把它导出的缩略图也放到：

- `访谈/01_参考模板/20190531-template-thumbnails/`

### 第二步：放好访谈资料

把原始纪要和提示词文档放到：

- `访谈/02_访谈资料/`

至少准备：

- 一份原始访谈纪要
- 一份访谈归纳提示词

### 第三步：先归纳，不急着做 PPT

把下面这段放到 Agent 对话框里，使用 `interview-summary-workflow`：

```text
使用 interview-summary-workflow。
我会提供本地访谈纪要和访谈归纳提示词，请先按固定步骤整理成结构化研究总结。
要求：
1. 不按发言顺序做流水账
2. 先拆段，再抽证据，再归纳观点，最后成稿
3. 保留用户真实观点，区分主持人、品牌方和用户反馈
4. 对证据不足的地方明确写“需问卷进一步验证”
5. 输出一份可直接给 PPT 使用的归纳总结 markdown
```

预期输出：

- `*_按提示词归纳总结.md`

### 第四步：再按参考模板生成可编辑 PPT

把下面这段放到 Agent 对话框里，使用 `avatr-report-ppt`：

```text
使用 avatr-report-ppt，并先读取 interview-summary-workflow 的归纳结果。
请基于本地参考 PPT 和访谈归纳稿，生成一版可编辑 PPT。
要求：
1. 最终 PPT 必须贴近参考 PPT 的配色框架、章节节奏和页型骨架
2. 不补造用户原话、数据和竞品结论
3. 文本框、表格、流程、矩阵尽量保持 PowerPoint 原生可编辑对象
4. 生成后导出预览图，检查字号、遮挡和留白
```

预期输出：

- `访谈/03_交付成品/*.pptx`
- `访谈/03_交付成品/*preview/`

### 第五步：最后做真实渲染检查

把下面这段放到 Agent 对话框里，使用 `powerpoint-automation`：

```text
使用 powerpoint-automation。
请对当前生成的 PPT 做真实渲染检查。
重点检查：
1. 字体大小是否失衡
2. 是否有文字遮挡
3. 表格和矩阵是否过密
4. 标题是否过长
5. 当前版本是否仍是可编辑对象，而不是整页图片
```

## 当前默认输出

当前主交付文件：

- `访谈/03_交付成品/座谈会笔录2-7月13日_参考模板可编辑版.pptx`

当前预览目录：

- `访谈/03_交付成品/0713-reference-editable-preview/`

当前可编辑生成脚本：

- `build-interview-report/build_0713_reference_editable_report.py`

## 写 PPT 时必须守住的规则

- 只用本地私有材料，不引用仓库里不存在的业务文件
- 先做访谈归纳，再做 PPT
- 标题优先直接写结论
- 不补造数据、用户原话和竞品判断
- 参考模板明确时，优先贴近该模板的配色和页型
- 最终交付优先保持 PowerPoint 原生可编辑对象

## Git 约束

仓库中默认忽略：

- `访谈/`
- `course-assets/`
- `scripts/`
- 本地截图、临时日志和中间文件

提交前只检查能力和文档，不提交客户材料。

## 一句话说明

这个项目现在的目标很明确：**把“参考模板 -> 访谈归纳 -> 可编辑 PPT”这条链路固定下来，并且以后可以照着重复做。**

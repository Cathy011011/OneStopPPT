---
name: avatr-report-ppt
description: Create or review the current interview-analysis PPT using the canonical project-report style references and keep visible copy faithful to the approved source material.
---

# 访谈分析报告 PPT 智能体入口

当需要基于当前项目的访谈材料继续生成、修订或复核 PPT 时，使用这个 skill。

## Canonical sources

先读下面这组唯一权威资料，不要再从其他目录找重复副本：

1. `f:\Codex\PPT自动化\.trae\skills\interview-summary-workflow\SKILL.md`
2. `f:\Codex\PPT自动化\.trae\skills\project-report-ppt-style\SKILL.md`
3. `f:\Codex\PPT自动化\.trae\skills\project-report-ppt-style\references\project-report-style.md`
4. `f:\Codex\PPT自动化\.trae\skills\project-report-ppt-style\references\slide-patterns.md`
5. `f:\Codex\PPT自动化\.trae\skills\project-report-ppt-style\references\project-report-analysis.json`

如用户明确给了新的本地参考模板 PPT，还要同时读取：

6. `f:\Codex\PPT自动化\访谈\01_参考模板\租车战略项目调研报告20190531(2).pptx`
7. `f:\Codex\PPT自动化\访谈\01_参考模板\20190531-template-thumbnails\`

如需重新抽取参考稿基线，运行：

- `f:\Codex\PPT自动化\.trae\skills\project-report-ppt-style\scripts\analyze_ppt_style.ps1`

## Standard local structure

本项目后续统一按下面的本地目录组织，不再把输入、参考和成品混在一起：

1. `f:\Codex\PPT自动化\访谈\01_参考模板\`
2. `f:\Codex\PPT自动化\访谈\02_访谈资料\`
3. `f:\Codex\PPT自动化\访谈\03_交付成品\`
4. `f:\Codex\PPT自动化\访谈\归档\`

## Current project assets

- 当前可编辑生成脚本：`f:\Codex\PPT自动化\build-interview-report\build_0713_reference_editable_report.py`
- 当前主交付 PPT：`f:\Codex\PPT自动化\访谈\03_交付成品\座谈会笔录2-7月13日_参考模板可编辑版.pptx`
- 业务材料：仅保存在本地私有目录，不进入公开仓库

## Use this skill for

- 基于原始访谈资料继续生成或修订 PPT
- 已有归纳稿时，把内容映射到本地参考模板页型
- 依据参考模板 PPT 的配色、版式和页型继续精修可编辑版本
- 复核当前最新 PPT 是否仍忠于访谈归纳稿
- 处理本地私有业务材料，不在公开仓库中保存素材本体

## Working rules

- visible copy 必须忠于用户批准的访谈材料和归纳稿
- 在生成 PPT 前，优先先跑 `interview-summary-workflow`
- 不补造数字、结论、用户原话和竞品判断
- 优先使用 PowerPoint 原生可编辑对象，避免整页截图
- 当用户明确指定 `01_参考模板` 下的参考 PPT 时，最终成品要优先贴近该模板的配色框架、章节节奏和页型骨架
- 当本 skill 与 `.trae/skills/project-report-ppt-style` 的规则冲突时，以用户明确指定的本地参考模板优先；如果用户没指定新模板，再以 `project-report-ppt-style` 为准

## Working flow

当输入是原始访谈纪要而不是现成 PPT 时，优先按下面顺序工作：

1. 先在 `02_访谈资料` 中提取正文，不直接把原始 `docx` 当汇报稿
2. 使用 `interview-summary-workflow` 按“拆段 -> 证据 -> 归纳 -> 成稿”沉淀 `*_按提示词归纳总结.md`
3. 读取 `01_参考模板` 中的参考 PPT 及缩略图，确定封面、目录、章节页、摘要页、矩阵页、行动页等页型
4. 生成 `03_交付成品` 下的可编辑 PPT
5. 导出 PowerPoint 预览，检查字号、遮挡、留白、是否仍为原生对象

## QA additions

- 长标题、目录三栏、矩阵页表格、流程页步骤卡、收尾页主结论，是最容易溢出或失衡的区域，必须单独复核
- 如果出现“字体盖住别的字体”，优先按“压缩文案 -> 调整版面层级 -> 缩小字号 -> 增加文本框高度”修正
- PowerPoint 实际渲染结果优先于脚本里的静态坐标判断
- 如果当前版本已经是原生可编辑对象，不要为了省事重新改回整页图片

## Kept files

本目录只保留：

- `SKILL.md`
- `agents/openai.yaml`

不要再复制保存一套 references / scripts。

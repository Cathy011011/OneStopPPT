---
name: avatr-report-ppt
description: Create or review the current interview-analysis PPT using the canonical project-report style references and keep visible copy faithful to the approved source material.
---

# 访谈分析报告 PPT 智能体入口

当需要基于当前项目的访谈材料继续生成、修订或复核 PPT 时，使用这个 skill。

## Canonical sources

先读下面这组唯一权威资料，不要再从其他目录找重复副本：

1. `f:\Codex\PPT自动化\.trae\skills\project-report-ppt-style\SKILL.md`
2. `f:\Codex\PPT自动化\.trae\skills\project-report-ppt-style\references\project-report-style.md`
3. `f:\Codex\PPT自动化\.trae\skills\project-report-ppt-style\references\slide-patterns.md`
4. `f:\Codex\PPT自动化\.trae\skills\project-report-ppt-style\references\project-report-analysis.json`

如需重新抽取参考稿基线，运行：

- `f:\Codex\PPT自动化\.trae\skills\project-report-ppt-style\scripts\analyze_ppt_style.ps1`

## Current project assets

- 当前生成脚本：`f:\Codex\PPT自动化\build-interview-report\build_benchmark_five_refs_safe.mjs`
- 业务材料：仅保存在本地私有目录，不进入公开仓库

## Use this skill for

- 基于访谈材料继续出新页或改现有页
- 依据参考 PPT 的研究报告风格做内容整理
- 检查当前最新 PPT 是否仍忠于访谈材料
- 沿用现有脚本继续生成最新版本
- 处理本地私有业务材料，不在公开仓库中保存素材本体

## Working rules

- visible copy 必须忠于用户批准的访谈材料
- 不补造数字、结论、用户原话和竞品判断
- 继续优先使用 PowerPoint 原生可编辑对象
- 当本 skill 与 `.trae/skills/project-report-ppt-style` 的规则冲突时，以后者为准

## Working flow

当输入是原始访谈纪要而不是现成 PPT 时，优先按下面顺序工作：

1. 先在本地提取正文，不直接把原始 `docx` 当成汇报稿使用
2. 先做脱敏，把姓名、具体职业、可识别经历、精确时间等改成研究口径
3. 先沉淀一份 `_脱敏汇总材料.md`，再开始写 PPT
4. 生成完成后，必须导出预览或做 PowerPoint 渲染检查，确认没有文字溢出和遮挡

## QA additions

- 长标题、摘要页大结论块、指标卡说明、收尾页主结论，都是最容易溢出的区域，必须单独复核
- 如果脚本生成版出现字号压住其他文字，优先按“压缩文案 -> 缩小字号 -> 增加文本框高度”顺序修正
- PowerPoint 实际渲染结果优先于脚本里的静态坐标判断

## Kept files

本目录只保留：

- `SKILL.md`
- `agents/openai.yaml`

不要再复制保存一套 references / scripts。

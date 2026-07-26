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

- 原始访谈材料：`f:\Codex\PPT自动化\访谈\会议笔录分析总结_优化版(1).docx`
- 参考 PPT：`f:\Codex\PPT自动化\访谈\项目报告.pptx`
- 最新交付 PPT：`f:\Codex\PPT自动化\访谈\会议笔录分析总结_优化版_访谈分析报告版.pptx`
- 参考图：`f:\Codex\PPT自动化\访谈\_codex_work_项目报告复刻\reference-contact-1.jpg` 到 `reference-contact-5.jpg`
- 当前生成脚本：`f:\Codex\PPT自动化\build-interview-report\build_benchmark_five_refs_safe.mjs`

## Use this skill for

- 基于访谈材料继续出新页或改现有页
- 依据参考 PPT 的研究报告风格做内容整理
- 检查当前最新 PPT 是否仍忠于访谈材料
- 沿用现有脚本继续生成最新版本

## Working rules

- visible copy 必须忠于用户批准的访谈材料
- 不补造数字、结论、用户原话和竞品判断
- 继续优先使用 PowerPoint 原生可编辑对象
- 当本 skill 与 `.trae/skills/project-report-ppt-style` 的规则冲突时，以后者为准

## Kept files

本目录只保留：

- `SKILL.md`
- `agents/openai.yaml`

不要再复制保存一套 references / scripts。

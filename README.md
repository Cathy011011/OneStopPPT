# OneStopPPT

一个面向中文研究型汇报场景的 PowerPoint 自动化项目。当前仓库聚焦于一条明确工作流：基于本地私有材料进行 PPT 生成、修订与 live 审查，但公开仓库不再包含任何业务敏感素材。

## 项目目标

- 保留可复用的风格规则与智能体入口
- 保留稳定的生成脚本
- 保留 `scientific-illustrator` 这条 PowerPoint live 审查与修订链路
- 不在仓库中保存任何业务 PPT、原始文档、截图或参考图片

## 当前仓库包含什么

### 1. 生成脚本

- `build-interview-report/build_benchmark_five_refs_safe.mjs`
  - 当前保留的稳定版生成脚本
- `build-interview-report/package.json`
  - 脚本目录的基础 Node 配置

### 2. 技能与规则

- `.trae/skills/avatr-report-ppt`
  - 当前项目的智能体入口
- `.trae/skills/project-report-ppt-style`
  - 参考报告风格规则、页型模板与分析基线
- `.trae/skills/powerpoint-automation`
  - PowerPoint 自动化相关通用能力

### 3. PowerPoint live 能力

- `scientific-illustrator/`
  - 保留该目录是为了继续使用 live PowerPoint 编辑、结构检查、导出预览和审查流程

## 目录结构

```text
OneStopPPT/
├─ .trae/
│  └─ skills/
│     ├─ avatr-report-ppt/
│     ├─ powerpoint-automation/
│     └─ project-report-ppt-style/
├─ build-interview-report/
│  ├─ build_benchmark_five_refs_safe.mjs
│  └─ package.json
└─ scientific-illustrator/
```

## 工作原则

- PPT 可见内容必须忠于已批准的访谈材料
- 不补造数字、结论、用户原话和竞品判断
- 尽量使用 PowerPoint 原生可编辑对象，而不是大面积图片平铺
- 当智能体入口与 canonical 风格规则冲突时，以 `.trae/skills/project-report-ppt-style` 为准

## 快速开始

### 1. 克隆仓库

```bash
git clone https://github.com/Cathy011011/OneStopPPT.git
cd OneStopPPT
```

### 2. 准备环境

建议环境：

- Windows
- Microsoft PowerPoint
- Node.js 18+
- PowerShell 5.1 或 PowerShell 7+

如果要运行本地生成脚本，可先安装脚本目录依赖：

```bash
cd build-interview-report
npm install
```

### 3. 查看关键资产

优先阅读这些文件：

1. `.trae/skills/avatr-report-ppt/SKILL.md`
2. `.trae/skills/project-report-ppt-style/SKILL.md`
3. `build-interview-report/build_benchmark_five_refs_safe.mjs`
4. `scientific-illustrator/README.md`

## 推荐工作流

### 路线 A：继续修改已有 PPT

适用于当前项目的日常迭代。

1. 在本地私有目录准备业务材料
2. 以 `.trae/skills/avatr-report-ppt` 作为任务入口
3. 如需套用研究报告风格，参考 `.trae/skills/project-report-ppt-style`
4. 在本地环境修改或重生成目标 PPT
5. 如需 live 审查，用 `scientific-illustrator` 打开并复核

### 路线 B：重新抽取参考风格基线

当本地参考 PPT 更新后执行：

```powershell
powershell -ExecutionPolicy Bypass -File ".\.trae\skills\project-report-ppt-style\scripts\analyze_ppt_style.ps1"
```

输出会更新 `project-report-ppt-style/references` 下的风格分析基线。注意：参考 PPT 本体不应提交到公开仓库。

### 路线 C：继续使用稳定版脚本生成

当前保留脚本：

```bash
node build-interview-report/build_benchmark_five_refs_safe.mjs
```

说明：

- 脚本会围绕当前访谈材料和参考图生成最新版本的访谈分析 PPT
- 该脚本是目前仓库中保留的稳定生成入口
- 业务材料路径请在本地私有环境中自行维护，不要提交到远端

## 为什么保留 scientific-illustrator

这个目录不是历史残留，而是项目里仍然有价值的一环。它主要用于：

- 打开 PowerPoint live 实例
- 做结构检查
- 做可编辑性审查
- 导出预览图
- 在 PowerPoint 内进行局部修订

如果后续完全不再需要 live 审查链路，可以再考虑拆掉；当前仓库默认保留。

## 技能说明

### avatr-report-ppt

当前项目的主入口 skill，用于：

- 基于访谈材料继续出新页或改现有页
- 校验当前 PPT 是否仍忠于访谈材料
- 串联 canonical 风格规则与现有生成脚本
- 所有业务材料默认仅存在于本地私有目录

### project-report-ppt-style

用于沉淀和复用参考研究报告的风格系统，核心包括：

- 字体与层级
- 版式边界
- 图表与表格习惯
- 页型模板
- 风格基线数据

### powerpoint-automation

用于通用 PPT 自动化能力补充，例如：

- PPTX 生成
- 模板分析
- 内容替换
- 通用流程与检查参考

## 仓库整理策略

当前仓库已经做过一次收敛，原则是：

- 只保留一套 canonical skill 体系
- 只保留脚本、规则、插件与公开可共享说明
- 删除业务 PPT、访谈文档、参考图和其他敏感素材
- 删除历史版本、草稿、重复 references 与中间预览

如果后续继续整理，建议仍然遵守这个原则，避免仓库再次堆满临时文件。

## 提交与更新建议

日常更新建议流程：

```bash
git status
git add README.md .gitignore
git add .trae
git add build-interview-report/build_benchmark_five_refs_safe.mjs
git add scientific-illustrator
git commit -m "update OneStopPPT project assets and documentation"
git push
```

## 注意事项

- `build-interview-report/node_modules/` 已在 `.gitignore` 中忽略，不提交到仓库
- PowerPoint live 相关能力依赖本机 Office/COM 环境
- 业务 PPT、访谈纪要、参考截图等敏感文件只应保存在本地私有目录
- 若敏感文件已经推送到远端，仅在新提交里删除并不能从 Git 历史中消失

## License

当前仓库未单独声明新的顶层许可证。若需要开源发布，建议在确认素材与文档可公开后，再补充仓库级 `LICENSE`。

# 组内工作常用 Skills

用于组内日常研究材料准备，基于个人使用习惯和实际工作经验整理，后续会结合使用反馈持续更新和完善。

## Skills

| Skill | 用途 |
|---|---|
| [科研项目申报书写作](grant-proposal-writing/SKILL.md) | 起草和修改申报内容、分配成果指标、编写预算说明。 |
| [横向项目与研究方案PPT制作](research-proposal-ppt/SKILL.md) | 按领域整合多份素材，用统一模板制作汇报页面，并可直接修改pptx。 |

两个 Skill 分别放在仓库根目录下的同名文件夹中，可独立安装和使用。使用时以当前任务要求、项目指南和指定模板为准。

## 安装与使用

将需要的完整 Skill 文件夹复制到 `~/.codex/skills/` 或 `~/.claude/skills/`，保留其中的全部文件。

- 使用 `$grant-proposal-writing` 起草或修改申报材料。
- 使用 `$research-proposal-ppt` 策划、制作或修改横向项目与研究方案PPT。

PPT Skill 自带模板 `assets/project-report-template.pptx` 和读取、修改、检查PPT的脚本，脚本依赖 `python-pptx` 和 `Pillow`。环境中有 PowerPoint 渲染能力时可配合使用，以检查实际页面。

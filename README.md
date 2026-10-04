# pptskills · 当前工作流

本仓库当前使用 **PPT6：imagegen 图片审核 → 人工批准 → 原生可编辑 PowerPoint → 实际渲染验收**。

## 从这里开始

- 执行入口：[PPT6 SKILL.md](.codex/skills/ppt6/SKILL.md)
- 项目约束：[AGENTS.md](AGENTS.md)
- 审核状态：[review-state.md](.codex/skills/ppt6/references/review-state.md)
- 文案和真实素材：[copy-and-asset-rules.md](.codex/skills/ppt6/references/copy-and-asset-rules.md)

## 工作流程

1. 核对任务范围、原始资料与当前 GPT-6 模型；已有模板先渲染分析。
2. 用内置 imagegen 生成完整页面审核稿，放入任务的“审核”文件夹。保留准确文案、提示词、素材来源和生成记录。
3. 展示审核稿，等待用户明确批准对应页面；修改的页面回到图片审核，其他页面的批准继续有效。
4. 批准后按版式重建可编辑 PPT。文字、公式、表格和连接符使用原生对象；科研证据图使用真实论文原图，图注与来源可核验。
5. 在实际 PowerPoint 中渲染检查，并验证文本、连接符和图片对象确实可编辑或可独立替换。

## 必须遵守

- 图片审核使用 imagegen 完成整页；禁止程序排版整页冒充生成审核稿。
- 科研图、设备照片和第三方 logo 保留真实来源；不得把生成内容作为研究证据。
- 流程箭头使用原生直线/肘形连接符，并连接到节点；原创装饰图标使用 imagegen，独立嵌入。
- 整页审核图不能充当正式 PPT 背景来伪装可编辑性。
- 审核记录按页面保存，未批准页面不进入正式可编辑制作。
- 过程文件集中在 agent-workspace/，审核与交付集中在 task/。

## 安装与使用

将本仓库的 .codex/skills/ppt6/ 放入项目的 .codex/skills/，把 AGENTS.md 中的工作流约束用于该项目。调用 `$ppt6` 开始任务；技能保留显式调用设置。

模型与审核检查只依赖 Python 标准库。可选替代图片服务需要 Python 3.11+ 及本机已安装的 imagegen CLI，且必须获得用户明确选择；本仓库不包含服务凭据或账号配置。

```powershell
python .codex/skills/ppt6/scripts/check_model_gate.py --model-id <actual-runtime-model-id>
python .codex/skills/ppt6/scripts/validate_review_state.py --state task/demo/审核/review_state.json --phase build
```

## 工作区

```text
.codex/skills/ppt6/        # 当前技能、参考规则和验证脚本
agent-workspace/<task>/    # 代码、过程素材、提示词、日志、渲染与 QA（不上传）
task/<task>/审核/          # 审核图片、生成记录与批准状态（不上传）
task/<task>/               # 参考资料和正式交付物（不上传）
assets/                   # 本地稳定源素材（不上传）
```

## 旧流程

旧的 template-first-ppt、母版库、模块合同和相关工具已退出当前 main 的执行入口。原始版本保存在 [archive/template-first-20261004](https://github.com/hanhan761/pptskills/tree/archive/template-first-20261004) 分支及 Git 历史中；开始新任务请使用本页的 PPT6。

## 公开范围

此仓库只发布脱敏后的通用技能、规则和验证脚本。个人绝对路径、内部服务标识、账号凭据、任务资料、真实项目图片、审核记录和正式 PPT 均未纳入本次发布。详见 [PRIVACY.md](PRIVACY.md)。

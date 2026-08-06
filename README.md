# pptskills

一个“统一母版、模板优先”的 PowerPoint 工作流。公开版只保留已审核、可机械验证的母版与内容模块，以及模板研发、合同化、组合和验收流程。

## 公开内容

- 1 个脱敏后的通用蓝白母版家族；
- 39 个活动内容模块及容量/几何合同；
- 模板研发的 `draft → revise → approved/rejected → 合同化 → 验证 → active` 全流程；
- 只负责复制、挂载、替换、渲染和验证的机械工具；
- Codex skill：`.codex/skills/template-first-ppt/`。

## 隐私边界

公开模板中的原始文案、研究数据、人物照片、机构标识、备注、批注、作者信息、绝对路径和外部链接均已移除或替换为中性占位内容。仓库不包含任务成品、研究原稿、素材源文件或内部审核记录。

使用模板时，请把图片占位替换为真实、可追溯且有授权的素材，并重新填写准确图注和来源。

## 快速验证

```powershell
pip install -r requirements.txt
python script/template_ppt.py verify-master --input "模板/母版/通用蓝白/母版源.pptx" --family "模板/母版/通用蓝白/family.json"
```

验证某个模块：

```powershell
python script/template_ppt.py verify-module --module "模板/模块/内容/<模块>.module.json" --family "模板/母版/通用蓝白/family.json"
```

详细规则见 `AGENTS.md` 和 `.codex/skills/template-first-ppt/SKILL.md`。

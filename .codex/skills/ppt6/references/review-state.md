# PPT6 审核状态

## 状态机

```text
draft → needs_revision → draft
draft → approved
draft → rejected
approved → build → final_qa
```

`approved` 只表示人已经确认图片稿；它不是 PPT 已完成。任何页面进入 `needs_revision` 或 `rejected` 时，都必须回到图片稿阶段。

## 最小状态文件

```json
{
  "schema_version": 1,
  "run_id": "2026-09-20-demo",
  "model_gate": {
    "model_id": "gpt-6",
    "passed": true
  },
  "approval_scope": "all_pages",
  "pages": [
    {
      "page_id": "slide-001",
      "draft_image": "draft/slide-001.png",
      "status": "approved",
      "human_confirmation": "图片稿通过，可以开始制作 PPT",
      "confirmed_at": "2026-09-20T12:00:00+08:00"
    }
  ]
}
```

## 有效批准

批准必须来自用户或明确的人工审核记录，并且覆盖所有要进入 PPT 的页面。有效内容至少包括：

- 明确的通过词，例如“图片稿通过”“可以开始制作 PPT”“第 1 页和第 2 页通过”；
- 被批准的页码或 `all_pages`；
- 时间和必要的修改备注。

“看着可以”“先这样”“继续看看”不构成批准。没有明确批准时，PPT6 只能继续整理图片稿或提出问题。

## 验证命令

```powershell
python scripts/check_model_gate.py --model-id gpt-6
python scripts/validate_review_state.py --state review_state.json --phase build
```

构建阶段必须同时满足：`model_gate.passed == true`、所有页面 `status == approved`、每页有图片稿路径、人工确认和确认时间。

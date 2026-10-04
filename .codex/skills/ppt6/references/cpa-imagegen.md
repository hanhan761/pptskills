# 已授权的可选 imagegen 替代路径

仅当内置 imagegen 不可用，且用户明确选择已有的 OpenAI-compatible 图片服务时使用。失败时说明原因；不得自行切换服务或用代码绘图替代。

脚本从本机 Codex 配置读取用户指定 provider 的 base_url 和 auth.command。令牌仅在内存及子进程环境中使用，不打印、不写入仓库或日志。仓库不提供账号、服务地址或模型目录。

先查询所选服务的认证模型目录，使用其实际支持的图片模型 ID。不要假定示例名称或产品名称就是 API ID。安装的 imagegen CLI 必须存在；缺失时停止并说明依赖。

示例（替换占位参数）：

```powershell
python .codex/skills/ppt6/scripts/cpa_imagegen.py --provider <configured-provider> -- generate --model <verified-image-model> --prompt <accurate-page-prompt> --out agent-workspace/demo/generated/page.png
```

这一替代路径仍遵守 GPT-6 模型检查、完整页面图片审核、人工批准和原生可编辑重建。

# CPA imagegen fallback

Use this only when the built-in `image_gen` tool is not registered and the user has chosen the configured CPA provider path.

`scripts/cpa_imagegen.py` reads the provider's existing `auth.command` from the local Codex config, obtains a short-lived token, and invokes the installed official `imagegen` CLI with:

- `OPENAI_BASE_URL` set to the provider's `base_url`;
- `OPENAI_API_KEY` set only in the child process;
- no token written to the repository, logs, or generated asset metadata.

Example:

```powershell
python .codex/skills/ppt6/scripts/cpa_imagegen.py -- generate `
  --model gpt-image-2 `
  --prompt "A clean editorial visual explaining a template-first PowerPoint repository" `
  --size 1536x1024 `
  --quality high `
  --out output/imagegen/repository-overview.png
```

Use the exact image model ID returned by the provider's authenticated `/v1/models` endpoint. Do not assume that a product label such as “image2.5” is the API model ID. The current local CPA catalog advertises `gpt-image-2`.

This fallback does not register a native `image_gen` tool inside the Codex turn. It only provides a safe provider-compatible execution path. It also does not bypass PPT6's GPT-6 model gate or its human image-review gate.

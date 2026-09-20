#!/usr/bin/env python3
"""Run the installed imagegen CLI through an OpenAI-compatible CPA provider.

The provider token is obtained from the existing Codex model-provider auth
command and is passed only to the child process. It is never written to disk
or printed.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tomllib
from pathlib import Path
from typing import Any


def _die(message: str) -> "NoReturn":
    print(f"cpa_imagegen: {message}", file=sys.stderr)
    raise SystemExit(2)


def _default_config() -> Path:
    return Path.home() / ".codex" / "config.toml"


def _default_imagegen_script() -> Path:
    return (
        Path.home()
        / ".codex"
        / "skills"
        / ".system"
        / "imagegen"
        / "scripts"
        / "image_gen.py"
    )


def _load_provider(config_path: Path, provider_name: str) -> tuple[str, dict[str, Any]]:
    if not config_path.is_file():
        _die(f"Codex config not found: {config_path}")
    try:
        config = tomllib.loads(config_path.read_text(encoding="utf-8"))
    except Exception as exc:
        _die(f"could not read Codex config: {type(exc).__name__}")

    providers = config.get("model_providers", {})
    provider = providers.get(provider_name)
    if not isinstance(provider, dict):
        _die(f"provider not found in Codex config: {provider_name}")

    base_url = provider.get("base_url")
    auth = provider.get("auth")
    if not isinstance(base_url, str) or not base_url.strip():
        _die(f"provider has no base_url: {provider_name}")
    if not isinstance(auth, dict):
        _die(f"provider has no command auth configuration: {provider_name}")
    return base_url.rstrip("/"), auth


def _get_provider_token(auth: dict[str, Any]) -> str:
    command = auth.get("command")
    args = auth.get("args", [])
    if not isinstance(command, str) or not command:
        _die("provider auth.command must be a non-empty string")
    if not isinstance(args, list) or not all(isinstance(item, str) for item in args):
        _die("provider auth.args must be a list of strings")

    timeout_ms = auth.get("timeout_ms", 10_000)
    try:
        timeout_seconds = max(float(timeout_ms) / 1000.0, 0.1)
    except (TypeError, ValueError):
        timeout_seconds = 10.0

    try:
        result = subprocess.run(
            [command, *args],
            check=False,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        _die(f"provider auth command failed: {type(exc).__name__}")
    if result.returncode != 0:
        _die(f"provider auth command exited with code {result.returncode}")

    token = result.stdout.strip()
    if not token:
        _die("provider auth command returned an empty token")
    return token


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Run the installed imagegen CLI through a Codex-configured CPA provider."
    )
    parser.add_argument("--provider", default="cpa_matterswarm")
    parser.add_argument("--config", type=Path, default=_default_config())
    parser.add_argument("--imagegen-script", type=Path, default=_default_imagegen_script())
    args, forwarded = parser.parse_known_args(argv)
    if forwarded and forwarded[0] == "--":
        forwarded = forwarded[1:]
    if not forwarded:
        parser.error("append imagegen arguments, for example: -- generate --prompt ...")
    if not args.imagegen_script.is_file():
        _die(f"installed imagegen CLI not found: {args.imagegen_script}")

    base_url, auth = _load_provider(args.config, args.provider)
    token = _get_provider_token(auth)

    child_env = os.environ.copy()
    child_env["OPENAI_API_KEY"] = token
    child_env["OPENAI_BASE_URL"] = base_url
    child = [sys.executable, str(args.imagegen_script), *forwarded]
    return subprocess.call(child, env=child_env)


if __name__ == "__main__":
    raise SystemExit(main())

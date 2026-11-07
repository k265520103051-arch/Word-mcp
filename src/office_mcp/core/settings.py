from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass


@dataclass(frozen=True)
class Settings:
    allow_vba: bool
    unsafe_gateway: bool
    call_timeout: float
    busy_retry: float


def _flag(env: Mapping[str, str], key: str) -> bool:
    return env.get(key, "0").strip().lower() in ("1", "true", "yes")


def _num(env: Mapping[str, str], key: str, default: float) -> float:
    try:
        return float(env[key])
    except (KeyError, ValueError):
        return default


def load_settings(env: Mapping[str, str] | None = None) -> Settings:
    env = os.environ if env is None else env
    return Settings(
        allow_vba=_flag(env, "WORD_MCP_ALLOW_VBA"),
        unsafe_gateway=_flag(env, "WORD_MCP_UNSAFE_GATEWAY"),
        call_timeout=_num(env, "WORD_MCP_CALL_TIMEOUT", 60.0),
        busy_retry=_num(env, "WORD_MCP_BUSY_RETRY", 10.0),
    )

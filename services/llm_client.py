"""OpenRouter API client with support for multiple LLM models."""

from __future__ import annotations

import json
import logging
import re
from typing import Any

from openrouter import OpenRouter

from config import LLM_MAX_TOKENS, LLM_TEMPERATURE, OPENROUTER_API_KEY

logger = logging.getLogger(__name__)


class LLMClient:
    """Thin wrapper around the OpenRouter SDK."""

    def __init__(self, api_key: str | None = None):
        self._api_key = api_key or OPENROUTER_API_KEY
        if not self._api_key or self._api_key.startswith("sk-or-YOUR"):
            raise ValueError(
                "Set a valid OPENROUTER_API_KEY in .env or pass it explicitly."
            )

    def _make_client(self) -> OpenRouter:
        return OpenRouter(api_key=self._api_key)

    def chat(
        self,
        model: str,
        messages: list[dict[str, str]],
        *,
        temperature: float = LLM_TEMPERATURE,
        max_tokens: int = LLM_MAX_TOKENS,
    ) -> str:
        """Send a chat completion request and return the text response."""
        with self._make_client() as client:
            response = client.chat.send(
                model=model,
                messages=messages,
                temperature=temperature,
                max_tokens=max_tokens,
                stream=False,
            )
        content = response.choices[0].message.content
        logger.info("LLM %s returned %d chars", model, len(content))
        return content

    def chat_json(
        self,
        model: str,
        messages: list[dict[str, str]],
        *,
        temperature: float = LLM_TEMPERATURE,
        max_tokens: int = LLM_MAX_TOKENS,
    ) -> dict[str, Any]:
        """Send a chat request expecting a JSON response."""
        raw = self.chat(
            model, messages, temperature=temperature, max_tokens=max_tokens
        )
        raw = _extract_json_block(raw)
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            logger.error("Failed to parse JSON from LLM response:\n%s", raw[:500])
            raise


def _extract_json_block(text: str) -> str:
    """Strip markdown fences (```json ... ```) if present."""
    m = re.search(r"```(?:json)?\s*\n(.*?)```", text, re.DOTALL)
    if m:
        return m.group(1).strip()
    text = text.strip()
    if text.startswith("{") or text.startswith("["):
        return text
    # Try to find first { or [
    for ch in ("{", "["):
        idx = text.find(ch)
        if idx != -1:
            return text[idx:]
    return text

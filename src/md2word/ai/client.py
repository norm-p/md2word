"""
Provider-aware LLM client.

Set AI_PROVIDER to one of:
  anthropic    — Anthropic direct API
  azure        — Azure AI Foundry (Anthropic endpoint)
  bedrock      — AWS Bedrock  (requires: pip install anthropic[bedrock])
  vertex       — Google Vertex AI  (requires: pip install anthropic[vertex])
  openai       — OpenAI direct API  (requires: pip install openai)
  azure-openai — Azure OpenAI  (requires: pip install openai)
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass
class LLMClient:
    """Unified wrapper. All AI modules call client.complete() — never the raw SDK."""

    _sdk_client: Any
    _model: str
    _family: str  # "anthropic" | "openai"

    def complete(self, system: str, messages: list[dict], max_tokens: int = 8192) -> str:
        """Send a completion request; return the assistant text."""
        if self._family == "anthropic":
            # Use streaming to avoid SDK timeout for large max_tokens values
            text_parts: list[str] = []
            with self._sdk_client.messages.stream(
                model=self._model,
                max_tokens=max_tokens,
                system=system,
                messages=messages,
            ) as stream:
                for text in stream.text_stream:
                    text_parts.append(text)
            return "".join(text_parts)
        else:  # openai
            chat_messages = [{"role": "system", "content": system}] + messages
            response = self._sdk_client.chat.completions.create(
                model=self._model,
                max_tokens=max_tokens,
                messages=chat_messages,
            )
            return response.choices[0].message.content


def _repair_json_strings(text: str) -> str:
    """Escape literal newlines/tabs inside JSON string values.

    LLMs sometimes embed literal newlines inside xml_content strings rather
    than using \\n escape sequences, which makes the JSON invalid.
    This scanner replaces bare newlines and carriage returns inside string
    literals with their escape equivalents.
    """
    result: list[str] = []
    in_string = False
    i = 0
    while i < len(text):
        c = text[i]
        if c == "\\" and in_string:
            # Consume the escape sequence as-is
            result.append(c)
            i += 1
            if i < len(text):
                result.append(text[i])
                i += 1
            continue
        if c == '"':
            in_string = not in_string
            result.append(c)
        elif in_string and c == "\n":
            result.append("\\n")
        elif in_string and c == "\r":
            result.append("\\r")
        elif in_string and c == "\t":
            result.append("\\t")
        else:
            result.append(c)
        i += 1
    return "".join(result)


def parse_llm_json(raw: str) -> Any:
    """Strip markdown fences and parse JSON from an LLM response.

    Handles:
    - ```json...``` fences
    - Preamble text before the JSON array/object
    - Literal newlines embedded inside JSON string values (common when LLMs
      produce multi-line XML inside xml_content fields)
    """
    text = raw.strip()
    # Strip ```json ... ``` or ``` ... ```
    text = re.sub(r"^```(?:json)?\s*\n?", "", text, count=1)
    text = re.sub(r"\n?```\s*$", "", text, count=1)
    text = text.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Try repairing escaped newlines inside strings
    try:
        return json.loads(_repair_json_strings(text))
    except json.JSONDecodeError:
        pass

    # Try finding the first JSON array or object in case there's preamble text
    for prefix in ["[", "{"]:
        idx = text.find(prefix)
        if idx > 0:
            try:
                return json.loads(_repair_json_strings(text[idx:]))
            except json.JSONDecodeError:
                continue

    # Final attempt without repair (will raise with the original error)
    return json.loads(text)


def get_client_or_none() -> LLMClient | None:
    """Build an LLMClient if credentials are available, otherwise return None.

    Returns None (instead of raising) when AI_MODEL or provider-specific API
    keys are missing from the environment. Callers use this to detect whether
    AI-augmented steps should run or be skipped.
    """
    try:
        return get_client()
    except (KeyError, ValueError, ImportError):
        return None


def load_env() -> None:
    """Load .env from cwd (default search) and, as a fallback, from the
    directory containing the running executable. This makes the frozen exe
    pick up a sibling .env regardless of where the user invoked it from."""
    from dotenv import load_dotenv

    load_dotenv()
    try:
        import sys
        exe_dir = Path(sys.executable).resolve().parent
        env_path = exe_dir / ".env"
        if env_path.is_file():
            load_dotenv(env_path, override=False)
    except Exception:
        pass


def get_client() -> LLMClient:
    """Build and return an LLMClient from environment variables."""
    load_env()

    provider = os.getenv("AI_PROVIDER", "anthropic")
    model = os.environ["AI_MODEL"]

    if provider == "anthropic":
        import anthropic

        sdk = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
        return LLMClient(_sdk_client=sdk, _model=model, _family="anthropic")

    if provider == "azure":
        import anthropic

        sdk = anthropic.Anthropic(
            api_key=os.environ["AZURE_ANTHROPIC_KEY"],
            base_url=os.environ["AZURE_ANTHROPIC_ENDPOINT"],
        )
        return LLMClient(_sdk_client=sdk, _model=model, _family="anthropic")

    if provider == "bedrock":
        import anthropic

        sdk = anthropic.AnthropicBedrock(
            aws_access_key=os.environ["AWS_ACCESS_KEY_ID"],
            aws_secret_key=os.environ["AWS_SECRET_ACCESS_KEY"],
            aws_region=os.getenv("AWS_REGION", "us-east-1"),
        )
        return LLMClient(_sdk_client=sdk, _model=model, _family="anthropic")

    if provider == "vertex":
        import anthropic

        sdk = anthropic.AnthropicVertex(
            project_id=os.environ["VERTEX_PROJECT_ID"],
            region=os.getenv("VERTEX_REGION", "us-east5"),
        )
        return LLMClient(_sdk_client=sdk, _model=model, _family="anthropic")

    if provider == "openai":
        import openai

        sdk = openai.OpenAI(api_key=os.environ["OPENAI_API_KEY"])
        return LLMClient(_sdk_client=sdk, _model=model, _family="openai")

    if provider == "azure-openai":
        import openai

        sdk = openai.AzureOpenAI(
            api_key=os.environ["AZURE_OPENAI_KEY"],
            azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-01"),
        )
        return LLMClient(_sdk_client=sdk, _model=model, _family="openai")

    raise ValueError(
        f"Unknown AI_PROVIDER '{provider}'. "
        "Choose: anthropic | azure | bedrock | vertex | openai | azure-openai"
    )

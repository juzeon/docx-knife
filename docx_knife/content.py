"""Content-source resolution and normalization (Phase 5).

Resolves ``ContentItem`` entries into paragraph-expanded text before any DOM
mutation. Every source variant (literal, JSONPath, file, command) shares the
same visible-mode newline expansion and the same raw-mode passthrough, so the
executor never sees source-specific quirks.

Only :class:`ContentResolverConfig` is part of the public package surface; the
resolver function and normalization helpers are internal to the package.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from re import Pattern
from typing import Any, Literal

from jsonpath_ng.exceptions import JsonPathParserError
from jsonpath_ng.ext import parse as parse_jsonpath

from ._models import (
    ContentItem,
    ContentSourceCommand,
    ContentSourceFile,
    ContentSourceJsonPath,
)
from .errors import InvalidContentError

SourceKind = Literal["literal", "jsonpath", "file", "command"]

_MAX_STDOUT_BYTES = 1024 * 1024
_STDERR_PREVIEW = 200
_ARGV_PREVIEW_CHARS = 60


@dataclass(frozen=True, slots=True)
class ContentResolverConfig:
    """Configuration for content-reference resolution.

    ``workspace_root`` is the base directory for relative source paths and the
    working directory of command references. When :meth:`Document.open` builds
    the default config it uses the caller's current working directory so that
    ``ContentSourceFile(path="notes.txt")`` resolves the same way a shell user
    would expect; both the CWD and the source document's parent directory are
    granted read access via ``input_roots``. ``input_roots`` restricts file and
    JSONPath sources to a declared set of directories via
    :meth:`pathlib.Path.is_relative_to` after ``resolve(strict=True)``.
    ``command_env_allowlist`` names the parent environment variables that may
    be inherited by command references; anything else must be provided through
    the explicit ``env`` mapping on the reference.
    """

    workspace_root: Path
    input_roots: tuple[Path, ...]
    command_env_allowlist: tuple[str, ...] = ("PATH", "LANG", "LC_ALL")


@dataclass(frozen=True, slots=True)
class ResolvedItem:
    """Result of resolving a single ``ContentItem``.

    In visible mode ``paragraphs`` carries the newline-expanded paragraph list
    from :func:`expand_visible`. In raw mode ``paragraphs`` is a single-element
    tuple carrying the untouched XML fragment string; downstream phases parse
    and validate it as WordprocessingML.
    """

    paragraphs: tuple[str, ...]
    raw: bool
    source_kind: SourceKind


def resolve_items(
    items: Sequence[ContentItem],
    *,
    raw: bool,
    config: ContentResolverConfig,
) -> tuple[ResolvedItem, ...]:
    """Resolve every ``ContentItem`` before any DOM mutation.

    All items are prevalidated for exclusive source cardinality first; a single
    offending index aborts the batch before any source is touched. JSONPath
    documents are loaded once per unique ``source`` path within one call.
    """
    for idx, item in enumerate(items):
        has_lit = item.content_literal is not None
        has_ref = item.content_ref is not None
        if has_lit == has_ref:
            raise InvalidContentError(
                raw=raw,
                reason=(f"item[{idx}] must provide exactly one of content_literal or content_ref"),
            )
    json_cache: dict[Path, Any] = {}
    return tuple(
        _resolve_one(idx, item, raw=raw, config=config, json_cache=json_cache)
        for idx, item in enumerate(items)
    )


def _resolve_one(
    idx: int,
    item: ContentItem,
    *,
    raw: bool,
    config: ContentResolverConfig,
    json_cache: dict[Path, Any],
) -> ResolvedItem:
    if item.content_literal is not None:
        text = item.content_literal
        kind: SourceKind = "literal"
    else:
        ref = item.content_ref
        assert ref is not None  # noqa: S101 - narrowed by cardinality check
        if isinstance(ref, ContentSourceJsonPath):
            text = _resolve_jsonpath(ref, raw=raw, config=config, cache=json_cache)
            kind = "jsonpath"
        elif isinstance(ref, ContentSourceFile):
            text = _resolve_file(ref, raw=raw, config=config)
            kind = "file"
        elif isinstance(ref, ContentSourceCommand):
            text = _resolve_command(ref, raw=raw, config=config)
            kind = "command"
        else:  # pragma: no cover - exhaustive
            raise InvalidContentError(raw=raw, reason=f"item[{idx}]: unknown content_ref type")
    if raw:
        paragraphs: tuple[str, ...] = (text,)
    else:
        paragraphs = expand_visible(text)
    return ResolvedItem(paragraphs=paragraphs, raw=raw, source_kind=kind)


def _resolve_within_roots(
    config: ContentResolverConfig, path_str: str, *, raw: bool, kind: str
) -> Path:
    candidate = Path(path_str)
    if not candidate.is_absolute():
        candidate = config.workspace_root / candidate
    try:
        resolved = candidate.resolve(strict=True)
    except FileNotFoundError as exc:
        raise InvalidContentError(raw=raw, reason=f"{kind} source not found: {path_str}") from exc
    for root in config.input_roots:
        try:
            resolved_root = root.resolve(strict=True)
        except FileNotFoundError:
            continue
        if resolved.is_relative_to(resolved_root):
            return resolved
    raise InvalidContentError(
        raw=raw,
        reason=f"{kind} path escapes allowed input roots: {path_str}",
    )


def _resolve_jsonpath(
    ref: ContentSourceJsonPath,
    *,
    raw: bool,
    config: ContentResolverConfig,
    cache: dict[Path, Any],
) -> str:
    source_path = _resolve_within_roots(config, ref.source, raw=raw, kind="jsonpath")
    if source_path not in cache:
        try:
            cache[source_path] = json.loads(source_path.read_text(encoding="utf-8"))
        except UnicodeDecodeError as exc:
            raise InvalidContentError(
                raw=raw,
                reason=f"jsonpath source is not valid UTF-8: {ref.source}",
            ) from exc
        except json.JSONDecodeError as exc:
            raise InvalidContentError(
                raw=raw,
                reason=f"jsonpath source is not valid JSON: {ref.source}",
            ) from exc
    data = cache[source_path]
    try:
        expr = parse_jsonpath(ref.path)
    except JsonPathParserError as exc:
        raise InvalidContentError(
            raw=raw,
            reason=f"invalid jsonpath expression {ref.path!r}: {exc}",
        ) from exc
    matches = expr.find(data)
    if not matches:
        raise InvalidContentError(
            raw=raw,
            reason=(f"jsonpath returned no value: path={ref.path}, source={ref.source}"),
        )
    if len(matches) > 1:
        raise InvalidContentError(
            raw=raw,
            reason=(
                f"jsonpath returned {len(matches)} values; exactly one required: "
                f"path={ref.path}, source={ref.source}"
            ),
        )
    value = matches[0].value
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (str, int, float)):
        return str(value)
    raise InvalidContentError(
        raw=raw,
        reason=(f"jsonpath value must be scalar, got {type(value).__name__}: path={ref.path}"),
    )


def _resolve_file(ref: ContentSourceFile, *, raw: bool, config: ContentResolverConfig) -> str:
    path = _resolve_within_roots(config, ref.path, raw=raw, kind="file")
    try:
        return path.read_text(encoding=ref.encoding)
    except UnicodeDecodeError as exc:
        raise InvalidContentError(
            raw=raw,
            reason=f"file {ref.path} could not be decoded as {ref.encoding}",
        ) from exc
    except LookupError as exc:
        raise InvalidContentError(raw=raw, reason=f"unknown encoding {ref.encoding!r}") from exc


def _resolve_command(ref: ContentSourceCommand, *, raw: bool, config: ContentResolverConfig) -> str:
    argv = list(ref.argv)
    if not argv or not all(isinstance(arg, str) and arg != "" for arg in argv):
        raise InvalidContentError(
            raw=raw, reason="command argv must be a non-empty list of strings"
        )
    cwd = _resolve_command_cwd(ref, raw=raw, config=config)
    env: dict[str, str] = {
        key: os.environ[key] for key in config.command_env_allowlist if key in os.environ
    }
    if ref.env is not None:
        env.update({str(k): str(v) for k, v in ref.env.items()})
    try:
        completed = subprocess.run(
            argv,
            cwd=cwd,
            env=env,
            capture_output=True,
            shell=False,
            check=False,
            timeout=ref.timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        raise InvalidContentError(
            raw=raw,
            reason=(f"command timed out after {ref.timeout_seconds}s: argv={_argv_preview(argv)}"),
        ) from exc
    except FileNotFoundError as exc:
        raise InvalidContentError(
            raw=raw, reason=f"command executable not found: {argv[0]!r}"
        ) from exc
    if completed.returncode != 0:
        stderr_preview = completed.stderr.decode("utf-8", errors="replace")[:_STDERR_PREVIEW]
        raise InvalidContentError(
            raw=raw,
            reason=(
                f"command exited with code {completed.returncode}: "
                f"argv={_argv_preview(argv)} stderr={stderr_preview!r}"
            ),
        )
    stdout: bytes = completed.stdout
    if len(stdout) > _MAX_STDOUT_BYTES:
        raise InvalidContentError(
            raw=raw,
            reason=(
                f"command stdout exceeds {_MAX_STDOUT_BYTES} bytes: argv={_argv_preview(argv)}"
            ),
        )
    try:
        text = stdout.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise InvalidContentError(
            raw=raw,
            reason=(f"command stdout is not valid UTF-8: argv={_argv_preview(argv)}"),
        ) from exc
    if text.endswith("\n"):
        text = text[:-1]
    return text


def _resolve_command_cwd(
    ref: ContentSourceCommand, *, raw: bool, config: ContentResolverConfig
) -> Path:
    workspace = config.workspace_root.resolve(strict=True)
    if ref.cwd is None:
        return workspace
    candidate = Path(ref.cwd)
    if not candidate.is_absolute():
        candidate = workspace / candidate
    try:
        resolved = candidate.resolve(strict=True)
    except FileNotFoundError as exc:
        raise InvalidContentError(raw=raw, reason=f"command cwd not found: {ref.cwd}") from exc
    if not resolved.is_relative_to(workspace):
        raise InvalidContentError(raw=raw, reason=f"command cwd escapes workspace: {ref.cwd}")
    return resolved


def _argv_preview(argv: Sequence[str]) -> str:
    joined = " ".join(argv)
    if len(joined) > _ARGV_PREVIEW_CHARS:
        joined = joined[: _ARGV_PREVIEW_CHARS - 1] + "\u2026"
    return joined


_PARAGRAPH_BREAK_RE = re.compile(r"\n{2,}")


def expand_visible(text: str) -> tuple[str, ...]:
    """Expand a visible-mode string into paragraph texts.

    Normalizes CRLF/CR to LF, splits on runs of two or more newlines to form
    paragraph boundaries, and leaves single embedded newlines in place for
    downstream ``<w:br/>`` projection.
    """
    unified = text.replace("\r\n", "\n").replace("\r", "\n")
    return tuple(_PARAGRAPH_BREAK_RE.split(unified))


_URL_RE: Pattern[str] = re.compile(r"https?://\S+")
_EMAIL_RE: Pattern[str] = re.compile(r"\S+@\S+\.\S+")
_CODE_SPAN_RE: Pattern[str] = re.compile(r"`[^`\n]*`")

_HALF_TO_FULL_PUNCT: dict[str, str] = {
    ",": "\uff0c",
    "?": "\uff1f",
    "!": "\uff01",
    ":": "\uff1a",
    ";": "\uff1b",
    "(": "\uff08",
    ")": "\uff09",
}


def normalize(text: str) -> str:
    """Deterministic Chinese punctuation and CJK/Latin spacing normalization.

    Off by default. Preserves the entire URL, email, and backtick code-span
    substrings byte-for-byte, and never trims leading or trailing whitespace.
    """
    if not text:
        return text
    protected = _find_protected_ranges(text)
    converted = _convert_punct(text, protected)
    return _apply_spacing(converted, _find_protected_ranges(converted))


def _find_protected_ranges(text: str) -> list[tuple[int, int]]:
    spans: list[tuple[int, int]] = []
    for pattern in (_URL_RE, _EMAIL_RE, _CODE_SPAN_RE):
        for match in pattern.finditer(text):
            spans.append(match.span())
    spans.sort()
    merged: list[tuple[int, int]] = []
    for start, end in spans:
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def _in_ranges(index: int, ranges: list[tuple[int, int]]) -> bool:
    return any(start <= index < end for start, end in ranges)


def _is_cjk(char: str) -> bool:
    return bool(char) and "\u4e00" <= char <= "\u9fff"


def _is_ascii_letter(char: str) -> bool:
    return bool(char) and (("a" <= char <= "z") or ("A" <= char <= "Z"))


def _is_ascii_digit(char: str) -> bool:
    return bool(char) and "0" <= char <= "9"


def _convert_punct(text: str, protected: list[tuple[int, int]]) -> str:
    out = list(text)
    length = len(text)
    for i, ch in enumerate(text):
        if _in_ranges(i, protected):
            continue
        left = text[i - 1] if i > 0 else ""
        right = text[i + 1] if i + 1 < length else ""
        if ch == ".":
            if _is_ascii_digit(left) or _is_ascii_digit(right):
                continue
            if _is_ascii_letter(left) or _is_ascii_letter(right):
                continue
            if _is_cjk(left) and (_is_cjk(right) or right == ""):
                out[i] = "\u3002"
        elif ch in _HALF_TO_FULL_PUNCT and (_is_cjk(left) or _is_cjk(right)):
            out[i] = _HALF_TO_FULL_PUNCT[ch]
    return "".join(out)


def _needs_cjk_latin_space(left: str, right: str) -> bool:
    latin_like_left = _is_ascii_letter(left) or _is_ascii_digit(left)
    latin_like_right = _is_ascii_letter(right) or _is_ascii_digit(right)
    return (_is_cjk(left) and latin_like_right) or (latin_like_left and _is_cjk(right))


def _apply_spacing(text: str, protected: list[tuple[int, int]]) -> str:
    length = len(text)
    out: list[str] = []
    i = 0
    while i < length:
        ch = text[i]
        out.append(ch)
        if _in_ranges(i, protected):
            i += 1
            continue
        j = i + 1
        while j < length and text[j] == " " and not _in_ranges(j, protected):
            j += 1
        if j >= length or _in_ranges(j, protected):
            out.extend(text[i + 1 : j])
            i = j
            continue
        gap = text[i + 1 : j]
        if _needs_cjk_latin_space(ch, text[j]):
            out.append(" ")
        else:
            out.extend(gap)
        i = j
    return "".join(out)


__all__ = [
    "ContentResolverConfig",
    "ResolvedItem",
    "SourceKind",
    "expand_visible",
    "normalize",
    "resolve_items",
]

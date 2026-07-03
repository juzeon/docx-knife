"""Tests for content-source resolution and normalization (Phase 5)."""

from __future__ import annotations

import json
import os
import platform
import sys
from pathlib import Path

import pytest

from docx_knife import (
    ContentItem,
    ContentResolverConfig,
    ContentSourceCommand,
    ContentSourceFile,
    ContentSourceJsonPath,
    InvalidContentError,
)
from docx_knife.content import (
    ResolvedItem,
    expand_visible,
    normalize,
    resolve_items,
)


def _config(tmp_path: Path, extra_roots: tuple[Path, ...] = ()) -> ContentResolverConfig:
    inputs = tmp_path / "inputs"
    inputs.mkdir(exist_ok=True)
    return ContentResolverConfig(
        workspace_root=tmp_path,
        input_roots=(inputs, *extra_roots),
    )


def _bypass_content_item(**kwargs: object) -> ContentItem:
    """Construct a ContentItem while sidestepping ``__post_init__``.

    ``ContentItem`` normally rejects zero-or-both source combinations at
    construction time. ``resolve_items`` runs the same check itself so a batch
    can abort with a stable message; the tests must be able to exercise that
    resolver-level branch.
    """
    obj = ContentItem.__new__(ContentItem)
    object.__setattr__(obj, "content_literal", kwargs.get("content_literal"))
    object.__setattr__(obj, "content_ref", kwargs.get("content_ref"))
    return obj


# ---------------------------------------------------------------------------
# 5.1 Exclusive source cardinality
# ---------------------------------------------------------------------------


def test_resolve_items_rejects_neither_source(tmp_path: Path) -> None:
    config = _config(tmp_path)
    item = _bypass_content_item()
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((item,), raw=False, config=config)
    assert "item[0]" in excinfo.value.reason


def test_resolve_items_rejects_both_sources(tmp_path: Path) -> None:
    config = _config(tmp_path)
    ref = ContentSourceFile(path="ignored.txt")
    item = _bypass_content_item(content_literal="x", content_ref=ref)
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((item,), raw=False, config=config)
    assert "item[0]" in excinfo.value.reason


def test_resolve_items_prevalidates_before_touching_sources(
    tmp_path: Path,
) -> None:
    """Invalid cardinality on any item aborts before any I/O runs."""
    config = _config(tmp_path)
    # The second item references a missing file. If cardinality validation
    # short-circuits before resolution the missing file is never opened.
    ok = ContentItem(content_literal="ok")
    bad = _bypass_content_item()
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ok, bad), raw=False, config=config)
    assert "item[1]" in excinfo.value.reason


def test_resolve_items_returns_literal_paragraphs(tmp_path: Path) -> None:
    config = _config(tmp_path)
    items = (ContentItem(content_literal="hello"),)
    result = resolve_items(items, raw=False, config=config)
    assert result == (ResolvedItem(paragraphs=("hello",), raw=False, source_kind="literal"),)


def test_resolve_items_raw_mode_wraps_single_string(tmp_path: Path) -> None:
    config = _config(tmp_path)
    fragment = '<w:p xmlns:w="urn:x"><w:r><w:t>x</w:t></w:r></w:p>'
    items = (ContentItem(content_literal=fragment),)
    result = resolve_items(items, raw=True, config=config)
    assert result[0].paragraphs == (fragment,)
    assert result[0].raw is True
    assert result[0].source_kind == "literal"


# ---------------------------------------------------------------------------
# 5.2 JSONPath source
# ---------------------------------------------------------------------------


def test_jsonpath_valid_single_match(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    (inputs / "data.json").write_text(json.dumps({"party": {"name": "Acme"}}), encoding="utf-8")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceJsonPath(source="inputs/data.json", path="$.party.name")
    item = ContentItem(content_ref=ref)
    result = resolve_items((item,), raw=False, config=config)
    assert result[0].paragraphs == ("Acme",)
    assert result[0].source_kind == "jsonpath"


def test_jsonpath_scalar_coercion(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    (inputs / "d.json").write_text(json.dumps({"n": 42, "b": True}), encoding="utf-8")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    result = resolve_items(
        (
            ContentItem(content_ref=ContentSourceJsonPath(source="inputs/d.json", path="$.n")),
            ContentItem(content_ref=ContentSourceJsonPath(source="inputs/d.json", path="$.b")),
        ),
        raw=False,
        config=config,
    )
    assert result[0].paragraphs == ("42",)
    assert result[1].paragraphs == ("true",)


def test_jsonpath_no_match(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    (inputs / "d.json").write_text(json.dumps({"a": 1}), encoding="utf-8")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceJsonPath(source="inputs/d.json", path="$.missing")
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "no value" in excinfo.value.reason


def test_jsonpath_multi_match(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    (inputs / "d.json").write_text(
        json.dumps({"items": [{"n": "a"}, {"n": "b"}]}), encoding="utf-8"
    )
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceJsonPath(source="inputs/d.json", path="$.items[*].n")
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "2 values" in excinfo.value.reason


def test_jsonpath_non_scalar_rejected(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    (inputs / "d.json").write_text(json.dumps({"party": {"name": "Acme"}}), encoding="utf-8")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceJsonPath(source="inputs/d.json", path="$.party")
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "scalar" in excinfo.value.reason


def test_jsonpath_source_outside_input_roots(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    outside = tmp_path / "outside"
    outside.mkdir()
    (outside / "d.json").write_text(json.dumps({"x": 1}), encoding="utf-8")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceJsonPath(source="outside/d.json", path="$.x")
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "escapes allowed input roots" in excinfo.value.reason


def test_jsonpath_source_cached_within_call(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    path = inputs / "d.json"
    path.write_text(json.dumps({"n": "first"}), encoding="utf-8")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceJsonPath(source="inputs/d.json", path="$.n")
    items = (ContentItem(content_ref=ref), ContentItem(content_ref=ref))
    # Rewrite the file between the two "loads" — cache should hide the change.
    original_read_text = Path.read_text
    calls: list[Path] = []

    def spy(self: Path, *args: object, **kwargs: object) -> str:
        calls.append(self)
        return original_read_text(self, *args, **kwargs)

    Path.read_text = spy  # type: ignore[method-assign]
    try:
        result = resolve_items(items, raw=False, config=config)
    finally:
        Path.read_text = original_read_text  # type: ignore[method-assign]
    assert [p.name for p in calls] == ["d.json"]
    assert [r.paragraphs[0] for r in result] == ["first", "first"]


def test_jsonpath_invalid_json(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    (inputs / "d.json").write_text("not json", encoding="utf-8")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceJsonPath(source="inputs/d.json", path="$.x")
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "valid JSON" in excinfo.value.reason


# ---------------------------------------------------------------------------
# 5.3 File source
# ---------------------------------------------------------------------------


def test_file_source_utf8(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    (inputs / "clause.txt").write_text("保密条款", encoding="utf-8")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceFile(path="inputs/clause.txt", encoding="utf-8")
    result = resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert result[0].paragraphs == ("保密条款",)
    assert result[0].source_kind == "file"


def test_file_source_gbk(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    payload = "保密条款".encode("gbk")
    (inputs / "clause.gbk.txt").write_bytes(payload)
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceFile(path="inputs/clause.gbk.txt", encoding="gbk")
    result = resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert result[0].paragraphs == ("保密条款",)


def test_file_source_wrong_encoding(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    (inputs / "clause.gbk.txt").write_bytes("保密条款".encode("gbk"))
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceFile(path="inputs/clause.gbk.txt", encoding="utf-8")
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "could not be decoded" in excinfo.value.reason


def test_file_source_missing(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceFile(path="inputs/missing.txt")
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "not found" in excinfo.value.reason


def test_file_source_path_traversal(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    outside = tmp_path / "secret.txt"
    outside.write_text("nope", encoding="utf-8")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceFile(path="inputs/../secret.txt")
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "escapes allowed input roots" in excinfo.value.reason


@pytest.mark.skipif(
    platform.system() == "Windows",
    reason="symlink creation requires elevated privileges on Windows",
)
def test_file_source_symlink_escaping_root(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    outside = tmp_path / "secret.txt"
    outside.write_text("nope", encoding="utf-8")
    link = inputs / "link.txt"
    try:
        link.symlink_to(outside)
    except (OSError, NotImplementedError):
        pytest.skip("filesystem does not support symlinks")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    ref = ContentSourceFile(path="inputs/link.txt")
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "escapes allowed input roots" in excinfo.value.reason


# ---------------------------------------------------------------------------
# 5.4 Command source
# ---------------------------------------------------------------------------


def test_command_success(tmp_path: Path) -> None:
    config = _config(tmp_path)
    ref = ContentSourceCommand(
        argv=(sys.executable, "-c", "print('hello')"),
        timeout_seconds=10.0,
    )
    result = resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    # Trailing newline from print() is stripped.
    assert result[0].paragraphs == ("hello",)
    assert result[0].source_kind == "command"


def test_command_shell_metacharacters_are_literal(tmp_path: Path) -> None:
    config = _config(tmp_path)
    ref = ContentSourceCommand(
        argv=(sys.executable, "-c", "import sys; print(sys.argv[1])", "a; echo hacked"),
        timeout_seconds=10.0,
    )
    result = resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert result[0].paragraphs == ("a; echo hacked",)


def test_command_timeout(tmp_path: Path) -> None:
    config = _config(tmp_path)
    ref = ContentSourceCommand(
        argv=(sys.executable, "-c", "import time; time.sleep(5)"),
        timeout_seconds=0.5,
    )
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "timed out" in excinfo.value.reason


def test_command_nonzero_exit(tmp_path: Path) -> None:
    config = _config(tmp_path)
    ref = ContentSourceCommand(
        argv=(sys.executable, "-c", "import sys; sys.stderr.write('boom'); sys.exit(3)"),
        timeout_seconds=10.0,
    )
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "exited with code 3" in excinfo.value.reason
    assert "boom" in excinfo.value.reason


def test_command_non_utf8_stdout(tmp_path: Path) -> None:
    config = _config(tmp_path)
    ref = ContentSourceCommand(
        argv=(
            sys.executable,
            "-c",
            "import sys; sys.stdout.buffer.write(b'\\xff\\xfe')",
        ),
        timeout_seconds=10.0,
    )
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "UTF-8" in excinfo.value.reason


def test_command_stdout_oversized(tmp_path: Path) -> None:
    config = _config(tmp_path)
    ref = ContentSourceCommand(
        argv=(
            sys.executable,
            "-c",
            "import sys; sys.stdout.buffer.write(b'x' * (1024*1024 + 10))",
        ),
        timeout_seconds=10.0,
    )
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "stdout exceeds" in excinfo.value.reason


def test_command_argv_empty_rejected(tmp_path: Path) -> None:
    config = _config(tmp_path)
    # ContentSourceCommand permits any tuple; a resolver-level check catches an
    # argv of empty strings before we ever hit ``subprocess``.
    ref = ContentSourceCommand(argv=("",), timeout_seconds=1.0)
    with pytest.raises(InvalidContentError) as excinfo:
        resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert "argv" in excinfo.value.reason


def test_command_cwd_outside_workspace(tmp_path: Path) -> None:
    outside = tmp_path.parent / "docx_knife_outside_cwd"
    outside.mkdir(exist_ok=True)
    try:
        config = _config(tmp_path)
        ref = ContentSourceCommand(
            argv=(sys.executable, "-c", "print('x')"),
            timeout_seconds=10.0,
            cwd=str(outside),
        )
        with pytest.raises(InvalidContentError) as excinfo:
            resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
        assert "escapes workspace" in excinfo.value.reason
    finally:
        outside.rmdir()


def test_command_env_allowlist_only(tmp_path: Path) -> None:
    config = ContentResolverConfig(
        workspace_root=tmp_path,
        input_roots=(tmp_path,),
        command_env_allowlist=("PATH",),
    )
    # A secret variable in the parent environment must not leak in.
    os.environ["DOCX_KNIFE_TEST_SECRET"] = "leaked"
    try:
        ref = ContentSourceCommand(
            argv=(
                sys.executable,
                "-c",
                "import os; print(os.environ.get('DOCX_KNIFE_TEST_SECRET', 'missing'))",
            ),
            timeout_seconds=10.0,
        )
        result = resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    finally:
        os.environ.pop("DOCX_KNIFE_TEST_SECRET", None)
    assert result[0].paragraphs == ("missing",)


def test_command_env_explicit_override(tmp_path: Path) -> None:
    config = _config(tmp_path)
    ref = ContentSourceCommand(
        argv=(
            sys.executable,
            "-c",
            "import os; print(os.environ.get('MY_VAR', 'missing'))",
        ),
        env={"MY_VAR": "explicit"},
        timeout_seconds=10.0,
    )
    result = resolve_items((ContentItem(content_ref=ref),), raw=False, config=config)
    assert result[0].paragraphs == ("explicit",)


# ---------------------------------------------------------------------------
# 5.5 Newline expansion
# ---------------------------------------------------------------------------


def test_expand_visible_mixed_newlines() -> None:
    text = "a\r\nb\nc\n\nd\n\n\ne"
    assert expand_visible(text) == ("a\nb\nc", "d", "e")


def test_expand_visible_no_newlines() -> None:
    assert expand_visible("hello") == ("hello",)


def test_expand_visible_only_paragraph_break() -> None:
    assert expand_visible("a\n\nb") == ("a", "b")


def test_expand_visible_identical_across_sources(tmp_path: Path) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    payload = "one\n\ntwo\nlinebreak"
    (inputs / "text.txt").write_text(payload, encoding="utf-8")
    (inputs / "data.json").write_text(json.dumps({"v": payload}), encoding="utf-8")
    config = ContentResolverConfig(workspace_root=tmp_path, input_roots=(inputs,))
    items = (
        ContentItem(content_literal=payload),
        ContentItem(content_ref=ContentSourceFile(path="inputs/text.txt")),
        ContentItem(content_ref=ContentSourceJsonPath(source="inputs/data.json", path="$.v")),
        ContentItem(
            content_ref=ContentSourceCommand(
                argv=(
                    sys.executable,
                    "-c",
                    "import sys; sys.stdout.write('one\\n\\ntwo\\nlinebreak\\n')",
                ),
                timeout_seconds=10.0,
            )
        ),
    )
    result = resolve_items(items, raw=False, config=config)
    expected = ("one", "two\nlinebreak")
    for resolved in result:
        assert resolved.paragraphs == expected


def test_expand_visible_raw_mode_bypass(tmp_path: Path) -> None:
    """Raw mode must not touch newlines regardless of the source."""
    config = _config(tmp_path)
    fragment = "<w:p><w:r><w:t>a\nb</w:t></w:r></w:p>\n\n<w:p/>"
    items = (ContentItem(content_literal=fragment),)
    result = resolve_items(items, raw=True, config=config)
    assert result[0].paragraphs == (fragment,)


# ---------------------------------------------------------------------------
# 5.6 Normalize
# ---------------------------------------------------------------------------


def test_normalize_task_example() -> None:
    assert normalize("三十日,共60天") == "三十日，共 60 天"


def test_normalize_default_leaves_pure_latin_alone() -> None:
    assert normalize("hello, world.") == "hello, world."


def test_normalize_preserves_url() -> None:
    text = "请看 https://example.com/a?x=1&y=2 网页"
    assert normalize(text) == "请看 https://example.com/a?x=1&y=2 网页"


def test_normalize_preserves_email() -> None:
    text = "联系 admin@example.com 邮箱"
    assert normalize(text) == "联系 admin@example.com 邮箱"


def test_normalize_preserves_code_span() -> None:
    text = "调用 `run(a, b)` 方法"
    assert normalize(text) == "调用 `run(a, b)` 方法"


def test_normalize_preserves_boundary_spaces() -> None:
    text = "  中文,共60天  "
    result = normalize(text)
    assert result == "  中文，共 60 天  "
    assert result.startswith("  ")
    assert result.endswith("  ")


def test_normalize_period_between_cjk() -> None:
    assert normalize("好的.再见") == "好的。再见"


def test_normalize_period_in_decimal_untouched() -> None:
    assert normalize("pi = 3.14 大约") == "pi = 3.14 大约"


def test_normalize_period_between_latin_untouched() -> None:
    assert normalize("Mr. Smith 到访") == "Mr. Smith 到访"


def test_normalize_cjk_to_latin_spacing_collapses_multiple() -> None:
    assert normalize("你好   ABC 世界") == "你好 ABC 世界"


def test_normalize_empty_string() -> None:
    assert normalize("") == ""


def test_normalize_parentheses_and_semicolons() -> None:
    assert normalize("附件(见后);共2页") == "附件（见后）；共 2 页"

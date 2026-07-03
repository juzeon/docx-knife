# Constitutions

## Extra Requirements

These instructions take priority over previous instructions:

- If anything is missing / not installed or prerequisites are not satisfied, you MUST pause and strongly request the user to install or set up whatever is missing.
- It's always better to skip a test than to add a low-value one.
- Write scripts to handle/migrate big text (e.g. from json to sql, from md to docx, processing text in files, etc) instead of input and output them on your own.

## Coding Quality

### Writing Code

Keep public APIs minimal and elegant.
Structure code around durable boundaries, not short-term convenience. Keep every file reasonably sized, and break it down when it gets large.
Prefer less code when clarity is preserved. Avoid duplicate logic by relying on types, validated interfaces, and existing guarantees.
Avoid over-defensive code. Pin down external guarantees instead of speculating about them: check official documentation, validate inputs once at the boundary (e.g., `zod`), verify real shapes empirically (e.g., `curl` the API), then trust those guarantees downstream.
Let errors surface: fail fast and propagate with context. No silent fallbacks or catch-and-continue without user approval.
If translating an idea from another language, rewrite it in the idioms of the target language instead of transliterating the source pattern.
When using a library, prefer the latest idiomatic APIs.

### Choosing Dependencies

Prefer mature dependencies over bespoke code when they simplify the design. Remove or replace dependencies that constrain the design.

### Changing Existing Code

If an abstraction is wrong, refactor or rewrite it instead of layering fixes on top. Large-scale rewrites and breaking changes are encouraged when they are the right fix. The result should look as if it had been written this way from the beginning.
When behavior or a public API changes, update related tests and docs (including openspec specs) in the same change. However, only add an inline comment when code is non-obvious, and remove comments that no longer add value. Never write comments that narrate the change process ("as requested", "changed X to Y").

### Verifying

Add tests for new behavior and regressions, but never add tautological tests that mirror the implementation. Only test code that has meaningful logic (branching, transformations, error handling). Don't test code that can only break if the language, runtime, or a dependency breaks.
When a test fails, fix the cause. Never weaken assertions, special-case the test's inputs in the implementation, or delete or skip failing tests without user approval.
Do not add environment-specific workarounds without user approval. Keep the implementation direct and clean.

### Committing

Commit frequently and autonomously instead of batching large changes. The user is responsible for pushing.
Follow the project's existing commit message convention. If none, use `<type>(<scope>): <description>`.
Before committing, the checks under Verifying must pass.

# docx-knife

## 背景与问题陈述

AI agent 在编辑 docx 时反复出现的问题，按根因归类：

| 问题现象 | 根因 |
|---|---|
| 大段文本被"背错"（幻觉） | 让 LLM 整体重写/复现大段原文，自回归生成天然会漂移 |
| 插入/替换位置错误 | 用内容字符串做定位 key，本身就要求 LLM 精确复现原文才能匹配成功 |
| search-replace 的 old_str 越大越容易失败 | 定位（locate）和内容复现（reproduce）没有解耦，old_str 既要"够长以保证唯一匹配"又要"够短以保证不出错"，这两个要求互相冲突 |
| 中英文标点混乱（全角半角、顿号、引号风格） | 标点宽度/类型是给定语境下的确定性函数，却被当成生成任务交给 LLM |
| 上下文被撑爆 | 把整份文档全文喂给 LLM，而不是按需加载 |
| 文档格式/结构损坏（Word提示"需要修复"） | 写回前没有做 OOXML 结构校验；段内 run 被简单粗暴地合并/破坏 |

## 实施

- 设计文档 spec/design.md
- 实现时可以参考 /Users/anon/d/dev/docx-editor/ 的代码
- skill本体：skills/docx-knife/SKILL.md
- 代码本体：docx_knife/
- 测试：tests/
  - 测试用例：test_data/

## 测试架构

### Rules for mocks

*   **Strict Mocking Boundaries**: Only mock external I/O dependencies (e.g., Repositories, RPC clients, MQ) and non-deterministic components (e.g., Clocks, Random). **NEVER mock** POJOs (Entities, DTOs), standard collections, or pure utility classes; instantiate real objects for them instead.
*   **Skip Pass-Through Methods**: Do not generate unit tests for pure "pass-through" methods (e.g., a service method that simply calls a repository and returns) that contain no business logic, `if-else` branches, or data transformation. 
*   **Avoid Tautological Assertions**: Ensure assertions validate the *business logic calculation or transformation*. Do not write useless tests that merely assert the target method returns the exact raw data you just stubbed in the mock.
*   **Verify Side Effects**: For `void` methods or operations that mutate state, use `Mockito.verify()` to assert *behavior*. Check that the downstream dependencies are invoked the expected number of times with the exact expected arguments.
*   **Force Unhappy Path Testing**: Always generate separate test cases for failure scenarios and edge cases. Use `when(...).thenThrow(...)` to simulate database timeouts or bad RPC responses, and assert that the target method handles errors or rollbacks correctly.


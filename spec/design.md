# DOCX Patch Engine 设计

## 1. 核心行为

系统以段落 ID 定位，以原子批次提交编辑，并在保存前备份旧文件。LLM 只生成结构化操作，不接触 XML、XPath、数组下标或 `w14:paraId`。

以下输入把两段内容插入 A 与 D 之间：

```json
{
  "operations": [
    {
      "op_id": "op_001",
      "op": "insert_para_after",
      "target_id": "p_000001",
      "items": [
        {"content_literal": "B"},
        {"content_literal": "C"}
      ]
    }
  ]
}
```

```text
编辑前：A, D
编辑后：A, B, C, D
```

执行器按 `items` 顺序，从原始 A 节点开始移动插入游标。批次中的任一操作失败时，DOM、锚点表和变更日志全部回滚。

Python API 同时提供可串联的段落对象。插入操作返回新段落，可继续作为下一次编辑的锚点：

```python
[para] = doc.insert_para_after(anchor, ["B"])
[para2] = para.insert_para_after(["C"])
[para3] = para2.insert_para_before(["B.5"])
```

段内编辑绑定段落 ID，并可跨 run 匹配：

```json
{
  "op_id": "op_002",
  "op": "replace_text",
  "target_id": "p_000042",
  "find": "违约责任",
  "occurrence": 0,
  "content_literal": "赔偿责任"
}
```

即使“违约”和“责任”位于不同 `<w:r>/<w:t>`，操作仍能命中；未命中的 run、格式和非文本节点保持不变。

## 2. 处理边界

引擎读取 DOCX ZIP 包中的 `word/document.xml`，处理正文及表格单元格内的 `<w:p>`，支持：

- 段落插入、删除、替换和一对多替换；
- 段内查找、替换、删除、前插和后插；
- 跨 `<w:t>`、`<w:r>` 的连续可见文本映射；
- 基于文本 diff 的段落改写；
- 原子批处理、失败回滚和保存前 `.bak` 备份；
- workspace 恢复、源文件漂移检查和原子保存。

以下结构不进入编辑面：SDT、修订标记生成、页眉页脚、脚注尾注、批注、文本框、图片、字段刷新。编辑范围穿越不支持结构时返回 `UnsupportedStructureError`，不降级为破坏性重建。

引擎仅替换 `word/document.xml`。其他 ZIP entry 原样复制，不新增 part，不修改 content types 或 relationships。

## 3. 公共 Python API

```python
with Document.open("contract.docx", force_recreate=False) as doc:
    snapshot = doc.list_paragraphs(start=1, limit=50)
    paragraph = snapshot.paragraphs[0]

    matches = doc.count_matches("三十日", paragraph_id=paragraph.id)
    if matches != 1:
        raise ValueError(f"expected one match, got {matches}")

    result = doc.batch_edit(
        operations=[
            EditOperation.replace_text(
                paragraph_id=paragraph.id,
                find="三十日",
                replacement="六十日",
                occurrence=0,
            )
        ],
    )
    doc.save("edited.docx", validate=True)
```

公共类型包括：

- `Document`、`DocumentSnapshot`、`ParagraphInfo`；
- `ParagraphLocation`、`TableContext`；
- `TextMap`、`TextPosition`、`TextMatch`；
- `EditOperation`、`EditResult`、`OperationResult`；
- 第 12 节定义的结构化异常。

### 3.1 查询契约

```text
paragraph_count() -> int
list_paragraphs(start=1, limit=None, max_chars=80) -> DocumentSnapshot
get_visible_text() -> str
find_text(text, occurrence=None, paragraph_id=None) -> TextMatch | list[TextMatch] | None
count_matches(text, paragraph_id=None) -> int
```

`DocumentSnapshot` 返回段落列表。`ParagraphInfo` 返回 ID、全局序号、完整文本、截断预览、样式及结构位置。分页只限制结果窗口，不改变 ID 或全局序号。

`TextMatch` 返回段落 ID、字符范围、XML 节点范围、是否跨节点及总匹配数。查询文本只用于发现目标；写操作仍须携带 `target_id`。

### 3.2 写入契约

所有写操作使用同一批次 envelope：

```json
{
  "operations": []
}
```

成功提交返回逐项结果：

```json
{
  "results": [
    {"op_id": "op_002", "target_id": "p_000042", "status": "success"}
  ]
}
```

批次只操作当前 `Document` 实例中的 DOM，不接受跨进程复用的操作。`save()` 每次覆盖已有目标文件前先生成同路径 `.bak` 备份。

## 4. 操作模型

### 4.1 段落操作

段落操作名包含 `para`，与段内文本操作明确区分。插入和替换统一接收 `items` 数组；单段内容也使用单元素数组。

`replace_para` 用一组新段落替换目标段落。目标 ID 被删除，每个结果段落获得新 ID：

```json
{
  "op_id": "op_010",
  "op": "replace_para",
  "target_id": "p_000042",
  "items": [
    {"content_literal": "新的段落文本"}
  ]
}
```

`delete_para` 删除目标节点及其 manifest 项，后续访问该 ID 返回 `ParagraphNotFoundError`：

```json
{"op_id": "op_011", "op": "delete_para", "target_id": "p_000043"}
```

`insert_para_before` 与 `insert_para_after` 按数组顺序插入：

```json
{
  "op_id": "op_012",
  "op": "insert_para_before",
  "target_id": "p_000042",
  "items": [
    {"content_literal": "第一条"},
    {"content_literal": "第二条"}
  ]
}
```

对象 API 与批次操作一一对应，并返回按文档顺序排列的 `list[Paragraph]`：

```python
new_paras = doc.insert_para_before(target, ["第一条", "第二条"])
more_paras = target.insert_para_after(["第三条"])
replacement = doc.replace_para(other_target, ["替换后的第一段", "替换后的第二段"])
doc.delete_para(obsolete_target)
```

### 4.2 段内操作

段内操作为：

```text
replace_text(find, replacement, occurrence=None)
delete_text(find, occurrence=None)
insert_text_before(find, text, occurrence=None)
insert_text_after(find, text, occurrence=None)
```

`find` 支持完整文本和首尾上下文两种选择器。字符串是 `{"exact": "..."}` 的简写：

```json
{"exact": "三十日内"}
```

```json
{"prefix": "如乙方未能", "suffix": "承担违约责任"}
```

首尾上下文选择器命中从 `prefix` 开始、到随后第一个 `suffix` 结束的连续范围，适合目标文本过长、只需提供前后若干字的场景。每个候选范围仍须经过唯一性检查。

执行规则：

1. 通过 `target_id` 取得段落，再用 `TextMap` 计算全部匹配；
2. 省略 `occurrence` 时必须恰好命中一次，否则返回 `AmbiguousTextMatchError`；
3. `occurrence >= 0` 表示按从 0 开始的序号操作一个匹配，越界时返回 `TextNotFoundError`；
4. `occurrence = -1` 表示操作全部匹配；影响字符位置的操作按从后向前顺序执行；
5. 每次操作后重建该段落的 `TextMap`，下一操作基于最新文本执行；
6. 新文本继承插入点或首个被替换字符所在 run 的 `<w:rPr>`；
7. 空 `<w:t>` 可以清理；仍含 drawing、tab、break 等内容的 run 不得删除；
8. 匹配范围穿越受保护结构时，整项操作失败。

执行器绝不自动选择相似文本。调用方需要确定匹配数量时可使用 `count_matches()`。

### 4.3 内容来源

`content_literal` 适合短文本。大量文本以及金额、日期、公司名、统计值等确定性内容优先使用 `content_ref`，避免 LLM 复制长文本或重写事实数据。

```json
{
  "content_ref": {
    "type": "jsonpath",
    "source": "contract.json",
    "path": "$.party_a.name"
  }
}
```

文本文件直接读入：

```json
{"content_ref": {"type": "file", "path": "clauses/confidentiality.txt", "encoding": "utf-8"}}
```

脚本标准输出作为内容：

```json
{
  "content_ref": {
    "type": "command",
    "argv": ["python", "scripts/render_clause.py", "--contract", "contract.json"],
    "timeout_seconds": 30
  }
}
```

`command` 不经过 shell，必须使用参数数组；工作目录限制在 workspace，超时、非零退出码或非 UTF-8 输出均拒绝执行。文件路径必须位于允许的输入根目录内。

模板内容也通过引用加载：

```json
{"content_ref": {"type": "template", "name": "confidentiality_clause_v1"}}
```

执行器先解析引用，再把结果作为普通文本写入。缺失键、模板或文件不存在、命令执行失败，以及要求单值却命中多值时拒绝执行。

## 5. 锚点

### 5.1 Agent-owned ID

解析 `word/document.xml` 时，引擎按文档顺序为每个可编辑 `<w:p>` 分配 ID：

```json
{
  "id": "p_000042",
  "part": "word/document.xml",
  "node_type": "w:p",
  "original_index": 42,
  "style_id": "Heading2",
  "preview": "如乙方未能按期交付……"
}
```

ID 是当前 `Document` 实例内的执行坐标。`w14:paraId` 可能缺失或重复，只作为诊断 metadata；index 只用于展示、日志和原始顺序记录。

manifest 在内存中绑定 ID 与 lxml 节点：

```text
target_id -> anchor_manifest -> node_ref
```

节点已脱离当前 XML tree 或 ID 已删除时拒绝执行。引擎不模糊重定位。`node_ref` 不可序列化；进程恢复后重新解析 XML 并生成新 ID，旧操作全部失效。

### 5.2 ID 生命周期

- 初次解析按文档顺序分配 `p_000001` 等 ID；
- 段内编辑不改变 ID 或节点引用；
- `replace_para` 删除原 ID，并为结果段落分配新 ID；
- 删除同步移除 manifest 项；
- 插入使用 workspace 内单调递增序列，已删除 ID 不复用。

## 6. TextMap

段落的可见文本可能分散在多个 run：

```xml
<w:p>
  <w:r><w:t>违约</w:t></w:r>
  <w:r><w:rPr><w:b/></w:rPr><w:t>责任</w:t></w:r>
</w:p>
```

`TextMap.text` 为连续字符串“违约责任”，每个字符反向映射到文本节点、节点偏移、全局文本偏移和所属 run：

```python
@dataclass(frozen=True)
class TextPosition:
    node_ref: Element
    node_offset: int
    text_offset: int
    run_ref: Element | None

@dataclass(frozen=True)
class TextMap:
    text: str
    positions: tuple[TextPosition, ...]
```

抽取按文档顺序拼接 `<w:t>`，并把 tab、换行等可见节点映射为对应字符。既有修订内容默认包含 `<w:ins>`、排除 `<w:del>`；策略必须可配置。字段、超链接、书签等结构记录在命中 metadata 中，供能力矩阵决定允许、保留或拒绝。

## 7. XML 修改策略

### 7.1 插入顺序

同一锚点的后插不能反复调用 `anchor.addnext()`，否则 B、C 会倒序。正确实现使用移动游标：

```python
cursor = anchor
for item in items:
    paragraph = build_paragraph(item)
    cursor.addnext(paragraph)
    cursor = paragraph
```

前插可以按数组倒序调用 `anchor.addprevious()`，也可以取得一次父节点位置后正序批量插入。临时位置仅服务单次 DOM 操作，不成为公共执行坐标。

### 7.2 新段落格式

插入或替换段落时：

- 复制锚点的 `<w:pPr>`；
- 优先复制锚点首个普通文本 run 的 `<w:rPr>`；
- 不存在普通文本 run 时创建无 `<w:rPr>` 的 run；
- 文本首尾含空格时设置 `xml:space="preserve"`。

整段替换可能丢失段内多 run 样式、局部加粗、超链接、字段、书签、批注范围和修订标记。目标段落包含这些结构时仍执行替换，并在 `OperationResult.warnings` 和 change log 中列出被移除的结构。段内操作未穿越这些结构时保留未命中内容。

### 7.3 文本规范化

`normalize_text=false` 是默认行为：不改标点、空格、引号或中英文混排，只执行 OOXML 必需的转义和 `xml:space` 处理。

`normalize_text=true` 启用基础中文标点与中英文空格规范化，同时跳过 URL、邮箱和代码片段。规范化不能删除文本首尾空格。

## 8. 表格段落

表格单元格内的 `<w:p>` 使用相同 ID 和操作 API。差异只体现在位置 metadata：

```json
{
  "table_context": {
    "table_index": 2,
    "row_index": 4,
    "physical_cell_index": 1,
    "logical_column_index": 2,
    "grid_span": 1,
    "nesting_depth": 0,
    "paragraph_index_in_cell": 0
  }
}
```

`logical_column_index` 计入 `w:gridSpan` 和 `w:gridBefore`。嵌套表格记录全局 table index 与 `nesting_depth`。垂直合并等无法可靠还原为视觉坐标的结构返回明确 metadata，不猜测布局。行列信息只用于展示和审计，不参与定位。

## 9. 规范化与原子批次

每条段落插入或替换操作直接携带有序 `items` 数组。数组位置定义结果顺序，不再引入额外排序字段。

冲突矩阵：

| 同一目标的操作 | 结果 |
| --- | --- |
| 多个 `insert_para_after` | 按操作顺序合并 `items` 数组后执行 |
| 多个 `insert_para_before` | 按操作顺序合并 `items` 数组后执行 |
| `items` 为空 | 拒绝 |
| 多个 `replace_para` | 拒绝 |
| `replace_para` + `delete_para` | 拒绝 |
| `delete_para` + `insert_para_after` | 拒绝 |
| `delete_para` + `insert_para_before` | 允许，先插入再删除 |
| `replace_para` + 前插或后插 | 允许，先插入再替换 |

批次按以下顺序执行：

1. 规范化全部操作；
2. 校验 schema、target、occurrence、内容引用和冲突；
3. 保存 DOM、manifest 和 change log 状态快照；
4. 按确定顺序在内存中应用全部操作；
5. 执行提交前校验；
6. 全部成功后提交；任一步失败则恢复完整快照。

预校验覆盖全部操作。回滚后文档状态必须与批次开始前一致；失败批次只记录一条不改变文档的审计事件。

## 10. Workspace 与保存

打开文档时解包到独立 workspace，并记录源文件指纹：

```json
{
  "source_path": "/path/to/source.docx",
  "source_size": 123456,
  "source_mtime_ns": 1710000000000000000,
  "source_sha256": "...",
  "created_at": "...",
  "backup_path": "/path/to/output.docx.bak"
}
```

保存流程：

1. `sync_check()` 比较源文件 size、mtime 和 SHA-256；任一变化都抛出 `WorkspaceSyncError`；
2. 目标文件已存在时，先把旧目标原子复制为 `<target>.bak`；已有 `.bak` 被本次备份替换；
3. 序列化 `word/document.xml`，不 pretty-print；
4. 将原 ZIP entry 复制到临时输出，仅替换主文档 XML；
5. 校验临时 DOCX；
6. 原子 rename 到目标路径；保存失败时保留原目标和 `.bak`。

正常关闭清理 workspace；异常退出保留现场。恢复 workspace 时重新建立锚点。`force_recreate` 显式丢弃旧 workspace，并从当前源文件重建。

## 11. 校验

提交前必须确认：

- `word/document.xml` 可重新解析且 XML 良构；
- 全部操作已消费，目标状态符合预期；
- 段落数量变化与操作一致；
- 未命中节点的 XML 内容不变；
- 每项操作的结果和警告均已记录。

`save(validate=true)` 使用 LibreOffice headless 打开并转换输出文件。环境缺少 LibreOffice 时返回 `ValidatorUnavailableError`，不能把未执行验证记为通过。

## 12. 错误模型

所有公共异常继承 `DocxKnifeError`，包含适合程序处理的可序列化字段：

```text
DocumentNotFoundError(path)
InvalidDocumentError(path, reason)
WorkspaceExistsError(workspace_path)
WorkspaceSyncError(source_path, expected_fingerprint, actual_fingerprint)
ParagraphNotFoundError(target_id)
TextNotFoundError(target_id, selector, occurrence, total_matches)
AmbiguousTextMatchError(target_id, selector, total_matches)
UnsupportedStructureError(target_id, structures, matched_range)
BatchOperationError(operation_index, op_id, reason, cause, rolled_back=true)
ValidationError(stage, checks, failed_check)
ValidatorUnavailableError(validator)
```

错误消息供人阅读，字段供 CLI、Agent 和测试稳定分支。文本预览必须截断。`BatchOperationError` 保留原始 cause。候选段落或文本只能作为只读诊断返回，不能触发自动重试写入。

## 13. Change Log

每次成功操作记录目标、警告及截断后的前后预览：

```json
{
  "op_id": "op_001",
  "op": "insert_para_after",
  "target_id": "p_000042",
  "status": "success",
  "warnings": [],
  "before": {"preview": "如乙方未能按期交付……"},
  "after": {
    "inserted_count": 2,
    "previews": ["第一条……", "第二条……"]
  }
}
```

失败批次记录错误类型、预期与实际状态以及 `rolled_back=true`。日志用于审阅、回归测试、诊断和证明未命中内容未被修改，不保存无边界的文档全文。

## 14. LLM 边界

LLM 输入仅包含用户要求、相关段落、可用内容引用和操作 schema。输出必须引用读取结果中的 `target_id`。长文本和确定性数据优先输出 `content_ref`，不复制到 `content_literal`。

LLM 不得输出 XML、XPath、index、`w14:paraId`、XML 字符偏移、不带目标 ID 的全局旧文本，或读取结果中不存在的目标 ID。执行器负责消歧、TextMap 映射、结构保护、规范化和回滚。

## 15. 验收测试

### 15.1 定位与顺序

- 重复或缺失 `w14:paraId` 的段落仍获得唯一 ID；
- A 后按数组顺序插入 B、C，结果严格为 A、B、C；
- D 前按数组顺序插入 B、C，结果严格为 B、C、D；
- 删除 A 后再要求在 A 后插入时，批次预校验失败。

### 15.2 段内编辑

- “违约”与“责任”分属两个 run 时，仍能替换“违约责任”；
- 同段有两个“30 日”且省略 `occurrence` 时返回 `AmbiguousTextMatchError`；`occurrence=1` 只修改第二处，`occurrence=-1` 修改全部；
- 首尾上下文选择器可跨 run 唯一定位长文本范围；
- 随机拆分 run 后，TextMap 可见文本和编辑结果保持一致。

### 15.3 一致性与保存

- 执行中失败后，DOM、manifest 和日志恢复到批次前状态；
- 外部修改源 DOCX 后，保存原路径触发 `WorkspaceSyncError` 且不覆盖外部版本；
- 覆盖已有目标文件前生成 `<target>.bak`，连续保存只保留上一次目标版本；
- 输出 DOCX 可重新打开；启用验证时可由 LibreOffice 转换；
- 未命中 XML 节点逐字节不变，ZIP 中其他 entry 内容不变。

### 15.4 表格位置

- 前一单元格 `gridSpan=2` 时，后一单元格的逻辑列跨过两列；
- 嵌套表格返回正确 `table_index` 和 `nesting_depth`；
- 表格位置 metadata 不参与操作定位。

## 16. 工程质量门槛

- 单元测试覆盖 TextMap、跨 run 操作、批次回滚、备份保存、表格位置及全部异常；
- 集成测试使用真实 DOCX 验证打开、编辑、保存、重新打开及可选 LibreOffice 校验；
- 属性测试覆盖随机 run 切分、未命中节点不变和失败批次状态不变；
- 测试默认超时 60 秒，LibreOffice 测试单独标记，缺少依赖时明确 skip；
- CI 执行格式化、lint、类型检查、单元测试、覆盖率和 Python 版本矩阵；
- 性能基准覆盖大文档分页、表格位置索引、TextMap 构建和批量编辑，禁止逐段重复扫描整棵 DOM。

交付物包括可安装 Python package、API reference、quickstart 和只调用公共 API 的 Agent Skill。核心库不依赖 UI，也不代理第三方编辑器能力。

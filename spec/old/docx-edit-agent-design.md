# DOCX 编辑 Agent Skill · 设计文档

## 0. 文档目的

本设计文档整合了关于"AI agent 编辑 docx 时出现幻觉、定位错误"问题的完整讨论，作为构建一个带 scripts 的 skill 的实施依据。目标读者是实现这个 skill 的工程师（可能是未来的自己）。

---

## 1. 背景与问题陈述

AI agent 在编辑 docx 时反复出现的问题，按根因归类：

| 问题现象 | 根因 |
|---|---|
| 大段文本被"背错"（幻觉） | 让 LLM 整体重写/复现大段原文，自回归生成天然会漂移 |
| 插入/替换位置错误 | 用内容字符串做定位 key，本身就要求 LLM 精确复现原文才能匹配成功 |
| search-replace 的 old_str 越大越容易失败 | 定位（locate）和内容复现（reproduce）没有解耦，old_str 既要"够长以保证唯一匹配"又要"够短以保证不出错"，这两个要求互相冲突 |
| 数字、日期等派生内容被写错 | 用 LLM 自由生成去"转写"本该精确的数据 |
| 中英文标点混乱（全角半角、顿号、引号风格） | 标点宽度/类型是给定语境下的确定性函数，却被当成生成任务交给 LLM |
| 上下文被撑爆 | 把整份文档全文喂给 LLM，而不是按需加载 |
| 文档格式/结构损坏（Word提示"需要修复"） | 写回前没有做 OOXML 结构校验；段内 run 被简单粗暴地合并/破坏 |
| 长文档、多轮编辑后错误复合放大 | 缺少分批处理和阶段性校验，一次性把大任务丢给模型 |

## 2. 设计目标 / 非目标

**目标**
- 让 LLM 只做"决策"（改哪里、改成什么、以及是否需要它自己写内容），把"执行"完全交给确定性代码。
- 把"生成保真度"问题尽量转化为"路由/映射正确性"问题——后者可穷举校验，前者不行。
- 支持大文档场景下的可控上下文占用。
- 编辑结果可审阅、可回滚、可审计溯源。

**非目标（本版本不解决）**
- 不追求消灭 LLM 生成侧的幻觉（`content_literal` 这部分的自由创作内容仍然依赖模型本身能力）。
- 不实现专用的小模型 "fast-apply" 基础设施（见 §10 未来扩展）。
- 不处理 .doc（老格式）、加密文档、宏文档的特殊解析（转换为前置步骤，用现有 skill 的 soffice 转换即可）。

## 3. 核心设计原则

以下原则是整份文档中所有模块设计的依据，先列出来方便实现时对照检查：

1. **最小化编辑，禁止整体重写**：任何时候都不让 LLM 输出整份文档或整段未改动内容。
2. **ID 寻址，定位与内容复现解耦**：用提前分配好的短 ID（复用 `w14:paraId`）做 key，LLM 不需要复现原文来定位。
3. **分层加载，按需检索**：先给"大纲索引"（ID+预览），LLM 圈定目标后才展开该节点全文；定位阶段优先用确定性检索（grep/向量），不占用 LLM 上下文。
4. **内容分类，能不生成就不生成**：把要写入的内容分成"照抄原文 / 派生计算值 / 外部数据 / 模板样板 / 真正原创文字"五类，只有最后一类必须走 LLM 自由生成（`content_literal`），其余走引用解析（`content_ref`）。
5. **确定性后处理层**：XML 转义、`xml:space="preserve"`、标点全半角归一化、数字/日期格式化——这些有确定正确答案的转换，一律不进 LLM 生成通道，统一在 apply 阶段强制执行。
6. **强 schema 约束**：LLM 输出必须是可校验的结构化 op，格式不对直接打回重试，不做"尽力解析"。
7. **写回前结构校验 + 写回后渲染回看**：保存前检查 ID 唯一性、引用完整性、XML 良构性；保存后转 PDF/图片做视觉核验。
8. **修订追踪兜底**：默认以 `<w:ins>/<w:del>` 落地编辑，把 AI 输出当"建议"而非"终稿"，人工在 Word 审阅视图里逐条确认。
9. **分批处理，可断点续传**：大文档按章节切分成独立事务，每批处理后落检查点，允许中途恢复。

## 4. 端到端流程

```
┌─────────────┐   ┌──────────────┐   ┌───────────┐   ┌────────────┐
│ unpack docx  │→ │ 分配/复用ID   │→ │ 构建大纲    │→ │ (可选)定位   │
│              │  │ w14:paraId   │  │ outline.json│  │ 检索候选ID  │
└─────────────┘   └──────────────┘   └───────────┘   └─────┬──────┘
                                                             ↓
┌─────────────┐   ┌──────────────┐   ┌───────────┐   ┌────────────┐
│  pack docx   │← │ 结构校验       │← │ apply_ops  │← │ LLM 生成    │
│  (不reformat) │  │ (重复ID/悬空   │  │ (转义/标点  │  │ 结构化op    │
│              │  │  引用/良构性)  │  │  /preserve)│  │ 列表(带schema)│
└──────┬──────┘   └──────────────┘   └───────────┘   └────────────┘
       ↓
┌─────────────┐
│ 渲染回看验证  │  → pdf/图片，人工或模型抽查
│ + track      │
│ changes审阅  │
└─────────────┘
```

每个箭头两端都是独立可测试的模块，中间没有"LLM 直接操作 XML"这一步。

## 5. 目录结构

按 Anthropic skill 规范组织（`SKILL.md` + `scripts/` + `references/`）：

```
docx-edit-agent/
├── SKILL.md                      # 触发条件 + 流程总览，指向下面各脚本
├── scripts/
│   ├── parse_and_index.py        # unpack + 分配ID + 生成大纲
│   ├── locate.py                 # 可选：关键词/向量检索候选段落
│   ├── op_schema.py              # op 的 pydantic 定义 + 校验
│   ├── resolve_refs.py           # content_ref 解析器
│   ├── normalize_text.py         # 标点归一化 + 转义 + preserve处理
│   ├── apply_ops.py              # 把 op 应用到 XML 树
│   ├── validate_structure.py     # 写回前结构校验
│   ├── track_changes.py          # 包装 w:ins/w:del + 生成change log
│   ├── pack.py                   # 打包，去symlink，不reformat
│   └── render_verify.py          # 转pdf/图片
├── references/
│   ├── ooxml_notes.md            # run切分、paraId规范、xml:space等踩坑记录
│   ├── op_schema.md              # op JSON 结构说明（给LLM读的简化版）
│   └── punctuation_rules.md      # 标点归一化规则表
└── assets/
    └── (可选：常用模板条款 template library)
```

## 6. 关键数据结构

### 6.1 大纲索引条目（outline.json）

```json
{
  "id": "60C4CAE9",
  "kind": "paragraph",
  "style": "Heading2",
  "section_path": ["第三章", "3.2 违约责任"],
  "preview_head": "如乙方未能按期交付……",
  "preview_tail": "……并承担相应赔偿责任。",
  "char_count": 214,
  "run_count": 3
}
```
只在 LLM 需要精读某段时，才用 `id` 单独取全文（含 run 结构），大纲阶段只给预览。

### 6.2 编辑指令（op）

```json
{"op": "replace", "target_id": "60C4CAE9", "content_literal": "……仅限真正需要LLM原创的文字"}
{"op": "replace", "target_id": "71A2B3C4", "content_ref": {"type": "jsonpath", "source": "q3_report.json", "path": "$.summary.total_revenue", "format": "currency_cny"}}
{"op": "replace", "target_id": "82D5E6F7", "content_ref": {"type": "script_output", "key": "computed_penalty_amount"}}
{"op": "replace", "target_id": "93G8H9I0", "content_ref": {"type": "template", "name": "confidentiality_clause_v3"}}
{"op": "insert_after", "target_id": "60C4CAE9", "content_literal": "……"}
{"op": "delete", "target_id": "60C4CAE9"}
```

`content_literal` 和 `content_ref` 二选一，schema 层面互斥校验。`target_id` 必须在当前会话的 ID 表里存在，否则整条 op 拒绝执行（不静默跳过，要显式报错让上层决定重试还是人工介入）。

### 6.3 可用引用目录（available_refs.json，喂给 LLM 用）

```json
{
  "available_refs": [
    {"key": "q3_report.json:$.summary.total_revenue", "type": "number", "desc": "Q3总营收"},
    {"key": "script_output:computed_penalty_amount", "type": "number", "desc": "违约金计算结果"},
    {"key": "template:confidentiality_clause_v3", "type": "boilerplate", "desc": "保密条款标准版"}
  ]
}
```
只列 key + 类型 + 一句话描述，不塞真实值，保持轻量。

## 7. 模块详细设计

### 7.1 `parse_and_index.py`

- unzip → 用 lxml 解析 `word/document.xml`（不要用正则扫全文本）。
- 遍历所有 `<w:p>`/`<w:tr>`：
  - 已有 `w14:paraId` 直接复用；缺失的按 `ST_LongHexNumber` 规范（唯一、`0 < value < 0x80000000`）生成并回填。
  - **注意**：paraId 只在本次会话内保证稳定——用户后续用 Word 手动编辑并保存，Word 可能重新生成所有 paraId。不要把它当跨会话的持久主键存到外部数据库。
- 非段落元素直接复用原生 ID：图片/关系用 `r:id`，脚注用脚注 ID，批注用 comment ID，不重新发明一套。
- 表格：`row_id`（表格行的paraId）+ `col_index` 组成寻址 key。
- 输出 `outline.json`（分层：先章节标题，段落列表可选择性展开，避免大文档下大纲本身也超预算）。

### 7.2 `locate.py`（可选，大文档/模糊指令时启用）

- 输入：自然语言描述（"关于违约责任的条款"）。
- 用关键词匹配 / 简单 embedding 检索大纲预览，返回 top-K 候选 `id`。
- 这一步不经过 LLM，纯代码检索，只有当候选数量仍然过多需要语义判断时才把候选全文交给 LLM 做最终确认。

### 7.3 `op_schema.py`

- 用 Pydantic 定义 op 的严格 schema，`content_literal` 与 `content_ref` 用 discriminated union 互斥。
- 校验失败：返回明确的错误原因（哪个字段、为什么），供上层决定是重新请求 LLM 生成还是直接报错终止，不做"尽量兼容解析"。
- 单条 op 的 `content_literal` 长度设一个软上限（比如 500 字），超限的直接拒绝并提示"请拆分成多条 op"——从源头限制单次生成的风险敞口。

### 7.4 `resolve_refs.py`

- 输入一条 `content_ref`，输出解析后的字符串。
- 校验：key 是否在 `available_refs` 目录里存在；解析出的数据类型是否与目标位置预期类型匹配（数字字段不能塞进模板类内容）。
- 格式化（货币、日期、中文数字大写等）在这里用确定性函数完成，绝不回退到"让 LLM 转写"。
- `type: source_span` 的情况（引用文档自身其他位置的原文）：直接从已解析的文档树里取，不经过任何生成。

### 7.5 `normalize_text.py`

对所有 `content_literal`（LLM 生成的自由文本）强制执行，`content_ref` 解析结果视情况决定是否跳过（模板/外部数据一般已经是规范格式，不需要再处理，除非明确要求统一风格）：

- **全角/半角标点**：逐字符判断左右相邻字符是否属于 CJK 区间，据此决定该标点应为全角还是半角；判断必须在**整段拼接后的全文视图**上做，不能逐 run 单独判断（否则标点卡在 run 边界时看不到相邻字符）。
- **局部惯例优先**：参考本段/本文档已有的标点风格（简体 `""` 还是繁体 `「」`）而不是套用固定默认值。
- **顿号 vs 逗号**：语义判断，不做强制自动改写，只在 lint 报告里标注"疑似列举场景建议用顿号"供人工确认。
- **中英文混排间距**（pangu spacing）：CJK 与半角字母/数字之间补空格。
- **内嵌英文/代码/URL 豁免**：给规则引擎加一层"忽略区间"标记，防止把 URL、代码片段内部的半角符号也转换掉。
- 归一化完成后，做 **XML 转义**（`&`/`<`/`>`/`"`）和 `xml:space="preserve"` 判断（新文本首尾若有空格，显式加此属性，否则会被吃掉）。

### 7.6 `apply_ops.py`

- 用 `outline.json` 里的 `id → 节点` 映射做 O(1) 定位，不重复搜索。
- 段内 run 处理策略二选一（按配置或按段落复杂度自动选择）：
  - **整段替换**：用首个 run 的 `<w:rPr>` 作为默认格式生成单一新 run，简单可靠，代价是段内局部强调格式会被拉平。
  - **Markdown 往返**（段内有多处局部格式变化时优先用这个）：先把该段落投影成 Markdown 给 LLM/规则处理，再用确定性转换器转回带 run 切分的 OOXML，保留局部格式。
- `insert_before`/`insert_after`：在目标节点前后插入新 `<w:p>`，新段落复用相邻段落的 `pPr`（避免样式突变）。
- 所有变更先在内存中的 XML 树上完成，最后统一序列化，不要每条 op 都单独 unzip/zip 一轮。

### 7.7 `validate_structure.py`

保存前跑一遍，等价于前面讨论过的 docx 结构审计清单：
- 有没有重复 paraId；
- 是否所有 op 引用的 `target_id`/`content_ref.key` 在实际应用后都已消费且无残留悬空引用；
- 书签、脚注引用、图片关系（`r:id`）是否仍能正确解析；
- XML 是否良构（能被 lxml 无错解析）。
校验不通过：拒绝打包，把失败原因和涉及的 op 一起报出来。

### 7.8 `track_changes.py`（默认开启，可关闭）

- 把 `apply_ops.py` 产生的变更包装成 `<w:ins>`/`<w:del>`（带 `w:id`/`w:author`/`w:date`），而不是直接落地成终稿。
- 生成一份 `change_log.txt`：按段落列出"插入/删除/替换"及来源（`content_literal` 还是具体的 `content_ref` key），方便人工快速审阅而不用逐字比对全文。
- 提供 `--no-track` 参数供确实需要直接落地的场景使用。

### 7.9 `pack.py`

- 复用现有 docx skill 的做法：清理 symlink（防御不可信输入），rezip，**不对 XML 做任何 reformat/pretty-print**（保持后续可能的二次编辑里字符串匹配可用）。

### 7.10 `render_verify.py`

- `soffice --headless --convert-to pdf` → `pdftoppm` 转图片 → 人工或模型抽查排版有没有明显错乱（表格塌陷、分页错位等结构校验覆盖不到的视觉问题）。

## 8. 大文档批处理与断点续传

- 按标题层级切成"处理单元"（比如按二级标题切章节），单元大小上限按 token 预算反推（大纲预览+目标段落全文，控制在单次请求可接受的范围内）。
- 维护 `progress.json`：
  ```json
  {"processed_sections": ["3.1", "3.2"], "pending_sections": ["3.3", "4.1"], "applied_ops": [...], "failed_ops": [...]}
  ```
- 每个单元走一遍完整的"定位→生成→apply→校验"小事务，成功才推进检查点，失败就停在当前单元等待重试/人工介入，不影响已完成部分。
- 这样即使中途中断（token预算耗尽、连接中断），也能从检查点续跑，而不用整份文档重来。

## 9. 风险清单（实现时逐条自检）

| 风险点 | 缓解措施 | 对应模块 |
|---|---|---|
| run 跨段落格式边界导致定位/替换错误 | 段落全文视图判断 + markdown往返保格式 | apply_ops.py |
| paraId 在用户手动编辑后被 Word 重新生成 | 仅作会话内临时寻址键，不做外部持久主键 | parse_and_index.py |
| 表格/图片/脚注等非段落元素无法用paraId寻址 | 复用各自原生ID体系 | parse_and_index.py |
| 超大文档喂爆上下文 | 分层大纲 + 按需展开 + 检索定位 | parse_and_index.py / locate.py |
| content_ref 的 key 不存在或类型不匹配 | 强校验，不匹配直接拒绝执行 | resolve_refs.py |
| 数字/日期被LLM转写出错 | 一律走确定性格式化函数 | resolve_refs.py |
| 标点全半角/顿号/引号风格混乱 | 归一化强制关卡，参考局部惯例 | normalize_text.py |
| 首尾空格被吃掉 | 显式设置 `xml:space="preserve"` | normalize_text.py / apply_ops.py |
| XML特殊字符破坏文档 | 统一转义 | normalize_text.py |
| 单条op文本过长风险敞口大 | 长度软上限，超限拒绝并要求拆分 | op_schema.py |
| 写回后Word提示"文档已损坏" | 写回前结构校验 | validate_structure.py |
| AI编辑直接覆盖终稿，人工无法把关 | 默认走track changes，人工审阅接受 | track_changes.py |
| 长任务中途失败需要从头重来 | 分批处理+检查点续传 | 整体pipeline |

## 10. 测试计划

- **单元测试**：每个脚本独立测试。重点覆盖 `normalize_text.py`（构造中英混排/顿号/引号/URL嵌入等边界样例）、`resolve_refs.py`（key缺失/类型不匹配的拒绝逻辑）、`apply_ops.py`（run级别替换的格式保真）。
- **陷阱样例集**：准备几份"陷阱"docx（多级标题+表格+脚注+批注+中英文混排+历史修订痕迹），跑全流程后人工核对渲染结果。
- **回归校验**：每次编辑前后做段落级 diff，抽样确认"未在 op 列表里的段落"字节级别未发生变化——这是检验"最小化编辑"原则是否真正落实的直接指标。
- **大文档压力测试**：构造一份超预算的长文档，验证分批+断点续传路径（人为在中途掐断，确认能从检查点正确恢复）。

## 11. SKILL.md 编写要点

按 progressive disclosure 原则，`SKILL.md` 本体只放：
- 触发条件（"编辑较大/较复杂的docx，尤其涉及数据填充、批量替换、格式敏感场景时触发"）。
- 流程总览（对应 §4 的流程图，用文字精简版）。
- 每个脚本一句话说明 + 何时调用它。
- 指向 `references/ooxml_notes.md`、`references/op_schema.md` 的引用（细节不在 SKILL.md 里展开，按需加载）。

控制在 500 行以内；细节（比如 paraId 规范的完整技术细节、标点规则表）放进 `references/`，由 Claude 按需读取。

## 12. 已知局限 & 未来扩展

- `content_literal` 部分仍然依赖 LLM 生成质量，本设计不能完全消灭这部分的幻觉风险，只是把风险敞口压缩到最小必要范围。
- 未引入专门训练的"apply模型"（类似 Morph Fast Apply 的思路），目前用确定性脚本代替；如果未来编辑吞吐量要求提高，可以考虑把 `apply_ops.py` 的合并逻辑换成专用小模型。
- 如果目标文档是自己控制生成的模板，后续可以把 ID 体系从 `w14:paraId` 升级为 Word 原生的内容控件（`w:sdt` + `w:tag`/`w:alias`），语义更明确、也不受 Word 重新保存后重新生成paraId的影响。
- 可以引入"草稿-核对"双 agent 流水线：一个 agent 生成 op 列表，另一个专门核对 `content_literal` 里出现的数字/专有名词是否都能在 `available_refs` 或原文里找到依据，作为 §7.5 之外再加一层语义级校验。

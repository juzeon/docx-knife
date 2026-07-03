## Why

现有 AI Agent 编辑 DOCX 时把定位、原文复现和内容生成耦合在一起，容易造成文本幻觉、定位失败、格式损坏与上下文膨胀。需要一个以稳定段落 ID 和结构化操作为边界的补丁引擎，让 LLM 只表达编辑意图，由确定性执行器负责 OOXML 映射、校验、回滚与安全保存。

## What Changes

- 新增可安装的 Python DOCX Patch Engine，仅编辑 `word/document.xml` 中正文与表格单元格内的可编辑段落。
- 提供段落分页、搜索、原始 XML 读取、文本匹配与计数 API，并以实例内段落 ID 作为唯一写入坐标。
- 提供段落插入、删除、一对多替换，以及可跨 run 的段内替换、删除、前插和后插操作。
- 引入 TextMap，在可见文本与 OOXML 节点之间建立可逆映射，并保护 tab、换行、分页符等原子结构。
- 提供 `content_literal`、受控文件、JSONPath 与无 shell 命令等内容来源，避免 LLM 复现长文本和确定性数据。
- 实现操作规范化、冲突检测、原子批次、完整回滚、结构化异常和有边界的变更日志。
- 实现源文件漂移检测、覆盖前 `.bak` 备份、临时 ZIP 校验与原子保存，保持其他 ZIP entry 不变。
- 提供公共 Python API、API reference、quickstart，以及只调用公共 API 且不暴露 raw XML 的 Agent Skill。

## Capabilities

### New Capabilities

- `document-query-and-anchors`: 文档打开、段落清单与搜索、稳定实例内 ID、SDT 排除规则及表格位置元数据。
- `text-map-editing`: 可见文本投影、跨 run 匹配、保留标记、段内编辑及受保护结构检测。
- `paragraph-editing`: 段落插入、批量删除、一对多替换、对象式可串联 API、raw XML 模式与格式继承。
- `content-sources-and-normalization`: 字面量与受控内容引用解析、换行展开和可选文本规范化。
- `atomic-batch-processing`: 操作 schema、冲突规则、确定性执行顺序、批次回滚、结构化错误与变更日志。
- `safe-document-save`: 源漂移检测、备份、OOXML/ZIP 校验、未命中结构保真及原子保存。

### Modified Capabilities

无。

## Impact

- 新增 Python package 及其公共类型、异常层、OOXML 解析与编辑核心。
- 新增对成熟 ZIP、XML、schema 校验、JSONPath 和属性测试依赖的评估与最小化集成。
- 新增真实 DOCX 集成测试、属性测试、性能基准和 Python 版本矩阵 CI。
- Agent 交互契约改为先查询段落 ID，再提交结构化操作；LLM 不再直接处理 XML、XPath、数组下标或 `w14:paraId`。

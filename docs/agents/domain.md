# Domain docs

Layout: single-context.

## Read before domain exploration

读取根 CONTEXT.md（如存在），以及 docs/adr/ 中与当前任务有关的 ADR。
文件尚不存在时继续工作；不为完成配置预先创建空领域文档。

实现或验收 Cash Rec 前，读取：
- Cash_Rec_精简改造_PRD.md
- Cash_Rec_安装版开发规格_PRD1.6.md

业务以已确认 PRD 和规格为依据，采用其领域词汇。
CONTEXT.md 中已有术语定义时保持一致，新增术语通过领域建模记录。
发现与 ADR 冲突的方案时，明确指出冲突及原因。

## Layout

- CONTEXT.md：项目领域词汇和模型
- docs/adr/：领域或架构决策记录

领域文档由 domain-modeling 等技能在实际澄清术语和决策时逐步创建。

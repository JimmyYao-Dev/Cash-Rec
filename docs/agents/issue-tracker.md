# Issue tracker: GitHub

Repository: JimmyYao-Dev/Cash-Rec
URL: https://github.com/JimmyYao-Dev/Cash-Rec

规格和开发任务存放于该仓库的 GitHub Issues，使用 gh CLI。
命令显式带 --repo JimmyYao-Dev/Cash-Rec，避免依赖本地远端配置。

## Operations

- 发布规格或任务：gh issue create --repo JimmyYao-Dev/Cash-Rec --title "<title>" --body-file "<UTF-8 body file>"
- 读取任务和评论：gh issue view <number> --repo JimmyYao-Dev/Cash-Rec --comments
- 列出任务：gh issue list --repo JimmyYao-Dev/Cash-Rec --state open --json number,title,body,labels
- 添加评论：gh issue comment <number> --repo JimmyYao-Dev/Cash-Rec --body-file "<UTF-8 comment file>"
- 应用或移除标签：gh issue edit <number> --repo JimmyYao-Dev/Cash-Rec --add-label "<label>"，或 --remove-label
- 关闭任务：gh issue close <number> --repo JimmyYao-Dev/Cash-Rec

标签含义读取 triage-labels.md。更新已有工作前先读取 Issue 及评论，
同一规格已有对应 Issue 时更新原 Issue，避免重复创建。

## Publication contract

“发布到 Issue Tracker”表示创建或更新 GitHub Issue，不代表推送源码、
业务数据或整个工作区。规格使用独立 Issue；后续实施任务分别建 Issue。

## Pull requests as a triage surface

PRs as a request surface: no.

## Wayfinding

使用父 Issue 和子任务表示工作地图及任务。
依赖优先使用 GitHub 原生 blocking 关系；不可用时写明 Blocked by。
认领时分配给执行者；完成后记录结果、关闭任务并更新父 Issue。

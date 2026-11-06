# First implementation batch

Parent specification: https://github.com/JimmyYao-Dev/Cash-Rec/issues/1

This batch implements Issues #2–#8: a Windows installer, prepared-file workflow,
Equity, Bank, PSP and TransactionExchange, and comparable previous balances.
Monthly reporting starts with Issue #9 and is outside this batch.

The integration branch is `feat/first-seven-tickets`. Implementers work in
isolated worktrees, and their changes merge here for installed-GUI acceptance.
The application uses a self-contained Windows x64 .NET desktop build so users
do not need a development environment. Business data stays in a separate workspace.

Business acceptance starts through the installed graphical application.
Independent workbook checks support that workflow. DEP/PNL formulas require
actual Excel editing; unexecuted checks remain explicit acceptance gaps.

Only source, build recipes and synthetic fixtures belong in this repository.
Local business inputs, recordings, report outputs and mapping backups stay local.

The parent issue remains open. Closing references are added only for batch
tickets whose implementation and required acceptance are complete.

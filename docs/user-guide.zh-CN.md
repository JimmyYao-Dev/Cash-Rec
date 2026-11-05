# Cash Rec 0.1.0 使用说明

本版处理单日 TMA/TMG，生成 Equity、Bank、PSP、TransactionExchange 四张基础表。月度 Breakdown 属于后续任务，不包含在本安装包中。

## 安装与启动

运行 `CashRec_Setup.exe`，随后从桌面或开始菜单启动 Cash Rec。安装包包含运行依赖，用户不需要安装 Python、.NET SDK 或开发工具。实际验收系统为 Windows 11 Enterprise 26100 x64；安装器最低系统检查为 Windows 10 19041 x64。尚未完成 Windows 10 或干净虚拟机兼容矩阵。

程序安装目录与业务工作区分开。升级或卸载程序不会删除独立的业务工作区。

## 准备单日文件

选择工作区和业务日期。日期目录使用 `YYYY.MM.DD`，例如 `2026.10.25`。日期目录可直接位于工作区，也可位于工作区的 `data/risk_and_cash_to_process` 下。

在日期目录的 `Needed files` 内放置以下文件，也可通过程序导入：

| 子目录 | 文件名 |
| --- | --- |
| Bank | NAB AU.csv |
| Bank | NAB GL.csv |
| MT4/MT4 | Transaction Summary Report.csv |
| MT4/Risk | Accounting Equity Report.csv |
| PSP+BANK | BANK BD.xlsx |
| PSP+BANK | PSP Balance Summary.xlsx |
| Rate | VFSC Rate.xlsx |

将完整映射文件 `account_mapping.xlsx` 放到工作区的 `config` 或 `db/tmc` 目录，也可通过界面导入。缺项草稿只供补分类，不能替代完整映射。

CSV 使用准备文件中的全部日期。PSP 只读取所选日期的 Amount。包含历史日期的汇率表只使用所选业务日。USDT 固定按 1 换算；PSP 真空白 Amount 按 0 计算。非法文本或无数值缓存的金额公式仍需修正。

## 生成与检查

1. 点击检查按钮，处理缺文件、日期不一致或不可读文件。
2. 点击生成按钮，等待任务结束。
3. 打开结果位置，查看 `CashRec_Base.xlsx` 及诊断文件。

检查文件通过仅表示输入齐备。生成成功还需要业务处理通过。缺银行分类、汇率或源金额错误时，按诊断中的文件、工作表和行号修正后重跑。

## 上期与人工 DEP

将此前日期目录与当前日期目录并列保留。程序使用最近有效且同口径的上期四表；有必要时从该日 Equity、NAB AU/GL、PSP 和汇率原文件重算余额。上期余额重算不需要该日交易文件。

没有有效上期时，程序保留当前四表并标记未完成。Previous、Movement、PNL 留空，不能作为正常零差额使用。

首次 DEP 为数值 0。在 Excel 修改 DEP 后，PNL 按 `Movement−DEP` 重算，Movement 保持当前与上期余额差。重跑保留可匹配业务键的人工 DEP；替换旧输出前会备份。无法匹配的人工值会单独保留供复核。

## 交付边界

原始文件、映射备份、诊断和输出保留在本地。用户安装包不含项目源码或开发测试工具。本版尚未实现月度汇总、跨月衔接及完整 Check 联动，后续由 Issue #9 起继续实施。

[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][string]$InstalledExe,
 [Parameter(Mandatory=$true)][string]$Workspace,
 [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
 [ValidateSet('Daily','Monthly')][string]$LockedOutput='Monthly',
 [string]$BusinessDate='2026-10-25',
 [string]$ExpectedMonthlyStage='Check'
)
$ErrorActionPreference='Stop'
$InstalledExe=(Resolve-Path -LiteralPath $InstalledExe).Path
$Workspace=(Resolve-Path -LiteralPath $Workspace).Path
$marker=Join-Path $Workspace '.cashrec-monthly-test-fixture.json'
if(-not(Test-Path -LiteralPath $marker) -or -not(Get-Content -LiteralPath $marker -Raw|ConvertFrom-Json).synthetic_only){throw 'Lock acceptance requires an explicitly marked synthetic workspace.'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$EvidenceDirectory=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
$daily=Join-Path $Workspace ($BusinessDate.Replace('-','.')+'/CashRec_Base.xlsx')
$monthly=Join-Path $Workspace ('monthly/CashRec_Breakdown_'+$BusinessDate.Substring(0,7).Replace('-','.')+'.xlsx')
$before=@{Daily=(Get-FileHash -LiteralPath $daily -Algorithm SHA256).Hash;Monthly=(Get-FileHash -LiteralPath $monthly -Algorithm SHA256).Hash}
$target=if($LockedOutput -eq 'Daily'){$daily}else{$monthly}
$result=[ordered]@{Entry='Installed real GUI with a test-owned file replacement lock';SyntheticOnly=$true;Passed=$false;LockedOutput=$LockedOutput;Before=$before;Limitation='A preflight lock rejection does not prove a post-replacement rollback was executed.'}
$handle=$null
try{
 # ShareRead permits ordinary source inspection but denies writing/deleting/replacing this owned target.
 $handle=[IO.File]::Open($target,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::Read)
 & (Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe') -NoProfile -File (Join-Path $PSScriptRoot 'Invoke-InstalledAcceptance.ps1') -InstalledExe $InstalledExe -Workspace $Workspace -BusinessDate $BusinessDate -EvidenceDirectory (Join-Path $EvidenceDirectory 'gui') -ExpectedOutcome Failure -ExpectedMonthlyStage $ExpectedMonthlyStage -ExpectedMonthlyStageOutcome Incomplete -ExpectedVersion '0.2.0' -CaseScale ('monthly-locked-'+$LockedOutput) *> (Join-Path $EvidenceDirectory 'ui-stdout.txt')
 if($LASTEXITCODE -ne 0){throw 'The installed GUI did not report the required failure for the owned output lock.'}
 $handle.Dispose();$handle=$null
 $after=@{Daily=(Get-FileHash -LiteralPath $daily -Algorithm SHA256).Hash;Monthly=(Get-FileHash -LiteralPath $monthly -Algorithm SHA256).Hash}
 $result.After=$after
 if($after.Daily -ne $before.Daily -or $after.Monthly -ne $before.Monthly){throw 'A failed GUI task modified one or both previous output workbooks.'}
 $ui=Get-Content -LiteralPath (Join-Path $EvidenceDirectory 'gui/ui-result.json') -Raw|ConvertFrom-Json
 if($ui.Run.Complete -or $ui.Run.DailyOutput -or $ui.Run.MonthlyOutput){throw 'The failed task incorrectly published output paths or complete status.'}
 if($ui.Run.Transaction.Status -notin @('NotCommitted','RolledBack')){throw 'The failed task did not provide a successful no-commit/rollback transaction status.'}
 if(@($ui.Run.Transaction.CommittedOutputs).Count -ne 0){throw 'A failed task published a committed output.'}
 $result.Transaction=$ui.Run.Transaction;$result.Passed=$true
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'locked-output-result.json'),($result|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
 $result|ConvertTo-Json -Depth 12
}catch{
 $result.Error=$_.Exception.Message
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'locked-output-result.json'),($result|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false));throw
}finally{if($null -ne $handle){$handle.Dispose()}}

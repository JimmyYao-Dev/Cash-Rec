[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][string]$InstalledExe,
 [Parameter(Mandatory=$true)][string]$Workspace,
 [Parameter(Mandatory=$true)][datetime]$BusinessDate,
 [Nullable[datetime]]$EndDate,
 [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
 [ValidateSet('Fixed','Reproduce')][string]$ExpectedDisposition='Fixed',
 [int]$TimeoutSeconds=60
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'InstalledUiHelpers.ps1')
$InstalledExe=(Resolve-Path -LiteralPath $InstalledExe).Path
$Workspace=(Resolve-Path -LiteralPath $Workspace).Path
$marker=Join-Path $Workspace '.cashrec-monthly-test-fixture.json'
if(-not(Test-Path -LiteralPath $marker) -or -not(Get-Content -LiteralPath $marker -Raw|ConvertFrom-Json).synthetic_only){throw 'Diagnostic ACL tests require a fresh explicitly marked synthetic workspace.'}
$diagnostics=Join-Path $Workspace '.cashrec/diagnostics'
[IO.Directory]::CreateDirectory($diagnostics)|Out-Null
$diagnostics=(Resolve-Path -LiteralPath $diagnostics).Path
if(-not $diagnostics.StartsWith($Workspace.TrimEnd([IO.Path]::DirectorySeparatorChar)+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Diagnostic ACL target escaped the synthetic workspace.'}
foreach($path in @($Workspace,(Split-Path -Parent $diagnostics),$diagnostics)){if((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'ACL tests cannot target a reparse-point workspace or diagnostics directory.'}}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$EvidenceDirectory=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
if($EvidenceDirectory -eq $diagnostics -or $EvidenceDirectory.StartsWith($diagnostics+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Evidence must be outside the intentionally unwritable diagnostic directory.'}
$isRange=$null -ne $EndDate
$last=if($isRange){[datetime]$EndDate}else{$BusinessDate}
if($last.Date -lt $BusinessDate.Date){throw 'This I/O test requires an ascending prepared range.'}
$expectedPaths=@()
for($date=$BusinessDate.Date;$date -le $last.Date;$date=$date.AddDays(1)){
 $expectedPaths+=Join-Path $Workspace ($date.ToString('yyyy.MM.dd')+'/CashRec_Base.xlsx')
 $expectedPaths+=Join-Path $Workspace ('monthly/CashRec_Breakdown_'+$date.ToString('yyyy.MM')+'.xlsx')
}
$expectedPaths=@($expectedPaths|Select-Object -Unique)
$originalAcl=Get-Acl -LiteralPath $diagnostics
$originalSddl=$originalAcl.GetSecurityDescriptorSddlForm([Security.AccessControl.AccessControlSections]::All)
[IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'diagnostics-original-acl.sddl.txt'),$originalSddl,[Text.UTF8Encoding]::new($false))
$result=[ordered]@{Entry='Installed real GUI Generate under a test-owned diagnostics-only CreateFiles denial';SyntheticOnly=$true;Passed=$false;ExpectedDisposition=$ExpectedDisposition;Workspace=$Workspace;DiagnosticsTarget=$diagnostics;BusinessDate=$BusinessDate.ToString('yyyy-MM-dd');EndDate=$last.ToString('yyyy-MM-dd');Range=$isRange;ProductVersion=[Diagnostics.FileVersionInfo]::GetVersionInfo($InstalledExe).ProductVersion;ExpectedReportPaths=$expectedPaths;AclRestored=$false;UsesTaskOrBatchJson=$false;UsesInternalBusinessApi=$false}
$window=$null;$process=$null;$failure=$null;$aclChanged=$false
$probe=Join-Path $diagnostics ('.owned-acl-probe-'+[guid]::NewGuid().ToString('N'))
try{
 # Establish that this fresh directory was writable before adding the one deny rule.
 [IO.File]::WriteAllText($probe,'synthetic ACL control');Remove-Item -LiteralPath $probe
 $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User
 $deniedAcl=Get-Acl -LiteralPath $diagnostics
 $rule=[Security.AccessControl.FileSystemAccessRule]::new($sid,[Security.AccessControl.FileSystemRights]::CreateFiles,[Security.AccessControl.InheritanceFlags]::None,[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Deny)
 [void]$deniedAcl.AddAccessRule($rule);Set-Acl -LiteralPath $diagnostics -AclObject $deniedAcl;$aclChanged=$true
 $creationDenied=$false
 try{[IO.File]::WriteAllText($probe,'must be denied')}catch{if($_.Exception.GetBaseException() -is [UnauthorizedAccessException]){$creationDenied=$true}else{throw}}
 if(-not$creationDenied){throw 'The diagnostics-only CreateFiles fault was not effective; no GUI task will run.'}
 [void](Get-ChildItem -LiteralPath $diagnostics -ErrorAction Stop)
 $result.FaultControl=@{DirectoryAlreadyExisted=$true;ReadAndListAllowed=$true;CreateFilesDenied=$true;Inheritance='None';Scope='This diagnostics directory only'}
 $process=Start-Process -FilePath $InstalledExe -PassThru
 $condition=[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$process.Id)
 $deadline=[DateTime]::UtcNow.AddSeconds(30)
 do{$window=[Windows.Automation.AutomationElement]::RootElement.FindFirst([Windows.Automation.TreeScope]::Children,$condition);if($null -eq $window){Start-Sleep -Milliseconds 200}}while($null -eq $window -and [DateTime]::UtcNow -lt $deadline)
 if($null -eq $window){throw 'Owned installed GUI unavailable.'}
 [void]$process.WaitForInputIdle(5000)
 (Find-Control $window 'workspaceBox').GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($Workspace)
 $result.StartDateInput=Set-OwnedDatePicker $window (Find-Control $window 'businessDatePicker') $BusinessDate $process.Id
 $toggle=(Find-Control $window 'rangeCheckBox').GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern)
 if(($toggle.Current.ToggleState -eq [Windows.Automation.ToggleState]::On) -ne $isRange){$toggle.Toggle()}
 if($isRange){$result.EndDateInput=Set-OwnedDatePicker $window (Find-Control $window 'endDatePicker') $last $process.Id}
 Capture-Window $window (Join-Path $EvidenceDirectory '01-direct-generate-no-check.png')
 # Checking inputs writes a separate diagnostic, so a real user directly invokes Generate.
 $run=Find-Control $window 'runButton';$started=[DateTime]::UtcNow;$watch=[Diagnostics.Stopwatch]::StartNew()
 Invoke-Control $run
 $deadline=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds);$stable=0;$lastText='';$finished=$false
 do{
  $modal=$window.FindFirst([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ClassNameProperty,'#32770'))
  if($null -ne $modal){
   Capture-Window $modal (Join-Path $EvidenceDirectory 'owned-error-dialog.png')
   $result.OwnedDialogObserved=$modal.Current.Name
   $ok=$modal.FindFirst([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,'1'))
   if($null -ne $ok){Invoke-Control $ok}
  }
  $text=Get-Text (Find-Control $window 'statusText')
  if($text -eq $lastText){$stable++}else{$stable=0;$lastText=$text}
  $filesReady=@($expectedPaths|Where-Object {-not(Test-Path -LiteralPath $_ -PathType Leaf)}).Count -eq 0
  if($run.Current.IsEnabled -and $stable -ge 3 -and $watch.Elapsed.TotalMilliseconds -gt 1500 -and ($filesReady -or $text -match '\u5931\u8d25|\u62d2\u7edd|denied|Error')){$finished=$true;break}
  Start-Sleep -Milliseconds 200
 }while([DateTime]::UtcNow -lt $deadline)
 if(-not$finished){throw 'Owned GUI did not settle after diagnostics-only failure.'}
 $watch.Stop();$result.RunMilliseconds=$watch.Elapsed.TotalMilliseconds
 $result.StatusText=Get-Text (Find-Control $window 'statusText')
 $result.ResultLabel=Get-Text (Find-Control $window 'resultLabel')
 $result.OpenResultsEnabled=(Find-Control $window 'openButton').Current.IsEnabled
 $visible=$result.ResultLabel+"`r`n"+$result.StatusText
 $visiblePaths=@($expectedPaths|Where-Object {$visible.IndexOf($_,[StringComparison]::OrdinalIgnoreCase) -ge 0})
 $files=@($expectedPaths|ForEach-Object{@{Path=$_;Exists=(Test-Path -LiteralPath $_ -PathType Leaf);WrittenAfterGenerate=if(Test-Path -LiteralPath $_){(Get-Item -LiteralPath $_).LastWriteTimeUtc -ge $started.AddSeconds(-1)}else{$false}}})
 $reportsCommitted=(@($files|Where-Object {-not$_.Exists -or -not$_.WrittenAfterGenerate}).Count -eq 0)
 # A failed range also prints a "partial=0" counter; require the actual result label.
 $partialVisible=$result.ResultLabel -match '^\s*(?:\u8303\u56f4)?(?:\u90e8\u5206|\u7ed3\u679c\u4e0d\u5b8c\u6574|\u62a5\u8868\u5df2\u751f\u6210.*\u8bb0\u5f55\u672a\u4fdd\u5b58)'
 $recordFailureVisible=$visible -match '\u8bb0\u5f55.*(\u672a\u4fdd\u5b58|\u672a\u80fd|\u5931\u8d25|\u65e0\u6cd5)|\u8bca\u65ad.*(\u5931\u8d25|\u672a\u80fd|\u65e0\u6cd5)|DiagnosticWriteFailed|denied|\u8bbf\u95ee.*\u62d2\u7edd'
 $fixed=$reportsCommitted -and $visiblePaths.Count -eq $expectedPaths.Count -and $partialVisible -and $recordFailureVisible -and $result.OpenResultsEnabled
 $result.Reports=$files;$result.AllReportPathsVisible=$visiblePaths.Count -eq $expectedPaths.Count
 $result.VisibleReportPathCount=$visiblePaths.Count;$result.PartialUiSignal=$partialVisible;$result.RecordFailureUiSignal=$recordFailureVisible
 $result.ExpectedFixedSatisfied=$fixed;$result.SymptomReproduced=$reportsCommitted -and -not$fixed
 Capture-Window $window (Join-Path $EvidenceDirectory '02-diagnostics-failure-visible-result.png')
 if(-not$reportsCommitted){throw 'The I/O fault occurred before the expected report pair(s) committed; this is not the target defect.'}
 $result.Passed=if($ExpectedDisposition -eq 'Fixed'){$fixed}else{$result.SymptomReproduced}
 if(-not$result.Passed){throw 'Committed reports were not retained as visible usable partial results under the diagnostics-only I/O failure.'}
}catch{$failure=$_.Exception.Message;$result.Error=$failure}
finally{
 try{
  if($aclChanged){Set-Acl -LiteralPath $diagnostics -AclObject $originalAcl}
  $restored=(Get-Acl -LiteralPath $diagnostics).GetSecurityDescriptorSddlForm([Security.AccessControl.AccessControlSections]::All)
  $result.AclRestored=$restored -eq $originalSddl
  if(-not$result.AclRestored){$result.Passed=$false;$failure='The synthetic diagnostics ACL was not restored byte-for-byte; use the saved SDDL.';$result.AclRestoreError=$failure}
 }catch{$result.Passed=$false;$failure=$_.Exception.Message;$result.AclRestoreError=$failure}
 if(Test-Path -LiteralPath $probe){try{Remove-Item -LiteralPath $probe}catch{}}
 if($null -ne $window){try{$window.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close()}catch{}}
 Write-Json (Join-Path $EvidenceDirectory 'diagnostic-failure-ui-result.json') $result
}
$result|ConvertTo-Json -Depth 20
if($failure){throw $failure}

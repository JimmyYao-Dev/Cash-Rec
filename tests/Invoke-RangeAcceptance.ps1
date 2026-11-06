[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][string]$InstalledExe,
 [Parameter(Mandatory=$true)][string]$Workspace,
 [Parameter(Mandatory=$true)][datetime]$StartDate,
 [Parameter(Mandatory=$true)][datetime]$EndDate,
 [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
 [ValidateSet('Success','Missing','Failed','Partial')][string]$ExpectedStatus='Success',
 [string]$ExpectedVersion='0.2.0',
 [int]$TimeoutSeconds=180,
 [string]$CaseScale='synthetic-small-range'
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'InstalledUiHelpers.ps1')
$InstalledExe=(Resolve-Path -LiteralPath $InstalledExe).Path
$Workspace=(Resolve-Path -LiteralPath $Workspace).Path
$diagnosticWorkspace=$Workspace
if((Split-Path -Leaf $Workspace) -match '^\d{4}\.\d{2}\.\d{2}$'){$diagnosticWorkspace=Split-Path -Parent $Workspace}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$EvidenceDirectory=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
$window=$null;$process=$null
$result=[ordered]@{Entry='Installed real range GUI via UI Automation';Passed=$false;Workspace=$Workspace;DiagnosticWorkspace=$diagnosticWorkspace;StartDate=$StartDate.ToString('yyyy-MM-dd');EndDate=$EndDate.ToString('yyyy-MM-dd');ExpectedStatus=$ExpectedStatus}
$result.Measurements=[ordered]@{CaseScale=$CaseScale;OS=[Environment]::OSVersion.VersionString;Architecture=$env:PROCESSOR_ARCHITECTURE;ProductVersion=[Diagnostics.FileVersionInfo]::GetVersionInfo($InstalledExe).ProductVersion;InstalledExecutableBytes=(Get-Item -LiteralPath $InstalledExe).Length}
try{
 $watch=[Diagnostics.Stopwatch]::StartNew()
 $process=Start-Process -FilePath $InstalledExe -PassThru
 $condition=[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$process.Id)
 $deadline=[DateTime]::UtcNow.AddSeconds(30)
 do{$window=[Windows.Automation.AutomationElement]::RootElement.FindFirst([Windows.Automation.TreeScope]::Children,$condition);if($null -eq $window){Start-Sleep -Milliseconds 200}}while($null -eq $window -and [DateTime]::UtcNow -lt $deadline)
 if($null -eq $window){throw 'Actual installed range window unavailable on this desktop.'}
 [void]$process.WaitForInputIdle(5000);$watch.Stop();$result.Measurements.LaunchToWindowMilliseconds=$watch.Elapsed.TotalMilliseconds
 $result.WindowTitle=$window.Current.Name
 if($ExpectedVersion -and ($result.WindowTitle -notlike ('*'+$ExpectedVersion+'*') -or -not $result.Measurements.ProductVersion.StartsWith($ExpectedVersion))){throw 'Installed range window version differs.'}
 (Find-Control $window 'workspaceBox').GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($Workspace)
 $startControl=Find-Control $window 'businessDatePicker'
 $result.StartDateInput=Set-OwnedDatePicker $window $startControl $StartDate $process.Id
 $toggle=(Find-Control $window 'rangeCheckBox').GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern)
 if($toggle.Current.ToggleState -ne [Windows.Automation.ToggleState]::On){$toggle.Toggle()}
 $endControl=$null;$deadline=[DateTime]::UtcNow.AddSeconds(5)
 do{try{$endControl=Find-Control $window 'endDatePicker'}catch{};if($null -eq $endControl -or $endControl.Current.IsOffscreen){Start-Sleep -Milliseconds 100}}while(($null -eq $endControl -or $endControl.Current.IsOffscreen) -and [DateTime]::UtcNow -lt $deadline)
 if($null -eq $endControl -or $endControl.Current.IsOffscreen){throw 'End date picker was not shown by the actual range checkbox.'}
 $result.EndDateInput=Set-OwnedDatePicker $window $endControl $EndDate $process.Id
 $result.RangeImportDisabled=-not(Find-Control $window 'importButton').Current.IsEnabled
 if(-not$result.RangeImportDisabled){throw 'Single-date import remained enabled in range mode.'}
 Capture-Window $window (Join-Path $EvidenceDirectory '01-range-selected.png')
 $run=Find-Control $window 'runButton';$since=[DateTime]::UtcNow;$watch=[Diagnostics.Stopwatch]::StartNew()
 Invoke-Control $run
 $deadline=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds);$batch=$null
 $result.BusyControls=[ordered]@{ObservedDuringRun=$false;AllDateWorkspaceControlsDisabled=$null}
 do{
  if(-not$run.Current.IsEnabled){
   $enabled=@('browseButton','businessDatePicker','endDatePicker','rangeCheckBox')|ForEach-Object{(Find-Control $window $_).Current.IsEnabled}
   if(-not$run.Current.IsEnabled){$result.BusyControls.ObservedDuringRun=$true;$result.BusyControls.AllDateWorkspaceControlsDisabled=($enabled -notcontains $true);if($enabled -contains $true){throw 'Date or workspace controls remained enabled while range processing was active.'}}
  }
  foreach($file in Get-ChildItem -LiteralPath (Join-Path $diagnosticWorkspace '.cashrec/diagnostics') -Filter 'batch-*.json' -ErrorAction SilentlyContinue|Where-Object {$_.LastWriteTimeUtc -ge $since.AddSeconds(-1)}|Sort-Object LastWriteTimeUtc -Descending){
   try{$candidate=Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8|ConvertFrom-Json}catch{continue}
   if($candidate.Batch -and $candidate.Operation -eq 'Generate' -and ([datetime]$candidate.StartDate).Date -eq $StartDate.Date -and ([datetime]$candidate.EndDate).Date -eq $EndDate.Date){$batch=@{File=$file.FullName;Value=$candidate};break}
  }
  if($null -ne $batch -and $run.Current.IsEnabled){break}
  Start-Sleep -Milliseconds 200
 }while([DateTime]::UtcNow -lt $deadline)
 if($null -eq $batch -or -not$run.Current.IsEnabled){throw 'The actual GUI did not finish the requested range with a matching batch diagnostic.'}
 $watch.Stop();$process.Refresh();$result.Measurements.RunToBatchDiagnosticMilliseconds=$watch.Elapsed.TotalMilliseconds;$result.Measurements.PeakWorkingSetBytes=$process.PeakWorkingSet64
 $result.DiagnosticFile=$batch.File;$result.Batch=$batch.Value;$result.StatusText=Get-Text (Find-Control $window 'statusText')
 if($batch.Value.Status -ne $ExpectedStatus){throw ('Actual range status '+$batch.Value.Status+' differs from '+$ExpectedStatus)}
 if([bool]$batch.Value.Complete -ne ($ExpectedStatus -eq 'Success')){throw 'Range overall complete flag contradicts its expected status.'}
 Capture-Window $window (Join-Path $EvidenceDirectory '02-range-result.png')
 $result.Passed=$true;Write-Json (Join-Path $EvidenceDirectory 'range-ui-result.json') $result
 $result|ConvertTo-Json -Depth 20
}catch{
 $result.Error=$_.Exception.Message
 if($null -ne $window){try{Capture-Window $window (Join-Path $EvidenceDirectory 'failure.png')}catch{}}
 Write-Json (Join-Path $EvidenceDirectory 'range-ui-result.json') $result;throw
}finally{
 if($null -ne $window){try{$window.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close()}catch{}}
}

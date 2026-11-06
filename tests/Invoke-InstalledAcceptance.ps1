[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][string]$InstalledExe,
 [Parameter(Mandatory=$true)][string]$Workspace,
 [datetime]$BusinessDate = '2026-10-25',
 [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
 [ValidateSet('Complete','Partial','Failure','PreflightReady','PreflightFailure')][string]$ExpectedOutcome='Complete',
 [string]$ExpectedMonthlyStage,
 [ValidateSet('Complete','Incomplete')][string]$ExpectedMonthlyStageOutcome,
 [string]$ExpectedVersion,
 [int]$TimeoutSeconds=120,
 [string]$CaseScale='synthetic-small',
 [switch]$PreflightOnly,
 [switch]$IsolateRuntimeEnvironment,
 [string]$ImportFile,
 [switch]$KeepOpen
)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
if (-not ('CashRecUiaNative' -as [type])) {
 Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CashRecUiaNative {
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd,IntPtr hdc,uint flags);
 [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 public static int ForegroundProcess() {uint pid;GetWindowThreadProcessId(GetForegroundWindow(),out pid);return (int)pid;}
}
'@
}
function Write-Json($Path,$Object) {
 [IO.File]::WriteAllText($Path,($Object|ConvertTo-Json -Depth 15),[Text.UTF8Encoding]::new($false))
}
function Find-Control($Window,$Id) {
 $condition=[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,$Id)
 $control=$Window.FindFirst([Windows.Automation.TreeScope]::Descendants,$condition)
 if($null -eq $control) {
  $condition=[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,$Id)
  $control=$Window.FindFirst([Windows.Automation.TreeScope]::Descendants,$condition)
 }
 if($null -eq $control){throw "Real installed UI control not found: $Id"}
 return $control
}
function Get-Text($Control) {
 $valuePattern=$null
 if($Control.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$valuePattern)){return $valuePattern.Current.Value}
 return $Control.Current.Name
}
function Invoke-Control($Control) {
 $pattern=$Control.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern)
 $pattern.Invoke()
}
function Assert-OwnForeground([int]$ProcessId) {
 if([CashRecUiaNative]::ForegroundProcess() -ne $ProcessId){throw 'Keyboard focus is outside the test application; no keystrokes will be sent.'}
}
function Capture-Window($Window,$Path) {
 [void][CashRecUiaNative]::SetForegroundWindow([IntPtr]$Window.Current.NativeWindowHandle)
 Start-Sleep -Milliseconds 200
 $rect=$Window.Current.BoundingRectangle
 $image=[Drawing.Bitmap]::new([int]$rect.Width,[int]$rect.Height)
 $graphics=[Drawing.Graphics]::FromImage($image)
 try {
  $dc=$graphics.GetHdc()
  try{$captured=[CashRecUiaNative]::PrintWindow([IntPtr]$Window.Current.NativeWindowHandle,$dc,2)}finally{$graphics.ReleaseHdc($dc)}
  if(-not $captured){throw 'Could not capture the owned test window; no desktop screenshot will be substituted.'}
  $stream=[IO.MemoryStream]::new()
  try {$image.Save($stream,[Drawing.Imaging.ImageFormat]::Png);[IO.File]::WriteAllBytes($Path,$stream.ToArray())}finally{$stream.Dispose()}
 } finally {$graphics.Dispose();$image.Dispose()}
}
function Wait-Diagnostic($Since,$Window,$Button) {
 $deadline=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
 while([DateTime]::UtcNow -lt $deadline){
  if($Button.Current.AutomationId -eq 'runButton' -and -not $Button.Current.IsEnabled){
   $browse=Find-Control $Window 'browseButton'
   $browseEnabled=$browse.Current.IsEnabled
   # UIA reads occur sequentially: completion may happen between the two controls.
   if(-not $browseEnabled){
    $result.BusyControls.ObservedDuringRun=$true
    $result.BusyControls.BrowseDisabled=$true
   }elseif(-not $Button.Current.IsEnabled){
    throw 'Workspace browse remained enabled during the actual installed GUI run.'
   }
  }
  $file=Get-ChildItem -LiteralPath (Join-Path $Workspace '.cashrec/diagnostics') -Filter ($BusinessDate.ToString('yyyy.MM.dd')+'-*.json') -ErrorAction SilentlyContinue |
   Where-Object {$_.LastWriteTimeUtc -ge $Since.AddSeconds(-1)} | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
  if($null -ne $file -and $Button.Current.IsEnabled){
   try {return @{File=$file.FullName;Value=(Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8|ConvertFrom-Json)}} catch {}
  }
  Start-Sleep -Milliseconds 200
 }
 throw "Installed GUI did not produce a completed diagnostic within $TimeoutSeconds seconds."
}
$InstalledExe=(Resolve-Path -LiteralPath $InstalledExe).Path
$Workspace=(Resolve-Path -LiteralPath $Workspace).Path
New-Item -ItemType Directory -Path $EvidenceDirectory -Force|Out-Null
$EvidenceDirectory=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
$process=$null;$window=$null;$launchWatch=$null;$runWatch=$null
$result=[ordered]@{Entry='Installed real GUI via UI Automation';Executable=$InstalledExe;Workspace=$Workspace;BusinessDate=$BusinessDate.ToString('yyyy-MM-dd');Expected=$ExpectedOutcome;Passed=$false}
$result.Measurements=[ordered]@{CaseScale=$CaseScale;OS=[Environment]::OSVersion.VersionString;Architecture=$env:PROCESSOR_ARCHITECTURE;AppFileVersion=[Diagnostics.FileVersionInfo]::GetVersionInfo($InstalledExe).FileVersion;InstalledExecutableBytes=(Get-Item -LiteralPath $InstalledExe).Length;LaunchToWindowMilliseconds=$null;RunToDiagnosticMilliseconds=$null;PeakWorkingSetBytes=$null;PreparedInputFiles=$null;PreparedInputBytes=$null}
try {
 $launchWatch=[Diagnostics.Stopwatch]::StartNew()
 if($IsolateRuntimeEnvironment){
  $child=[Diagnostics.ProcessStartInfo]::new()
  $child.FileName=$InstalledExe
  $child.WorkingDirectory=Split-Path -Parent $InstalledExe
  $child.UseShellExecute=$false
  $child.EnvironmentVariables['PATH']=(Join-Path $env:WINDIR 'System32')+';'+$env:WINDIR
  $absentRuntime=Join-Path ([IO.Path]::GetTempPath()) ('CashRec-absent-runtime-'+[guid]::NewGuid().ToString('N'))
  foreach($key in @('DOTNET_ROOT','DOTNET_ROOT_X64','DOTNET_ROOT_X86','DOTNET_ROOT(x86)')){$child.EnvironmentVariables[$key]=$absentRuntime}
  $child.EnvironmentVariables['DOTNET_MULTILEVEL_LOOKUP']='0'
  $child.EnvironmentVariables.Remove('PYTHONHOME')
  $child.EnvironmentVariables.Remove('PYTHONPATH')
  $result.RuntimeIsolation=@{Enabled=$true;Scope='Child process environment only';Path=$child.EnvironmentVariables['PATH'];DotnetRoot=$absentRuntime;MultilevelLookup='0';PythonVariablesRemoved=$true;CleanVirtualMachine=$false}
  $process=[Diagnostics.Process]::Start($child)
 }else{
  $result.RuntimeIsolation=@{Enabled=$false}
  $process=Start-Process -FilePath $InstalledExe -PassThru
 }
 $deadline=[DateTime]::UtcNow.AddSeconds(30)
 $condition=[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$process.Id)
 do {
  $window=[Windows.Automation.AutomationElement]::RootElement.FindFirst([Windows.Automation.TreeScope]::Children,$condition)
  if($null -eq $window){Start-Sleep -Milliseconds 200}
 }while($null -eq $window -and [DateTime]::UtcNow -lt $deadline)
 if($null -eq $window){throw 'Installed window unavailable. On Windows, run this harness in the interactive user desktop; the restricted sandbox desktop cannot see user windows.'}
 $launchWatch.Stop()
 [void]$process.WaitForInputIdle(5000)
 $result.Measurements.LaunchToWindowMilliseconds=[math]::Round($launchWatch.Elapsed.TotalMilliseconds,1)
 $result.WindowTitle=$window.Current.Name
 if($ExpectedVersion -and ($result.WindowTitle -notlike ('*'+$ExpectedVersion+'*') -or -not $result.Measurements.AppFileVersion.StartsWith($ExpectedVersion))){throw 'The installed window or executable version differs from the expected release.'}
 $workspaceControl=Find-Control $window 'workspaceBox'
 $workspaceControl.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($Workspace)
 $dateControl=Find-Control $window 'businessDatePicker'
 $result.DateInputAttempts=@()
 for($attempt=1;$attempt -le 3;$attempt++){
  [void][CashRecUiaNative]::SetForegroundWindow([IntPtr]$window.Current.NativeWindowHandle)
  $dateControl.SetFocus()
  Start-Sleep -Milliseconds 100
  $record=[ordered]@{Attempt=$attempt;Focused=$dateControl.Current.HasKeyboardFocus;Segments=@()}
  if(-not $record.Focused){$record.Observed=[string]$dateControl.GetCurrentPropertyValue([Windows.Automation.ValuePattern]::ValueProperty);$result.DateInputAttempts+=$record;continue}
  $segmentIndex=-1
  foreach($keys in @('{HOME}',$BusinessDate.ToString('yyyy'),'{RIGHT}',$BusinessDate.ToString('MM'),'{RIGHT}',$BusinessDate.ToString('dd'),'{TAB}')){
   $segmentIndex++
   Assert-OwnForeground $process.Id
   [Windows.Forms.SendKeys]::SendWait($keys)
   Start-Sleep -Milliseconds 80
   $segmentValue=[string]$dateControl.GetCurrentPropertyValue([Windows.Automation.ValuePattern]::ValueProperty)
   $record.Segments+=$segmentValue
   if($segmentIndex -in @(1,3,5)){
    $parsedSegment=[datetime]::ParseExact($segmentValue,'yyyy.MM.dd',[Globalization.CultureInfo]::InvariantCulture)
    if($parsedSegment.Year -ne $BusinessDate.Year -or ($segmentIndex -ge 3 -and $parsedSegment.Month -ne $BusinessDate.Month) -or ($segmentIndex -ge 5 -and $parsedSegment.Day -ne $BusinessDate.Day)){$record.InvalidSegment=$segmentIndex;break}
   }
  }
  $observedDate=$dateControl.GetCurrentPropertyValue([Windows.Automation.ValuePattern]::ValueProperty)
  $record.Observed=[string]$observedDate
  $result.DateInputAttempts+=$record
  if([string]$observedDate -eq $BusinessDate.ToString('yyyy.MM.dd')){break}
 }
 if([string]$observedDate -ne $BusinessDate.ToString('yyyy.MM.dd')){throw 'The actual date picker value did not change to the requested date.'}
 if($ImportFile){
  $ImportFile=(Resolve-Path -LiteralPath $ImportFile).Path
  $hashBefore=(Get-FileHash -LiteralPath $ImportFile -Algorithm SHA256).Hash
  $import=Find-Control $window 'importButton'
  $import.SetFocus()
  Assert-OwnForeground $process.Id
  [Windows.Forms.SendKeys]::SendWait(' ')
  $dialog=$null;$dialogDeadline=[DateTime]::UtcNow.AddSeconds(30)
  do{
   $dialog=$window.FindFirst([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ClassNameProperty,'#32770'))
   if($null -eq $dialog){Start-Sleep -Milliseconds 200}
  }while($null -eq $dialog -and [DateTime]::UtcNow -lt $dialogDeadline)
  if($null -eq $dialog){throw 'The actual import file dialog was not observed.'}
  $fileName=$null;$fileNamePattern=$null
  $fileNames=$dialog.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,'1148'))
  foreach($candidate in $fileNames){$candidatePattern=$null;if($candidate.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$candidatePattern)){$fileName=$candidate;$fileNamePattern=$candidatePattern;break}}
  Capture-Window $dialog (Join-Path $EvidenceDirectory '00-import-file-dialog.png')
  if($null -eq $fileNamePattern){
   Assert-OwnForeground $process.Id
   [Windows.Forms.SendKeys]::SendWait('%n')
   [Windows.Forms.SendKeys]::SendWait('^a')
   $escapedPath=[regex]::Replace($ImportFile,'[+^%~(){}\[\]]',{param($match) '{'+$match.Value+'}'})
   [Windows.Forms.SendKeys]::SendWait($escapedPath+'{ENTER}')
  }else{
   $fileNamePattern.SetValue($ImportFile)
   $openCondition=[Windows.Automation.AndCondition]::new([Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,'1'),[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::Button))
   $open=$dialog.FindFirst([Windows.Automation.TreeScope]::Descendants,$openCondition)
   if($null -eq $open){throw 'The actual import dialog Open button was not found.'}
   Invoke-Control $open
  }
  Start-Sleep -Milliseconds 500
  if((Get-FileHash -LiteralPath $ImportFile -Algorithm SHA256).Hash -ne $hashBefore){throw 'Import modified the original source file.'}
  $result.Import=@{Source=$ImportFile;OriginalSha256=$hashBefore;OriginalUnchanged=$true}
  Capture-Window $window (Join-Path $EvidenceDirectory '00-imported.png')
 }
 $check=Find-Control $window 'checkButton'
 $start=[DateTime]::UtcNow
 Invoke-Control $check
 $preflight=Wait-Diagnostic $start $window $check
 if(([datetime]$preflight.Value.BusinessDate).Date -ne $BusinessDate.Date){throw 'The real date control did not accept the requested date.'}
 $result.Preflight=$preflight.Value
 $preparedFiles=@($preflight.Value.Inputs.PSObject.Properties|ForEach-Object{Get-Item -LiteralPath $_.Value -ErrorAction SilentlyContinue})
 $result.Measurements.PreparedInputFiles=$preparedFiles.Count
 $result.Measurements.PreparedInputBytes=($preparedFiles|Measure-Object -Property Length -Sum).Sum
 Capture-Window $window (Join-Path $EvidenceDirectory '01-preflight.png')
 if($PreflightOnly){
  $outcome=if($preflight.Value.CanRun){'PreflightReady'}else{'PreflightFailure'}
 }else{
  $run=Find-Control $window 'runButton'
  $start=[DateTime]::UtcNow
  $runWatch=[Diagnostics.Stopwatch]::StartNew()
  $result.BusyControls=[ordered]@{ObservedDuringRun=$false;BrowseDisabled=$null;UnobservedMeaning='Processing may finish before UI Automation observes the busy state.'}
  Invoke-Control $run
  $runDiagnostic=$null
  $deadline=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
  do {
   $candidate=Wait-Diagnostic $start $window $run
   if(($candidate.Value.PSObject.Properties.Name -contains 'Complete') -and
      (-not $ExpectedMonthlyStage -or ($candidate.Value.PSObject.Properties.Name -contains 'MonthlyStage'))){$runDiagnostic=$candidate}
   if($null -eq $runDiagnostic){Start-Sleep -Milliseconds 200}
  }while($null -eq $runDiagnostic -and [DateTime]::UtcNow -lt $deadline)
  if($null -eq $runDiagnostic){throw 'The generated report diagnostic was not observed.'}
  $runWatch.Stop()
  $result.Measurements.RunToDiagnosticMilliseconds=[math]::Round($runWatch.Elapsed.TotalMilliseconds,1)
  $result.DiagnosticFile=$runDiagnostic.File
  $result.Run=$runDiagnostic.Value
  if($ExpectedMonthlyStage){
   if($runDiagnostic.Value.MonthlyStage -ne $ExpectedMonthlyStage){throw 'The actual installed task reported a different monthly stage.'}
   if($ExpectedMonthlyStageOutcome){
    $wantedStageComplete=$ExpectedMonthlyStageOutcome -eq 'Complete'
    if([bool]$runDiagnostic.Value.MonthlyStageComplete -ne $wantedStageComplete){throw 'The actual installed monthly stage completion did not match the expected outcome.'}
    if($wantedStageComplete -and -not(Test-Path -LiteralPath $runDiagnostic.Value.MonthlyOutput -PathType Leaf)){throw 'The actual installed monthly stage reported completion without its output workbook.'}
   }
  }
  $outcome=if($runDiagnostic.Value.Complete){'Complete'}elseif($runDiagnostic.Value.Output){'Partial'}else{'Failure'}
  Capture-Window $window (Join-Path $EvidenceDirectory '02-generated.png')
 }
 $result.Observed=$outcome
 $result.Status=Get-Text (Find-Control $window 'statusText')
 $process.Refresh()
 $result.Measurements.PeakWorkingSetBytes=$process.PeakWorkingSet64
 $result.Passed=$outcome -eq $ExpectedOutcome
 if(-not $result.Passed){throw "Expected $ExpectedOutcome but installed GUI reported $outcome"}
 Write-Json (Join-Path $EvidenceDirectory 'ui-result.json') $result
 $result|ConvertTo-Json -Depth 15
} catch {
 $result.Error=$_.Exception.Message
 if($null -ne $window){try {Capture-Window $window (Join-Path $EvidenceDirectory 'failure.png')}catch {}}
 Write-Json (Join-Path $EvidenceDirectory 'ui-result.json') $result
 throw
} finally {
 if(-not $KeepOpen -and $null -ne $window){
  try {$modal=$window.FindFirst([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ClassNameProperty,'#32770'));if($null -ne $modal){$modal.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close()}}catch {}
  try {$window.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close()}catch {}
 }
}


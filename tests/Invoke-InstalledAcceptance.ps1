[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][string]$InstalledExe,
 [Parameter(Mandatory=$true)][string]$Workspace,
 [datetime]$BusinessDate = '2026-10-25',
 [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
 [ValidateSet('Complete','Partial','Failure','PreflightReady','PreflightFailure')][string]$ExpectedOutcome='Complete',
 [int]$TimeoutSeconds=120,
 [string]$CaseScale='synthetic-small',
 [switch]$PreflightOnly,
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
  $graphics.CopyFromScreen([int]$rect.X,[int]$rect.Y,0,0,$image.Size)
  $stream=[IO.MemoryStream]::new()
  try {$image.Save($stream,[Drawing.Imaging.ImageFormat]::Png);[IO.File]::WriteAllBytes($Path,$stream.ToArray())}finally{$stream.Dispose()}
 } finally {$graphics.Dispose();$image.Dispose()}
}
function Wait-Diagnostic($Since,$Window,$Button) {
 $deadline=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
 while([DateTime]::UtcNow -lt $deadline){
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
 $process=Start-Process -FilePath $InstalledExe -PassThru
 $deadline=[DateTime]::UtcNow.AddSeconds(30)
 $condition=[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$process.Id)
 do {
  $window=[Windows.Automation.AutomationElement]::RootElement.FindFirst([Windows.Automation.TreeScope]::Children,$condition)
  if($null -eq $window){Start-Sleep -Milliseconds 200}
 }while($null -eq $window -and [DateTime]::UtcNow -lt $deadline)
 if($null -eq $window){throw 'Installed window unavailable. On Windows, run this harness in the interactive user desktop; the restricted sandbox desktop cannot see user windows.'}
 $launchWatch.Stop()
 $result.Measurements.LaunchToWindowMilliseconds=[math]::Round($launchWatch.Elapsed.TotalMilliseconds,1)
 $workspaceControl=Find-Control $window 'workspaceBox'
 $workspaceControl.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).SetValue($Workspace)
 $dateControl=Find-Control $window 'businessDatePicker'
 [void][CashRecUiaNative]::SetForegroundWindow([IntPtr]$window.Current.NativeWindowHandle)
 $dateControl.SetFocus()
 Assert-OwnForeground $process.Id
 [Windows.Forms.SendKeys]::SendWait('{HOME}'+$BusinessDate.ToString('yyyy')+'{RIGHT}'+$BusinessDate.ToString('MM')+'{RIGHT}'+$BusinessDate.ToString('dd')+'{TAB}')
 $observedDate=$dateControl.GetCurrentPropertyValue([Windows.Automation.ValuePattern]::ValueProperty)
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
  Invoke-Control $run
  $runDiagnostic=$null
  $deadline=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
  do {
   $candidate=Wait-Diagnostic $start $window $run
   if($candidate.Value.PSObject.Properties.Name -contains 'Complete'){$runDiagnostic=$candidate}
   if($null -eq $runDiagnostic){Start-Sleep -Milliseconds 200}
  }while($null -eq $runDiagnostic -and [DateTime]::UtcNow -lt $deadline)
  if($null -eq $runDiagnostic){throw 'The generated report diagnostic was not observed.'}
  $runWatch.Stop()
  $result.Measurements.RunToDiagnosticMilliseconds=[math]::Round($runWatch.Elapsed.TotalMilliseconds,1)
  $result.DiagnosticFile=$runDiagnostic.File
  $result.Run=$runDiagnostic.Value
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


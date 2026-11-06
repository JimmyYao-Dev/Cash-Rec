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
function Set-OwnedDatePicker($Window,$Control,[datetime]$RequestedDate,[int]$OwnedProcessId) {
 $attempts=@()
 for($attempt=1;$attempt -le 3;$attempt++){
  [void][CashRecUiaNative]::SetForegroundWindow([IntPtr]$window.Current.NativeWindowHandle)
  $Control.SetFocus()
  Start-Sleep -Milliseconds 100
  $record=[ordered]@{Attempt=$attempt;Focused=$Control.Current.HasKeyboardFocus;Segments=@()}
  if(-not $record.Focused){$record.Observed=[string]$Control.GetCurrentPropertyValue([Windows.Automation.ValuePattern]::ValueProperty);$attempts+=$record;continue}
  $segmentIndex=-1
  foreach($keys in @('{HOME}',$RequestedDate.ToString('yyyy'),'{RIGHT}',$RequestedDate.ToString('MM'),'{RIGHT}',$RequestedDate.ToString('dd'),'{TAB}')){
   $segmentIndex++
   Assert-OwnForeground $OwnedProcessId
   [Windows.Forms.SendKeys]::SendWait($keys)
   Start-Sleep -Milliseconds 80
   $segmentValue=[string]$Control.GetCurrentPropertyValue([Windows.Automation.ValuePattern]::ValueProperty)
   $record.Segments+=$segmentValue
   if($segmentIndex -in @(1,3,5)){
    $parsedSegment=[datetime]::ParseExact($segmentValue,'yyyy.MM.dd',[Globalization.CultureInfo]::InvariantCulture)
    if($parsedSegment.Year -ne $RequestedDate.Year -or ($segmentIndex -ge 3 -and $parsedSegment.Month -ne $RequestedDate.Month) -or ($segmentIndex -ge 5 -and $parsedSegment.Day -ne $RequestedDate.Day)){$record.InvalidSegment=$segmentIndex;break}
   }
  }
  $observedDate=$Control.GetCurrentPropertyValue([Windows.Automation.ValuePattern]::ValueProperty)
  $record.Observed=[string]$observedDate
  $attempts+=$record
  if([string]$observedDate -eq $RequestedDate.ToString('yyyy.MM.dd')){break}
 }
 if([string]$observedDate -ne $RequestedDate.ToString('yyyy.MM.dd')){throw 'The actual date picker value did not change to the requested date.'}
 return [pscustomobject]@{Attempts=$attempts;Observed=[string]$observedDate}
}

[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][string]$Workspace,
 [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
 [string]$BusinessDate='2026-10-25',
 [string]$PreviousDate='2026-10-22'
)
$ErrorActionPreference='Stop'
$Workspace=(Resolve-Path -LiteralPath $Workspace).Path
$marker=Join-Path $Workspace '.cashrec-monthly-test-fixture.json'
if(-not(Test-Path -LiteralPath $marker) -or -not(Get-Content -LiteralPath $marker -Raw|ConvertFrom-Json).synthetic_only){throw 'Only explicitly marked synthetic workspaces may be edited.'}
[IO.Directory]::CreateDirectory($EvidenceDirectory)|Out-Null
$EvidenceDirectory=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CashRecRerunExcel {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [StructLayout(LayoutKind.Sequential)] public struct Rect {public int Left,Top,Right,Bottom;}
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd,out Rect rect);
 [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd,IntPtr hdc,uint flags);
}
'@
function Find-DateRow($Sheet,[string]$Date,[int]$First){
 for($r=$First;$r -le $Sheet.UsedRange.Rows.Count;$r++){
  $v=$Sheet.Cells.Item($r,1).Value2
  if($v -is [double] -and [datetime]::FromOADate($v).ToString('yyyy-MM-dd') -eq $Date){return $r}
 }
 throw ('Missing setup date '+$Date+' on '+$Sheet.Name)
}
function Capture-Owned($Excel,[string]$Path){
 $rect=[CashRecRerunExcel+Rect]::new();[void][CashRecRerunExcel]::GetWindowRect([IntPtr]$Excel.Hwnd,[ref]$rect)
 $bitmap=[Drawing.Bitmap]::new($rect.Right-$rect.Left,$rect.Bottom-$rect.Top);$graphics=[Drawing.Graphics]::FromImage($bitmap)
 try{$dc=$graphics.GetHdc();try{$ok=[CashRecRerunExcel]::PrintWindow([IntPtr]$Excel.Hwnd,$dc,2)}finally{$graphics.ReleaseHdc($dc)}
  if(-not $ok){throw 'Owned Excel capture failed'}
  $stream=[IO.MemoryStream]::new();try{$bitmap.Save($stream,[Drawing.Imaging.ImageFormat]::Png);[IO.File]::WriteAllBytes($Path,$stream.ToArray())}finally{$stream.Dispose()}
 }finally{$graphics.Dispose();$bitmap.Dispose()}
}
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$excel=$null;$book=$null;$owned=$false
$result=[ordered]@{Entry='Actual owned Excel; intentional synthetic rerun precondition';Passed=$false;ExistingExcelPids=$existing;Workspace=$Workspace;BusinessDate=$BusinessDate;PreviousDate=$PreviousDate;ManualDEP=@();MonthlyManual=@()}
try{
 $excel=New-Object -ComObject Excel.Application
 [uint32]$created=0;[void][CashRecRerunExcel]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$created)
 $owned=$existing -notcontains [int]$created
 if(-not $owned){throw 'Existing user Excel instance will not be edited or closed.'}
 $result.ExcelPid=$created;$result.OwnNewInstance=$true
 $excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
 $daily=Join-Path $Workspace ($BusinessDate.Replace('-','.')+'/CashRec_Base.xlsx')
 $book=$excel.Workbooks.Open($daily,0,$false);$sheet=$book.Worksheets.Item('Equity')
 $headers=@{};for($c=1;$c -le $sheet.UsedRange.Columns.Count;$c++){$headers[[string]$sheet.Cells.Item(1,$c).Value2]=$c}
 $manual=@{'TMA/Standard/USD'=30;'TMA/IB/USD'=0;'TMA/IB/AUD'=-4;'TMG/Standard/USC'=$null;'TMG/Standard/EUR'=7}
 for($r=2;$r -le $sheet.UsedRange.Rows.Count;$r++){
  $key=([string]$sheet.Cells.Item($r,$headers.Branch).Value2)+'/'+([string]$sheet.Cells.Item($r,$headers.'Retail Type').Value2)+'/'+([string]$sheet.Cells.Item($r,$headers.Currency).Value2)
  if(-not $manual.ContainsKey($key)){throw 'Unexpected synthetic Equity key in Excel setup.'}
  $cell=$sheet.Cells.Item($r,$headers.DEP)
  if($null -eq $manual[$key]){$cell.ClearContents()|Out-Null}else{$cell.Value2=[double]$manual[$key]}
  $cell.NumberFormat='$#,##0.0000;[Red]-$#,##0.0000'
  $result.ManualDEP+=@{Key=$key;Value=$cell.Value2;NumberFormat=$cell.NumberFormat}
 }
 # Real Excel sorting and tab movement are intentional inputs to the rerun test.
 $range=$sheet.Range($sheet.Cells.Item(2,1),$sheet.Cells.Item($sheet.UsedRange.Rows.Count,$sheet.UsedRange.Columns.Count))
 [void]$range.Sort($sheet.Cells.Item(2,$headers.Currency),2,$null,$null,1,$null,1,2)
 $book.Worksheets.Item('Bank').Move($book.Worksheets.Item(1))
 $excel.CalculateFull();$sheet.Activate();$excel.ActiveWindow.Zoom=85
 $book.Save();$book.SaveCopyAs((Join-Path $EvidenceDirectory 'seeded-daily.xlsx'))
 Capture-Owned $excel (Join-Path $EvidenceDirectory 'excel-seeded-dep.png')
 $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
 $monthly=Join-Path $Workspace ('monthly/CashRec_Breakdown_'+$BusinessDate.Substring(0,7).Replace('-','.')+'.xlsx')
 $book=$excel.Workbooks.Open($monthly,0,$false);$sheet=$book.Worksheets.Item('Break down')
 foreach($day in @($BusinessDate,$PreviousDate)){
  $r=Find-DateRow $sheet $day 2
  $numbers=if($day -eq $BusinessDate){@{7=-5;12=0;13=$null;14=-2.5}}else{@{7=0;12=-3;13=$null;14=$null}}
  foreach($column in @(7,12,13,14)){
   $cell=$sheet.Cells.Item($r,$column)
   if($null -eq $numbers[$column]){$cell.ClearContents()|Out-Null}else{$cell.Value2=[double]$numbers[$column]}
   if($column -eq 7){$cell.NumberFormat='$#,##0.0000;[Red]-$#,##0.0000'}
   $result.MonthlyManual+=@{Date=$day;Column=$column;Value=$cell.Value2;NumberFormat=$cell.NumberFormat}
  }
 }
 foreach($s in $book.Worksheets){
  $first=if($s.Name -eq 'Break down'){2}else{3}
  $range=$s.Range($s.Cells.Item($first,1),$s.Cells.Item($s.UsedRange.Rows.Count,$s.UsedRange.Columns.Count))
  [void]$range.Sort($s.Cells.Item($first,1),2,$null,$null,1,$null,1,2)
 }
 $sheet.Move($book.Worksheets.Item(1));$excel.CalculateFull();$sheet.Activate();$excel.ActiveWindow.Zoom=85
 $book.Save();$book.SaveCopyAs((Join-Path $EvidenceDirectory 'seeded-monthly.xlsx'))
 Capture-Owned $excel (Join-Path $EvidenceDirectory 'excel-seeded-monthly.png')
 $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
 $result.DailySHA256=(Get-FileHash -LiteralPath $daily -Algorithm SHA256).Hash
 $result.MonthlySHA256=(Get-FileHash -LiteralPath $monthly -Algorithm SHA256).Hash
 $result.Passed=$true
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'excel-rerun-setup.json'),($result|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
 $result|ConvertTo-Json -Depth 12
}catch{
 $result.Error=$_.Exception.Message
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'excel-rerun-setup.json'),($result|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false));throw
}finally{
 if($null -ne $book){try{$book.Close($false)}catch{}}
 if($owned -and $null -ne $excel){try{$excel.Quit()}catch{}}
 foreach($item in @($sheet,$range,$book,$excel)){if($null -ne $item){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)}catch{}}}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

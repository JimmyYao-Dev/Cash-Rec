[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][string]$Workbook,
 [Parameter(Mandatory=$true)][string]$Oracle,
 [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
 [ValidateSet('ValidateBalances','EditCheck')][string]$Mode='ValidateBalances'
)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
if(-not ('CashRecMonthlyExcelWindow' -as [type])){
 Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CashRecMonthlyExcelWindow {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [StructLayout(LayoutKind.Sequential)] public struct Rect {public int Left,Top,Right,Bottom;}
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd,out Rect rect);
 [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd,IntPtr hdc,uint flags);
}
'@
}
function Assert-Number($Actual,$Expected,$Context){
 if($null -eq $Actual -or $Actual -is [string] -or [math]::Abs([double]$Actual-[double]$Expected) -gt 0.00000001){throw ($Context+' differs from the independent synthetic expectation.')}
}
function Normalize-Header($Value){return ([regex]::Replace([string]$Value,'[^A-Za-z0-9]','')).ToLowerInvariant()}
function Find-DateRow($Sheet,[string]$Date,[int]$First){
 $found=0
 for($row=$First;$row -le $Sheet.UsedRange.Rows.Count;$row++){
  $value=$Sheet.Cells.Item($row,1).Value2
  if($null -eq $value){continue}
  $day=$null
  if($value -is [double]){try{$day=[datetime]::FromOADate($value).ToString('yyyy-MM-dd')}catch{}}
  else{try{$day=[datetime]::ParseExact(([string]$value).Replace('.','-'),'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).ToString('yyyy-MM-dd')}catch{}}
  if($day -eq $Date){if($found){throw 'Duplicate monthly business date found in actual Excel.'};$found=$row}
 }
 if(-not $found){throw 'Selected monthly business date not found in actual Excel.'}
 return $found
}
function Grouped-Headers($Sheet){
 $headers=@{};$group=''
 for($column=1;$column -le $Sheet.UsedRange.Columns.Count;$column++){
  $top=$Sheet.Cells.Item(1,$column)
  $value=$top.MergeArea.Cells.Item(1,1).Value2
  if($value){$group=Normalize-Header $value}
  $field=Normalize-Header $Sheet.Cells.Item(2,$column).Value2
  if($group -and $field){$key=$group+'/'+$field;if($headers.ContainsKey($key)){throw 'Duplicate grouped header in actual Excel.'};$headers[$key]=$column}
 }
 return $headers
}
function Check-Value($Sheet,[int]$Row,[int]$Column,$Wanted,$Context){
 if(-not $Column){throw ($Context+' column missing in actual Excel.')}
 $value=$Sheet.Cells.Item($Row,$Column).Value2
 Assert-Number $value $Wanted $Context
 return $value
}
function Capture-OwnedWindow($Excel,[string]$Path){
 Start-Sleep -Milliseconds 300
 $rect=[CashRecMonthlyExcelWindow+Rect]::new()
 if(-not[CashRecMonthlyExcelWindow]::GetWindowRect([IntPtr]$Excel.Hwnd,[ref]$rect)){throw 'Owned Excel window rectangle unavailable.'}
 $bitmap=[Drawing.Bitmap]::new($rect.Right-$rect.Left,$rect.Bottom-$rect.Top)
 $graphics=[Drawing.Graphics]::FromImage($bitmap)
 try{
  $dc=$graphics.GetHdc()
  try{$captured=[CashRecMonthlyExcelWindow]::PrintWindow([IntPtr]$Excel.Hwnd,$dc,2)}finally{$graphics.ReleaseHdc($dc)}
  if(-not $captured){throw 'Owned Excel capture failed; desktop capture will not be substituted.'}
  $stream=[IO.MemoryStream]::new()
  try{$bitmap.Save($stream,[Drawing.Imaging.ImageFormat]::Png);[IO.File]::WriteAllBytes($Path,$stream.ToArray())}finally{$stream.Dispose()}
 }finally{$graphics.Dispose();$bitmap.Dispose()}
}
function Cell-State($Sheet,[int]$Row,[int]$Column){
 $cell=$Sheet.Cells.Item($Row,$Column)
 return @{Value=$cell.Value2;HasFormula=[bool]$cell.HasFormula;NumberFormat=[string]$cell.NumberFormat;ValueType=if($null -eq $cell.Value2){'Null'}else{$cell.Value2.GetType().FullName}}
}
$Workbook=(Resolve-Path -LiteralPath $Workbook).Path
$oracleData=Get-Content -LiteralPath $Oracle -Raw -Encoding UTF8|ConvertFrom-Json
if(-not $oracleData.synthetic_only){throw 'Only an independently specified synthetic oracle may be used.'}
New-Item -ItemType Directory -Force -Path $EvidenceDirectory|Out-Null
$EvidenceDirectory=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
$sourceHash=(Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$excel=$null;$book=$null;$sheet=$null;$owned=$false
$result=[ordered]@{Entry='Actual owned installed Excel application';Mode=$Mode;Workbook=$Workbook;Oracle=$Oracle;Passed=$false;ExistingExcelPids=$existing}
try{
 $excel=New-Object -ComObject Excel.Application
 [uint32]$createdPid=0
 [void][CashRecMonthlyExcelWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$createdPid)
 $owned=$existing -notcontains [int]$createdPid
 $result.ExcelPid=$createdPid;$result.OwnNewInstance=$owned
 if(-not $owned){throw 'Excel did not create a new isolated process; no existing instance will be changed or closed.'}
 $excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
 $book=$excel.Workbooks.Open($Workbook,0,$false)
 $excel.CalculateFull()
 $date=[string]$oracleData.business_date
 if($Mode -eq 'ValidateBalances'){
  $checks=@()
  foreach($entity in @('TMA','TMG','Total')){
   $sheet=$book.Worksheets.Item($entity)
   $row=Find-DateRow $sheet $date 3
   $headers=Grouped-Headers $sheet
   $record=[ordered]@{Sheet=$entity;BusinessDate=$date;Row=$row;Values=@{}}
   foreach($field in @('Insto','Retail','Wholesale','PSP','PSP Pending','Settle')){
    $name=Normalize-Header $field
    $record.Values['BANK/'+$field]=Check-Value $sheet $row $headers['bank/'+$name] $oracleData.balances.$entity.$field ($entity+' BANK/'+$field)
    $record.Values['DIFF/'+$field]=Check-Value $sheet $row $headers['diff/'+$name] $oracleData.diffs.$entity.$field ($entity+' DIFF/'+$field)
   }
   $record.Values['SETTLE PNL']=Check-Value $sheet $row $headers['bank/settlepnl'] $oracleData.settle_pnl.$entity ($entity+' SETTLE PNL')
   $checks+=$record
  }
  $result.Checks=$checks
  $sheet=$book.Worksheets.Item('TMA');$row=Find-DateRow $sheet $date 3
  $sheet.Activate();$sheet.Cells.Item($row,14).Select()|Out-Null
 }else{
  $sheet=$book.Worksheets.Item('Break down');$row=Find-DateRow $sheet $date 2
  foreach($pair in @(@(6,'Diff'),@(8,'To settle pnl (TMA)'),@(9,'To Settle adj (TMG)'),@(10,'To Settle (WS)'),@(11,'To Commission'),@(16,'Check'))){[void](Check-Value $sheet $row $pair[0] $oracleData.breakdown.($pair[1]) ('Break down '+$pair[1]))}
  $formula=[string]$sheet.Cells.Item($row,16).Formula
  $wanted='=F{0}-G{0}-H{0}-I{0}-J{0}-K{0}-L{0}-M{0}-N{0}' -f $row
  if($formula.Replace('$','').Replace(' ','') -ne $wanted){throw 'Actual Excel Check is not the required relative subtraction formula.'}
  $before=@{}
  foreach($column in @(7,12,13,14,15)){
   $state=Cell-State $sheet $row $column
   if($null -ne $state.Value -or $state.HasFormula){throw 'Initial monthly manual/Error cell is not truly blank in actual Excel.'}
   $before[[string]$column]=$state
  }
  $result.Before=$before;$result.CheckFormula=$formula
  $sheet.Cells.Item($row,7).Value2=[double]-5
  $sheet.Cells.Item($row,12).Value2=[double]0
  $sheet.Cells.Item($row,13).ClearContents()|Out-Null
  $sheet.Cells.Item($row,14).Value2=[double]-2.5
  $excel.CalculateFull()
  $negativeExpected=[double]$oracleData.breakdown.Check+7.5
  [void](Check-Value $sheet $row 16 $negativeExpected 'Check after negative/zero/blank manual edits')
  $result.NegativeEdits=@{ToLP=(Cell-State $sheet $row 7);BankFee=(Cell-State $sheet $row 12);Conversion=(Cell-State $sheet $row 13);TimeDiff=(Cell-State $sheet $row 14);Check=$sheet.Cells.Item($row,16).Value2}
  $sheet.Cells.Item($row,7).Value2=[double]0;$excel.CalculateFull()
  [void](Check-Value $sheet $row 16 ([double]$oracleData.breakdown.Check+2.5) 'Check after explicit zero')
  $result.ExplicitZero=Cell-State $sheet $row 7
  $sheet.Cells.Item($row,7).ClearContents()|Out-Null;$excel.CalculateFull()
  [void](Check-Value $sheet $row 16 ([double]$oracleData.breakdown.Check+2.5) 'Check after clearing manual cell')
  $result.Cleared=Cell-State $sheet $row 7
  $sheet.Cells.Item($row,7).Value2=[double]-5;$excel.CalculateFull()
  [void](Check-Value $sheet $row 16 $negativeExpected 'Final Check')
  foreach($column in @(7,12,13,14,15)){if([string]$sheet.Cells.Item($row,$column).NumberFormat -ne $before[[string]$column].NumberFormat){throw 'Actual Excel edit changed a manual cell format.'}}
  $sheet.Activate();$sheet.Cells.Item($row,7).Select()|Out-Null
 }
 $excel.ActiveWindow.Zoom=85
 $copy=Join-Path $EvidenceDirectory ('Excel-monthly-'+$Mode+'.xlsx')
 $book.SaveCopyAs($copy);$result.EditedCopy=$copy
 $screenshot=Join-Path $EvidenceDirectory ('03-excel-monthly-'+$Mode+'.png')
 Capture-OwnedWindow $excel $screenshot
 $result.Screenshot=$screenshot;$result.ScreenshotCapture='PrintWindow on the owned Excel HWND'
 $result.SourceUnchanged=(Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash -eq $sourceHash
 if(-not $result.SourceUnchanged){throw 'Actual Excel verification modified the original generated workbook.'}
 $result.Passed=$true
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'excel-monthly-result.json'),($result|ConvertTo-Json -Depth 15),[Text.UTF8Encoding]::new($false))
 $result|ConvertTo-Json -Depth 15
}catch{
 $result.Error=$_.Exception.Message
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'excel-monthly-result.json'),($result|ConvertTo-Json -Depth 15),[Text.UTF8Encoding]::new($false))
 throw
}finally{
 if($null -ne $book){try{$book.Close($false)}catch{}}
 if($owned -and $null -ne $excel){try{$excel.Quit()}catch{}}
 foreach($item in @($sheet,$book,$excel)){if($null -ne $item){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)}catch{}}}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

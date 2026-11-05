[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][string]$Workbook,
 [Parameter(Mandatory=$true)][string]$Oracle,
 [Parameter(Mandatory=$true)][string]$EvidenceDirectory
)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
if(-not ('CashRecExcelWindow' -as [type])){
 Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CashRecExcelWindow {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
 [StructLayout(LayoutKind.Sequential)] public struct Rect {public int Left,Top,Right,Bottom;}
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd,out Rect rect);
 [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd,IntPtr hdc,uint flags);
}
'@
}
$Workbook=(Resolve-Path -LiteralPath $Workbook).Path
$oracleData=Get-Content -LiteralPath $Oracle -Raw -Encoding UTF8|ConvertFrom-Json
if(-not $oracleData.synthetic_only){throw 'The independent oracle must identify synthetic test data.'}
$case=$oracleData.manual_formula_case
New-Item -ItemType Directory -Force -Path $EvidenceDirectory|Out-Null
$EvidenceDirectory=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$excel=$null;$book=$null;$sheet=$null;$owned=$false
$result=[ordered]@{Entry='Actual installed Excel application';SourceWorkbook=$Workbook;Oracle=$Oracle;Passed=$false;ExistingExcelPids=$existing}
try{
 $excel=New-Object -ComObject Excel.Application
 [uint32]$createdPid=0
 [void][CashRecExcelWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$createdPid)
 $owned=$existing -notcontains [int]$createdPid
 $result.ExcelPid=$createdPid
 $result.OwnNewInstance=$owned
 if(-not $owned){throw 'COM did not create an isolated Excel process; existing user instance will not be modified or closed.'}
 $excel.Visible=$true
 $excel.DisplayAlerts=$false
 $excel.AskToUpdateLinks=$false
 $book=$excel.Workbooks.Open($Workbook,0,$false)
 $sheet=$book.Worksheets.Item($case.sheet)
 $headers=@{}
 for($column=1;$column -le $sheet.UsedRange.Columns.Count;$column++){
  $name=[string]$sheet.Cells.Item(1,$column).Value2
  if($name){$headers[$name]=$column}
 }
 $aliases=@('Retail Type','IB/Standard','IB Status','Category')
 $typeHeader=$aliases|Where-Object {$headers.ContainsKey($_)}|Select-Object -First 1
 if(-not $typeHeader){throw 'Equity customer type header missing'}
 $rowFound=0
 for($row=2;$row -le $sheet.UsedRange.Rows.Count;$row++){
  if([string]$sheet.Cells.Item($row,$headers.Branch).Value2 -eq $case.key.Branch -and
     [string]$sheet.Cells.Item($row,$headers[$typeHeader]).Value2 -eq $case.key.'Retail Type' -and
     [string]$sheet.Cells.Item($row,$headers.Currency).Value2 -eq $case.key.Currency){$rowFound=$row;break}
 }
 if(-not $rowFound){throw 'Synthetic formula row not found'}
 foreach($required in @('DEP','PNL','Movement')){if(-not $headers.ContainsKey($required)){throw "Header missing: $required"}}
 $before=@{DEP=$sheet.Cells.Item($rowFound,$headers.DEP).Value2;PNL=$sheet.Cells.Item($rowFound,$headers.PNL).Value2;Movement=$sheet.Cells.Item($rowFound,$headers.Movement).Value2}
 $pnlFormula=[string]$sheet.Cells.Item($rowFound,$headers.PNL).Formula
 if(-not $pnlFormula.StartsWith('=')){throw 'PNL is not an Excel formula'}
 $sheet.Activate()
 $sheet.Cells.Item($rowFound,$headers.DEP).Select()|Out-Null
 $sheet.Cells.Item($rowFound,$headers.DEP).Value2=[double]$case.edit.DEP
 $excel.CalculateFull()
 $actualMovement=[double]$sheet.Cells.Item($rowFound,$headers.Movement).Value2
 $actualPnl=[double]$sheet.Cells.Item($rowFound,$headers.PNL).Value2
 $result.Before=$before
 $result.After=@{DEP=$sheet.Cells.Item($rowFound,$headers.DEP).Value2;PNL=$actualPnl;Movement=$actualMovement}
 $result.PnlFormula=$pnlFormula
 $result.Passed=([math]::Abs($actualMovement-[double]$case.expected_after_excel_recalculation.Movement) -lt 0.00000001) -and ([math]::Abs($actualPnl-[double]$case.expected_after_excel_recalculation.PNL) -lt 0.00000001)
 $copy=Join-Path $EvidenceDirectory 'Excel-edited-synthetic-result.xlsx'
 $book.SaveCopyAs($copy)
 $result.EditedCopy=$copy
 [void][CashRecExcelWindow]::SetForegroundWindow([IntPtr]$excel.Hwnd)
 Start-Sleep -Milliseconds 300
 $rect=[CashRecExcelWindow+Rect]::new()
 [void][CashRecExcelWindow]::GetWindowRect([IntPtr]$excel.Hwnd,[ref]$rect)
 $bitmap=[Drawing.Bitmap]::new($rect.Right-$rect.Left,$rect.Bottom-$rect.Top)
 $graphics=[Drawing.Graphics]::FromImage($bitmap)
 try{
  $dc=$graphics.GetHdc()
  try{$captured=[CashRecExcelWindow]::PrintWindow([IntPtr]$excel.Hwnd,$dc,2)}finally{$graphics.ReleaseHdc($dc)}
  if(-not $captured){throw 'Could not capture the owned Excel window; no desktop screenshot will be substituted.'}
  $stream=[IO.MemoryStream]::new()
  try{$bitmap.Save($stream,[Drawing.Imaging.ImageFormat]::Png);[IO.File]::WriteAllBytes((Join-Path $EvidenceDirectory '03-excel-dep-pnl.png'),$stream.ToArray())}finally{$stream.Dispose()}
 }finally{$graphics.Dispose();$bitmap.Dispose()}
 $result.ScreenshotCapture='PrintWindow on the owned Excel window'
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'excel-result.json'),($result|ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
 if(-not $result.Passed){throw 'Actual Excel recalculation differs from independent expected result'}
 $result|ConvertTo-Json -Depth 10
}catch{
 $result.Error=$_.Exception.Message
 [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'excel-result.json'),($result|ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
 throw
}finally{
 if($null -ne $book){try{$book.Close($false)}catch{}}
 if($owned -and $null -ne $excel){try{$excel.Quit()}catch{}}
 foreach($item in @($sheet,$book,$excel)){if($null -ne $item){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)}catch{}}}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}


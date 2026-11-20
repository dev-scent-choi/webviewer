# 문서 하나를 MS Office로 열어 PDF(정답)로 내보내고, 같은 형식으로 다시 저장한다. ref_office.ps1이 문서마다 따로 띄운다
# (한 Office 프로세스로 여러 문서를 돌리면 가끔 숨은 창에서 멈췄다 — 문서마다 새로 띄우고 시간 제한을 건다).
# 띄운 Office 프로세스 번호를 -PidFile에 먼저 적어, 멈추면 부른 쪽이 그 번호로만 끈다.
param([string]$Kind, [string]$In, [string]$Pdf, [string]$Out, [string]$PidFile)
$ErrorActionPreference = 'Stop'
$proc = @{ docx = 'WINWORD'; pptx = 'POWERPNT'; xlsx = 'EXCEL' }[$Kind]
$prog = @{ docx = 'Word.Application'; pptx = 'PowerPoint.Application'; xlsx = 'Excel.Application' }[$Kind]
# 앞 문서의 PowerPoint가 끝나는 데 몇 초 걸린다 — 사라질 때까지 기다린다(사용자가 띄운 창이면 20초 뒤 멈춘다)
for ($k = 0; $Kind -eq 'pptx' -and $k -lt 40 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue); $k++) { Start-Sleep -Milliseconds 500 }
$before = @(Get-Process $proc -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
if ($Kind -eq 'pptx' -and $before.Count) { throw 'PowerPoint가 이미 켜져 있다(한 프로세스 앱이라 그 창에 붙는다) — 닫고 다시 실행' }
$app = New-Object -ComObject $prog
$mine = @(Get-Process $proc -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } | Where-Object { $before -notcontains $_ })
Set-Content $PidFile ($mine -join ',')
try {
  if ($Kind -eq 'docx') {
    $app.Visible = $false; $app.DisplayAlerts = 0
    $doc = $app.Documents.Open($In, $false, $true, $false)
    $doc.ExportAsFixedFormat($Pdf, 17)
    # 호환 모드 문서는 SaveAs2가 숨은 '형식 올리기' 창에서 멈출 수 있어 원본 그대로 둔다
    if ($doc.CompatibilityMode -ge 15) { $doc.SaveAs2($Out, 16) } else { Copy-Item $In $Out -Force }
    $doc.Close(0)
  } elseif ($Kind -eq 'pptx') {
    $app.AutomationSecurity = 3
    $pres = $app.Presentations.Open($In, -1, 0, 0); $pres.SaveAs($Pdf, 32); $pres.Close()
    $pres = $app.Presentations.Open($In, 0, 0, 0); $pres.SaveAs($Out, 24); $pres.Close()
  } else {
    # 화면과 비슷하게: 저장한 뒤 시트마다 격자선·행/열 머리 인쇄, 폭 1쪽 맞춤으로 바꿔 PDF로 내보내고 저장 안 하고 닫는다
    $app.Visible = $false; $app.DisplayAlerts = $false; $app.AutomationSecurity = 3
    $wb = $app.Workbooks.Open($In, 0, $true)
    $app.CalculateFull(); $wb.SaveAs($Out, 51)
    $app.PrintCommunication = $false  # 쪽 설정을 바꿀 때마다 프린터와 통신하면 시트마다 수십 초 걸렸다
    foreach ($ws in $wb.Worksheets) {
      if ($ws.Visible -ne -1) { continue }
      $ps = $ws.PageSetup; $ps.PrintGridlines = $true; $ps.PrintHeadings = $true; $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false
      $ps.CenterHeader = ''; $ps.CenterFooter = ''
    }
    $app.PrintCommunication = $true
    $wb.ExportAsFixedFormat(0, $Pdf, 0, $true, $true); $wb.Close($false)
  }
} finally {
  try { $app.Quit() } catch {}
  Start-Sleep -Milliseconds 500
  foreach ($i in $mine) { if (Get-Process -Id $i -ErrorAction SilentlyContinue) { Stop-Process -Id $i -Force } }
}

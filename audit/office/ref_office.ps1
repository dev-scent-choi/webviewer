# 정답 그림 만들기: raw 문서마다 ref_one.ps1을 따로 띄워 MS Office로 PDF(정답)를 만들고 같은 형식으로 다시 저장한다(실제 사용자 파일 모양).
# 문서 하나에 90초 제한 — 넘으면 그 문서가 띄운 Office 프로세스만 번호로 끄고 다음으로 넘어간다.
# 사용: powershell -File audit/office/ref_office.ps1 -Raw audit/office/corpus/raw -Corpus audit/office/corpus -RefDir docs/quality-audit/out/ref [-Only docx,pptx,xlsx]
param([string]$Raw, [string]$Corpus, [string]$RefDir, [string]$Only = 'docx,pptx,xlsx', [int]$TimeoutSec = 90)
$ErrorActionPreference = 'Stop'
$Raw = (Resolve-Path $Raw).Path; New-Item -ItemType Directory -Force $RefDir | Out-Null; $RefDir = (Resolve-Path $RefDir).Path
$Corpus = (Resolve-Path $Corpus).Path
$one = Join-Path $PSScriptRoot 'ref_one.ps1'
$pidFile = Join-Path $env:TEMP "office-ref-$PID.txt"
foreach ($kind in $Only.Split(',')) {
  New-Item -ItemType Directory -Force (Join-Path $Corpus $kind) | Out-Null
  foreach ($f in @(Get-ChildItem (Join-Path $Raw $kind) -Filter "*.$kind")) {
    $pdf = Join-Path $RefDir ($f.BaseName + '.pdf'); $out = Join-Path $Corpus "$kind\$($f.Name)"
    if ((Test-Path $pdf) -and (Test-Path $out)) { "skip    $kind $($f.Name)"; continue }  # 이미 만든 것은 건너뛴다(다시 하려면 지운다)
    Set-Content $pidFile ''
    $p = Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$one`"", '-Kind', $kind, '-In', "`"$($f.FullName)`"", '-Pdf', "`"$pdf`"", '-Out', "`"$out`"", '-PidFile', "`"$pidFile`"") -PassThru -WindowStyle Hidden
    if ($p.WaitForExit($TimeoutSec * 1000) -and (Test-Path $pdf)) { "ok      $kind $($f.Name)" }
    else {
      "TIMEOUT $kind $($f.Name)"
      foreach ($i in ((Get-Content $pidFile -ErrorAction SilentlyContinue) -split ',' | Where-Object { $_ })) { Stop-Process -Id ([int]$i) -Force -ErrorAction SilentlyContinue }
      if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
    }
  }
}
Remove-Item $pidFile -ErrorAction SilentlyContinue

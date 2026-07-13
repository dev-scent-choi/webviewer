# 대형 페이지 문서를 고배율로 반복 렌더링해 캐시 메모리가 예산 안에서 안정되는지 확인
$ErrorActionPreference = 'Stop'
$docs = (curl.exe -s http://127.0.0.1:8000/api/documents | ConvertFrom-Json).data
$pageDocs = $docs | Where-Object { $_.status -eq 'ready' -and $_.viewKind -ne 'grid' } | Select-Object -First 6

$body = '{"viewer":"load","canPrint":false,"canDownload":false,"canCopy":false}'
$tokens = @()
foreach ($d in $pageDocs) {
  $r = ($body | curl.exe -s -X POST -H "Content-Type: application/json" -d "@-" ("http://127.0.0.1:8000/api/documents/" + $d.id + "/tokens")) | ConvertFrom-Json
  if ($r.success) { $tokens += [pscustomobject]@{ token = $r.data.token; pages = [int]$d.pageCount } }
}
Write-Output ("tokens: {0}" -f $tokens.Count)

$node = Get-Process node | Sort-Object WorkingSet64 -Descending | Select-Object -First 1
Write-Output ("before: {0:n0} MB" -f ($node.WorkingSet64 / 1MB))

# 여러 문서 × 여러 배율 × 여러 페이지를 20 라운드 반복 요청
$scales = @(4, 8, 12, 16)
for ($round = 1; $round -le 20; $round++) {
  foreach ($t in $tokens) {
    $page = (Get-Random -Minimum 1 -Maximum ([Math]::Max(2, $t.pages + 1)))
    $scale = $scales | Get-Random
    curl.exe -s -o NUL ("http://127.0.0.1:8000/api/view/pages/{0}?token={1}&scale={2}" -f $page, $t.token, $scale)
  }
  if ($round % 5 -eq 0) {
    $n = Get-Process -Id $node.Id
    Write-Output ("round {0}: {1:n0} MB" -f $round, ($n.WorkingSet64 / 1MB))
  }
}
$n = Get-Process -Id $node.Id
Write-Output ("after: {0:n0} MB (peak WS)" -f ($n.PeakWorkingSet64 / 1MB))

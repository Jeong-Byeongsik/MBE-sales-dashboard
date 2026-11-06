# Windows: 내 PC의 모든 Claude Code 세션에 영업 스킬과 공통 규칙을 설치합니다.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dst = Join-Path $env:USERPROFILE '.claude'
New-Item -ItemType Directory -Force -Path (Join-Path $dst 'skills') | Out-Null
Copy-Item -Recurse -Force (Join-Path $root '.claude\skills\*') (Join-Path $dst 'skills')
Copy-Item -Force (Join-Path $root 'claude-config\sales-rules.md') (Join-Path $dst 'sales-rules.md')
$md = Join-Path $dst 'CLAUDE.md'
if (-not (Test-Path $md)) { New-Item -ItemType File -Path $md | Out-Null }
if (-not (Select-String -Path $md -Pattern '@sales-rules.md' -Quiet)) { Add-Content -Path $md -Value "`n@sales-rules.md" }
Write-Host '설치 완료: 영업 스킬과 공통 규칙(sales-rules.md)'

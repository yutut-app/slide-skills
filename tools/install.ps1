# Windows：5本のスキルを %USERPROFILE%\.claude\skills に写して入れる。
#   powershell -ExecutionPolicy Bypass -File tools\install.ps1
# 同じ名前のフォルダは**鏡写し（robocopy /MIR）**にする。古いファイルが残らない
# （実績: フォルダを上から写しただけで、参照ファイル1本が別の中身のまま残った）。
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Dest = if ($env:CLAUDE_SKILLS_DIR) { $env:CLAUDE_SKILLS_DIR } else { Join-Path $env:USERPROFILE ".claude\skills" }
$Old  = Join-Path $env:USERPROFILE ".claude\skills-old"
$py = if (Get-Command py -ErrorAction SilentlyContinue) { "py" } else { "python" }
& $py "$Root\tools\bundle_content.py" --check | Out-Null
if ($LASTEXITCODE -ne 0) { Write-Host "中身の規約の写しが古い。先に $py tools\bundle_content.py"; exit 1 }
New-Item -ItemType Directory -Force -Path $Dest | Out-Null
Get-ChildItem -Directory (Join-Path $Root "skills") | ForEach-Object {
  if (Test-Path (Join-Path $_.FullName "SKILL.md")) {
    robocopy $_.FullName (Join-Path $Dest $_.Name) /MIR /XD __pycache__ /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "写せなかった: $($_.Name)" }
    Write-Host "入れた     $($_.Name)"
  }
}
foreach ($n in "slider-craft","slider-reskin","create-slide") {
  $p = Join-Path $Dest $n
  if (Test-Path $p) { New-Item -ItemType Directory -Force -Path $Old | Out-Null; Move-Item $p $Old -Force; Write-Host "移した     $n → $Old（使わない旧スキル。中身は残る）" }
}
& $py "$Root\tools\verify_install.py" $Dest

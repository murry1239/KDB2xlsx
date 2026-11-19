# 원시자료 정리봇 바탕화면·시작 메뉴 바로가기 만들기
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$pyw = Join-Path $root ".venv\Scripts\pythonw.exe"
if (-not (Test-Path $pyw)) { Write-Host "INSTALL.bat 을 먼저 실행하세요."; exit 1 }
$shell = New-Object -ComObject WScript.Shell
$targets = @(
  [Environment]::GetFolderPath("Desktop"),
  (Join-Path ([Environment]::GetFolderPath("StartMenu")) "Programs")
)
foreach ($dir in $targets) {
  $lnk = $shell.CreateShortcut((Join-Path $dir "원시자료 정리봇.lnk"))
  $lnk.TargetPath = $pyw
  $lnk.Arguments = '"' + (Join-Path $root "app.py") + '"'
  $lnk.WorkingDirectory = $root
  $lnk.IconLocation = "$env:SystemRoot\System32\imageres.dll,-5302"
  $lnk.Description = "KAERS 원시자료 정리봇"
  $lnk.Save()
  Write-Host "바로가기 생성: $($lnk.FullName)"
}

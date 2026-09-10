[CmdletBinding()]
param(
  [string]$CredentialPath = (Join-Path $env:LOCALAPPDATA 'JD-SupplyChain\bbcc-pages-password.xml'),
  [string]$AppAssets = ''
)
$ErrorActionPreference = 'Stop'
$Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Python = Join-Path $Root '.venv\Scripts\python.exe'
if (-not $AppAssets) {
  $appsRoot = if ($env:JD_SUPPLYCHAIN_APPS_ROOT) {
    $env:JD_SUPPLYCHAIN_APPS_ROOT
  } else {
    Join-Path $env:USERPROFILE 'repos\jd-supplychain-apps'
  }
  $AppAssets = Join-Path $appsRoot 'apps\jd_free_goods_bbcc_cost_simulation\assets'
}
if (-not (Test-Path -LiteralPath $Python)) { throw "找不到公开站构建环境：$Python" }
if (-not (Test-Path -LiteralPath $AppAssets)) { throw "找不到BBCC私有应用数据层：$AppAssets" }
if (-not (Test-Path -LiteralPath $CredentialPath)) { throw "未找到BBCC页加密密码文件：$CredentialPath" }
$securePassword = Import-Clixml -LiteralPath $CredentialPath
if ($securePassword -isnot [Security.SecureString]) { throw "密码文件格式无效：$CredentialPath" }
$pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
try {
  $env:BBCC_PAGES_PASSWORD = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
  & $Python -X utf8 (Join-Path $PSScriptRoot 'build-bbcc-data.py') --app-assets $AppAssets --password-env BBCC_PAGES_PASSWORD --self-test
  if ($LASTEXITCODE -ne 0) { throw "加密BBCC数据构建失败，退出码$LASTEXITCODE" }
} finally {
  Remove-Item Env:BBCC_PAGES_PASSWORD -ErrorAction SilentlyContinue
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
}
Write-Host '加密BBCC数据已就绪；密码未写入仓库或命令参数。' -ForegroundColor Green

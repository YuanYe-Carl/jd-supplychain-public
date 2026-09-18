[CmdletBinding()]
param(
    [ValidateRange(1, 10)][int]$SnapshotCount = 3,
    [string]$SourceDirectory = '',
    [string]$AppAssets = '',
    [string]$StatePath = (Join-Path $env:LOCALAPPDATA 'JD-SupplyChain\fulfillment-pages-state.json'),
    [switch]$Force,
    [switch]$NoPush
)

$ErrorActionPreference = 'Stop'
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Builder = Join-Path $PSScriptRoot 'build-fulfillment-data.ps1'
if (-not $SourceDirectory) {
    $SourceDirectory = if ($env:JD_FULFILLMENT_SOURCE_DIR) {
        $env:JD_FULFILLMENT_SOURCE_DIR
    } else {
        Join-Path $env:USERPROFILE 'Procter and Gamble\JD PS 铁军 - 文档\17 SND\18. 代发治理\拆单或代发判断数据基础'
    }
}
if (-not $AppAssets) {
    # apps 仓库通常与本仓库平级，先找同级目录再回退到 repos。
    $appsRoot = if ($env:JD_SUPPLYCHAIN_APPS_ROOT) {
        $env:JD_SUPPLYCHAIN_APPS_ROOT
    } elseif (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $RepoRoot) 'jd-supplychain-apps')) {
        Join-Path (Split-Path -Parent $RepoRoot) 'jd-supplychain-apps'
    } else {
        Join-Path $env:USERPROFILE 'repos\jd-supplychain-apps'
    }
    $AppAssets = Join-Path $appsRoot 'apps\jd_fulfillment_decision_tool\assets'
}
$RelativeOutputs = @(
    'data/fulfillment-status.json'
    'data/fulfillment-snapshots'
)
$GitBase = @('-c', "safe.directory=$RepoRoot", '-C', $RepoRoot)

function Get-SourceFingerprint {
    $inventoryDirectory = if ($env:JD_INVENTORY_DIR) {
        $env:JD_INVENTORY_DIR
    } else {
        Join-Path $env:USERPROFILE 'Procter and Gamble\JD PS 铁军 - 文档\03 库存管理\每日库存'
    }
    $latest = Get-ChildItem -LiteralPath $inventoryDirectory -Filter '*.xlsx' -File |
        Where-Object Name -NotLike '~$*' |
        Sort-Object Name -Descending |
        Select-Object -First 1
    if (-not $latest) { throw "没有找到库存切片：$inventoryDirectory" }
    $paths = @(
        $latest.FullName
        (Join-Path $SourceDirectory '宝洁直送明细.xlsx')
        (Join-Path $SourceDirectory '11区域对应关系.xlsx')
        (Join-Path $SourceDirectory '轻货仓对应关系.xlsx')
    )
    $items = foreach ($path in $paths) {
        $item = Get-Item -LiteralPath $path
        [ordered]@{ path = $item.FullName; ticks = $item.LastWriteTimeUtc.Ticks; length = $item.Length }
    }
    return [ordered]@{ snapshot_count = $SnapshotCount; sources = @($items) }
}

function ConvertTo-StableJson {
    param([Parameter(Mandatory)]$Value)
    return $Value | ConvertTo-Json -Depth 5 -Compress
}

function Invoke-Git {
    param([Parameter(Mandatory)][string[]]$Arguments)
    & git @GitBase @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "git $($Arguments -join ' ')失败，退出码$LASTEXITCODE"
    }
}

function Invoke-EngineTest {
    param([Parameter(Mandatory)][string]$Path)

    $node = Get-Command node.exe -ErrorAction SilentlyContinue
    if ($node) {
        & $node.Source $Path
        if ($LASTEXITCODE -ne 0) {
            throw "浏览器履约引擎测试失败，退出码$LASTEXITCODE"
        }
        return
    }

    $code = Join-Path $env:ProgramFiles 'Microsoft VS Code\Code.exe'
    if (-not (Test-Path -LiteralPath $code)) {
        throw '未找到 Node.js 或 VS Code 内置 Node 运行时'
    }

    $stdoutPath = [IO.Path]::GetTempFileName()
    $stderrPath = [IO.Path]::GetTempFileName()
    $previousRunAsNode = $env:ELECTRON_RUN_AS_NODE
    try {
        $env:ELECTRON_RUN_AS_NODE = '1'
        $process = Start-Process `
            -FilePath $code `
            -ArgumentList "`"$Path`"" `
            -WorkingDirectory $RepoRoot `
            -Wait `
            -PassThru `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath
        Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue
        Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue
        if ($process.ExitCode -ne 0) {
            throw "浏览器履约引擎测试失败，退出码$($process.ExitCode)"
        }
    } finally {
        if ($null -eq $previousRunAsNode) {
            Remove-Item Env:ELECTRON_RUN_AS_NODE -ErrorAction SilentlyContinue
        } else {
            $env:ELECTRON_RUN_AS_NODE = $previousRunAsNode
        }
        Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function Write-PublicationState {
    $stateDirectory = Split-Path -Parent $StatePath
    New-Item -ItemType Directory -Path $stateDirectory -Force | Out-Null
    $fingerprint | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $StatePath -Encoding utf8
}

function Test-AllowedPath {
    param([Parameter(Mandatory)][string]$Path)
    $normalized = $Path.Replace('\', '/')
    foreach ($allowed in $RelativeOutputs) {
        if ($normalized -eq $allowed -or $normalized.StartsWith("$allowed/")) {
            return $true
        }
    }
    return $false
}

$changes = @(& git @GitBase status --porcelain)
if ($LASTEXITCODE -ne 0) { throw '无法读取Git工作区状态' }
$fingerprint = Get-SourceFingerprint
$needsBuild = $Force -or -not (Test-Path -LiteralPath (Join-Path $RepoRoot 'data\fulfillment-status.json'))
if (-not $needsBuild -and (Test-Path -LiteralPath $StatePath)) {
    try {
        $previous = Get-Content -LiteralPath $StatePath -Raw | ConvertFrom-Json
        $needsBuild = (ConvertTo-StableJson $fingerprint) -ne (ConvertTo-StableJson $previous)
    } catch {
        $needsBuild = $true
    }
} elseif (-not $needsBuild) {
    $needsBuild = $true
}

if (-not $needsBuild) {
    Write-Host '履约数据源没有变化，无需重建或推送'
    exit 0
}

if (-not $NoPush) {
    Invoke-Git @('fetch', 'origin', 'main')
    $counts = (& git @GitBase rev-list --left-right --count origin/main...HEAD) -split '\s+'
    if ($LASTEXITCODE -ne 0 -or $counts.Count -lt 2) { throw '无法比较本地main与origin/main' }
    $behind = [int]$counts[0]
    $ahead = [int]$counts[1]
    if ($behind -gt 0 -and $ahead -gt 0) { throw '本地main与origin/main已分叉，请人工处理' }
    if ($behind -gt 0) {
        if ($changes) { throw '远端有更新且本地有未提交修改，自动发布已停止' }
        Invoke-Git @('pull', '--ff-only', 'origin', 'main')
    } elseif ($ahead -gt 0) {
        Invoke-Git @('-c', 'http.postBuffer=524288000', 'push', 'origin', 'main')
    }
}

& $Builder -SnapshotCount $SnapshotCount -AppAssets $AppAssets
if ($LASTEXITCODE -ne 0) { throw "履约密文构建失败，退出码$LASTEXITCODE" }
Invoke-EngineTest (Join-Path $PSScriptRoot 'test-fulfillment-engine.js')

$addArguments = @('add', '--') + $RelativeOutputs
Invoke-Git $addArguments
& git @GitBase diff --cached --quiet -- @RelativeOutputs
if ($LASTEXITCODE -eq 0) {
    if (-not $NoPush) { Write-PublicationState }
    Write-Host '履约工具没有变化，无需提交'
    exit 0
}
if ($LASTEXITCODE -ne 1) { throw '无法检查待提交履约文件' }

if ($NoPush) {
    Write-Host '履约工具已构建并暂存；按-NoPush要求未提交或推送'
    exit 0
}

$commitArguments = @('commit', '-m', 'data: 更新订单履约加密数据', '--') + $RelativeOutputs
Invoke-Git $commitArguments
Invoke-Git @('-c', 'http.postBuffer=524288000', 'push', 'origin', 'main')
Write-PublicationState
Write-Host '订单履约分析已推送，GitHub Pages将自动部署' -ForegroundColor Green
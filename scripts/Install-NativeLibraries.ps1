param([string]$SevenZipExe)
$ErrorActionPreference = 'Stop'
if (-not $SevenZipExe) {
    $candidates = @('C:\App\7-Zip\7z.exe', "$env:ProgramFiles\7-Zip\7z.exe")
    $command = Get-Command 7z -ErrorAction SilentlyContinue
    if ($command) { $candidates += $command.Source }
    $SevenZipExe = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}
if (-not $SevenZipExe) { throw '请安装 7-Zip，或通过 -SevenZipExe 指定 7z.exe（仅用于解包安装文件）。' }
$repo = Split-Path $PSScriptRoot -Parent
$cache = Join-Path $repo 'work/native-downloads'
$destination = Join-Path $repo 'ZipExtractor/native'
New-Item -ItemType Directory -Path $cache,$destination -Force | Out-Null
$packages = @(
    @{ Name='7zip'; File='7z2603-x64.exe'; Url='https://github.com/ip7z/7zip/releases/download/26.03/7z2603-x64.exe'; Hash='0859C524B8A63551848F0C246ABDDCB1D0B7B656B0FBFE879F8D85E61A9E6EDD'; Dll='7z.dll'; Target='7z.dll' },
    @{ Name='unrar'; File='unrardll-720.exe'; Url='https://www.rarlab.com/rar/unrardll-720.exe'; Hash='B1C0D7B9771DA282709AA50FCABB751F5CCD75A58A7C54729590C65F82DB4C7F'; Dll='x64/UnRAR64.dll'; Target='UnRAR.dll' }
)
foreach ($package in $packages) {
    $download = Join-Path $cache $package.File
    if (-not (Test-Path -LiteralPath $download)) { Invoke-WebRequest -Uri $package.Url -OutFile $download }
    if ((Get-FileHash -LiteralPath $download -Algorithm SHA256).Hash -ne $package.Hash) {
        throw "下载文件校验失败：$download；请检查文件或上游版本变更。"
    }
    $unpack = Join-Path $cache $package.Name
    & $SevenZipExe x $download "-o$unpack" -y | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "安装包解包失败：$download" }
    Copy-Item -LiteralPath (Join-Path $unpack $package.Dll) -Destination (Join-Path $destination $package.Target)
    Copy-Item -LiteralPath (Join-Path $unpack 'license.txt') -Destination (Join-Path $destination ($package.Name + '-license.txt'))
}
Write-Output "已安装 7-Zip 26.03 / UnRAR 7.20 x64 到 $destination。重新 build/publish 即可复制至输出目录。"

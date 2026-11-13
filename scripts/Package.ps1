param([string]$BinaryPath)
$ErrorActionPreference='Stop'
$taskRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if(-not $BinaryPath){$BinaryPath=Join-Path $taskRoot 'source\build\LessonZoomGuard.dll'}
$taskBinary=[IO.Path]::GetFullPath($BinaryPath)
if(-not(Test-Path -LiteralPath $taskBinary)){throw '未找到 DLL，请先运行 source\Build.ps1，或通过 -BinaryPath 指定已编译的 DLL。'}
$taskDist=Join-Path $taskRoot 'dist'
[void][IO.Directory]::CreateDirectory($taskDist)
$taskStage=Join-Path $taskDist ('stage-'+[Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($taskStage)
$taskEntries=@('Install.cmd','Install.ps1','Uninstall.cmd','Uninstall.ps1','README.md','LICENSE','CHANGELOG.md','source\AddIn.cs','source\Build.ps1','source\DEVELOPMENT.md','scripts\Package.ps1')
foreach($taskEntry in $taskEntries){
    $taskDestination=Join-Path $taskStage $taskEntry
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($taskDestination))
    [IO.File]::Copy((Join-Path $taskRoot $taskEntry),$taskDestination,$true)
}
[IO.File]::Copy($taskBinary,(Join-Path $taskStage 'LessonZoomGuard.dll'),$true)
$taskManifest=Get-Content -LiteralPath (Join-Path $taskRoot 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$taskManifest.dllSha256=(Get-FileHash -LiteralPath $taskBinary -Algorithm SHA256).Hash
$taskManifest.sourceSha256=(Get-FileHash -LiteralPath (Join-Path $taskRoot 'source\AddIn.cs') -Algorithm SHA256).Hash
[IO.File]::WriteAllText((Join-Path $taskStage 'manifest.json'),($taskManifest | ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
$taskZip=Join-Path $taskDist 'PPT-ZoomGuard-0.2.0-preview-win64.zip'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$taskPending=Join-Path $taskDist ('package-'+[Guid]::NewGuid().ToString('N')+'.zip')
[IO.Compression.ZipFile]::CreateFromDirectory($taskStage,$taskPending,[IO.Compression.CompressionLevel]::Optimal,$false)
[IO.File]::Copy($taskPending,$taskZip,$true)
[IO.File]::Delete($taskPending)
$taskChecksum=Join-Path $taskDist 'SHA256SUMS.txt'
[IO.File]::WriteAllText($taskChecksum,((Get-FileHash -LiteralPath $taskZip -Algorithm SHA256).Hash.ToLowerInvariant()+'  '+[IO.Path]::GetFileName($taskZip)+[Environment]::NewLine),[Text.UTF8Encoding]::new($false))
Write-Host ('安装包：'+$taskZip)
Write-Host ('校验文件：'+$taskChecksum)

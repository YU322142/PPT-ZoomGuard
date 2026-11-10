$ErrorActionPreference='Stop'
$taskBuild=Join-Path $PSScriptRoot 'build'
[void][IO.Directory]::CreateDirectory($taskBuild)
$taskSdk=Join-Path ${env:ProgramFiles(x86)} 'Microsoft SDKs\Windows\v10.0A\bin'
$taskCandidates=@(
    (Join-Path $taskSdk 'NETFX 4.8.1 Tools\x64\TlbImp.exe'),
    (Join-Path $taskSdk 'NETFX 4.8 Tools\x64\TlbImp.exe'),
    (Join-Path $taskSdk 'NETFX 4.7 Tools\x64\TlbImp.exe')
)
$taskImporter=$taskCandidates | Where-Object {Test-Path -LiteralPath $_} | Select-Object -First 1
if(-not $taskImporter){throw '需要 Windows SDK 的 .NET Framework 工具 TlbImp.exe。'}
$taskTypeLib=Join-Path $env:CommonProgramFiles 'DESIGNER\MSADDNDR.OLB'
if(-not(Test-Path -LiteralPath $taskTypeLib)){throw '未找到 Office Add-in Designer 类型库 MSADDNDR.OLB。'}
$taskInterop=Join-Path $taskBuild 'ExtensibilityArray.dll'
& $taskImporter $taskTypeLib ('/out:'+$taskInterop) /namespace:Extensibility /sysarray /silent
if($LASTEXITCODE -ne 0){throw '类型库转换失败。'}
$taskCompiler=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$taskOutput=Join-Path $taskBuild 'LessonZoomGuard.dll'
& $taskCompiler /nologo /target:library /platform:anycpu /optimize+ /r:System.Windows.Forms.dll /r:System.Drawing.dll ('/link:'+$taskInterop) ('/out:'+$taskOutput) (Join-Path $PSScriptRoot 'AddIn.cs')
if($LASTEXITCODE -ne 0){throw '编译失败。'}
Write-Host ('编译完成：'+$taskOutput)
Get-FileHash -LiteralPath $taskOutput -Algorithm SHA256

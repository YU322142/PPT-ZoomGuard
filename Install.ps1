param([switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
$taskProductId = '{F1C4E46D-337F-4504-9F69-81A47FE21183}'
$taskProgId = 'LessonZoomGuard.Connect'
$taskProductRoot = Join-Path $env:LOCALAPPDATA 'LessonZoomGuard'
$taskInstallRoot = [IO.Path]::GetFullPath((Join-Path $taskProductRoot '0.1.0'))
$taskAllowedRoot = [IO.Path]::GetFullPath($taskProductRoot).TrimEnd('\') + '\'
if (-not $taskInstallRoot.StartsWith($taskAllowedRoot,[StringComparison]::OrdinalIgnoreCase)) { throw '安装目录校验失败。' }
$taskDll = Join-Path $PSScriptRoot 'LessonZoomGuard.dll'
$taskManifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if ((Get-FileHash -LiteralPath $taskDll -Algorithm SHA256).Hash -ne $taskManifest.dllSha256) { throw '安装文件校验失败，请重新解压原始安装包。' }
$taskOffice = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if (-not $taskOffice -or $taskOffice.Platform -ne 'x64') { throw '当前测试版需要 64 位桌面版 PowerPoint。' }
Write-Host ('检测到 Office：' + $taskOffice.VersionToReport + '，64 位。')
if ($CheckOnly) { Write-Host '文件与安装环境检查通过。尚未安装。'; exit 0 }
$taskSession = [Diagnostics.Process]::GetCurrentProcess().SessionId
if (Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $taskSession }) { throw '请先保存课件并完全关闭 PowerPoint，然后重新运行安装。' }
$taskHive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::Registry64)
$taskAddinPath = 'Software\Microsoft\Office\PowerPoint\Addins\' + $taskProgId
$taskExisting = $taskHive.OpenSubKey($taskAddinPath)
if ($taskExisting) {
    $taskExistingOwner = $taskExisting.GetValue('ProductId')
    $taskExisting.Dispose()
    if ($taskExistingOwner -ne $taskProductId) { $taskHive.Dispose(); throw '发现同名的旧注册，请先卸载旧测试版，再安装。' }
}
try {
    [void][IO.Directory]::CreateDirectory($taskInstallRoot)
    $taskInstalledDll = Join-Path $taskInstallRoot 'LessonZoomGuard.dll'
    [IO.File]::Copy($taskDll,$taskInstalledDll,$true)
    Unblock-File -LiteralPath $taskInstalledDll -ErrorAction SilentlyContinue
    $taskAssembly = [Reflection.AssemblyName]::GetAssemblyName($taskInstalledDll).FullName
    $taskClassPath = 'Software\Classes\CLSID\' + $taskProductId
    $taskKey = $taskHive.CreateSubKey($taskClassPath)
    $taskKey.SetValue('','Lesson Zoom Guard'); $taskKey.Dispose()
    $taskKey = $taskHive.CreateSubKey($taskClassPath + '\InprocServer32')
    $taskKey.SetValue('',(Join-Path ([Environment]::GetFolderPath('System')) 'mscoree.dll'))
    $taskKey.SetValue('ThreadingModel','Both')
    $taskKey.SetValue('Class','LessonZoomGuard.Connect')
    $taskKey.SetValue('Assembly',$taskAssembly)
    $taskKey.SetValue('RuntimeVersion','v4.0.30319')
    $taskKey.SetValue('CodeBase',([Uri]$taskInstalledDll).AbsoluteUri)
    $taskKey.Dispose()
    $taskKey = $taskHive.CreateSubKey($taskClassPath + '\ProgId')
    $taskKey.SetValue('',$taskProgId); $taskKey.Dispose()
    $taskKey = $taskHive.CreateSubKey('Software\Classes\' + $taskProgId + '\CLSID')
    $taskKey.SetValue('',$taskProductId); $taskKey.Dispose()
    $taskKey = $taskHive.CreateSubKey($taskAddinPath)
    $taskKey.SetValue('FriendlyName','课程缩放防误触（实验版）')
    $taskKey.SetValue('Description','正常放映时过滤双指操作，使用原有放大镜时临时放开。')
    $taskKey.SetValue('ProductId',$taskProductId)
    $taskKey.SetValue('InstallRoot',$taskInstallRoot)
    $taskKey.SetValue('Version','0.1.0')
    $taskKey.SetValue('CommandLineSafe',1,[Microsoft.Win32.RegistryValueKind]::DWord)
    $taskKey.SetValue('LoadBehavior',3,[Microsoft.Win32.RegistryValueKind]::DWord)
    $taskKey.Dispose()
    Write-Host '安装完成。打开 PowerPoint，先用测试课件进行验证。' -ForegroundColor Green
    Write-Host '批注工具需要在教室大屏上另行实测。'
} finally { $taskHive.Dispose() }

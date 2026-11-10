$ErrorActionPreference = 'Stop'
$taskProductId = '{F1C4E46D-337F-4504-9F69-81A47FE21183}'
$taskProgId = 'LessonZoomGuard.Connect'
$taskSession = [Diagnostics.Process]::GetCurrentProcess().SessionId
if (Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $taskSession }) { throw '请先保存课件并完全关闭 PowerPoint，然后重新运行卸载。' }
$taskHive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::Registry64)
$taskAddinPath = 'Software\Microsoft\Office\PowerPoint\Addins\' + $taskProgId
$taskKey = $taskHive.OpenSubKey($taskAddinPath)
if (-not $taskKey) { $taskHive.Dispose(); Write-Host '未发现已安装的课程缩放防误触。'; exit 0 }
$taskOwner = $taskKey.GetValue('ProductId')
$taskInstallRoot = [string]$taskKey.GetValue('InstallRoot')
$taskKey.Dispose()
if ($taskOwner -ne $taskProductId) { $taskHive.Dispose(); throw '注册信息不属于此安装包，未做更改。' }
try {
    $taskPaths = @(
        $taskAddinPath,
        ('Software\Classes\CLSID\' + $taskProductId),
        ('Software\Classes\' + $taskProgId)
    )
    foreach ($taskPath in $taskPaths) { $taskHive.DeleteSubKeyTree($taskPath,$false) }
} finally { $taskHive.Dispose() }
$taskAllowedRoot = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'LessonZoomGuard')).TrimEnd('\') + '\'
if ($taskInstallRoot) {
    $taskResolved = [IO.Path]::GetFullPath($taskInstallRoot)
    if ($taskResolved.StartsWith($taskAllowedRoot,[StringComparison]::OrdinalIgnoreCase)) {
        $taskDll = Join-Path $taskResolved 'LessonZoomGuard.dll'
        if (Test-Path -LiteralPath $taskDll) { Remove-Item -LiteralPath $taskDll -Force }
        if ([IO.Directory]::Exists($taskResolved) -and [IO.Directory]::GetFileSystemEntries($taskResolved).Length -eq 0) { [IO.Directory]::Delete($taskResolved,$false) }
    }
}
Write-Host '卸载完成。PowerPoint 将恢复原有的触摸行为。' -ForegroundColor Green

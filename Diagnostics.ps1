$ErrorActionPreference='Stop'
$taskReport=[Collections.Generic.List[string]]::new()
$taskReport.Add('PPT-ZoomGuard 诊断（仅查询，不安装或启用组件）')
$taskReport.Add('时间：'+[DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))
$taskManifest=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'manifest.json') -Raw -Encoding UTF8|ConvertFrom-Json
$taskReport.Add('安装包版本：'+$taskManifest.version)
$taskRegistry=Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\LessonZoomGuard.Connect' -ErrorAction SilentlyContinue
if($taskRegistry){
    $taskReport.Add('注册版本：'+$taskRegistry.Version)
    $taskReport.Add('注册加载状态：'+$taskRegistry.LoadBehavior)
    $taskAllowed=[IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'LessonZoomGuard')).TrimEnd('\')+'\'
    if($taskRegistry.ProductId -eq '{F1C4E46D-337F-4504-9F69-81A47FE21183}' -and $taskRegistry.InstallRoot){
        $taskInstalled=[IO.Path]::GetFullPath([string]$taskRegistry.InstallRoot)
        if($taskInstalled.StartsWith($taskAllowed,[StringComparison]::OrdinalIgnoreCase)){
            $taskInstalledDll=Join-Path $taskInstalled 'LessonZoomGuard.dll'
            if(Test-Path -LiteralPath $taskInstalledDll){
                $taskHash=(Get-FileHash -LiteralPath $taskInstalledDll -Algorithm SHA256).Hash
                $taskReport.Add('已安装 DLL SHA256：'+$taskHash)
                $taskReport.Add('与本包 DLL 一致：'+($taskHash -eq $taskManifest.dllSha256))
            }
        }
    }
}else{$taskReport.Add('未发现当前用户的组件安装注册。')}
$taskApp=$null;$taskAddin=$null;$taskDiagnostics=$null
try{
    $taskApp=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    $taskAddin=$taskApp.COMAddIns.Item('LessonZoomGuard.Connect')
    $taskReport.Add('当前 PowerPoint 实例中已连接：'+$taskAddin.Connect)
    if($taskAddin.Connect){
        $taskDiagnostics=$taskAddin.Object
        if($taskDiagnostics){$taskReport.Add([string]$taskDiagnostics.GetStatus())}
        else{$taskReport.Add('组件没有返回运行状态。')}
    }
}catch{$taskReport.Add('无法读取当前组件：'+$_.Exception.Message)}
finally{
    foreach($taskCom in @($taskDiagnostics,$taskAddin,$taskApp)){
        if($taskCom -and [Runtime.InteropServices.Marshal]::IsComObject($taskCom)){
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($taskCom)
        }
    }
}
$taskLog=Join-Path $env:LOCALAPPDATA 'LessonZoomGuard\guard.log'
if(Test-Path -LiteralPath $taskLog){
    $taskReport.Add('最近日志（不包含课件文字或截图）：')
    foreach($taskLine in (Get-Content -LiteralPath $taskLog -Tail 40 -Encoding UTF8)){$taskReport.Add($taskLine)}
}
$taskReportText=$taskReport -join [Environment]::NewLine
Write-Output $taskReportText
try{
    $taskReportFile=Join-Path $PSScriptRoot 'ZoomGuard-Diagnostics.txt'
    [IO.File]::WriteAllText($taskReportFile,$taskReportText,[Text.UTF8Encoding]::new($true))
    Write-Output ('诊断已保存到：'+$taskReportFile)
}catch{Write-Output '无法在此文件夹保存诊断，请复制窗口中的文字。'}

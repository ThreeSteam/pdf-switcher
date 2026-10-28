param([Parameter(Mandatory=$true)][string]$JobFile)
$ErrorActionPreference = 'Stop'
$word = $doc = $wordProcess = $null
$tempOutput = $null
$result = @{Success=$false; Skipped=$false; Error=''; Output=''; Mode=''}
function Phase([string]$Text) {
    @{Phase=$Text; At=[DateTimeOffset]::Now.ToString('o')} | ConvertTo-Json | Set-Content -LiteralPath ($JobFile+'.phase') -Encoding UTF8
}
try {
    $job = Get-Content -LiteralPath $JobFile -Raw -Encoding UTF8 | ConvertFrom-Json
    $source = [IO.Path]::GetFullPath([string]$job.Source)
    if (!(Test-Path -LiteralPath $source -PathType Leaf) -or [IO.Path]::GetExtension($source).ToLowerInvariant() -notin @('.doc','.docx','.docm')) { throw '文件不存在或不是支持的 Word 格式。' }
    $target = [IO.Path]::ChangeExtension($source,'.pdf')
    $result.Output = $target
    if ([IO.File]::Exists($target) -and !$job.Overwrite) {
        $result.Skipped=$true
    } else {
        $tempOutput = [IO.Path]::Combine([IO.Path]::GetDirectoryName($target),('.pdf-switcher-'+[guid]::NewGuid().ToString('N')+'.pdf'))
        # Verify destination permissions before starting Word.
        $probe=[IO.File]::Create($tempOutput); $probe.Dispose()
        Phase '启动 Word'
        . "$PSScriptRoot\WordNative.ps1"
        $wordExe=$null
        foreach ($key in @('Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE','Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE','Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE')) {
            $entry=Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
            if ($entry -and $entry.'(default)') { $wordExe=$entry.'(default)'; break }
        }
        if (!$wordExe -or !(Test-Path -LiteralPath $wordExe)) { throw '未找到 Microsoft Word 桌面版，请先安装并完成首次启动。' }
        $wordProcess=Start-Process -FilePath $wordExe -ArgumentList '/x','/a','/w' -WindowStyle Hidden -PassThru
        @{Id=$wordProcess.Id; Ticks=$wordProcess.StartTime.Ticks} | ConvertTo-Json | Set-Content -LiteralPath ($JobFile+'.owner') -Encoding UTF8
        for ($i=0; $i -lt 80; $i++) {
            if ($wordProcess.HasExited) { throw 'Word 提前退出，请检查安装和激活状态。' }
            $handle=[WordNativeObject]::FindMainWindow($wordProcess.Id)
            if ($handle -ne [IntPtr]::Zero) {
                try { $native=[WordNativeObject]::FromMainWindow($handle); if ($native) {$word=$native.Application; break} } catch {}
            }
            Start-Sleep -Milliseconds 500
        }
        if (!$word) {throw 'Word 启动超时。'}
        $word.Visible=$false; $word.DisplayAlerts=0; $word.AutomationSecurity=3
        $oldLinks=$word.Options.UpdateLinksAtOpen; $word.Options.UpdateLinksAtOpen=$false
        Phase '读取文档'
        # A deliberately incorrect password prevents a password prompt from blocking automation.
        $doc=$word.Documents.Open([string]$source,$false,$true,$false,'PDF-Switcher-no-password')
        $view=$doc.ActiveWindow.View
        $markup=[bool]$view.ShowRevisionsAndComments -and (([int]$doc.Revisions.Count -gt 0 -and ([bool]$view.ShowInsertionsAndDeletions -or [bool]$view.ShowFormatChanges)) -or ([int]$doc.Comments.Count -gt 0 -and [bool]$view.ShowComments))
        if ($job.Review -eq 'clean') {$markup=$false}
        if ($job.Review -eq 'markup') {
            $markup=$true; $view.ShowRevisionsAndComments=$true
            $view.ShowInsertionsAndDeletions=$true; $view.ShowFormatChanges=$true; $view.ShowComments=$true
        }
        $doc.PrintRevisions=$markup
        $item=[int]0
        if ($markup) {$item=[int]7; $result.Mode='含修订标记'} else {$result.Mode='清洁版'}
        Phase ('导出 PDF · '+$result.Mode)
        $doc.ExportAsFixedFormat([string]$tempOutput,[int]17,$false,[int]0,[int]0,[int]1,[int]1,$item)
        $stream=[IO.File]::OpenRead($tempOutput)
        try {$header=New-Object byte[] 5; [void]$stream.Read($header,0,5); if ([Text.Encoding]::ASCII.GetString($header) -ne '%PDF-') {throw 'Word 未生成有效的 PDF。'}} finally {$stream.Dispose()}
        Phase '保存结果'
        if ([IO.File]::Exists($target)) {
            if (!$job.Overwrite) {$result.Skipped=$true}
            else { [IO.File]::Replace($tempOutput,$target,[NullString]::Value); $tempOutput=$null; $result.Success=$true }
        } else { [IO.File]::Move($tempOutput,$target); $tempOutput=$null; $result.Success=$true }
    }
} catch { $result.Error=$_.Exception.Message }
finally {
    if ($doc) {try {$doc.Close(0)} catch {}; try {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)} catch {}}
    if ($word) {
        try {if ($null -ne $oldLinks) {$word.Options.UpdateLinksAtOpen=$oldLinks}} catch {}
        try {$word.NormalTemplate.Saved=$true; $word.Quit(0)} catch {}
        try {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)} catch {}
    }
    if ($wordProcess) {try {if (!$wordProcess.WaitForExit(3000)) {$wordProcess.Kill()}} catch {}; $wordProcess.Dispose()}
    if ($tempOutput -and [IO.File]::Exists($tempOutput)) {try {[IO.File]::Delete($tempOutput)} catch {}}
    $result | ConvertTo-Json | Set-Content -LiteralPath ($JobFile+'.result') -Encoding UTF8
}

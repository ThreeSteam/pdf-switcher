$ErrorActionPreference='Stop'
$framework=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$src=Join-Path $PSScriptRoot 'src'
$out=Join-Path $PSScriptRoot 'portable'
[void][IO.Directory]::CreateDirectory($out)
# Windows PowerShell reads scripts as ANSI without a BOM; normalize the embedded workers.
foreach($name in @('ConvertWorker.ps1','WordNative.ps1')) {
    $p=Join-Path $src $name
    [IO.File]::WriteAllText($p,[IO.File]::ReadAllText($p,[Text.Encoding]::UTF8),[Text.UTF8Encoding]::new($true))
}
& "$framework\csc.exe" /nologo /target:winexe /platform:anycpu "/win32icon:$src\App.ico" "/out:$out\PDF Switcher.exe" /reference:System.Web.Extensions.dll /reference:System.Xaml.dll "/reference:$framework\WPF\PresentationFramework.dll" "/reference:$framework\WPF\PresentationCore.dll" "/reference:$framework\WPF\WindowsBase.dll" "/resource:$src\MainWindow.xaml,MainWindow.xaml" "/resource:$src\Theme.xaml,Theme.xaml" "/resource:$src\ConvertWorker.ps1,ConvertWorker.ps1" "/resource:$src\WordNative.ps1,WordNative.ps1" "$src\App.cs"
if($LASTEXITCODE -ne 0){throw 'Build failed'}
Write-Output "Built: $out\PDF Switcher.exe"

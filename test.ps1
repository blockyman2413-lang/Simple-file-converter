$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
& .\build.ps1
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /target:exe /out:"$PSScriptRoot\bin\Tests.exe" /reference:"$PSScriptRoot\bin\RianFile.exe" /reference:Microsoft.CSharp.dll /reference:System.Core.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll "$PSScriptRoot\Tests.cs" "$PSScriptRoot\UiTests.cs"
if ($LASTEXITCODE -ne 0) { throw 'Test build failed.' }
& "$PSScriptRoot\bin\Tests.exe" | Tee-Object -FilePath "$PSScriptRoot\test-log.txt"
if ($LASTEXITCODE -ne 0) { throw 'Tests failed. See test-log.txt.' }
& $compiler /nologo /target:exe /out:"$PSScriptRoot\bin\SlowEngine.exe" "$PSScriptRoot\SlowEngine.cs"
if ($LASTEXITCODE -ne 0) { throw 'Timeout fixture build failed.' }
& $compiler /nologo /target:exe /out:"$PSScriptRoot\bin\FailureTests.exe" /reference:"$PSScriptRoot\bin\RianFile.exe" "$PSScriptRoot\FailureTests.cs"
if ($LASTEXITCODE -ne 0) { throw 'Failure test build failed.' }
& "$PSScriptRoot\bin\FailureTests.exe" | Tee-Object -FilePath "$PSScriptRoot\test-log.txt" -Append
if ($LASTEXITCODE -ne 0) { throw 'Failure tests failed. See test-log.txt.' }

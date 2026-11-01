# Tray lifecycle regression

Run this Windows-only console test after building the application in Release mode.
It requires .NET Framework 4.8.1 and a .NET SDK that supports the test project.
From the repository root, in PowerShell:

```powershell
$appOutput = (Resolve-Path .\Build\Release).Path
dotnet build .\tests\TrayLifecycle\TrayLifecycle.csproj -c Release "-p:AppBuildDirectory=$appOutput"
if ($LASTEXITCODE -ne 0) { throw 'Test build failed' }
$testOutput = '.\tests\TrayLifecycle\bin\Release\net481'
Get-ChildItem -LiteralPath $appOutput -File |
    Where-Object { $_.Extension -in @('.dll', '.exe') -or $_.Name -eq 'PowerSwitcher.TrayApp.exe.config' } |
    Copy-Item -Destination $testOutput -Force
& "$testOutput\TrayLifecycle.exe"
if ($LASTEXITCODE -ne 0) { throw 'Tray lifecycle regression failed' }
```

If the application was built elsewhere, set `$appOutput` to that directory instead.
Run in an interactive Windows desktop session: the test briefly registers real tray
icons, then removes them in its cleanup path. It does not start the application's
normal startup handler, acquire its production mutex, read or write user settings,
register global shortcuts, or switch real power plans. Power management and
configuration use test implementations.

The test sends `TaskbarCreated` only to its own main window. It checks that 12
notifications leave one registered icon and one active tray power subscription,
then checks background theme refresh, a refresh queued before closing, direct
window closure, and application exit. It does not restart Explorer or simulate
Windows logoff or shutdown.

On the original code, the first notification fails with `registered=2`.

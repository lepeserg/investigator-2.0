# Entry point for Windows PowerShell 5.1; Python is not needed to start.
[CmdletBinding()]
param(
    # Not named $Profile: that would shadow the automatic $PROFILE variable.
    # The alias keeps the documented `-Profile <name>` invocation working.
    [Alias('Profile')]
    [ValidateSet('menu', 'base', 'documents', 'ocr', 'audio', 'all')]
    [string]$SetupProfile = 'menu',
    [switch]$CheckOnly
)
$ErrorActionPreference = 'Stop'

function Find-Python312 {
    $candidates = @((Join-Path $PSScriptRoot '.venv\Scripts\python.exe'))
    if ($env:LOCALAPPDATA) { $candidates += Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe' }
    if ($env:ProgramFiles) { $candidates += Join-Path $env:ProgramFiles 'Python312\python.exe' }
    $pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($pythonCommand -and $pythonCommand.Source -notlike '*\WindowsApps\*') { $candidates += $pythonCommand.Source }
    foreach ($candidate in $candidates | Select-Object -Unique) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            try {
                & $candidate -c 'import sys; sys.exit(0 if sys.version_info[:2] == (3,12) and sys.maxsize > 2**32 else 1)' 2>$null | Out-Null
                if ($LASTEXITCODE -eq 0) { return $candidate }
            } catch { continue }
        }
    }
    $launcher = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($launcher) {
        $savedPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            $found = & $launcher.Source -3.12 -c 'import sys; sys.exit(1) if sys.maxsize <= 2**32 else print(sys.executable)' 2>$null
            if ($LASTEXITCODE -eq 0 -and $found -and (Test-Path -LiteralPath ([string]$found) -PathType Leaf)) { return [string]$found }
        } catch {
            # A Store alias may exist even when its application is unavailable.
        } finally { $ErrorActionPreference = $savedPreference }
    }
    return $null
}

try {
    if ($env:OS -ne 'Windows_NT') { throw 'Windows is required.' }
    if ($SetupProfile -eq 'menu') {
        Write-Host 'Investigator 2.0 - select components'
        Write-Host '1. Base (Python + document libraries)'
        Write-Host '2. Documents (base + LibreOffice if Word is absent)'
        Write-Host '3. Scans / OCR (base + Tesseract + Russian language)'
        Write-Host '4. Audio (base + FFmpeg + WhisperX; large download)'
        Write-Host '5. All components'
        Write-Host '0. Exit'
        $choice = Read-Host 'Select 0-5'
        if ($choice -eq '0') { exit 0 }
        $choices = @{ '1'='base'; '2'='documents'; '3'='ocr'; '4'='audio'; '5'='all' }
        if (-not $choices.ContainsKey($choice)) { throw 'Invalid selection.' }
        $SetupProfile = $choices[$choice]
    }
    $taskPython = Find-Python312
    if (-not $taskPython) {
        if ($CheckOnly) { throw 'Python 3.12 x64 is missing. Run setup without -CheckOnly.' }
        $taskWinget = Get-Command winget.exe -ErrorAction SilentlyContinue
        if (-not $taskWinget) { throw 'Install/update App Installer from https://aka.ms/getwinget and retry.' }
        Write-Host 'Python 3.12 x64 is required. Package: Python.Python.3.12 (WinGet).'
        if ((Read-Host 'Install Python for this user? [y/N]') -ne 'y') { exit 2 }
        & $taskWinget.Source install --id Python.Python.3.12 --exact --source winget --architecture x64 --scope user
        if ($LASTEXITCODE -ne 0) { throw "Python installer failed: $LASTEXITCODE" }
        $env:PATH = [Environment]::GetEnvironmentVariable('Path', 'User') + ';' + [Environment]::GetEnvironmentVariable('Path', 'Machine')
        $taskPython = Find-Python312
        if (-not $taskPython) { throw 'Python was not detected. Restart this window and retry.' }
    }
    $env:PYTHONUTF8 = '1'
    $env:PYTHONIOENCODING = 'utf-8'
    $taskArgs = @((Join-Path $PSScriptRoot 'setup_windows.py'), '--profile', $SetupProfile)
    if ($CheckOnly) { $taskArgs += '--check' }
    & $taskPython @taskArgs
    exit $LASTEXITCODE
} catch {
    Write-Host ("Setup incomplete: " + $_.Exception.Message) -ForegroundColor Red
    exit 1
}

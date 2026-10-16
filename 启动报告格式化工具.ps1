$ErrorActionPreference = 'Stop'

$appRoot = Split-Path -Parent $PSCommandPath
$url = 'http://127.0.0.1:8501'
$port = 8501

function Test-LocalPort {
    param([int] $Port)

    $client = [System.Net.Sockets.TcpClient]::new()
    try {
        $connection = $client.BeginConnect('127.0.0.1', $Port, $null, $null)
        if (-not $connection.AsyncWaitHandle.WaitOne(500)) { return $false }
        $client.EndConnect($connection)
        return $true
    }
    catch {
        return $false
    }
    finally {
        $client.Dispose()
    }
}

try {
    if (-not (Test-LocalPort -Port $port)) {
        $venvPython = Join-Path $appRoot '.venv\Scripts\python.exe'
        $python = if (Test-Path $venvPython) {
            $venvPython
        }
        else {
            (Get-Command python -ErrorAction Stop).Source
        }

        & $python -c 'import streamlit' 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw 'Streamlit is missing. In the project folder, run: python -m pip install -r requirements.txt'
        }

        $appPath = Join-Path $appRoot 'app.py'
        $arguments = @(
            '-m', 'streamlit', 'run', ('"{0}"' -f $appPath),
            '--server.address', '127.0.0.1',
            '--server.port', $port,
            '--server.headless', 'true'
        )
        $logDirectory = Join-Path $appRoot 'logs'
        New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        $logTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $stdoutLog = Join-Path $logDirectory "streamlit-$logTimestamp.out.log"
        $stderrLog = Join-Path $logDirectory "streamlit-$logTimestamp.err.log"
        Start-Process -FilePath $python -ArgumentList $arguments -WorkingDirectory $appRoot -WindowStyle Hidden `
            -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog

        $deadline = (Get-Date).AddSeconds(45)
        do {
            Start-Sleep -Milliseconds 500
            if (Test-LocalPort -Port $port) { break }
        } while ((Get-Date) -lt $deadline)

        if (-not (Test-LocalPort -Port $port)) {
            throw 'The service did not start within 45 seconds. Check the Python and Streamlit installation.'
        }
    }

    Start-Process $url
}
catch {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show($_.Exception.Message, 'Report Formatter failed to start', 'OK', 'Error') | Out-Null
    exit 1
}

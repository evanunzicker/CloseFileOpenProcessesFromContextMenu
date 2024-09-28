param (
    [string]$filePath
)

Add-Type -AssemblyName System.Windows.Forms

$result = [System.Windows.Forms.MessageBox]::Show(
    "Do you want to continue?",
    "Confirmation",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::None
)

if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    Write-Host "Continuing the script..."
} else {
    Write-Host "Script canceled."
    [System.Windows.Forms.MessageBox]::Show("The script has been aborted", "Aborted", [System.Windows.Forms.MessageBoxButtons]::Close, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

Write-Host "File Path Received: '$filePath'"

function Get-OpenFileProcessesUsingHandle {
    param (
        [string]$file
    )

    Write-Host "Checking for processes using handle.exe: '$file'"
    $fileProcesses = @()
#Replace 'PATH' with actual path to handle.exe
    $handleOutput = & "C:\PATH\handle.exe" $file -accepteula 2>&1

    foreach ($line in $handleOutput) {
        if ($line -match '^(?<pid>\d+)\s+(?<processName>.+?)\s+[\[\(]') {
            $fileProcesses += [PSCustomObject]@{
                ProcessName = $matches['processName']
                ProcessId = [int]$matches['pid']
            }
            Write-Host "Detected process $($matches['processName']) with PID $($matches['pid'])"
        }
    }

    return $fileProcesses
}

function Get-OpenFileProcessesUsingWMI {
    param (
        [string]$file
    )

    Write-Host "Checking for processes using WMI: '$file'"
    $fileProcesses = @()

    $processes = Get-WmiObject Win32_Process

    foreach ($process in $processes) {
        try {
            $processId = $process.ProcessId
            $processHandle = Get-Process -Id $processId -ErrorAction Stop
            
            foreach ($module in $processHandle.Modules) {
                if ($module.FileName -eq $file) {
                    $fileProcesses += [PSCustomObject]@{
                        ProcessName = $process.Name
                        ProcessId = $processId
                    }
                    Write-Host "Detected module $($module.FileName) in process $($process.Name) with PID $processId"
                }
            }
        } catch {
            Write-Host "Could not check process with PID $processId $($_.Exception.Message)"
        }
    }

    return $fileProcesses
}

$fileProcessesFromHandle = Get-OpenFileProcessesUsingHandle -file $filePath
$fileProcessesFromWMI = Get-OpenFileProcessesUsingWMI -file $filePath

$fileProcesses = $fileProcessesFromHandle + $fileProcessesFromWMI
$fileProcesses = $fileProcesses | Select-Object -Unique

if ($fileProcesses.Count -eq 0) {
    Write-Host "No processes are holding the file open."
} else {
    Write-Host "The following processes are holding the file open:"
    $fileProcesses | ForEach-Object {
        Write-Host "Process: $($_.ProcessName), PID: $($_.ProcessId)"
    }

    $userInput = 'Y'
    
    if ($userInput -eq 'Y') {
        foreach ($process in $fileProcesses) {
            try {
                Stop-Process -Id $process.ProcessId -ErrorAction Stop
                Write-Host "Closed process $($process.ProcessName) with PID $($process.ProcessId)"
            } catch {
                Write-Host "Failed to close process $($process.ProcessName) with PID $($process.ProcessId): $($_.Exception.Message)"
                try {
                    Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop
                    Write-Host "Force killed process $($process.ProcessName) with PID $($process.ProcessId)"
                } catch {
                    Write-Host "Failed to force kill process $($process.ProcessName) with PID $($process.ProcessId): $($_.Exception.Message)"
                }
            }
        }
    } else {
        Write-Host "Processes were not closed. Please close them manually if needed."
    }
}
[System.Windows.Forms.MessageBox]::Show("The process(es) have been closed", "Process Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::None)

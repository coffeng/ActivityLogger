# === CONFIG ===
$ScriptName = "ActivityLogger.ps1"
$ScriptPath = "$PSScriptRoot\$ScriptName"
$LogPath = "$PSScriptRoot\ActivityLog.csv"
$TaskName = "ActivityLoggerTask"
$IntervalSeconds = 10

# === DEFINE APP TAGS ===
$AppCategories = @{
    "Excel"     = "Work - Office"
    "Word"      = "Work - Office"
    "PowerPoint"= "Work - Office"
    "Outlook"   = "Email"
    "Chrome"    = "Web Browsing"
    "Edge"      = "Web Browsing"
    "YouTube"   = "Entertainment"
    "Teams"     = "Meetings"
    "Slack"     = "Communication"
    "Zoom"      = "Meetings"
    "Notepad"   = "Notes"
}

# === FUNCTION: Get Active Window Title ===
function Get-ActiveWindowTitle {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);
    }
"@
    $buffer = New-Object System.Text.StringBuilder 1024
    $handle = [Win32]::GetForegroundWindow()
    [Win32]::GetWindowText($handle, $buffer, $buffer.Capacity) | Out-Null
    return $buffer.ToString()
}

# === FUNCTION: Get Idle Time ===
function Get-IdleTime {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class IdleTime {
        [StructLayout(LayoutKind.Sequential)]
        public struct LASTINPUTINFO {
            public uint cbSize;
            public uint dwTime;
        }

        [DllImport("user32.dll")]
        public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        public static uint GetIdleTime() {
            LASTINPUTINFO lii = new LASTINPUTINFO();
            lii.cbSize = (uint)Marshal.SizeOf(lii);
            GetLastInputInfo(ref lii);
            return ((uint)Environment.TickCount - lii.dwTime) / 1000;
        }
    }
"@
    return [IdleTime]::GetIdleTime()
}

# === FUNCTION: Get App-Specific Details ===
function Get-WindowDetails {
    param ($windowTitle)

    if ($windowTitle -like "* - Excel") {
        try {
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            if ($excel.ActiveWorkbook -ne $null) {
                return $excel.ActiveWorkbook.FullName
            }
        } catch {}
    } elseif ($windowTitle -like "* - Word") {
        try {
            $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
            if ($word.ActiveDocument -ne $null) {
                return $word.ActiveDocument.FullName
            }
        } catch {}
    } elseif ($windowTitle -like "*Outlook*") {
        try {
            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            $inspector = $outlook.ActiveInspector()
            if ($inspector -ne $null -and $inspector.CurrentItem -ne $null) {
                return "Subject: " + $inspector.CurrentItem.Subject
            }
        } catch {}
    } elseif ($windowTitle -match " - (Google Chrome|Microsoft Edge)$") {
        return $windowTitle -replace " - (Google Chrome|Microsoft Edge)$", ""
    }

    return ""
}

# === FUNCTION: Categorize App ===
function Get-AppCategory {
    param ($windowTitle, $windowDetails)
    $combined = "$windowTitle $windowDetails"
    foreach ($key in $AppCategories.Keys) {
        if ($combined -like "*$key*") {
            return $AppCategories[$key]
        }
    }
    return "Uncategorized"
}

# === FUNCTION: Schedule Task at Login ===
function Ensure-ScheduledTask {
    if (-not (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)) {
        $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""
        $Trigger = New-ScheduledTaskTrigger -AtLogOn
        Register-ScheduledTask -Action $Action -Trigger $Trigger -TaskName $TaskName -Description "Logs user activity when active app changes" -User $env:USERNAME
    }
}

# === MAIN ===
Ensure-ScheduledTask

if (!(Test-Path $LogPath)) {
    "StartTime,EndTime,WindowTitle,DurationSeconds,WindowDetails,Category" | Out-File -FilePath $LogPath -Encoding UTF8
}

$prevWindow = ""
$startTime = Get-Date
$wasIdle = $false

while ($true) {
    $idleTime = Get-IdleTime
    $isIdle = $idleTime -ge $IntervalSeconds
    $currentWindow = if (-not $isIdle) { Get-ActiveWindowTitle } else { $prevWindow }

    $shouldLog = $false

    if ($isIdle -and -not $wasIdle) {
        $shouldLog = $true
    } elseif (-not $isIdle -and $wasIdle) {
        $startTime = Get-Date
    } elseif ($currentWindow -ne $prevWindow -and -not $isIdle) {
        $shouldLog = $true
    }

    if ($shouldLog -and $prevWindow -ne "") {
        $endTime = Get-Date
        $duration = [math]::Round(($endTime - $startTime).TotalSeconds, 1)
        $details = Get-WindowDetails -windowTitle $prevWindow
        $category = Get-AppCategory -windowTitle $prevWindow -windowDetails $details
        "$($startTime.ToString("yyyy-MM-dd HH:mm:ss")),$($endTime.ToString("yyyy-MM-dd HH:mm:ss")),""$prevWindow"",$duration,""$details"",""$category""" |
            Out-File -FilePath $LogPath -Encoding UTF8 -Append
        $startTime = Get-Date
    }

    if (-not $isIdle) {
        $prevWindow = $currentWindow
    }

    $wasIdle = $isIdle
    Start-Sleep -Seconds $IntervalSeconds
}

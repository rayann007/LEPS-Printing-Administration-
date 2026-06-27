
# =======================================
# LEPS PRINTING ADMINISTRATION V2.0
# Company: LEONI
# Author: Rayan Gazzah [LEPS ADMIN]

$script:AppVersion = "2.0"
# DEGLA  
# LEPS ADMINISTRATOR
# Dedicated for my beloved mother !
# =======================================

# ---- Required Assemblies ----
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Management.Automation

# ---- App Info ----
$script:AppName     = "LEPS PRINTING ADMINISTRATION V2.0"
$script:CompanyName = "LEONI"

# ---- Server Map ----
$script:ServerMap = [ordered]@{
    "1" = @{Label="Porsche/Lamborghini [LTN-1]"; Name="svltnporpnj03.leoni.local"}
    "2" = @{Label="Mercedes-Benz [LTN-1]";       Name="svltnmlbe7pnj03.leoni.local"}
    "3" = @{Label="MLB [LTN-1]";                 Name="svltnmlbe6pnj03.leoni.local"}
    "4" = @{Label="K8X [LTN-4]";                 Name="svltnk8xpnj03.leoni.local"}
    "5" = @{Label="G2X [LTN-4]";                 Name="svltng2xpnj03.leoni.local"}
    "6" = @{Label="High Voltage [LTN-2]";        Name="svtn4mebhvpnj03.leoni.local"}
    "7" = @{Label="NAX SZ [LTN-4]";              Name="svtn4naxszpnj03.leoni.local"}
    "8" = @{Label="NAX BZ [LTN-1]";              Name="svtn4naxbzpnj03.leoni.local"}
    "9" = @{Label="V2L [LTN-2]";                 Name="svtn4v2lmebpnj03.leoni.local"}
}
$ServerMap = $script:ServerMap


# ---- Global State ----
$script:CurrentServerLabel = $null
$script:CurrentServerTag   = ""

function Get-ServerTagFromLabel {
    param([string]$label)

    if ([string]::IsNullOrWhiteSpace($label)) { return "" }

    $start = $label.IndexOf("[")
    $end   = $label.IndexOf("]")
    if ($start -ge 0 -and $end -gt $start) {
        $inner = $label.Substring($start+1, $end-$start-1)
        $innerClean = $inner -replace "-", ""
        return ("[{0}]" -f $innerClean)
    }

    return ""
}

$Global:CimSession        = $null
$Global:LiveStatusRunning = $false
$Global:LiveStatusTimer   = $null
$Global:SessionWatchdogTimer = $null
$Global:WatchdogAlertShown   = $false
$Global:IsRefreshingPrinters = $false

# Theme state
$script:CurrentTheme = "Light"

# Self-healing thresholds
$script:SelfHealing_MaxJobAgeMinutes  = 30
$script:SelfHealing_MaxJobsThreshold  = 30

$printerCollection = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]

# ---- Audit Log (new) ----
$script:LogDirectory = Join-Path (Get-Location) "Logs"
if (-not (Test-Path $script:LogDirectory)) {
    New-Item -Path $script:LogDirectory -ItemType Directory -Force | Out-Null
}

$script:LogFile = Join-Path $script:LogDirectory ("LEPS_Audit_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd'))

$script:CurrentFilterText   = ""
$script:FilterPlaceholder   = "Search by Name, Driver or IP (e.g. pac10)"
$script:Window              = $null
$script:printerTable        = $null
$script:LogTextBox          = $null
$script:StatusFilterBox     = $null
$script:HasJobsOnly        = $false
$script:JobsWindow         = $null

# --------------------------------
# Utility: Log + Log View
# --------------------------------
function Refresh-LogView {
    try {
        if (-not $script:LogTextBox) { return }
        if (Test-Path $script:LogFile) {
            $lines = Get-Content -Path $script:LogFile -Tail 200
            $script:LogTextBox.Text = ($lines -join [Environment]::NewLine)
            $script:LogTextBox.ScrollToEnd()
        } else {
            $script:LogTextBox.Text = "<No log entries yet>"
        }
    } catch {
        # avoid GUI crash due to log read
    }
}

function Log {
    param(
        [string]$msg,
        [string]$ActionType   = "",
        [string]$Target       = "",
        [string]$Result       = "",
        [string]$ErrorMessage = ""
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $user      = $env:USERNAME

    # Structured audit line
    $auditLine = "{0}|{1}|{2}|{3}|{4}|{5}|{6}" -f `
        $timestamp,
        $user,
        ($ActionType   -replace '\|','/'),
        ($Target       -replace '\|','/'),
        ($Result       -replace '\|','/'),
        ($ErrorMessage -replace '\|','/'),
        ($msg          -replace '\|','/')

    try {
        Add-Content -Path $script:LogFile -Value $auditLine
    } catch {
        # ignore file write errors
    }

    # Human-readable line for Activity Log
    $displayLine = "$timestamp - $msg"
    try {
        if ($script:LogTextBox) {
            if ([string]::IsNullOrEmpty($script:LogTextBox.Text) -or
                $script:LogTextBox.Text -eq "<No log entries yet>") {
                $script:LogTextBox.Text = $displayLine
            } else {
                $script:LogTextBox.Text += [Environment]::NewLine + $displayLine
            }
            $script:LogTextBox.ScrollToEnd()
        }
    } catch {
        # ignore GUI log errors
    }
}


function Show-Message {
    param(
        [string]$msg,
        [string]$title = "Info"
    )
    [System.Windows.MessageBox]::Show($msg, $title) | Out-Null
}

function Show-ProgressDialog {
    param(
        [string]$Title   = "Connecting...",
        [string]$Message = "Please wait"
    )

    $win = New-Object System.Windows.Window -Property @{
        Title               = $Title
        Width               = 320
        Height              = 120
        WindowStartupLocation = "CenterScreen"
        ResizeMode          = "NoResize"
        WindowStyle         = "ToolWindow"
        ShowInTaskbar       = $false
        Topmost             = $true
    }

    $sp = New-Object System.Windows.Controls.StackPanel -Property @{ Margin = "12" }

    $label = New-Object System.Windows.Controls.TextBlock -Property @{
        Text       = $Message
        Margin     = "0,0,0,8"
        FontWeight = 'Bold'
    }

    $pb = New-Object System.Windows.Controls.ProgressBar -Property @{
        IsIndeterminate = $true
        Height          = 16
    }

    $sp.Children.Add($label) | Out-Null
    $sp.Children.Add($pb)    | Out-Null

    $win.Content = $sp
    $win.Show()

    return $win
}

function Show-InputBox {
    param(
        [string]$prompt,
        [string]$title,
        [string]$default = ""
    )

    $win = New-Object System.Windows.Window -Property @{
        Title               = $title
        SizeToContent       = "WidthAndHeight"
        WindowStartupLocation = "CenterScreen"
        ResizeMode          = "NoResize"
        WindowStyle         = "ToolWindow"
        Topmost             = $true
    }

    $stack = New-Object System.Windows.Controls.StackPanel
    $stack.Margin = "10"

    $lbl = New-Object System.Windows.Controls.TextBlock -Property @{
        Text   = $prompt
        Margin = "0,0,0,5"
    }

    $txt = New-Object System.Windows.Controls.TextBox -Property @{
        Text  = $default
        Width = 300
    }

    $btn = New-Object System.Windows.Controls.Button -Property @{
        Content    = "OK"
        Width      = 80
        Margin     = "0,5,0,0"
        IsDefault  = $true
        HorizontalAlignment = "Right"
    }

    $stack.Children.Add($lbl) | Out-Null
    $stack.Children.Add($txt) | Out-Null
    $stack.Children.Add($btn) | Out-Null

    $win.Content = $stack

    $result = $null
    $btn.Add_Click({
        $script:result = $txt.Text
        $win.Close()
    })

    $win.ShowDialog() | Out-Null
    return $result
}

function Show-CredentialDialog {
    param(
        [string]$serverFqdn
    )

    $win = New-Object System.Windows.Window -Property @{
        Title                 = "Credentials for $serverFqdn"
        SizeToContent         = "WidthAndHeight"
        WindowStartupLocation = "CenterScreen"
        ResizeMode            = "NoResize"
        WindowStyle           = "ToolWindow"
        Topmost               = $true
    }

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = "10"

    $text = New-Object System.Windows.Controls.TextBlock
    $text.Text = "Enter admin credentials for $serverFqdn"
    $text.Margin = "0,0,0,10"
    $panel.Children.Add($text) | Out-Null

    # Username row
    $userRow = New-Object System.Windows.Controls.StackPanel
    $userRow.Orientation = "Horizontal"
    $userRow.Margin = "0,0,0,5"

    $lblUser = New-Object System.Windows.Controls.TextBlock
    $lblUser.Text = "Username:"
    $lblUser.Width = 80
    $lblUser.VerticalAlignment = "Center"
    $userRow.Children.Add($lblUser) | Out-Null

    $txtUser = New-Object System.Windows.Controls.TextBox
    $txtUser.Width = 220
    $userRow.Children.Add($txtUser) | Out-Null

    $panel.Children.Add($userRow) | Out-Null

    # Password row
    $passRow = New-Object System.Windows.Controls.StackPanel
    $passRow.Orientation = "Horizontal"
    $passRow.Margin = "0,0,0,10"

    $lblPass = New-Object System.Windows.Controls.TextBlock
    $lblPass.Text = "Password:"
    $lblPass.Width = 80
    $lblPass.VerticalAlignment = "Center"
    $passRow.Children.Add($lblPass) | Out-Null

    $pwdBox = New-Object System.Windows.Controls.PasswordBox
    $pwdBox.Width = 220
    $passRow.Children.Add($pwdBox) | Out-Null

    $panel.Children.Add($passRow) | Out-Null

    # Buttons row
    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"

    $okBtn = New-Object System.Windows.Controls.Button
    $okBtn.Content = "OK"
    $okBtn.Width = 70
    $okBtn.Margin = "0,0,5,0"

    $cancelBtn = New-Object System.Windows.Controls.Button
    $cancelBtn.Content = "Cancel"
    $cancelBtn.Width = 70

    $btnRow.Children.Add($okBtn) | Out-Null
    $btnRow.Children.Add($cancelBtn) | Out-Null

    $panel.Children.Add($btnRow) | Out-Null

    # Button handlers: only manage dialog result, we read values AFTER ShowDialog returns
    $okBtn.Add_Click({
        $u = $txtUser.Text.Trim()
        $p = $pwdBox.Password

        if (-not $u -or -not $p) {
            [System.Windows.MessageBox]::Show(
                "Please enter both username and password.",
                "Missing data",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
            return
        }

        $win.DialogResult = $true
        $win.Close()
    })

    $cancelBtn.Add_Click({
        $win.DialogResult = $false
        $win.Close()
    })

    $win.Content = $panel

    # Show dialog synchronously
    $dialogResult = $win.ShowDialog()
    if (-not $dialogResult) {
        return $null
    }

    # After OK, collect values and build credential
    $finalUser = $txtUser.Text.Trim()
    $finalPass = $pwdBox.Password

    if (-not $finalUser -or -not $finalPass) {
        return $null
    }

    $secure = ConvertTo-SecureString $finalPass -AsPlainText -Force
    return New-Object System.Management.Automation.PSCredential($finalUser, $secure)
}
function Set-Theme {
    param(
        [ValidateSet("Light","Dark")]
        [string]$Mode
    )

    if (-not $script:Window) { return }

    $script:CurrentTheme = $Mode

    try {
        # Main window & root dock
        $rootDock = $script:Window.FindName("RootDock")
        if ($Mode -eq "Dark") {
            $bgDark   = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(17,24,39))
            $bgPanel  = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(31,41,55))
            $fgLight  = [System.Windows.Media.Brushes]::White

            $script:Window.Background = $bgDark
            $script:Window.Foreground = $fgLight
            if ($rootDock) { $rootDock.Background = $bgDark }

            $header = $script:Window.FindName("HeaderBorder")
            if ($header) { $header.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(15,23,42)) }

            # Main cards
            $bordersToDark = @(
                $script:Window.FindName("LeftPanelBorder"),
                $script:Window.FindName("MainDashboardBorder"),
                $script:Window.FindName("MainDetailsBorder"),
                $script:Window.FindName("MainLogBorder"),
                $script:Window.FindName("QuickActionsBorder"),
                $script:Window.FindName("ServerControlsBorder")
            )
            foreach ($b in $bordersToDark) {
                if ($b) { $b.Background = $bgPanel }
            }

            if ($script:printerTable) {
                $script:printerTable.Background = $bgPanel
                $script:printerTable.Foreground = $fgLight
            }
        } else {
            # Light theme: restore defaults
            $bgLightWindow = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(229,231,235))
            $bgPanelLight  = [System.Windows.Media.Brushes]::White
            $fgDark        = [System.Windows.Media.Brushes]::Black

            $script:Window.Background = $bgLightWindow
            $script:Window.Foreground = $fgDark
            if ($rootDock) { $rootDock.Background = $bgLightWindow }

            $header = $script:Window.FindName("HeaderBorder")
            if ($header) { $header.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(31,41,55)) }

            $bordersToLight = @(
                $script:Window.FindName("LeftPanelBorder"),
                $script:Window.FindName("MainDashboardBorder"),
                $script:Window.FindName("MainDetailsBorder"),
                $script:Window.FindName("MainLogBorder"),
                $script:Window.FindName("QuickActionsBorder"),
                $script:Window.FindName("ServerControlsBorder")
            )
            foreach ($b in $bordersToLight) {
                if ($b) { $b.Background = $bgPanelLight }
            }

            if ($script:printerTable) {
                $script:printerTable.Background = $bgPanelLight
                $script:printerTable.Foreground = $fgDark
            }
        }
    } catch {
        # don't kill UI for theme issues
    }
}

# --------------------------------
# Session Watchdog (connection health)
# --------------------------------
function Start-SessionWatchdog {
    if ($Global:SessionWatchdogTimer) {
        $Global:SessionWatchdogTimer.Stop()
        $Global:SessionWatchdogTimer = $null
    }

    if (-not $Global:CimSession) { return }

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(300)
    $timer.Add_Tick({
        if (-not $Global:CimSession) { return }
        try {
            # Lightweight health check (runs every 120s)
            Get-CimInstance -CimSession $Global:CimSession -ClassName Win32_OperatingSystem -ErrorAction Stop | Out-Null
        } catch {
            if (-not $Global:WatchdogAlertShown) {
                $Global:WatchdogAlertShown = $true
                Log "Session watchdog detected a lost connection: $($_.Exception.Message)" -ActionType "Watchdog" -Result "Error" -ErrorMessage $_.Exception.Message
                Show-Message "Connection to the print server appears lost. Please reconnect." "Connection Lost"
            } else {
                Log "Session watchdog ping failed again: $($_.Exception.Message)" -ActionType "Watchdog" -Result "Error" -ErrorMessage $_.Exception.Message
            }
            try {
                if ($Global:LiveStatusRunning) { Stop-LiveStatus }
            } catch {}
        }
    })

    $Global:SessionWatchdogTimer = $timer
    $Global:WatchdogAlertShown   = $false
    $timer.Start()

    Log "Session watchdog started (300s interval)." -ActionType "Watchdog" -Result "Info"
    if ($statusBar) {
        $statusBar.Text = "Watchdog active (connection monitored every 300s)..."
    }
}

function Stop-SessionWatchdog {
    if ($Global:SessionWatchdogTimer) {
        $Global:SessionWatchdogTimer.Stop()
        $Global:SessionWatchdogTimer = $null
    }
    Log "Session watchdog stopped." -ActionType "Watchdog" -Result "Info"
    if ($statusBar) {
        $statusBar.Text = "Watchdog stopped."
    }
}

# --------------------------------
# Advanced filter parsing (name:, status:, vlan:, ip:, driver:)
# --------------------------------
function Parse-PrinterFilterSpec {
    param([string]$filterText)

    $spec = [ordered]@{
        TextTerms = @()
        Name      = $null
        Status    = $null
        VLAN      = $null
        IP        = $null
        Driver    = $null
    }

    if ([string]::IsNullOrWhiteSpace($filterText)) {
        return $spec
    }

    $parts = $filterText.Split(' ',[System.StringSplitOptions]::RemoveEmptyEntries)
    foreach ($p in $parts) {
        if ($p -match '^(?i)name:(.+)$')   { $spec.Name   = $Matches[1]; continue }
        if ($p -match '^(?i)status:(.+)$') { $spec.Status = $Matches[1]; continue }
        if ($p -match '^(?i)vlan:(.+)$')   { $spec.VLAN   = $Matches[1]; continue }
        if ($p -match '^(?i)ip:(.+)$')     { $spec.IP     = $Matches[1]; continue }
        if ($p -match '^(?i)driver:(.+)$') { $spec.Driver = $Matches[1]; continue }
        $spec.TextTerms += $p
    }

    return $spec
}
# --------------------------------
# Filtering (Text + Status)
# --------------------------------

function Apply-PrinterFilter {
    param(
        [string]$filterText,
        [string]$statusFilter
    )

    $script:CurrentFilterText = $filterText

    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($printerCollection)
    if (-not $view) { return }

    if ([string]::IsNullOrWhiteSpace($filterText) -and ($statusFilter -eq "All" -or -not $statusFilter)) {
        $view.Filter = $null
        return
    }

    $filterSpec = Parse-PrinterFilterSpec -filterText $filterText

    $view.Filter = {
        param($item)
        if (-not $item) { return $false }

        $name   = [string]$item.Name
        $driver = [string]$item.Driver
        $ip     = [string]$item.IP
        $vlan   = [string]$item.VLAN
        $queue  = 0
        if ($item.PSObject.Properties.Match("QueueCount").Count -gt 0 -and $item.QueueCount -ne $null) {
            [int]::TryParse($item.QueueCount.ToString(), [ref]$queue) | Out-Null
        }

        # Structured filters
        if ($filterSpec.Name -and $name -notlike "*$($filterSpec.Name)*") { return $false }
        if ($filterSpec.IP   -and $ip   -notlike "*$($filterSpec.IP)*")   { return $false }
        if ($filterSpec.Driver -and $driver -notlike "*$($filterSpec.Driver)*") { return $false }
        if ($filterSpec.VLAN -and $vlan -notlike "*$($filterSpec.VLAN)*") { return $false }

        # Text terms (each term must match at least one of Name/Driver/IP)
        $matchesText = $true
        foreach ($t in $filterSpec.TextTerms) {
            $pt = $t.ToLower()
            $termMatch = (($name   -and $name.ToLower().Contains($pt)) -or
                          ($driver -and $driver.ToLower().Contains($pt)) -or
                          ($ip     -and $ip.ToLower().Contains($pt)))
            if (-not $termMatch) {
                $matchesText = $false
                break
            }
        }

        if (-not $matchesText) { return $false }

        # Status filter (UI combo)
        $matchesStatus = $true
        switch ($statusFilter) {
            "Online" {
                $matchesStatus = $item.PrinterStatus -in @("Ready","Idle","Printing")
            }
            "Offline" {
                $matchesStatus = $item.PrinterStatus -eq "Offline"
            }
            "Error" {
                $matchesStatus = $item.PrinterStatus -eq "Error"
            }
            default {
                $matchesStatus = $true
            }
        }

        # Has jobs only filter from checkbox
        if ($script:HasJobsOnly -and $queue -le 0) {
            return $false
        }

        # Extra status filter from advanced syntax (status:offline etc.)
        if ($filterSpec.Status) {
            $wanted = $filterSpec.Status.ToLower()
            $cur    = ([string]$item.PrinterStatus).ToLower()
            if ($wanted -eq "online") {
                if ($item.PrinterStatus -notin @("Ready","Idle","Printing")) { return $false }
            } elseif ($wanted -eq "offline") {
                if ($item.PrinterStatus -ne "Offline") { return $false }
            } elseif ($wanted -eq "error") {
                if ($item.PrinterStatus -ne "Error") { return $false }
            } else {
                if ($cur -notlike "*$wanted*") { return $false }
            }
        }

        return $matchesStatus
    }
}

# --------------------------------
# Server Connection (WITH RETRY)
# --------------------------------

function Connect-ToServer {
    param([string]$serverFqdn)

    while ($true) {
        $cred = $null
        try {
            $cred = Show-CredentialDialog -serverFqdn $serverFqdn
        } catch {
            Show-Message "Credential dialog failed. Not connecting." "Credentials"
            Log ("Credential dialog threw an exception for {0}: {1}" -f $serverFqdn, $_.Exception.Message)
            return $false
        }

        if (-not $cred) {
            Show-Message "No credentials provided. Not connecting." "Credentials"
            Log "No credentials provided for $serverFqdn (dialog canceled or empty)."
            return $false
        }

        $progressWin = Show-ProgressDialog -Title "Connecting" -Message "Connecting to $serverFqdn..."

        try {
            $Global:CimSession = New-CimSession -ComputerName $serverFqdn -Credential $cred -ErrorAction Stop
            Log "Connected to $serverFqdn with explicit credentials"
            Update-ResourceMonitor
            $progressWin.Close()
            return $true
        } catch {
            $msg = $_.Exception.Message
            $progressWin.Close()
            Log "Failed to connect $($serverFqdn): $msg"

            $fullMsg = "Failed to connect to $($serverFqdn):`n$msg`n`nDo you want to try again?"
            $choice = [System.Windows.MessageBox]::Show(
                $fullMsg,
                "Connection / Authentication Error",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )

            if ($choice -ne [System.Windows.MessageBoxResult]::Yes) {
                Show-Message "Not connected to $serverFqdn." "Connection"
                return $false
            }
            # else: loop and ask again
        }
    }
}

function Disconnect-Server {
    if ($Global:CimSession) {
        $serverName = $Global:CimSession.ComputerName
        try {
            Remove-CimSession $Global:CimSession -ErrorAction SilentlyContinue
        } catch {
            # ignore errors on close
        }
        $Global:CimSession = $null
        Log "Disconnected from server $serverName"
        $script:CurrentServerLabel = $null
        $script:CurrentServerTag   = ""
    } else {
        Log "Disconnect called but no CimSession existed"
    }
}

# --------------------------------
# Dashboard & Details
# --------------------------------
function Update-Dashboard {
    param($Window)

    try {
        $Window.FindName("TotalPrinters").Text  = "$($printerCollection.Count)"
        $online = ($printerCollection | Where-Object { $_.PrinterStatus -in @("Idle","Ready","Printing") }).Count
        $offline = ($printerCollection | Where-Object { $_.PrinterStatus -in @("Offline","Error") }).Count
        $Window.FindName("OnlinePrinters").Text  = "$online"
        $Window.FindName("OfflinePrinters").Text = "$offline"

        $totalJobs = 0
        foreach ($p in $printerCollection) {
            if ($p.PSObject.Properties.Match("QueueCount").Count -gt 0 -and $p.QueueCount -ne $null) {
                [int]$val = 0
                if ([int]::TryParse($p.QueueCount.ToString(), [ref]$val)) {
                    $totalJobs += $val
                }
            }
        }
        $Window.FindName("TotalJobs").Text = "$totalJobs"
    } catch {
        Log "Update-Dashboard failed: $($_.Exception.Message)"
    }
}

function Update-ResourceMonitor {
    try {
        $healthBlock = $script:Window.FindName("ServerHealth")
        if (-not $healthBlock) { return }

        if (-not $Global:CimSession) {
            $healthBlock.Text = "CPU: - | RAM: - | Disk C: - (not connected)"
            return
        }

        $cpuObj = Get-CimInstance -CimSession $Global:CimSession -ClassName Win32_Processor -ErrorAction SilentlyContinue
        $cpu    = ($cpuObj | Measure-Object -Property LoadPercentage -Average).Average

        $os = Get-CimInstance -CimSession $Global:CimSession -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($os) {
            $totalMem = [double]$os.TotalVisibleMemorySize
            $freeMem  = [double]$os.FreePhysicalMemory
            $usedMemPct = if ($totalMem -gt 0) { [math]::Round((( $totalMem - $freeMem ) / $totalMem) * 100, 1) } else { 0 }
        } else {
            $usedMemPct = $null
        }

        $disk = Get-CimInstance -CimSession $Global:CimSession -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
        if ($disk) {
            $usedDiskPct = if ($disk.Size -gt 0) {
                [math]::Round((1 - ($disk.FreeSpace / $disk.Size)) * 100, 1)
            } else { 0 }
        } else {
            $usedDiskPct = $null
        }

        # Spooler uptime removed from resource monitor
        # (no spooler uptime calculation)

        $cpuText  = if ($cpu      -ne $null) { "{0}%" -f [math]::Round($cpu,1) } else { "-" }
        $ramText  = if ($usedMemPct -ne $null) { "{0}%" -f $usedMemPct } else { "-" }
        $diskText = if ($usedDiskPct -ne $null) { "{0}%" -f $usedDiskPct } else { "-" }

        $healthBlock.Text = "CPU: {0} | RAM: {1} | Disk C: {2}" -f $cpuText,$ramText,$diskText
    } catch {
        # don't break UI because of monitoring
    }
}



function Update-SelectedPrinterDetails {
    try {
        $sel = $script:printerTable.SelectedItem

        $nameBlock   = $script:Window.FindName("DetailName")
        $ipBlock     = $script:Window.FindName("DetailIP")
        $statusBlock = $script:Window.FindName("DetailStatus")
        $driverBlock = $script:Window.FindName("DetailDriver")
        $vlanBlock   = $script:Window.FindName("DetailVLAN")
        $swpBlock    = $script:Window.FindName("DetailSwitchPort")

        if (-not $sel) {
            $nameBlock.Text   = "Name: -"
            $ipBlock.Text     = "IP: -"
            $statusBlock.Text = "Status: -"
            $driverBlock.Text = "Driver: -"
            $vlanBlock.Text   = "VLAN: -"
            $swpBlock.Text    = "Switch Port: -"
            return
        }

        $nameBlock.Text   = "Name: {0}"         -f $sel.Name
        $ipBlock.Text     = "IP: {0}"           -f $sel.IP
        $statusBlock.Text = "Status: {0}"       -f $sel.PrinterStatus
        $driverBlock.Text = "Driver: {0}"       -f $sel.Driver
        $vlanBlock.Text   = "VLAN: {0}"         -f $sel.VLAN
        $swpBlock.Text    = "Switch Port: {0}"  -f $sel.SwitchPort
    } catch {
        Log "Update-SelectedPrinterDetails failed: $($_.Exception.Message)"
    }
}

# --------------------------------
# Data Refresh
# --------------------------------
function Refresh-PrintersTableAsync {
    param(
        [switch]$LightWeight,        # Light = quick status-only refresh (no job enumeration)
        [switch]$SuppressErrorPopup  # When set, don't show popup on refresh failure
    )

    if (-not $Global:CimSession) { return }

    # Prevent overlapping refreshes (manual + live status + other actions)
    if ($Global:IsRefreshingPrinters) {
        return
    }

    $Global:IsRefreshingPrinters = $true
    $progressWin = $null

    if ($statusBar) {
        $statusBar.Text = "Refreshing printers..."
    }
    Log "Refresh (UI) started." -ActionType "Refresh" -Result "Info"

    try {
        # Only show progress dialog for "deep" refreshes
        if (-not $LightWeight) {
            $progressWin = Show-ProgressDialog -Title "Loading printers..." -Message "Please wait while printers and queues are loaded..."
        }

        # Keep a snapshot of existing printers (to preserve VLAN/SwitchPort etc.)
        $existingByName = @{}
        foreach ($old in $printerCollection) {
            if ($old.Name) {
                $existingByName[$old.Name] = $old
            }
        }

        # Get all printers once
        $printers = Get-Printer -CimSession $Global:CimSession -ErrorAction Stop

        $newList = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]

        foreach ($p in $printers | Sort-Object Name) {
            $existing = $null
            if ($existingByName.ContainsKey($p.Name)) {
                $existing = $existingByName[$p.Name]
            }

            # Queue count (deep refresh uses per-printer query like V1)
            $qCount = 0
            if (-not $LightWeight) {
                try {
                    $jobs    = Get-PrintJob -PrinterName $p.Name -CimSession $Global:CimSession -ErrorAction SilentlyContinue
                    $jobsArr = @($jobs)
                    if ($jobsArr.Count -gt 0) {
                        $qCount = ($jobsArr | Measure-Object -Property TotalPages -Sum).Sum
                        if (-not $qCount -or $qCount -eq 0) {
                            $qCount = $jobsArr.Count   # fallback to job count
                        }
                    }
                } catch {
                    $qCount = 0
                }
            } elseif ($existing) {
                # For lightweight refresh, preserve previous queue count
                $qCount = $existing.QueueCount
            }

            # Normalize printer status into UI values
            $status = $p.PrinterStatus
            switch ($status) {
                'Normal'   { $status = 'Ready' }
                'Idle'     { $status = 'Idle' }
                'Printing' { $status = 'Printing' }
                'Offline'  { $status = 'Offline' }
                'Error'    { $status = 'Error' }
                $null      { $status = 'Other' }
                default    { $status = 'Other' }
            }

            # Preserve VLAN / SwitchPort if they were filled before
            $vlan       = if ($existing) { $existing.VLAN }       else { "" }
            $switchPort = if ($existing) { $existing.SwitchPort } else { "" }

            $obj = [PSCustomObject]@{
                Name          = $p.Name
                IP            = ($p.PortName -replace "^IP_","")
                Driver        = $p.DriverName
                PrinterStatus = $status
                QueueCount    = $qCount
                VLAN          = $vlan
                SwitchPort    = $switchPort
                ServerTag    = $script:CurrentServerTag
            }

            $newList.Add($obj)
        }

        # Replace collection in one shot
        $printerCollection.Clear()
        foreach ($obj in $newList) {
            $printerCollection.Add($obj)
        }

        # Reapply filter + status combo
        if ($script:StatusFilterBox) {
            $statusItem = $script:StatusFilterBox.SelectedItem
            $statusText = "All"
            if ($statusItem -and $statusItem.Content) { $statusText = $statusItem.Content.ToString() }

            $filterBox = $script:Window.FindName("FilterBox")
            $effective = if ($filterBox) { Get-FilterTextEffective -raw $filterBox.Text } else { "" }
            Apply-PrinterFilter -filterText $effective -statusFilter $statusText
        } else {
            Apply-PrinterFilter -filterText $script:CurrentFilterText -statusFilter "All"
        }

        Update-Dashboard      -Window $script:Window
        Update-SelectedPrinterDetails
    } catch {
        Log "Refresh (UI) is running in background..." `
            -ActionType "Refresh" `
            -Result "Info" `
            -ErrorMessage $_.Exception.Message

        if (-not $SuppressErrorPopup) {
            Show-Message "Refresh is loading…
The printer list is updating in the background." "Refreshing"
        }
    } finally {
        if ($statusBar) {
            $statusBar.Text = "Refresh completed."
        }
        if ($progressWin) {
            try { $progressWin.Close() } catch {}
        }
        $Global:IsRefreshingPrinters = $false
    }
}

# --------------------------------
# Live Status (Auto Refresh)
# --------------------------------
function Start-LiveStatus {
    if (-not $Global:CimSession) {
        Show-Message "Connect to a server first." "No Connection"
        return
    }

    if ($Global:LiveStatusRunning) {
        Show-Message "Live status already running." "Live Status"
        return
    }

    # Determine interval from AutoRefreshInterval ComboBox (default 30s)
    $intervalSec = 30
    try {
        $cmb = $script:Window.FindName("AutoRefreshInterval")
        if ($cmb -and $cmb.SelectedItem) {
            $txt = $cmb.SelectedItem.Content.ToString()
            if ($txt -match '(\d+)') {
                $intervalSec = [int]$Matches[1]
            }
        }
    } catch {
        # fallback to default 10s
    }

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds($intervalSec)
    $timer.Add_Tick({
        try {
            Refresh-PrintersTableAsync -LightWeight -SuppressErrorPopup
        } catch {
            Log "Live status refresh error: $($_.Exception.Message)" -ActionType "LiveStatus" -Result "Error" -ErrorMessage $_.Exception.Message
        }
    })

    $Global:LiveStatusTimer   = $timer
    $Global:LiveStatusRunning = $true
    $timer.Start()

    Show-Message ("Live status started (refresh every {0} seconds)." -f $intervalSec) "Live Status"
    Log ("Live status started (interval {0}s)" -f $intervalSec) -ActionType "LiveStatus" -Result "Success"
}


function Stop-LiveStatus {
    if ($Global:LiveStatusTimer) {
        try { $Global:LiveStatusTimer.Stop() } catch {}
        $Global:LiveStatusTimer = $null
    }
    $Global:LiveStatusRunning = $false
    Show-Message "Live status stopped." "Live Status"
    Log "Live status stopped"
}

# --------------------------------
# Actions on Selected Printers
# --------------------------------
function GUI-RestartSpooler {
    if (-not $Global:CimSession) {
        Show-Message "Connect to server first." "No Connection"
        return
    }

    $server = $Global:CimSession.ComputerName
    $q = "Are you sure you want to restart the Print Spooler on server $server ?"
    $choice = [System.Windows.MessageBox]::Show(
        $q,
        "Confirm Restart Spooler",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($choice -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $progressWin = $null
    try {
        $progressWin = Show-ProgressDialog -Title "Restarting Spooler" -Message "Please wait while the spooler is restarted on $server..."

        # Use CIM to control the remote Spooler service via the existing CimSession
        $svc = Get-CimInstance -CimSession $Global:CimSession -ClassName Win32_Service -Filter "Name='Spooler'" -ErrorAction Stop
        $null = Invoke-CimMethod -InputObject $svc -MethodName StopService -ErrorAction Stop
        Start-Sleep -Seconds 2
        $null = Invoke-CimMethod -InputObject $svc -MethodName StartService -ErrorAction Stop

        Show-Message "Print Spooler restarted successfully." "Action"
        Log "Spooler restarted on $server" -ActionType "RestartSpooler" -Target $server -Result "Success"
        Start-Sleep -Seconds 2
        Refresh-PrintersTableAsync -SuppressErrorPopup
    } catch {
        Show-Message "Failed to restart spooler:`n$($_.Exception.Message)`n`nSee Activity Log for detailed errors." "Error"
        Log "Spooler restart failed on ${server}: $($_.Exception.Message)" -ActionType "RestartSpooler" -Target $server -Result "Error" -ErrorMessage $_.Exception.Message
    } finally {
        if ($progressWin) {
            try { $progressWin.Close() } catch {}
        }
    }
}


function GUI-ClearQueueSelected {
    if (-not $Global:CimSession) {
        Show-Message "Connect to server first." "No Connection"
        return
    }

    $selected = $script:printerTable.SelectedItems
    if (-not $selected -or $selected.Count -eq 0) {
        Show-Message "Select printers first." "No Selection"
        return
    }

    $count = $selected.Count

    # Compute total jobs across selected printers for better confirmation
    $totalJobs = 0
    foreach ($p in $selected) {
        try {
            $jobs = Get-PrintJob -PrinterName $p.Name -CimSession $Global:CimSession -ErrorAction SilentlyContinue
            if ($jobs) {
                $totalJobs += ($jobs | Measure-Object).Count
            }
        } catch {
            # ignore job count errors here, they'll be handled during actual clear
        }
    }

    $q = "You are about to clear queues for $count printer(s), total jobs: $totalJobs. Are you sure?"
    $choice = [System.Windows.MessageBox]::Show(
        $q,
        "Confirm Clear Queue",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($choice -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $failedPrinters = @()

    foreach ($printer in $selected) {
        try {
            Get-PrintJob -PrinterName $printer.Name -CimSession $Global:CimSession -ErrorAction SilentlyContinue |
                Remove-PrintJob -ErrorAction SilentlyContinue
            Log "Cleared queue for $($printer.Name)" -ActionType "ClearQueue" -Target $printer.Name -Result "Success"
        } catch {
            $failedPrinters += $printer.Name
            Log "Failed to clear queue for $($printer.Name): $($_.Exception.Message)" -ActionType "ClearQueue" -Target $printer.Name -Result "Error" -ErrorMessage $_.Exception.Message
        }
    }

    if ($failedPrinters.Count -gt 0) {
        Show-Message ("Queues cleared, but these printers had errors:`n{0}`n`nSee Activity Log for detailed errors." -f ($failedPrinters -join ", ")) "Clear Queue"
    } else {
        Show-Message "Selected printer queues cleared." "Action"
    }

    Refresh-PrintersTableAsync -SuppressErrorPopup
}

function GUI-SelfHealingServer {
    if (-not $Global:CimSession) {
        Show-Message "Connect to server first." "No Connection"
        return
    }

    $server = $Global:CimSession.ComputerName

    $summary = @()
    $summary += "Self-Healing will:"
    $summary += "- Scan printers for Offline/Error status."
    $summary += "- Look for large / old queues (>{0} jobs or age >{1} min)." -f $script:SelfHealing_MaxJobsThreshold,$script:SelfHealing_MaxJobAgeMinutes
    $summary += "- Clear jobs that are obviously stuck (older than threshold)."
    $summary += ""
    $summary += "Run on server: $server ?"

    $msgBoxText = ($summary -join "`n")
    $choice = [System.Windows.MessageBox]::Show(
        $msgBoxText,
        "Confirm Self-Healing",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    if ($choice -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $problemPrinters = @()
    $clearedJobs     = 0

    foreach ($p in $printerCollection) {
        $isProblem = $false

        if ($p.PrinterStatus -in @("Offline","Error")) {
            $isProblem = $true
        }

        try {
            $jobs = Get-PrintJob -PrinterName $p.Name -CimSession $Global:CimSession -ErrorAction SilentlyContinue
        } catch {
            $jobs = @()
        }

        if ($jobs -and $jobs.Count -ge $script:SelfHealing_MaxJobsThreshold) {
            $isProblem = $true
        }

        if ($jobs) {
            foreach ($j in @($jobs)) {
                $age = (Get-Date) - $j.SubmitTime
                if ($age.TotalMinutes -ge $script:SelfHealing_MaxJobAgeMinutes) {
                    try {
                        Remove-PrintJob -InputObject $j -ErrorAction SilentlyContinue
                        $clearedJobs++
                    } catch {
                        Log "Self-Healing failed to remove job $($j.Id) on $($p.Name): $($_.Exception.Message)" -ActionType "SelfHealing" -Target $p.Name -Result "Error" -ErrorMessage $_.Exception.Message
                    }
                }
            }
        }

        if ($isProblem) {
            $problemPrinters += $p.Name
        }
    }

    Log ("Self-Healing executed on {0}. Problem printers (before): {1}. Cleared jobs: {2}" -f $server,($problemPrinters -join ", "),$clearedJobs) -ActionType "SelfHealing" -Target $server -Result "Success"

    Refresh-PrintersTableAsync -SuppressErrorPopup

    $resultLines = @()
    $resultLines += "Self-Healing finished on $server."
    $resultLines += "Problem printers detected: " + ($problemPrinters.Count)
    if ($problemPrinters.Count -gt 0) {
        $resultLines += "List: " + ($problemPrinters -join ", ")
    }
    $resultLines += "Cleared jobs (older than $($script:SelfHealing_MaxJobAgeMinutes) minutes): $clearedJobs"

    Show-Message ($resultLines -join "`n") "Self-Healing Summary"
}

function GUI-ShowServerPerformance {
    if (-not $Global:CimSession) {
        Show-Message "Connect to server first." "No Connection"
        return
    }

    $server = $Global:CimSession.ComputerName
    $progressWin = $null

    try {
        $progressWin = Show-ProgressDialog -Title "Reading server performance" -Message "Collecting CPU/RAM/disk metrics from $server..."

        # CPU
        $cpuObj = Get-CimInstance -CimSession $Global:CimSession -ClassName Win32_Processor -ErrorAction SilentlyContinue
        $cpu = $null
        if ($cpuObj) {
            $cpu = ($cpuObj | Measure-Object -Property LoadPercentage -Average).Average
        }

        # RAM
        $os = Get-CimInstance -CimSession $Global:CimSession -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
        $ramText = "N/A"
        if ($os) {
            $totalMem = [double]$os.TotalVisibleMemorySize
            $freeMem  = [double]$os.FreePhysicalMemory
            if ($totalMem -gt 0) {
                $usedMemPct = [math]::Round((( $totalMem - $freeMem ) / $totalMem) * 100, 1)
                $ramText = "$usedMemPct %"
            }
        }

        # Disk C:
        $diskText = "N/A"
        $disk = Get-CimInstance -CimSession $Global:CimSession -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
        if ($disk -and $disk.Size -gt 0) {
            $usedDiskPct = [math]::Round((1 - ($disk.FreeSpace / $disk.Size)) * 100, 1)
            $diskText = "$usedDiskPct %"
        }

        if ($progressWin) {
            try { $progressWin.Close() } catch {}
        }

        # Build small window
        $win = New-Object System.Windows.Window
        $win.Title = "Server Performance - $server"
        $win.SizeToContent = 'WidthAndHeight'
        $win.WindowStartupLocation = 'CenterOwner'
        $win.ResizeMode = 'NoResize'
        if ($script:Window) { $win.Owner = $script:Window }

        $panel = New-Object System.Windows.Controls.StackPanel
        $panel.Margin = '10'

        $titleBlock = New-Object System.Windows.Controls.TextBlock
        $titleBlock.Text = "Current resource usage for $server"
        $titleBlock.FontWeight = 'Bold'
        $titleBlock.Margin = '0,0,0,8'
        $panel.Children.Add($titleBlock)

        $cpuText = if ($cpu -ne $null) { ("{0} %" -f [math]::Round($cpu,1)) } else { "N/A" }

        $cpuBlock = New-Object System.Windows.Controls.TextBlock
        $cpuBlock.Text = "CPU: $cpuText"
        $cpuBlock.Margin = '0,2,0,2'
        $panel.Children.Add($cpuBlock)

        $ramBlock = New-Object System.Windows.Controls.TextBlock
        $ramBlock.Text = "RAM Used: $ramText"
        $ramBlock.Margin = '0,2,0,2'
        $panel.Children.Add($ramBlock)

        $diskBlock = New-Object System.Windows.Controls.TextBlock
        $diskBlock.Text = "Disk C: Used: $diskText"
        $diskBlock.Margin = '0,2,0,2'
        $panel.Children.Add($diskBlock)

        $infoBlock = New-Object System.Windows.Controls.TextBlock
        $infoBlock.TextWrapping = 'Wrap'
        $infoBlock.Margin = '0,10,0,0'
        $infoBlock.Text = "Note: This is a lightweight snapshot. Use Live Status + Dashboard for continuous monitoring."
        $panel.Children.Add($infoBlock)

        $win.Content = $panel
        $win.ShowDialog() | Out-Null

        Log ("Server performance window opened for {0} (CPU={1}, RAM={2}, Disk={3})" -f $server,$cpuText,$ramText,$diskText) -ActionType "ServerPerformance" -Target $server -Result "Success"
    } catch {
        if ($progressWin) {
            try { $progressWin.Close() } catch {}
        }
        Show-Message "Failed to read server performance:`n$($_.Exception.Message)" "Server Performance"
        Log ("Server performance failed for {0}: {1}" -f $server,$_.Exception.Message) -ActionType "ServerPerformance" -Target $server -Result "Error" -ErrorMessage $_.Exception.Message
    }
}
function GUI-DeployPrinterWizard {
    # For now, reuse the Add Printer logic as a simple deployment wizard
    GUI-AddPrinter
}

function GUI-PingSelected {
    $selected = $script:printerTable.SelectedItems
    if (-not $selected -or $selected.Count -eq 0) {
        Show-Message "Select printers first." "No Selection"
        return
    }

    $results = @()
    foreach ($printer in $selected) {
        $ip = $printer.IP
        if ([string]::IsNullOrWhiteSpace($ip)) {
            $msg = "$($printer.Name): No IP defined"
            $results += $msg
            Log $msg -ActionType "Ping" -Target $printer.Name -Result "Error" -ErrorMessage "No IP"
            continue
        }

        try {
            $pings = Test-Connection -ComputerName $ip -Count 2 -ErrorAction SilentlyContinue
            if ($pings) {
                $avg = [math]::Round( ( ($pings | Measure-Object -Property ResponseTime -Average).Average ), 1 )
                $msg = "$($printer.Name) ($ip): Reachable, avg RTT ${avg}ms"
                $printer.PrinterStatus = "Ready"
                Log $msg -ActionType "Ping" -Target "$($printer.Name) ($ip)" -Result "Reachable"
            } else {
                $msg = "$($printer.Name) ($ip): UNREACHABLE"
                $printer.PrinterStatus = "Offline"
                Log $msg -ActionType "Ping" -Target "$($printer.Name) ($ip)" -Result "Unreachable"
            }
            $results += $msg
        } catch {
            $msg = "$($printer.Name) ($ip): ERROR - $($_.Exception.Message)"
            $printer.PrinterStatus = "Error"
            $results += $msg
            Log $msg -ActionType "Ping" -Target "$($printer.Name) ($ip)" -Result "Error" -ErrorMessage $_.Exception.Message
        }
    }

    # Reapply filter and update dashboard/details
    $statusItem = $script:StatusFilterBox.SelectedItem
    $statusText = "All"
    if ($statusItem -and $statusItem.Content) { $statusText = $statusItem.Content.ToString() }

    $filterBox = $script:Window.FindName("FilterBox")
    $effective = if ($filterBox) { Get-FilterTextEffective -raw $filterBox.Text } else { "" }
    Apply-PrinterFilter -filterText $effective -statusFilter $statusText

    Update-Dashboard -Window $script:Window
    Update-SelectedPrinterDetails

    # Force UI refresh so status / colors update immediately
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:printerTable.ItemsSource)
    if ($view) { $view.Refresh() }

    $text = $results -join "`n"
    Show-Message $text "Ping Results"
}


function GUI-TestPrintSelected {
    if (-not $Global:CimSession) {
        Show-Message "Connect to server first." "No Connection"
        return
    }

    $selected = $script:printerTable.SelectedItems
    if (-not $selected -or $selected.Count -eq 0) {
        Show-Message "Select printers first." "No Selection"
        return
    }

    $ok = 0; $fail = 0
    foreach ($printer in $selected) {
        try {
            $filter = "Name='{0}'" -f ($printer.Name.Replace("'","''"))
            $wmiPrinter = Get-CimInstance -CimSession $Global:CimSession `
                                          -ClassName Win32_Printer `
                                          -Filter $filter `
                                          -ErrorAction Stop

            if (-not $wmiPrinter) {
                Log "Test page FAILED for $($printer.Name): Win32_Printer not found with filter [$filter]"
                $fail++
                continue
            }

            $result = Invoke-CimMethod -CimSession $Global:CimSession `
                                       -InputObject $wmiPrinter `
                                       -MethodName PrintTestPage `
                                       -ErrorAction Stop

            Log "Test page requested for $($printer.Name) - ReturnValue: $($result.ReturnValue)"
            if ($result.ReturnValue -eq 0) {
                $ok++
            } else {
                $fail++
            }
        } catch {
            Log "Test page FAILED for $($printer.Name): $($_.Exception.Message)"
            $fail++
        }
    }

    Show-Message "Test print requested. OK: $ok | Failed: $fail`nSee Activity Log for details." "Test Page"
}

function GUI-InspectQueueSelected {
    if (-not $Global:CimSession) {
        Show-Message "Connect to server first." "No Connection"
        return
    }

    $selected = $script:printerTable.SelectedItems
    if (-not $selected -or $selected.Count -eq 0) {
        Show-Message "Select printers first." "No Selection"
        return
    }

    $rows       = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
    $totalJobs  = 0
    $totalPages = 0

    foreach ($printer in $selected) {
        try {
            $jobs = Get-PrintJob -PrinterName $printer.Name -CimSession $Global:CimSession -ErrorAction SilentlyContinue

            if (-not $jobs -or $jobs.Count -eq 0) {
                Log "InspectQueue: 0 jobs returned for $($printer.Name)" -ActionType "InspectJobs" -Target $printer.Name -Result "Info"
                continue
            }

            foreach ($j in $jobs) {
                $age       = (Get-Date) - $j.SubmitTime
                $ageString = "{0}d {1}h {2}m {3}s" -f $age.Days, $age.Hours, $age.Minutes, $age.Seconds

                $pages = 0
                if ($j.TotalPages) {
                    try {
                        $pages = [int]$j.TotalPages
                    } catch {
                        $pages = 0
                    }
                }

                $rows.Add([PSCustomObject]@{
                    Printer   = $printer.Name
                    JobId     = $j.Id
                    Document  = $j.DocumentName
                    User      = $j.Submitter
                    Pages     = $pages
                    Status    = $j.JobStatus
                    Submitted = $j.SubmitTime
                    Age       = $ageString
                })

                $totalJobs++
                $totalPages += $pages
            }
        } catch {
            Log "Job inspection failed for $($printer.Name): $($_.Exception.Message)" -ActionType "InspectJobs" -Target $printer.Name -Result "Error" -ErrorMessage $_.Exception.Message
        }
    }

    if ($rows.Count -eq 0) {
        Show-Message "No print jobs found for selected printers." "Job Queue"
        Log "InspectQueue: no jobs found." -ActionType "InspectJobs" -Result "Success"
        return
    }

    Log "InspectQueue opened for $($selected.Count) printer(s), $($rows.Count) job(s)." -ActionType "InspectJobs" -Result "Success"

    $titleSuffix = "Printer(s): {0} | Jobs: {1} | Pages: {2}" -f $selected.Count, $totalJobs, $totalPages

    # Close any previous jobs inspector window before opening a new one
    if ($script:JobsWindow -and $script:JobsWindow.IsVisible) {
        try { $script:JobsWindow.Close() } catch {}
    }

    $win = New-Object System.Windows.Window -Property @{
        Title               = "Job Queue Inspector - $titleSuffix"
        Width               = 900
        Height              = 400
        WindowStartupLocation = "CenterOwner"
        Owner               = $script:Window
    }

    $grid = New-Object System.Windows.Controls.DataGrid -Property @{
        AutoGenerateColumns = $true
        ItemsSource         = $rows
        IsReadOnly          = $true
    }

    $win.Content = $grid
    $script:JobsWindow = $win
    $win.ShowDialog() | Out-Null
}


function GUI-PauseSelected {
    if (-not $Global:CimSession) {
        Show-Message "Connect to server first." "No Connection"
        return
    }

    $selected = $script:printerTable.SelectedItems
    if (-not $selected -or $selected.Count -eq 0) {
        Show-Message "Select printers first." "No Selection"
        return
    }

    # Make sure PrintManagement module is available (for Set-Printer / Suspend-PrintJob on some systems)
    try {
        if (-not (Get-Module -Name PrintManagement -ErrorAction SilentlyContinue)) {
            Import-Module -Name PrintManagement -ErrorAction Stop
            Log "PrintManagement module imported for Pause operation." -ActionType "Module" -Result "Success"
        }
    } catch {
        Log "Failed to import PrintManagement module for Pause: $($_.Exception.Message)" -ActionType "Module" -Result "Error" -ErrorMessage $_.Exception.Message
    }

    foreach ($printer in $selected) {
        try {
            $spCmd = Get-Command Set-Printer -ErrorAction SilentlyContinue
            if ($spCmd -and $spCmd.Parameters.ContainsKey('AcceptingJobs')) {
                Set-Printer -Name $printer.Name -CimSession $Global:CimSession -AcceptingJobs $false -ErrorAction Stop
            } else {
                Log "Set-Printer -AcceptingJobs not available, skipping AcceptingJobs toggle for $($printer.Name)." -ActionType "Pause" -Target $printer.Name -Result "Info"
            }

            Get-PrintJob -PrinterName $printer.Name -CimSession $Global:CimSession -ErrorAction SilentlyContinue |
                Suspend-PrintJob -ErrorAction SilentlyContinue

            Log "Paused $($printer.Name)" -ActionType "Pause" -Target $printer.Name -Result "Success"
        } catch {
            Log "Pause failed for $($printer.Name): $($_.Exception.Message)" -ActionType "Pause" -Target $printer.Name -Result "Error" -ErrorMessage $_.Exception.Message
        }
    }

    Show-Message "Selected printers paused." "Action"
    Refresh-PrintersTableAsync -SuppressErrorPopup
}

function GUI-ResumeSelected {
    if (-not $Global:CimSession) {
        Show-Message "Connect to server first." "No Connection"
        return
    }

    $selected = $script:printerTable.SelectedItems
    if (-not $selected -or $selected.Count -eq 0) {
        Show-Message "Select printers first." "No Selection"
        return
    }

    # Make sure PrintManagement module is available (for Set-Printer / Resume-PrintJob on some systems)
    try {
        if (-not (Get-Module -Name PrintManagement -ErrorAction SilentlyContinue)) {
            Import-Module -Name PrintManagement -ErrorAction Stop
            Log "PrintManagement module imported for Resume operation." -ActionType "Module" -Result "Success"
        }
    } catch {
        Log "Failed to import PrintManagement module for Resume: $($_.Exception.Message)" -ActionType "Module" -Result "Error" -ErrorMessage $_.Exception.Message
    }

    foreach ($printer in $selected) {
        try {
            $spCmd = Get-Command Set-Printer -ErrorAction SilentlyContinue
            if ($spCmd -and $spCmd.Parameters.ContainsKey('AcceptingJobs')) {
                Set-Printer -Name $printer.Name -CimSession $Global:CimSession -AcceptingJobs $true -ErrorAction Stop
            } else {
                Log "Set-Printer -AcceptingJobs not available, skipping AcceptingJobs toggle for $($printer.Name)." -ActionType "Resume" -Target $printer.Name -Result "Info"
            }

            Get-PrintJob -PrinterName $printer.Name -CimSession $Global:CimSession -ErrorAction SilentlyContinue |
                Resume-PrintJob -ErrorAction SilentlyContinue

            Log "Resumed $($printer.Name)" -ActionType "Resume" -Target $printer.Name -Result "Success"
        } catch {
            Log "Resume failed for $($printer.Name): $($_.Exception.Message)" -ActionType "Resume" -Target $printer.Name -Result "Error" -ErrorMessage $_.Exception.Message
        }
    }

    Show-Message "Selected printers resumed." "Action"
    Refresh-PrintersTableAsync -SuppressErrorPopup
}


function GUI-ExportCSV {
    if ($printerCollection.Count -eq 0) {
        Show-Message "Nothing to export." "Export"
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        Filter          = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        FileName        = "LEPS_Printers_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        OverwritePrompt = $true
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $printerCollection |
                Sort-Object Name |
                Select-Object Name,IP,Driver,PrinterStatus,QueueCount,VLAN,SwitchPort |
                Export-Csv -NoTypeInformation -Path $dlg.FileName -Encoding UTF8

            Show-Message "Exported to:`n$($dlg.FileName)" "Export"
            Log "Export CSV: $($dlg.FileName)" -ActionType "ExportCSV" -Target $dlg.FileName -Result "Success"
        } catch {
            Show-Message "Export failed:`n$($_.Exception.Message)" "Export Error"
            Log "Export CSV failed: $($_.Exception.Message)" -ActionType "ExportCSV" -Target $dlg.FileName -Result "Error" -ErrorMessage $_.Exception.Message
        }
    }
}

function GUI-ExportTXT {
    if ($printerCollection.Count -eq 0) {
        Show-Message "Nothing to export." "Export"
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        Filter          = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        FileName        = "LEPS_Printers_$(Get-Date -Format 'yyyyMMdd_HHmm').txt"
        OverwritePrompt = $true
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $lines = $printerCollection |
                Sort-Object Name |
                ForEach-Object {
                    "{0};{1};{2};{3};{4};{5};{6}" -f $_.Name,$_.IP,$_.Driver,$_.PrinterStatus,$_.QueueCount,$_.VLAN,$_.SwitchPort
                }

            Set-Content -Path $dlg.FileName -Value $lines -Encoding UTF8
            Show-Message "Exported to:`n$($dlg.FileName)" "Export"
            Log "Export TXT: $($dlg.FileName)" -ActionType "ExportTXT" -Target $dlg.FileName -Result "Success"
        } catch {
            Show-Message "Export failed:`n$($_.Exception.Message)" "Export Error"
            Log "Export TXT failed: $($_.Exception.Message)" -ActionType "ExportTXT" -Target $dlg.FileName -Result "Error" -ErrorMessage $_.Exception.Message
        }
    }
}

function GUI-ExportSelection {
    if ($printerCollection.Count -eq 0) {
        Show-Message "Nothing to export." "Export"
        return
    }

    $msg = "Choose export format:`n`nYes = CSV (recommended)`nNo = TXT"
    $result = [System.Windows.MessageBox]::Show(
        $msg,
        "Export Format",
        [System.Windows.MessageBoxButton]::YesNoCancel,
        [System.Windows.MessageBoxImage]::Question
    )

    switch ($result) {
        ([System.Windows.MessageBoxResult]::Yes) { GUI-ExportCSV }
        ([System.Windows.MessageBoxResult]::No)  { GUI-ExportTXT }
        default { return }
    }
}

function GUI-ImportConfig {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Filter      = "Text or CSV files (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*"
        Multiselect = $false
        Title       = "Select a printer configuration file"
    }

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $file = $dlg.FileName
    try {
        $lines = Get-Content -Path $file -ErrorAction Stop
        if (-not $lines -or $lines.Count -eq 0) {
            Show-Message "Selected file is empty." "Import"
            return
        }

        $targetName = Show-InputBox "Enter the printer name you want to inspect (exact or partial match):" "Import Printer Config"
        if (-not $targetName) { return }

        $matches = @()

        foreach ($line in $lines) {
            # Expect: Name;IP;Driver;Status;QueueCount;VLAN;SwitchPort
            $parts = $line -split ';'
            if ($parts.Count -lt 3) { continue }

            $name = $parts[0]
            if ($name -like "*$targetName*") {
                $obj = [PSCustomObject]@{
                    Name        = $parts[0]
                    IP          = if ($parts.Count -gt 1) { $parts[1] } else { "" }
                    Driver      = if ($parts.Count -gt 2) { $parts[2] } else { "" }
                    Status      = if ($parts.Count -gt 3) { $parts[3] } else { "" }
                    QueuePages  = if ($parts.Count -gt 4) { $parts[4] } else { "" }
                    VLAN        = if ($parts.Count -gt 5) { $parts[5] } else { "" }
                    SwitchPort  = if ($parts.Count -gt 6) { $parts[6] } else { "" }
                }
                $matches += $obj
            }
        }

        if (-not $matches -or $matches.Count -eq 0) {
            Show-Message "No printer in the file matched '$targetName'." "Import"
            return
        }

        $summary = $matches | ForEach-Object {
            "Name: {0}`nIP: {1}`nDriver: {2}`nStatus: {3}`nQueue Pages: {4}`nVLAN: {5}`nSwitch Port: {6}`n`n" -f `
                $_.Name,$_.IP,$_.Driver,$_.Status,$_.QueuePages,$_.VLAN,$_.SwitchPort
        } -join "`n"

        [System.Windows.MessageBox]::Show(
            $summary,
            "Imported Printer Configuration(s)",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null

        Log "Imported printer configuration(s) from $file for search '$targetName'" -ActionType "ImportConfig" -Target $file -Result "Success"
    } catch {
        Show-Message "Import failed:`n$($_.Exception.Message)" "Import Error"
        Log "Import config failed: $($_.Exception.Message)" -ActionType "ImportConfig" -Target $file -Result "Error" -ErrorMessage $_.Exception.Message
    }
}

function GUI-ImportServers {
    # Let the admin choose a server definition file.
    # Supported formats:
    # - JSON: [ { "Name": "svltng2xpnj03.leoni.local", "Label": "G2X [LTN-4]" }, ... ]
    # - CSV/TXT: lines with either:
    #     Fqdn;Label
    #     or Fqdn,Label
    #   (header row is optional and will be skipped if it contains "fqdn" or "label").
    $dlg = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Filter      = "JSON, CSV or TXT (*.json;*.csv;*.txt)|*.json;*.csv;*.txt|All files (*.*)|*.*"
        Multiselect = $false
        Title       = "Import print servers list"
    }

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $file = $dlg.FileName
    $ext  = [System.IO.Path]::GetExtension($file).ToLowerInvariant()

    $newMap = [ordered]@{}

    try {
        if ($ext -eq ".json") {
            $data = Get-Content -Path $file -Raw | ConvertFrom-Json
            $idx  = 1
            foreach ($entry in $data) {
                if (-not $entry.Name -or -not $entry.Label) { continue }
                $key = [string]$idx
                $newMap[$key] = @{
                    Name  = [string]$entry.Name
                    Label = [string]$entry.Label
                }
                $idx++
            }
        } else {
            $lines = Get-Content -Path $file | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

            $idx = 1
            foreach ($line in $lines) {
                $trim = $line.Trim()
                if ([string]::IsNullOrWhiteSpace($trim)) { continue }

                # Header detection (for CSV with header row)
                if ($trim -match 'fqdn' -and $trim -match 'label') {
                    continue
                }

                # First try ; then ,
                $parts = $trim -split ';'
                if ($parts.Count -lt 2) {
                    $parts = $trim -split ','
                }
                if ($parts.Count -lt 2) {
                    continue
                }

                $fqdn  = $parts[0].Trim()
                $label = $parts[1].Trim()

                if ([string]::IsNullOrWhiteSpace($fqdn) -or [string]::IsNullOrWhiteSpace($label)) {
                    continue
                }

                $key = [string]$idx
                $newMap[$key] = @{
                    Name  = $fqdn
                    Label = $label
                }
                $idx++
            }
        }

        if ($newMap.Count -eq 0) {
            Show-Message "No valid servers found in file.`nExpected JSON or CSV/TXT with FQDN and label." "Import Servers"
            return
        }

        # Replace global server map
        $script:ServerMap = $newMap
        $Global:ServerMap = $newMap

        # Rebuild the Print Servers combo with new entries
        Refresh-ServerCombo

        Show-Message ("Imported {0} server(s) from:`n{1}" -f $newMap.Count, $file) "Import Servers"
        Log "Imported servers list from $file ($($newMap.Count) entries)" -ActionType "ImportServers" -Target $file -Result "Success"
    } catch {
        Show-Message "Import Servers failed:`n$($_.Exception.Message)" "Import Servers Error"
        Log "Import Servers failed: $($_.Exception.Message)" -ActionType "ImportServers" -Target $file -Result "Error" -ErrorMessage $_.Exception.Message
    }
}

# --------------------------------
# Add Server manually
# --------------------------------
function GUI-AddServer {
    # Prompt the admin for a new print server definition.
    # The FQDN is required, the label is optional (defaults to the FQDN).
    $fqdn = Show-InputBox "Server FQDN:" "Add Server"
    if (-not $fqdn) {
        return
    }

    $label = Show-InputBox "Friendly name (label):" "Add Server"
    if (-not $label) {
        $label = $fqdn
    }

    # Trim whitespace from inputs
    $fqdn  = $fqdn.Trim()
    $label = $label.Trim()

    # Prevent duplicates
    foreach ($key in $script:ServerMap.Keys) {
        if ($script:ServerMap[$key].Name -eq $fqdn) {
            Show-Message "Server '$fqdn' already exists in the list." "Add Server"
            return
        }
    }

    # Determine the next available numeric key
    $newKey = 1
    if ($script:ServerMap.Count -gt 0) {
        $intKeys = $script:ServerMap.Keys | ForEach-Object { [int]$_ } | Sort-Object
        $newKey  = ($intKeys[-1] + 1)
    }
    $newKeyStr = [string]$newKey

    # Add the server into both script and global maps
    $script:ServerMap[$newKeyStr] = @{
        Name  = $fqdn
        Label = $label
    }
    $Global:ServerMap = $script:ServerMap

    # Refresh the server combo so the new server appears
    Refresh-ServerCombo

    Show-Message "Server '$fqdn' added to the list. This change lasts for the current session." "Add Server"
    Log "Server added: $fqdn ($label)" -ActionType "AddServer" -Result "Success"
}

# --------------------------------
# Select all printers in the table
# --------------------------------
function GUI-SelectAllPrinters {
    if (-not $script:printerTable) {
        return
    }
    try {
        # Try DataGrid.SelectAll() if available
        if ($script:printerTable.GetType().GetMethod("SelectAll")) {
            $script:printerTable.SelectAll()
        } else {
            # Fallback: manually select every item
            $script:printerTable.SelectedItems.Clear()
            foreach ($row in $script:printerTable.Items) {
                $script:printerTable.SelectedItems.Add($row)
            }
        }
        Log "All printers selected." -ActionType "SelectAll" -Result "Success"
    } catch {
        Log "Select all printers failed: $($_.Exception.Message)" -ActionType "SelectAll" -Result "Error" -ErrorMessage $_.Exception.Message
    }
}

function GUI-AddPrinter {
    if (-not $Global:CimSession) {
        Show-Message "Connect to server first." "No Connection"
        return
    }

    $pName  = Show-InputBox "Printer Name:" "Add Printer"
    if (-not $pName) { return }

    $ipAddr = Show-InputBox "Printer IP Address:" "Add Printer"
    if (-not $ipAddr) { return }

    $driver = Show-InputBox "Driver Name (as installed on server):" "Add Printer"
    if (-not $driver) { return }

    try {
        $portName = "IP_$ipAddr"
        $port = Get-PrinterPort -CimSession $Global:CimSession -Name $portName -ErrorAction SilentlyContinue

        if (-not $port) {
            Add-PrinterPort -CimSession $Global:CimSession -Name $portName -PrinterHostAddress $ipAddr -ErrorAction Stop
            Log "Added port $portName ($ipAddr)"
        }

        $drv = Get-PrinterDriver -CimSession $Global:CimSession -Name $driver -ErrorAction SilentlyContinue
        if (-not $drv) {
            Show-Message "Driver '$driver' not found on server. Install the driver first." "Add Printer"
            Log "Add printer failed: Driver '$driver' not found"
            return
        }

        Add-Printer -CimSession $Global:CimSession -Name $pName -DriverName $driver -PortName $portName -ErrorAction Stop
        Show-Message "Printer '$pName' added on server." "Add Printer"
        Log "Added printer $pName ($ipAddr) Driver:$driver"
        Refresh-PrintersTableAsync -SuppressErrorPopup
    } catch {
        Show-Message "Add printer failed:`n$($_.Exception.Message)" "Error"
        Log "Add printer failed: $($_.Exception.Message)"
    }
}

# --------------------------------
# Copy Name / IP (clipboard)
# --------------------------------
function GUI-CopyNameSelected {
    $selected = $script:printerTable.SelectedItems
    if (-not $selected -or $selected.Count -eq 0) {
        Show-Message "Select at least one printer." "Copy Name"
        return
    }

    $names = $selected | ForEach-Object { $_.Name } | Where-Object { $_ }
    if (-not $names -or $names.Count -eq 0) {
        Show-Message "No printer names found." "Copy Name"
        return
    }

    $text = ($names -join "`r`n")
    [System.Windows.Clipboard]::SetText($text)
    Log "Copied printer name(s) to clipboard: $($names -join ', ')"
    Show-Message "Copied:`n$text" "Copy Name"
}

function GUI-CopyIPSelected {
    $selected = $script:printerTable.SelectedItems
    if (-not $selected -or $selected.Count -eq 0) {
        Show-Message "Select at least one printer." "Copy IP"
        return
    }

    $ips = $selected | ForEach-Object { $_.IP } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if (-not $ips -or $ips.Count -eq 0) {
        Show-Message "No IP addresses found for selected printers." "Copy IP"
        return
    }

    $text = ($ips -join "`r`n")
    [System.Windows.Clipboard]::SetText($text)
    Log "Copied IP(s) to clipboard: $($ips -join ', ')"
    Show-Message "Copied:`n$text" "Copy IP"
}

# --------------------------------
# XAML (GUI Layout with LEONI Branding + Tutorial)
# --------------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="LEPS PRINTING ADMINISTRATION V2.4" Height="800" Width="1300"
        WindowStartupLocation="CenterScreen"
        ShowInTaskbar="True"
        Background="#FFE5E7EB">
    <DockPanel Name="RootDock" Margin="10">
        <!-- Top Panel / Branding -->
        <Border Name="HeaderBorder" DockPanel.Dock="Top" Background="#FF1F2937" Padding="10" CornerRadius="8" Margin="0,0,0,10">
            <DockPanel LastChildFill="False">
                <StackPanel DockPanel.Dock="Left">
                    <TextBlock Text="LEONI"
                               FontSize="14"
                               FontWeight="Bold"
                               Foreground="#FF60A5FA"
                               Margin="0,0,0,2"/>
                    <TextBlock Text="LEPS PRINTING ADMINISTRATION V2.4"
                               FontWeight="Bold"
                               Foreground="White"/>
                    <TextBlock Text="Centralized Printer Monitoring &amp; Management"
                               FontSize="12"
                               Foreground="#FF93C5FD"/>
                </StackPanel>
                <StackPanel DockPanel.Dock="Right"
                            Orientation="Vertical"
                            VerticalAlignment="Center"
                            HorizontalAlignment="Right"
                            Margin="0,0,10,0">
                    <Button Name="ThemeToggleButton"
                            Content="Dark Mode"
                            Margin="0,0,0,4"
                            Padding="10,4"
                            Background="#FF4B5563"
                            Foreground="White"
                            FontSize="11"/>
                    <Button Name="HelpButton"
                            Content="Help &amp; About"
                            Margin="0,0,0,4"
                            Padding="10,4"
                            Background="#FF312E81"
                            Foreground="White"
                            FontSize="11"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- Server Controls + Quick Actions -->
        <StackPanel DockPanel.Dock="Top" Orientation="Vertical" Margin="0,0,0,10">
            <Border Name="ServerControlsBorder" Background="White" CornerRadius="8" Padding="8" Margin="0,0,0,10">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*" />
                        <ColumnDefinition Width="Auto" />
                    </Grid.ColumnDefinitions>
                    <StackPanel Orientation="Horizontal" Grid.Column="0">
                        <TextBlock Text="Print Servers:" VerticalAlignment="Center" Margin="0,4,8,0" FontWeight="Bold"/>
                        <ComboBox Name="ServerCombo" Width="340" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <Button Name="ConnectButton" Content="Connect" Width="100" Margin="0,0,8,0"
                                Background="#FF22C55E" Foreground="White" FontWeight="Bold"/>
                        <Button Name="DisconnectBtn" Content="Disconnect" Width="100" Margin="0,0,8,0"
                                Background="#FFEF4444" Foreground="White"/>
                        <Button Name="RefreshButton" Content="Refresh" Width="100" Margin="0,0,8,0"
                                Background="#FFCBD5F5"/>
                        <Button Name="AddPrinterBtn" Content="Add Server" Width="120" Margin="0,0,8,0"
                                Background="#FF3B82F6" Foreground="White"
                                ToolTip="Add a new print server to the list (session only)"/>



                        <CheckBox Name="AutoRefreshCheck" Content="Auto Refresh" VerticalAlignment="Center" Margin="0,0,6,0"/>
                        <ComboBox Name="AutoRefreshInterval" Width="80" SelectedIndex="0" VerticalAlignment="Center">
                            <ComboBoxItem Content="30 s"/>
                            <ComboBoxItem Content="60 s"/>
                            <ComboBoxItem Content="120 s"/>
                        </ComboBox>
                        <Button Name="ToggleLeftPanelBtn" Content="Hide Left Panel" Width="130" Margin="10,0,0,0"
                                Background="#FFE5E7EB"/>
                    </StackPanel>
                    <Button Name="CloseBtn"
                            Grid.Column="1"
                            Content="Close"
                            Width="90"
                            Margin="10,0,0,0"
                            Background="#FFEF4444"
                            Foreground="White"
                            FontWeight="Bold"/>
                </Grid>
            </Border>

            <Border Name="QuickActionsBorder" Background="White" CornerRadius="8" Padding="8" HorizontalAlignment="Stretch">
                <GroupBox Header="[ Quick Actions ]" BorderThickness="0" Margin="0" FontWeight="Bold">
                    <StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                            <TextBlock Text="Server Actions" FontWeight="Bold" Margin="0,0,8,0" TextDecorations="Underline"/>
                        </StackPanel>
                        <WrapPanel Margin="0,0,0,4">
                            <Button Name="QA3"  Content="Restart Spooler" Width="140" Margin="4"
                                    Background="#FFF97373" FontWeight="Bold"
                                    ToolTip="Restart the print spooler service on the connected server"/>
                            <Button Name="QA14" Content="Self-Healing" Width="140" Margin="4"
                                    Background="#FFDCFCE7" FontWeight="Bold"
                                    ToolTip="Run automatic checks and safe fixes on the connected server"/>
                            <Button Name="QA15" Content="Server Performance" Width="150" Margin="4"
                                    Background="#FFDBEAFE" FontWeight="Bold"
                                    ToolTip="Show CPU, RAM and disk usage for the connected server"/>
                            <Button Name="QA16" Content="Deploy Printer" Width="140" Margin="4"
                                    Background="#FF3B82F6" Foreground="White" FontWeight="Bold"
                                    ToolTip="Open the deployment wizard to add a new printer"/>
                            <Button Name="QA5"  Content="Live Status" Width="140" Margin="4"
                                    Background="#FFBFDBFE" FontWeight="Bold"
                                    ToolTip="Toggle automatic auto-refresh using the selected interval"/>
                        </WrapPanel>

                        <StackPanel Orientation="Horizontal" Margin="0,8,0,4">
                            <TextBlock Text="Printer Actions" FontWeight="Bold" Margin="0,0,8,0" TextDecorations="Underline"/>
                        </StackPanel>
                        <WrapPanel Margin="0,0,0,4">
                            <Button Name="QA6"  Content="Ping"            Width="90"  Margin="4"
                                    Background="#FFDBEAFE"
                                    ToolTip="Ping one or more selected printers and show reachability"/>
                            <Button Name="QA7"  Content="Test Page"       Width="100" Margin="4"
                                    Background="#FFDBEAFE"
                                    ToolTip="Send a Windows test page to selected printers"/>
                            <Button Name="QA13" Content="View Jobs"       Width="110" Margin="4"
                                    Background="#FFDBEAFE"
                                    ToolTip="Inspect job queue (with job age) for selected printers"/>
                            <Button Name="QA8"  Content="Select All"     Width="100" Margin="4"
                                    Background="#FFDBEAFE"
                                    ToolTip="Select all printers in the table"/>
                            <!-- Removed Pause/Resume quick action buttons -->
                            <Button Name="QA4"  Content="Clear Queue"     Width="140" Margin="4"
                                    Background="#FFF97373" FontWeight="Bold"
                                    ToolTip="Delete all pending jobs for selected printers"/>
                            <Button Name="QA2"  Content="Errors"          Width="90"  Margin="4"
                                    Background="#FFDBEAFE"
                                    ToolTip="Show printers with non-OK status"/>
                            <Button Name="QA1"  Content="List"            Width="90"  Margin="4"
                                    Background="#FFDBEAFE"
                                    ToolTip="Show a simple text list of all printers"/>
                        </WrapPanel>

                        <StackPanel Orientation="Horizontal" Margin="0,8,0,4">
                            <TextBlock Text="Export &amp; Import" FontWeight="Bold" Margin="0,0,8,0" TextDecorations="Underline"/>
                        </StackPanel>
                        <WrapPanel>
                            <Button Name="QA10" Content="Export CSV/TXT"  Width="140" Margin="4"
                                    ToolTip="Export printers to CSV or TXT file"/>
                            <Button Name="QA12" Content="Import Config"   Width="120" Margin="4"
                                    ToolTip="Import printer settings from a file"/>
                            <Button Name="QA18" Content="Import Servers"  Width="140" Margin="4"
                                    ToolTip="Import print server list (FQDN + label) from a file"/>
                        </WrapPanel>
                    </StackPanel>
                </GroupBox>
            </Border>
        </StackPanel>

        <!-- Bottom Panel / Status Bar -->
        <DockPanel DockPanel.Dock="Bottom" Margin="0,10,0,0">
            <TextBlock Name="StatusBar" Text="" VerticalAlignment="Center" FontStyle="Italic"/>
        </DockPanel>
        <!-- Main Content -->
        <Grid Name="MainGrid">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="3*"/>
            </Grid.ColumnDefinitions>

            <!-- Left: Search + Printers (Card) -->
            <Border Name="LeftPanelBorder" Grid.Column="0" Background="White" CornerRadius="8" Padding="8" Margin="0,0,8,0">
                <StackPanel>
                    <TextBlock Text="Printers" FontSize="18" FontWeight="Bold" Margin="0,0,0,8" TextDecorations="Underline"/>

                    <!-- Search / Filter Bar -->
                    <StackPanel Orientation="Vertical" Margin="0,0,0,6">
                        <TextBlock Text="[ Search &amp; Filter ]" FontWeight="Bold" Margin="0,0,0,2"/>
                        <StackPanel Orientation="Horizontal">
                            <TextBox Name="FilterBox" Width="260" Margin="0,0,6,0"/>
                            <TextBlock Text="Status:" VerticalAlignment="Center" Margin="10,0,6,0"/>
                            <ComboBox Name="StatusFilterBox" Width="120" SelectedIndex="0">
                                <ComboBoxItem Content="All"/>
                                <ComboBoxItem Content="Online"/>
                                <ComboBoxItem Content="Offline"/>
                                <ComboBoxItem Content="Error"/>
                            </ComboBox>
                            <CheckBox Name="HasJobsCheckBox" Content="Has jobs only" VerticalAlignment="Center" Margin="10,0,0,0"/>
                            <Button Name="FilterBtn" Content="Apply" Width="70" Margin="10,0,0,0" Background="#FF3B82F6" Foreground="White" FontWeight="Bold"/>
                        </StackPanel>
                    </StackPanel>

                    <DataGrid Name="PrinterTable" AutoGenerateColumns="False"
                              SelectionMode="Extended" SelectionUnit="FullRow" Height="500"
                              CanUserSortColumns="True">
                        <DataGrid.RowStyle>
                            <Style TargetType="DataGridRow">
                                <Style.Triggers>
                                    <!-- Online -->
                                    <DataTrigger Binding="{Binding PrinterStatus}" Value="Ready">
                                        <Setter Property="Background" Value="LightGreen"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding PrinterStatus}" Value="Idle">
                                        <Setter Property="Background" Value="LightGreen"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding PrinterStatus}" Value="Printing">
                                        <Setter Property="Background" Value="LightYellow"/>
                                    </DataTrigger>

                                    <!-- Offline / Error -->
                                    <DataTrigger Binding="{Binding PrinterStatus}" Value="Offline">
                                        <Setter Property="Background" Value="LightCoral"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding PrinterStatus}" Value="Error">
                                        <Setter Property="Background" Value="LightCoral"/>
                                    </DataTrigger>

                                    <!-- Other/Unknown -->
                                    <DataTrigger Binding="{Binding PrinterStatus}" Value="Other">
                                        <Setter Property="Background" Value="LightGray"/>
                                    </DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.RowStyle>

                        <DataGrid.ContextMenu>
                            <ContextMenu>
                                <MenuItem Header="Copy Name"/>
                                <MenuItem Header="Copy IP"/>
                                <Separator/>
                                <MenuItem Name="CtxViewJobs" Header="View Jobs (selected)"/>
                                <Separator/>
                                <MenuItem Header="Test Page (selected)"/>
                                <!-- Removed Pause/Resume items from context menu -->
                                <MenuItem Header="Clear Queue (selected)"/>
                            </ContextMenu>
                        </DataGrid.ContextMenu>

                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Srv"         Binding="{Binding ServerTag}"     Width="Auto"/>
                            <DataGridTextColumn Header="Name"        Binding="{Binding Name}"          Width="2*"/>

                            <DataGridTemplateColumn Header="Status" Width="*">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <StackPanel Orientation="Horizontal">
                                            <Ellipse Width="10" Height="10" Margin="0,0,4,0">
                                                <Ellipse.Style>
                                                    <Style TargetType="Ellipse">
                                                        <Setter Property="Fill" Value="Gray"/>
                                                        <Style.Triggers>
                                                            <DataTrigger Binding="{Binding PrinterStatus}" Value="Ready">
                                                                <Setter Property="Fill" Value="Green"/>
                                                            </DataTrigger>
                                                            <DataTrigger Binding="{Binding PrinterStatus}" Value="Idle">
                                                                <Setter Property="Fill" Value="Green"/>
                                                            </DataTrigger>
                                                            <DataTrigger Binding="{Binding PrinterStatus}" Value="Printing">
                                                                <Setter Property="Fill" Value="Yellow"/>
                                                            </DataTrigger>
                                                            <DataTrigger Binding="{Binding PrinterStatus}" Value="Offline">
                                                                <Setter Property="Fill" Value="Red"/>
                                                            </DataTrigger>
                                                            <DataTrigger Binding="{Binding PrinterStatus}" Value="Error">
                                                                <Setter Property="Fill" Value="Red"/>
                                                            </DataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </Ellipse.Style>
                                            </Ellipse>
                                            <TextBlock Text="{Binding PrinterStatus}"/>
                                        </StackPanel>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>

                            <DataGridTextColumn Header="IP"          Binding="{Binding IP}"            Width="*"/>
                            <DataGridTextColumn Header="VLAN"        Binding="{Binding VLAN}"         Width="*"/>
                            <DataGridTextColumn Header="Switch Port" Binding="{Binding SwitchPort}"   Width="*"/>
                            <DataGridTextColumn Header="Driver"      Binding="{Binding Driver}"        Width="2*"/>
                            <DataGridTextColumn Header="QP"          Binding="{Binding QueueCount}"   Width="60"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </StackPanel>
            </Border>

                    <!-- Right: Dashboard + Details + Log (Cards) -->
        <Grid Grid.Column="1" Margin="8,0,0,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- Dashboard Statistics -->
            <Border Grid.Row="0" Background="White" CornerRadius="8" Padding="8" Margin="0,0,0,8" HorizontalAlignment="Stretch">
                <GroupBox Header="[ Dashboard Statistics ]" Margin="0" BorderThickness="0" FontWeight="Bold">
                    <StackPanel Orientation="Vertical" Margin="4">
                        <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                            <TextBlock Text="Total Printers:" FontWeight="Bold" Margin="0,0,6,0"/>
                            <TextBlock Name="TotalPrinters"  Text="0" FontSize="18" FontWeight="Bold"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                            <TextBlock Text="Online:" FontWeight="Bold" Margin="0,0,6,0"/>
                            <TextBlock Name="OnlinePrinters" Text="0" FontSize="16" Foreground="Green"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                            <TextBlock Text="Offline:" FontWeight="Bold" Margin="0,0,6,0"/>
                            <TextBlock Name="OfflinePrinters" Text="0" FontSize="16" Foreground="Red"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                            <TextBlock Text="Total Jobs:" FontWeight="Bold" Margin="0,0,6,0"/>
                            <TextBlock Name="TotalJobs" Text="0" FontSize="16" Foreground="DarkBlue"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                            <TextBlock Text="Server Health:" FontWeight="Bold" Margin="0,0,6,0"/>
                            <TextBlock Name="ServerHealth"
                                       Text="CPU: - | RAM: - | Disk C: -"
                                       FontSize="12"/>
                            <Button Name="RefreshMonitorBtn"
                                    Content="Refresh"
                                    Margin="8,0,0,0"
                                    Padding="8,2"
                                    ToolTip="Refresh server CPU/RAM/disk snapshot"/>
                        </StackPanel>
                    </StackPanel>
                </GroupBox>
            </Border>

            <!-- Printer Details -->
            <Border Grid.Row="1" Background="White" CornerRadius="8" Padding="8" Margin="0,0,0,8" HorizontalAlignment="Stretch">
                <GroupBox Header="[ Printer Details ]" Margin="0" BorderThickness="0" FontWeight="Bold">
                    <StackPanel>
                        <TextBlock Name="DetailName"       Text="Name: -"         Margin="0,2,0,0"/>
                        <TextBlock Name="DetailIP"         Text="IP: -"           Margin="0,2,0,0"/>
                        <TextBlock Name="DetailStatus"     Text="Status: -"       Margin="0,2,0,0"/>
                        <TextBlock Name="DetailDriver"     Text="Driver: -"       Margin="0,2,0,0"/>
                        <TextBlock Name="DetailVLAN"       Text="VLAN: -"         Margin="0,2,0,0"/>
                        <TextBlock Name="DetailSwitchPort" Text="Switch Port: -"  Margin="0,2,0,0"/>
                    </StackPanel>
                </GroupBox>
            </Border>

            <!-- Activity Log -->
            <Border Name="MainLogBorder" Grid.Row="2" Background="White" CornerRadius="8" Padding="8" HorizontalAlignment="Stretch">
                <GroupBox Header="[ Activity Log ]" Margin="0" BorderThickness="0" FontWeight="Bold">
                    <StackPanel>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,0,4">
                            <Button Name="ClearLogBtn" Content="Clear View" Width="90" Height="22"
                                    Background="#FFCBD5F5" FontSize="11"/>
                        </StackPanel>
                        <TextBox Name="LogTextBox" IsReadOnly="True"
                                 VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto"
                                 TextWrapping="NoWrap"
                                 AcceptsReturn="True"/>


                    </StackPanel>
                </GroupBox>
            </Border>
        </Grid>
        </Grid>

        
    </DockPanel>
</Window>
"@

# --------------------------------
# Load XAML and Wire Events
# --------------------------------
$reader          = New-Object System.Xml.XmlNodeReader $xaml
$script:Window   = [System.Windows.Markup.XamlReader]::Load($reader)
$script:Window.WindowState = 'Maximized'

# Populate server combo

function Refresh-ServerCombo {
    try {
        $serverCombo = $script:Window.FindName("ServerCombo")
        if (-not $serverCombo) { return }

        $serverCombo.Items.Clear()

        foreach ($k in $ServerMap.Keys) {
            # Build a colored label: FQDN normal, friendly label in accent color
            $tb = New-Object System.Windows.Controls.TextBlock
            $tb.FontSize = 13
            $tb.Inlines.Add((New-Object System.Windows.Documents.Run($ServerMap[$k].Name)))
            $tb.Inlines.Add(" | ")

            $labelRun = New-Object System.Windows.Documents.Run($ServerMap[$k].Label)
            $labelRun.Foreground = [System.Windows.Media.Brushes]::DarkBlue
            $tb.Inlines.Add($labelRun)

            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $tb
            $item.Tag     = $ServerMap[$k].Name   # keep FQDN for connect

            [void]$serverCombo.Items.Add($item)
        }

        # Auto-select the first server if none is selected
        if ($serverCombo.Items.Count -gt 0 -and -not $serverCombo.SelectedItem) {
            $serverCombo.SelectedIndex = 0
        }
    } catch {
        # do not break UI because of combo refresh
    }
}


Refresh-ServerCombo

# Bind DataGrid and other controls
$script:printerTable      = $script:Window.FindName("PrinterTable")
$script:printerTable.ItemsSource = $printerCollection

$ctx = $script:printerTable.ContextMenu
if ($ctx) {
    $ctxView = $ctx.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "View Jobs (selected)" }
    if ($ctxView) {
        $ctxView.Add_Click({
            GUI-InspectQueueSelected
        })
    }
    $ctxClear = $ctx.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "Clear Queue (selected)" }
    if ($ctxClear) {
        $ctxClear.Add_Click({
            GUI-ClearQueueSelected
        })
    }
    $ctxTest = $ctx.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "Test Page (selected)" }
    if ($ctxTest) {
        $ctxTest.Add_Click({
            GUI-TestPageSelected
        })
    }
}


$filterBox                = $script:Window.FindName("FilterBox")
$filterBtn                = $script:Window.FindName("FilterBtn")
$script:StatusFilterBox   = $script:Window.FindName("StatusFilterBox")

$connectBtn      = $script:Window.FindName("ConnectButton")
$refreshBtn      = $script:Window.FindName("RefreshButton")
$addPrinterBtn       = $script:Window.FindName("AddPrinterBtn")

$autoRefreshChk      = $script:Window.FindName("AutoRefreshCheck")
$autoRefreshInterval = $script:Window.FindName("AutoRefreshInterval")
$disconnectBtn       = $script:Window.FindName("DisconnectBtn")
$closeBtn            = $script:Window.FindName("CloseBtn")
$statusBar           = $script:Window.FindName("StatusBar")
function Set-ConnectionState {
    param(
        [bool]$Connected
    )

    try {
        if ($refreshBtn)          { $refreshBtn.IsEnabled          = $Connected }
        # Add Server button does not depend on connection
        if ($addPrinterBtn)       { $addPrinterBtn.IsEnabled       = $true }
        if ($autoRefreshChk)      { $autoRefreshChk.IsEnabled      = $Connected }
        if ($autoRefreshInterval) { $autoRefreshInterval.IsEnabled = $Connected }

        # Exclude Resume (formerly QA9). Include QA8 (Select All) in enable/disable list
        $qaNames = "QA1","QA2","QA3","QA4","QA5","QA6","QA7","QA8","QA10","QA12","QA13","QA14","QA15","QA16","QA18"
        foreach ($name in $qaNames) {
            $btn = $script:Window.FindName($name)
            if ($btn) { $btn.IsEnabled = $Connected }
        }
    } catch {
        # Don't break UI if something goes wrong here
    }
}
$script:LogTextBox   = $script:Window.FindName("LogTextBox")
$helpButton          = $script:Window.FindName("HelpButton")
$themeToggleButton   = $script:Window.FindName("ThemeToggleButton")
$toggleLeftPanelBtn  = $script:Window.FindName("ToggleLeftPanelBtn")
$clearLogBtn         = $script:Window.FindName("ClearLogBtn")

if ($clearLogBtn -and $script:LogTextBox) {
    $clearLogBtn.Add_Click({
        $script:LogTextBox.Text = "<Log view cleared>"
        $script:LogTextBox.ScrollToEnd()
    })
}


if ($themeToggleButton) {
    $themeToggleButton.Add_Click({
        if ($script:CurrentTheme -eq "Light") {
            Set-Theme -Mode "Dark"
            $themeToggleButton.Content = "Light Mode"
        } else {
            Set-Theme -Mode "Light"
            $themeToggleButton.Content = "Dark Mode"
        }
    })
    # Initialize theme
    Set-Theme -Mode $script:CurrentTheme
}

if ($toggleLeftPanelBtn) {
    $mainGrid   = $script:Window.FindName("MainGrid")
    $leftBorder = $script:Window.FindName("LeftPanelBorder")
    $script:LeftPanelCollapsed = $false

    $toggleLeftPanelBtn.Add_Click({
        if (-not $mainGrid -or -not $leftBorder) { return }

        if (-not $script:LeftPanelCollapsed) {
            $mainGrid.ColumnDefinitions[0].Width = New-Object System.Windows.GridLength(0, [System.Windows.GridUnitType]::Star)
            $leftBorder.Visibility = "Collapsed"
            $toggleLeftPanelBtn.Content = "Show Left Panel"
            $script:LeftPanelCollapsed = $true
        } else {
            $mainGrid.ColumnDefinitions[0].Width = New-Object System.Windows.GridLength(3, [System.Windows.GridUnitType]::Star)
            $leftBorder.Visibility = "Visible"
            $toggleLeftPanelBtn.Content = "Hide Left Panel"
            $script:LeftPanelCollapsed = $false
        }
    })
}

if ($helpButton) {
    $helpButton.Add_Click({
        # Build a custom WPF window for Help & About
        $helpWin                       = New-Object System.Windows.Window
        $helpWin.Title                 = "Help & About - LEPS PRINTING ADMINISTRATION"
        $helpWin.SizeToContent         = "Manual"
        $helpWin.Width                 = 1000
        $helpWin.Height                = 700
        $helpWin.ResizeMode            = "CanResize"
        $helpWin.WindowStartupLocation = "CenterOwner"
        $helpWin.Owner                 = $script:Window

        $rootBorder            = New-Object System.Windows.Controls.Border
        $rootBorder.Background  = [System.Windows.Media.Brushes]::White
        $rootBorder.Padding     = "15"
        $rootBorder.CornerRadius= "8"

        $panel                  = New-Object System.Windows.Controls.StackPanel
        $panel.Orientation      = "Vertical"
        $panel.Margin           = "0,0,0,10"

        # Title
        $title                  = New-Object System.Windows.Controls.TextBlock
        $title.Text             = "LEPS PRINTING ADMINISTRATION V2.0"
        $title.FontSize         = 18
        $title.FontWeight       = "Bold"
        $title.Foreground       = [System.Windows.Media.Brushes]::DarkBlue
        $panel.Children.Add($title)

        $subtitle               = New-Object System.Windows.Controls.TextBlock
        $subtitle.Text          = "Centralized Printer Monitoring & Management"
        $subtitle.FontSize      = 14
        $subtitle.Foreground    = [System.Windows.Media.Brushes]::DarkSlateGray
        $subtitle.Margin        = "0,2,0,10"
        $panel.Children.Add($subtitle)

        # QUICK START
        $qsHeader               = New-Object System.Windows.Controls.TextBlock
        $qsHeader.Text          = "QUICK START"
        $qsHeader.FontWeight    = "Bold"
        $qsHeader.Margin        = "0,8,0,2"
        $qsHeader.Foreground    = [System.Windows.Media.Brushes]::DarkSlateBlue
        $panel.Children.Add($qsHeader)

        $qsText                 = New-Object System.Windows.Controls.TextBlock
        $qsText.TextWrapping    = "Wrap"
        $qsText.FontSize        = 13
        $qsText.Text            = @"
1. Select a print server from the 'Print Servers' list and click Connect.
2. Enter your administrator credentials when prompted.
3. Wait for printers to load in the main table on the left.
4. Use Search & Filter + Status to quickly find a specific printer or subset.
5. Use the Quick Actions section to manage printers in BULK (multi-selection).
"@
        $panel.Children.Add($qsText)

        # MAIN FEATURES
        $featHeader             = New-Object System.Windows.Controls.TextBlock
        $featHeader.Text        = "MAIN FEATURES"
        $featHeader.FontWeight  = "Bold"
        $featHeader.Margin      = "0,8,0,2"
        $featHeader.Foreground  = [System.Windows.Media.Brushes]::DarkSlateBlue
        $panel.Children.Add($featHeader)

                $featText               = New-Object System.Windows.Controls.TextBlock
        $featText.TextWrapping  = "Wrap"
        $featText.FontSize      = 13
        $featText.Text          = @"
- [Multi-Server Support]:
  Manage several print servers from a single interface.

- [Bulk Printer Actions] (multi-selection):
  Ping, Test Page, Clear Queue on many printers at once.

- [Job Queue Inspector] (View Jobs):
  See Job ID, Document, User, Pages, Status, Submit time and Job Age for selected printers.

- [Dashboard & Status]:
  Colored rows and live counters for Total / Online / Offline, updated on refresh and ping.

- [Server Resource Monitor]:
  Lightweight CPU, RAM and Disk C: usage view for the current print server.

- [Export & Import]:
  Export printer list to CSV or TXT, import previous exports to inspect printer configuration.

- [Activity & Audit Log]:
  On-screen Activity Log plus structured daily log files in the Logs folder for auditing.
"@
        $panel.Children.Add($featText)


        
        # TUTORIAL - BUTTON GUIDE
                $tutHeader               = New-Object System.Windows.Controls.TextBlock
        $tutHeader.Text          = "TUTORIAL - BUTTON GUIDE"
        $tutHeader.FontWeight    = "Bold"
        $tutHeader.Margin        = "0,8,0,2"
        $tutHeader.Foreground    = [System.Windows.Media.Brushes]::DarkSlateBlue
        $panel.Children.Add($tutHeader)

        $tutText                 = New-Object System.Windows.Controls.TextBlock
        $tutText.TextWrapping    = "Wrap"
        $tutText.FontSize        = 13
        $tutText.Text            = @"
[Connection & Server Area]
- [Print Servers]: Choose the target print server (e.g. svltng2xpnj03.leoni.local | G2X [LTN-4]).
- [Connect]: Opens a secure session using your admin credentials.
- [Disconnect]: Closes the session and clears the view.
- [Refresh]: Reloads printers and updates dashboard and resource monitor.
 - [Add Server]: Manually add a new print server to the list for this session.
- [Auto Refresh]: Automatically refreshes the view using the selected interval.

[Quick Actions - Server]
- [Restart Spooler]: Stops then starts the Spooler service on the connected server.
- [Live Status]: Toggles automatic refresh using the selected interval.

[Quick Actions - Printers]
- [Ping]: Reachability test for all selected printers.
- [Test Page]: Sends a Windows test page on each selected printer.
 - [Select All]: Select all printers in the table.
 - [View Jobs]: Opens the Job Queue Inspector with Job Age per print job.
- [Clear Queue]: Deletes all jobs for each selected printer after confirmation.
- [Errors]: Shows only printers with Offline or Error status.
- [List]: Shows a simple text list (Name | IP | Status) of all loaded printers.

[Export & Reporting]
- [Export CSV/TXT]: Export the printer list including status and basic network info.
- [Import Config]: Load a previous export and inspect one or more printers.

Tip:
Use CTRL + Click or SHIFT + Click in the printer table to select multiple printers,
then apply any action (Ping, Test Page, Clear Queue, View Jobs) in bulk.
"@
        $panel.Children.Add($tutText)


# KEYBOARD SHORTCUTS
        $keysHeader             = New-Object System.Windows.Controls.TextBlock
        $keysHeader.Text        = "KEYBOARD SHORTCUTS"
        $keysHeader.FontWeight  = "Bold"
        $keysHeader.Margin      = "0,8,0,2"
        $keysHeader.Foreground  = [System.Windows.Media.Brushes]::DarkSlateBlue
        $panel.Children.Add($keysHeader)

        $keysText               = New-Object System.Windows.Controls.TextBlock
        $keysText.TextWrapping  = "Wrap"
        $keysText.FontSize      = 13
        $keysText.Text          = @"
 - CTRL + R : Full refresh of printers and dashboard.
 - CTRL + A : Select all printers in the table.
 - CTRL + F : Focus the Search / Filter box.
 - F5       : Refresh the Server Resource Monitor.
 - CTRL + L : Toggle Live Status (start / stop auto refresh).
"@
        $panel.Children.Add($keysText)


# TIPS
        $tipsHeader             = New-Object System.Windows.Controls.TextBlock
        $tipsHeader.Text        = "TIPS FOR TROUBLESHOOTING"
        $tipsHeader.FontWeight  = "Bold"
        $tipsHeader.Margin      = "0,8,0,2"
        $tipsHeader.Foreground  = [System.Windows.Media.Brushes]::DarkSlateBlue
        $panel.Children.Add($tipsHeader)

        $tipsText               = New-Object System.Windows.Controls.TextBlock
        $tipsText.TextWrapping  = "Wrap"
        $tipsText.FontSize      = 13
        $tipsText.Text          = @"
- Always connect with an account that has administrative rights on the print server.
- Use Ping and Test Page together to separate network issues from printer or driver issues.
- Enable Live Status while troubleshooting and disable it when not needed.
"@
        $panel.Children.Add($tipsText)

        # ABOUT & DEDICATION
        $aboutHeader            = New-Object System.Windows.Controls.TextBlock
        $aboutHeader.Text       = "ABOUT & DEDICATION"
        $aboutHeader.FontWeight = "Bold"
        $aboutHeader.Margin     = "0,8,0,2"
        $aboutHeader.Foreground = [System.Windows.Media.Brushes]::DarkSlateBlue
        $panel.Children.Add($aboutHeader)

        $aboutText              = New-Object System.Windows.Controls.TextBlock
        $aboutText.TextWrapping = "Wrap"
        $aboutText.FontStyle    = "Italic"
        $aboutText.FontSize     = 16
        $aboutText.Text         = @"
"@
        $panel.Children.Add($aboutText)

        $thanksText             = New-Object System.Windows.Controls.TextBlock
        $thanksText.TextWrapping= "Wrap"
        $thanksText.Margin      = "0,6,0,0"
        $thanksText.FontSize    = 16
        $thanksText.Text        = @"
Special thanks to the LEPS Administration Team and to LEONI for their trust and support.
Together we keep our systems reliable and our users supported.
"@
        $panel.Children.Add($thanksText)

        # Close button
        $btnPanel                       = New-Object System.Windows.Controls.StackPanel
        $btnPanel.Orientation           = "Horizontal"
        $btnPanel.HorizontalAlignment   = "Right"
        $btnPanel.Margin                = "0,12,0,0"

        $closeBtnHelp                  = New-Object System.Windows.Controls.Button
        $closeBtnHelp.Content          = "Close"
        $closeBtnHelp.Width            = 80
        $closeBtnHelp.Margin           = "0,0,0,0"
        $closeBtnHelp.Padding          = "8,3"
        $closeBtnHelp.Background       = [System.Windows.Media.Brushes]::DarkSlateBlue
        $closeBtnHelp.Foreground       = [System.Windows.Media.Brushes]::White
        $closeBtnHelp.FontWeight       = "Bold"
        $closeBtnHelp.Add_Click({ $helpWin.Close() })

        $btnPanel.Children.Add($closeBtnHelp)
        $panel.Children.Add($btnPanel)

        $scrollViewer                  = New-Object System.Windows.Controls.ScrollViewer
        $scrollViewer.VerticalScrollBarVisibility = "Auto"
        $scrollViewer.Content          = $panel

        $rootBorder.Child  = $scrollViewer
        $helpWin.Content   = $rootBorder
        $helpWin.ShowDialog() | Out-Null
    })
}

# Initial log view: start clean for this app session (no old entries from previous runs)
if ($script:LogTextBox) {
    $script:LogTextBox.Text = ""
}
Log ("--- Application session started (version {0}) ---" -f $script:AppVersion) -ActionType "Session" -Result "Info"


# ---- Placeholder behavior for search box ----
$filterBox.Text       = $script:FilterPlaceholder
$filterBox.Foreground = 'Gray'

$filterBox.Add_GotFocus({
    param($s,$e)
    if ($s.Text -eq $script:FilterPlaceholder) {
        $s.Text = ""
        $s.Foreground = 'Black'
    }
})

$filterBox.Add_LostFocus({
    param($s,$e)
    if ([string]::IsNullOrWhiteSpace($s.Text)) {
        $s.Text = $script:FilterPlaceholder
        $s.Foreground = 'Gray'
    }
})

function Get-FilterTextEffective {
    param([string]$raw)
    if ($raw -eq $script:FilterPlaceholder) { return "" }
    return $raw
}

# ---- Context Menu wiring (right-click on the table) ----
$ctxMenu = $script:printerTable.ContextMenu
if ($ctxMenu) {
    $ctxCopyName = $ctxMenu.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "Copy Name" }
    $ctxCopyIP   = $ctxMenu.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "Copy IP" }
    $ctxPing     = $ctxMenu.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "Ping selected" }
    $ctxTest     = $ctxMenu.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "Test Page (selected)" }
    # Pause/Resume context menu items have been removed, so we do not look them up
    # $ctxPause    = $ctxMenu.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "Pause selected" }
    # $ctxResume   = $ctxMenu.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "Resume selected" }
    $ctxClear    = $ctxMenu.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] -and $_.Header -eq "Clear Queue (selected)" }

    if ($ctxCopyName) { $ctxCopyName.Add_Click({ GUI-CopyNameSelected }) }
    if ($ctxCopyIP)   { $ctxCopyIP.Add_Click({ GUI-CopyIPSelected }) }
    if ($ctxPing)     { $ctxPing.Add_Click({ GUI-PingSelected }) }
    if ($ctxTest)     { $ctxTest.Add_Click({ GUI-TestPrintSelected }) }
    # No handlers for Pause/Resume since those actions are no longer available
    if ($ctxClear)    { $ctxClear.Add_Click({ GUI-ClearQueueSelected }) }
}


# Connect / Refresh / Add Printer buttons
$connectBtn.Add_Click({
    # Always re-fetch ServerCombo from the Window
    $serverComboCtrl = $script:Window.FindName("ServerCombo")
    if (-not $serverComboCtrl) {
        Show-Message "Server list control not found in the UI." "Internal Error"
        return
    }

    # Try to use the selected item
    $sel = $serverComboCtrl.SelectedItem

    # If nothing selected but there are items → auto-select first
    if (-not $sel -and $serverComboCtrl.Items.Count -gt 0) {
        $serverComboCtrl.SelectedIndex = 0
        $sel = $serverComboCtrl.SelectedItem
    }

    if (-not $sel) {
        Show-Message "Choose a server first." "No Server"
        return
    }

    # If we used ComboBoxItem with Tag, prefer that
    if ($sel -is [System.Windows.Controls.ComboBoxItem] -and $sel.Tag) {
        $fqdn = [string]$sel.Tag
    } else {
        # Fallback – try to read text
        $fqdn = [string]$sel
        if ($fqdn -like "*|*") {
            $fqdn = $fqdn.Split(" | ")[0]
        }
    }

    # Try to find friendly label for this server (e.g. G2X [LTN-4])
    $friendlyLabel = $null
    foreach ($k in $ServerMap.Keys) {
        if ($ServerMap[$k].Name -eq $fqdn) {
            $friendlyLabel = $ServerMap[$k].Label
            break
        }
    }
    if (-not $friendlyLabel) { $friendlyLabel = $fqdn }

    $script:CurrentServerLabel = $friendlyLabel
    $script:CurrentServerTag   = Get-ServerTagFromLabel -label $friendlyLabel


    if (Connect-ToServer $fqdn) {
        # Start a fresh Activity Log view for this server
        if ($script:LogTextBox) {
            $script:LogTextBox.Text = ""
        }
        Log ("--- New session connected to {0} ---" -f $fqdn) -ActionType "Session" -Target $fqdn -Result "Info"
        # Lightweight first load for faster connection experience
        if ($statusBar) {
            $statusBar.Text = "Loading printers from $fqdn..."
        }
        Refresh-PrintersTableAsync -LightWeight -SuppressErrorPopup
        Start-SessionWatchdog
        Set-ConnectionState -Connected:$true
        $statusBar.Text = "Connected to $fqdn successfully."

        [System.Windows.MessageBox]::Show(
            "Connected successfully to server: $friendlyLabel",
            "Connection Successful",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
    } else {
        $statusBar.Text = "Connection failed."
    }
})

$addPrinterBtn.Add_Click({
    # Add Server does not require an active CimSession. We allow server addition anytime.
    QuickAction '13'
})






$refreshBtn.Add_Click({
    if (-not $Global:CimSession) {
        Show-Message "Connect to a server first." "No Connection"
        return
    }
    try {
        Refresh-PrintersTableAsync
    } catch {
        Log "Manual refresh error: $($_.Exception.Message)" -ActionType "Refresh" -Result "Error" -ErrorMessage $_.Exception.Message
    }
})

# Inline search / filter events
$filterBtn.Add_Click({
    $statusItem = $script:StatusFilterBox.SelectedItem
    $statusText = "All"
    if ($statusItem -and $statusItem.Content) { $statusText = $statusItem.Content.ToString() }

    $effective = Get-FilterTextEffective -raw $filterBox.Text
    Apply-PrinterFilter -filterText $effective -statusFilter $statusText
})

$filterBox.Add_KeyDown({
    param($sender,$e)
    if ($e.Key -eq "Enter") {
        $statusItem = $script:StatusFilterBox.SelectedItem
        $statusText = "All"
        if ($statusItem -and $statusItem.Content) { $statusText = $statusItem.Content.ToString() }

        $effective = Get-FilterTextEffective -raw $sender.Text
        Apply-PrinterFilter -filterText $effective -statusFilter $statusText
    }
})

$script:StatusFilterBox.Add_SelectionChanged({
    $statusItem = $script:StatusFilterBox.SelectedItem
    $statusText = "All"
    if ($statusItem -and $statusItem.Content) { $statusText = $statusItem.Content.ToString() }

    $effective = Get-FilterTextEffective -raw $filterBox.Text
    Apply-PrinterFilter -filterText $effective -statusFilter $statusText
})

# Selection changed
$hasJobsCheck = $script:Window.FindName("HasJobsCheckBox")
if ($hasJobsCheck) {
    $hasJobsCheck.Add_Checked({
        $script:HasJobsOnly = $true
        $statusItem = $script:StatusFilterBox.SelectedItem
        $statusText = "All"
        if ($statusItem -and $statusItem.Content) { $statusText = $statusItem.Content.ToString() }
        $effective = Get-FilterTextEffective -raw $filterBox.Text
        Apply-PrinterFilter -filterText $effective -statusFilter $statusText
    })
    $hasJobsCheck.Add_Unchecked({
        $script:HasJobsOnly = $false
        $statusItem = $script:StatusFilterBox.SelectedItem
        $statusText = "All"
        if ($statusItem -and $statusItem.Content) { $statusText = $statusItem.Content.ToString() }
        $effective = Get-FilterTextEffective -raw $filterBox.Text
        Apply-PrinterFilter -filterText $effective -statusFilter $statusText
    })
}
# → update detail panel
$script:printerTable.Add_SelectionChanged({
    Update-SelectedPrinterDetails
})
# Quick Action dispatcher
function QuickAction {
    param([string]$id)

    switch ($id) {
        '1' { # List
            if ($printerCollection.Count -eq 0) {
                Show-Message "No printers loaded." "List"
                return
            }
            $text = ($printerCollection | Sort-Object Name |
                ForEach-Object { "$($_.Name) - $($_.IP) - $($_.PrinterStatus)" }) -join "`n"
            [System.Windows.MessageBox]::Show($text, "Printer List") | Out-Null
        }
        '2' { # Errors (Offline + Error only)
            $err = $printerCollection | Where-Object {
                $_.PrinterStatus -in @("Offline","Error")
            }

            if ($err -and $err.Count -gt 0) {
                $text = ($err | Sort-Object Name |
                    ForEach-Object { "$($_.Name) | $($_.IP) | $($_.PrinterStatus)" }) -join "`n"

                [System.Windows.MessageBox]::Show($text, "Offline / Error Printers") | Out-Null
                Log "Errors quick view opened. Count: $($err.Count)" -ActionType "ErrorsView" -Result "Success"
            } else {
                Show-Message "No Offline or Error printers detected." "Errors"
                Log "Errors quick view opened. None found." -ActionType "ErrorsView" -Result "Success"
            }
        }
        '3' { GUI-RestartSpooler }
        '4' { GUI-ClearQueueSelected }
        '5' { # Live Status toggle
            if (-not $Global:LiveStatusRunning) { Start-LiveStatus } else { Stop-LiveStatus }
            $autoRefreshChk.IsChecked = $Global:LiveStatusRunning
        }
        '6' { GUI-PingSelected }
        '7' { GUI-TestPrintSelected }
        '8' { # Select All printers
            GUI-SelectAllPrinters
        }
        '10'{ GUI-ExportSelection }
        '12'{ GUI-ImportConfig }
        '13'{ GUI-AddServer }
        '14'{ GUI-InspectQueueSelected }
        '15'{ GUI-SelfHealingServer }
        '16'{ GUI-ShowServerPerformance }
        '17'{ GUI-DeployPrinterWizard }
        '18'{ GUI-ImportServers }
    }
}



# Quick Actions buttons
$script:Window.FindName("QA1").Add_Click({ QuickAction '1' })
$script:Window.FindName("QA2").Add_Click({ QuickAction '2' })
$script:Window.FindName("QA3").Add_Click({ QuickAction '3' })
$script:Window.FindName("QA4").Add_Click({ QuickAction '4' })
$script:Window.FindName("QA5").Add_Click({ QuickAction '5' })
$script:Window.FindName("QA6").Add_Click({ QuickAction '6' })
$script:Window.FindName("QA7").Add_Click({ QuickAction '7' })
$script:Window.FindName("QA8").Add_Click({ QuickAction '8' })
$script:Window.FindName("QA10").Add_Click({ QuickAction '10' })
$script:Window.FindName("QA12").Add_Click({ QuickAction '12' })
$script:Window.FindName("QA13").Add_Click({ QuickAction '14' })
$script:Window.FindName("QA14").Add_Click({ QuickAction '15' })
$script:Window.FindName("QA15").Add_Click({ QuickAction '16' })
$script:Window.FindName("QA16").Add_Click({ QuickAction '17' })
$script:Window.FindName("QA18").Add_Click({ QuickAction '18' })

$script:Window.FindName("RefreshMonitorBtn").Add_Click({ Update-ResourceMonitor })
# Auto-refresh checkbox
$autoRefreshChk.Add_Checked({
    if (-not $Global:LiveStatusRunning) { Start-LiveStatus }
})
$autoRefreshChk.Add_Unchecked({
    if ($Global:LiveStatusRunning) { Stop-LiveStatus }
})
$disconnectBtn.Add_Click({
    # If we are not connected and nothing is loaded, show gentle info
    if (-not $Global:CimSession -and $printerCollection.Count -eq 0) {
        Show-Message "You are not connected to any server yet." "No Connection"
        return
    }

    if ($Global:LiveStatusRunning) { Stop-LiveStatus }
    Stop-SessionWatchdog
    Disconnect-Server
    $printerCollection.Clear()
    Update-Dashboard -Window $script:Window
    Update-SelectedPrinterDetails
    Update-ResourceMonitor
    Set-ConnectionState -Connected:$false
    $statusBar.Text = "Disconnected successfully."
})

$closeBtn.Add_Click({
    if ($Global:LiveStatusRunning) { Stop-LiveStatus }
    Stop-SessionWatchdog
    Disconnect-Server
    $script:Window.Close()
})

# Keyboard shortcuts (global)
try {
    $refreshCmd = New-Object System.Windows.Input.RoutedCommand
    $script:Window.CommandBindings.Add( (New-Object System.Windows.Input.CommandBinding($refreshCmd, {
        Refresh-PrintersTableAsync
    })) )
    $script:Window.InputBindings.Add( (New-Object System.Windows.Input.KeyBinding($refreshCmd,
        [System.Windows.Input.Key]::R,
        [System.Windows.Input.ModifierKeys]::Control)) )

    $filterCmd = New-Object System.Windows.Input.RoutedCommand
    $script:Window.CommandBindings.Add( (New-Object System.Windows.Input.CommandBinding($filterCmd, {
        $fb = $script:Window.FindName("FilterBox")
        if ($fb) { $fb.Focus(); $fb.SelectAll() }
    })) )
    $script:Window.InputBindings.Add( (New-Object System.Windows.Input.KeyBinding($filterCmd,
        [System.Windows.Input.Key]::F,
        [System.Windows.Input.ModifierKeys]::Control)) )

    $monitorCmd = New-Object System.Windows.Input.RoutedCommand
    $script:Window.CommandBindings.Add( (New-Object System.Windows.Input.CommandBinding($monitorCmd, {
        Update-ResourceMonitor
    })) )
    $script:Window.InputBindings.Add( (New-Object System.Windows.Input.KeyBinding($monitorCmd,
        [System.Windows.Input.Key]::F5,
        [System.Windows.Input.ModifierKeys]::None)) )

    $liveToggleCmd = New-Object System.Windows.Input.RoutedCommand
    $script:Window.CommandBindings.Add( (New-Object System.Windows.Input.CommandBinding($liveToggleCmd, {
        if ($Global:LiveStatusRunning) {
            Stop-LiveStatus
        } else {
            Start-LiveStatus
        }
    })) )
    $script:Window.InputBindings.Add( (New-Object System.Windows.Input.KeyBinding($liveToggleCmd,
        [System.Windows.Input.Key]::L,
        [System.Windows.Input.ModifierKeys]::Control)) )

    # Full refresh via Ctrl+R
    $refreshCmd = New-Object System.Windows.Input.RoutedCommand
    $script:Window.CommandBindings.Add( (New-Object System.Windows.Input.CommandBinding($refreshCmd, {
        if ($Global:CimSession) {
            try { Refresh-PrintersTableAsync } catch { }
        } else {
            Show-Message "Connect to a server first." "No Connection"
        }
    })) )
    $script:Window.InputBindings.Add( (New-Object System.Windows.Input.KeyBinding($refreshCmd,
        [System.Windows.Input.Key]::R,
        [System.Windows.Input.ModifierKeys]::Control)) )

    # Select all printers via Ctrl+A
    $selectAllCmd = New-Object System.Windows.Input.RoutedCommand
    $script:Window.CommandBindings.Add( (New-Object System.Windows.Input.CommandBinding($selectAllCmd, {
        GUI-SelectAllPrinters
    })) )
    $script:Window.InputBindings.Add( (New-Object System.Windows.Input.KeyBinding($selectAllCmd,
        [System.Windows.Input.Key]::A,
        [System.Windows.Input.ModifierKeys]::Control)) )
} catch {
    # ignore shortcut wiring errors
}

# Start with disconnected state
Set-ConnectionState -Connected:$false
Update-ResourceMonitor
Update-Dashboard -Window $script:Window
Update-SelectedPrinterDetails

$script:Window.ShowDialog() | Out-Null

# SPO/PnP Persistent Session Broker
# Maintains a live PowerShell session with either SPO or PnP module loaded.
# IMPORTANT: SPO and PnP modules CANNOT coexist in the same session (CSOM assembly conflicts).
# Use -Mode SPO for tenant admin/multi-geo, -Mode PnP for site-level/list/page operations.
#
# SPO mode (default):
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File SPO-SessionBroker.ps1 -AdminUrl "https://tenant-admin.sharepoint.com/"
# PnP mode:
#   pwsh -NoProfile -ExecutionPolicy Bypass -File SPO-SessionBroker.ps1 -Mode PnP -AdminUrl "https://tenant-admin.sharepoint.com/" -PnPClientId "<app-id>"

param(
    [ValidateSet("SPO", "PnP")]
    [string]$Mode = "SPO",
    [string]$AdminUrl,
    [string]$PnPClientId   # Required for PnP mode: Entra ID app registration client ID
)

# Session directory varies by mode to allow both brokers to run simultaneously
$sessionSuffix = if ($Mode -eq "PnP") { ".pnp-session" } else { ".spo-session" }
$sessionDir = Join-Path $env:USERPROFILE $sessionSuffix
$cmdPath    = Join-Path $sessionDir "command.ps1"
$outPath    = Join-Path $sessionDir "output.txt"
$readyPath  = Join-Path $sessionDir "ready.marker"
$pidPath    = Join-Path $sessionDir "broker.pid"
$stopPath   = Join-Path $sessionDir "shutdown.marker"

# Clean up any previous session
if (Test-Path $sessionDir) { Remove-Item $sessionDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $sessionDir -Force | Out-Null

# Save PID so we can check if broker is alive
Set-Content -Path $pidPath -Value $PID

Write-Host "$Mode Session Broker starting..." -ForegroundColor Cyan
Write-Host "Session directory: $sessionDir" -ForegroundColor Gray
Write-Host "Broker PID: $PID" -ForegroundColor Gray

$connected = $false
$moduleVersion = "N/A"

if ($Mode -eq "PnP") {
    # --- PnP Mode: Load PnP.PowerShell only (requires pwsh / PowerShell 7+) ---
    try {
        $pnpModule = Get-Module -Name PnP.PowerShell -ListAvailable -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $pnpModule) {
            Write-Host "PnP.PowerShell NOT INSTALLED. Run: Install-Module PnP.PowerShell -Scope CurrentUser" -ForegroundColor Red
            exit 1
        }
        Import-Module PnP.PowerShell -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop
        $moduleVersion = $pnpModule.Version.ToString()
        Write-Host "PnP.PowerShell $moduleVersion loaded." -ForegroundColor Green
    } catch {
        Write-Host "PnP.PowerShell failed to load: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    # Connect PnP
    if ($AdminUrl) {
        $tenantRoot = $AdminUrl -replace '-admin\.sharepoint\.com.*', '.sharepoint.com'
        Write-Host "Connecting PnP to $tenantRoot ..." -ForegroundColor Cyan
        try {
            if ($PnPClientId) {
                Connect-PnPOnline -Url $tenantRoot -Interactive -ClientId $PnPClientId -ErrorAction Stop
            } else {
                Connect-PnPOnline -Url $tenantRoot -Interactive -ErrorAction Stop
            }
            $connected = $true
            Write-Host "PnP connected successfully." -ForegroundColor Green
        } catch {
            Write-Host "PnP connection failed: $($_.Exception.Message)" -ForegroundColor Red
            if ($_.Exception.Message -match "consent|AADSTS65001|AADSTS50011|admin_consent|permission|client_id") {
                Write-Host "" -ForegroundColor Red
                Write-Host "=== PnP TENANT SETUP REQUIRED ===" -ForegroundColor Red
                Write-Host "PnP PowerShell needs an Entra ID app registration with admin consent." -ForegroundColor Yellow
                Write-Host "Steps:" -ForegroundColor Yellow
                Write-Host "  1. Register app: Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'PnP Admin' -Tenant <tenant>.onmicrosoft.com -Interactive" -ForegroundColor Cyan
                Write-Host "  2. Use the returned Client ID with: -PnPClientId <client-id>" -ForegroundColor Cyan
            }
        }
    }
} else {
    # --- SPO Mode: Load Microsoft.Online.SharePoint.PowerShell only ---
    # Auto-detect module path: check standard locations for the SPO module
    $spoModuleFound = Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable -ErrorAction SilentlyContinue
    if (-not $spoModuleFound) {
        # Search common non-standard locations (e.g., OneDrive-synced Documents folder)
        $candidatePaths = @(
            "$env:USERPROFILE\Documents\PowerShell\Modules",
            "$env:USERPROFILE\Documents\WindowsPowerShell\Modules"
        )
        # Also check OneDrive paths dynamically
        Get-ChildItem "$env:USERPROFILE\OneDrive*" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $candidatePaths += "$($_.FullName)\Documents\PowerShell\Modules"
            $candidatePaths += "$($_.FullName)\Documents\WindowsPowerShell\Modules"
        }
        foreach ($p in $candidatePaths) {
            if (Test-Path "$p\Microsoft.Online.SharePoint.PowerShell") {
                $env:PSModulePath = "$p;$env:PSModulePath"
                break
            }
        }
    }
    try {
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop
        $spoMod = Get-Module Microsoft.Online.SharePoint.PowerShell
        $moduleVersion = $spoMod.Version.ToString()
        Write-Host "SPO module $moduleVersion loaded." -ForegroundColor Green
    } catch {
        Write-Host "SPO module failed to load: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    # Connect SPO
    if ($AdminUrl) {
        Write-Host "Connecting SPO to $AdminUrl ..." -ForegroundColor Cyan
        try {
            Get-SPOTenant -ErrorAction Stop | Out-Null
            $connected = $true
            Write-Host "SPO already connected." -ForegroundColor Green
        } catch {
            try {
                Connect-SPOService -Url $AdminUrl -ErrorAction Stop
                $connected = $true
                Write-Host "SPO connected successfully." -ForegroundColor Green
            } catch {
                Write-Host "SPO connection failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "No AdminUrl provided. Use Connect-SPOService in a command to connect." -ForegroundColor Yellow
    }
}

# --- Signal ready ---
$readyInfo = @{
    Timestamp     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Mode          = $Mode
    Connected     = $connected
    ModuleVersion = $moduleVersion
    Shell         = "PowerShell $($PSVersionTable.PSVersion)"
}
Set-Content -Path $readyPath -Value ($readyInfo | ConvertTo-Json -Compress)
Write-Host "$Mode Session broker READY. Waiting for commands at: $cmdPath" -ForegroundColor Green
Write-Host "---" -ForegroundColor Gray

# --- Command processing loop ---
while ($true) {
    # Check for shutdown signal
    if (Test-Path $stopPath) {
        Write-Host "Shutdown requested. Cleaning up..." -ForegroundColor Yellow
        if ($Mode -eq "SPO") { Disconnect-SPOService -ErrorAction SilentlyContinue }
        else { Disconnect-PnPOnline -ErrorAction SilentlyContinue }
        Remove-Item $sessionDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "$Mode Session broker stopped." -ForegroundColor Green
        exit 0
    }

    # Check for command file
    if (Test-Path $cmdPath) {
        $timestamp = Get-Date -Format 'HH:mm:ss'
        Write-Host "[$timestamp] Executing command..." -ForegroundColor Cyan

        # Clear previous output
        if (Test-Path $outPath) { Remove-Item $outPath -Force }

        try {
            # Read command content and execute via Invoke-Expression in the current scope
            # This preserves connection state (module-scoped) across commands
            $scriptContent = Get-Content -Path $cmdPath -Raw
            $result = (Invoke-Expression $scriptContent) *>&1 | Out-String
            Set-Content -Path $outPath -Value $result -Encoding UTF8
            Write-Host "[$timestamp] Command completed successfully." -ForegroundColor Green
        } catch {
            $errorMsg = "ERROR: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
            Set-Content -Path $outPath -Value $errorMsg -Encoding UTF8
            Write-Host "[$timestamp] Command failed: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Remove command file to signal: "output is ready"
        Remove-Item $cmdPath -Force
    }

    Start-Sleep -Milliseconds 500
}

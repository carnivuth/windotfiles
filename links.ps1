# Create-WebShortcuts.ps1
# Reads an array of links from a JSON file and creates .lnk shortcut files
# that open each URL in a new Firefox tab.
#
# Usage:
#   .\links.ps1 -JsonFile "links.json"
#   .\links.ps1 -JsonFile "links.json" -OutputFolder "C:\Custom\Path"
#   .\links.ps1 -JsonFile "links.json" -FirefoxExe "D:\Firefox\firefox.exe"

param (
    [Parameter(Mandatory = $false, HelpMessage = "Path to the JSON file containing the links.")]
    [string]$JsonFile = ".\sites.json",

    [Parameter(Mandatory = $false, HelpMessage = "Destination folder for shortcuts. Defaults to the user's Start Menu Programs folder.")]
    [string]$OutputFolder = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Firefox Web Apps",

    [Parameter(Mandatory = $false, HelpMessage = "Path to the Firefox executable. Auto-detected if omitted.")]
    [string]$FirefoxExe = ""
)

# ── Resolve Firefox path ──────────────────────────────────────────────────────
if (-not $FirefoxExe) {
    $candidates = @(
        "$env:ProgramFiles\Mozilla Firefox\firefox.exe",
        "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe",
        "$env:LocalAppData\Mozilla Firefox\firefox.exe"
    )
    $FirefoxExe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}

$FJsonFIle= ".\sites.json"

if (-not $FirefoxExe -or -not (Test-Path $FirefoxExe)) {
    Write-Error "Firefox executable not found. Use -FirefoxExe to specify the path manually."
    exit 1
}

Write-Host "Using Firefox at: $FirefoxExe"
Write-Host ""

# ── Validate JSON file exists ─────────────────────────────────────────────────
if (-not (Test-Path -Path $JsonFile)) {
    Write-Error "JSON file not found: $JsonFile"
    exit 1
}

# ── Parse JSON ────────────────────────────────────────────────────────────────
try {
    $links = Get-Content -Path $JsonFile -Raw | ConvertFrom-Json
} catch {
    Write-Error "Failed to parse JSON file: $_"
    exit 1
}

if ($links.Count -eq 0) {
    Write-Warning "No entries found in the JSON file."
    exit 0
}

# ── Ensure output folder exists ───────────────────────────────────────────────
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Host "Created folder: $OutputFolder"
}

# ── Create shortcuts ──────────────────────────────────────────────────────────
$created = 0
$skipped = 0
$failed  = 0

foreach ($entry in $links) {
    # Validate required fields
    if (-not $entry.site -or -not $entry.link) {
        Write-Warning "Skipping entry with missing 'site' or 'link' field: $($entry | ConvertTo-Json -Compress)"
        $skipped++
        continue
    }

    $siteName = $entry.site.Trim()
    $url      = $entry.link.Trim()

    # Sanitise filename — strip characters illegal in Windows file names
    $safeName     = $siteName -replace '[\\/:*?"<>|]', '_'
    $shortcutPath = Join-Path -Path $OutputFolder -ChildPath "$safeName.lnk"

    try {
        $shell    = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)

        $shortcut.TargetPath       = $FirefoxExe
        $shortcut.Arguments        = "-new-tab `"$url`""
        $shortcut.Description      = $siteName
        $shortcut.WorkingDirectory = Split-Path $FirefoxExe

        $shortcut.Save()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)    | Out-Null

        Write-Host "  [OK] $siteName  ->  $shortcutPath" -ForegroundColor Green
        $created++
    } catch {
        Write-Warning "  [FAIL] Could not create shortcut for '$siteName': $_"
        $failed++
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Done. Created: $created  |  Skipped: $skipped  |  Failed: $failed" -ForegroundColor Cyan
Write-Host "Shortcuts saved to: $OutputFolder"

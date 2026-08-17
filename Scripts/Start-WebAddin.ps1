#Requires -Version 5.1

<#
.SYNOPSIS
    Sideloads an ExploreTheWorld Office web add-in, starts its Blazor server, and launches the Office app.
.DESCRIPTION
    Node-free equivalent of the `npm run start-local` script (office-addin-debugging start).
    Performs the same four steps without requiring Node.js:

      1. Verifies the ASP.NET Core HTTPS development certificate is trusted
         (Office refuses to load a task pane from an untrusted https origin).
      2. Sideloads Assets\manifest.local.xml by writing it to the Office developer
         registry key, HKCU\Software\Microsoft\Office\16.0\WEF\Developer. The value
         name is the manifest <Id> GUID and the data is the full manifest path.
      3. Starts the add-in's Blazor server on its https/http port pair and waits
         until the port accepts connections (reuses an already-running server).
      4. Opens Word / Excel / PowerPoint via COM with Visible = true and creates a
         blank document, so the "ETW (Web)" ribbon tab is ready to click.

    Run with -Unregister to remove the sideload registry entry again.

    If Node.js is installed, `npm run start-local` in the add-in project folder
    remains a supported alternative - see docs/project-templates.md.
.PARAMETER OfficeApp
    Which add-in to launch: Word, Excel, or PowerPoint.
.PARAMETER Unregister
    Removes this add-in's sideload registry entry and exits. Does not stop the server.
.PARAMETER NoLaunch
    Sideloads and starts the server but does not open the Office application.
.PARAMETER Configuration
    Build configuration to run: Debug (default) or Release.
.PARAMETER TimeoutSeconds
    How long to wait for the add-in server to start listening. Default 60.
.EXAMPLE
    .\Scripts\Start-WebAddin.ps1 -OfficeApp Word
.EXAMPLE
    .\Scripts\Start-WebAddin.ps1 -OfficeApp PowerPoint -Configuration Release
.EXAMPLE
    .\Scripts\Start-WebAddin.ps1 -OfficeApp Excel -Unregister
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Word', 'Excel', 'PowerPoint')]
    [string] $OfficeApp,

    [switch] $Unregister,

    [switch] $NoLaunch,

    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Debug',

    [int] $TimeoutSeconds = 60
)

$ErrorActionPreference = 'Stop'

$RepoRoot = Split-Path -Parent $PSScriptRoot

# Port pairs match Properties\launchSettings.json and OfficeWebAddinServerFixture.cs - keep all three in sync.
$AddIns = @{
    'Word'       = @{ Folder = 'AL.MsOfficeWordBlazorWebAddIn';       HttpsPort = 7100; HttpPort = 5100; ProgId = 'Word.Application' }
    'Excel'      = @{ Folder = 'AL.MsOfficeExcelBlazorWebAddIn';      HttpsPort = 7101; HttpPort = 5101; ProgId = 'Excel.Application' }
    'PowerPoint' = @{ Folder = 'AL.MsOfficePowerPointBlazorWebAddIn'; HttpsPort = 7102; HttpPort = 5102; ProgId = 'PowerPoint.Application' }
}

$AddIn        = $AddIns[$OfficeApp]
$ProjectDir   = Join-Path $RepoRoot "src\$($AddIn.Folder)"
$ProjectPath  = Join-Path $ProjectDir "ExploreTheWorld.$($AddIn.Folder).csproj"
$ManifestPath = Join-Path $ProjectDir 'Assets\manifest.local.xml'
$ExePath      = Join-Path $ProjectDir "bin\$Configuration\net10.0\ExploreTheWorld.$($AddIn.Folder).exe"
$BaseUrl      = "https://localhost:$($AddIn.HttpsPort)"
$DeveloperKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'

function Write-Step {
    param([string] $Number, [string] $Text)
    Write-Host (" [{0}] {1}" -f $Number, $Text) -ForegroundColor Yellow
}

function Test-PortOpen {
    param([int] $Port)

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect('127.0.0.1', $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne(500)) {
            return $false
        }
        $client.EndConnect($async)
        return $true
    }
    catch {
        return $false
    }
    finally {
        $client.Close()
    }
}

if (-not (Test-Path -LiteralPath $ManifestPath)) {
    throw "Manifest not found: $ManifestPath"
}

[xml] $Manifest = Get-Content -LiteralPath $ManifestPath -Raw
$AddInId = $Manifest.OfficeApp.Id
if ([string]::IsNullOrWhiteSpace($AddInId)) {
    throw "Could not read <Id> from $ManifestPath"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " ExploreTheWorld - $OfficeApp Web Add-in" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# --- Unregister and exit -----------------------------------------------------
if ($Unregister) {
    if (Test-Path -LiteralPath $DeveloperKey) {
        $Values = Get-ItemProperty -Path $DeveloperKey
        if ($Values.PSObject.Properties.Name -contains $AddInId) {
            Remove-ItemProperty -Path $DeveloperKey -Name $AddInId
            Write-Host " Removed sideload entry $AddInId" -ForegroundColor Green
        }
        else {
            Write-Host " No sideload entry found for $AddInId - nothing to do." -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host " Developer key does not exist - nothing to do." -ForegroundColor DarkGray
    }
    Write-Host ""
    exit 0
}

# --- 1. HTTPS development certificate ----------------------------------------
Write-Step '1/4' 'Checking ASP.NET Core HTTPS development certificate...'

dotnet dev-certs https --check --trust | Out-Null
if ($LASTEXITCODE -eq 0) {
    Write-Host "       Trusted certificate found." -ForegroundColor Green
}
else {
    Write-Warning "No trusted HTTPS development certificate. Office will refuse to load the task pane from $BaseUrl."
    Write-Warning "Run this once, then re-run this script:  dotnet dev-certs https --trust"
}

# --- 2. Sideload the manifest ------------------------------------------------
Write-Step '2/4' 'Sideloading manifest.local.xml...'

if (-not (Test-Path -LiteralPath $DeveloperKey)) {
    New-Item -Path $DeveloperKey -Force | Out-Null
}
New-ItemProperty -Path $DeveloperKey -Name $AddInId -Value $ManifestPath -PropertyType String -Force | Out-Null

Write-Host "       $AddInId" -ForegroundColor Green
Write-Host "       -> $ManifestPath" -ForegroundColor DarkGray

# --- 3. Start the add-in server ----------------------------------------------
Write-Step '3/4' "Starting add-in server on $BaseUrl ..."

if (Test-PortOpen $AddIn.HttpsPort) {
    Write-Host "       Already listening - reusing the running server." -ForegroundColor Green
}
else {
    if (-not (Test-Path -LiteralPath $ExePath)) {
        Write-Host "       $Configuration build not found - building..." -ForegroundColor DarkGray
        dotnet build $ProjectPath -c $Configuration --nologo -v quiet
        if ($LASTEXITCODE -ne 0) {
            throw "Build failed for $ProjectPath"
        }
    }

    # Start-Process inherits the current process environment; restore afterwards.
    $PreviousUrls        = $env:ASPNETCORE_URLS
    $PreviousEnvironment = $env:ASPNETCORE_ENVIRONMENT
    try {
        $env:ASPNETCORE_URLS        = "https://localhost:$($AddIn.HttpsPort);http://localhost:$($AddIn.HttpPort)"
        $env:ASPNETCORE_ENVIRONMENT = 'Development'
        $ServerProcess = Start-Process -FilePath $ExePath -WorkingDirectory $ProjectDir -PassThru
    }
    finally {
        $env:ASPNETCORE_URLS        = $PreviousUrls
        $env:ASPNETCORE_ENVIRONMENT = $PreviousEnvironment
    }

    $Deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $Ready    = $false
    while ((Get-Date) -lt $Deadline) {
        if ($ServerProcess.HasExited) {
            throw "Add-in server exited with code $($ServerProcess.ExitCode) before it started listening."
        }
        if (Test-PortOpen $AddIn.HttpsPort) {
            $Ready = $true
            break
        }
        Start-Sleep -Milliseconds 500
    }

    if (-not $Ready) {
        throw "Add-in server did not start listening on $BaseUrl within $TimeoutSeconds seconds."
    }

    Write-Host "       Ready (PID $($ServerProcess.Id))." -ForegroundColor Green
    Write-Host "       Stop it later with:  Stop-Process -Id $($ServerProcess.Id)" -ForegroundColor DarkGray
}

# --- 4. Launch the Office application ----------------------------------------
if ($NoLaunch) {
    Write-Step '4/4' 'Skipping Office launch (-NoLaunch).'
}
else {
    Write-Step '4/4' "Launching $OfficeApp ..."

    $Application = New-Object -ComObject $AddIn.ProgId

    if ($OfficeApp -eq 'PowerPoint') {
        # PowerPoint rejects Visible = msoFalse; -1 is msoTrue. Add the presentation
        # WithWindow = msoTrue so the window surfaces.
        $Application.Visible = -1
        $Application.Presentations.Add(-1) | Out-Null
    }
    else {
        $Application.Visible = $true
        if ($OfficeApp -eq 'Word') {
            $Application.Documents.Add() | Out-Null
        }
        else {
            $Application.Workbooks.Add() | Out-Null
        }
    }

    Write-Host "       $OfficeApp is open and visible." -ForegroundColor Green
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Ready - open the 'ETW (Web)' ribbon tab" -ForegroundColor Cyan
Write-Host " Task pane: $BaseUrl" -ForegroundColor Cyan
Write-Host " Undo sideload: Scripts\Start-WebAddin.ps1 -OfficeApp $OfficeApp -Unregister" -ForegroundColor DarkGray
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

exit 0

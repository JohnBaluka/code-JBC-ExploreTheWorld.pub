#Requires -Version 7.0
<#
.SYNOPSIS
    Runs all ExploreTheWorld test projects in the prescribed order.
.DESCRIPTION
    Executes dotnet test on each test project sequentially, collects pass/fail results,
    prints a summary, and exits with code 1 if any project failed.
#>

$ErrorActionPreference = 'Continue'

$RepoRoot = Split-Path -Parent $PSScriptRoot

$TestProjects = @(
    'src\Tests\UnitTests\ExploreTheWorld.UnitTests.csproj'
    'src\Tests\UnitTests._netF\ExploreTheWorld.UnitTests._netF.csproj'
    'src\Tests\IntegrationTests\ExploreTheWorld.IntegrationTests.csproj'
    'src\Tests\IntegrationTests._netF\ExploreTheWorld.IntegrationTests._netF.csproj'
    'src\Tests\SqliteDbTests._netF\ExploreTheWorld.SqliteDbTests._netF.csproj'
    'src\Tests\OpenXmlLibTests\ExploreTheWorld.OpenXmlLibTests.csproj'
    'src\Tests\OpenXmlLibTests._netF\ExploreTheWorld.OpenXmlLibTests._netF.csproj'
    'src\Tests\RazorTests\ExploreTheWorld.RazorTests.csproj'
    'src\Tests\WebAppTests\ExploreTheWorld.WebAppTests.csproj'
    'src\Tests\OqtaneTests\ExploreTheWorld.OqtaneTests.csproj'
    'src\Tests\WinFormAppTests\ExploreTheWorld.WinFormAppTests.csproj'
    'src\Tests\WinFormAppTests._netF\ExploreTheWorld.WinFormAppTests._netF.csproj'
    'src\Tests\MauiAppTests\ExploreTheWorld.MauiAppTests.csproj'
    'src\Tests\OfficeAddinTests\ExploreTheWorld.OfficeAddinTests.csproj'
    'src\Tests\OfficeAddinTests._netF\ExploreTheWorld.OfficeAddinTests._netF.csproj'
    'src\Tests\OfficeWebAddinTests\ExploreTheWorld.OfficeWebAddinTests.csproj'
    'src\Tests\AccessDbTests\ExploreTheWorld.AccessDbTests.csproj'
)

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
$StartTime = Get-Date

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " ExploreTheWorld — Run All Tests" -ForegroundColor Cyan
Write-Host " Started: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

foreach ($RelPath in $TestProjects) {
    $ProjectPath = Join-Path $RepoRoot $RelPath
    $ProjectName = [System.IO.Path]::GetFileNameWithoutExtension($RelPath)

    Write-Host ""
    Write-Host "----------------------------------------" -ForegroundColor DarkGray
    Write-Host " Running: $ProjectName" -ForegroundColor Yellow
    Write-Host "----------------------------------------" -ForegroundColor DarkGray

    $ProjectStart = Get-Date
    dotnet test $ProjectPath --no-build 2>&1
    $ExitCode = $LASTEXITCODE

    if ($ExitCode -ne 0) {
        # Retry without --no-build in case the project hasn't been built yet
        dotnet test $ProjectPath 2>&1
        $ExitCode = $LASTEXITCODE
    }

    $Elapsed = (Get-Date) - $ProjectStart
    $Status  = if ($ExitCode -eq 0) { 'PASS' } else { 'FAIL' }

    $Results.Add([PSCustomObject]@{
        Project  = $ProjectName
        Status   = $Status
        ExitCode = $ExitCode
        Elapsed  = $Elapsed
    })

    if ($ExitCode -eq 0) {
        Write-Host " PASS  $ProjectName  ($([int]$Elapsed.TotalSeconds)s)" -ForegroundColor Green
    } else {
        Write-Host " FAIL  $ProjectName  ($([int]$Elapsed.TotalSeconds)s)" -ForegroundColor Red
    }
}

$TotalElapsed = (Get-Date) - $StartTime
$PassCount    = ($Results | Where-Object Status -eq 'PASS').Count
$FailCount    = ($Results | Where-Object Status -eq 'FAIL').Count

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Summary  ($([int]$TotalElapsed.TotalSeconds)s total)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

foreach ($R in $Results) {
    $Color = if ($R.Status -eq 'PASS') { 'Green' } else { 'Red' }
    $Line  = " [{0}]  {1}  ({2}s)" -f $R.Status, $R.Project, [int]$R.Elapsed.TotalSeconds
    Write-Host $Line -ForegroundColor $Color
}

Write-Host ""
Write-Host " Passed: $PassCount  |  Failed: $FailCount" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($FailCount -gt 0) {
    exit 1
}
exit 0

#requires -Version 5.1

<#
.SYNOPSIS
Converts legacy .xls workbooks to .xlsx or .xlsm by using desktop Excel.

.DESCRIPTION
Inspects legacy .xls workbooks without running their macros. Workbooks that
contain a VBA project are converted to .xlsm; all others are converted to
.xlsx. The source files are never deleted or overwritten, and existing output
files are skipped.

Normal execution creates the converted workbooks. Add -DryRun to validate the
inputs and show the planned destinations without creating output folders or
workbooks. Progress is written to stderr. Results use key=value lines by default;
-Json selects JSON and -FullOutput includes unchanged/skipped items.

.PARAMETER SourcePath
Source folder containing the legacy .xls workbooks to convert. Defaults to the
current folder. This is not the output destination.

.PARAMETER OutputDirectory
Optional destination folder. When omitted, each converted workbook is saved
next to its source .xls workbook.

.PARAMETER Recurse
Includes .xls workbooks in subfolders. With the default output behavior, each
converted workbook is saved next to its source. When OutputDirectory is set,
the output is flat, so duplicate output names are reported as collisions.

.PARAMETER DryRun
Shows and validates the conversion plan without creating output folders or
workbooks. Desktop Excel is still started and each source workbook is opened
read-only so that the required output format can be determined accurately.

.PARAMETER Write
Deprecated compatibility switch. Normal execution already writes. Do not use
this switch together with -DryRun.

.EXAMPLE
.\scripts\Convert-XlsToOpenXml.ps1 -SourcePath 'C:\path\to\legacy-workbooks'

.EXAMPLE
.\scripts\Convert-XlsToOpenXml.ps1 -SourcePath 'C:\path\to\legacy-workbooks' -DryRun
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$SourcePath = ".",

    [string]$OutputDirectory,

    [switch]$Recurse,

    [switch]$DryRun,

    [switch]$Write,

    [switch]$Json,

    [switch]$FullOutput,

    [switch]$Quiet
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

function Write-ResultValue {
    param($Value, [string]$Key = "")
    if ($Value -is [System.Collections.IDictionary]) {
        if ($Value.Count -eq 0) { [Console]::Out.WriteLine("$Key={}") }
        foreach ($child in $Value.Keys) {
            $next = if ($Key) { "$Key.$child" } else { [string]$child }
            Write-ResultValue $Value[$child] $next
        }
    } elseif ($Value -is [System.Management.Automation.PSCustomObject]) {
        foreach ($property in $Value.PSObject.Properties) {
            Write-ResultValue $property.Value "$Key.$($property.Name)"
        }
    } elseif ($Value -is [System.Collections.IList]) {
        if ($Value.Count -eq 0) { [Console]::Out.WriteLine("$Key=[]") }
        for ($index = 0; $index -lt $Value.Count; $index++) {
            Write-ResultValue $Value[$index] "$Key[$index]"
        }
    } else {
        $text = if ($null -eq $Value) { "null" } elseif ($Value -is [bool]) { $Value.ToString().ToLowerInvariant() } else { [string]$Value }
        $text = [regex]::Replace($text, '[\x00-\x1f\x7f-\x9f\u2028\u2029]', { param($m) '\u{0:x4}' -f [int][char]$m.Value })
        [Console]::Out.WriteLine("$Key=$text")
    }
}
function Write-FinalResult {
    param($Value)
    if ($Json) { $Value | ConvertTo-Json -Depth 30 -Compress } else { Write-ResultValue $Value }
}
try {
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($DryRun -and $Write) {
    throw "-DryRun and the deprecated -Write switch cannot be combined."
}

$writeEnabled = -not $DryRun

function Write-ProgressLine {
    param(
        [Parameter(Mandatory)]
        [string]$Level,

        [Parameter(Mandatory)]
        [string]$Message
    )

    if (-not $Quiet -or $Level -in @("ERROR", "CRITICAL")) {
        [Console]::Error.WriteLine("[{0}] {1}", $Level, $Message)
    }
}

function Add-Result {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Results,

        [Parameter(Mandatory)]
        [string]$Status,

        [Parameter(Mandatory)]
        [System.IO.FileInfo]$Source,

        [string]$Destination,

        [AllowNull()]
        [Nullable[bool]]$HasVba,

        [string]$Message = ""
    )

    $format = if ($null -eq $HasVba) {
        $null
    }
    elseif ($HasVba) {
        "xlsm"
    }
    else {
        "xlsx"
    }

    [void]$Results.Add([pscustomobject][ordered]@{
        status      = $Status
        source      = $Source.FullName
        destination = $Destination
        format      = $format
        hasVba      = $HasVba
        message     = $Message
    })
}

$sourceItem = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
if (-not $sourceItem.PSIsContainer) {
    throw "SourcePath must be a folder: $($sourceItem.FullName)"
}

$sourceRoot = $sourceItem.FullName
$hasExplicitOutput = -not [string]::IsNullOrWhiteSpace($OutputDirectory)
$outputRoot = if ($hasExplicitOutput) {
    [System.IO.Path]::GetFullPath($OutputDirectory)
}
else {
    $null
}
$outputLabel = if ($hasExplicitOutput) { $outputRoot } else { "alongside each input workbook" }

$searchArguments = @{
    LiteralPath = $sourceRoot
    Filter      = "*.xls"
    File        = $true
}
if ($Recurse) {
    $searchArguments.Recurse = $true
}

# Windows wildcard matching may return names such as book.xlsx for "*.xls".
# Require the actual extension to be exactly .xls before opening anything.
$files = @(
    Get-ChildItem @searchArguments |
        Where-Object { $_.Extension -ieq ".xls" } |
        Sort-Object FullName
)
$results = [System.Collections.Generic.List[object]]::new()
$plannedDestinations = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
$excel = $null

if ($Write) {
    Write-ProgressLine "WARNING" (
        "-Write is deprecated and no longer required; normal execution writes. " +
        "Use -DryRun for a read-only preview."
    )
}

Write-ProgressLine "INFO" ("Mode={0} | Source={1} | Output={2} | Files={3}" -f (
        $(if ($writeEnabled) { "write" } else { "dry-run" })
    ), $sourceRoot, $outputLabel, $files.Count)

if ($files.Count -gt 0) {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3
    }
    catch {
        $message = "Unable to start desktop Excel: $($_.Exception.Message)"
        foreach ($file in $files) {
            Add-Result -Results $results -Status "error" -Source $file -HasVba $null -Message $message
        }
        Write-ProgressLine "ERROR" $message
    }

    if ($null -ne $excel) {
        if ($writeEnabled -and $hasExplicitOutput) {
            [void](New-Item -ItemType Directory -Path $outputRoot -Force)
        }

        try {
            foreach ($file in $files) {
                $workbook = $null

                try {
                    # UpdateLinks=0 and ReadOnly=$true keep inspection non-mutating.
                    $workbook = $excel.Workbooks.Open($file.FullName, 0, $true)
                    $hasVba = [bool]$workbook.HasVBProject
                    $extension = if ($hasVba) { ".xlsm" } else { ".xlsx" }
                    $fileFormat = if ($hasVba) { 52 } else { 51 }
                    $destinationDirectory = if ($hasExplicitOutput) {
                        $outputRoot
                    }
                    else {
                        $file.DirectoryName
                    }
                    $destination = Join-Path $destinationDirectory ($file.BaseName + $extension)

                    if (-not $plannedDestinations.Add($destination)) {
                        $message = "Another source workbook maps to the same output name."
                        Add-Result -Results $results -Status "collision" -Source $file `
                            -Destination $destination -HasVba $hasVba -Message $message
                        Write-ProgressLine "ERROR" ("[collision] {0} -> {1}" -f $file.FullName, $destination)
                        continue
                    }

                    if (Test-Path -LiteralPath $destination) {
                        $message = "The destination already exists; it was not overwritten."
                        Add-Result -Results $results -Status "skipped-existing" -Source $file `
                            -Destination $destination -HasVba $hasVba -Message $message
                        Write-ProgressLine "WARNING" ("[skipped-existing] {0} -> {1}" -f $file.Name, $destination)
                        continue
                    }

                    if (-not $writeEnabled) {
                        Add-Result -Results $results -Status "would-convert" -Source $file `
                            -Destination $destination -HasVba $hasVba
                        Write-ProgressLine "INFO" ("[would-convert] {0} -> {1}" -f $file.Name, $destination)
                        continue
                    }

                    [void]$workbook.SaveAs($destination, $fileFormat)
                    Add-Result -Results $results -Status "converted" -Source $file `
                        -Destination $destination -HasVba $hasVba
                    Write-ProgressLine "SUCCESS" ("[converted] {0} -> {1}" -f $file.Name, $destination)
                }
                catch {
                    $message = $_.Exception.Message
                    Add-Result -Results $results -Status "error" -Source $file -HasVba $null -Message $message
                    Write-ProgressLine "ERROR" ("[error] {0} | {1}" -f $file.FullName, $message)
                }
                finally {
                    if ($null -ne $workbook) {
                        $workbook.Close($false)
                        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
                    }
                }
            }
        }
        finally {
            $excel.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
else {
    Write-ProgressLine "INFO" "No .xls workbooks were found."
}

$counts = [ordered]@{}
foreach ($status in @("would-convert", "converted", "skipped-existing", "collision", "error")) {
    $counts[$status] = @($results | Where-Object status -EQ $status).Count
}

$summary = [ordered]@{
    mode            = if ($writeEnabled) { "write" } else { "dry-run" }
    sourcePath      = $sourceRoot
    outputMode      = if ($hasExplicitOutput) { "specified-directory" } else { "alongside-input" }
    outputDirectory = $outputRoot
    recursive       = [bool]$Recurse
    filesFound      = $files.Count
    counts          = $counts
}

$summary.status = if (($counts.error + $counts.collision) -gt 0) { "error" } else { "success" }
$summary.results = @($results | Where-Object { $FullOutput -or $_.status -notin @("skipped", "unchanged") })
$summary.omittedResults = $results.Count - $summary.results.Count
if ($summary.omittedResults -gt 0) { $summary.detailsHint = "Use -FullOutput for every item." }
Write-FinalResult $summary

if (($counts.error + $counts.collision) -gt 0) {
    exit 1
}

} catch {
    [Console]::Error.WriteLine("[ERROR] {0}", $_.Exception.Message)
    Write-FinalResult ([ordered]@{status="error"; sourcePath=$SourcePath; message=$_.Exception.Message})
    exit 1
}

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [Alias('Input', 'In')]
    [string]$InputPath,

    [Parameter(Mandatory)]
    [Alias('Output', 'Out')]
    [string]$OutputPath,

    # Default = Failed. Set to '' (empty string) to export all statuses.
    [string]$Status = 'Failed',

    # Excel-only niceties (used when -OutputPath ends with .xlsx)
    [string]$WorksheetName = 'ZTA Tasks',
    [string]$TableName = 'ZTATasks'
)

function Convert-ZTAReportToRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InputPath,
        [string]$Status = 'Failed'
    )

    if (-not (Test-Path $InputPath)) { throw "Input file not found: $InputPath" }

    # Read JSON, handle UTF-8 BOM
    $bytes = [System.IO.File]::ReadAllBytes($InputPath)
    $hasBOM = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    if ($hasBOM) { $content = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3) }
    else { $content = [System.Text.Encoding]::UTF8.GetString($bytes) }

    $data = $content | ConvertFrom-Json
    if (-not $data.Tests) { throw "JSON did not contain a 'Tests' array." }

    $tests = $data.Tests
    if ($Status -ne '') { $tests = $tests | Where-Object { $_.TestStatus -eq $Status } }

    $bucketMap = @{
        'Identity'   = 'Identity & Access'
        'Devices'    = 'Devices & Endpoint'
        'Apps'       = 'Apps & Access Governance'
        'Data'       = 'Policies & Governance'
        'Network'    = 'Network & Threat Protection'
        'Monitoring' = 'Monitoring & Response'
        'Default'    = 'Policies & Governance'
    }

    $rows = $tests | ForEach-Object {
        $pillar = [string]$_.TestPillar
        $bucket = $bucketMap[$pillar]; if (-not $bucket) { $bucket = $bucketMap['Default'] }

        $desc = ($_.TestDescription -replace '\r?\n', ' ').Trim()
        $result = ($_.TestResult -replace '\r?\n', ' ').Trim()

        [pscustomobject]@{
            'Task Title'          = $_.TestTitle
            'Bucket'              = $bucket
            'Category'            = $_.TestCategory
            'Pillar'              = $pillar
            'Risk'                = $_.TestRisk
            'Impact'              = $_.TestImpact
            'Implementation Cost' = $_.TestImplementationCost
            'Status'              = 'Not Started'
            'Description'         = $desc
            'Result Summary'      = $result
            'Label.High'          = ($_.TestRisk -eq 'High')
            'Label.Medium'        = ($_.TestRisk -eq 'Medium')
            'Label.Low'           = ($_.TestRisk -eq 'Low')
            'Label.QuickWin'      = ($_.TestImplementationCost -eq 'Low')
        }
    }

    return , $rows
}

$rows = Convert-ZTAReportToRows -InputPath $InputPath -Status $Status

$ext = [System.IO.Path]::GetExtension($OutputPath).ToLowerInvariant()
switch ($ext) {
    '.xlsx' {
        try { Import-Module ImportExcel -ErrorAction Stop }
        catch {
            throw "The 'ImportExcel' module is required for .xlsx output.
Install it with:  Install-Module ImportExcel -Scope CurrentUser"
        }

        # Write to Excel as a proper Excel Table
        $rows | Export-Excel `
            -Path $OutputPath `
            -WorksheetName $WorksheetName `
            -TableName $TableName `
            -TableStyle 'Medium2' `
            -AutoSize `
            -FreezeTopRow `
            -AutoFilter `
            -ClearSheet

        Write-Host "Wrote $($rows.Count) rows to Excel table '$TableName' on sheet '$WorksheetName': $OutputPath"
    }
    '.csv' {
        $rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $OutputPath
        Write-Host "Wrote $($rows.Count) rows to CSV: $OutputPath"
    }
    default {
        throw "Unknown output type for '$OutputPath'. Use .xlsx or .csv."
    }
}

<# 
.SYNOPSIS
Convert ZeroTrustAssessmentReport.json to a Planner-friendly CSV across all pillars.

.EXAMPLES
# 1) Run as a script (tab-completes -input / -output)
.\Convert-ZTAReportToCsv.ps1 -input '.\ZeroTrustAssessmentReport.json' -output '.\zerotrust_tasks_allpillars.csv'

# 2) All statuses (Failed + Passed, etc.)
.\Convert-ZTAReportToCsv.ps1 -input '.\ZeroTrustAssessmentReport.json' -output '.\zerotrust_tasks_allpillars_allstatus.csv' -Status ''

# 3) Still possible to dot-source and call the function directly
. .\Convert-ZTAReportToCsv.ps1
Convert-ZTAReportToCsv -InputPath $env:USERPROFILE\Downloads\ZeroTrustAssessmentReport.json `
                       -OutputPath "$env:USERPROFILE\Downloads\zerotrust_tasks.csv"
#>

# --- Script parameters (so tab completion works on the script itself) ---
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [Alias('Input', 'In')]
    [string]$InputPath,

    [Parameter(Mandatory)]
    [Alias('Output', 'Out')]
    [string]$OutputPath,

    # Default = Failed. Set to '' (empty string) to export all statuses.
    [string]$Status = 'Failed'
)

function Convert-ZTAReportToCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,

        [Parameter(Mandatory)]
        [string]$OutputPath,

        # Default = Failed. Set to '' (empty string) to export all statuses.
        [string]$Status = 'Failed'
    )

    if (-not (Test-Path $InputPath)) {
        throw "Input file not found: $InputPath"
    }

    # --- Read JSON safely (handles UTF-8 BOM) ---
    $bytes = [System.IO.File]::ReadAllBytes($InputPath)
    $hasBOM = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    if ($hasBOM) { $content = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3) }
    else { $content = [System.Text.Encoding]::UTF8.GetString($bytes) }

    $data = $content | ConvertFrom-Json
    if (-not $data.Tests) { throw "JSON did not contain a 'Tests' array." }

    $tests = $data.Tests
    if ($Status -ne '') { $tests = $tests | Where-Object { $_.TestStatus -eq $Status } }

    # Optional: map pillars to your Planner buckets (edit to match your board)
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

        $isHigh = $_.TestRisk -eq 'High'
        $isMedium = $_.TestRisk -eq 'Medium'
        $isLow = $_.TestRisk -eq 'Low'
        $isQuick = $_.TestImplementationCost -eq 'Low'

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
            'Label.High'          = $isHigh
            'Label.Medium'        = $isMedium
            'Label.Low'           = $isLow
            'Label.QuickWin'      = $isQuick
        }
    }

    $rows | Export-Csv -NoTypeInformation -Encoding UTF8 $OutputPath
    Write-Host "Wrote $($rows.Count) rows to: $OutputPath"
}

# --- Invoke the function using the script's parameters ---
Convert-ZTAReportToCsv @PSBoundParameters

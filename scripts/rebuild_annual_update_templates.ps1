param(
    [Parameter(Mandatory = $true)] [string] $PropertySource,
    [Parameter(Mandatory = $true)] [string] $ConcessionSource,
    [Parameter(Mandatory = $true)] [string] $FutureSource,
    [Parameter(Mandatory = $true)] [string] $PropertyTarget,
    [Parameter(Mandatory = $true)] [string] $ConcessionTarget,
    [Parameter(Mandatory = $true)] [string] $FutureTarget
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function U {
    param([int[]] $Codes)
    return -join ($Codes | ForEach-Object { [char]$_ })
}

$TextPropertyNature = U 20135,26435
$TextConcessionNature = U 29305,35768,32463,33829,26435
$TextInitial = U 26399,21021
$TextTerminal = U 26399,26411,22238,25910
$TextResidual = U 27531,20540
$TextWholeProject = U 39033,30446,25972,20307
$TextPropertyType = U 22253,21306,22522,30784,35774,26045
$TextConcessionType = U 20132,36890,22522,30784,35774,26045

function Remove-PersonalMetadata {
    param([Parameter(Mandatory = $true)] $Workbook)
    foreach ($code in 4, 7, 8, 99) {
        try {
            $Workbook.RemoveDocumentInformation($code)
        } catch {
        }
    }
}

function Sanitize-PropertyTemplate {
    param([Parameter(Mandatory = $true)] $Workbook)
    $ws = $Workbook.Worksheets.Item(1)
    if ($ws.UsedRange.Rows.Count -gt 149) {
        $ws.Rows("150:" + $ws.Rows.Count).Delete()
    }
    $ws.Rows("4:149").ClearContents()
    $ws.Cells.Item(4, 1).Value = '000000.SH'
    $ws.Cells.Item(4, 2).Value = 'Template Fund'
    $ws.Cells.Item(4, 3).Value = $TextPropertyType
    $ws.Cells.Item(4, 4).Value = $TextPropertyNature
    $ws.Cells.Item(4, 5).Value = $TextWholeProject
    $ws.Cells.Item(4, 9).Value = 2027
    $ws.Cells.Item(4, 10).Value = 0
    $ws.Cells.Item(4, 13).Value = 2027
    $ws.Cells.Item(4, 14).Value = [datetime]'2027-06-30'
    $ws.Cells.Item(4, 24).Value = 'Template Report Period'
    $ws.Cells.Item(4, 25).Value = [datetime]'2026-12-31'

    $ws.Cells.Item(110, 13).Value = $TextInitial
    $ws.Cells.Item(110, 14).Value = [datetime]'2026-12-31'

    $ws.Cells.Item(148, 21).Value = $TextResidual
    $ws.Cells.Item(149, 20).Value = [datetime]'2037-06-30'

    $ws.Activate() | Out-Null
    $ws.Range('A1').Select() | Out-Null
}

function Sanitize-ConcessionTemplate {
    param([Parameter(Mandatory = $true)] $Workbook)
    $ws = $Workbook.Worksheets.Item(1)
    if ($ws.UsedRange.Rows.Count -gt 16) {
        $ws.Rows("17:" + $ws.Rows.Count).Delete()
    }
    $ws.Rows("4:16").ClearContents()
    $ws.Cells.Item(4, 9).Value = $TextInitial
    $ws.Cells.Item(4, 10).Value = [datetime]'2026-12-31'
    $ws.Cells.Item(4, 27).Value = 0.0326

    $ws.Cells.Item(5, 1).Value = '000000.SH'
    $ws.Cells.Item(5, 2).Value = 'Template Fund'
    $ws.Cells.Item(5, 3).Value = $TextConcessionType
    $ws.Cells.Item(5, 4).Value = $TextConcessionNature
    $ws.Cells.Item(5, 5).Value = $TextWholeProject
    $ws.Cells.Item(5, 9).Value = 2027
    $ws.Cells.Item(5, 10).Value = [datetime]'2027-06-30'
    $ws.Cells.Item(5, 11).Value = 0
    $ws.Cells.Item(5, 13).Value = 'Template Report Period'
    $ws.Cells.Item(5, 14).Value = [datetime]'2026-12-31'

    $ws.Cells.Item(16, 1).Value = '000000.SH'
    $ws.Cells.Item(16, 2).Value = 'Template Fund'
    $ws.Cells.Item(16, 3).Value = $TextConcessionType
    $ws.Cells.Item(16, 4).Value = $TextConcessionNature
    $ws.Cells.Item(16, 5).Value = $TextWholeProject
    $ws.Cells.Item(16, 9).Value = $TextTerminal
    $ws.Cells.Item(16, 10).Value = [datetime]'2036-06-30'
    $ws.Cells.Item(16, 11).Value = 0

    $ws.Activate() | Out-Null
    $ws.Range('A1').Select() | Out-Null
}

function Sanitize-FutureTemplate {
    param([Parameter(Mandatory = $true)] $Workbook)
    $ws = $Workbook.Worksheets.Item(1)
    if ($ws.UsedRange.Rows.Count -gt 2) {
        $ws.Rows("3:" + $ws.Rows.Count).Delete()
    }
    $ws.Rows('2:2').ClearContents()
    $ws.Activate() | Out-Null
    $ws.Range('A1').Select() | Out-Null
}

function Save-PropertyTemplate {
    param(
        [Parameter(Mandatory = $true)] [string] $SourcePath,
        [Parameter(Mandatory = $true)] [string] $TargetPath
    )
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    try {
        $workbook = $excel.Workbooks.Open($targetPath, 0, $false)
        try {
            Sanitize-PropertyTemplate $workbook
            Remove-PersonalMetadata $workbook
            $workbook.Save()
        } finally {
            $workbook.Close($true)
        }
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Save-ConcessionTemplate {
    param(
        [Parameter(Mandatory = $true)] [string] $SourcePath,
        [Parameter(Mandatory = $true)] [string] $TargetPath
    )
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    try {
        $workbook = $excel.Workbooks.Open($targetPath, 0, $false)
        try {
            Sanitize-ConcessionTemplate $workbook
            Remove-PersonalMetadata $workbook
            $workbook.Save()
        } finally {
            $workbook.Close($true)
        }
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Save-FutureTemplate {
    param(
        [Parameter(Mandatory = $true)] [string] $SourcePath,
        [Parameter(Mandatory = $true)] [string] $TargetPath
    )
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    try {
        $workbook = $excel.Workbooks.Open($targetPath, 0, $false)
        try {
            Sanitize-FutureTemplate $workbook
            Remove-PersonalMetadata $workbook
            $workbook.Save()
        } finally {
            $workbook.Close($true)
        }
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Save-PropertyTemplate $PropertySource $PropertyTarget
Save-ConcessionTemplate $ConcessionSource $ConcessionTarget
Save-FutureTemplate $FutureSource $FutureTarget

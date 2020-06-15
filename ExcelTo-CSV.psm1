<#
.SYNOPSIS
    ExcelTo-CSV - convert Excel file to CSV or TXT format

.DESCRIPTION
    Script for converting an Excel workbook to CSV or TXT file(s).

.NOTES
    File Name:  ExcelTo-CSV.psm1
    Author:     Marcus Libäck <marcus.liback@gmail.com>
    Requires:   PowerShell v4

.EXAMPLE
    ExcelTo-CSV -Sheet "SheetName" SheetIndex <sheet number> -InFile Foo.xlsx -ExportFile Foo.pdf
#>

# Command line parameters

# -----------------------------------------------------------------------------
# Support functions
# -----------------------------------------------------------------------------
function PathToAbsolute([string] $path)
{
    if ($path -eq "") {
        $path = Convert-Path .
    }
    if (-not $(Split-Path -IsAbsolute $path)) {
        $path = Convert-Path (Join-Path $(Convert-Path .) $path)
    }

    return $path
}

function CleanupExcelInstance
{
    if ($worksheet) {
        while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)){}
    }

    if ($workbook) {
        $workbook.Close()
        while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)){}
        while( [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($workbook)){}
    }

    if ($excel) {
        $excel.Quit()
        while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)){}
        [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($excel)
    }
    [System.GC]::Collect()
}

function SafeFileName ([string] $name) {
    $invalidChars = $([string][System.IO.Path]::GetInvalidFileNameChars() -join '')
    $regex = "[{0}]" -f [regex]::Escape($invalidChars)
    return ($name -replace $regex)
}

# function SaveExcelSheet ([System.__ComObject] $worksheet, [Int] $format,  [System.IO.Path] $file) {
#     try {
#         $worksheet = $workbook.sheets.item($sheet)
#         $filename = "$($_ExportPath)\$($_ExportFile)" + $_FileFormat[0]
#         Write-Output "Saving sheet:`t$($worksheet.name) -> $filename"
#         $worksheet.SaveAs("$_ExportPath\$_ExportFile" + $_FileFormat[0], $_FileFormat[1], 0, 0, 0, 0, 0, 0, 0, $_UseCulture)
#     } catch {
#         Write-Output "" "Error! No such sheet: `"$sheet`", exiting script!"
#         Write-Output "" "Message:" "--------" "$($_.Exception.Message)" "$($_.Exception.ItemName)"
#         CleanupExcelInstance
#         exit 1
#     }
# }

function ExcelTo-CSV {
    # -----------------------------------------------------------------------------
    # Set parameters & settings
    # -----------------------------------------------------------------------------
    # Set up path and filename for the exported PDF file
    if ($InFile) {$_InPath    = Split-Path $InFile -Parent}
    if ($InFile) {$_InFile    = Split-Path $InFile -Leaf}

    # Set up path and filename for the exported PDF file
    if ($ExportFile) {
        # Remove double file extension if necessary
        if ($ExportFile -match ".*\.(txt)|(csv)$") {
            $ExportFile = $ExportFile.Substring(0, $ExportFile.Length-4)
        }

        $_ExportPath    = Split-Path $ExportFile -Parent
        $_ExportFile    = Split-Path $ExportFile -Leaf

    } else {
        $_ExportPath    = $_InPath
        $_ExportFile    = $(Get-Item $InFile -ErrorAction SilentlyContinue).BaseName
        #$_ExportPath    = Convert-Path .
        #$_ExportFile    = "$((Get-Item $MyInvocation.MyCommand.Definition).BaseName)"
    }

    # Set up file format for export
    if ($TXT) {
        $_FileFormat = ".txt", 20
    } else {
        $_FileFormat = ".csv", 6
    }

    # Use local culture (eg. ';' as separator instead of ',')
    if ($UseCulture) {
        $_UseCulture = $true
    } else {
        $_UseCulture = $false
    }

    # Convert relative paths to absolute
    $_InPath        = PathToAbsolute($_InPath)
    $_ExportPath    = PathToAbsolute($_ExportPath)

    # -----------------------------------------------------------------------------
    # Main program
    # -----------------------------------------------------------------------------

    # Create new Excel COM-object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible          = $false
    $excel.Interactive      = $false
    $excel.DisplayAlerts    = $false

    # Open Excel file
    try {
        Write-Output "Opening file:`t$_InPath\$_InFile"
        $workbook = $excel.workbooks.open("$_InPath\$_InFile")
    } catch {
        Write-Output "" "Error! Could not open Excel file, exiting script!"
        Write-Output "" "Message:" "--------" "$($_.Exception.Message)" "$($_.Exception.ItemName)"
        CleanupExcelInstance
        exit 1
    }

    # Refresh all sheets before continuing (ensures all data is current)
    if ($Refresh) {
        try {
            Write-Output "Excel:`t`tRefreshing all tables, formulas, connections etc"
            $workbook.RefreshAll()
            $excel.Application.CalculateUntilAsyncQueriesDone()
        } catch {
            Write-Output "" "Error! Could not refresh Excel workbook, exiting script!"
            Write-Output "" "Message:" "--------" "$($_.Exception.Message)" "$($_.Exception.ItemName)"
            CleanupExcelInstance
            exit 1
        }
    }

    if ($SheetName -or $SheetIndex) {
        # Use either SheetName or SheetIndex
        if ($SheetName) {
            $sheet = $SheetName
        } elseif ($SheetIndex) {
            $sheet = $SheetIndex
        }

        # Try to save designated sheet
        try {
            $worksheet  = $workbook.sheets.item($sheet)
            $filename   = "$($_ExportPath)\$($_ExportFile)" + $_FileFormat[0]
            Write-Output "Saving sheet:`t$($worksheet.name) -> $filename"
            $worksheet.SaveAs("$_ExportPath\$_ExportFile" + $_FileFormat[0], $_FileFormat[1], 0, 0, 0, 0, 0, 0, 0, $_UseCulture)
        } catch {
            Write-Output "" "Error! No such sheet: `"$sheet`", exiting script!"
            Write-Output "" "Message:" "--------" "$($_.Exception.Message)" "$($_.Exception.ItemName)"
            CleanupExcelInstance
            exit 1
        }
    } else {
        # Save all sheets if nothing is specified
        for ($i = 1; $i -le $workbook.sheets.count; $i++) {
            $worksheet  = $workbook.sheets.item($i)
            $filename   = "$($_ExportPath)\$($_ExportFile)_$($worksheet.Name)" + $_FileFormat[0]
            Write-Output "Saving sheet:`t$($worksheet.name) -> $filename"
            $worksheet.SaveAs($filename, $_FileFormat[1], 0, 0, 0, 0, 0, 0, 0, $_UseCulture)
        }
    }

    Write-Output "All done! Exiting script."
    CleanupExcelInstance
}

Export-ModuleMember -Function ExcelTo-CSV

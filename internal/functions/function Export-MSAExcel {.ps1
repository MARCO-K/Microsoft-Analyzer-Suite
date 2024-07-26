<#
.SYNOPSIS
Exports data to an Excel file with specific formatting and conditional formatting.

.DESCRIPTION
This function exports an array of objects to an Excel file, applying specific formatting to the top row and conditional formatting based on a blacklist.

.PARAMETER exportdata
An array of objects to be exported to Excel.

.PARAMETER name
The name of the Excel file and worksheet.

.EXAMPLE
Export-MSAExcel -exportdata $data -name "SignInLogs"
Exports the data in $data to an Excel file named "SignInLogs.xlsx" with appropriate formatting.

.NOTES
Ensure that the Export-Excel module is installed and available in the environment.
#>
function Export-MSAExcel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [object[]]$exportdata,

        [Parameter(Mandatory = $true)]
        [string]$name
    )

    # Define the output folder path
    $outputFolder = Join-Path -Path $OUTPUT_FOLDER -ChildPath "SignInLogsGraph\XLSX\$name.xlsx"

    # Check if there is data to export
    if ($exportdata.Count -gt 0) {
        Write-PSFMessage -Level Verbose -Message "Exporting data to $outputFolder"

        # Export data to Excel with specified parameters and formatting
        $exportdata | Export-Excel -Path $outputFolder -NoNumberConversion * -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter -WorkSheetname $name -CellStyleSB {
            param($WorkSheet)

            # Define formatting for the top row
            $BackgroundColor = [System.Drawing.Color]::FromArgb(50, 60, 220)
            $n = ($exportdata | Get-Member -MemberType Noteproperty | Measure-Object).Count
            $end = ([char](64 + $n), (+$n - 64))[$n -ge 65]
            $header = '"A1:$end"'
            Set-Format -Address $WorkSheet.Cells[$header] -BackgroundColor $BackgroundColor -FontColor White

            # Set horizontal alignment for columns A to E
            $WorkSheet.Cells[$header].Style.HorizontalAlignment = "Center"

            # Retrieve the blacklist content based on the name
            $Blacklist = ($supportlist | Where-Object { $_.Name -like $name }).Content
            if ($Blacklist.Count -gt 0) {
                $Col = ($Blacklist | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -eq $name }).Name
                foreach ($value in $Blacklist.$Col) {
                    $ConditionValue = 'NOT(ISERROR(FIND("{0}",$A1)))' -f $value
                    Add-ConditionalFormatting -Address $WorkSheet.Cells[$header] -WorkSheet $WorkSheet -RuleType 'Expression' -ConditionValue $ConditionValue -BackgroundColor Red
                }
            }
        }
    } else {
        Write-PSFMessage -Level Warning -Message "No data to export"
    }
}
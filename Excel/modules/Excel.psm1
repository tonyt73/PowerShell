<#
 .Synopsis
  Nice helper function for with the Export-Excel module

 .Description
  Functions to:
    * Create new Graphs
    * Import CSV data into a worksheet
    * Add new Charts to an existing Graph
        - with support for a 2nd axis
    * Ability to set graph colours and series names

  .Example
   # TODO: Examples
#>

Set-StrictMode -Version 3.0

Function New-SheetFromCsv {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $ExcelFile,
        [Parameter(Mandatory)]
        $WorksheetName,
        [Parameter(Mandatory, ValueFromPipeline)]
        [String[]]$CsvData
    )

    # Annoyingly to process an array from the pipeline we need to reconstruct the 
    # array that is pass into the pipeline, still powershell breaks it up for us.
    # Splitting it up works for more "process" driven modules, but here we want
    # list of CSV data in one array to pass into the ConvertFrom-Csv module
    Begin {
        $csv = @()
    }
    Process {
        foreach ($data in $CsvData) {
            $csv += $data
        }
    }
    End {
        $csv | ConvertFrom-Csv | Export-Excel -Path $ExcelFile -WorksheetName $WorksheetName -AutoSize -AutoFilter -FreezeTopRow
    }
}

Function New-ExcelGraph {
    [CmdletBinding(DefaultParameterSetName="ColorsFile")]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        $ExcelFile,
        [Parameter(Mandatory)]
        $XSeries,
        [Parameter(Mandatory)]
        $YSeries,
        [Parameter(Mandatory)]
        $Title,
        [Parameter(Mandatory)]
        $WorksheetName,
        [Parameter(Mandatory=$false)]
        $ChartType = "Area",
        [Parameter(Mandatory=$false)]
        $MajorTickMark = "None",
        [Parameter(Mandatory=$false)]
        $MinorTickMark = "None",
        [Parameter(Mandatory=$false)]
        [switch]$ReverseXAxis,
        [Parameter(Mandatory, ParameterSetName="Color")]
        $BackgroundColor = "White",
        [Parameter(Mandatory=$false)]
        $Width = 1500,
        [Parameter(Mandatory=$false)]
        $Height = 750,
        [Parameter(Mandatory, ParameterSetName="ColorsFile")]
        $ColorsFile
    )

    $chartDef = New-ExcelChartDefinition -XRange $XSeries -YRange $YSeries -Title $Title -Height $Height -Width $Width -ChartType $ChartType -Row 1 -Column 0 -LegendPosition Bottom
    $xl = Export-Excel -Path $ExcelFile -WorksheetName $WorksheetName -ExcelChartDefinition $chartDef -AutoNameRange -AutoFilter -AutoSize -PassThru
    # find the chart we added
    $dc = $xl.$WorksheetName.Drawings.Count - 1
    $chart = $xl.$WorksheetName.Drawings[$dc]
    # set its background color
    if ($PSCmdlet.ParameterSetName -eq "ColorsFile") {
        $color = Get-ColorFromFile -ColorsFile $ColorsFile -Color "Background"
    } else {
        $color = $BackgroundColor
    }
    $chart.Fill.Color = $color
    $chart.PlotArea.Fill.Color = $color
    $chart.XAxis.MajorTickMark = $MajorTickMark
    $chart.XAxis.MinorTickMark = $MinorTickMark
    if ($ReverseXAxis) {
        $chart.XAxis.Orientation = "MaxMin"
    }

    return $xl, $chart
}

Function Add-ExcelGraphSeries {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        $Chart,
        [Parameter(Mandatory)]
        $ChartType,
        [Parameter(Mandatory)]
        $XSeries,
        [Parameter(Mandatory)]
        $YSeries,
        [Parameter(Mandatory=$false)]
        [Switch]$UseSecondaryAxis
    )

    $newChart = $Chart.PlotArea.ChartTypes.Add($ChartType)
    $YSeries.foreach({
        $null = $newChart.Series.Add($_, $XSeries)
    })
    $null = $newChart.UseSecondaryAxis = $UseSecondaryAxis
}

Function Set-ExcelGraphSeries {
    [CmdletBinding(DefaultParameterSetName="ColorsFile")]
    param(
        [Parameter(Mandatory,ValueFromPipeline)]
        $Chart,
        [Parameter(Mandatory)]
        $Headers,
        [Parameter(Mandatory,ParameterSetName="Colors")]
        $Colors,
        [Parameter(Mandatory,ParameterSetName="ColorsFile")]
        $ColorsFile
    )

    #iterate all chart types and their series'
    $i = 0
    $Chart.PlotArea.ChartTypes.foreach({
        # get the chart type for the serie
        $ct = $_.ChartType
        $_.Series.foreach({
            if ($PSCmdlet.ParameterSetName -eq "ColorsFile") {
                $series = $_.Series.Split("!")[1]
                $color = Get-ColorFromFile -ColorsFile $ColorsFile -Color $series
            } else {
                $color = $Colors[$i]
            }
            # if line, then set line color, else set the fill color
            if ($ct -eq "Line") {
                $null = $_.LineColor = $color
            } else {
                $null = $_.Fill.Color = $color
            }
            # set the serie header
            $null = $_.Header = $Headers[$i]
            $i++
        })
    })
}

# https://www.powershellgallery.com/packages/SSRS/1.3.0/Content/New-XmlNamespaceManager.ps1
function New-XmlNamespaceManager ($XmlDocument, $DefaultNamespacePrefix) {

    $script:ErrorActionPreference = 'Stop'

    $NsMgr = New-Object -TypeName System.Xml.XmlNamespaceManager -ArgumentList $XmlDocument.NameTable
    $DefaultNamespace = $XmlDocument.DocumentElement.GetAttribute('xmlns')
    if ($DefaultNamespace -and $DefaultNamespacePrefix) {
        $NsMgr.AddNamespace($DefaultNamespacePrefix, $DefaultNamespace)
    }
    return ,$NsMgr # unary comma wraps $NsMgr so it isn't unrolled
}

# EPPlus and thus the Export-Excel module; doesn't expose the Excel property 'Multi-Level Category Labels'
# This setting allows us to group column names nicely and correctly
# Big Thank you to https://github.com/MiguelRozalen for his solution in C#.
#  He had some redundant code that I removed and the function works as expected.
# https://github.com/JanKallman/EPPlus/issues/189
# This function manipulates the base XML object and applies the flag.
Function Enable-MultiLevelLabel {
    [CmdletBinding(DefaultParameterSetName="ColorsFile")]
    param(
        [Parameter(Mandatory,ValueFromPipeline)]
        $Chart
    )
    $chartXml = $Chart.ChartXml
    $nsm = New-XmlNamespaceManager $chartXml
    $nsuri = $chartXml.DocumentElement.NamespaceURI
    $null = $nsm.AddNamespace("c", $nsuri) 

    $noMultiLvlLblNode = $chartXml.CreateElement("c:noMultiLvlLbl", $nsuri)
    $att = $chartXml.CreateAttribute("val")
    $att.Value = "0"
    $null = $noMultiLvlLblNode.Attributes.Append($att)

    $catAxNode = $chartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea/c:catAx", $nsm)
    $null = $catAxNode.AppendChild($noMultiLvlLblNode)
}

Function Export-ChartAsImage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        $Xl,
        [Parameter(Mandatory)]
        $WorksheetName,
        [Parameter(Mandatory)]
        $ExcelFile
    )


}

Function Start-Excel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        $XL
    )

    Close-ExcelPackage $XL -Show
}

$g_ColorsFileMap = @{}

Function Read-ColorsFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$ColorsFile
    )

    if (Test-Path $ColorsFile) {
        # Reset the color mappings
        $mapping = Get-Content -Path $ColorsFile | ConvertFrom-Json -AsHashtable
        $g_ColorsFileMap[$ColorsFile] = $mapping    
    } else {
        Write-Warning "Color settings file '$ColorsFile' was not found."
    }
}

Function Get-ColorFromFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$ColorsFile,
        [Parameter(Mandatory)]
        [String]$Color
    )

    if (-not ($g_ColorsFileMap.ContainsKey($ColorsFile))) {
        Read-ColorsFile $ColorsFile
    }
    if ($g_ColorsFileMap[$ColorsFile].ContainsKey($Color)) {
        return $g_ColorsFileMap[$ColorsFile][$Color]
    }
    return "Black"
}

Export-ModuleMember -Function New-SheetFromCsv
Export-ModuleMember -Function New-ExcelGraph
Export-ModuleMember -Function Add-ExcelGraphSeries
Export-ModuleMember -Function Set-ExcelGraphSeries
Export-ModuleMember -Function Start-Excel
Export-ModuleMember -Function Export-ChartAsImage
Export-ModuleMember -Function Enable-MultiLevelLabel
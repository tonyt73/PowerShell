<#
.SYNOPSIS Sorts an xml file by element and attribute names. Useful for diffing XML files.
.LINK https://danielsmon.com/2017/03/10/diff-xml-via-sorting-xml-elements-and-attributes/
#>
 
param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    # The path to the XML file to be sorted
    [string]$XmlPath
)
 
process {
    if (-not (Test-Path $XmlPath)) {
        Write-Warning "Skipping $XmlPath, as it was not found."
        continue;
    }
 
    $fullXmlPath = (Resolve-Path $XmlPath)
    [xml]$xml = Get-Content $fullXmlPath
    Write-Output "Sorting $fullXmlPath"
 
    function SortChildNodes($node, $depth = 0, $maxDepth = 32) {
        if ($node.HasChildNodes -and $depth -lt $maxDepth) {
            foreach ($child in $node.ChildNodes) {
                SortChildNodes $child ($depth + 1) $maxDepth
            }
        }
		
		# Need to ignore the root level and leave it unsorted.
		# Rad Studio kicks up a stink if its order is changed.
		if ($depth -ge 1) {
	 
			$sortedAttributes = $node.Attributes | Select-Object
			$sortedChildren = $node.ChildNodes | Sort-Object { $_.OuterXml }
	 
			$node.RemoveAll()
	 
			foreach ($sortedAttribute in $sortedAttributes) {
				[void]$node.Attributes.Append($sortedAttribute)
			}
	 
			foreach ($sortedChild in $sortedChildren) {
				[void]$node.AppendChild($sortedChild)
			}
		}
    }
 
    SortChildNodes $xml.DocumentElement
 
    $xml.Save($fullXmlPath)
}
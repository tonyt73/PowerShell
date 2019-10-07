<#
.SYNOPSIS Sorts all the Rad Studio project files
#>
 
process {
	Get-ChildItem -Path *.cbproj* | & '.\SortXML.ps1'
	Get-ChildItem -Path *.groupproj | & '.\SortXML.ps1'
	Get-ChildItem -Path *.xml | & '.\SortXML.ps1'
}
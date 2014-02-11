<#
.SYNOPSIS
   MamlDocBuild.ps1
.DESCRIPTION
   Builds MAML documentation for a targeted assembly using Gary Lapointe's MAML-building assembly;
   more information is available at http://blog.falchionconsulting.com/index.php/2011/01/creating-powershell-help-files-dynamically/
   Note that this script assumes that the assembly is located in the same folder as the script (i.e.,
   the current folder).
.NOTES
   Author: Sean McDonough
   Last Revision: 16-November-2011
.PARAMETER assemblyPath
   The path to the assembly which will be documented (e.g., "c:\assemblies\MyAssembly.dll")
.PARAMETER outputPath
   The path and filename of the XML help file to generate (e.g., "C:\output\help\MyAssembly.dll-help.xml")
.EXAMPLE
   MamlDocBuild.ps1 "c:\assemblies\MyAssembly.dll" "C:\output\help\"
#>
param 
(
	[string] $assemblyPath = "$(Read-Host 'Path to assembly to document [e.g. c:\assembly\MyAssembly.dll]')",
	[string] $outputPath = "$(Read-Host 'Path for XML file output [e.g. c:\output\help\]')"
)

function BuildDoc($assemblyPath, $outputPath)
{
	[System.Reflection.Assembly]::LoadFrom("Lapointe.PowerShell.MamlGenerator.dll") | Out-Null
    Write-Output "Generating MAML for '$assemblyPath'"
	[Lapointe.PowerShell.MamlGenerator.CmdletHelpGenerator]::GenerateHelp($assemblyPath, $outputPath, $true)
	Write-Output "MAML file created at '$outputPath'"
}
BuildDoc $assemblyPath $outputPath
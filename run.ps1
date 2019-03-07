param([Switch]$admin)
Add-Type -AssemblyName System.Windows.Forms

$classFilePath = "$PSScriptRoot\luokka.csv" # $PSScriptRoot is the folder where this script is located
$addonSyncPath = "\\10.132.0.97\Addons"
$defaultGateway = "10.132.0.1"
$internetGateway = "10.132.0.3"

. $PSScriptRoot\hallinta.ps1

param([Switch]$admin)
Add-Type -AssemblyName System.Windows.Forms

# $classFilePath = "$PSScriptRoot\luokka.csv" # $PSScriptRoot is the folder where this script is located
$addonSyncPath = "\\10.132.0.24\Addons"
$vbs3Path = "C:\Program Files\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI"
$steelBeastsPath = "C:\Program Files\eSim Games\SB Pro FI\Release"

. $PSScriptRoot\hallinta.ps1

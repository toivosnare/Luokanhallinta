Add-Type -AssemblyName System.Windows.Forms # For some reason this needs to be done in seperate script

$classFilePath = "$PSScriptRoot\luokka.csv" # $PSScriptRoot is the folder where this script is located
$username = $(whoami.exe) # Gets username of currently logged on user
$password = ""
$addonSyncPath = "\\10.132.0.97\Addons"
$addonSyncUsername = ""
$addonSyncPassword = ""
$debug = $false

. $PSScriptRoot\hallinta.ps1

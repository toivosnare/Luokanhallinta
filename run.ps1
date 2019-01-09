param(
    [String]$path = "$PSScriptRoot\luokka.csv", # Default values are used if alternatives
    [Switch]$debug = $false                     # are not supplied as command line arguments
)
Add-Type -AssemblyName System.Windows.Forms # For some reason this needs to be done in seperate script
. $PSScriptRoot\hallinta.ps1 -path $path -debug:$debug

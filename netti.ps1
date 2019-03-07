param(
    [Switch]$internetOn,
    [String]$defaultGateway = "10.132.0.1",
    [String]$internetGateway = "10.132.0.3"
)

$alias = Get-NetAdapter -Physical | Where-Object Status -eq "Up" | Select-Object -First 1 -ExpandProperty InterfaceAlias
Remove-NetRoute -InterfaceAlias $alias -Confirm:$false
New-NetRoute -InterfaceAlias $alias -DestinationPrefix "10.132.0.0/16" -NextHop $defaultGateway
if($internetOn)
{
    New-NetRoute -InterfaceAlias $alias -DestinationPrefix "0.0.0.0/0" -NextHop $internetGateway
}
Set-NetConnectionProfile -InterfaceAlias $alias -NetworkCategory Private
Restart-NetAdapter -InterfaceAlias $alias
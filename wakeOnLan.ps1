param(
    [String[]]$target,
    [Int]$port = 9
)

$broadcast = [Net.IPAddress]::Parse("255.255.255.255")
foreach($mac in $target)
{
    $mac = (($mac.replace(":", "")).replace("-", "")).replace(".", "")
    $t = 0, 2, 4, 6, 8, 10 | ForEach-Object {[Convert]::ToByte($mac.substring($_, 2), 16)}
    $packet = (,[Byte]255 * 6) + ($t * 16) # Creates the magic packet
    $UDPclient = [System.Net.Sockets.UdpClient]::new()
    $UDPclient.Connect($broadcast, $port)
    $UDPclient.Send($packet, 102) # Sends the magic packet
}
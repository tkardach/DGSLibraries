
# Returns ipconfig /all from the remote server as an array
function Get-ServerIPAddresses ($Session) {
    Write-Progress "Getting IP information from $($ServerName)"

    # Get the ipconfig and netstat information from the remote server
    $ipconfig = Invoke-Command -Session $Session -ScriptBlock {ipconfig /all}

    # return the parsed commands as a string array
    $InitServerList = Parse-IPConfig $ipconfig

    $ServerList = @()

    # convert the returned parsed list into an array of servers and ports
    foreach($server in $InitServerList) {
        $ServerList += $server.Server
    }

    $ServerList = $ServerList | select -Unique

    return $ServerList
}

# Returns netstat -na from the remote server as an array
function Get-ServerActivePorts ($Session) {
    Write-Progress "Getting Port information from $($ServerName)"

    # Remotely get the returned info from netstat -na
    $netstat = Invoke-Command -Session $Session -ScriptBlock {netstat -na}

    # Parse the netstat -na returned value
    $InitPortList = Parse-Netstat $netstat

    $PortList = @()

    # Convert it into an array of integers
    foreach($port in $InitPortList) {
        if ($port.Status -eq "LISTENING") {
            $PortList += $port.Port
        }
    }

    $PortList = $PortList | select -Unique

    return $PortList
}
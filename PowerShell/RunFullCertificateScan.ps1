#region Import Modules
Set-Location $PSScriptRoot

# CertificateAdministration Module.
$ModulePath = "$($PSScriptRoot)\Modules\"
$ModuleName = 'CertificateAdministration'
$CertModule= "$($ModulePath)$($ModuleName).psm1"
Import-Module -Force $CertModule
get-module

# ExceptionHandler Module.
$ModulePath = "$($PSScriptRoot)\Modules\"
$ModuleName = 'ExceptionHandler'
$ExceptMod = "$($ModulePath)$($ModuleName).psm1"
Import-Module -Force $ExceptMod
get-module

# IPInformation Module.
$ModulePath = "$($PSScriptRoot)\Modules\"
$ModuleName = 'IPInformation'
$IPMod = "$($ModulePath)$($ModuleName).psm1"
Import-Module -Force $IPMod
get-module


# NetworkCommandParser Module.
$ModulePath = "$($PSScriptRoot)\Modules\"
$ModuleName = 'NetworkCommandParser'
$ParseModule = "$($ModulePath)$($ModuleName).psm1"
Import-Module -Force $ParseModule
get-module
#endregion

#region Functions
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

# Creates a PSCredential object from user input
function Get-PSCredential {
    # Get Username
    $user = Read-Host "Enter Username: "

    # Get Password
    $pw = Read-Host "Enter Password: " -AsSecureString
     #| Out-File ".\Text\Password.txt" 

    # Convert password to secure string
    $secureStringPwd = $pw | ConvertFrom-SecureString | ConvertTo-SecureString 

    # Clear password from local variables
    $pw = "" 

    $cred = New-Object System.Management.Automation.PSCredential -ArgumentList $user, $secureStringPwd

    # Clear secure password from local variables
    $secureStringPwd = ""

    return $cred
}

# Runs the certificate check on the list of ports and returns the certificates
function Get-ServerCertificates ($RootDir, $ServerDir, $ServerName, $Ports, $Timeout) {

    #region Initialize Variables
    $CertInformation = @()   # Array for all certificate checks on TCP connections

    # Set up Report variables
    $HostName  = $ServerName
    $ResultDir = ".\ErrorData\ResultData"
    $RunDate   = (Get-Date -uFormat "%Y%m%d-%H%M")

    # Destination path
    $ExportServerCSVPath = "$($ServerDir)$($ServerName).csv"
    $ExportGeneralCSVPath = "$($RootDir)Certificates.csv"


    # Status array to maintain the progress of the application
    $StatusArray = new-object psobject -Property @{
									                'DateTime' = $RunDate
									                'AllCertificatesScanned' = $false
                                                    'ExcelArrayExported' = $false
								                  }
    #endregion

    try 
    {
        # Scan for a certificate on each port
        foreach($port in $Ports) {
            Write-Progress "Scanning Certificate on $($ServerName):$($port)"
            $CertInformation += Get-CertificateInformation -ComputerName $ServerName -Port $port -Timeout $Timeout
        }
        $StatusArray.AllCertificatesScanned = $true
    }
    catch [Exception]
    {
        Report-Error $RunDate $HostName $ResultDir $StatusArray 
    }

    #region Export Results to Excel
    try 
    {
        # Export server information to it's personal csv and the consolidated csv
        $certs = $CertInformation | Where-object {$_ -ne $null}
        $certs | Export-Csv -Path $ExportServerCSVPath -NoTypeInformation
        $certs | Export-Csv -Path $ExportGeneralCSVPath -NoTypeInformation -Append
        $StatusArray.ExcelArrayExported = $true
    }
    catch [Exception]
    {
        Report-Error $RunDate $HostName $ResultDir $StatusArray
    }
    #endregion


    #region Report Final Status
    Report-Status $StatusArray $ResultDir $HostName
    #endregion
}
#endregion



#region Initialize Variables

# Import the list of servers you will be checking for certificates
$ImportServers = import-csv '.\CSV\Servers.csv'

# Create the root directory where the certificate information will be stored
$RootCertDir = "$($PSScriptRoot)\CertificateResults\"

# Get user credentials for remoting into servers
$Credential = $null

# Test Server used to Validate user credentials
$TestServer = "YO00APPD20.DGSDEV.DGS.CA.GOV"

#endregion


#region Create PS Remoting Credentials
try 
{
    # Try and create a PSSession using the inputted credentials; if it throws an error, exit the program
    $Credential = Get-PSCredential
    $Sess = New-PSSession -ComputerName $TestServer -Credential $Credential -ErrorAction SilentlyContinue
    
    Remove-PSSession $Sess
}
catch
{
    Write-Host "Logon Failure"
    exit
}
#endregion




# Iterate through each server and check it for certificates
foreach ($server in $ImportServers) {
    if (-NOT $server) { continue; }  # Do not continue if server name is null
    
    $ServerName = $server.Server

    # Create directory to store server specific certificates
    $ServerCertDir = "$($RootCertDir)$($ServerName)\"
    if (-NOT (Test-Path $ServerCertDir)){
        New-Item -ItemType directory -Path $ServerCertDir
    }

    # Get IP Addresses and active ports on remote machine
    try 
    {
        $Sess = New-PSSession -ComputerName $ServerName -Credential $Credential -ErrorAction SilentlyContinue
        $IPAddrs = Get-ServerIPAddresses $Sess
        $Ports   = Get-ServerActivePorts $Sess
        Remove-PSSession $Sess
    }
    catch { Write-Host "Error connecting to $($ServerName)" }
    
    if ($IPAddrs -is [array]) {
        # If Get-ServerIPAddresses returned multiple IP Addresses (Multi-homed), find certificates on all of them
        foreach ($addr in $IPAddrs) { Get-ServerCertificates $RootCertDir $ServerCertDir $addr $Ports 2000 }

    } else {
        # If Get-ServerIPAddresses returned a single address, find certificates on it
        Get-ServerCertificates $RootCertDir $ServerCertDir $IPAddrs $Ports 2000
    }
}

# Wipe credentials from local stack
$Credential = ""
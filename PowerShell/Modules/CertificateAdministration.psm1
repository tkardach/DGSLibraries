<#
TO-DO:
Add help attributes to function header comments.
Create/Test SCCM install package.
#>

#region TCP Certificate Scanning

#region Class Definition
# Create class if it has not already been created
function Get-Certificate2AsPSObject ([string]$Machine, [int]$Port, [Security.Cryptography.X509Certificates.X509Certificate2]$Certificate) {
    return New-Object -TypeName PSObject -Property @{
        # Create Certificate Object and add to CertificateList
        IPAddress       = $Machine
        Port            = $Port
        Archived        = $Certificate.Archived
        Extensions      = $Certificate.Extensions
        FriendlyName    = $Certificate.FriendlyName
        Handle          = $Certificate.Handle
        HasPrivateKey   = $Certificate.HasPrivateKey
        Issuer          = $Certificate.Issuer
        IssuerName      = $Certificate.GetIssuerName()
        NotAfter        = $Certificate.NotAfter
        NotBefore       = $Certificate.NotBefore
        PrivateKey      = $Certificate.PrivateKey
        PublicKey       = $Certificate.GetPublicKey()
        RawData         = $Certificate.RawData
        SerialNumber    = $Certificate.SerialNumber
        SignAlgorithm   = $Certificate.SignatureAlgorithm.FriendlyName
        SubjectName     = $Certificate.SubjectName.Name
        Thumbprint      = $Certificate.Thumbprint
        Version         = $Certificate.Version
    }
}

#endregion

function Get-ScannedPorts {
<#
    .SOURCE
        Derived from script posted at:
            https://www.nextofwindows.com/a-simple-network-port-scanner-in-powershell
    
    .SYNOPSIS
        Checks a list of machine names and ports, and scans all ports of each machine for TCP connections.

    .DESCRIPTION
        Checks a list of machine names and ports, and scans all ports of each machine for TCP connections.
        Returns a list containing: Server, Port, and Status of TCP connection.

    .PARAMETER ComputerNames
        Array of all servers to be scanned

    .PARAMETER Ports         
        Array of all ports to be scanned per server

    .PARAMETER Timeout       
        Number of milliseconds to wait before timing out the TCP connection attempt

    .RETURN
        Table containing all ScannedPorts

    .NOTES

    .EXAMPLE
        Get-ScannedPorts -ComputerNames $serverList -Ports $portList -timeout 100

        Server                                                                                                                                                                           Port Status                                                                                   
        ------                                                                                                                                                                           ---- ------                                                                                   
        example1.exampleserver.com                                                                                                                                                         80 Connected                                                                                
        example1.exampleserver.com                                                                                                                                                        443 Connected                                                                                
        example1.exampleserver.com                                                                                                                                                       1433 Connection Timed Out                                                                     
        example1.exampleserver.com                                                                                                                                                        135 Connected                                                                                
        example1.exampleserver.com                                                                                                                                                        139 Connected                                                                                
        example2.exampleserver.com                                                                                                                                                         80 Connected                                                                                
        example2.exampleserver.com                                                                                                                                                        443 Connected                                                                                
        example2.exampleserver.com                                                                                                                                                       1433 Connection Timed Out                                                                     
        example2.exampleserver.com                                                                                                                                                        135 Connected                                                                                
        example2.exampleserver.com                                                                                                                                                        139 Connected                                                                                
        example3.exampleserver.com                                                                                                                                                         80 Connected                                                                                
        example3.exampleserver.com                                                                                                                                                        443 Connected                                                                                
        example3.exampleserver.com                                                                                                                                                       1433 Connection Timed Out                                                                     
        example3.exampleserver.com                                                                                                                                                        135 Connected                                                                                
        example3.exampleserver.com                                                                                                                                                        139 Connected             
        
    .INPUT

    .OUTPUT

    .REMARKS

#>

    #region Parameters
    Param(
            [parameter(mandatory=$true,helpmessage="Array of machine names to scan")]
            [Alias('IPAddrs', 'MachineNames')]
            [string[]]$ComputerNames,
            [parameter(mandatory=$true,helpmessage="Array of ports to scan over the machines")]
            [int[]]$Ports = 443,
            [parameter(mandatory=$false,helpmessage="Timeout for TCP connection")]
            [int]$Timeout = 1000
         )
    #endregion

    #region Begin
    Begin
    {
        # Table of ScannedPorts to be returned
        $scanTable = @()
    }
    #endregion

    #region Process
    Process
    {
        #region Machine/Port Loop
        # Nested for loops check each port on every server
        foreach($server in $ComputerNames) {
            foreach($port in $Ports) {
                #region Try-Block
                $scanTable += Invoke-TestTCPConnection -ComputerName $server -Port $port -Timeout $Timeout
            }
        }
        #endregion
    }
    #endregion

    #return the table of ScannedPorts
    END
    {
        return $scanTable
    }
}


function Get-CertificateInformationTCP {
    <#
    .SOURCE
        Derived from script posted at:
            https://isc.sans.edu/forums/diary/Assessing+Remote+Certificates+with+Powershell/20645/
    
    .SYNOPSIS
        Retrieves a Certificate on a specified port on a remote machine.

    .DESCRIPTION
        Retrieves a Certificate on a specified port on a remote machine. Returns certificate information
        such as: machine name, port number, certificate name, issuer name, issue date, expiration date,
        error message (if there is one).

    .PARAMETER MachineName
        A remote system with we are checking for a certificate.

    .PARAMETER Port
        The port to be checked on the machine we are scanning.

	.PARAMETER TCPTimeout
		Number of milliseconds the program will wait before giving up on a TCP connection.
	
	.PARAMETER SSLTimeout
		Number of milliseconds the program will wait before giving up on the SSL Negotation.
		
	.PARAMETER StatusReport
		A list of the current status of the running 
	
    .NOTES

    .EXAMPLE
        Get-CertificateInformation -MachineName "ex-pc" -Port 443

        Machine    : EX
        Port       : 443
        Name       : CN=EX
        IssuerName : DC=com, DC=example, CN=EX
        IssueDate  : 10/25/2016 4:56:02 PM
        ExpDate    : 10/25/2018 4:56:02 PM
        Error      : 

    .EXAMPLE
        Get-CertificateInformation -MachineName "ex-pc" -Port 135

        Machine    : EX
        Port       : 135
        Name       : 
        IssuerName : 
        IssueDate  : 
        ExpDate    : 
        Error      : Timed Out During Client Authentication

    .INPUTS

    .OUTPUTS

    .REMARKS

    #>
    #region Parameters
    Param (
			[parameter(mandatory=$true,helpmessage="Hostname to resolve remote system")]
            [Alias('IPAddr', 'MachineName')]
			[string]$ComputerName,
			[parameter(mandatory=$true,helpmessage="TCP port number that SSL application is listening on")]
			[int]$Port,
			[parameter(helpmessage="Number of milliseconds to wait before giving up on TCP connection")]
			[int]$TCPTimeout=1000,
			[parameter(helpmessage="Number of milliseconds to wait before giving up on SSL Negotiation")]
			[int]$SSLTimeout=2000
          )
    #endregion

    #region Begin
    Begin
    {
		$machine = $ComputerName;
        $obj = New-Object Certificate
    }
    #endregion

    #region Process
	Process
	{
        #region TCP Connection
		try {
            # Get IPHostEntry from IP or Machine Name
            $hosttask = [System.Net.Dns]::GetHostEntryAsync($machine)
			# Create a timer for IPHostEntry Finding to be established
			$wait = $hosttask.AsyncWaitHandle.WaitOne(1000,$false)
			
			# Begin TCP Connection
			$tcpclient = New-Object -TypeName system.Net.Sockets.TcpClient
			
			if (!$wait) {
				$iar = $tcpclient.BeginConnect($machine, $Port, $null, $null)
			}
			else 
			{
				$machine = $hosttask.Result.HostName
				$iar = $tcpclient.BeginConnect($machine, $Port, $null, $null)
			}
			
			Write-Progress "Establishing TCP Connection $($machine):$($Port)"
			# Create a timer for TCP Connection to be established
			$wait = $iar.AsyncWaitHandle.WaitOne($SSLTimeout,$false)

			if (!$wait)
			{              
                $obj.Machine    = $machine
                $obj.Port       = $Port
                $obj.Status     = "Certificate Not Found"
				$obj.Error      = "Timed Out During TCP Connection Attempt"
			}
			else
			{
				$wait = $iar.AsyncWaitHandle.Close()
                
                #region Certificate-Checking
				try
				{
					# Create ssl stream on existing tcp connection
					$stream = New-Object system.net.security.sslstream($tcpclient.GetStream())
					# Send hostname on cert to try SSL negotiation
					$streamtask = $stream.AuthenticateAsClientAsync($machine)

					# Create a timer for Authentication task
					$wait = $streamtask.AsyncWaitHandle.WaitOne($SSLTimeout, $false)

					if (!$wait)
					{	
                        # Handle Authentication Timeout, add Certificate with error message to CertificateList
                        $obj.Machine    = $machine
                        $obj.Port       = $Port
                        $obj.Status     = "Certificate Not Found"
						$obj.Error      = "Timed Out During Client Authentication"
					}
					else
					{
						# Close Async Wait Handler
						$wait = $streamtask.AsyncWaitHandle.Close()

                        if ($stream.IsAuthenticated)
                        {
						    # Get certificate from SSLStream
						    $cert2 = New-Object system.security.cryptography.x509certificates.x509certificate2($stream.RemoteCertificate)
			  
						    # Create Certificate Object and add to CertificateList
                            $obj.Machine    = $machine
                            $obj.Port       = $Port
						    $obj.Name       = $cert2.GetName()
						    $obj.IssuerName = $cert2.GetIssuerName()
						    $obj.IssueDate  = $cert2.GetEffectiveDateString()
						    $obj.ExpDate    = $cert2.GetExpirationDateString()

                            if((get-date) -gt $cert2.GetExpirationDateString()){
                                $obj.ExpIn      = "Certificate has expired!"
                            }
                            else
                            {
                                $timespan = New-TimeSpan -end $cert2.GetExpirationDateString()
                                $obj.ExpIn      = "{0} Days - {1} Hours - {2} Minutes" -f $timespan.days,$timespan.hours,$timespan.minutes
                            }
                            
                            $obj.Status     = "Certificate Found"
						    $obj.Error      = ""
                        }
                        else
                        {
                            # Create Certificate object with error message and add to CertificateList
	                        $obj.Machine    = $machine
                            $obj.Port       = $Port
                            $obj.Status     = "Certificate Not Found"
					        $obj.Error      = "Client Authentication Failed"
                        }
					}
				}
				# Catch generic errors relating to the SSLStream
				catch
				{
					# Create Certificate object with error message and add to CertificateList
	                $obj.Machine    = $machine
                    $obj.Port       = $Port
                    $obj.Status     = "Certificate Not Found"
					$obj.Error      = $_.Exception.Message
				}
				finally
				{
					# Make sure the SSLStream gets closed
					$stream.Close()
                    $stream.Dispose()
				}
                #endregion
			}
		}
		#endregion	
		# Catch generic error
		catch 
		{
			# Create Certificate object with error message and add to CertificateList         
            $obj.Machine    = $machine
            $obj.Port       = $Port			
            $obj.Status     = "Certificate Not Found"
			$obj.Error      = $_.Exception.Message
		}
		finally
		{
			# Close the TCP Connection
			$tcpclient.close()
		}
        #endregion
	}

    #region End
    End
    {
        return $obj
    }
    #endregion
}


function Get-CertificateInformation {
    <#
    .SOURCE
        Derived from script posted at:
            https://isc.sans.edu/forums/diary/Assessing+Remote+Certificates+with+Powershell/20645/
    
    .SYNOPSIS
        Retrieves a Certificate on a specified port on a remote machine.

    .DESCRIPTION
        Retrieves a Certificate on a specified port on a remote machine. Returns certificate information
        such as: machine name, port number, certificate name, issuer name, issue date, expiration date,
        error message (if there is one).

    .PARAMETER MachineName
        A remote system with we are checking for a certificate.

    .PARAMETER Port
        The port to be checked on the machine we are scanning.

	.PARAMETER TCPTimeout
		Number of milliseconds the program will wait before giving up on a TCP connection.
	
	.PARAMETER SSLTimeout
		Number of milliseconds the program will wait before giving up on the SSL Negotation.
		
	.PARAMETER StatusReport
		A list of the current status of the running 
	
    .NOTES

    .EXAMPLE
        Get-CertificateInformation -MachineName "ex-pc" -Port 443

        Machine    : EX
        Port       : 443
        Name       : CN=EX
        IssuerName : DC=com, DC=example, CN=EX
        IssueDate  : 10/25/2016 4:56:02 PM
        ExpDate    : 10/25/2018 4:56:02 PM
        Error      : 

    .EXAMPLE
        Get-CertificateInformation -MachineName "ex-pc" -Port 135

        Machine    : EX
        Port       : 135
        Name       : 
        IssuerName : 
        IssueDate  : 
        ExpDate    : 
        Error      : Timed Out During Client Authentication

    .INPUTS

    .OUTPUTS

    .REMARKS

    #>
    Param (
			[parameter(mandatory=$true,helpmessage="Hostname to resolve remote system")]
            [Alias('IPAddr', 'MachineName')]
			[string]$ComputerName,
			[parameter(mandatory=$true,helpmessage="TCP port number that SSL application is listening on")]
			[int]$Port,
			[parameter(helpmessage="Number of milliseconds to wait before giving up on SSL Negotiation")]
			[int]$Timeout=2000
          )

    Begin
    {
		$machine = $ComputerName
        $obj = $null
    }

	Process
	{
        Start-Sleep -s .5
        Write-Progress "Checking for Certificate on $($machine):$($Port)"
                
        #region Certificate-Checking
	    try
	    {
            #create our webrequest object for the ssl connection
            $request = [Net.WebRequest]::Create("https://$machine`:$Port")

            # Send hostname on cert to try SSL negotiation
            $requestTask = $request.GetResponseAsync()
            $wait = $requestTask.AsyncWaitHandle.WaitOne($Timeout, $false)

		    if (!$wait)
		    {
                $request.abort()
                $obj = $null
		    }
		    else
		    {
                # Close Async Wait Handler
			    $wait = $requestTask.AsyncWaitHandle.Close()

                if ($cert2 = [Security.Cryptography.X509Certificates.X509Certificate2]$request.ServicePoint.Certificate.Handle)
                {
				    # Create Certificate Object and add to CertificateList
                    $obj = Get-Certificate2AsPSObject -Machine $machine -Port $Port -Certificate $cert2
                }
                else
                {
                    # Create Certificate object with error message and add to CertificateList
	                $obj = $null
                }
		    }
	    }
	    # Catch generic errors relating to the SSLStream
	    catch
	    {
		    # Create Certificate object with error message and add to CertificateList
	        $obj = $null
	    }
	    finally
	    {
		    # Make sure the SSLStream gets closed
            if($request) { $request.abort() }
	    }
    }

    End
    {
        return $obj
    }
}

#endregion
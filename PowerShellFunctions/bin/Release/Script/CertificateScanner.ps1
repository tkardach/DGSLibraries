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
    #region Parameters
    Param (
			[parameter(mandatory=$true,helpmessage="Hostname to resolve remote system")]
            [Alias('IPAddr', 'Machine', 'ComputerName')]
			[string]$MachineName,
			[parameter(mandatory=$true,helpmessage="TCP port number that SSL application is listening on")]
			[int]$Port,
			[parameter(helpmessage="Number of milliseconds to wait before giving up on SSL/TLS Negotiation")]
			[int]$Timeout=2000
          )
    #endregion

    #region Process
	Process
	{
        Write-Progress "Checking for Certificate on $($MachineName):$($Port)"

        #region Certificate-Checking
	    try
	    {
            #create our webrequest object for the ssl connection
            $request = [Net.WebRequest]::Create("https://$MachineName`:$Port")

            # Send hostname on cert to try SSL negotiation
            $requestTask = $request.GetResponseAsync()
            $wait = $requestTask.AsyncWaitHandle.WaitOne($Timeout, $false)

		    if (!$wait)
		    {
                $request.abort();
                return $null
		    }
		    else
		    {
                # Close Async Wait Handler
			    $wait = $requestTask.AsyncWaitHandle.Close()
                if ($cert2 = [Security.Cryptography.X509Certificates.X509Certificate2]$request.ServicePoint.Certificate.Handle)
                {
                    return Get-Certificate2AsPSObject -Machine $MachineName -Port $Port -Certificate $cert2
                }
                else
                {
                    # Create Certificate object with error message and add to CertificateList
	                return $null
                }
		    }
	    }
	    # Catch generic errors relating to the SSLStream
	    catch
	    {
		    # Create Certificate object with error message and add to CertificateList
	        return $null
	    }
	    finally
	    {
		    # Make sure the SSLStream gets closed
            if($request) { $request.abort(); }
	    }
        #endregion
    }
}

Get-CertificateInformation -MachineName "10.4.19.83" -Port 443 -Timeout 2000
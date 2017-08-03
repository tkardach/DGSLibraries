function Get-NetstatInformation
{
	<#
		.SYNOPSIS
			Returns all of the active ports running on the system in an array.

		.DESCRIPTION
			Returns an array containing all of the port numbers that are active along with their status, i.e. "ESTABLISHED", 
			"LISTENING", etc...

		.RETURN
			Array of {Port, Status}

		.NOTES

		.EXAMPLE
			Status : ESTABLISHED
			Port   : 135  

			Status : LISTENING
			Port   : 6129       

			Status : LISTENING
			Port   : 60615    


			Status : CLOSE-WAIT
			Port   : 56284


			Status : CLOSE-WAIT
			Port   : 65436


			Status : TIME-WAIT
			Port   : 5357 


		.INPUT
		
		.OUTPUT
		
		.REMARKS
			[::]:PORTS ARE ALMOST ALWAYS REPEATED, SO THEY'RE NOT CAPTURED
		
	#>
	Param(
		[parameter(helpmessage="Additional option to add to the netstat command (ommit '-')")]
		[string]$Options="na"
	)
	Begin 
	{
		#ARRAY TO HOLD ALL NETSTAT ENTRIES
		$netstatInfo = @()
		
		#OPTIONS FOR THE NETSTAT COMMAND
		$Options = $Options.replace("-","")
		
		#EXECUTE COMMAND AND STORE DATA
		$initialData = netstat -$Options
		
		#START AT THE 5TH LINE OF CONSOLE
		$initialData = $initialData[4..$initialData.Count].Trim()

		#DIFFERENTIATE THE PORTS BY ITS STATE
		$established = $initialData | findstr "ESTABLISHED"
		$listening = $initialData | findstr "LISTENING"
		$close_wait = $initialData | findstr "CLOSE_WAIT"
		$time_wait = $initialData | findstr "TIME_WAIT"
	}

	Process
	{
		if($established -match "TCP    "){
			$tcpData_est = $established.SubString(0,26) |Where-Object {$_ -match "TCP"} |ForEach-Object{$_.Split(":")[1]}
			$established_ports = $tcpData_est | ?{$_ -ne ""} #GO THROUGH THIS ARRAY, CREATE OBJECT FOR EACH PORT AND APPEND THAT OBJECT INTO A NEW ARRAY   
		}    
		

		#MAKE OBJECT TO EXPORT TO CSV
		foreach( $num in $established_ports){
			$est = @{
					Port = $num.Replace(" ", "")
					Status = "ESTABLISHED" 
			}
			$est_obj = New-Object -TypeName PSObject -Property $est
			$netstatInfo += $est_obj
		}

		#GO THROUGH LISTENING PORTS AND EXTRACT THE PORT NUMBER
		if($listening -match "TCP    "){
			$tcpData_list = $listening.SubString(0,26) |Where-Object {$_ -match "TCP"} |ForEach-Object{$_.Split(":")[1]}
			$listening_ports = $tcpData_list| ?{$_ -ne ""} 
		} 
		

		#MAKE OBJECT TO EXPORT TO CSV
		foreach($num in $listening_ports){
			$est = @{
					Port = $num.Replace(" ", "")
					Status = "LISTENING" 
			}
			$lis_obj = New-Object -TypeName PSObject -Property $est
			$netstatInfo += $lis_obj
		}

		#GO THROUGH CLOSE-WAIT PORTS AND EXTRACT THE PORT NUMBER
		if($close_wait -match "TCP    "){
			$tcpData_clos = $close_wait.SubString(0,26) |Where-Object {$_ -match "TCP"} |ForEach-Object{$_.Split(":")[1]}
			$close_wait_ports = $tcpData_clos| ?{$_ -ne ""} 
		}
		

		#MAKE OBJECT TO EXTRACT TO CSV
		$c = @()
		foreach($num in $close_wait_ports){
			$est = @{
					Port = $num.Replace(" ", "")
					Status = "CLOSE-WAIT" 
			}
			$close_obj = New-Object -TypeName PSObject -Property $est
			$netstatInfo += $close_obj
		}

		#GO THROUGH CLOSE-WAIT PORTS AND EXTRACT THE PORT NUMBER
		if($time_wait -match "TCP    "){
			$tcpData_time = $time_wait.SubString(0,26) |Where-Object {$_ -match "TCP"} |ForEach-Object{$_.Split(":")[1]}
			$time_wait_ports = $tcpData_time| ?{$_ -ne ""} 
		}
		
		#GO THROUGH TIME_WAIT PORTS AND EXTRACT THE PORT NUMBER
		foreach($num in $time_wait_ports){
			$est = @{
					Port = $num.Replace(" ", "")
					Status = "TIME-WAIT" 
			}
			$time_obj = New-Object -TypeName PSObject -Property $est
			$netstatInfo += $time_obj
		}
	}
	
	End
	{
		return $netstatInfo
	}
	
}

function Get-IPConfigAllIPv4
{
<#
    .SYNOPSIS
		Returns the IP Addresses from inputting the "ipconfig /all" command in the form of a string array of IPs

    .DESCRIPTION
		Parses the "ipconfig /all" command into an array of IPs

    .RETURN
		Array of IP addresses as strings

    .NOTES

    .EXAMPLE
		11.111.11.11
		11.111.11.12
		11.111.11.13
		11.111.11.14
		11.111.11.15
		11.111.11.16
		11.111.11.17

    .INPUT

    .OUTPUT

    .REMARKS

#>
	Begin
	{
		$iparray = @()
		# Store ipconfig /all in a string
		$ipconfig = ipconfig /all
		# Get rid of everything excpet for IPv4 Addresses
		$ipv4 = $ipconfig | findstr "IPv4" | foreach-object{$_.Split(":")[1]}
		# Remove the "(Preferred)" string
		$serverList = $ipv4 -replace ".{12}$"
		
	}
	
	Process
	{
		$lines = $serverList | Measure-Object -Line
		if ($lines.Lines -ne 1)  {
			for($i=1; $i -lt $serverList.Length; $i++) {
				$iparray += new-object -TypeName psobject -Property @{
                            Server = $serverList[$i].replace(" ", "")
                        }
			}
		}
		else 
		{	
           $iparray += new-object -TypeName psobject -Property @{
                        Server = $serverList.replace(" ", "")
                    }
		    
		}
        
	}
	
	End
	{
		return $iparray
	}
}
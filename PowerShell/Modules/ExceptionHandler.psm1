function Report-Error
{
    <#
        .SYNOPSIS
            Collects exception details and records it to an Excel sheet.
			
        .DESCRIPTION
            
        .PARAMETER 
			
        .PARAMETER
			
        .PARAMETER
			
        .RETURN
            
        .NOTES
			
        .EXAMPLE
            
        .INPUT
			
        .OUTPUT
			
        .REMARKS
			
    #>

	#region Parameters
    Param(
            [parameter(helpmessage="Date and Time of program execution")]
            [string]$RunDate = (get-date -uformat "%Y%m%d-%H%M"),
            [parameter(mandatory=$true,helpmessage="Name of the machine for which the error was thrown")]
            [string]$HostName,
            [parameter(mandatory=$true,helpmessage="Directory where the exception information will be stored")]
            [string]$ResultDirectory,
			[parameter(helpmessage="An array containing the status of the program leading to the exception")]
			[string[]]$StatusArray,
			[parameter(helpmessage="Record for the error")]
			[String[]]$ErrorRecord
         )
    #endregion
	
	Begin
	{
		$ErrorFilePath = "$($ResultDirectory)\$($HostName)-Error.csv"

		$ErrorArray = @()
		$ExceptionValue = $ErrorRecord.Exception
		$FullyQualifiedErrorId = $ErrorRecord.FullyQualifiedErrorId
		$ScriptStackTrace = $ErrorRecord.ScriptStackTrace
	   
		$ErrorArray +=[pscustomobject]@{
										'$DateTime' = $RunDate
										'$Exception' = $ExceptionValue
										'$FullyQualifiedErrorId' = $FullyQualifiedErrorId
										'$ScriptStackTrace' = $ScriptStackTrace
										}
	}

	Process
	{
		for ($i = 0; $Exception; $i++, ($Exception = $Exception.InnerException))
		{   
			$ExceptionValue = $Exception
			$FullyQualifiedErrorId = $ErrorRecord.FullyQualifiedErrorId
			$ScriptStackTrace = $ErrorRecord.ScriptStackTrace

			$ErrorArray +=[pscustomobject]@{
											'$DateTime' = $RunDate
											'$Exception' = $ExceptionValue
											'$FullyQualifiedErrorId' = $FullyQualifiedErrorId
											'$ScriptStackTrace' = $ScriptStackTrace
										   }
		}

		try
		{
			$ErrorArray | export-csv $ErrorFilePath -NoTypeInformation -Append
		}
		catch [Exception]
		{
			Report-Status $statusArray $resultDir $hostName
			exit
		}
		finally
		{

		}
	}

	End
	{
		Report-Status $StatusArray $ResultDirectory $HostName
	}
}

function Report-Status($statusArray,$resultDir,$hostName)
{
	<#
	
	#>
    #region Write status record
    $ResultFilePath = "$($resultDir)\$($hostName).csv"

    try
    {
        $statusArray | export-csv $ResultFilePath -NoTypeInformation -Append
    }
    catch [Exception]
    {
        exit
    }
    finally
    {

    }    
    #endregion
}

function Resolve-Error ($ErrorRecord=$Error[0])
{
	$ErrorRecord | Format-List * -Force
	$ErrorRecord.InvocationInfo |Format-List *
	$Exception = $ErrorRecord.Exception

	for ($i = 0; $Exception; $i++, ($Exception = $Exception.InnerException))
	{   
		"$i" * 80
		$Exception |Format-List * -Force
	}
}
function Send-OutlookEmail{
<#  
    .SYNOPSIS
        Sends an email through Outlook using the user account of the person currently signed in.

    .DESCRIPTION
        Accepts parameters that specify the email details, and sends an email through Outlook
        using that information. The message will be sent from the currently signed in user.

    .PARAMETER To
        Array or string containing the To: recipient(s)

    .PARAMETER CC         
        Array or string containing the CC: recipient(s)

    .PARAMETER BCC       
        Array or string containing the BCC: recipient(s)

    .NOTES
        The string format for sending to multiple recipients is as follows:
            "example@domain.net; example2@domain.net; example3@domain.net"

        If any of the recipient related parameters (To, CC, BCC) are in the form of an array,
        they will be converted to this type of string automatically.                       
        
    .INPUT

    .OUTPUT
        Sends an email.

    .REMARKS
#>

    Param (
        [parameter(mandatory=$true,helpmessage="The To: recipient(s) of the email")]
        [string[]]$To,
        [parameter(helpmessage="The CC: recipient(s) of the email")]
        [string[]]$CC,
        [parameter(helpmessage="The BCC: recipient(s) of the email")]
        [string[]]$BCC,
        [parameter(helpmessage="The Subject of the email")]
        [string]$Subject,
        [parameter(helpmessage="The Body of the email")]
        [string]$Body
    )

    Begin
    {
        # Convert arrays to strings
        $strTo = ""
        if ($To -is [array]) {
            foreach($rec in $To) {
                $strTo = $strTo + "$rec; "
            }
        }

        $strCC = ""
        if ($CC -is [array]) {
            foreach($rec in $CC) {
                $strCC = $strCC + "$rec; "
            }
        }

        $strBCC = ""
        if ($BCC -is [array]) {
            foreach($rec in $BCC) {
                $strBCC = $strBCC + "$rec; "
            }
        }
    }

    Process {
        # Create Outlook object
        $outlook = New-Object -ComObject Outlook.Application

        $mail = $outlook.CreateItem(0)

        # Set email fields
        $mail.To = $strTo
        $mail.CC = $strCC
        $mail.BCC= $strBCC

        $mail.Subject = $Subject
        $mail.Body = $Body

        # Send email
        $mail.Send()
    }

    End {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)
        $outlook = $null
    }
}

function Get-ADUserEmail {
<#  
    .SYNOPSIS
        Retrieves an Active Directory user's email address when given their common name.

    .DESCRIPTION
        Accepts a domain and a common name as parameters. Returns the email address associated
        with the common name if it exists.

    .PARAMETER Domain
        Domain name for the Active Directory server

    .PARAMETER CommonName         
        Common name for the Active Directory user being searched

    .NOTES
        Nothing will be returned if the common name is inputted incorrectly. The domain
        is set to the DGSACCOUNTS domain by default.

    .RETURN
        Email of Active Directory user.
        
    .INPUT
        

    .OUTPUT
        

    .REMARKS

#>
    Param(
        [parameter(helpmessage="Domain name for the Active Directory server")]
        [string]$Domain="dgsaccounts.dgs.ca.gov",
        [parameter(mandatory=$true, helpmessage="Common name for the Active Directory user being searched")]
        [string]$CommonName
    )

    Process {
        # Create a DomainEntry object using the dgsaccounts domain
        $domainEntry = new-object -typename System.DirectoryServices.DirectoryEntry -ArgumentList "LDAP://$Domain"

        # Create a DomainSearcher to search through the dgsaccounts domain
        $domainSearcher = new-object -typename System.DirectoryServices.DirectorySearcher -ArgumentList $domainEntry

        # Set the filter to find the $CommonName parameter
        $domainSearcher.Filter = "(&(objectClass=user)(CN=$CommonName))"

        # Find the user
        $user = $domainSearcher.FindOne()

		[string]$userEmail = $user.Properties["mail"]
        if ($user) {
            return New-Object PSObject -Property @{
				email = $userEmail
			}
		}
    }
}
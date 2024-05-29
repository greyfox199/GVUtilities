function Out-GVLogFile {
     <#
	    .SYNOPSIS
	    Sends data to log file and optional console
	    .DESCRIPTION
	    The function sends data to a log file and optoinally displays to console
	    .PARAMETER objLogFile
	    Mandatory object of logfile to be written to
        .PARAMETER WriteToLog
	    Mandatory boolean controlling whether to write to log file or not
        .PARAMETER LogString
	    Mandatory string of data to write to log
        .PARAMETER LogType
        Mandatory string of log type of log string
        must be one of these values
            -Info
            -Warning
            -Error
            -Debug
        .PARAMETER DisplayInConsole
	    boolean controlling whether to display log string in console or not
	    .EXAMPLE
	    Out-GVLogFile -LogFileObject $objLogFile -WriteToLog $true -LogString "Test log string" -LogType "Warning" -DisplayInConsole $true
	    .INPUTS
	    System.String
	    .OUTPUTS
	    System.PSObject
	    .NOTES
	    TBD.
	    .LINK
	TBD
	#>
    param (
        [Parameter(Mandatory=$true)]
        $LogFileObject,
        [Parameter(Mandatory=$true)]
        [bool]$WriteToLog=$true,
        [Parameter(Mandatory=$true)]
        [string]$LogString,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Info','Warning','Error','Debug')]
        [string]$LogType,
        [Parameter()]
        [bool]$DisplayInConsole=$true
    )
    if ($DisplayInConsole -eq $true) {
        if ($LogType.toLower() -eq "info") {
            write-host $LogString
        }
        if ($LogType.toLower() -eq "warning") {
            write-host $LogString -ForegroundColor Yellow
        }
        if ($LogType.toLower() -eq "error") {
            write-host $LogString -ForegroundColor Red
        }
        if ($LogType.toLower() -eq "debug") {
            write-host $LogString -ForegroundColor Green
        }
    }
    if ($WriteToLog -eq $true) {
        $LogFileObject.writeline($LogString)
        $LogFileObject.flush()
    }
}

function Send-GVMailMessage {
    <#
	    .SYNOPSIS
	    Sends data to log file and optional console
	    .DESCRIPTION
	    The sends a mail message via the Microsoft graph API.  It requires an Azure application registration created with a valid client secret, and "mail.send" api permissions
	    .PARAMETER Sender
	    Mandatory string of sender email addresses of mail message
        .PARAMETER Recipients
	    Mandatory string of recipient email addresses of mail message.  Multiple recipients are separated with a comma
        .PARAMETER Subject
	    Mandatory string of subject of mail message
        .PARAMETER Body
	    Mandatory string of body of mail message
        .PARAMETER ContentType
	    Mandatory string of content type of message
        must be one of these values
            -Text
            -HTML
        .PARAMETER SaveToSentItems
	    boolean flagging whether to save mail message to sent items folder or not
        default value is True
        .PARAMETER TenantID
	    Mandatory string of tenant ID of azure instance
        .PARAMETER AppID
	    Mandatory string  of applicaton ID of azure applicatoin registration
        .PARAMETER Sender
	    Mandatory string of client secret for azure appli8cation rgistration
	    .EXAMPLE
	    Send-GVMailMessage -sender "sender@contoso.com" -TenantID "12345" -AppID "6789" -ClientSecret "guest" -subject "Test Subject" -body "Test Body" -ContentType "HTML" -Recipient "recipient@contoso.com"
	    .INPUTS
	    System.String
	    .OUTPUTS
	    System.PSObject
	    .NOTES
	    TBD.
	    .LINK
	TBD
	#>
    param (
        [Parameter (Mandatory = $true)]
        [String]$Sender,
        [Parameter (Mandatory = $true)]
        [String]$Recipients,
        [Parameter (Mandatory = $true)]
        [String]$Subject,
        [Parameter (Mandatory = $true)]
        [String]$Body,
        [Parameter (Mandatory = $true)]
        [ValidateSet('Text','HTML')]
        [String]$ContentType,
        [Parameter()]
        [bool]$SaveToSentItems=$false,
        [Parameter (Mandatory = $true)]
        [String]$TenantID,
        [Parameter (Mandatory = $true)]
        [String]$AppID,
        [Parameter (Mandatory = $true)]
        [String]$ClientSecret
    )

    $Uri = "https://login.microsoftonline.com/$($TenantID)/oauth2/v2.0/token"
        $PostData = @{
        client_id = $AppID
        scope = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type = "client_credentials"
    }
    $TokenRequest = Invoke-WebRequest -Method Post -Uri $Uri -ContentType "application/x-www-form-urlencoded" -Body $PostData -UseBasicParsing
    $Token = ($tokenRequest.Content | ConvertFrom-Json).access_token

    $Headers = @{
        "Authorization" = "Bearer $($Token)"
    }

    $Uri =  "https://graph.microsoft.com/v1.0/users/$($Sender)/sendMail"

    $RecipientsTemp = $Recipients.split(",")

    if ($RecipientsTemp.Count -eq 1) {
        $RecipientBody = "
        {
            ""emailAddress"": {
                ""address"": ""$($RecipientsTemp)""
            }
        }"
    } else {
        $intCounter = 1
        foreach ($Recipient in $RecipientsTemp) {
            if ($intCounter -eq $RecipientsTemp.Count) {
                $RecipientBody = $RecipientBody + "
                {
                    ""emailAddress"": {
                        ""address"": ""$($Recipient)""
                    }
                }"
            } else {
                $RecipientBody = $RecipientBody + "
                {
                    ""emailAddress"": {
                        ""address"": ""$($Recipient)""
                    }
                },"
            }

            $intCounter = $intCounter + 1
        }
    }

    $message = "{
        ""message"": {
            ""subject"": ""$($Subject)"",
            ""body"": {
                ""contentType"": ""$($ContentType)"",
                ""content"": ""$($Body)""
            },
            ""toRecipients"": [
                $($RecipientBody)
            ]
        },
        ""saveToSentItems"": ""$($SaveToSentItems)""
    }"

    Invoke-RestMethod -Uri $Uri -Headers $Headers -Method POST -Body $message -ContentType "application/json"
}
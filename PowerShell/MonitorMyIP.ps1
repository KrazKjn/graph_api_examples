<#
.SYNOPSIS
A script to monitor a public IP address and send an email to the email associated with a Microsft Account.

.PARAMETER PhoneNumber
The phone number of the user to send a text (Script Modification Required).

.PARAMETER CheckEveryMinutes
The number of minutes to wait to re-check. Must be an integer.
#>

$previousIp = "Unknown"

# Check if the parameters are provided
param (
    [string]$PhoneNumber,
    [int]$CheckEveryMinutes = 30
)

$smsAddress = $null
if ($PhoneNumber -ne $null) {
    $smsAddress = "$PhoneNumber@mms.att.net" # Example: "number@txt.att.net" for AT&T
}

# Install the Microsoft.Graph module if not already installed
# Install-Module -Name Microsoft.Graph -Scope CurrentUser

# Import the module
#Import-Module Microsoft.Graph -Scope Global

# Import only the Users module
Import-Module -Name Microsoft.Graph.Users

# Import only the Mail module
Import-Module -Name Microsoft.Graph.Mail

# Import only the Devices module
Import-Module -Name Microsoft.Graph.DeviceManagement

function Send-Notification {
    param (
        [string]$address,
        [string]$currentIp,
        [string]$previousIp
    )

    $subject = "Public IP Address Change Detected"
    $body = "Your public IP address has changed to: $currentIp"
    if ($previousIp.Length -gt 0) {
        $body = "$body from $previousIp."
    }

    # Define the email message parameters
    $emailMessage = @{
        "message" = @{
            "subject" = $subject
            "body" = @{
                "contentType" = "Text"
                "content" = $body
            }
            "toRecipients" = @(
                @{
                    "emailAddress" = @{
                        "address" = $address
                    }
                }
            )
        }
        "saveToSentItems" = "false"
    }

    # Send the email
    Write-Host "Sending Email to [$address]: $subject"
    Send-MgUserMail -UserId $user.Id -BodyParameter $emailMessage
}

function Format-SecondsAsTime {
    param (
        [int]$totalSeconds
    )
    
    $timeSpan = [System.TimeSpan]::FromSeconds($totalSeconds)
    $formattedTime = $timeSpan.ToString("hh\:mm\:ss")
    return $formattedTime
}

function ShowWaitStatus {

    param (
        [int]$seconds = 10,
        [string]$waitMessage
    )

    $startTime = Get-Date
    $endDate = $startTime.AddSeconds($seconds)
    while ((Get-Date) -lt $endDate) {
        # Calculate the time difference
        $timeDifference = $endDate - (Get-Date)

        $percentComplete = [math]::Round((($seconds - $timeDifference.TotalSeconds) / $seconds) * 100)
        $formattedTime = Format-SecondsAsTime -totalSeconds $timeDifference.TotalSeconds
        Write-Progress -Activity "Countdown Timer" -Status "$formattedTime $waitMessage" -PercentComplete $percentComplete
        Start-Sleep -Seconds 1
    }
}

# Define the required scopes
$scopes = @("User.Read", "Mail.Send")

# Authenticate the user and get an access token
$userLoginResults = Connect-MgGraph -Scopes "User.Read", "Mail.Send"

$user = Get-MgUser -UserId "me"

$waitSeconds = $CheckEveryMinutes * 60

while ($true) {
    try {
        $currentIp = (Invoke-WebRequest -Uri "http://ifconfig.me/ip").Content.Trim()
        if ($currentIp -ne $previousIp) {
            Write-Host "Public IP Address Changed from $previousIp to $currentIp."
            # Send notification
            Send-Notification -address $user.Mail -currentIp $currentIp -previousIp $previousIp
            Send-Notification -address $smsAddress -currentIp $currentIp -previousIp $previousIp
            $previousIp = $currentIp
        }
    } catch {
        Write-Error "Failed to retrieve public IP address: $_"
    }
    ShowWaitStatus -seconds $waitSeconds -waitMessage "until Next IP Address Check"
}

$communicationendpointurl = "your-commservice-instance.region.communication.azure.com" # Update with your communication endpoint URL

$EmailRecipient = "user@example.com"
$emailBody = @"
<html>
<body>
<p>Dear User,</p>
<p>Your action was <b>successful</b>.</p>
<p>Please proceed.</p>
</body>
</html>
"@



$emailSubject = "Subject1234"

function sendemail() {
    $ResourceID = 'https://communication.azure.com'
    $Uri = "$($env:IDENTITY_ENDPOINT)?api-version=2018-02-01&resource=$ResourceID"

    try {
        Write-Output "Getting access token..."
        $AzToken = Invoke-WebRequest -Uri $Uri -Method GET -Headers @{ Metadata = "true" } -UseBasicParsing |
            Select-Object -ExpandProperty Content |
            ConvertFrom-Json |
            Select-Object -ExpandProperty access_token

        Write-Output "Access Token retrieved successfully."
    }
    catch {
        Write-Error "Failed to get access token: $_"
        return
    }

    $uri = "https://$communicationendpointurl/emails:send?api-version=2023-03-31"

    $headers = @{
        "Content-Type"  = "application/json"
        "Authorization" = "Bearer $AzToken"
    }

    $apiResponse = @{
        headers = @{
            id = (New-Guid).Guid
        }
        senderAddress = 'no-reply@example.com'
        content = @{
            subject = $emailSubject
            html    = $emailBody
        }
        recipients = @{
            to = @(
                @{
                    address     = $EmailRecipient
                    displayName = $EmailRecipient
                }
            )
        }
        replyTo = @(
            @{
                address     = "support@example.com"
                displayName = "Support Team"
            }
        )
        userEngagementTrackingDisabled = $true
    }

    $body = $apiResponse | ConvertTo-Json -Depth 10

    try {
        Write-Output "Sending email..."
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body -UseBasicParsing
        Write-Output "Email sent successfully. Response: $response"
    }
    catch {
        Write-Error "Failed to send email: $_"
    }
}

sendemail

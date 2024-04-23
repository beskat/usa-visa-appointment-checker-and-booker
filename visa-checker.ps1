$currentApptDate = "2025-01-14" # Date format must be in YYYY-MM-DD
$userEmail = "john.doe@example.com"
$userPassword = "Password123"
$scheduleId = 123456789
$locationId = 987654321
$senderEmail = "john.doe@example.com"
$senderPassword = "AppPassword456"
$recipientEmail = "jane.smith@example.com"
$baseUrl = "https://ais.usvisa-info.com/en-tr/niv"
$loginUrl = "$baseUrl/users/sign_in"
$dateUrl = "$baseUrl/schedule/$scheduleId/appointment/days/$locationId.json?appointments%5Bexpedite%5D=false"
$continueActionsUrl = "$baseUrl/schedule/$scheduleId/continue_actions"
$bookUrl = "$baseUrl/schedule/$scheduleId/appointment"
$waitMinute = 1

function Send-Email
{
    $emailSubject = "Closest Appointment Date Alert"
    $emailBody = "The closest appointment date is $closestDate, which is less than $currentApptDate."
    $smtpServer = "smtp.gmail.com"
    $smtpPort = 587

    Send-MailMessage -From $senderEmail `
                 -To $recipientEmail `
                 -Subject $emailSubject `
                 -Body $emailBody `
                 -SmtpServer $smtpServer `
                 -Port $smtpPort `
                 -Credential (New-Object System.Management.Automation.PSCredential($senderEmail, (ConvertTo-SecureString $senderPassword -AsPlainText -Force))) `
                 -UseSSL
}

function Get-CsrfToken
{
    param ([string] $Content)

    $Content -match '<meta name="csrf-token" content="([^"]+)"' | Out-Null
    return $matches[1]
}

while ($true)
{
    $firstResponse = Invoke-WebRequest -Uri $loginUrl `
    -Headers @{
        "User-Agent" = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
        "Accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7"
        "Accept-Encoding" = "gzip, deflate, br"
        "Accept-Language" = "en-US,en;q=0.9"
        "Cache-Control" = "no-cache"
        "Origin" = "https://ais.usvisa-info.com"
        "DNT" = "1"
        "Pragma" = "no-cache"
        "Sec-Fetch-Dest" = "document"
        "Sec-Fetch-Mode" = "navigate"
        "Sec-Fetch-Site" = "none"
        "Sec-Fetch-User" = "?1"
    }

    $csrfToken = Get-CsrfToken -Content $firstResponse.Content

    $loginRequestHeaders = @{
        "User-Agent" = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
        "Accept" = "*/*;q=0.5, text/javascript, application/javascript, application/ecmascript, application/x-ecmascript"
        "Accept-Encoding" = "gzip, deflate, br"
        "Accept-Language" = "en-US,en;q=0.9"
        "Cache-Control" = "no-cache"
        "DNT" = "1"
        "Cookie" = $firstResponse.Headers['Set-Cookie'] -join '; '
        "Origin" = "https://ais.usvisa-info.com"
        "Pragma" = "no-cache"
        "Referer" = $loginUrl
        "Sec-Fetch-Dest" = "empty"
        "Sec-Fetch-Mode" = "cors"
        "Sec-Fetch-Site" = "same-origin"
        "X-CSRF-Token" = $csrfToken
        "X-Requested-With" = "XMLHttpRequest"
    }

    $loginResponse = Invoke-WebRequest -UseBasicParsing -Uri $loginUrl `
    -Method POST `
    -Headers $loginRequestHeaders `
    -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
    -Body @{
        'user[email]' = $userEmail
        'user[password]' = $userPassword
        'policy_confirmed' = '1'
        'commit' = 'Sign In'
    }

    $dateResponse = Invoke-WebRequest -UseBasicParsing -Uri $dateUrl `
    -Headers @{
        "User-Agent" = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
        "Cookie" = $loginResponse.Headers['Set-Cookie'] -join '; '
        "X-Requested-With" = "XMLHttpRequest"
    }

    if ($dateResponse.StatusCode -ne 200)
    {
        Write-Output "$( Get-Date ) - HTTP error $( $dateResponse.StatusCode ): $( $dateResponse.StatusDescription )"
        Start-Sleep -Seconds ($waitMinute * 60)
        continue
    }

    $dateData = $dateResponse.Content | ConvertFrom-Json

    $closestDate = ($dateData.date | Sort-Object { [DateTime]$_ } | Select-Object -First 1)

    if ( [string]::IsNullOrEmpty($closestDate))
    {
        Write-Output "$( Get-Date ) - No dates available."
        Start-Sleep -Seconds ($waitMinute * 60)
        continue
    }
    Write-Output "$( Get-Date ) - The closest date is: $closestDate"

    if ([DateTime]::Parse($closestDate) -lt [DateTime]::Parse($currentApptDate))
    {
        Send-Email

        $timeUrl = "$baseUrl/schedule/$scheduleId/appointment/times/$locationId.json?date=$closestDate&appointments%5Bexpedite%5D=false"
        $timeResponse = Invoke-WebRequest -UseBasicParsing -Uri $timeUrl`
                    -Headers @{
            "User-Agent" = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
            "Cookie" = $loginResponse.Headers['Set-Cookie'] -join '; '
            "X-Requested-With" = "XMLHttpRequest"
        } | Select-Object -ExpandProperty Content | ConvertFrom-Json

        $lastAvailableTime = $timeResponse.available_times[-1]
        Write-Output "Last available time: $lastAvailableTime"

        $continueActionsResponse = Invoke-WebRequest -UseBasicParsing -Uri $continueActionsUrl `
                    -Headers @{
            "User-Agent" = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
            "Cookie" = $loginResponse.Headers['Set-Cookie'] -join '; '
        }

        $csrfTokenActions = Get-CsrfToken -Content $continueActionsResponse.Content

        if (-not [string]::IsNullOrEmpty($lastAvailableTime) -and -not [string]::IsNullOrEmpty($csrfTokenActions))
        {

            $bookResponse = Invoke-WebRequest -UseBasicParsing -Uri $bookUrl `
                        -Method POST `
                        -Headers @{
                "User-Agent" = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
                "X-CSRF-Token" = $csrfTokenActions
                "Cookie" = $continueActionsResponse.Headers['Set-Cookie'] -join '; '
            } `
                        -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
                        -Body @{
                "authenticity_token" = $csrfTokenActions
                "confirmed_limit_message" = "1"
                "use_consulate_appointment_capacity" = "true"
                "appointments[consulate_appointment][facility_id]" = $locationId
                "appointments[consulate_appointment][date]" = $closestDate
                "appointments[consulate_appointment][time]" = $lastAvailableTime
            }
            break
        }
    }
    Start-Sleep -Seconds ($waitMinute * 60)
}
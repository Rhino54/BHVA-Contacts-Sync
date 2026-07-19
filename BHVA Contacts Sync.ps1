Import-Module Microsoft.Playwright

###########################################################################
# PURE POWERSHELL TOTP GENERATOR (NO NUGET, NO DLLS)
###########################################################################

function Get-TotpCode {
    param([string]$Base32Secret)

    $alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"
    $bytes = New-Object System.Collections.Generic.List[byte]

    $buffer = 0
    $bitsLeft = 0

    foreach ($c in $Base32Secret.ToUpper()) {
        $val = $alphabet.IndexOf($c)
        if ($val -lt 0) { continue }

        $buffer = ($buffer -shl 5) -bor $val
        $bitsLeft += 5

        if ($bitsLeft -ge 8) {
            $bitsLeft -= 8
            $bytes.Add([byte]($buffer -shr $bitsLeft))
            $buffer = $buffer -band ((1 -shl $bitsLeft) - 1)
        }
    }

    $key = $bytes.ToArray()

    $counter = [Math]::Floor((Get-Date -UFormat %s) / 30)
    $counterBytes = [BitConverter]::GetBytes([UInt64]$counter)
    [Array]::Reverse($counterBytes)

    $hmac = New-Object System.Security.Cryptography.HMACSHA1
    $hmac.Key = $key
    $hash = $hmac.ComputeHash($counterBytes)

    $offset = $hash[19] -band 0x0F
    $binary =
        (($hash[$offset] -band 0x7F) -shl 24) -bor
        (($hash[$offset+1] -band 0xFF) -shl 16) -bor
        (($hash[$offset+2] -band 0xFF) -shl 8) -bor
        (($hash[$offset+3] -band 0xFF))

    return $binary % 1000000
}

###########################################################################
# CONFIGURATION — HARD‑CODED PATHS AND CREDENTIALS
###########################################################################

# Hotmail credentials (hard‑code or replace with your own)
$HOTMAIL_USERNAME = "YOUR_HOTMAIL_USERNAME"
$HOTMAIL_PASSWORD = "YOUR_HOTMAIL_PASSWORD"
$HOTMAIL_TOTP_SECRET = "YOUR_TOTP_SECRET"

# BHVA export file (hard‑coded folder you provided)
$BHVA_CSV = "C:\Users\hotpi\OneDrive\Documents\Contacts\BHVA.csv"

# Microsoft Graph App Registration (hard‑code your values)
$CLIENT_ID     = "YOUR_CLIENT_ID"
$CLIENT_SECRET = "YOUR_CLIENT_SECRET"
$TENANT_ID     = "YOUR_TENANT_ID"

###########################################################################
# LOGIN TO HOTMAIL USING PLAYWRIGHT
###########################################################################

$pw = Initialize-Playwright
$browser = $pw.Chromium.LaunchAsync(@{ headless = $false }).Result
$page = $browser.NewPageAsync().Result

$page.GotoAsync("https://login.live.com").Wait()

$page.FillAsync("#i0116", $HOTMAIL_USERNAME).Wait()
$page.ClickAsync("#idSIButton9").Wait()

$page.FillAsync("#i0118", $HOTMAIL_PASSWORD).Wait()
$page.ClickAsync("#idSIButton9").Wait()

$totp = Get-TotpCode -Base32Secret $HOTMAIL_TOTP_SECRET
$page.WaitForSelectorAsync("input[type='tel']").Wait()
$page.FillAsync("input[type='tel']", $totp).Wait()
$page.ClickAsync("#idSubmit_SA_Confirm").Wait()

Write-Host "Logged into Hotmail."

###########################################################################
# EXPORT HOTMAIL CONTACTS
###########################################################################

$page.GotoAsync("https://outlook.live.com/people/").Wait()
$page.ClickAsync("button[aria-label='Manage']").Wait()
$page.ClickAsync("text=Export contacts").Wait()

Start-Sleep -Seconds 5

$downloadFolder = "$env:USERPROFILE\Downloads"
$hotmailFile = Get-ChildItem $downloadFolder -Filter "*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

$hotmailContacts = Import-Csv $hotmailFile.FullName

###########################################################################
# LOAD BHVA CONTACTS (ANY FORMAT)
###########################################################################

$bhvaContacts = Import-Csv $BHVA_CSV

###########################################################################
# AUTO-DETECT EMAIL COLUMN
###########################################################################

$emailColumn = ($bhvaContacts[0].PSObject.Properties.Name | Where-Object {
    $_ -match "email"
})[0]

if (-not $emailColumn) {
    throw "BHVA CSV does not contain an email column."
}

###########################################################################
# AUTO-DETECT NAME COLUMNS
###########################################################################

$firstColumn = ($bhvaContacts[0].PSObject.Properties.Name | Where-Object {
    $_ -match "first"
})[0]

$lastColumn = ($bhvaContacts[0].PSObject.Properties.Name | Where-Object {
    $_ -match "last"
})[0]

###########################################################################
# BUILD HOTMAIL EMAIL HASHSET
###########################################################################

$hotmailEmails = $hotmailContacts.EmailAddress | ForEach-Object { $_.ToLower() }

###########################################################################
# FILTER BHVA CONTACTS THAT ARE NOT IN HOTMAIL
###########################################################################

$newContacts = $bhvaContacts | Where-Object {
    $_.$emailColumn -and ($hotmailEmails -notcontains $_.$emailColumn.ToLower())
}

Write-Host "New contacts to add: $($newContacts.Count)"

###########################################################################
# ADD NEW CONTACTS TO HOTMAIL USING GRAPH API
###########################################################################

$token = (Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TENANT_ID/oauth2/v2.0/token" -Body @{
    client_id     = $CLIENT_ID
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $CLIENT_SECRET
    grant_type    = "client_credentials"
}).access_token

foreach ($c in $newContacts) {

    $body = @{
        givenName      = $firstColumn ? $c.$firstColumn : $null
        surname        = $lastColumn  ? $c.$lastColumn  : $null
        emailAddresses = @(@{ address = $c.$emailColumn })
    } | ConvertTo-Json

    Invoke-RestMethod -Method Post `
        -Uri "https://graph.microsoft.com/v1.0/users/$HOTMAIL_USERNAME/contacts" `
        -Headers @{ Authorization = "Bearer $token" } `
        -Body $body `
        -ContentType "application/json"

    Write-Host "Added: $($c.$emailColumn)"
}

Write-Host "BHVA sync complete — Hotmail contacts preserved."

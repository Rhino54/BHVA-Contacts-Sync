param(
    [switch]$DryRun
)

# =========================
# Config
# =========================

$TenantId        = $env:GRAPH_TENANT_ID
$ClientId        = $env:GRAPH_CLIENT_ID
$ClientSecret    = $env:GRAPH_CLIENT_SECRET

$HotmailUser     = $env:HOTMAIL_IMAP_USERNAME
$HotmailPassword = $env:HOTMAIL_IMAP_PASSWORD

$BhvUser1        = "scortese@bristolharbourvillage.org"
$BhvUser2        = "bod.bhva@bristolharbourvillage.org"

$BhvaCategory    = "BHVA"

# =========================
# REST auth – app-only token for BHV
# =========================

function Get-GraphToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    $response = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body
    return $response.access_token
}

$AccessToken = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
$AuthHeader  = @{ Authorization = "Bearer $AccessToken" }

# =========================
# IMAP helpers – Hotmail Contacts (fail-fast, F2-Loose, C1)
# =========================

function Connect-ImapHotmail {
    param(
        [string]$Username,
        [string]$Password
    )

    try {
        # TCP connect
        $client = New-Object System.Net.Sockets.TcpClient
        $client.Connect("imap-mail.outlook.com", 993)

        # SSL stream with forced TLS 1.2
        $sslStream = New-Object System.Net.Security.SslStream(
            $client.GetStream(),
            $false,
            { param($sender,$cert,$chain,$errors) return $true }
        )

        $sslStream.AuthenticateAsClient(
            "imap-mail.outlook.com",
            $null,
            [System.Security.Authentication.SslProtocols]::Tls12,
            $false
        )

        $reader = New-Object System.IO.StreamReader($sslStream)
        $writer = New-Object System.IO.StreamWriter($sslStream)
        $writer.AutoFlush = $true

        # Server greeting
        $greet = $reader.ReadLine()
        if (-not $greet) {
            throw "IMAP server did not send a greeting."
        }

        # LOGIN
        $writer.WriteLine("a1 LOGIN $Username $Password")
        $loginResp = $reader.ReadLine()

        if ($loginResp -notmatch "a1 OK") {
            throw "IMAP LOGIN failed: $loginResp"
        }

        # SELECT Contacts folder
        $writer.WriteLine('a2 SELECT "Contacts"')
        while ($true) {
            $line = $reader.ReadLine()
            if (-not $line) { throw "IMAP SELECT failed: no response." }
            if ($line -match '^a2 OK') { break }
            if ($line -match '^a2 NO' -or $line -match '^a2 BAD') {
                throw "IMAP SELECT Contacts failed: $line"
            }
        }

        return [pscustomobject]@{
            Client = $client
            Stream = $sslStream
            Reader = $reader
            Writer = $writer
        }
    }
    catch {
        throw "IMAP connection failed: $($_.Exception.Message)"
    }
}

function Convert-VCardToContactObject {
    param(
        [string]$VCard
    )

    $lines = $VCard -split "`r?`n"

    $displayName = $null
    $emails      = @()
    $phones      = @()
    $notes       = $null
    $categories  = @()

    foreach ($line in $lines) {
        if ($line -like "FN:*") {
            $displayName = $line.Substring(3)
        } elseif ($line -like "EMAIL*:*") {
            $emails += $line.Split(":")[1]
        } elseif ($line -like "TEL*:*") {
            $phones += $line.Split(":")[1]
        } elseif ($line -like "NOTE:*") {
            $notes = $line.Substring(5)
        } elseif ($line -like "CATEGORIES:*") {
            $rawCats = $line.Substring(11)
            # split on comma or semicolon
            $categories = $rawCats -split "[,;]" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        }
    }

    return [pscustomobject]@{
        id             = $null
        displayName    = $displayName
        emailAddresses = $emails | Where-Object { $_ } | ForEach-Object { @{ address = $_ } }
        mobilePhone    = $phones.Count -gt 0 ? $phones[0] : $null
        businessPhones = $phones.Count -gt 1 ? $phones[1..($phones.Count-1)] : @()
        homePhones     = @()
        personalNotes  = $notes
        categories     = $categories
        parentFolderId = "HotmailContacts"
    }
}

function Get-HotmailContactsViaIMAP {
    param(
        [string]$Username,
        [string]$Password
    )

    $session = Connect-ImapHotmail -Username $Username -Password $Password
    $reader  = $session.Reader
    $writer  = $session.Writer

    $contacts = @()

    # FETCH all messages in Contacts as vCards
    $writer.WriteLine("a3 FETCH 1:* BODY[]")
    while ($true) {
        $line = $reader.ReadLine()
        if (-not $line) { break }
        if ($line -match '^a3 ') { break }

        if ($line -match '^\* \d+ FETCH') {
            # Next lines until a line with only ")" contain the vCard
            $vcard = New-Object System.Text.StringBuilder
            while ($true) {
                $bodyLine = $reader.ReadLine()
                if ($bodyLine -eq ")") { break }
                [void]$vcard.AppendLine($bodyLine)
            }

            $vc = $vcard.ToString()
            if ($vc -match 'BEGIN:VCARD') {
                $contact = Convert-VCardToContactObject -VCard $vc

                # F2-Loose + C1: include any vCard whose categories contain "BHVA" (case-insensitive)
                $hasBhva = $false
                foreach ($cat in $contact.categories) {
                    if ($cat -and ($cat.ToLower().Contains("bhva"))) {
                        $hasBhva = $true
                        break
                    }
                }

                if ($hasBhva) {
                    $contacts += $contact
                }
            }
        }
    }

    # LOGOUT
    $writer.WriteLine("a4 LOGOUT")
    $null = $reader.ReadLine()

    $session.Stream.Dispose()
    $session.Client.Close()

    return $contacts
}

# =========================
# Helpers
# =========================

function Normalize-Phone {
    param([string]$Number)

    if (-not $Number) { return $null }

    $digits = ($Number -replace '[^\d]', '')
    if (-not $digits) { return $null }

    if ($digits.Length -eq 10) {
        return "+1$digits"
    }

    if ($digits.StartsWith("1") -and $digits.Length -eq 11) {
        return "+$digits"
    }

    return "+$digits"
}

function Build-MatchKey {
    param(
        [string]$DisplayName,
        [string[]]$Emails
    )

    $name = ($DisplayName ?? "").Trim().ToLower()

    if ($Emails -and $Emails.Count -gt 0) {
        $primaryEmail = $Emails[0].Trim().ToLower()
        return "$primaryEmail|$name"
    }

    return $name
}

function Get-ContactsForUser {
    param([string]$UserId)

    $contacts = @()
    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/contacts`?$top=100"

    while ($uri) {
        $page = Invoke-RestMethod -Method GET -Uri $uri -Headers $AuthHeader
        if ($page.value) {
            $contacts += $page.value
        }
        $uri = $page.'@odata.nextLink'
    }

    return $contacts
}

# =========================
# 1. Read contacts
# =========================

$hotmailContacts = Get-HotmailContactsViaIMAP -Username $HotmailUser -Password $HotmailPassword
$bhv1Contacts    = Get-ContactsForUser -UserId $BhvUser1
$bhv2Contacts    = Get-ContactsForUser -UserId $BhvUser2

# =========================
# 2. Filter BHVA from Hotmail (already filtered by IMAP, but keep for safety)
# =========================

$hotmailBhva = $hotmailContacts | Where-Object {
    $_.categories -and ($_.categories | Where-Object { $_.ToLower().Contains("bhva") })
}

# =========================
# 3. Treat all BHV contacts as BHVA
# =========================

$allBhvContacts = @()
$allBhvContacts += $bhv1Contacts
$allBhvContacts += $bhv2Contacts

# =========================
# 4. Build unified BHVA set
# =========================

$unifiedBhva = @{}
$sourceContacts = @()
$sourceContacts += $hotmailBhva
$sourceContacts += $allBhvContacts

foreach ($c in $sourceContacts) {
    $emails = @()
    if ($c.emailAddresses) {
        $emails = $c.emailAddresses | ForEach-Object { $_.address }
    }

    $key = Build-MatchKey -DisplayName $c.displayName -Emails $emails

    if (-not $unifiedBhva.ContainsKey($key)) {
        $unifiedBhva[$key] = [ordered]@{
            SourceContacts = @($c)
        }
    } else {
        $unifiedBhva[$key].SourceContacts += $c
    }
}

function Merge-Contacts {
    param([object[]]$Contacts)

    $hotmail = $Contacts | Where-Object { $_.parentFolderId -eq "HotmailContacts" }
    if ($hotmail.Count -gt 0) {
        $base = $hotmail[0]
    } else {
        $base = $Contacts | Sort-Object {
            ($_.displayName, $_.companyName, $_.jobTitle, $_.mobilePhone, $_.businessPhones, $_.homePhones, $_.emailAddresses) |
                Where-Object { $_ } |
                Measure-Object
        } -Descending | Select-Object -First 1
    }

    $merged = [ordered]@{
        id             = $base.id
        displayName    = $base.displayName
        companyName    = $base.companyName
        jobTitle       = $base.jobTitle
        mobilePhone    = $base.mobilePhone
        businessPhones = $base.businessPhones
        homePhones     = $base.homePhones
        categories     = $base.categories
        personalNotes  = $base.personalNotes
        emailAddresses = $base.emailAddresses
    }

    $allEmails = @()
    foreach ($c in $Contacts) {
        if ($c.emailAddresses) {
            $allEmails += $c.emailAddresses | ForEach-Object { $_.address }
        }
    }
    $allEmails = $allEmails | Where-Object { $_ } | Select-Object -Unique
    $merged.emailAddresses = $allEmails | ForEach-Object {
        @{ address = $_ }
    }

    $allPhones = @()
    foreach ($c in $Contacts) {
        if ($c.mobilePhone)      { $allPhones += $c.mobilePhone }
        if ($c.businessPhones)   { $allPhones += $c.businessPhones }
        if ($c.homePhones)       { $allPhones += $c.homePhones }
    }

    $normalized = $allPhones |
        Where-Object { $_ } |
        ForEach-Object { Normalize-Phone $_ } |
        Where-Object { $_ } |
        Select-Object -Unique

    if ($normalized.Count -gt 0) {
        $merged.mobilePhone    = $normalized[0]
        if ($normalized.Count -gt 1) {
            $merged.businessPhones = $normalized[1..($normalized.Count-1)]
        } else {
            $merged.businessPhones = @()
        }
    }

    $allNotes = @()
    foreach ($c in $Contacts) {
        if ($c.personalNotes) { $allNotes += $c.personalNotes }
    }
    if ($allNotes.Count -gt 0) {
        $merged.personalNotes = "Merged notes from $($allNotes.Count) source contact(s)."
    }

    $cats = @()
    foreach ($c in $Contacts) {
        if ($c.categories) { $cats += $c.categories }
    }
    $cats += $BhvaCategory
    $merged.categories = $cats | Where-Object { $_ } | Select-Object -Unique

    $merged.SourceSummary = @{
        Count     = $Contacts.Count
        Mailboxes = ($Contacts | ForEach-Object { $_.parentFolderId }) -join ";"
    }

    return $merged
}

$mergedBhvaContacts = @()

foreach ($entry in $unifiedBhva.GetEnumerator()) {
    $mergedBhvaContacts += Merge-Contacts -Contacts $entry.Value.SourceContacts
}

# =========================
# REST helpers for BHV write operations
# =========================

function New-GraphContact {
    param(
        [string]$UserId,
        [hashtable]$Contact
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/contacts"
    $body = $Contact | ConvertTo-Json -Depth 6
    Invoke-RestMethod -Method POST -Uri $uri -Headers $AuthHeader -Body $body -ContentType "application/json"
}

function Update-GraphContact {
    param(
        [string]$UserId,
        [string]$ContactId,
        [hashtable]$Contact
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/contacts/$ContactId"
    $body = $Contact | ConvertTo-Json -Depth 6
    Invoke-RestMethod -Method PATCH -Uri $uri -Headers $AuthHeader -Body $body -ContentType "application/json"
}

function Remove-GraphContact {
    param(
        [string]$UserId,
        [string]$ContactId
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/contacts/$ContactId"
    Invoke-RestMethod -Method DELETE -Uri $uri -Headers $AuthHeader
}

# =========================
# 5. Apply to Hotmail (read-only SPOT)
# =========================

$changes = @()

# =========================
# 6. Apply to BHV mailboxes
# =========================

function Build-GraphPayloadFromMerged {
    param(
        [object]$m
    )

    # Ensure arrays are valid for Graph
    $emails = @()
    if ($m.emailAddresses) {
        $emails = $m.emailAddresses |
            Where-Object { $_.address } |
            ForEach-Object { @{ address = $_.address } }
    }

    $businessPhones = @()
    if ($m.businessPhones) {
        $businessPhones = $m.businessPhones | Where-Object { $_ }
    }

    $homePhones = @()
    if ($m.homePhones) {
        $homePhones = $m.homePhones | Where-Object { $_ }
    }

    $categories = @()
    if ($m.categories) {
        $categories = $m.categories | Where-Object { $_ }
    }

    return @{
        displayName    = $m.displayName
        companyName    = $m.companyName
        jobTitle       = $m.jobTitle
        mobilePhone    = $m.mobilePhone
        businessPhones = $businessPhones
        homePhones     = $homePhones
        categories     = $categories
        personalNotes  = $m.personalNotes
        emailAddresses = $emails
    }
}

function Sync-ToBhv {
    param(
        [string]$UserId,
        [object[]]$MergedBhva,
        [object[]]$ExistingContacts
    )

    $index = @{}
    foreach ($c in $ExistingContacts) {
        $emails = @()
        if ($c.emailAddresses) {
            $emails = $c.emailAddresses | ForEach-Object { $_.address }
        }
        $key = Build-MatchKey -DisplayName $c.displayName -Emails $emails
        if (-not $index.ContainsKey($key)) {
            $index[$key] = $c
        }
    }

    foreach ($m in $MergedBhva) {
        $emails = @()
        if ($m.emailAddresses) {
            $emails = $m.emailAddresses | ForEach-Object { $_.address }
        }
        $key = Build-MatchKey -DisplayName $m.displayName -Emails $emails

        $phones = @()
        if ($m.mobilePhone)      { $phones += $m.mobilePhone }
        if ($m.businessPhones)   { $phones += $m.businessPhones }

        if ($index.ContainsKey($key)) {
            $existing = $index[$key]

            $changes += [pscustomobject]@{
                Action          = "UpdateBhv"
                Mailbox         = $UserId
                Key             = $key
                Name            = $m.displayName
                Emails          = ($emails -join ", ")
                Phones          = ($phones -join ", ")
                Categories      = ($m.categories -join ", ")
                NotesSummary    = $m.personalNotes
                SourceCount     = $m.SourceSummary.Count
                SourceMailboxes = $m.SourceSummary.Mailboxes
            }

            if (-not $DryRun) {
                $payload = Build-GraphPayloadFromMerged -m $m
                Update-GraphContact -UserId $UserId -ContactId $existing.id -Contact $payload
            }
        } else {
            $changes += [pscustomobject]@{
                Action          = "CreateBhv"
                Mailbox         = $UserId
                Key             = $key
                Name            = $m.displayName
                Emails          = ($emails -join ", ")
                Phones          = ($phones -join ", ")
                Categories      = ($m.categories -join ", ")
                NotesSummary    = $m.personalNotes
                SourceCount     = $m.SourceSummary.Count
                SourceMailboxes = $m.SourceSummary.Mailboxes
            }

            if (-not $DryRun) {
                $payload = Build-GraphPayloadFromMerged -m $m
                New-GraphContact -UserId $UserId -Contact $payload
            }
        }
    }

    foreach ($c in $ExistingContacts) {
        $emails = @()
        if ($c.emailAddresses) {
            $emails = $c.emailAddresses | ForEach-Object { $_.address }
        }
        $key = Build-MatchKey -DisplayName $c.displayName -Emails $emails

        $existsInUnified = $MergedBhva | Where-Object {
            $me = @()
            if ($_.emailAddresses) {
                $me = $_.emailAddresses | ForEach-Object { $_.address }
            }
            (Build-MatchKey -DisplayName $_.displayName -Emails $me) -eq $key
        }

        if (-not $existsInUnified) {
            $changes += [pscustomobject]@{
                Action          = "DeleteBhv"
                Mailbox         = $UserId
                Key             = $key
                Name            = $c.displayName
                Emails          = ($emails -join ", ")
                Phones          = ($c.mobilePhone, ($c.businessPhones -join ";")) -join ", "
                Categories      = ($c.categories -join ", ")
                NotesSummary    = $c.personalNotes
                SourceCount     = 1
                SourceMailboxes = $c.parentFolderId
            }

            if (-not $DryRun) {
                Remove-GraphContact -UserId $UserId -ContactId $c.id
            }
        }
    }
}

Sync-ToBhv -UserId $BhvUser1 -MergedBhva $mergedBhvaContacts -ExistingContacts $bhv1Contacts
Sync-ToBhv -UserId $BhvUser2 -MergedBhva $mergedBhvaContacts -ExistingContacts $bhv2Contacts

# =========================
# 7. Logging
# =========================

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$logPath   = "sync-log-$timestamp.csv"

if ($changes -and $changes.Count -gt 0) {
    $changes | Export-Csv -Path $logPath -NoTypeInformation
} else {
    "" | Export-Csv -Path $logPath -NoTypeInformation
}

Write-Host "Sync complete. Changes logged to $logPath"
if ($DryRun) {
    Write-Host "Dry run mode: no changes were written."
}

# =========================
# 8. Local save on Hotpi (Windows)
# =========================

$RunningOnWindows = $PSVersionTable.OS -match "Windows"

if ($RunningOnWindows) {
    $localDir = "C:\Users\hotpi\Documents\BHVA Not Shared\BHVA Sync Logs"
    if (-not (Test-Path $localDir)) {
        New-Item -ItemType Directory -Path $localDir | Out-Null
    }

    $localPath = Join-Path $localDir (Split-Path $logPath -Leaf)
    Copy-Item $logPath $localPath -Force

    Write-Host "Local log copy saved to $localPath"
}

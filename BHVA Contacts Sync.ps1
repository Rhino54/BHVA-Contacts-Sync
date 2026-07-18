param(
    [switch]$DryRun
)

# =========================
# Config
# =========================

$TenantId        = $env:GRAPH_TENANT_ID
$ClientId        = $env:GRAPH_CLIENT_ID
$ClientSecret    = $env:GRAPH_CLIENT_SECRET

$HotmailUser     = "hotpie1@hotmail.com"
$BhvUser1        = "scortese@bristolharbourvillage.org"
$BhvUser2        = "bod.bhva@bristolharbourvillage.org"

$BhvaCategory    = "BHVA"

# =========================
# Connect to Graph
# =========================

# Ensure Microsoft.Graph PowerShell module is installed (Ubuntu runners need this)
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Force -Scope CurrentUser
}

Import-Module Microsoft.Graph -Force

$SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force

# Correct Kiota credential object (works on Windows + Ubuntu)
$Cred = New-MgClientSecretCredential `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -ClientSecret $SecureSecret

Connect-MgGraph -ClientSecretCredential $Cred -Scopes "Contacts.ReadWrite", "User.Read.All"

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
    $page = Get-MgUserContact -UserId $UserId -PageSize 100

    $contacts += $page

    while ($null -ne $page.NextLink) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $page.NextLink
        $contacts += $page.value
    }

    return $contacts
}

# =========================
# 1. Read contacts
# =========================

$hotmailContacts = Get-ContactsForUser -UserId $HotmailUser
$bhv1Contacts    = Get-ContactsForUser -UserId $BhvUser1
$bhv2Contacts    = Get-ContactsForUser -UserId $BhvUser2

# =========================
# 2. Filter BHVA from Hotmail
# =========================

$hotmailBhva = $hotmailContacts | Where-Object {
    $_.categories -contains $BhvaCategory
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

    $hotmail = $Contacts | Where-Object { $_.parentFolderId -like "*$HotmailUser*" }
    if ($hotmail.Count -gt 0) {
        $base = $hotmail[0]
    } else {
        $base = $Contacts | Sort-Object {
            ($_.displayName, $_.companyName, $_.jobTitle, $_.mobilePhone, $_.businessPhones, $_.homePhones, $_.emailAddresses) |
                Where-Object { $_ } |
                Measure-Object
        } -Descending | Select-Object -First 1
    }

    $merged = $base.PSObject.Copy()

    $allEmails = @()
    foreach ($c in $Contacts) {
        if ($c.emailAddresses) {
            $allEmails += $c.emailAddresses | ForEach-Object { $_.address }
        }
    }
    $allEmails = $allEmails | Where-Object { $_ } | Select-Object -Unique

    $merged.emailAddresses = $allEmails | ForEach-Object {
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphEmailAddress]@{
            address = $_
        }
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
        $merged.mobilePhone   = $normalized[0]
        $merged.businessPhones = $normalized[1..($normalized.Count-1)]
    }

    $allBodies = $Contacts | Where-Object { $_.body -and $_.body.content } |
        ForEach-Object { $_.body.content }

    if ($allBodies.Count -gt 0) {
        $merged.body = @{
            contentType = "Text"
            content     = ($allBodies -join "`n--- BHV notes ---`n")
        }
    }

    $cats = @()
    foreach ($c in $Contacts) {
        if ($c.categories) { $cats += $c.categories }
    }
    $cats += $BhvaCategory
    $merged.categories = $cats | Select-Object -Unique

    return $merged
}

$mergedBhvaContacts = @()

foreach ($entry in $unifiedBhva.GetEnumerator()) {
    $mergedBhvaContacts += Merge-Contacts -Contacts $entry.Value.SourceContacts
}

# =========================
# 5. Apply to Hotmail
# =========================

$hotmailIndex = @{}
foreach ($c in $hotmailContacts) {
    $emails = @()
    if ($c.emailAddresses) {
        $emails = $c.emailAddresses | ForEach-Object { $_.address }
    }
    $key = Build-MatchKey -DisplayName $c.displayName -Emails $emails
    if (-not $hotmailIndex.ContainsKey($key)) {
        $hotmailIndex[$key] = $c
    }
}

$changes = @()

foreach ($m in $mergedBhvaContacts) {
    $emails = @()
    if ($m.emailAddresses) {
        $emails = $m.emailAddresses | ForEach-Object { $_.address }
    }
    $key = Build-MatchKey -DisplayName $m.displayName -Emails $emails

    if ($hotmailIndex.ContainsKey($key)) {
        $existing = $hotmailIndex[$key]

        $changes += [ordered]@{
            Action   = "UpdateHotmail"
            Key      = $key
            Name     = $m.displayName
            Emails   = ($emails -join ", ")
        }

        if (-not $DryRun) {
            Update-MgUserContact -UserId $HotmailUser -ContactId $existing.Id -BodyParameter $m
        }
    } else {
        $changes += [ordered]@{
            Action   = "CreateHotmail"
            Key      = $key
            Name     = $m.displayName
            Emails   = ($emails -join ", ")
        }

        if (-not $DryRun) {
            New-MgUserContact -UserId $HotmailUser -BodyParameter $m
        }
    }
}

# =========================
# 6. Apply to BHV mailboxes
# =========================

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

        if ($index.ContainsKey($key)) {
            $existing = $index[$key]

            $changes += [ordered]@{
                Action   = "UpdateBhv"
                Mailbox  = $UserId
                Key      = $key
                Name     = $m.displayName
                Emails   = ($emails -join ", ")
            }

            if (-not $DryRun) {
                Update-MgUserContact -UserId $UserId -ContactId $existing.Id -BodyParameter $m
            }
        } else {
            $changes += [ordered]@{
                Action   = "CreateBhv"
                Mailbox  = $UserId
                Key      = $key
                Name     = $m.displayName
                Emails   = ($emails -join ", ")
            }

            if (-not $DryRun) {
                New-MgUserContact -UserId $UserId -BodyParameter $m
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
            $changes += [ordered]@{
                Action   = "DeleteBhv"
                Mailbox  = $UserId
                Key      = $key
                Name     = $c.displayName
            }

            if (-not $DryRun) {
                Remove-MgUserContact -UserId $UserId -ContactId $c.Id -Confirm:$false
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

$changes | Export-Csv -Path $logPath -NoTypeInformation

Write-Host "Sync complete. Changes logged to $logPath"
if ($DryRun) {
    Write-Host "Dry run mode: no changes were written."

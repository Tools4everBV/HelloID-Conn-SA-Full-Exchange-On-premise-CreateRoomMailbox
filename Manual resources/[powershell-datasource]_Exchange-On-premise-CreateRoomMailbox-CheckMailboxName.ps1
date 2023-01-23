###############################################################################
# HelloID-Conn-SA-Full-Exchange-On-premise-CreateRoomMailbox
# [powershell-datasource]_Exchange-On-premise-Check-names-CreateRoomMailbox
###############################################################################

$RoomMailboxName = $datasource.roomname
$emailAddress = $datasource.emailaddress

# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck #-SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -Authentication Default -AllowRedirection -SessionOption $sessionOption
    $null = Import-PSSession -Session $exchangeSession -AllowClobber -CommandName Get-Mailbox, Get-AcceptedDomain
    Write-Information "Successfully connected to Exchange using the URI '[$ExchangeConnectionUri]'"
} catch {
    Write-Error "Error connecting to Exchange using the URI '[$exchangeConnectionUri]', Message '$($_.Exception.Message)'"
}

try {
    $i = 1
    $newRoomMailboxName = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($RoomMailboxName))
    $eMail = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($eMailAddress))

    $domainName = (Get-AcceptedDomain)[0].DomainName

    # create lookup table
    $allMailboxes = Get-Mailbox | Select-Object PrimarySmtpAddress, Name
    $allMailboxesGroupedByPrimary = $allMailboxes | Group-Object -Property PrimarySmtpAddress -AsHashTable
    $allMailboxesGroupedByName = $allMailboxes | Group-Object -Property Name -AsHashTable

    # Check if a roommailbox with name '$newRoomMailboxName' already exists
    if (-not($allMailboxesGroupedByName[$newRoomMailboxName])) {
        $roomName = $newRoomMailboxName
    }else{
        $i=1
        Do {
            $roomName = $newRoomMailboxName + $i
        } While (!$allMailboxesGroupedByName[$newRoomMailboxName])
    }

    if (-not($allMailboxesGroupedByPrimary[$eMail])) {
        $newEmail = $eMail
    }else{
        $i=1
        Do {
            $newName = $email.Split("@")[0]
            $newEmail = "$newName$i"+"@$domainName"
            $i++
        } While (!$allMailboxesGroupedByPrimary[$eMail])
    }

    $returnObject = @{
        displayname       = $roomName
        userPrincipalName = "$roomName@$domainName"
        EmailAddress = $newEmail
    }
    Write-Output $returnObject
} catch {
    Write-Error "Error generating name, Message '$($_.Exception.Message)'"
}

# Disconnect from Exchange
try{
    Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    Write-Information "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]"        
} catch {
    Write-Error "Error disconnecting from Exchange.  Error: $($_.Exception.Message)"
}


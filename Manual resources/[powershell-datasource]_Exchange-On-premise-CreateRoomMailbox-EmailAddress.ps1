###############################################################################
# HelloID-Conn-SA-Full-Exchange-On-premise-CreateRoomMailbox
# [powershell-datasource]_Exchange-On-premise-Check-names-CreateRoomMailbox
###############################################################################

$emailAddress = $datasource.emailaddress
# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername,$adminSecurePassword
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -Authentication Basic -AllowRedirection -SessionOption $sessionOption
    $null = Import-PSSession -Session $exchangeSession -AllowClobber -CommandName Get-Mailbox, Get-AcceptedDomain
    #Write-Information "Successfully connected to Exchange '$ExchangeConnectionUri'"
} catch {
    Write-Error "Error connecting to Exchange using the URI '$exchangeConnectionUri', Message '$($_.Exception.Message)'"
}

try {
    $i = 1
    $eMail = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($eMailAddress))
    $domainName = (Get-AcceptedDomain)[0].DomainName

    # create lookup table
    $allMailboxes = Get-Mailbox | Select-Object PrimarySmtpAddress
    $allMailboxesGroupedByPrimary = $allMailboxes | Group-Object -Property PrimarySmtpAddress -AsHashTable
    
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
        EmailAddress = $newEmail
    }
    Write-Output $returnObject
} catch {
    Write-Error "Error generating name, Message '$($_.Exception.Message)'"
}

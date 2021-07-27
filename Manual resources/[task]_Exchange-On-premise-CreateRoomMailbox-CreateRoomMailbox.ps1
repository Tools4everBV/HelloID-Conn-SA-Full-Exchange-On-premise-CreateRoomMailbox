##############################################################
# HelloID-Conn-SA-Full-Exchange-On-Premises-RoomMailboxCreate
##############################################################

# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername,$adminSecurePassword
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -Authentication Basic -AllowRedirection -SessionOption $sessionOption
    Import-PSSession $exchangeSession
    HID-Write-Status -Message "Successfully connected to Exhchange '$ExchangeConnectionUri'" -Event Information
} catch {
    HID-Write-Status -Message "Error connecting to Exchange using the URI [$exchangeConnectionUri]" -Event Error
    HID-Write-Status -Message "Error at line: $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)" -Event Error
    HID-Write-Summary -Message "Failed to connect to Exchange using the URI [$exchangeConnectionUri]" -Event Failed
}

function New-Password {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [int]
        $length,

        [int]
        $SpecialCharCount = 4
    )

    Add-Type -AssemblyName System.Web
    $password = [System.Web.Security.Membership]::GeneratePassword($length, $SpecialCharCount)

    Write-Output $password
}

# Create mailbox
try {
    $password = New-Password -Length 10 -SpecialCharCount 2
    $roomMailboxParams = @{
        Name              = $RoomName
        DisplayName       = $RoomName
        ResourceCapacity  = $ResourceCapacity
        Password          = (ConvertTo-SecureString -AsPlainText $password -Force)
        PrimarySmtpAddress= $EmailAddress
    }
    $null = New-Mailbox @roomMailboxParams -Room
    HID-Write-Status -Message "Successfully created room mailbox $RoomName" -Event Success
    HID-Write-Summary -Message "Successfully created room mailbox $RoomName" -Event Success
} catch {
    $ex = $_
    HID-Write-Status -Message "Error creating room mailbox '$RoomName', Message '$ex.Exception.Message'" -Event Error
    HID-Write-Summary -Message "Error creating room mailbox '$RoomName', Message '$ex.Exception.Message'" -Event Error
}

# Disconnect from Exchange
Remove-PsSession -Session $exchangeSession
HID-Write-Status -Message "Successfully disconnected from Exchange" -Event Success

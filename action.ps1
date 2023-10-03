# HelloID-Task-SA-Target-ExchangeOnPremises-SharedMailboxRevokeFullAccess
#########################################################################
# Form mapping
$formObject = @{
    DisplayName     = $form.DisplayName
    MailboxIdentity = $form.MailboxIdentity
    UsersToRemove   = [array]$form.Users
}

[bool]$IsConnected = $false
try {
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos  -ErrorAction Stop
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber -CommandName 'Remove-MailboxPermission'
    $IsConnected = $true

    # <Action logic here>
    foreach ($user in $formObject.UsersToRemove) {
        Write-Information "Executing ExchangeOnPremises action: [SharedMailboxRevokeFullAccess] for: [$($user.UserPrincipalName)]"

        $ParamsRemoveMailboxPermission = @{
            Identity        = $formObject.MailboxIdentity
            User            = $user.UserPrincipalName
            AccessRights    = 'FullAccess'
            InheritanceType = 'All'
        }
        $null = Remove-MailboxPermission @ParamsRemoveMailboxPermission -Confirm:$false

        $auditLog = @{
            Action            = 'RevokeMembership'
            System            = 'ExchangeOnPremises'
            TargetIdentifier  = $formObject.MailboxIdentity
            TargetDisplayName = $formObject.MailboxIdentity
            Message           = "ExchangeOnPremises action: [SharedMailboxRevokeFullAccess][$($user.UserPrincipalName)] from [$($formObject.DisplayName)] executed successfully"
            IsError           = $false
        }
        Write-Information -Tags 'Audit' -MessageData $auditLog
        Write-Information "ExchangeOnPremises action: [SharedMailboxRevokeFullAccess][$($user.UserPrincipalName)] from [$($formObject.DisplayName)] executed successfully"
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'RevokeMembership'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.MailboxIdentity
        Message           = "Could not execute ExchangeOnPremises action: [SharedMailboxRevokeFullAccess][$($user.UserPrincipalName)] from [$($formObject.DisplayName)] , error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnPremises action: [SharedMailboxRevokeFullAccess][$($user.UserPrincipalName)] from [$($formObject.DisplayName)] , error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        Remove-PSSession -Session $exchangeSession -Confirm:$false  -ErrorAction Stop
    }
}
#########################################################################

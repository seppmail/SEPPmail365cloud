
Write-verbose "Export Exo-Config as JSON"
New-SC365ExOReport -jsonBackup

Write-verbose "Rename Connectors from SEPPmail to BKP-SC"
$oldIbc = Get-InboundConnector -Identity "[SEPPmail.cloud]*"
Set-InboundConnector -Identity $oldIbc.Identity -Name ($oldIbc.Identity -replace 'SEPPmail.Cloud','SC BKP')

$oldObc = Get-OutBoundConnector -Identity "[SEPPmail.cloud]*"
Set-OutBoundConnector -Identity $oldObc.Identity -Name ($oldObc.Identity -replace 'SEPPmail.Cloud','SC BKP')
#FIXME Check if connectors with Backup names exist already

Write-Verbose "Rename existing SEPPmail rules"
$oldTrpRls = Get-TransportRule -Identity '*SEPPmail.cloud*'
foreach ($rule in $oldTrpRls) {
    Set-TransportRule -Identity $rule.Name -Name ($rule.Name -replace 'SEPPmail.Cloud','SC BKP') -WhatIf
}

Write-Verbose "Creating new connectors disabled" 
New-SC365connectors -disabled

Write-Verbose "Creating new Transport Rules" 
New-SC365Rules -Disabled

Read-Host "I have rearranged the transport rules for a working setup"
do {
    # Frage den Benutzer nach einer Eingabe
    $response = Read-Host "I have rearranged the transport rules for a working setup and i know the risks of this step (Y/N)"

    if ($response -match '^[Yy]$') {
        Write-Host "Sie haben 'Y' gewählt. Fortfahren..."
       
        # Do the critical stuff - enable new stuff
        $newIbc = Get-InboundConnector -Identity "[SEPPmail.cloud]*"
        Set-InboundConnector -Identity $newIbc.Identity -Enabled:$true
        $newObc = Get-OutboundConnector -Identity "[SEPPmail.cloud]*"
        Set-OutboundConnector -Identity $newObc.Identity -Enabled:$true

        $newTrpRls = Get-Transportrule -Identity '*SEPPmail.cloud*'
            foreach ($rule in $newTrpRls) {
                Enable-Transportrule -Identity $rule.Name 
        }

         # Do the less critical stuff - disable old stuff
         $bkpIbc = Get-InboundConnector -Identity "*SC BKP*"
         Set-InboundConnector -Identity $bkpIbc.Identity -Enabled:$true
         $bkpObc = Get-OutboundConnector -Identity "*SC BKP*"
         Set-OutboundConnector -Identity $bkpObc.Identity -Enabled:$true

         $bkpTrpRls = Get-TransportRule -Identity '*SC BKP*'
             foreach ($rule in $bkpTrpRls) {
                 Enable-TransportRule -Identity $rule.Name 
        }
        
        break
    } elseif ($response -match '^[Nn]$') {
        Write-Host "Script has been stopped, go to the Exchange Online Admin page and check your configuration"
        $proceed = $false
        break
    } else {
        Write-Host "Invalid character, please choose 'Y' or 'N'." -ForegroundColor Red
    }
} while ($true)


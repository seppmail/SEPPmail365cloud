
Write-verbose "Export Exo-Config as JSON"
New-SC365ExOReport -jsonBackup

Write-verbose "Rename Connectors from SEPPmail to BKP-SC"
$oldibc = Get-InboundConnector -Identity "[SEPPmail.cloud]*"
Set-InboundConnector -Identity $oldibc.Identity -Name ($oldibc.Name -replace 'SEPPmail.Cloud','SC BKP')

$oldobc = Get-OutBoundConnector -Identity "[SEPPmail.cloud]*"
Set-OutBoundConnector -Identity $oldobc.Identity -Name ($oldobc.Name -replace 'SEPPmail.Cloud','SC BKP')
#FIXME Check if connectors with Backup names exist already

Write-Verbose "Rename existing SEPPmail rules"
$oldtrprls = Get-Transportrule -Identity '*SEPPmail.cloud*'
foreach ($rule in $oldtrprls) {
    Set-TransportRule -Identity $rule.Name -Name ($rule.Name -replace 'SEPPmail.Cloud','SC BKP') -WhatIf
}

Write-Verbose "Creating new connectors disabled" 
New-SC365connectors -enabled -disabled

Write-Verbose "Creating new Transport Rules" 
New-SC365Rules -Disabled

Read-Host "I have rearranged the transport rules for a working setup"
do {
    # Frage den Benutzer nach einer Eingabe
    $response = Read-Host "I have rearranged the transport rules for a working setup and i know the risks of this step (Y/N)"

    if ($response -match '^[Yy]$') {
        Write-Host "Sie haben 'Y' gewählt. Fortfahren..."
       
        # Do the critical stuff - enable new stuff
        $newibc = Get-InboundConnector -Identity "[SEPPmail.cloud]*"
        Set-InboundConnector -Identity $newibc.Identity -Enabled:$true
        $newobc = Get-OutboundConnector -Identity "[SEPPmail.cloud]*"
        Set-OutboundConnector -Identity $newobc.Identity -Enabled:$true

        $newtrprls = Get-Transportrule -Identity '*SEPPmail.cloud*'
            foreach ($rule in $newtrprls) {
                Enable-Transportrule -Identity $rule.Name 
        }

         # Do the less critical stuff - disable old stuff
         $bkpibc = Get-InboundConnector -Identity "*SC BKP*"
         Set-InboundConnector -Identity $bpkibc.Identity -Enabled:$true
         $bkpobc = Get-OutboundConnector -Identity "*SC BKP*"
         Set-OutboundConnector -Identity $bpkobc.Identity -Enabled:$true

         $bkptrprls = Get-Transportrule -Identity '*SC BKP*'
             foreach ($rule in $bpktrprls) {
                 Enable-Transportrule -Identity $rule.Name 
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


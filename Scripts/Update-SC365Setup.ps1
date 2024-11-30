[CmdletBinding()]
param()

Write-verbose "Export Exo-Config as JSON"
#TODO: New-SC365ExOReport -jsonBackup

#

if ((Get-InboundConnector -Identity '[SC BK]*' -ea SilentlyContinue) -or (Get-OutboundConnector -Identity '[SC BK]*' -ea SilentlyContinue) -or (Get-TransportRule -Identity '[SC BK]*' -ea SilentlyContinue)) {

    Write-Verbose "Rename existing SEPPmail.cloud rules and disable those routing to the Outbound Connector."
    $oldTrpRls = Get-TransportRule -Identity '[SEPPmail.cloud]*'
    foreach ($rule in $oldTrpRls) {
        if (($rule.state -eq 'Enabled') -and ($rule.RouteMessageOutboundConnector)) {
            Disable-TransportRule -Identity $rule.identity -Confirm:$false
        }
        Set-TransportRule -Identity $rule.Name -Name ($rule.Name -replace 'SEPPmail.Cloud','SC BKP')
    }
    
    Write-Verbose "Rename existing SEPPmail.cloud Inbound Connector"
    try {
        Write-verbose 'Rename Connectors from SEPPmail to BKP-SC'
        $oldIbc = Get-InboundConnector -Identity '[SEPPmail.cloud]*'
        foreach ($ic in $oldIbc) {
            Set-InboundConnector -Identity $Ic.Identity -Name ($Ic.Identity -replace 'SEPPmail.Cloud','SC BKP')
            }
        }
    catch {
        Write-Warning "No inbound Connector found with [SEPPmail.cloud] in the name"
        }
    
    Write-Verbose "Rename existing SEPPmail.cloud Outbound Connector"
    $oldObc = Get-OutBoundConnector -Identity "[SEPPmail.cloud]*"
    foreach ($oc in $oldObc) {
        Set-OutBoundConnector -Identity $oc.Identity -Name ($Oc.Identity -replace 'SEPPmail.Cloud','SC BKP')
    }
    
    Write-verbose "Enable transport rules so mailflow works during upgrade"
    $bkpRls = Get-TransportRule -Identity '*SC BKP*'
    foreach ($rule in $bkpRls) {
        if ($rule.status -ne 'Enabled') {
            Enable-TransportRule -Identity $rule.identity -Confirm:$false
        }
    }
    Write-Verbose "Old Setup running with SC BKP as name"
    
    Write-Verbose "Creating new connectors disabled" 
    New-SC365connectors -Disabled:$true
    
    Write-Verbose "Enabling new outbound connector"
    $newObc = Get-OutboundConnector -Identity "[SEPPmail.cloud]*"
    Set-OutboundConnector -Identity $newObc.Identity -Enabled:$true
    
    Write-Verbose "Creating new Transport Rules" 
    New-SC365Rules -Disabled:$true
    
    #Read-Host "I have rearranged the transport rules for a working setup"
    do {
        # Frage den Benutzer nach einer Eingabe
        $response = Read-Host "I have rearranged/customized the Transport Rules for a working setup and i KNOW the RISKS of this step (Y/N)"
    
        if ($response -match '^[Yy]$') {
            Write-Host "Sie haben 'Y' gewählt. Fortfahren..."
           
            # Do the critical stuff - enable Inbound Connector stuff
            $newIbc = Get-InboundConnector -Identity "[SEPPmail.cloud]*"
            foreach ($ic in $newIbc) {
                Set-InboundConnector -Identity $ic.Identity -Enabled:$true
            }
            $newTrpRls = Get-Transportrule -Identity '*SEPPmail.cloud*'
                foreach ($rule in $newTrpRls) {
                    Enable-Transportrule -Identity $rule.Name 
            }
    
             # Do the less critical stuff - disable old stuff
            $bkpIbc = Get-InboundConnector -Identity "[SC BKP]*"
            Write-Verbose "Disabling old Inbound Connectors"
            foreach ($ic in $bkpIbc) {
                Set-InboundConnector -Identity $bkpIbc.Identity -Enabled:$false
            }
            $bkpObc = Get-OutboundConnector -Identity "[SC BKP]*"
            Write-Verbose "Disabling old Outbound Connectors"
            foreach ($oc in $bkpObc) {
                Set-OutboundConnector -Identity $bkpObc.Identity -Enabled:$false
            }
            
            $bkpTrpRls = Get-TransportRule -Identity '[SC BKP]*'
            Write-Verbose "Disabling old Transport Rules"
            foreach ($rule in $bkpTrpRls) {
                     Disable-TransportRule -Identity $rule.Name -confirm:$false
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
    
} else {
    Write-Error 'STOPPING - Found Existing Backup Objects - clean up the environment from SC BKP objects (rules and conenctors) and TRY again'
    break
}


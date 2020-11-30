<#
  .SYNOPSIS
  This script adds new objects to the SCOM Run As Account Distribution List.

  .DESCRIPTION
  The script gets all credential alerts from SCOM and adds the object to the right Run As Accounts distribution list.
  Additional accounts must be added to the correct section outlined by "#####".

  .PARAMETERS
  -AlertId : Optionally it is possible to run the script from a SCOM interal Notification Command Channel and provide the Alert ID.
  -Debug : if $true the script will write more information to the PowerShell console and the Operations Manager Event Log.

  .EXAMPLE
  Just run the script and it will get all related alerts:
  Update-ScomRunAsAccountDistribution.ps1

  .EXAMPLE
  Just run the script and it will get all related alerts:
  .\Update-ScomRunAsAccountDistribution.ps1

  .EXAMPLE
  Configure the script in a Notification Command Channel:
  Command:                 C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe
  Command Line Parameter:  -Command "& '"C:\SCOM\Scripts\Update-ScomRunAsAccountDistribution.ps1"'" -alertId '$Data/Context/DataItem/AlertId$'
  Startup Folder:          C:\Windows\system32\WindowsPowerShell\v1.0
  
  .CONFIGURATION
  Here is an example of how additional accounts must be defined in the section outlined by "#####".

    if ($alert.Parameters -like "*<replace with your Run As Profile ID>*") {
        $RunAsAccount = "<replace with your Run As Account Name"
    }

  .CREDITS
  Credits to Matthew Long for his initial blog on that: https://matthewlong.wordpress.com/2013/01/25/scom-2012-update-run-as-account-distribution-via-powershell

  .AUTHOR
  Patrick Seidl
  s2 - seidl solutions
  SystemCenterRocks.com
#>

param(
    $alertId,
    [bool]$debug = $true
)

Import-Module OperationsManager

if ($Debug -eq $true) {
    $objApi = new-object -comObject "MOM.ScriptAPI"
    $objAPI.LogScriptEvent("Update-ScomRunAsAccountDistribution", 130, 0, "Script triggered to update Run As Account Distribution List
    $alertId
    ")
}

# get credential alerts from SCOM
if ($alertId) {
    $alerts = Get-SCOMAlert -Id $alertId
} else {
    $alerts = Get-SCOMAlert -ResolutionState 0 -Name "System Center Management Health Service Credentials Not Found Alert Message"
}

if ($alerts) {
    foreach ($alert in $alerts) {
        "-"*70
        $alert.PrincipalName

######### RUN AS ACCOUNT DEFINITION MUST BE ADDED BELOW THIS LINE ######### 
        # check if it is related to the right product and set the correct account name
        # the sample below expects the account for SQL to be named like "SQL Server Run As Account (MyDomainName)"; so there are such accounts for every domain in the environment
        if ($alert.Parameters -like "*Microsoft.SQLServer.*") {
            $RunAsAccount = "SQL Server Run As Account ("+$alert.NetbiosDomainName+")"
        }
        if ($alert.Parameters -like "*Microsoft.SystemCenter.ServiceManager.DatabaseWriteActionAccount*") {
            $RunAsAccount = "Service Manager Database Account"
        }
        
######### RUN AS ACCOUNT DEFINITION MUST BE ADDED ABOVE THIS LINE ######### 

        # move on only if there is an account defined (see "if" statements above)
        if ($RunAsAccount.Length -gt 0) {
            # get the run as account from the MG
            $runas = Get-SCOMRunAsAccount -Name $RunAsAccount

            if ($debug) {
                Write-Host -ForegroundColor Green "Objects in the distribution list before the update:"
                ((Get-SCOMRunAsDistribution $runas).securedistribution).DisplayName
                $objAPI.LogScriptEvent("Update-ScomRunAsAccountDistribution", 131, 0, "Objects to be added to the distribution list:
                "+$alert.PrincipalName+" --> "+$RunAsAccount
                )
                $objAPI.LogScriptEvent("Update-ScomRunAsAccountDistribution", 131, 0, "Objects in the distribution list before the update:
                "+((Get-SCOMRunAsDistribution $runas).securedistribution).DisplayName )
            }
            # how many objects are in the distribution list
            $firstCount = ((Get-SCOMRunAsDistribution $runas).securedistribution).count
            Write-Host "Amount of objects in the distribution list before the update:
            $RunAsAccount : $firstCount"

            # generate the new distribution list
            if ($firstCount -gt 0) {$monitoringObjects = (Get-SCOMRunAsDistribution $runas).securedistribution}
            $monitoringObjects += $alert | % {get-ScomMonitoringObject -id $_.MonitoringObjectId}

            # write the new distribution list
            $managementGroup = Get-ScomManagementGroup
            [Microsoft.SystemCenter.OperationsManagerV10.Commands.OMV10Utility]::ApproveRunasAccountForDistribution($managementGroup, $runas, $monitoringObjects)

            if ($debug) {
                Write-Host -ForegroundColor Green "Objects in the distribution list after the update:"
                ((Get-SCOMRunAsDistribution $runas).securedistribution).DisplayName
                $objAPI.LogScriptEvent("Update-ScomRunAsAccountDistribution", 132, 0, "Objects in the distribution list after the update:
                "+((Get-SCOMRunAsDistribution $runas).securedistribution).DisplayName )
            }

            # how many objects are in the distribution list
            $lastCount = ((Get-SCOMRunAsDistribution $runas).securedistribution).count
            Write-Host "Amount of objects in the distribution list after the update:
            $RunAsAccount : $lastCount" 

            # clean up for next run
            $RunAsAccount = $null

            if ($firstCount -lt $lastCount) {
                # close the alert
                $alert.ResolutionState = 255
                $alert.Update(“Alert closed by script after adding the object to the Run As Accounts distribution list.”)
                if ($debug) {
                    $objAPI.LogScriptEvent("Update-ScomRunAsAccountDistribution", 133, 0, "Succesfully updated Run As Account distribution list.")
                }
            } else {
                Write-Host -ForegroundColor Red "Something went wrong: No changes have been applied."
                if ($debug) {
                    $objAPI.LogScriptEvent("Update-ScomRunAsAccountDistribution", 134, 1, "Something went wrong: No changes have been applied.")
                }
                $alert.Update(“Alert has not been closed by script since something went wrong when trying to add the object to the Run As Accounts distribution list.”)
            }
        }  else {
            Write-Host -ForegroundColor Yellow "No Run As Account matches to alert."
            if ($debug) {
                $objAPI.LogScriptEvent("Update-ScomRunAsAccountDistribution", 135, 2, "No Run As Account matches to alert.")
            }
        }
    }
} else {
    Write-Host -ForegroundColor Yellow "No alerts have been found."
    if ($debug) {
        $objAPI.LogScriptEvent("Update-ScomRunAsAccountDistribution", 136, 2, "No alerts have been found.")
    }
}
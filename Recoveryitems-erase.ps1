<#
Allows you to check the occupied space of the mailbox. Create retention policies. And delete items stored in the recoveryitems folder.
Version 1.0 October 2022
malamos@outlook.com
#>
Write-Host "*****  Script Delete Recovery Items  ******" -ForegroundColor Green

Write-Host "Don't Forget Install Module Exchange" -ForegroundColor red
Write-Host "Install-Module -Name ExchangeOnlineManagement" -ForegroundColor yellow


$buzon= Read-Host -Prompt "upn"

Write-Host "Size Recovery Items" -ForegroundColor red
Get-MailboxFolderStatistics -Identity $buzon -FolderScope RecoverableItems | ft Identity, ItemsInFolder, FolderAndSubfolderSize

#disblehold
Write-Host "Enable Archive Mailbox" -ForegroundColor yellow
Enable-Mailbox $buzon -Archive 
Enable-Mailbox $buzon -AutoExpandingArchive 
Set-Mailbox $buzon -RemoveDelayHoldApplied 
Set-Mailbox $buzon -RetainDeletedItemsFor 0
Set-Mailbox $buzon -LitigationHoldEnabled $false
Set-Mailbox $buzon -SingleItemRecoveryEnabled $false
Set-CASMailbox $buzon -EwsEnabled $false -ActiveSyncEnabled $false -MAPIEnabled $true -OWAEnabled $true -ImapEnabled $false -PopEnabled $false
Set-Mailbox $buzon -RemoveDelayHoldApplied
Set-Mailbox $buzon -RemoveDelayReleaseHoldApplied

#AddPolicy
Write-Host "Create Policy" -ForegroundColor yellow
New-RetentionPolicyTag -Name "MoverDeleteditems" -Type RecoverableItems -AgeLimitForRetention 1 -RetentionAction MoveToArchive 
New-RetentionPolicy "Borrar Elementos Eliminados" -RetentionPolicyTagLinks "MoverDeletedItems"
Set-Mailbox -identity $buzon -RetentionPolicy "Borrar Elementos Eliminados"
set-mailbox -Identity  $buzon -ElcProcessingDisabled $false 
Set-Mailbox $buzon -RemoveDelayHoldApplied

#Start
Write-Host "Start maintenance" -ForegroundColor yellow
Start-ManagedFolderAssistant -Identity $buzon
Search-Mailbox $buzon -SearchQuery size>0 -SearchDumpsterOnly -DeleteContent

Write-Host "Check size RecoverableItems " -ForegroundColor Red
Write-Host "Get-MailboxFolderStatistics -Identity $buzon -FolderScope RecoverableItems | ft Identity, ItemsInFolder, FolderAndSubfolderSize" -ForegroundColor Yellow
# Import BT module
Import-Module 'C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll'

# Set environment authentication
$username = "ericks@rotadooeste.com.br"
$pwd = ConvertTo-SecureString -String "Cro@2021" -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential $Username,$Pwd

# Get both bt and mw tickets
$mwTicket = Get-MW_Ticket -Credentials $cred
$btTicket = Get-BT_Ticket -Credentials $cred -ServiceType BitTitan
$scopedBtTicket = Get-BT_Ticket -Ticket $btTicket -OrganizationId $customer.OrganizationId

# Get customer, filtered by company name
# Mailbox connector is scoped under customer
$customer = Get-BT_Customer -Ticket $btTicket -CompanyName 'default'


# Import csv file containing sitename (used as project name), url (Source and Destination), EdnPointName (Source and Destination) and Library to imported (Source and Destination)
$Sites = Import-Csv -Path "C:\Projetos\O365 & Infra\Rota do Oeste\Bit Titan\Add-MW_MailboxConnector_SharePoint.csv"

# Set up export and import configurations
# You can choose to provide admin credentials and set UseAdministrativeCredentials to true, then you do not need to provide usernames and passwords in mailbox creation.

If(Test-Path $Sites){    
     foreach($Site in $Sites){
        $exportConfigurationMgmtSvc = New-Object -TypeName ManagementProxy.ManagementService.SharePointConfiguration -Property @{
                "Url" = $site.UrlSrc;
                "AdministrativeUsername" = "ericks@odbsa.onmicrosoft.com";
                "AdministrativePassword" = "K@rate2020dell2dan";
                "UseAdministrativeCredentials" = $true;
        }
        
        $importConfigurationMgmtSvc = New-Object -TypeName ManagementProxy.ManagementService.SharePointBetaConfiguration -Property @{
                "UseSharePointOnlineProvidedStorage" = $true;
                "Url" = $site.UrlDst;
                "AdministrativeUsername" = "bittitan@rotadooestebr.onmicrosoft.com";
                "AdministrativePassword" = "20Partner03@";
                "UseAdministrativeCredentials" = $true;
                "AzureAccountKey" = $null;
                "AzureStorageAccountName" = $null;
                  
        }
        
        $exportConfigurationWebApi = New-Object -TypeName MigrationProxy.WebApi.SharePointConfiguration -Property @{
                "Url" = $site.UrlSrc;
                "AdministrativeUsername" = "ericks@odbsa.onmicrosoft.com";
                "AdministrativePassword" = "K@rate2020dell2dan";
                "UseAdministrativeCredentials" = $true;
        }
        
        $importConfigurationWebApi = New-Object -TypeName MigrationProxy.WebApi.SharePointBetaConfiguration -Property @{
                "UseSharePointOnlineProvidedStorage" = $true;
                "Url" = $site.UrlDst;
                "AdministrativeUsername" = "bittitan@rotadooestebr.onmicrosoft.com";
                "AdministrativePassword" = "20Partner03@";
                "UseAdministrativeCredentials" = $true;
                "AzureAccountKey" = $null;
                "AzureStorageAccountName" = $null;      
        }
    
        # Create a new EndPoint
        $endpointSRC = Add-BT_Endpoint -Ticket $scopedBtTicket -Configuration $exportConfigurationMgmtSvc -Type SharePoint -Name $site.EndPointNameExport
        $endpointDST = Add-BT_Endpoint -Ticket $scopedBtTicket -Configuration $importConfigurationMgmtSvc -Type SharePointBeta -Name $site.EndPointNameImport
        
        # Create a new project
        $connector = Add-MW_MailboxConnector -ticket $mwTicket -ProjectType Storage `
            -ImportType SharePointBeta -ExportType SharePoint -Name $site.Name -UserId $mwTicket.UserId `
            -ImportConfiguration $importConfigurationWebApi -ExportConfiguration $exportConfigurationWebApi `
            -ZoneRequirement SouthAmerica -AdvancedOptions "UseApplicationPermission=1" `
            -SelectedExportEndpointId $endpointSRC.Id -SelectedImportEndpointId $endpointDST.Id -OrganizationId $customer.OrganizationId

        # Create an migration item with export, import credentials
        $mailbox = Add-MW_Mailbox -Ticket $mwTicket -ConnectorId $connector.Id `
	        -ExportLibrary $site.ExportLibrary -ImportLibrary $site.ImportLibrary

        Write-Host -NoNewline -ForegroundColor Green "[ OK ] - "
        Write-Host -f Green "Projeto" $site.Name "addicionado com sucesso"

        # Perform a Verify Credential migration
        $migration = Add-MW_MailboxMigration -Ticket $mwTicket -ConnectorId $connector.Id -MailboxId $mailbox.Id -UserId $mwTicket.UserId -Type Verification

        # Start a Full migration
        #$migration = Add-MW_MailboxMigration -Ticket $mwTicket -ConnectorId $connector.Id -MailboxId $mailbox.Id -UserId $mwTicket.UserId -Type Full -ItemTypes All
    }

}
else
{
	Write-Host -ForegroundColor Red "Csv file was not found." 
}
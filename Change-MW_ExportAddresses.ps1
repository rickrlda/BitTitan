#This script needs to run on the MigrationWiz Powershell
#The script will change the source address of your MigrationWiz Project
#The script reads from a CSV file
#03/27/2016

Import-Module 'C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll'

#Retrieve credentials.
$username = "ericks@rotadooeste.com.br" 
$password = "Cro@2021" | ConvertTo-SecureString -AsPlainText -Force 
$Creds = New-Object System.Management.Automation.PSCredential $Username,$Password

#Retrieve ticket.
Try
{
	$ticket = Get-MW_Ticket -Credentials $Creds
}
catch [Exception]
{
	Write-Host -ForegroundColor Red "Failed to connect to MigrationWiz"
	Write-Host -ForegroundColor Red $_.Exception.Message
	Write-Host
	Exit;
}

#Ask for the name of the project.
$Logfile = "C:\Projetos\O365 & Infra\Rota do Oeste\Bit Titan\Change-MW_ExportAddresses"
$ConnectorName = "OneDrive For Business" #"Please enter the name of your project"
$Connector = Get-MW_MailboxConnector -Ticket $ticket -FilterBy_String_Name $ConnectorName

#Check if the project name is unique.
if ($Connector.Count -eq 1)
{
	#Ask for the name of the csv files
	$CSVPath = "C:\Projetos\O365 & Infra\Rota do Oeste\Bit Titan\Change-MW_ExportAddresses.csv" #"Please enter the path of your CSV file (example: C:\scripts\test.csv)"

	#Check if the file exists.
	if (Test-Path $CSVPath)
	{
		#Loop through the csv file.
		Import-Csv -Path $CSVPath | ForEach-Object {
			#Get the mailbox(es).
			$mailboxes = Get-MW_Mailbox -Ticket $ticket -FilterBy_Guid_ConnectorId $Connector.id -FilterBy_String_ExportEmailAddress $_.UserPrincipalName
			if ($mailboxes)
			{
				#Loops through each mailbox.
				foreach ($mailbox in $mailboxes)
				{
					$Result = Set-MW_Mailbox -Ticket $ticket -ConnectorId $Connector.id -mailbox $Mailbox -ExportEmailAddress $_.NewUPN -ExportUserName $_.NewUPN
					Write-Host -NoNewline -ForegroundColor Green "[  OK  ] "
					Write-Host "Renamed ""$($_.UserPrincipalName)"" to ""$($_.NewUPN)""."
				}
			}
			else
			{
				Write-Host -NoNewline -ForegroundColor DarkYellow "[ FAIL ] "
				Write-Host  "Could not find user ""$($_.UserPrincipalName)"""
                Write-Output "$($_.UserPrincipalName)" | Out-file $Logfile\UserNotFound.csv -Append | Out-Null
			}
		}
	}
	else
	{
		Write-Host -ForegroundColor Red "The csv file ""$CSVPath"" was not found." 
	}
}
else
{
	Write-Host "Failed to find a unique project named ""$ConnectorName"" in your account" -ForegroundColor Red
}
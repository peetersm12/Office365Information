#Format and display errors
Function Update-MessageBlock{
	param(
		$type,
		$message
	)
	$Run = New-Object System.Windows.Documents.run
	
	if($type -eq "Critical"){
		$Run.Foreground = "Red"
		$Run.Text = $message
		$uiHash.MessageBlock.Inlines.Add($Run)
		$uiHash.MessageBlock.Inlines.Add((New-Object System.Windows.Documents.LineBreak)) 
	}
	elseif($type -eq "Warning"){
		$Run.Foreground = "Orange"
		$Run.Text = $message
		$uiHash.MessageBlock.Inlines.Add($Run)
		$uiHash.MessageBlock.Inlines.Add((New-Object System.Windows.Documents.LineBreak)) 
	}
	elseif($type -eq "Status"){
		$Run.Foreground = "Black"
		$Run.Text = $message
		$uiHash.MessageBlock.Inlines.Add($Run)
		$uiHash.MessageBlock.Inlines.Add((New-Object System.Windows.Documents.LineBreak)) 
	}
	elseif($type -eq "Okay"){
		$Run.Foreground = "Green"
		$Run.Text = $message
		$uiHash.MessageBlock.Inlines.Add($Run)
		$uiHash.MessageBlock.Inlines.Add((New-Object System.Windows.Documents.LineBreak)) 
	}
}

function Connect-ToOffice365(){
	param(
		$credential
	)

	if($credential){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Connecting to Office 365...Please Wait"
		})
			
		try{
			#connect to Office 365		
			Connect-MsolService -credential $credential	

			$tenant = ((Get-MsolAccountSku).AccountSkuId).split(":")[0]

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				$uiHash.Office365Tenant.Text = "$($tenant).onmicrosoft.com"
				update-MessageBlock -type "Okay" -Message "Tenant name is $($tenant).onmicrosoft.com...Please Wait"
			})    
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Okay" -Message "Connected to Office365...Please Wait"
				$uiHash.Users_CheckBox.isEnabled = $true
				$uiHash.Groups_CheckBox.isEnabled = $true
				$uiHash.Guests_CheckBox.isEnabled = $true
				$uiHash.Contacts_CheckBox.isEnabled = $true
				$uiHash.DeletedUsers_CheckBox.isEnabled = $true
				$uiHash.Domains_CheckBox.isEnabled = $true
				$uiHash.Subscriptions_CheckBox.isEnabled = $true
				$uiHash.Roles_CheckBox.isEnabled = $true
				$uiHash.AADAll_CheckBox.isEnabled = $true
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error connecting to Office365 $($_.exception.Message)...Please Wait"
				$uiHash.Users_CheckBox.isEnabled = $false
				$uiHash.Groups_CheckBox.isEnabled = $false
				$uiHash.Guests_CheckBox.isEnabled = $false
				$uiHash.Contacts_CheckBox.isEnabled = $false
				$uiHash.DeletedUsers_CheckBox.isEnabled = $false
				$uiHash.Domains_CheckBox.isEnabled = $false
				$uiHash.Subscriptions_CheckBox.isEnabled = $false
				$uiHash.Roles_CheckBox.isEnabled = $false
				$uiHash.AADAll_CheckBox.isEnabled = $false
			})
		}
	
		
		try{
			#connect to Exchange Online
			$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
			
			Import-PSSession $Session
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Okay" -Message "Connected to Exchange Online...Please Wait"
				$uiHash.ExchangeMailboxes_CheckBox.isEnabled = $true
				$uiHash.ExchangeGroups_CheckBox.isEnabled = $true
				$uiHash.ExchangeDevices_CheckBox.isEnabled = $true
				$uiHash.ExchangeContacts_CheckBox.isEnabled = $true
				$uiHash.ExchangeArchives_CheckBox.isEnabled = $true
				$uiHash.ExchangePublicFolders_CheckBox.isEnabled = $true
				$uiHash.ExchangeRetentionPolicies_CheckBox.isEnabled = $true
				$uiHash.ExchangeAll_CheckBox.isEnabled = $true
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error connecting to Exchange Online $($_.exception.Message)...Please Wait"
				$uiHash.ExchangeMailboxes_CheckBox.isEnabled = $false
				$uiHash.ExchangeGroups_CheckBox.isEnabled = $false
				$uiHash.ExchangeDevices_CheckBox.isEnabled = $false
				$uiHash.ExchangeContacts_CheckBox.isEnabled = $false
				$uiHash.ExchangeArchives_CheckBox.isEnabled = $false
				$uiHash.ExchangePublicFolders_CheckBox.isEnabled = $false
				$uiHash.ExchangeRetentionPolicies_CheckBox.isEnabled = $false
				$uiHash.ExchangeAll_CheckBox.isEnabled = $false
			})
		}

		try{
			#first get SharePoint URL by tenant domain
			$FQDN = Get-MsolDomain | where-object{$_.name -like "*.onmicrosoft.com"}
			$name = ($fqdn.name).split(".")[0]
			
			#import SharePoint Online Module
			Import-Module Microsoft.Online.SharePoint.PowerShell -WarningAction:SilentlyContinue

			#connect to SharePoint Online
			Connect-SPOService -url "https://$($Name)-admin.sharepoint.com" -credential $Credential
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Okay" -Message "Connected to SharePoint Online...Please Wait"
				$uiHash.SiteCollections_CheckBox.isEnabled = $true
				$uiHash.Webs_CheckBox.isEnabled = $true
				$uiHash.ContentTypes_CheckBox.isEnabled = $true
				$uiHash.Lists_CheckBox.isEnabled = $true
				$uiHash.Features_CheckBox.isEnabled = $true
				$uiHash.SharePointGroups_CheckBox.isEnabled = $true
				$uiHash.SharePointAll_CheckBox.isEnabled = $true
				$uiHash.Teams_CheckBox.isEnabled = $true
			})		
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error connecting to SharePoint Online $($_.exception.Message)"
				$uiHash.SiteCollections_CheckBox.isEnabled = $false
				$uiHash.Webs_CheckBox.isEnabled = $false
				$uiHash.ContentTypes_CheckBox.isEnabled = $false
				$uiHash.Lists_CheckBox.isEnabled = $false
				$uiHash.Features_CheckBox.isEnabled = $false
				$uiHash.SharePointGroups_CheckBox.isEnabled = $false
				$uiHash.Teams_CheckBox.isEnabled = $false
				$uiHash.SharePointAll_CheckBox.isEnabled = $false
			})
		}	

		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Done connecting to Office 365"
		})
	}
}

Function New-LogFile($FullPath){
	if(![System.IO.File]::Exists($FullPath)){
		New-Item $fullPath -ItemType file
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Okay" -Message "Logfile created at $($fullPath)"
		})
	}
	else{
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Warning" -Message "Logfile not created at $($fullPath) $($_.exception.Message), Please try again"
		})
	}
	
}


function Get-Checks($LogPath, $CheckedArray, $ManualSiteCollection, $credentials){
	$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
		update-MessageBlock -type "Status" -Message "Started retrieving information...Please Wait"
	})

	$temp = "<?xml version='1.0' encoding='utf-8' standalone='yes'?> `r`n" + 
	"<!--Office365 Inventarisation Log--> `r`n" + 
	"<Default>"

	Add-Content -Path $LogPath -Value $temp

	#region AAD
	#Check Azure Active Directory Users
	If ($CheckedArray.contains("Users_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Azure Active Directory Users...Please Wait"
		})

		try{
			$temp = "<AADUsers> `r`n"

			$users = get-msoluser -all -ea stop | where-object{$_.UserPrincipalName -notlike "*#ext#*"} | Select-Object DisplayName, FirstName, LastName, UserPrincipalName, Title, Department, Office, PhoneNumber, MobilePhone, CloudAnchor, IsLicensed, @{Name="License"; Expression = {$_.licenses.accountskuid}} 
			
			ForEach ($user in $users) { 
				If (-NOT [System.String]::IsNullOrEmpty($user)) { 
					$temp += "<AADUser> `r`n" +
						"<DisplayName>$($user.DisplayName)</DisplayName> `r`n" + 
						"<FirstName>$($user.FirstName)</FirstName> `r`n" + 
						"<LastName>$($user.LastName)</LastName> `r`n" + 
						"<UserPrincipalName>$($user.UserPrincipalName)</UserPrincipalName> `r`n" + 
						"<Title>$($user.Title)</Title> `r`n" + 
						"<Department>$($user.Department)</Department> `r`n" + 
						"<Office>$($user.Office)</Office> `r`n" + 
						"<PhoneNumber>$($user.PhoneNumber)</PhoneNumber> `r`n" + 
						"<MobilePhone>$($user.MobilePhone)</MobilePhone> `r`n" + 
						"<CloudAnchor>$($user.CloudAnchor)</CloudAnchor> `r`n" + 
						"<IsLicensed>$($user.IsLicensed)</IsLicensed> `r`n" + 
						"<License>$($user.License)</License> `r`n" + 
						"</AADUser> `r`n"
				}
			}

			$temp += "</AADUsers>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Azure Active Directory Users...Please Wait"
				$uiHash.Users_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Azure Active Directory Users $($_.exception.Message)...Please Wait"
				$uiHash.Users_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Azure Active Directory Groups
	If ($CheckedArray.contains("Groups_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Azure Active Directory Groups...Please Wait"
		})

		try{
			$temp = "<AADGroups> `r`n"

			$groups = get-msolGroup | Select-Object DisplayName, EmailAddress, GroupType, ValidationStatus
			
			ForEach ($group in $groups) { 
				If (-NOT [System.String]::IsNullOrEmpty($group)) {  
					$temp += "<AADGroup> `r`n" +
						"<GroupType>$($group.GroupType)</GroupType> `r`n" + 
						"<DisplayName>$($group.DisplayName)</DisplayName> `r`n" + 
						"<EmailAddress>$($group.EmailAddress)</EmailAddress> `r`n" + 
						"<ValidationStatus>$($group.ValidationStatus)</ValidationStatus> `r`n" +
						"</AADGroup> `r`n"				
				}
			}

			$temp += "</AADGroups>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Azure Active Directory Groups...Please Wait"
				$uiHash.Groups_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Azure Active Directory Groups $($_.exception.Message)...Please Wait"
				$uiHash.Groups_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Azure Active Directory Guests
	If ($CheckedArray.contains("Guests_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Azure Active Directory guests...Please Wait"
		})

		try{
			$temp = "<AADGuests> `r`n"

			$Guests = Get-MsolUser -all | Where-Object{$_.UserPrincipalName -like "*#ext#*"} | Select-Object SignInName, UserPrincipalName, DisplayName, WhenCreated

			ForEach ($Guest in $Guests) { 
				If (-NOT [System.String]::IsNullOrEmpty($Guest)) {  
					$temp += "<AADGuest> `r`n" +
						"<SignInName>$($Guest.SignInName)</SignInName> `r`n" + 
						"<UserPrincipalName>$($Guest.UserPrincipalName)</UserPrincipalName> `r`n" + 
						"<DisplayName>$($Guest.DisplayName)</DisplayName> `r`n" + 
						"<WhenCreated>$($Guest.WhenCreated)</WhenCreated> `r`n" +
						"</AADGuest> `r`n"				
				}
			}

			$temp += "</AADGuests>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Azure Active Directory guests...Please Wait"
				$uiHash.Guests_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Azure Active Directory guests $($_.exception.Message)...Please Wait"
				$uiHash.Guests_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Azure Active Directory Contacts
	If ($CheckedArray.contains("Contacts_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Azure Active Directory Contacts...Please Wait"
		})

		try{
			$temp = "<AADContacts> `r`n"

			$Contacts = Get-Msolcontact -all | Select-Object DisplayName, EmailAddress

			ForEach ($Contact in $Contacts) { 
				If (-NOT [System.String]::IsNullOrEmpty($Contact)) {  
					$temp += "<AADContact> `r`n" +
						"<DisplayName>$($Contact.DisplayName)</DisplayName> `r`n" + 
						"<EmailAddress>$($Contact.EmailAddress)</EmailAddress> `r`n" + 
						"</AADContact> `r`n"				
				}
			}

			$temp += "</AADContacts>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Azure Active Directory Contacts...Please Wait"
				$uiHash.Contacts_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Azure Active Directory Contacts $($_.exception.Message)...Please Wait"
				$uiHash.Contacts_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Azure Active Directory DeletedUsers
	If ($CheckedArray.contains("DeletedUsers_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Azure Active Directory Deleted Users...Please Wait"
		})

		try{
			$temp = "<AADDeletedUsers> `r`n"

			$DeletedUsers =  Get-MsolUser -All -ReturnDeletedUsers | Select-Object SignInName, UserPrincipalName, DisplayName, SoftDeletionTimestamp, IsLicensed, @{Name="License"; Expression = {$_.licenses.accountskuid}} 

			ForEach ($DeletedUser in $DeletedUsers) { 
				If (-NOT [System.String]::IsNullOrEmpty($DeletedUser)) {  
					$temp += "<AADDeletedUser> `r`n" +
						"<SignInName>$($DeletedUser.SignInName)</SignInName> `r`n" + 
						"<UserPrincipalName>$($DeletedUser.UserPrincipalName)</UserPrincipalName> `r`n" + 
						"<DisplayName>$($DeletedUser.DisplayName)</DisplayName> `r`n" + 
						"<SoftDeletionTimestamp>$($DeletedUser.SoftDeletionTimestamp)</SoftDeletionTimestamp> `r`n" + 
						"<IsLicensed>$($DeletedUser.IsLicensed)</IsLicensed> `r`n" + 
						"<License>$($DeletedUser.License)</License> `r`n" + 
						"</AADDeletedUser> `r`n"				
				}
			}

			$temp += "</AADDeletedUsers>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Azure Active Directory Deleted Users...Please Wait"
				$uiHash.DeletedUsers_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Azure Active Directory Deleted Users $($_.exception.Message)...Please Wait"
				$uiHash.DeletedUsers_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Azure Active Directory Domains
	If ($CheckedArray.contains("Domains_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Azure Active Directory Domains...Please Wait"
		})

		try{
			$temp = "<AADDomains> `r`n"

			$Domains =  Get-MsolDomain | Select-Object Name, Status, Authentications

			ForEach ($Domain in $Domains) { 
				If (-NOT [System.String]::IsNullOrEmpty($Domain)) {  
					$temp += "<AADDomain> `r`n" +
						"<Name>$($Domain.Name)</Name> `r`n" + 
						"<Status>$($Domain.Status)</Status> `r`n" + 
						"<Authentications>$($Domain.Authentications)</Authentications> `r`n" +
						"</AADDomain> `r`n"				
				}
			}

			$temp += "</AADDomains>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Azure Active Directory Domains...Please Wait"
				$uiHash.Domains_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Azure Active Directory Domains $($_.exception.Message)...Please Wait"
				$uiHash.Domains_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Azure Active Directory Subscriptions
	If ($CheckedArray.contains("Subscriptions_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Azure Active Directory Subscriptions...Please Wait"
		})

		try{
			$temp = "<AADSubscriptions> `r`n"

			$Subscriptions = Get-MsolAccountSku | Select-Object AccountSkuID, ActiveUnits, ConsumedUnits, LockedOutUnits, SkuPartNumber
				

			ForEach ($Subscription in $Subscriptions) { 
				If (-NOT [System.String]::IsNullOrEmpty($Subscription)) { 
					$NextLifecycleDate = Get-MsolSubscription | where-object{$_.SkuPartNumber -eq $Subscription.SkuPartNumber} | Select-Object NextLifecycleDate
					$NextLifecycleDate = $NextLifecycleDate.NextLifecycleDate.tostring().replace(" 00:00:00","")
					$temp += "<AADSubscription> `r`n" +
						"<AccountSkuID>$($Subscription.AccountSkuID)</AccountSkuID> `r`n" + 
						"<ActiveUnits>$($Subscription.ActiveUnits)</ActiveUnits> `r`n" + 
						"<ConsumedUnits>$($Subscription.ConsumedUnits)</ConsumedUnits> `r`n" +
						"<LockedOutUnits>$($Subscription.LockedOutUnits)</LockedOutUnits> `r`n" +
						"<NextLifecycleDate>$($NextLifecycleDate)</NextLifecycleDate> `r`n" +
						"</AADSubscription> `r`n"				
				}
			}

			$temp += "</AADSubscriptions>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Azure Active Directory Subscriptions...Please Wait"
				$uiHash.Subscriptions_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Azure Active Directory Subscriptions $($_.exception.Message)...Please Wait"
				$uiHash.Subscriptions_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Azure Active Directory Roles
	If ($CheckedArray.contains("Roles_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Azure Active Directory Roles...Please Wait"
		})

		try{
			$temp = "<AADRoles> `r`n"

			$Roles = Get-MsolRole | Select-Object Name, IsEnabled
				

			ForEach ($Role in $Roles) { 
				If (-NOT [System.String]::IsNullOrEmpty($Role)) { 
					$NextLifecycleDate = Get-MsolRole | where-object{$_.SkuPartNumber -eq $Role.SkuPartNumber} | Select-Object NextLifecycleDate
					$NextLifecycleDate = $NextLifecycleDate.NextLifecycleDate.tostring().replace(" 00:00:00","")
					$temp += "<AADRole> `r`n" +
						"<Name>$($Role.Name)</Name> `r`n" + 
						"<IsEnabled>$($Role.IsEnabled)</IsEnabled> `r`n" + 
						"</AADRole> `r`n"				
				}
			}

			$temp += "</AADRoles>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Azure Active Directory Roles...Please Wait"
				$uiHash.Roles_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Azure Active Directory Roles $($_.exception.Message)...Please Wait"
				$uiHash.Roles_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}
	#endRegion

	#region Exchange Online
	#Check Exchange Mailboxes
	If ($CheckedArray.contains("ExchangeMailboxes_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Exchange Online Mailboxes...Please Wait"
		})

		try{
			$temp = "<ExchangeMailboxes> `r`n"

			$Mailboxes = Get-Mailbox | Sort-Object DisplayName | Select-Object DisplayName, Alias, PrimarySMTPAddress, ArchiveStatus, UsageLocation, WhenMailboxCreated

			ForEach ($Mailbox in $Mailboxes) { 
				If (-NOT [System.String]::IsNullOrEmpty($Mailbox)) {
					$statistics = Get-MailboxStatistics $Mailbox.Alias -WarningAction:SilentlyContinue | Select-Object ItemCount, TotalItemSize, LastLogonTime, @{name="TotalItemSizeMB"; expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}
					
					$temp += "<ExchangeMailbox> `r`n" +
						"<DisplayName>$($Mailbox.DisplayName)</DisplayName> `r`n" + 
						"<Alias>$($Mailbox.Alias)</Alias> `r`n" + 
						"<PrimarySMTPAddress>$($Mailbox.PrimarySMTPAddress)</PrimarySMTPAddress> `r`n" + 
						"<ArchiveStatus>$($Mailbox.ArchiveStatus)</ArchiveStatus> `r`n" + 
						"<UsageLocation>$($Mailbox.UsageLocation)</UsageLocation> `r`n" + 
						"<WhenMailboxCreated>$($Mailbox.WhenMailboxCreated)</WhenMailboxCreated> `r`n" +
						"<ItemCount>$($statistics.ItemCount)</ItemCount> `r`n" + 
						"<TotalItemSize>$($statistics.TotalItemSize)</TotalItemSize> `r`n" + 
						"<TotalItemSizeMB>$($statistics.TotalItemSizeMB)</TotalItemSizeMB> `r`n" +
						"<LastLogonTime>$($statistics.LastLogonTime)</LastLogonTime> `r`n" +
						"</ExchangeMailbox> `r`n"				
				}
			}

			$temp += "</ExchangeMailboxes>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Exchange Online Mailboxes...Please Wait"
				$uiHash.ExchangeMailboxes_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Exchange Online Mailboxes $($_.exception.Message)...Please Wait"
				$uiHash.ExchangeMailboxes_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Exchange Groups
	If ($CheckedArray.contains("ExchangeGroups_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Exchange Online Groups...Please Wait"
		})

		try{
			$temp = "<ExchangeGroups> `r`n"

			$Groups = Get-Group | Where-Object{$_.RecipientTypeDetails -ne "RoleGroup"} | Sort-Object DisplayName | Select-Object DisplayName, RecipientTypeDetails, @{Name="Owner"; Expression = {$_.ManagedBy}}, WindowsEmailAddress

			ForEach ($Group in $Groups) { 
				If (-NOT [System.String]::IsNullOrEmpty($Group)) { 
					$temp += "<ExchangeGroup> `r`n" +
						"<DisplayName>$($Group.DisplayName)</DisplayName> `r`n" + 
						"<RecipientTypeDetails>$($Group.RecipientTypeDetails)</RecipientTypeDetails> `r`n" + 
						"<Owner>$($Group.Owner)</Owner> `r`n" + 
						"<WindowsEmailAddress>$($Group.WindowsEmailAddress)</WindowsEmailAddress>" +
						"</ExchangeGroup> `r`n"				
				}
			}

			$temp += "</ExchangeGroups>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Exchange Online Groups...Please Wait"
				$uiHash.ExchangeGroups_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Exchange Online Groups $($_.exception.Message)...Please Wait"
				$uiHash.ExchangeGroups_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Exchange Devices
	If ($CheckedArray.contains("ExchangeDevices_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Exchange Online Devices...Please Wait"
		})

		try{
			$temp = "<ExchangeDevices> `r`n"

			$Devices = Get-MobileDevice | Select-Object Name, FriendlyName, DeviceOS, DeviceType, FirstSyncTime, IsValid, DeviceAccessState, @{Name="Owner"; Expression = {$_.Id.split("\")[0]}}

			ForEach ($Device in $Devices) { 
				If (-NOT [System.String]::IsNullOrEmpty($Device)) { 
					$temp += "<ExchangeDevice> `r`n" +
						"<Name>$($Device.Name)</Name> `r`n" + 
						"<FriendlyName>$($Device.FriendlyName)</FriendlyName> `r`n" + 
						"<DeviceOS>$($Device.DeviceOS)</DeviceOS> `r`n" + 
						"<DeviceType>$($Device.DeviceType)</DeviceType>" +
						"<FirstSyncTime>$($Device.FirstSyncTime)</FirstSyncTime> `r`n" + 
						"<DeviceAccessState>$($Device.DeviceAccessState)</DeviceAccessState> `r`n" + 
						"<Owner>$($Device.Owner)</Owner> `r`n" + 
						"</ExchangeDevice> `r`n"				
				}
			}

			$temp += "</ExchangeDevices>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Exchange Online Devices...Please Wait"
				$uiHash.ExchangeDevices_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Exchange Online Devices $($_.exception.Message)...Please Wait"
				$uiHash.ExchangeDevices_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Exchange Contacts
	If ($CheckedArray.contains("ExchangeContacts_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Exchange Online Contacts...Please Wait"
		})

		try{
			$temp = "<ExchangeContacts> `r`n"

			$Contacts = Get-Contact | select-object DisplayName, WindowsEmailAddress, WhenCreated

			ForEach ($Contact in $Contacts) { 
				If (-NOT [System.String]::IsNullOrEmpty($Contact)) { 
					$temp += "<ExchangeContact> `r`n" +
						"<DisplayName>$($Contact.DisplayName)</DisplayName> `r`n" + 
						"<WindowsEmailAddress>$($Contact.WindowsEmailAddress)</WindowsEmailAddress> `r`n" + 
						"<WhenCreated>$($Contact.WhenCreated)</WhenCreated> `r`n" + 
						"</ExchangeContact> `r`n"				
				}
			}

			$temp += "</ExchangeContacts>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Exchange Online Contacts...Please Wait"
				$uiHash.ExchangeContacts_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Exchange Online Contacts $($_.exception.Message)...Please Wait"
				$uiHash.ExchangeContacts_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Exchange Archives
	If ($CheckedArray.contains("ExchangeArchives_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Exchange Online Archives...Please Wait"
		})

		try{
			$temp = "<ExchangeArchives> `r`n"

			$Archives = Get-Mailbox -Archive | sort-object DisplayName | Select-Object DisplayName, Alias, PrimarySMTPAddress, ArchiveStatus, UsageLocation, WhenMailboxCreated

			ForEach ($Archive in $Archives) { 
				If (-NOT [System.String]::IsNullOrEmpty($Archive)) { 
					$statistics = Get-MailboxStatistics $Archive.Alias -WarningAction:SilentlyContinue | Select-Object ItemCount, TotalItemSize, LastLogonTime, @{name="TotalItemSizeMB"; expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}
					
					$temp += "<ExchangeArchive> `r`n" +
						"<DisplayName>$($Archive.DisplayName)</DisplayName> `r`n" + 
						"<Alias>$($Archive.Alias)</Alias> `r`n" + 
						"<PrimarySMTPAddress>$($Archive.PrimarySMTPAddress)</PrimarySMTPAddress> `r`n" + 
						"<ArchiveStatus>$($Archive.ArchiveStatus)</ArchiveStatus> `r`n" + 
						"<UsageLocation>$($Archive.UsageLocation)</UsageLocation> `r`n" + 
						"<WhenMailboxCreated>$($Archive.WhenMailboxCreated)</WhenMailboxCreated> `r`n" + 
						"<ItemCount>$($statistics.ItemCount)</ItemCount> `r`n" + 
						"<TotalItemSize>$($statistics.TotalItemSize)</TotalItemSize> `r`n" + 
						"<TotalItemSizeMB>$($statistics.TotalItemSizeMB)</TotalItemSizeMB> `r`n" +
						"<LastLogonTime>$($statistics.LastLogonTime)</LastLogonTime> `r`n" +
						"</ExchangeArchive> `r`n"				
				}
			}

			$temp += "</ExchangeArchives>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Exchange Online Archives...Please Wait"
				$uiHash.ExchangeArchives_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Exchange Online Archives $($_.exception.Message)...Please Wait"
				$uiHash.ExchangeArchives_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Exchange PublicFolders
	If ($CheckedArray.contains("ExchangePublicFolders_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Exchange Online Public Folders...Please Wait"
		})

		try{
			$temp = "<ExchangePublicFolders> `r`n"

			$PublicFolders = get-publicfolder \ -recurse | get-publicfolderstatistics | Select-Object Name, ItemCount, CreationTime, LastAccessTime, LastUserAccessTime, TotalItemSize, @{name="TotalItemSizeMB"; expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}} 

			ForEach ($PublicFolder in $PublicFolders) { 
				If (-NOT [System.String]::IsNullOrEmpty($PublicFolder)) { 
					$temp += "<ExchangePublicFolder> `r`n" +
						"<Name>$($PublicFolder.Name)</Name> `r`n" + 
						"<ItemCount>$($PublicFolder.ItemCount)</ItemCount> `r`n" + 
						"<CreationTime>$($PublicFolder.CreationTime)</CreationTime> `r`n" +  
						"<LastAccessTime>$($PublicFolder.LastAccessTime)</LastAccessTime> `r`n" + 
						"<LastUserAccessTime>$($PublicFolder.LastUserAccessTime)</LastUserAccessTime> `r`n" + 
						"<TotalItemSize>$($PublicFolder.TotalItemSize)</TotalItemSize> `r`n" +  
						"<TotalItemSizeMB>$($PublicFolder.TotalItemSizeMB)</TotalItemSizeMB> `r`n" +
						"</ExchangePublicFolder> `r`n"				
				}
			}

			$temp += "</ExchangePublicFolders>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Exchange Online Public Folders...Please Wait"
				$uiHash.ExchangePublicFolders_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Exchange Online Public Folders $($_.exception.Message)...Please Wait"
				$uiHash.ExchangePublicFolders_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check Exchange RetentionPolicies
	If ($CheckedArray.contains("ExchangeRetentionPolicies_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all Exchange Online Retention Policies...Please Wait"
		})

		try{
			$temp = "<ExchangeRetentionPolicies> `r`n"

			$RetentionPolicies = Get-RetentionPolicy

			ForEach ($RetentionPolicie in $RetentionPolicies) { 
				If (-NOT [System.String]::IsNullOrEmpty($RetentionPolicie)) { 
					$tags = $RetentionPolicies.retentionpolicytaglinks
					foreach ($tag in $tags){
						$temp += "<ExchangeRetentionPolicie> `r`n" +
							"<Policy>$($RetentionPolicie.name)</Policy> `r`n" + 
							"<PolicyTag>$($tag)</PolicyTag> `r`n" + 
							"</ExchangeRetentionPolicie> `r`n"		
					}		
				}
			}

			$temp += "</ExchangeRetentionPolicies>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved Exchange Online Retention Policies...Please Wait"
				$uiHash.ExchangeRetentionPolicies_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving Exchange Online Retention Policies $($_.exception.Message)...Please Wait"
				$uiHash.ExchangeRetentionPolicies_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}
	#endRegion

	#region SharePoint Online
	#Check SharePoint Online Site Collections
	If ($CheckedArray.contains("SiteCollections_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all SharePoint Online Site Collections...Please Wait"
		})

		try{
			$temp = "<SiteCollections> `r`n"

			$SiteCollections = Get-SPOSite | Sort-Object Url | Select-Object Status, Url, Title, StorageUsageCurrent, StorageQuota, LocaleId, Template, WebsCount, LastContentModifiedDate

			ForEach ($SiteCollection in $SiteCollections) { 
				If (-NOT [System.String]::IsNullOrEmpty($SiteCollection)) { 
					$temp += "<SiteCollection> `r`n" +
						"<Status>$($SiteCollection.Status)</Status> `r`n" + 
						"<Url>$($SiteCollection.Url)</Url> `r`n" + 
						"<Title>$($SiteCollection.Title)</Title> `r`n" +  
						"<StorageUsageCurrent>$($SiteCollection.StorageUsageCurrent)</StorageUsageCurrent> `r`n" + 
						"<StorageQuota>$($SiteCollection.StorageQuota)</StorageQuota> `r`n" + 
						"<LocaleId>$($SiteCollection.LocaleId)</LocaleId> `r`n" +  
						"<WebsCount>$($SiteCollection.WebsCount)</WebsCount> `r`n" +
						"<Template>$($SiteCollection.Template)</Template> `r`n" +  
						"<LastContentModifiedDate>$($SiteCollection.LastContentModifiedDate)</LastContentModifiedDate> `r`n" +
						"</SiteCollection> `r`n"				
				}
			}

			$temp += "</SiteCollections>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved SharePoint Online Site Collections...Please Wait"
				$uiHash.SiteCollections_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving SharePoint Online Site Collections $($_.exception.Message)...Please Wait"
				$uiHash.SiteCollections_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check SharePoint Online Webs
	If ($CheckedArray.contains("Webs_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all SharePoint Online Webs...Please Wait"
		})

		try{
			$temp = "<Webs> `r`n"

			if ($SPOCredentials){
				#Do nothing
			}			
			else{$SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.Username, $Credentials.password)}
			
			$Webs = @()

			if($ManualSiteCollection){
				$Webs = Get-SPOWebs -url $ManualSiteCollection -spocredentials $SPOCredentials
				$webs = $webs | sort-object url
			}
			else{
				$SiteCollections = Get-SPOSite | Select-Object Url
				foreach($SiteCollection in $SiteCollections){
					$Webs += Get-SPOWebs -url $SiteCollection.url -spocredentials $SPOCredentials
				}
				$webs = $webs | sort-object url
			}

			ForEach ($Web in $Webs) { 
				If (-NOT [System.String]::IsNullOrEmpty($Web)) { 
					$temp += "<Web> `r`n" +
						"<Url>$($Web.url)</Url> `r`n" +
						"<ServerRelativeUrl>$($Web.ServerRelativeUrl)</ServerRelativeUrl> `r`n" + 
						"<title>$($Web.title)</title> `r`n" + 
						"<created>$($Web.created)</created> `r`n" +  
						"<Template>$($web.WebTemplate)</Template> `r`n" +  
						"</Web> `r`n"				
				}
			}

			$temp += "</Webs>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved SharePoint Online Webs...Please Wait"
				$uiHash.Webs_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving SharePoint Online Webs $($_.exception.Message)...Please Wait"
				$uiHash.Webs_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check SharePoint Online Teams
	If ($CheckedArray.contains("Teams_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all SharePoint Online Teams...Please Wait"
		})

		try{
			$temp = "<Teams> `r`n"
			
			$Teams = @()

			Connect-MicrosoftTeams -Credential $credentials
			$Teams = get-team
			
			ForEach ($Team in $Teams) { 
				$owners = get-teamuser -groupid $team.groupid | where-object {$_.role -eq "owner"}
				$ownerString = ""
				foreach($owner in $owners){
					$ownerString += "$($owner.user);"
				}

				If (-NOT [System.String]::IsNullOrEmpty($Team)) { 
					$temp += "<Team> `r`n" +
						"<ID>$($team.GroupId)</ID> `r`n" +
						"<DisplayName>$($Team.DisplayName)</DisplayName> `r`n" + 
						"<Owners>$($ownerString)</Owners> `r`n" +  
						"</Team> `r`n"				
				}
			}

			$temp += "</Teams>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved SharePoint Online Teams...Please Wait"
				$uiHash.Teams_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving SharePoint Online Teams $($_.exception.Message)...Please Wait"
				$uiHash.Teams_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check SharePoint Online ContentTypes
	If ($CheckedArray.contains("ContentTypes_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all SharePoint Online Content Types...Please Wait"
		})

		try{
			$temp = "<ContentTypes> `r`n"
			
			$ContentTypes = @()

			if($ManualSiteCollection){
				Connect-PnPOnline -url $ManualSiteCollection -Credentials $credentials
				$ContentTypes = Get-PnPContentType -Connection $connection | sort-object Group
				
				ForEach ($ContentType in $ContentTypes) { 
					If (-NOT [System.String]::IsNullOrEmpty($ContentType)) { 
						$temp += "<ContentType> `r`n" +
							"<Url>$($ManualSiteCollection)</Url> `r`n" +
							"<Name>$($ContentType.Name)</Name> `r`n" + 
							"<Id>$($ContentType.Id)</Id> `r`n" + 
							"<Group>$($ContentType.Group)</Group> `r`n" +  
							"</ContentType> `r`n"				
					}
				}
			}
			else{
				$SiteCollections = Get-SPOSite | Select-Object Url
				foreach($SiteCollection in $SiteCollections){
					Connect-PnPOnline -url $SiteCollection.url -Credentials $credentials
					$ContentTypes += Get-PnPContentType -Connection $connection | sort-object Group

					ForEach ($ContentType in $ContentTypes) { 
						If (-NOT [System.String]::IsNullOrEmpty($ContentType)) { 
							$temp += "<ContentType> `r`n" +
								"<Url>$($SiteCollection.url)</Url> `r`n" +
								"<Name>$($ContentType.Name)</Name> `r`n" + 
								"<Id>$($ContentType.Id)</Id> `r`n" + 
								"<Group>$($ContentType.Group)</Group> `r`n" +  
								"</ContentType> `r`n"				
						}
					}
				}
			}

			$temp += "</ContentTypes>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved SharePoint Online Content Types...Please Wait"
				$uiHash.ContentTypes_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving SharePoint Online Content Types $($_.exception.Message)...Please Wait"
				$uiHash.ContentTypes_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check SharePoint Online Lists
	If ($CheckedArray.contains("Lists_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all SharePoint Online Lists...Please Wait"
		})

		try{
			$temp = "<Lists> `r`n"

			if ($SPOCredentials){
				#Do nothing
			}			
			else{$SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.Username, $Credentials.password)}

			$Lists = @()

			if($ManualSiteCollection){
				$Lists = Get-SPOLists -url $ManualSiteCollection -spocredentials $SPOCredentials
				$Lists = $Lists | sort-object url
			}
			else{
				$SiteCollections = Get-SPOSite | Select-Object Url
				foreach($SiteCollection in $SiteCollections){
					$Lists += Get-SPOLists -url $SiteCollection.url -spocredentials $SPOCredentials
				}
				$Lists = $Lists | sort-object url
			}

			ForEach ($List in $Lists) { 
				If (-NOT [System.String]::IsNullOrEmpty($List)) { 
					$temp += "<List> `r`n" +
						"<Url>$($List.url)</Url> `r`n" +
						"<title>$($List.title)</title> `r`n" + 
						"<itemcount>$($List.itemcount)</itemcount> `r`n" + 
						"<BaseTemplate>$($List.BaseTemplate)</BaseTemplate> `r`n" +  
						"<created>$($List.created)</created> `r`n" + 
						"<LastItemModifiedDate>$($List.LastItemModifiedDate)</LastItemModifiedDate> `r`n" + 
						"<Id>$($List.Id)</Id> `r`n" + 
						"</List> `r`n"				
				}
			}

			$temp += "</Lists>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved SharePoint Online Lists...Please Wait"
				$uiHash.Lists_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving SharePoint Online Lists $($_.exception.Message)...Please Wait"
				$uiHash.Lists_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check SharePoint Online Features
	If ($CheckedArray.contains("Features_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all SharePoint Online Features...Please Wait"
		})

		try{
			$temp = "<Features> `r`n"

			if ($SPOCredentials){
				#Do nothing
			}			
			else{$SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.Username, $Credentials.password)}

			$Features = @()

			if($ManualSiteCollection){
				$Features = Get-SPOFeatures -url $ManualSiteCollection -spocredentials $SPOCredentials
				$Features = $Features | sort-object url
			}
			else{
				$SiteCollections = Get-SPOSite | Select-Object Url
				foreach($SiteCollection in $SiteCollections){
					$Features += Get-SPOFeatures -url $SiteCollection.url -spocredentials $SPOCredentials
				}
				$Features = $Features | sort-object url
			}

			ForEach ($Feature in $Features) { 
				If (-NOT [System.String]::IsNullOrEmpty($Feature)) { 
					$temp += "<Feature> `r`n" +
						"<Url>$($Feature.url)</Url> `r`n" +
						"<DisplayName>$($Feature.DisplayName)</DisplayName> `r`n" + 
						"<active>$($Feature.active)</active> `r`n" +
						"</Feature> `r`n"				
				}
			}

			$temp += "</Features>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved SharePoint Online Features...Please Wait"
				$uiHash.Features_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving SharePoint Online Features $($_.exception.Message)...Please Wait"
				$uiHash.Features_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}

	#Check SharePoint Online Groups
	If ($CheckedArray.contains("SharePointGroups_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Retrieving all SharePoint Online Groups...Please Wait"
		})

		try{
			$temp = "<SharePointGroups> `r`n"

			if ($SPOCredentials){
				#Do nothing
			}			
			else{$SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.Username, $Credentials.password)}

			$Groups = @()
			
			if($ManualSiteCollection){
				$Groups = Get-SPOGroups -url $ManualSiteCollection -spocredentials $SPOCredentials
				$Groups = $Groups | sort-object url
			}
			else{
				$SiteCollections = Get-SPOSite | Select-Object Url
				foreach($SiteCollection in $SiteCollections){
					$Groups += Get-SPOGroups -url $SiteCollection.url -spocredentials $SPOCredentials
				}
				$Groups = $Groups | sort-object url
			}

			ForEach ($Group in $Groups) { 
				If (-NOT [System.String]::IsNullOrEmpty($Group)) { 
					$temp += "<SharePointGroup> `r`n" +
						"<Url>$($Group.url)</Url> `r`n" +
						"<DisplayName>$($Group.DisplayName)</DisplayName> `r`n" + 
						"<Permission>$($Group.Permission)</Permission> `r`n" +
						"</SharePointGroup> `r`n"				
				}
			}

			$temp += "</SharePointGroups>"
			Add-Content -Path $LogPath -Value $temp

			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
				update-MessageBlock -type "Okay" -Message "Retrieved SharePoint Online Groups...Please Wait"
				$uiHash.SharePointGroups_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				update-MessageBlock -type "Critical" -Message "Error retrieving SharePoint Online Groups $($_.exception.Message)...Please Wait"
				$uiHash.SharePointGroups_Image.Source = "$pwd\Images\Check_Error.ico"
			})			
		}
	}
	#endRegion

	$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
		update-MessageBlock -type "Status" -Message "Completed retrieving information"
	})

	Add-Content -Path $LogPath -Value "</Default>"

	New-HTMLReport -XMLPath $LogPath -CheckedArray $CheckedArray
}

function Get-SPOWebs(){
	param(
		$URL,
		$SPOCredentials
	)
	$context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
	$context.Credentials = $SPOCredentials
	
	$web = $context.Web
	$context.Load($web)
	$context.Load($web.Webs)
	
	try{
		$context.ExecuteQuery()
		
		If (-NOT [System.String]::IsNullOrEmpty($web)) {
			$WebTemp = New-Object -TypeName PSObject
				$WebTemp | Add-Member -MemberType NoteProperty -Name "url" -Value $web.url
				$WebTemp | Add-Member -MemberType NoteProperty -Name "ServerRelativeUrl" -Value $web.ServerRelativeUrl
				$WebTemp | Add-Member -MemberType NoteProperty -Name "Title" -Value $web.Title
				$WebTemp | Add-Member -MemberType NoteProperty -Name "created" -Value $web.created
				$WebTemp | Add-Member -MemberType NoteProperty -Name "WebTemplate" -Value $web.WebTemplate
					
			$Webs += $WebTemp
		}

		foreach($web in $web.Webs) {	
			Get-SPOWebs -url $web.Url -spocredentials $SPOCredentials
		}

		
	}
	catch{
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Warning" -Message "Error retrieving SharePoint Online Web $($URL) $($_.exception.Message)...Please Wait"
		})			
	}

	return $Webs
 
}

function Get-SPOLists(){
	param(
		$URL,
		$SPOCredentials
	)
	$context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
	$context.Credentials = $SPOCredentials
	
	$web = $context.Web
	$context.Load($web)
	$context.Load($web.Webs)
	$context.load($web.lists)
	
	try{
		$context.ExecuteQuery()
		
		$TempArray = @()
		If (-NOT [System.String]::IsNullOrEmpty($web)) {
			foreach($list in $web.lists){
				$ListTemp = New-Object -TypeName PSObject
					$ListTemp | Add-Member -MemberType NoteProperty -Name "url" -Value $web.url
					$ListTemp | Add-Member -MemberType NoteProperty -Name "title" -Value $List.title
					$ListTemp | Add-Member -MemberType NoteProperty -Name "itemcount" -Value $List.itemcount
					$ListTemp | Add-Member -MemberType NoteProperty -Name "BaseTemplate" -Value $List.BaseTemplate
					$ListTemp | Add-Member -MemberType NoteProperty -Name "created" -Value $List.created
					$ListTemp | Add-Member -MemberType NoteProperty -Name "LastItemModifiedDate" -Value $List.LastItemModifiedDate
					$ListTemp | Add-Member -MemberType NoteProperty -Name "Id" -Value $List.Id
						
					$TempArray += $ListTemp
			}
			$Lists += $TempArray
		}

		foreach($web in $web.Webs) {	
			Get-SPOLists -url $web.Url -spocredentials $SPOCredentials
		}

		
	}
	catch{
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Warning" -Message "Error retrieving SharePoint Online List $($URL)/$($List.title) $($_.exception.Message)...Please Wait"
		})			
	}

	return $Lists
 
}

function Get-SPOFeatures(){
	param(
		$URL,
		$SPOCredentials
	)
	$context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
	$context.Credentials = $SPOCredentials
	
	$web = $context.Web
	$context.Load($web)
	$context.Load($web.Webs)
	$context.load($web.Features)
	
	try{
		$context.ExecuteQuery()
		$TempArray = @()

		If (-NOT [System.String]::IsNullOrEmpty($web)) {
			foreach($Feature in $web.Features){
				if($web.Features.DefinitionId -eq $null){      
					$active = "False"    
				}      
				else{      
					$active = "True" 
				}   
					
				$FeatureTemp = New-Object -TypeName PSObject
					$FeatureTemp | Add-Member -MemberType NoteProperty -Name "url" -Value $web.url
					$FeatureTemp | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $Feature.DisplayName
					$FeatureTemp | Add-Member -MemberType NoteProperty -Name "active" -Value $active
						
				$TempArray += $FeatureTemp
			}
			$Features += $TempArray
		}

		foreach($web in $web.Webs) {	
			Get-SPOFeatures -url $web.Url -spocredentials $SPOCredentials -Features $Features
		}

		
	}
	catch{
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Warning" -Message "Error retrieving SharePoint Online Feature $($URL) - $($Feature.title) $($_.exception.Message)...Please Wait"
		})			
	}

	return $Features
 
}

function Get-SPOGroups(){
	param(
		$URL,
		$SPOCredentials
	)
	
	$context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
	$context.Credentials = $SPOCredentials
	
	$web = $context.Web
	$context.Load($web)
	$context.Load($web.Webs)
	$context.load($web.RoleAssignments)

	try{
		$context.ExecuteQuery()
		$TempArray = @()

		If (-NOT [System.String]::IsNullOrEmpty($web)) {
			foreach($RoleAssignment in $web.RoleAssignments){
				$context.Load($RoleAssignment.Member)
				$context.Load($RoleAssignment.RoleDefinitionBindings)
				$context.ExecuteQuery()

				foreach($RoleDefinition in $RoleAssignment.RoleDefinitionBindings)
				{
				$context.Load($RoleDefinition)
				$context.ExecuteQuery()
				
					if ($RoleDefinition.Name -ne "Limited Access") 
					{ 
						$GroupTemp = New-Object -TypeName PSObject
							$GroupTemp | Add-Member -MemberType NoteProperty -Name "url" -Value $web.url
							$GroupTemp | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $RoleAssignment.Member.Title
							$GroupTemp | Add-Member -MemberType NoteProperty -Name "Permission" -Value $RoleDefinition.Name
								
						$TempArray += $GroupTemp
					}
				}
			}
			$Groups += $TempArray
		}

		foreach($web in $web.Webs) {	
			Get-SPOGroups -url $web.Url -spocredentials $SPOCredentials
		}

		
	}
	catch{
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Warning" -Message "Error retrieving SharePoint Online Group $($URL) - $($Group.title) $($_.exception.Message)...Please Wait"
		})			
	}

	return $Groups
 
}
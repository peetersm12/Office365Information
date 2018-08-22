function New-HTMLReport($XMLPath, $CheckedArray){
	$HTMLPath = $XMLPath.replace("XML", "HTML")
	New-LogFile -FullPath $HTMLPath

	$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
		update-MessageBlock -type "Status" -Message "Generating HTML file...Please Wait"
	})

	$Sections = @()
	$AADLinks = @()

	[xml]$XmlDocument = Get-Content -Path $XMLPath

	#region AAD
	#Check licensed
	If ($CheckedArray.contains("Users_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Azure Active Directory licensed Users section...Please Wait"
		})

		$temp = "<section id=`"AADLicensedUsers`">" +
    		"<button class=`"collapsible`">Azure AD Licensed Users</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>DisplayName</th>" +
			"<th>FirstName</th>" +
			"<th>LastName</th>" +
			"<th>UserPrincipalName</th>" +
			"<th>Title</th>" +
			"<th>Department</th>" +
			"<th>Office</th>" +
			"<th>PhoneNumber</th>" +
			"<th>MobilePhone</th>" +
			"<th>CloudAnchor</th>" +
			"<th>IsLicensed</th>" +
			"<th>License</th>" +
			"</tr>"
			
		foreach($user in $xmldocument.default.aadusers.aaduser){
			if($user.IsLicensed -eq "True"){
				$temp += "<tr>" +
					"<td>$($user.DisplayName)</td>" +
					"<td>$($user.FirstName)</td>" +
					"<td>$($user.LastName)</td>" +
					"<td>$($user.UserPrincipalName)</td>" +
					"<td>$($user.Title)</td>" +
					"<td>$($user.Department)</td>" +
					"<td>$($user.Office)</td>" +
					"<td>$($user.PhoneNumber)</td>" +
					"<td>$($user.MobilePhone)</td>" +
					"<td>$($user.CloudAnchor)</td>" +
					"<td>$($user.IsLicensed)</td>" +
					"<td>$($user.License)</td>" +
					"</tr>"
			}
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$AADLinks += "<a href=`"#AADLicensedUsers`">Azure AD Licensed Users</a>"
	}

	#Check unlicensed
	If ($CheckedArray.contains("Users_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "adding Azure Active Directory unlicensed Users section...Please Wait"
		})

		$temp = "<section id=`"AADUnlicensedUsers`">" +
    		"<button class=`"collapsible`">Azure AD Unlicensed Users</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>DisplayName</th>" +
			"<th>FirstName</th>" +
			"<th>LastName</th>" +
			"<th>UserPrincipalName</th>" +
			"<th>Title</th>" +
			"<th>Department</th>" +
			"<th>Office</th>" +
			"<th>PhoneNumber</th>" +
			"<th>MobilePhone</th>" +
			"<th>CloudAnchor</th>" +
			"<th>IsLicensed</th>" +
			"<th>License</th>" +
			"</tr>"
			
		foreach($user in $xmldocument.default.aadusers.aaduser){
			if($user.IsLicensed -eq "False"){
				$temp += "<tr>" +
					"<td>$($user.DisplayName)</td>" +
					"<td>$($user.FirstName)</td>" +
					"<td>$($user.LastName)</td>" +
					"<td>$($user.UserPrincipalName)</td>" +
					"<td>$($user.Title)</td>" +
					"<td>$($user.Department)</td>" +
					"<td>$($user.Office)</td>" +
					"<td>$($user.PhoneNumber)</td>" +
					"<td>$($user.MobilePhone)</td>" +
					"<td>$($user.CloudAnchor)</td>" +
					"<td>$($user.IsLicensed)</td>" +
					"<td>$($user.License)</td>" +
					"</tr>"
			}
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$AADLinks += "<a href=`"#AADUnlicensedUsers`">Azure AD Unlicensed Users</a>"
	}

	#Check Groups
	If ($CheckedArray.contains("Groups_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Azure Active Directory Groups section...Please Wait"
		})

		$temp = "<section id=`"AADGroups`">" +
    		"<button class=`"collapsible`">Azure AD Groups</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>GroupType</th>" +
			"<th>DisplayName</th>" +
			"<th>EmailAddress</th>" +
			"<th>ValidationStatus</th>" +
			"</tr>"
			
		foreach($group in $xmldocument.default.aadgroups.aadgroup){
			$temp += "<tr>" +
				"<td>$($group.GroupType)</td>" +
				"<td>$($group.DisplayName)</td>" +
				"<td>$($group.EmailAddress)</td>" +
				"<td>$($group.ValidationStatus)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$AADLinks += "<a href=`"#AADGroups`">Azure AD Groups</a>"
	}

	#Check Guests
	If ($CheckedArray.contains("Guests_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Azure Active Directory Guests section...Please Wait"
		})

		$temp = "<section id=`"AADGuests`">" +
    		"<button class=`"collapsible`">Azure AD Guests</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>SignInName</th>" +
			"<th>UserPrincipalName</th>" +
			"<th>DisplayName</th>" +
			"<th>WhenCreated</th>" +
			"</tr>"
			
		foreach($guest in $xmldocument.default.aadguests.aadguest){
			$temp += "<tr>" +
				"<td>$($guest.SignInName)</td>" +
				"<td>$($guest.UserPrincipalName)</td>" +
				"<td>$($guest.DisplayName)</td>" +
				"<td>$($guest.WhenCreated)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$AADLinks += "<a href=`"#AADGuests`">Azure AD Guests</a>"
	}

	#Check Contacts
	If ($CheckedArray.contains("Contacts_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Azure Active Directory Contacts section...Please Wait"
		})

		$temp = "<section id=`"AADContacts`">" +
    		"<button class=`"collapsible`">Azure AD Contacts</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>DisplayName</th>" +
			"<th>EmailAddress</th>" +
			"</tr>"
			
		foreach($Contact in $xmldocument.default.aadContacts.aadContact){
			$temp += "<tr>" +
				"<td>$($Contact.DisplayName)</td>" +
				"<td>$($Contact.EmailAddress)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$AADLinks += "<a href=`"#AADContacts`">Azure AD Contacts</a>"
	}

	#Check DeletedUsers
	If ($CheckedArray.contains("DeletedUsers_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Azure Active Directory Deleted Users section...Please Wait"
		})

		$temp = "<section id=`"AADDeletedUsers`">" +
    		"<button class=`"collapsible`">Azure AD DeletedUsers</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>SignInName</th>" +
			"<th>UserPrincipalName</th>" +
			"<th>DisplayName</th>" +
			"<th>SoftDeletionTimestamp</th>" +
			"<th>IsLicensed</th>" +
			"<th>License</th>" +
			"</tr>"
			
		foreach($DeletedUser in $xmldocument.default.aadDeletedUsers.aadDeletedUser){
			$temp += "<tr>" +
				"<td>$($DeletedUser.SignInName)</td>" +
				"<td>$($DeletedUser.UserPrincipalName)</td>" +
				"<td>$($DeletedUser.DisplayName)</td>" +
				"<td>$($DeletedUser.SoftDeletionTimestamp)</td>" +
				"<td>$($DeletedUser.IsLicensed)</td>" +
				"<td>$($DeletedUser.License)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$AADLinks += "<a href=`"#AADDeletedUsers`">Azure AD Deleted Users</a>"
	}

	#Check Domains
	If ($CheckedArray.contains("Domains_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Azure Active Directory Domains section...Please Wait"
		})

		$temp = "<section id=`"AADDomains`">" +
    		"<button class=`"collapsible`">Azure AD Domains</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Name</th>" +
			"<th>Status</th>" +
			"<th>Authentications</th>" +
			"</tr>"
			
		foreach($Domain in $xmldocument.default.aadDomains.aadDomain){
			$temp += "<tr>" +
				"<td>$($Domain.Name)</td>" +
				"<td>$($Domain.Status)</td>" +
				"<td>$($Domain.Authentications)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$AADLinks += "<a href=`"#AADDomains`">Azure AD Domains</a>"
	}

	#Check Subscriptions
	If ($CheckedArray.contains("Subscriptions_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Azure Active Directory Subscriptions section...Please Wait"
		})

		$temp = "<section id=`"AADSubscriptions`">" +
    		"<button class=`"collapsible`">Azure AD Subscriptions</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>AccountSkuID</th>" +
			"<th>ActiveUnits</th>" +
			"<th>ConsumedUnits</th>" +
			"<th>LockedOutUnits</th>" +
			"<th>NextLifecycleDate</th>" +
			"</tr>"
			
		foreach($Subscription in $xmldocument.default.aadSubscriptions.aadSubscription){
			$temp += "<tr>" +
				"<td>$($Subscription.AccountSkuID)</td>" +
				"<td>$($Subscription.ActiveUnits)</td>" +
				"<td>$($Subscription.ConsumedUnits)</td>" +
				"<td>$($Subscription.LockedOutUnits)</td>" +
				"<td>$($Subscription.NextLifecycleDate)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$AADLinks += "<a href=`"#AADSubscriptions`">Azure AD Subscriptions</a>"
	}

	#Check Roles
	If ($CheckedArray.contains("Roles_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Azure Active Directory Roles section...Please Wait"
		})

		$temp = "<section id=`"AADRoles`">" +
    		"<button class=`"collapsible`">Azure AD Roles</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Name</th>" +
			"<th>IsEnabled</th>" +
			"</tr>"
			
		foreach($Role in $xmldocument.default.aadRoles.aadRole){
			$temp += "<tr>" +
				"<td>$($Role.Name)</td>" +
				"<td>$($Role.IsEnabled)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$AADLinks += "<a href=`"#AADRoles`">Azure AD Roles</a>"
	}
	#endRegion

	#region Exchange Online
	#Check Exchange Mailboxes
	If ($CheckedArray.contains("ExchangeMailboxes_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Exchange Online Mailboxes section...Please Wait"
		})

		$temp = "<section id=`"ExchangeMailbox`">" +
    		"<button class=`"collapsible`">Exchange Online Mailboxes</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>DisplayName</th>" +
			"<th>Alias</th>" +
			"<th>PrimarySMTPAddress</th>" +
			"<th>ArchiveStatus</th>" +
			"<th>UsageLocation</th>" +
			"<th>WhenMailboxCreated</th>" +
			"<th>ItemCount</th>" +
			"<th>TotalItemSize</th>" +
			"<th>TotalItemSizeMB</th>" +
			"<th>LastLogonTime</th>" +
			"</tr>"

		foreach($Mailbox in $xmldocument.default.ExchangeMailboxes.ExchangeMailbox){
			$temp += "<tr>" +
				"<td>$($Mailbox.DisplayName)</td>" +
				"<td>$($Mailbox.Alias)</td>" +
				"<td>$($Mailbox.PrimarySMTPAddress)</td>" +
				"<td>$($Mailbox.ArchiveStatus)</td>" +
				"<td>$($Mailbox.UsageLocation)</td>" +
				"<td>$($Mailbox.WhenMailboxCreated)</td>" +
				"<td>$($Mailbox.ItemCount)</td>" +
				"<td>$($Mailbox.TotalItemSize)</td>" +
				"<td>$($Mailbox.TotalItemSizeMB)</td>" +
				"<td>$($Mailbox.LastLogonTime)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$ExchangeLinks += "<a href=`"#ExchangeMailbox`">Exchange Online Mailboxes</a>"
	}

	#Check Exchange Groups
	If ($CheckedArray.contains("ExchangeGroups_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Exchange Online Groups section...Please Wait"
		})

		$temp = "<section id=`"ExchangeGroup`">" +
    		"<button class=`"collapsible`">Exchange Online Groups</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>DisplayName</th>" +
			"<th>RecipientTypeDetails</th>" +
			"<th>Owner</th>" +
			"<th>WindowsEmailAddress</th>" +
			"</tr>"
			
		foreach($Group in $xmldocument.default.ExchangeGroups.ExchangeGroup){
			$temp += "<tr>" +
				"<td>$($Group.DisplayName)</td>" +
				"<td>$($Group.RecipientTypeDetails)</td>" +
				"<td>$($Group.Owner)</td>" +
				"<td>$($Group.WindowsEmailAddress)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$ExchangeLinks += "<a href=`"#ExchangeGroup`">Exchange Online Groups</a>"
	}

	#Check Exchange Devices
	If ($CheckedArray.contains("ExchangeDevices_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Exchange Online Devices section...Please Wait"
		})

		$temp = "<section id=`"ExchangeDevice`">" +
    		"<button class=`"collapsible`">Exchange Online Devices</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Name</th>" +
			"<th>FriendlyName</th>" +
			"<th>DeviceOS</th>" +
			"<th>DeviceType</th>" +
			"<th>FirstSyncTime</th>" +
			"<th>DeviceAccessState</th>" +
			"<th>Owner</th>" +
			"</tr>"
			
		foreach($Device in $xmldocument.default.ExchangeDevices.ExchangeDevice){
			$temp += "<tr>" +
				"<td>$($Device.Name)</td>" +
				"<td>$($Device.FriendlyName)</td>" +
				"<td>$($Device.DeviceOS)</td>" +
				"<td>$($Device.DeviceType)</td>" +
				"<td>$($Device.FirstSyncTime)</td>" +
				"<td>$($Device.DeviceAccessState)</td>" +
				"<td>$($Device.Owner)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$ExchangeLinks += "<a href=`"#ExchangeDevice`">Exchange Online Devices</a>"
	}

	#Check Exchange Contacts
	If ($CheckedArray.contains("ExchangeContacts_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Exchange Online Contacts section...Please Wait"
		})

		$temp = "<section id=`"ExchangeContact`">" +
    		"<button class=`"collapsible`">Exchange Online Contacts</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>DisplayName</th>" +
			"<th>WindowsEmailAddress</th>" +
			"<th>WhenCreated</th>" +
			"</tr>"
			
		foreach($Contact in $xmldocument.default.ExchangeContacts.ExchangeContact){
			$temp += "<tr>" +
				"<td>$($Contact.DisplayName)</td>" +
				"<td>$($Contact.WindowsEmailAddress)</td>" +
				"<td>$($Contact.WhenCreated)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$ExchangeLinks += "<a href=`"#ExchangeContact`">Exchange Online Contacts</a>"
	}

	#Check Exchange Archives
	If ($CheckedArray.contains("ExchangeArchives_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Exchange Online Archives section...Please Wait"
		})

		$temp = "<section id=`"ExchangeArchive`">" +
    		"<button class=`"collapsible`">Exchange Online Archives</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>DisplayName</th>" +
			"<th>Alias</th>" +
			"<th>PrimarySMTPAddress</th>" +
			"<th>ArchiveStatus</th>" +
			"<th>UsageLocation</th>" +
			"<th>WhenMailboxCreated</th>" +
			"<th>ItemCount</th>" +
			"<th>TotalItemSize</th>" +
			"<th>TotalItemSizeMB</th>" +
			"<th>LastLogonTime</th>" +
			"</tr>"
			
		foreach($Archive in $xmldocument.default.ExchangeArchives.ExchangeArchive){
			$temp += "<tr>" +
				"<td>$($Archive.DisplayName)</td>" +
				"<td>$($Archive.Alias)</td>" +
				"<td>$($Archive.PrimarySMTPAddress)</td>" +
				"<td>$($Archive.ArchiveStatus)</td>" +
				"<td>$($Archive.UsageLocation)</td>" +
				"<td>$($Archive.WhenMailboxCreated)</td>" +
				"<td>$($Archive.ItemCount)</td>" +
				"<td>$($Archive.TotalItemSize)</td>" +
				"<td>$($Archive.TotalItemSizeMB)</td>" +
				"<td>$($Archive.LastLogonTime)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$ExchangeLinks += "<a href=`"#ExchangeArchive`">Exchange Online Archives</a>"
	}

	#Check Exchange PublicFolders
	If ($CheckedArray.contains("ExchangePublicFolders_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Exchange Online Public Folders section...Please Wait"
		})

		$temp = "<section id=`"ExchangePublicFolder`">" +
    		"<button class=`"collapsible`">Exchange Online Public Folders</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Name</th>" +
			"<th>ItemCount</th>" +
			"<th>CreationTime</th>" +
			"<th>LastAccessTime</th>" +
			"<th>LastUserAccessTime</th>" +
			"<th>TotalItemSize</th>" +
			"<th>TotalItemSizeMB</th>" +
			"</tr>"

		foreach($PublicFolder in $xmldocument.default.ExchangePublicFolders.ExchangePublicFolder){
			$temp += "<tr>" +
				"<td>$($PublicFolder.Name)</td>" +
				"<td>$($PublicFolder.ItemCount)</td>" +
				"<td>$($PublicFolder.CreationTime)</td>" +
				"<td>$($PublicFolder.LastAccessTime)</td>" +
				"<td>$($PublicFolder.LastUserAccessTime)</td>" +
				"<td>$($PublicFolder.TotalItemSize)</td>" +
				"<td>$($PublicFolder.TotalItemSizeMB)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$ExchangeLinks += "<a href=`"#ExchangePublicFolder`">Exchange Online Public Folders</a>"
	}

	#Check Exchange RetentionPolicies
	If ($CheckedArray.contains("ExchangeRetentionPolicies_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding Exchange Online Retention Policies section...Please Wait"
		})

		$temp = "<section id=`"ExchangeRetentionPolicie`">" +
    		"<button class=`"collapsible`">Exchange Online Retention Policies</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Policy</th>" +
			"<th>PolicyTag</th>" +
			"</tr>"
			
		foreach($RetentionPolicie in $xmldocument.default.ExchangeRetentionPolicies.ExchangeRetentionPolicie){
			$temp += "<tr>" +
				"<td>$($RetentionPolicie.Policy)</td>" +
				"<td>$($RetentionPolicie.PolicyTag)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$ExchangeLinks += "<a href=`"#ExchangeRetentionPolicie`">Exchange Online Retention Policies</a>"
	}
	#endRegion

	#region SharePoint Online
	#Check SharePoint Online Site Collections
	If ($CheckedArray.contains("SiteCollections_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding SharePoint Online Site Collections section...Please Wait"
		})

		$temp = "<section id=`"SiteCollections`">" +
    		"<button class=`"collapsible`">SharePoint Online Site Collections</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Status</th>" +
			"<th>Url</th>" +
			"<th>Title</th>" +
			"<th>StorageUsageCurrent</th>" +
			"<th>StorageQuota</th>" +
			"<th>LocaleId</th>" +
			"<th>WebsCount</th>" +
			"<th>Template</th>" +
			"<th>LastContentModifiedDate</th>" +
			"</tr>"
			
		foreach($SiteCollection in $xmldocument.default.SiteCollections.SiteCollection){
			$temp += "<tr>" +
				"<td>$($SiteCollection.Status)</td>" +
				"<td>$($SiteCollection.Url)</td>" +
				"<td>$($SiteCollection.Title)</td>" +
				"<td>$($SiteCollection.StorageUsageCurrent)</td>" +
				"<td>$($SiteCollection.StorageQuota)</td>" +
				"<td>$($SiteCollection.LocaleId)</td>" +
				"<td>$($SiteCollection.WebsCount)</td>" +
				"<td>$($SiteCollection.Template)</td>" +
				"<td>$($SiteCollection.LastContentModifiedDate)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$SharePointLinks += "<a href=`"#SiteCollections`">SharePoint Online Site Collections</a>"
	}

	#Check SharePoint Online Webs
	If ($CheckedArray.contains("Webs_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding SharePoint Online Webs section...Please Wait"
		})

		$temp = "<section id=`"Webs`">" +
    		"<button class=`"collapsible`">SharePoint Online Webs</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Url</th>" +
			"<th>ServerRelativeUrl</th>" +
			"<th>title</th>" +
			"<th>created</th>" +
			"<th>Template</th>" +
			"</tr>"
			
		foreach($Web in $xmldocument.default.Webs.Web){
			$temp += "<tr>" +
				"<td>$($Web.Url)</td>" +
				"<td>$($Web.ServerRelativeUrl)</td>" +
				"<td>$($Web.title)</td>" +
				"<td>$($Web.created)</td>" +
				"<td>$($Web.Template)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$SharePointLinks += "<a href=`"#Webs`">SharePoint Online Webs</a>"
	}

	#Check SharePoint Online Teams
	If ($CheckedArray.contains("Teams_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding SharePoint Online Content Types section...Please Wait"
		})

		$temp = "<section id=`"Teams`">" +
    		"<button class=`"collapsible`">SharePoint Online Teams</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>ID</th>" +
			"<th>DisplayName</th>" +
			"<th>Owners</th>" +
			"</tr>"
			
		foreach($Team in $xmldocument.default.Teams.Team){
			$temp += "<tr>" +
				"<td>$($Team.ID)</td>" +
				"<td>$($Team.DisplayName)</td>" +
				"<td>$($Team.Owners)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$SharePointLinks += "<a href=`"#Teams`">SharePoint Online Teams</a>"
	}

	#Check SharePoint Online ContentTypes
	If ($CheckedArray.contains("ContentTypes_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding SharePoint Online Content Types section...Please Wait"
		})

		$temp = "<section id=`"ContentTypes`">" +
    		"<button class=`"collapsible`">SharePoint Online ContentTypes</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Url</th>" +
			"<th>Name</th>" +
			"<th>Id</th>" +
			"<th>Group</th>" +
			"</tr>"
			
		foreach($ContentType in $xmldocument.default.ContentTypes.ContentType){
			$temp += "<tr>" +
				"<td>$($ContentType.Url)</td>" +
				"<td>$($ContentType.Name)</td>" +
				"<td>$($ContentType.Id)</td>" +
				"<td>$($ContentType.Group)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$SharePointLinks += "<a href=`"#ContentTypes`">SharePoint Online ContentTypes</a>"
	}

	#Check SharePoint Online Lists
	If ($CheckedArray.contains("Lists_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding SharePoint Online Lists section...Please Wait"
		})

		$temp = "<section id=`"Lists`">" +
    		"<button class=`"collapsible`">SharePoint Online Lists</button>" +
    		"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Url</th>" +
			"<th>title</th>" +
			"<th>itemcount</th>" +
			"<th>BaseTemplate</th>" +
			"<th>created</th>" +
			"<th>LastItemModifiedDate</th>" +
			"<th>Id</th>" +
			"</tr>"
			
		foreach($List in $xmldocument.default.Lists.List){
			$temp += "<tr>" +
				"<td>$($List.Url)</td>" +
				"<td>$($List.title)</td>" +
				"<td>$($List.itemcount)</td>" +
				"<td>$($List.BaseTemplate)</td>" +
				"<td>$($List.created)</td>" +
				"<td>$($List.LastItemModifiedDate)</td>" +
				"<td>$($List.Id)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$SharePointLinks += "<a href=`"#Lists`">SharePoint Online Lists</a>"
	}

	#Check SharePoint Online Features
	If ($CheckedArray.contains("Features_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding SharePoint Online Features section...Please Wait"
		})

		$temp = "<section id=`"Features`">" +
			"<button class=`"collapsible`">SharePoint Online Features</button>" +
			"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Url</th>" +
			"<th>DisplayName</th>" +
			"<th>active</th>" +
			"</tr>"
			
		foreach($Feature in $xmldocument.default.Features.Feature){
			$temp += "<tr>" +
				"<td>$($Feature.Url)</td>" +
				"<td>$($Feature.DisplayName)</td>" +
				"<td>$($Feature.active)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$SharePointLinks += "<a href=`"#Features`">SharePoint Online Features</a>"
	}

	#Check SharePoint Online Groups
	If ($CheckedArray.contains("SharePointGroups_CheckBox")){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			update-MessageBlock -type "Status" -Message "Adding SharePoint Online Groups section...Please Wait"
		})

		$temp = "<section id=`"SharePointGroups`">" +
			"<button class=`"collapsible`">SharePoint Online Groups</button>" +
			"<div class=`"content`">" +
			"<p><table>" +
			"<tr>" +
			"<th>Url</th>" +
			"<th>DisplayName</th>" +
			"<th>Permission</th>" +
			"</tr>"
			
		foreach($Group in $xmldocument.default.SharePointGroups.SharePointGroup){
			$temp += "<tr>" +
				"<td>$($Group.Url)</td>" +
				"<td>$($Group.DisplayName)</td>" +
				"<td>$($Group.Permission)</td>" +
				"</tr>"
		}

		$temp += "</table></p>" +
			"</section>"
		$Sections += $temp

		$SharePointLinks += "<a href=`"#SharePointGroups`">SharePoint Online Groups</a>"
	}
	#endRegion
	
	#Full HTML
	$HTML = "<!DOCTYPE html>
	<html xmlns=`"http://www.w3.org/1999/xhtml`">
	<head>
	<style>
	* {
		box-sizing: border-box;
	}
	
	body {
		margin: 0;
	}
	
	.navbar {
		overflow: hidden;
		background-color: #333;
		font-family: Arial, Helvetica, sans-serif;
		position: fixed; /* Set the navbar to fixed position */
		top: 0; /* Position the navbar at the top of the page */
		width: 100%; /* Full width */
	}
	
	.navbar a {
		float: left;
		font-size: 16px;
		color: white;
		text-align: center;
		padding: 14px 16px;
		text-decoration: none;
	}
	
	.navbar .collapse {
	  float: right;
	}
	
	.dropdown {
		float: left;
		overflow: hidden;
		width: 100%; /* Full width */
	}
	
	.dropdown .dropbtn {
		font-size: 16px;    
		border: none;
		outline: none;
		color: grey;
		padding: 14px 16px;
		background-color: inherit;
		font: inherit;
		margin: 0;
	}
	
	.navbar a:hover, .dropdown:hover .dropbtn {
		background-color: red;
	}
	
	.dropdown-content {
		display: none;
		/* position: absolute; */
		background-color: #f9f9f9;
		width: 100%;
		left: 0;
		box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.2);
		z-index: 1;
	}
	
	.dropdown-content .header {
		background: red;
		padding: 16px;
		color: white;
	}
	
	.dropdown:hover .dropdown-content {
		display: block;
	}
	
	/* Create three equal columns that floats next to each other */
	.column {
		float: left;
		width: 33%;
		padding: 10px;
		background-color: #ccc;
		min-height: 250px;
	}
	
	.column a {
		float: none;
		color: black;
		padding: 16px;
		text-decoration: none;
		display: block;
		text-align: left;
	}
	
	.column a:hover {
		background-color: #ddd;
	}
	
	/* Clear floats after the columns */
	.row:after {
		content: `"`";
		display: table;
		clear: both;
	}
	
	.split {
	  margin-top: 45px;
	}
	
	.collapsible {
		background-color: #777;
		color: white;
		cursor: pointer;
		padding: 18px;
		width: 100%;
		border: none;
		text-align: left;
		outline: none;
		font-size: 15px;
	}
	
	.active, .collapsible:hover {
		background-color: #555;
	}
	
	.collapsible:after {
		content: `"\002B`";
		color: white;
		font-weight: bold;
		float: right;
	}
	
	.active:after {
		content: `"\2212`";
	}
	
	.content {
		padding: 0 18px;
		max-height: 0;
		overflow: hidden;
		transition: max-height 0.2s ease-out;
		background-color: #f1f1f1;
	}

	table {
		font-family: arial, sans-serif;
		border-collapse: collapse;
		width: 100%;
		display: block;
        overflow-x: auto;
        white-space: nowrap;
	}
	
	td, th {
		border: 1px solid #dddddd;
		text-align: left;
		padding: 8px;
	}
	
	tr:nth-child(even) {
		background-color: #dddddd;
	}
	
	/* Responsive layout - makes the three columns stack on top of each other instead of next to each other */
	@media screen and (max-width: 600px) {
		.column {
			width: 100%;
			height: auto;
		}
	}
	</style>
	</head>
	<body>
	<nav class=`"navbar`">
		<div class=`"dropdown`">
			<button class=`"dropbtn`">Menu</button>
			<div class=`"dropdown-content`">
				<div class=`"row`">
					<div class=`"column`">
					<h3>Azure Active Directory</h3>"

					foreach($AADLink in $AADLinks){
						$HTML += $AADLink
					}
					
					$HTML += "</div>
					<div class=`"column`">
					<h3>Exchange Online</h3>"

					foreach($ExchangeLink in $ExchangeLinks){
						$HTML += $ExchangeLink
					}
					
					$HTML += "</div>
					<div class=`"column`">
					<h3>SharePoint Online</h3>"

					foreach($SharePointLink in $SharePointLinks){
						$HTML += $SharePointLink
					}
					
					$HTML += "</div>
				</div>
			</div>
		</div> 
	</nav> 
	
	<div class=split></div>"
	

	foreach($section in $sections){
		$HTML += $section
	}
	
	$HTML += "<footer></footer>
	
	<script>
	var coll = document.getElementsByClassName(`"collapsible`");
	var i;
	
	for (i = 0; i < coll.length; i++) {
		coll[i].addEventListener(`"click`", function() {
		this.classList.toggle(`"active`");
		var content = this.nextElementSibling;
		if (content.style.maxHeight){
			content.style.maxHeight = null;
		} else {
			content.style.maxHeight = content.scrollHeight + `"px`";
		} 
		});
	}
	</script>
	
	</body>
	</html>"

	add-content $HTML -path $HTMLPath

	Invoke-Item $HTMLPath

	$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
		update-MessageBlock -type "Okay" -Message "Generated HTML file"
	})
}
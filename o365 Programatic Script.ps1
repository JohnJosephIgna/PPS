#Custom Powershell Script created by yours truly
#John Joseph Igna
#This needs to be Ran as Admin or on ISE running in Admin Mode

cls
clear-host

write-host("")
write-host("o365 Programatic tool by John Joseph Igna") -F Green
write-host("v 0.0.1") -F Green
write-host("")
write-host("Please select from one of the following choices") -F Green
write-host("")
write-host("General Operation")
write-host("1 - Just connect to o365") -F Yellow
write-host("")
write-host("Mailbox Operations")
write-host("2 - Gather DL Memberships") -F Yellow
write-host("3 - Gather Forwarding Report") -F Yellow
write-host("4 - Get User Mailbox Access") -F Yellow
write-host("")
write-host("Calendar Operations")
write-host("5 - Get User Calendar Access") -F Yellow
write-host("6 - Add User Calendar Access") -F Yellow
write-host("")
write-host("0 - Exit") -F Red
write-host("")

$MainChoice = read-host "Enter Choice "

if ($MainChoice -eq 1) {

    Get-PSSession | Remove-PSSession
    $Office365Credentials  = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365credentials -Authentication Basic –AllowRedirection
    Import-PSSession $Session -AllowClobber | Out-Null

} elseif ($MainChoice -eq 2) {

    Get-PSSession | Remove-PSSession

    #User input for filename
    $CSVFilename = read-host "Enter File Name "
    $OutputFile = $CSVFilename + '.csv'

    $Office365Credentials  = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365credentials -Authentication Basic –AllowRedirection
    Import-PSSession $Session -AllowClobber | Out-Null

    $arrDLMembers = @{}
    Out-File -FilePath $OutputFile -InputObject "Distribution Group DisplayName,Distribution Group Email,Member DisplayName, Member Email, Member Type" -Encoding UTF8
    $objDistributionGroups = Get-DistributionGroup -ResultSize Unlimited

    Foreach ($objDistributionGroup in $objDistributionGroups)  
    {      
     
    write-host "Processing $($objDistributionGroup.DisplayName)..."  
  
    #Get members of this group  
    $objDGMembers = Get-DistributionGroupMember -Identity $($objDistributionGroup.PrimarySmtpAddress)  
         
    write-host "Found $($objDGMembers.Count) members..."  
      
    #Iterate through each member  
    Foreach ($objMember in $objDGMembers) {  
        Out-File -FilePath $OutputFile -InputObject "$($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)" -Encoding UTF8 -append   
    }  
    }

} elseif ($MainChoice -eq 3) {
        
    Get-PSSession | Remove-PSSession
    $Office365Credentials  = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365credentials -Authentication Basic –AllowRedirection
    Import-PSSession $Session -AllowClobber | Out-Null

    Get-Mailbox | Where {$_.ForwardingAddress -ne $null} | Select Name, PrimarySMTPAddress, ForwardingAddress, DeliverToMailboxAndForward
    Get-Mailbox | Where {$_.ForwardingSMTPAddress -ne $null} | Select Name, PrimarySMTPAddress, ForwardingAddress, DeliverToMailboxAndForward

} elseif ($MainChoice -eq 4) {

    Get-PSSession | Remove-PSSession
    $Office365Credentials  = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365credentials -Authentication Basic –AllowRedirection
    Import-PSSession $Session -AllowClobber | Out-Null

    write-host("")
    $Targetusermbx = read-host "Enter user Alias "

    get-mailbox | get-mailboxpermission -User "$Targetusermbx"

} elseif ($MainChoice -eq 5) {
        
    write-host("Search for User First using any part of the name") -F Cyan
    $UserSearch = read-host "Enter First Name "
    get-mailbox *$UserSearch* | Select Name

    write-host("")
    write-host("Enter the Full Name of the Target Calendar Below")
    $UserTarget = read-host "Enter Target Name "
    ForEach ($mbx in Get-Mailbox) {Get-MailboxFolderPermission ($mbx.Name + ":\Calendar") | Where-Object {$_.User -like "$UserTarget"} | Select @{Name="Calendar Of";expression={($mbx.name)}},User,AccessRights}

} elseif ($MainChoice -eq 6) {
        
    write-host("")
    write-host("Search for User First using any part of the name") -F Cyan
    $UserSearch1 = read-host "Enter Search Criteria "
    get-mailbox *$UserSearch* | Select Name,Alias

    write-host("")
    write-host("Enter the alias of the Target Calendar Below")
    $UserTarget = read-host "Enter Target Alias "
    write-host("")
    write-host("Search for User First that needs access using any part of the name") -F Cyan
    $UserSearch2 = read-host "Enter Search Criteria "
    get-mailbox *$UserSearch2* | Select Name,Alias

    write-host("")
    write-host("Enter the alias of the Target User Below")
    $UsertoAdd = read-host "Enter Target Alias "

    write-host("")
    write-host("Select Permission to Apply")
    write-host("1 - Reviewer (Read)") -F Yellow
    write-host("2 - Author (Read/Write Own/Del Own") -F Yellow
    write-host("3 - Editor (Read/Write)") -F Yellow
    write-host("4 - Publishing Editor (Full)") -F Yellow
    write-host("")

    $PermissionCoice = read-host "Select permission "

    if ($PermissionCoice -eq 1) {
            
        $Perm = "Reviewer"

    } elseif ($PermissionCoice -eq 2) {

        $Perm = "Author"

    } elseif ($PermissionCoice -eq 3) {

        $Perm = "Editor"

    } elseif ($PermissionCoice -eq 4) {

        $Perm = "PublishingEditor"

    }

        Set-MailboxFolderPermission "'$UserTarget':\Calendar" -User "$UsertoAdd" -AccessRights "$Perm"

    }

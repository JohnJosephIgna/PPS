#Custom Powershell Script created by yours truly
#John Joseph Igna
#This needs to be Ran as Admin or on ISE running in Admin Mode

cls
clear-host

#Check for Prerequisites First
#MSOnline
if (Test-path 'C:\Program Files\WindowsPowerShell\Modules\MSOnline') {
    Write-Host ("SUCCESS: MSOnline Module Found") -F Green
    $MSOlcheck = 1
} else {
    Write-Host ("Error: MSOnline Module not installed, some scripts might fail to run") -F Red
    $MSOlcheck = 0
}

#Chocolatey
if (Test-path 'C:\ProgramData\chocolatey') {
    Write-Host ("SUCCESS: Chocolatey Installed") -F Green
    $Chococheck = 1
} else {
    Write-Host ("Error: Chocolatey not installed, Some chocolatey commands might not run") -F Red
    $Chococheck = 0
}

#Main Menu
#=========
write-host("")
write-host("Welcome, please select from one of the following choices") -F Green
write-host("")
write-host("1 - AD Operations") -F Yellow
write-host("2 - o365 Operations") -F Yellow
write-host("3 - Chocolatey (SW Installation Tool)") -F Yellow
write-host("")
write-host("0 - Exit") -F Red
write-host("")

$MainChoice = read-host "Enter Choice "

#AD Menu and Operations After this Line
#======================================
if ($MainChoice -eq 1) {
    
    cls
    write-host("AD OPERATIONS") -F Green
    write-host("")
    write-host("Note that thse commands should be ran on the On-Premise server where AD module is Located") -F Cyan
    write-host("Usually its the Primary DC") -F Cyan
    write-host("")
    write-host("1 - Get AD User Status") -F Yellow
    write-host("2 - Get User Membership") -F Yellow
    write-host("")

    $SubChoice1 = read-host "Enter Choice "
    write-host("")

    if ($SubChoice1 -eq 1) {

        $CSVFilename = read-host "Enter File Name "
        write-host("")

        if ($CSVFilename -eq "") {
            $CSVFilename = "ADUserReport"
        }

        $OutputFile = $CSVFilename + '.csv'
        $Outputpath = 'C:\Users\' + $env:UserName + '\Desktop\' + $Outputfile

        write-host ("")
        $Exporttype = Read-host "Export Data (S)tatus / (L)ogon / (F)ull "
        write-host("")
        
        if ($Exporttype -eq $NULL) {
            $Exporttype = "F"
        }

        if ($Exporttype -eq "S" -OR $Exporttype -eq "s") {
    
            get-aduser -Filter * -Properties * | select Name,Enabled
            write-host("Successfully Exported to Desktop") -F Green

        } elseif ($Exporttype -eq "L" -OR $Exporttype -eq "l") {

            get-aduser -Filter * -Properties * | select Name,@{N='LastLogonDate';E={[DateTime]::FromFileTime($_.LastLogon).ToString('dd/MM/yyyy')}} | export-csv -Path "$Outputpath"
            write-host("Successfully Exported to Desktop") -F Green

        } elseif ($Exporttype -eq "F" -OR $Exporttype -eq "f") {

            get-aduser -Filter * -Properties * | select Name,Enabled,@{N='LastLogonDate';E={[DateTime]::FromFileTime($_.LastLogon).ToString('dd/MM/yyyy')}} | export-csv -Path "$Outputpath"
            write-host("Successfully Exported to Desktop") -F Green

        }

    } elseif ($SubChoice1 -eq 2) {

        write-host("Search for User First") -F Cyan
        $UserSearch1 = read-host "Enter First Name "
        write-host("")
        get-aduser -Filter "name -like '*$UserSearch1*'" | ft Name,samaccountname

        $AliasSearch = read-host "Enter Target Alias (SAM Account Name) "
        write-host("")
        Get-ADPrincipalGroupMembership $AliasSearch | sort name | select name

    }

#o365 Menu and Operations After this Line
#========================================
} elseif ($MainChoice -eq 2) {

    #Check if New Instance or Already logged in
    cls
    write-host("o365 Operations") -F Green
    write-host("1 - Setup New Instance") -F Green
    write-host("2 - Already Connected") -F Green
    write-host("")

    $Setconnection = read-host "Enter Choice "
    write-host("")

    if ($Setconnection -eq 1) {
        
        Get-PSSession | Remove-PSSession
        $Office365Credentials  = Get-Credential
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365credentials -Authentication Basic –AllowRedirection
        Import-PSSession $Session -AllowClobber | Out-Null

        cls
        write-host ("New Connection Established") -F Green
        write-host ("")

    } elseif ($Setconnection -eq 2) {
        
        cls
        write-host ("Using Existing connection") -F Green
        write-host ("")

    }

    write-host("o365 Operations") -F Green
    write-host("")
    write-host("1 - o365 Mailbox Report        | 5 - Mailbox Statistics Report") -F Yellow
    write-host("2 - o365 Mailbox Sign-In Satus | 6 - Get User Mailbox Access") -F Yellow
    write-host("3 - Gather DL Memberships      | 7 - Get User Calendar Access") -F Yellow
    write-host("4 - Gather Forwarding Report   | 8 - Add User Calendar Access") -F Yellow
    write-host("")
    write-host("0 - Custom Commands (Advanced Users)") -F Yellow
    write-host("")

    $SubChoice2 = read-host "Enter Choice "
    write-host("")

    if ($SubChoice2 -eq 1) {

        $CSVFilename = read-host "Enter File Name "
        write-host("")
        $OutputFile = $CSVFilename + '.csv'
        $OutputPath = 'C:\Users\' + $env:UserName + '\Desktop\' + $Outputfile

        Get-Mailbox -ResultSize Unlimited | Select Identity, UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails | Export-Csv -Path "$Outputpath"
        write-host("Successfully Exported to Desktop") -F Green

    } elseif ($SubChoice2 -eq 2) {
        
        if ($MSOlcheck -eq 0) {
            write-host ("ERROR - MSOnline Module not installed, please install it first") -F Red
            exit
        }

        connect-msolservice

        $CSVFilename = read-host "Enter File Name "
        write-host("")
        $OutputFile = $CSVFilename + '.csv'
        $OutputPath = 'C:\Users\' + $env:UserName + '\Desktop\' + $Outputfile
        get-msoluser | Select Userprincipalname, IsLicensed, BlockCredential | Export-csv "$OutputPath"
        write-host("Successfully Exported to Desktop") -F Green
    
    } elseif ($SubChoice2 -eq 3) {

        #User input for filename
        $CSVFilename = read-host "Enter File Name "
        write-host("")
        $OutputFile = $CSVFilename + '.csv'

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
            Foreach ($objMember in $objDGMembers)  
            {  
                Out-File -FilePath $OutputFile -InputObject "$($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)" -Encoding UTF8 -append   
            }  
        }

    } elseif ($SubChoice2 -eq 4) {

        Get-Mailbox | Where {$_.ForwardingAddress -ne $null} | Select Name, PrimarySMTPAddress, ForwardingAddress, DeliverToMailboxAndForward
        write-host("")
        Get-Mailbox | Where {$_.ForwardingSMTPAddress -ne $null} | Select Name, PrimarySMTPAddress, ForwardingAddress, DeliverToMailboxAndForward

    } elseif ($SubChoice2 -eq 5) {

        write-host("Search for User First") -F Cyan
        $UserSearch = read-host "Enter First Name "
        write-host("")
        get-mailbox *$UserSearch*

        $UserTarget = read-host "Enter Target Alias "
        write-host("")

        $CSVFilename = read-host "Enter File Name "
        write-host("")
        $OutputFile = $CSVFilename + '.csv'
        $OutputPath = 'C:\Users\' + $env:UserName + '\Desktop\' + $Outputfile

        Get-MailboxFolderStatistics "$UserTarget" | select Name, FolderPath, ItemsInFolderAndSubfolders, FolderSize | export-csv -Path "$OutputPath"
        write-host("Successfully Exported to Desktop") -F Green

    } elseif ($SubChoice2 -eq 6) {

        $Targetusermbx = read-host "Enter user Alias "
        write-host("")

        get-mailbox | get-mailboxpermission -User "$Targetusermbx"

    } elseif ($SubChoice2 -eq 7) {
        
        write-host("Search for User First") -F Cyan
        $UserSearch = read-host "Enter First Name "
        write-host("")
        get-mailbox *$UserSearch*

        $UserTarget = read-host "Enter Target Name "
        write-host("")
        ForEach ($mbx in Get-Mailbox) {Get-MailboxFolderPermission ($mbx.Name + ":\Calendar") | Where-Object {$_.User -like "$UserTarget"} | Select @{Name="Calendar Of";expression={($mbx.name)}},User,AccessRights}

    } elseif ($SubChoice2 -eq 8) {

        write-host("Search for User First") -F Cyan
        $UserSearch1 = read-host "Enter First Name "
        write-host("")
        get-mailbox *$UserSearch*

        $UserTarget = read-host "Enter Target Alias "
        write-host("")
        write-host("Search for User To Get Access") -F Cyan
        $UserSearch2 = read-host "Enter First Name User "
        write-host("")
        get-mailbox *$UserSearch*

        $UsertoAdd = read-host "Enter Target Alias "
        write-host("")

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

    } elseif ($SubChoice2 -eq 0) {

        write-host ("Custom Commands Selected, you may start putting commands below..") -F Green
        write-host ("")
        exit

    }

#Chocolatey Menu and Operations After this Line
#==============================================
} elseif ($MainChoice -eq 3) {

    write-host("")
    write-host("1 - Install Chocolatey") -F Yellow
    write-host("2 - Upgrade Chocolatey") -F Yellow
    write-host("")

    $SubChoice3 = read-host "Enter Choice "

    if ($SubChoice3 -eq 1) {

        Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

    } elseif ($SubChoice3 -eq 2) {

        choco upgrade chocolatey -Y

    }

} elseif ($MainChoice -eq 0) {
    exit
}


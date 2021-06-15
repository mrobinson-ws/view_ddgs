#Requires -Module ExchangeManagementOnline
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()
$quitboxOutput = ""


#Test And Connect To Microsoft Exchange Online If Needed
try {
    Write-Verbose -Message "Testing connection to Microsoft Exchange Online"
    Get-Mailbox -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Microsoft Exchange Online"
}
catch {
    Write-Verbose -Message "Connecting to Microsoft Exchange Online"
    Connect-ExchangeOnline
}

while($quitboxOutput -ne "NO")
    $DDG = Get-DynamicDistributionGroup | Out-GridView -OutputMode Single
    $GroupMembers = Get-Recipient -RecipientPreviewFilter $DDG.RecipientFilter -OrganizationalUnit $DDG.RecipientContainer
    Write-Host "Members Of $($DDG.Name)"
    foreach($GroupMember in $GroupMembers){
        "$($GroupMember.DisplayName)" + " -- $($GroupMember.PrimarySMTPAddress)"
        [pscustomobject]@{
                "User Name" = $GroupMember.Displayname
                Email = $GroupMember.PrimarySMTPAddress
            } | Export-Csv -Path c:\users\$env:USERNAME\Downloads\$(get-date -f yyyy-MM-dd)_members_of_$($DDG.Name).csv -NoTypeInformation -Append
    }
    Write-Host "Outputting to CSV as well, It will be in your Downloads Directory"
    $quitboxOutput = [System.Windows.Forms.MessageBox]::Show("Do you need to check another Dynamic Distribution Group?" , "Group Membership Export Complete" , 4)
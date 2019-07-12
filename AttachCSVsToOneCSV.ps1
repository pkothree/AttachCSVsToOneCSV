###
# Danny Davis
# twitter: twitter.com/pko3
# github: github.com/pkothree
# Created: 07/10/19
# Modified: 07/12/19
# Description: Download files from SharePoint, update a csv und upload the file again
# The file which you want to append to needs to exist and also needs to have the same headers as the csv
# the final csv will be send via e-Mail
###

Import-Module SharePointPnPPowerShellOnline

$cred = Get-Credential
$url = "https://TENANT.sharepoint.com/sites/SITECOLLECTION/"
Connect-PnPOnline $url -Credential $cred
$statFile = "csv" # Path to statistics file
$statLocalPath = "c:\temp\" # Local path where the statitics file is stored
$localStatFile = "final.csv" # Name of file which will contain the collection of data
$attachmentPath = $statLocalPath + $localStatFile
$items = [System.Collections.ArrayList]@()

$listItems = Get-PnPListItem -List $statFile

foreach($listItem in $listItems)
{
    if($listItem.FieldValues.FileLeafRef -ne "final.csv")
    {
        $items.Add($listItem.FieldValues.FileLeafRef)
        Get-PnPFile -Url $listItem.FieldValues.FileRef -Path $statLocalPath -AsFile -Force
    }
}

# Go through item in the list and append the content to the $localStatFile
foreach($item in $items)
{
    # Get filename and attach to path
    $file = $statLocalPath + $item 
    # Get the file for the data collection
    $stat = $statLocalPath + $localStatFile
    $import = Import-Csv -Path $file
    # Store the data in the collection file
    $import | Export-Csv -Path $stat -NoTypeInformation -Append
}

# Send the file as an attachment
Send-MailMessage -From 'User1 <user1@domain.com>' -To 'User2 <user2@domain.com>' -Subject 'Sending the Attachment' -Body "Here's the CSV." -Attachments $attachmentPath -Priority Normal -SmtpServer 'smtp.office365.com' -Port "587" -UseSSL -Credential $cred
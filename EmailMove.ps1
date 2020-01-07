$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)

# Mailbox name
$mailbox = "lli@oxfordproperties.com"          #############
# Name of folder containing reported phishing emails
$folderName = "inbox"
# Create the Outlook object to access mailboxes
$Outlook = New-Object -ComObject Outlook.Application;
# Grab the folder containing the emails from the phishing exercise
$Folder = $Outlook.Session.Folders($mailbox).Folders.Item($folderName)
#access the subfolder from inbox
$targetfolder = $inbox.Folders | where-object { $_.name -eq "SDReport" }    #############

#define the destination subfolder
$destinationfolder=$inbox.Folders | where-object{ $_.name -eq "MoveToHere"}        #############

# Grab the emails in the folder
$Emails = $targetfolder.Items


$emailcount = $emails.Count
for($t=$emailcount-1; $t -ge 0; $t--){
    
    $i=$emails.getlast()
    #write-host "this is $t email $i" -ForegroundColor yellow
    
    $i.move($destinationfolder)} 
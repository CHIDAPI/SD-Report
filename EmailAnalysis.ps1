

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
# Path to save attachments to
$filepath = "C:\temp\JJ\"        ###############

foreach($email in $Emails){
    # Output Sender Name for Testing
    #Write-Host $email.SenderName
    # The number of email attachments
    $intCount = $email.Attachments.Count

    # If the email has attachments, let's open the .msg email
    if($intCount -gt 0) {
        # Let's go through those attachments
        for($i=1; $i -le $intCount; $i++) {

            # The attachment being looked at
            $attachment = $email.Attachments.Item($i)

            # subject
            $subject = $email.Subject
            #write-host "email subject is $subject"
            if($attachment.FileName -like "*.zip"){
                $attachmentPath = $filepath+$attachment.FileName
                $attachmentName=$attachment.FileName
                $attachment.SaveAsFile($attachmentPath)
                Expand-Archive -path $attachmentPath -destinationpath $filepath}
                }
                }
                }

    




Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook
write-host $xlFixedFormat
$excel = New-Object -ComObject excel.application
$excel.visible = $false
$folderpath = "C:\temp\JJ"    ##############
$filetype ="*xls"
Get-ChildItem -Path $folderpath -Include $filetype -recurse | 
ForEach-Object
{
	$path = ($_.fullname).substring(0, ($_.FullName).lastindexOf("."))
	
	"Converting $path"
	$workbook = $excel.workbooks.open($_.fullname)

	$path += ".xlsx"
	$workbook.saveas($path, $xlFixedFormat)
	$workbook.close()
	

	#$oldFolder = $path.substring(0, $path.lastIndexOf("\")) + "\old"
	
	#write-host $oldFolder
	#if(-not (test-path $oldFolder))
	#{
	#	new-item $oldFolder -type directory
	#}
	
	#move-item $_.fullname $oldFolder
	
}
$excel.Quit()
$excel = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()


python "C:\Users\l_li\Desktop\PowerShell\SD Cisco Report _Jamie\Excel.py"      ############
sleep 2
python "C:\Users\l_li\Desktop\PowerShell\SD Cisco Report _Jamie\Excel2.py"     ############
sleep 2
python "C:\Users\l_li\Desktop\PowerShell\SD Cisco Report _Jamie\ExcelCSQ.py"
sleep 2


$emailcount = $emails.Count
for($t=$emailcount-1; $t -ge 0; $t--){
    
    $i=$emails.getlast()
    #write-host "this is $t email $i" -ForegroundColor yellow
    
    $i.move($destinationfolder)} 

 
$Today = (get-date).tostring("yyyy-MM-dd")

$Items = Get-ChildItem c:\temp\JJ
foreach ($Item in $Items)
{
    if ($item.name -match "^Oxford*")
    {
        #move-item c:\temp\JJ\$item C:\temp\JJ\Archive\[$Today]$item 
        remove-item c:\temp\JJ\$item
    }
}
# Path to directories
$sourceDir = "C:\UserManagement\UserDeletion"
$workingDir = "C:\UserManagement\Working"
$errorDir = "C:\UserManagement\Error"
$reportDir = "C:\UserManagement\Report"
$doneDir = "C:\UserManagement\Done"
$emailRecipients = ""  # Update with your team's email address

# Get the AD domain name
$adDomainName = (Get-ADDomain).Name
$adDomainNetBIOSName = (Get-ADDomain).NetBIOSName

# Get all csv files in the source folder
$csvFiles = Get-ChildItem -Path $sourceDir -Filter *.csv

foreach($file in $csvFiles){
    # Move file to working directory
    Move-Item -Path $file.FullName -Destination $workingDir

    # Construct new file path
    $newFilePath = Join-Path -Path $workingDir -ChildPath $file.Name

    # Try to parse the CSV file, if parsing fails, the file is not properly formatted
    try {
        $users = Import-Csv -Path $newFilePath
    }
    catch {
        Write-Output "File $($file.Name) is not properly formatted."
        # Move the file to the error folder
        Move-Item -Path $newFilePath -Destination $errorDir
        continue
    }

    # Initialize report
    $reportContent = @()

    foreach($user in $users){
        if($user.Username){
            # Check if the user exists
            $ADUser = Get-ADUser -Filter "SamAccountName -eq '$($user.Username)'" -Properties Description
            if($ADUser) {
                # Check if the Description contains 'CUSTOMER'
                if($ADUser.Description -like '*CUSTOMER*') {
                    try{
                        # Delete the user
                        Remove-ADUser -Identity $user.Username -Confirm:$false
                        # Log to the report
                        $reportContent += "$($user.Username),Deletion Successful"
                    }
                    catch {
                        # Log to the report
                        $reportContent += "$($user.Username),Deletion Failed"
                    }
                } else {
                    # Log to the report
                    $reportContent += "$($user.Username),DELETION NOT ALLOWED"
                }
            } else {
                # Log to the report
                $reportContent += "$($user.Username),User Not Found"
            }
        }
        else{
            Write-Output "Information missing in file $($file.Name)."
            # Move the file to the error folder
            Move-Item -Path $newFilePath -Destination $errorDir
            continue
        }
    }

    # Finalize the database changes
    echo . | &  "C:\Program Files (x86)\eCTDmanager\EditDB.exe" C:\eCTDmanager-Data\NamedUsers\NamedUsers.sqlite --finalize

    # Create a report
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $reportFilePath = "$reportDir\userdeletion_report_$timestamp.csv"
    $reportContent | Out-File -FilePath $reportFilePath

    # Move the file to the done folder
    Move-Item -Path $newFilePath -Destination $doneDir

    # Send email with the report
    try {
        $smtpServer = ""  # Update with your SMTP server
        $smtpUsername = ""  # Update with your SMTP username
        $smtpPassword = ""  # Update with your SMTP password
        $emailSubject = "User Deletion Report for $($adDomainNetBIOSName)"
        $emailBody = "The user deletion process has completed. Please find the report attached."

        $attachment = New-Object Net.Mail.Attachment($reportFilePath)
        $message = New-Object Net.Mail.MailMessage
        $message.From = ""  # Update with the appropriate email address
        $message.To.Add($emailRecipients)
        $message.Subject = $emailSubject
        $message.Body = $emailBody
        $message.Attachments.Add($attachment)

        $smtp = New-Object Net.Mail.SmtpClient($smtpServer, 587)
        $smtp.EnableSsl = $true
        $smtp.Credentials = New-Object System.Net.NetworkCredential($smtpUsername, $smtpPassword)
        $smtp.Send($message)

        Write-Output "Report emailed successfully."
    }
    catch {
        Write-Output "Error sending email: $_"
        # Move report file to error directory
        Move-Item -Path $reportFilePath -Destination $errorDir
    }
}

Write-Output "User deletion process completed."


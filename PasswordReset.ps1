# Paths to directories
$sourceFile = "C:\UserManagement\PasswordReset\PasswordReset.csv"
$workingDir = "C:\UserManagement\Working"
$errorDir = "C:\UserManagement\Error"
$reportDir = "C:\UserManagement\Report"
$doneDir = "C:\UserManagement\Done"
$emailRecipients = "johndoe@test.com"  # Update with your team's email address

# Get the AD domain name
$adDomainName = (Get-ADDomain).Name
$adDomainNetBIOSName = (Get-ADDomain).NetBIOSName

# Check if the source file exists
if (!(Test-Path -Path $sourceFile)) {
    Write-Output "No CSV file found at $sourceFile"
    exit
}

# Move the CSV file from source to working directory
Move-Item -Path $sourceFile -Destination $workingDir

# Process the CSV file in the working directory
Get-ChildItem -Path $workingDir -Filter "PasswordReset.csv" | ForEach-Object {
    $csvFile = $_.FullName
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $reportFile = Join-Path -Path $reportDir -ChildPath ($_.BaseName + "_report_" + $timestamp + ".csv")

    # Try to import the CSV file
    try {
        $users = Import-Csv -Path $csvFile -ErrorAction Stop
    }
    catch {
        Write-Output "Error reading file $csvFile. Message: $_"
        # Move the erroneous CSV file to the error directory
        Move-Item -Path $csvFile -Destination $errorDir
        continue
    }

    $reportData = @()

    # Function to generate random passwords
    function New-RandomPassword {
        Add-Type -AssemblyName 'System.Web'
        $length = 16
        $NumberOfAlphaNumericCharacters = 5
        $password = [System.Web.Security.Membership]::GeneratePassword($length,$NumberOfAlphaNumericCharacters)
        $password
    }

    # Process each user in the CSV file
    $users | ForEach-Object {
        $username = $_.Username

        # Check if the user exists
        $userExist = Get-ADUser -Filter "SamAccountName -eq '$username'"

        if ($userExist) {
            $password = New-RandomPassword
            $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force

            # Reset the user's password
            try {
                Set-ADAccountPassword -Identity $username -NewPassword $securePassword -Reset -ErrorAction Stop
                $reportData += [pscustomobject]@{
                    UserName = $username
                    Password = $password
                    Status = "Password Reset"
                }
            }
            catch {
                Write-Output "Error resetting password for user $username. Message: $_"
                $reportData += [pscustomobject]@{
                    UserName = $username
                    Password = "NULL"
                    Status = "Error: $_"
                }
            }
        }
        else {
            Write-Output "User $username does not exist."
            $reportData += [pscustomobject]@{
                UserName = $username
                Password = "NULL"
                Status = "User does not exist"
            }
            return
        }
    }

    # Export report data to CSV file
    $reportData | Export-Csv -Path $reportFile -NoTypeInformation

    # Move the CSV file to the done directory
    Move-Item -Path $csvFile -Destination $doneDir

    # Send email with the report
    $smtpServer = ""  # Update with your SMTP server
    $smtpUsername = ""  # Update with your SMTP username
    $smtpPassword = ""  # Update with your SMTP password
    $emailSubject = "Password Reset Report for $($adDomainNetBIOSName)"
    $emailBody = "The password reset process has completed. Please find the report attached."

    $attachment = New-Object Net.Mail.Attachment($reportFile)
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
}

Write-Output "Password reset process completed."

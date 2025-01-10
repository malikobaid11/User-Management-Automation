# Path to directories
$sourceDir = "C:\UserManagement\UserCreation"
$workingDir = "C:\UserManagement\Working"
$errorDir = "C:\UserManagement\Error"
$reportDir = "C:\UserManagement\Report"
$doneDir = "C:\UserManagement\Done"
$emailRecipients = ""  # Update with your team's email address

# Get the AD domain name
$adDomainName = (Get-ADDomain).Name
$adDomainNetBIOSName = (Get-ADDomain).NetBIOSName

# Check if the source directory contains any CSV file
if (!(Get-ChildItem -Path $sourceDir -Filter "*.csv")) {
    Write-Output "No CSV files found in the source directory"
    exit
}

# Get CSV files from source directory and move to working directory
Get-ChildItem -Path $sourceDir -Filter "*.csv" | ForEach-Object {
    Move-Item -Path $_.FullName -Destination $workingDir
}

# Process each CSV file in the working directory
Get-ChildItem -Path $workingDir -Filter "*.csv" | ForEach-Object {
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

    # Process each user in the CSV file
    $users | ForEach-Object {
        $UserEmail = $_.Email
        $userExist = Get-ADUser -Filter  "UserPrincipalName -eq '$UserEmail'"

        if ($userExist) {
            $reportData += [pscustomobject]@{
                UserName = $_.UserName
                Password = "NULL"
                Status = "Already exists"
                LicenseType = "N/A"
            }
            return
        }

        # Check if any field is missing
        if (![string]::IsNullOrEmpty($_.FirstName) -and ![string]::IsNullOrEmpty($_.LastName) -and ![string]::IsNullOrEmpty($_.Username) -and ![string]::IsNullOrEmpty($_.Email) -and ![string]::IsNullOrEmpty($_.LicenseType)) {
            # Get the email domain from the user email
            $emailDomain = $_.Email.Split('@')[1]

            # Check if email domain matches the system's domain
            if ($emailDomain -ne $adDomainName) {
                Write-Output "Invalid active directory domain for user $($_.UserName)."
                $reportData += [pscustomobject]@{
                    UserName = $_.UserName
                    Password = "NULL"
                    Status = "Invalid domain"
                    LicenseType = "N/A"
                }
                return
            }

            # Function to generate random passwords.
            function New-RandomPassword {
                Add-Type -AssemblyName 'System.Web'
                $length = 16
                $NumberOfAlphaNumericCharacters = 5
                $password = [System.Web.Security.Membership]::GeneratePassword($length,$NumberOfAlphaNumericCharacters)
                if ($ConvertToSecureString.IsPresent) {
                    ConvertTo-SecureString -String $password -AsPlainText -Force
                } else {
                    $password
                }
            }

            $password = New-RandomPassword
            $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force

            # Create the user in Active Directory
            try {
                $newUser = New-ADUser -Name $_.Username -GivenName $_.FirstName -Surname $_.LastName -UserPrincipalName $_.Email -Description "CUSTOMER" -AccountPassword $securePassword -Enabled $True -PassThru -ErrorAction Stop

                NetBIOSName\$($_.UserName) --codes=$licenseTypeCode --enable
                $reportData += [pscustomobject]@{
                    UserName = $_.UserName
                    Password = $password
                    Status = "Created"
                    LicenseType = $_.LicenseType
                }
            }
            catch {
                Write-Output "Error creating user $($_.UserName). Message: $_"
                $reportData += [pscustomobject]@{
                    UserName = $_.UserName
                    Password = $password
                    Status = "Error: $_"
                    LicenseType = $_.LicenseType
                }
            }
        }
        else {
            Write-Output "Missing information for user $($_.UserName)."
            $reportData += [pscustomobject]@{
                UserName = $_.UserName
                Password = "NULL"
                Status = "Missing information"
                LicenseType = "N/A"
            }
        }
    }

    echo . | &  "C:\Program Files (x86)\eCTDmanager\EditDB.exe"  C:\eCTDmanager-Data\NamedUsers\NamedUsers.sqlite --finalize

    # Export report data to CSV file
    $reportData | Export-Csv -Path $reportFile -NoTypeInformation

    # Move the CSV file to the done directory
    Move-Item -Path $csvFile -Destination $doneDir

    # Send email with the report
    $smtpServer = ""  # Update with your SMTP server
    $smtpUsername = ""  # Update with your SMTP username
    $smtpPassword = ""  # Update with your SMTP password
    $emailSubject = "User Creation Report for $($adDomainNetBIOSName)"
    $emailBody = "The user creation process has completed. Please find the report attached."

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

Write-Output "User creation process completed."

# User-Management-Automation
Automation for user management in Azure AD and Active Directory Using Powershell

Scripts
1. UserCreation.ps1

Purpose:
Automates the creation of new user accounts in Active Directory, assigning roles and generating reports for the process.

Workflow:

    Reads user information from CSV files in the source directory.
    Validates user data (e.g., email domain).
    Generates a secure random password for each user.
    Creates users in Active Directory.
    Logs success or errors to a report CSV file.
    Sends an email notification with the process report.

Key Parameters:

    $sourceDir: Directory containing input CSV files.
    $workingDir: Temporary working directory for processing files.
    $reportDir: Directory where reports are saved.
    $doneDir: Directory for processed CSV files.
    $errorDir: Directory for problematic CSV files.

Logs and Reports:

    Success and error logs are stored in the report directory.
    Email notifications include detailed reports.

Dependencies:

    Active Directory cmdlets.

2. UserDeletion.ps1

Purpose:
Streamlines the deletion of user accounts while maintaining compliance with business rules.

Workflow:

    Reads user information from CSV files in the source directory.
    Verifies that users exist and are eligible for deletion (e.g., marked as "CUSTOMER").
    Deletes users from Active Directory.
    Logs success or errors to a report CSV file.
    Sends an email notification with the process report.

Key Parameters:

    Same directory structure as UserCreation.ps1.

Logs and Reports:

    Stores a detailed report of the deletion process.
    Emails reports to specified recipients.

Dependencies:

    Active Directory cmdlets.

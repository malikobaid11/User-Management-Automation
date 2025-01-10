# Usage Guide for User Management Scripts

## Prerequisites
- PowerShell 5.1 or higher.
- Active Directory cmdlets installed and configured.
- Proper permissions for user creation/deletion.

## Directory Setup
Ensure the following directories exist and are accessible:
- `C:\UserManagement\UserCreation` for user creation files.
- `C:\UserManagement\UserDeletion` for user deletion files.
- `C:\UserManagement\PasswordReset` for password reset files.
- `C:\UserManagement\Working` for temporary files.
- `C:\UserManagement\Error` for problematic files.
- `C:\UserManagement\Report` for generated reports.
- `C:\UserManagement\Done` for processed files.

## Running the Scripts
1. **User Creation**
   ```powershell
   .\UserCreation.ps1

2. **User Creation**
   ```powershell
   .\UserDeletion.ps1

3. **User Creation**
   ```powershell
   .\PasswordReset.ps1

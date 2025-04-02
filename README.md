# GuestUserReport - Export Office 365 Guest User and Membership Report

## Overview
This PowerShell script exports a report of all guest users in an Office 365 tenant using the Microsoft Graph PowerShell module. It includes details like user display name, email, company, account age, group memberships, and invitation status.

## Features

- Connects to Microsoft Graph to retrieve guest user information.
- Supports filtering guests based on account age.
- Exports the generated report to a CSV file.
- Provides the option to open the report upon completion.

## Prerequisites

- PowerShell installed on your machine.
- Microsoft Graph PowerShell module must be installed. The script installs it automatically if not present.
- Appropriate permissions to access guest user information in Office 365.

## Parameters

- `-StaleGuests`: Filters the report to include only guests older than the specified number of days.
- `-RecentlyCreatedGuests`: Filters the report to include only guests created within the specified number of days.
- `-TenantId`: Specifies the Tenant ID for authentication.
- `-ClientId`: Specifies the Client ID for authentication.
- `-CertificateThumbprint`: Specifies the certificate thumbprint for authentication.

## Usage

1. Clone this repository:
   ```bash
   git clone https://github.com/YourGitHubUsername/ExportO365Guests.git
   cd ExportO365Guests
   ```

2. Open PowerShell and run the script:
   ```powershell
   .\ExportO365Guests.ps1
   ```

3. Use the available parameters as needed. For example:
   ```powershell
   # Example command to generate a report for guest users older than 30 days
   .\ExportO365Guests.ps1 -StaleGuests 30

   # Example command to generate a report for guest users created within the last 7 days
   .\ExportO365Guests.ps1 -RecentlyCreatedGuests 7
   ```

## Notes

- The generated report includes details such as display name, user principal name, email address, company, creation time, account age, invitation status, and group memberships.
- The output file is saved with a timestamp in the specified directory.
- The script disconnects from Microsoft Graph after execution to ensure security.

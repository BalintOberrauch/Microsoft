# Sync GAL to Personal Contacts in Outlook

This PowerShell script synchronizes contacts from the Global Address List (GAL) to your personal contacts in Outlook. If a contact already exists, the script updates the contact information. The script also handles cases where the GAL is not accessible by attempting to use the Offline Global Address List.

## Prerequisites

- Windows operating system
- Microsoft Outlook installed and configured
- PowerShell

## Use Case

This script is particularly useful when you don't have administrative access to Microsoft Graph or cannot access the Exchange server directly. By leveraging the Outlook COM object, the script can access the Global Address List (GAL) and synchronize it with your personal contacts in Outlook.

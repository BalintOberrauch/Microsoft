<#
.SYNOPSIS
Synchronizes contacts from the Global Address List (GAL) to personal contacts in Outlook.

.DESCRIPTION
This script loads the Outlook COM object, retrieves the Global Address List (GAL), and copies contacts from the GAL 
to the user's personal contacts in Outlook. If a contact already exists, it updates the contact information.
It also handles cases where the GAL is not accessible by attempting to use the Offline Global Address List.

.NOTES
Author: Balint Oberrauch
Date: 2024-07-08

.EXAMPLE
Run the script in PowerShell to sync contacts:
    .\Sync-GALToContacts.ps1

#>

# Load Outlook COM object
Write-Output "Loading Outlook COM object..."
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# List all available address lists
Write-Output "Listing all available address lists..."
$addressLists = $namespace.AddressLists
foreach ($addressList in $addressLists) {
    Write-Output "Found Address List: $($addressList.Name)"
}

# Get the Global Address List (GAL)
Write-Output "Retrieving the Global Address List (GAL)..."
$gal = $addressLists | Where-Object { $_.Name -eq "Global Address List" }

if ($null -eq $gal) {
    Write-Error "Global Address List not found."
    exit
} else {
    Write-Output "Global Address List found."
}

# Check if GAL has any entries
$galEntries = $gal.AddressEntries
if ($null -eq $galEntries -or $galEntries.Count -le 0) {
    Write-Error "Global Address List is empty or inaccessible. Attempting to access Offline Global Address List..."

    # Attempt to access Offline Global Address List
    $offlineGal = $addressLists | Where-Object { $_.Name -eq "Offline Global Address List" }

    if ($null -eq $offlineGal) {
        Write-Error "Offline Global Address List not found."
        exit
    } else {
        Write-Output "Offline Global Address List found."
        $galEntries = $offlineGal.AddressEntries
    }

    if ($null -eq $galEntries -or $galEntries.Count -le 0) {
        Write-Error "Offline Global Address List is also empty or inaccessible."
        exit
    } else {
        Write-Output "Offline Global Address List contains $($galEntries.Count) entries."
    }
} else {
    Write-Output "Global Address List contains $($galEntries.Count) entries."
}

# Get the Contacts folder
Write-Output "Retrieving the Contacts folder..."
$contactsFolder = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderContacts)

if ($null -eq $contactsFolder) {
    Write-Error "Contacts folder not found."
    exit
} else {
    Write-Output "Contacts folder found."
}

# Iterate through each entry in the GAL
foreach ($entry in $galEntries) {
    # Skip if the entry is not a user
    if ($entry.AddressEntryUserType -ne [Microsoft.Office.Interop.Outlook.OlAddressEntryUserType]::olExchangeUserAddressEntry) {
        Write-Output "Skipping non-user entry: $($entry.Name)"
        continue
    }
    
    $user = $entry.GetExchangeUser()
    if ($null -eq $user) {
        Write-Output "Skipping entry: Unable to retrieve user information for $($entry.Name)."
        continue
    }

    Write-Output "Processing user: $($user.Name) <$($user.PrimarySmtpAddress)>"

    # Check if contact already exists in personal contacts
    $existingContact = $contactsFolder.Items | Where-Object { $_.Email1Address -eq $user.PrimarySmtpAddress }
    
    if ($null -eq $existingContact) {
        Write-Output "Creating new contact for $($user.Name)..."
        # Create new contact
        $contact = $contactsFolder.Items.Add([Microsoft.Office.Interop.Outlook.OlItemType]::olContactItem)
        $contact.FirstName = $user.FirstName
        $contact.LastName = $user.LastName
        $contact.Email1Address = $user.PrimarySmtpAddress
        $contact.JobTitle = $user.JobTitle
        $contact.CompanyName = $user.CompanyName
        $contact.BusinessTelephoneNumber = $user.BusinessTelephoneNumber
        $contact.MobileTelephoneNumber = $user.MobileTelephoneNumber
        $contact.Save()
        Write-Output "Contact for $($user.Name) created."
    } else {
        Write-Output "Updating existing contact for $($user.Name)..."
        # Update existing contact
        $contact = $existingContact
        $contact.FirstName = $user.FirstName
        $contact.LastName = $user.LastName
        $contact.JobTitle = $user.JobTitle
        $contact.CompanyName = $user.CompanyName
        $contact.BusinessTelephoneNumber = $user.BusinessTelephoneNumber
        $contact.MobileTelephoneNumber = $user.MobileTelephoneNumber
        $contact.Save()
        Write-Output "Contact for $($user.Name) updated."
    }
}

Write-Output "Contacts sync completed."

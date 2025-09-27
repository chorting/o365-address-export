# Export O365 Mail Objects with SMTP Addresses

## Overview

This PowerShell script connects to Exchange Online and exports all mail-enabled objects (Mailboxes, Distribution Groups, Dynamic Distribution Groups, Unified Groups, and Mail Contacts) along with their SMTP addresses. It supports two modes:
- **Single SMTP Search**: Search for a specific SMTP address.
- **All SMTP Addresses**: Export all mail-enabled objects.

## Features

- Secret-based authentication to Exchange Online.
- GUI prompts for action selection and file destination.
- Memory-efficient Excel export using chunked writing.
- CSV fallback if Excel is not available.
- Error logging for Excel write failures.

## Prerequisites

- PowerShell 5.1 or later
- ExchangeOnlineManagement module

## Optional Prerequisites

- Registered Azure AD App with `User.Read` Graph API permission and secret for authentication
- Excel installed (optional for Excel output)

## Usage

1. Open PowerShell and run the script.
2. Choose between:
   - Single SMTP Search
   - All SMTP Addresses
3. Provide the SMTP address or choose a destination file.
4. The script will connect to Exchange Online and export the results.

## Output

- Excel file (`.xlsx`) with headers and data.
- CSV file (`.csv`) if Excel is not available.
- Log file (`excel_write_errors.log`) if any Excel write errors occur.

## Author

Christopher Horting

## Example

See `example_output.xlsx` and `example_output.csv` for sample output.

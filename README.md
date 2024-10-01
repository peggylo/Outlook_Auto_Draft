# Outlook Auto Draft

This script automates creating personalized email drafts in Outlook using an Excel file.

## Features

- Reads recipient details from an Excel file.
- Creates customized email drafts in Outlook.
- Attaches files based on a specific file naming pattern.
- Supports environment variables for sensitive information.

## Usage

1. Prepare an Excel file with columns: `school`, `name`, `no`, `email`, `report`.
2. Store the draft email template in your Outlook drafts folder.
3. The script will replace placeholders like `{school}`, `{name}`, `{no}`, and `{report}` with the actual data from the Excel file.
4. Attachments are added based on the `no` field.

## Requirements

- Python 3.x
- `pandas`
- `pywin32`
- Microsoft Outlook

## Environment Variables

Make sure to set the following environment variables:

- `file_path`: Path to the Excel file.
- `attachment_folder`: Path to the folder with attachments.
- `MAIL_SUBJECT`: Email subject.

Example for setting environment variables:

```bash
setx file_path "\\path\to\file.xlsx"
setx attachment_folder "\\path\to\attachments"
setx MAIL_SUBJECT "Your Email Subject"

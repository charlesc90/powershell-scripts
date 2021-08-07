# powershell-scripts

This is a collection of PowerShell scripts to assist various administrative tasks.

# Download-EmailHyperlinks.ps1

This script will download files from hyperlinks in your Outlook emails. You need to -

1. Find the EntryID for the mailbox you want to download from.
2. Use a regex to ensure you are downloading the file from the proper hyperlink
3. Use a regex to match some pattern in the email to name the file
4. Select an output directory

# Extract-DOCEmbeddedPDFs.ps1

This script extracts PDFs embedded in Microsoft Word files. It requires 7zip.

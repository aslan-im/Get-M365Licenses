# Get-M365Licenses
## Introduction
Simple script for collecting data about your M365 licenses in friendly view. Report exports to excel file (ImportExcel module is required).

## Getting Started
Run the script in powershell
```powershell
$Credential = Get-Credential 
.\Get-M365Licenses.ps1 -Credential $Credential
```
If you have the MFA configured, you see the authorization window.

## Prerequisites
- Credential for account with access to MSOL
- Installed module MSOnline
- Installed module ImportExcel
- Powershell version 5.1

# M365 License Optimization Dashboard

Enterprise-grade Microsoft 365 licensing optimization and waste detection platform.

## Features

- Detects disabled users with licenses
- Detects inactive licensed users
- Detects duplicate premium licensing
- Detects unlicensed active users
- Estimates monthly/yearly licensing waste
- Generates dark-mode HTML dashboards
- Exports CSV reports

## Requirements

```powershell
Install-Module Microsoft.Graph.Users
Install-Module Microsoft.Graph.Reports
Install-Module Microsoft.Graph.Identity.DirectoryManagement
```

## Usage

```powershell
.\M365-License-Optimization-Dashboard.ps1 -OpenReport
```

## Configure

Edit:

```json
config.json
```

to match your tenant pricing.

## Safety

This tool is read-only.
It does not modify tenant licensing.

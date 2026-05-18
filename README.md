# M365 License Optimization Dashboard

Enterprise-grade Microsoft 365 licensing audit and optimization platform.

## Features

- Detects inactive licensed users
- Detects disabled licensed users
- Estimates wasted licensing costs
- Generates executive HTML dashboards
- Exports CSV findings
- Supports Microsoft Graph

## Requirements

```powershell
Install-Module Microsoft.Graph.Users
Install-Module Microsoft.Graph.Reports
Install-Module Microsoft.Graph.Identity.DirectoryManagement
```

## Usage

```powershell
.\M365-License-Optimization-Dashboard.ps1
```

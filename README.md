# AI Outlook Troubleshooting Agent

An AI-powered Outlook troubleshooting plugin built with OpenClaw. The agent can perform Outlook operations, execute automated diagnostics using PowerShell, collect logs, test Microsoft Graph connectivity, and assist with Outlook issue resolution.

---

## Features

- Read Outlook Inbox
- Send Outlook Emails
- Read Calendar Events
- Manage Contacts
- Outlook Diagnostics
- Office Repair
- Outlook Restart
- OST/PST Scan
- Navigation Pane Reset
- COM Add-in Diagnostics
- DNS Diagnostics
- Network Diagnostics
- Event Viewer Log Collection
- Microsoft Graph Connectivity Tests
- PowerShell Command Execution

---

# Prerequisites

Install the following before running the project.

## 1. Node.js

Version

```
>= 20.x
```

Download

https://nodejs.org/

---

## 2. npm

Comes with Node.js

Verify

```bash
node -v
npm -v
```

---

## 3. Git

Download

https://git-scm.com/

Verify

```bash
git --version
```

---

## 4. Windows

The troubleshooting tools require

- Windows 10/11
- PowerShell 5.1 or PowerShell 7+
- Microsoft Outlook Desktop

---

## 5. Microsoft 365 Account

Required for Microsoft Graph API.

---

## 6. Azure App Registration

Create an application in Microsoft Entra ID.

Grant Microsoft Graph permissions.

Required scopes:

- User.Read
- Mail.Read
- Mail.ReadWrite
- Mail.Send
- Calendars.Read
- Calendars.ReadWrite
- Contacts.Read
- Contacts.ReadWrite
- offline_access

---

# Clone Repository

```bash
git clone https://github.com/<username>/<repo>.git

cd <repo>
```

---

# Install Dependencies

```bash
npm install
```

or

```bash
pnpm install
```

---

# Required Packages

Core

```text
openclaw
@sinclair/typebox
typescript
tsx
dotenv
```

Microsoft Graph

```text
@azure/msal-node
@microsoft/microsoft-graph-client
```

Utilities

```text
node-fetch
child_process
path
os
fs
```

Development

```text
typescript
tsx
@types/node
```

---

# Environment Variables

Create

```
.env
```

Example

```env
CLIENT_ID=xxxxxxxxxxxxxxxx

TENANT_ID=xxxxxxxxxxxxxxxx

EXECUTION_MODE=simulation
```

For real execution

```env
EXECUTION_MODE=execution
```

---

# Build

```bash
npm run build
```

---

# Run

```bash
npm start
```

or

```bash
npm run dev
```

---

# PowerShell Scripts

Place PowerShell scripts under

```
src/scripts/powershell/
```

Example

```
Check-OutlookRunning.ps1

Repair-Office.ps1

Repair-PST.ps1

Scan-OST.ps1

Reset-NavigationPane.ps1

Collect-EventViewerLogs.ps1

Test-ExchangeConnection.ps1
```

---

# Authentication

The project uses Microsoft Device Code Authentication.

The first login will prompt

```
https://microsoft.com/devicelogin
```

Enter the displayed code.

Subsequent logins use cached tokens.

---

# Simulation Mode

Default

```
EXECUTION_MODE=simulation
```

Returns mocked responses without modifying the system.

---

# Execution Mode

```
EXECUTION_MODE=execution
```

Runs actual PowerShell scripts.

Windows only.

---

# Project Structure

```
src/

│

├── auth.ts

├── graph.ts

├── mail.ts

├── inbox.ts

├── calendar.ts

├── contacts.ts

├── index.ts

│

└── scripts/

    └── powershell/

        ├── Check-OutlookRunning.ps1

        ├── Repair-Office.ps1

        ├── Scan-OST.ps1

        ├── Repair-PST.ps1

        ├── Reset-NavigationPane.ps1

        ├── Collect-EventViewerLogs.ps1

        └── Test-ExchangeConnection.ps1
```

---

# Troubleshooting

### Outlook Authentication

Delete cached tokens and authenticate again.

---

### PowerShell Execution Policy

```powershell
Set-ExecutionPolicy RemoteSigned
```

---

### Build Errors

```bash
rm -rf node_modules

npm install

npm run build
```

---

### Verify Plugin

```bash
npm run validate
```

---

# Future Enhancements

- Autonomous troubleshooting workflow
- AI-driven tool selection
- Help Desk integration (ServiceNow/Jira)
- Hardware diagnostics
- Automatic issue verification
- Intelligent escalation

---

# License

MIT
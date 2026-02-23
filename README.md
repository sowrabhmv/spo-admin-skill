# spo-admin — SharePoint Online Administration Skill for Claude Code

A Claude Code skill for managing SharePoint Online tenants via PowerShell. Uses a **PnP-first approach** — PnP PowerShell is the default for most operations, with the SPO module reserved for multi-geo, cross-geo moves, and SPO-only features.

## What This Skill Does

When you invoke `/spo-admin` in Claude Code, it gives Claude the ability to:

- Connect to your SPO tenant and run PowerShell cmdlets in a persistent session (no repeated MFA prompts)
- Manage sites, users, permissions, storage, sharing, hub sites, and more
- Run multi-geo operations (cross-geo moves, per-geo queries, storage quotas)
- Use PnP PowerShell for list/library management, modern pages, content types, and site provisioning
- Look up official Microsoft documentation and explain results in plain language

## Prerequisites

| Requirement | Details |
|-------------|---------|
| **Claude Code** | Installed and working ([install guide](https://docs.anthropic.com/en/docs/claude-code/overview)) |
| **Windows** | The broker uses `powershell.exe` (5.1) for SPO and `pwsh` (7+) for PnP |
| **PowerShell 7+** | Required for PnP PowerShell. Install from [github.com/PowerShell/PowerShell](https://github.com/PowerShell/PowerShell/releases) |
| **SPO PowerShell Module** | `Microsoft.Online.SharePoint.PowerShell` — install via `Install-Module Microsoft.Online.SharePoint.PowerShell` |
| **PnP PowerShell** (optional) | `PnP.PowerShell` — install in pwsh via `Install-Module PnP.PowerShell -Scope CurrentUser` |
| **SPO Admin access** | You need SharePoint Admin or Global Admin role on your M365 tenant |

## Installation

### Step 1: Clone this repo

```bash
git clone https://github.com/sowrabhmv/spo-admin-skill.git
```

### Step 2: Copy skill files to Claude Code's skill directory

```bash
# Create the skill directory
mkdir -p ~/.claude/skills/spo-admin

# Copy the skill definition and broker script
cp spo-admin-skill/skill.md ~/.claude/skills/spo-admin/
cp spo-admin-skill/SPO-SessionBroker.ps1 ~/.claude/skills/spo-admin/
```

On Windows (PowerShell):
```powershell
# Create the skill directory
New-Item -ItemType Directory -Path "$env:USERPROFILE\.claude\skills\spo-admin" -Force

# Copy the skill definition and broker script
Copy-Item "spo-admin-skill\skill.md" "$env:USERPROFILE\.claude\skills\spo-admin\"
Copy-Item "spo-admin-skill\SPO-SessionBroker.ps1" "$env:USERPROFILE\.claude\skills\spo-admin\"
```

### Step 3: Install PowerShell modules (if not already installed)

**SPO module** (Windows PowerShell 5.1):
```powershell
# In powershell.exe
Install-Module Microsoft.Online.SharePoint.PowerShell -Force
```

**PnP module** (PowerShell 7+ only):
```powershell
# In pwsh
Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
```

### Step 4: First run — tenant configuration

Start Claude Code and run:

```
/spo-admin connect to my tenant
```

The skill will:
1. Ask for your tenant name (e.g., `contoso` if your SharePoint URL is `contoso.sharepoint.com`)
2. Create a `tenant-config.json` in the skill directory with your tenant details
3. Start the SPO session broker and prompt for MFA in your browser (one-time per session)

### Step 5: PnP setup (optional, one-time)

If you need PnP PowerShell features (list/library management, modern pages, provisioning), tell Claude:

```
/spo-admin set up PnP PowerShell for this tenant
```

The skill will:
1. Check if PnP module is installed locally
2. Register an Entra ID app on your tenant (opens browser for Global Admin consent)
3. Save the app's Client ID to your config
4. Verify the PnP connection works

**Note:** PnP app registration requires **Global Administrator** or **Application Developer** role in Entra ID.

## Usage

Once set up, just invoke the skill with any SharePoint admin task:

```
/spo-admin show me all sites sorted by storage usage
/spo-admin check multi-geo move status across all geos
/spo-admin list all external users on the tenant
/spo-admin get the items in the "Documents" library on https://contoso.sharepoint.com/sites/MySite
/spo-admin create a modern page called "Welcome" on the team site
```

## Architecture

### Dual-Broker Design

SPO and PnP PowerShell modules **cannot coexist** in the same PowerShell session (CSOM assembly version conflict). The skill uses two separate background processes:

| Broker | Shell | Session Directory | Module | Use Case |
|--------|-------|-------------------|--------|----------|
| **PnP** (primary) | `pwsh` (7+) | `~\.pnp-session\\` | PnP.PowerShell | **Default for most operations** — sites, tenant settings, lists, libraries, pages, permissions |
| **SPO** | `powershell.exe` (5.1) | `~\.spo-session\\` | Microsoft.Online.SharePoint.PowerShell | SPO-only ops — multi-geo, cross-geo moves, external users, site health |

Both brokers can run simultaneously. The skill automatically routes commands to the correct broker based on which cmdlets are needed.

### Session Broker

The `SPO-SessionBroker.ps1` script:
- Runs as a background process
- Loads the module and authenticates once (MFA only on first connect)
- Watches for command files, executes them in the live session, returns output
- Keeps the connection alive across multiple commands (no repeated logins)

### Config File

`tenant-config.json` stores all tenant-specific settings locally. It is **never committed to the repo** (listed in `.gitignore`). Each user creates their own config during first-time setup.

```json
{
  "TenantName": "contoso",
  "AdminUrl": "https://contoso-admin.sharepoint.com/",
  "TenantRoot": "https://contoso.sharepoint.com",
  "PnPClientId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "OnMicrosoftDomain": "contoso.onmicrosoft.com"
}
```

## Files

| File | Purpose | Committed? |
|------|---------|------------|
| `skill.md` | Skill definition — instructions, cmdlet reference, broker protocol | Yes |
| `SPO-SessionBroker.ps1` | Background session broker script | Yes |
| `tenant-config.json` | Per-user tenant config (tenant name, Client ID) | No (gitignored) |
| `.gitignore` | Excludes tenant-config.json | Yes |

## Troubleshooting

| Issue | Fix |
|-------|-----|
| MFA prompt every time | The broker may have died. Delete the `.spo-session` or `.pnp-session` folder and let Claude restart it |
| PnP consent error | A Global Admin needs to run the PnP app registration. Use `/spo-admin set up PnP PowerShell` |
| SPO module not found | Install it: `Install-Module Microsoft.Online.SharePoint.PowerShell` in `powershell.exe` |
| PnP module not found | Install it: `Install-Module PnP.PowerShell -Scope CurrentUser` in `pwsh` |
| "Cannot load assembly" error | SPO and PnP are being loaded in the same session. The broker's `-Mode` parameter prevents this — ensure you're using the latest `SPO-SessionBroker.ps1` |
| Connection timeout | Use a longer timeout on the Bash command (e.g., `timeout: 300000`) |

## License

Internal use only. Do not distribute outside your organization.

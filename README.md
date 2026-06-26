# Outlook-Purview-Sensitivity

Displays Microsoft Purview sensitivity labels (PII, Confidential, etc.) as a sortable column in the Outlook message list — across **all** folders, not just the Inbox.

Built as a VSTO add-in for classic Outlook (desktop), 64-bit.

## Features

- Adds a **PurviewLabel** user-defined field to every folder you open
- Automatically switches non-TableView folders (Sent Items, subfolders) to Messages view so the column always appears
- Reads the `msip_labels` MAPI property via `PropertyAccessor` to extract the label name
- Stamps the label into the custom field so it appears as a sortable column
- No modification of email subjects, categories, or body content
- Startup health check logs whether the add-in loaded successfully

## System Requirements

| Component | Requirement |
|-----------|------------|
| Windows | Windows 10 or later (x64) |
| Outlook | Microsoft Outlook 2016+ (64-bit) |
| .NET | .NET Framework 4.7.2 |
| VSTO Runtime | Visual Studio 2010 Tools for Office Runtime 4.0 |
| Build tooling | Visual Studio 2022 (Community or higher) |

> **Note:** The VSTO runtime is included with Office 2016+ and Visual Studio. If missing, it is downloaded automatically by `setup.exe`.

## How It Works

1. On startup, the add-in runs a health check and logs status to the debug output
2. When any folder is opened (including Sent Items, subfolders, shared mailboxes), it:
   - Adds the `PurviewLabel` user-defined property if missing
   - Switches non-TableView views to **Messages** view to display the column
   - Adds the `PurviewLabel` column to the view
3. The first ~50 labeled items in the folder are stamped automatically on folder switch
4. The label name is parsed from the `msip_labels` MAPI string (e.g., `PII High`, `General`)

## Build & Debug (for development)

```powershell
# 1. Clone the repo
git clone <repo-url>
cd outlook-purrrview-sensitivity

# 2. Open in Visual Studio 2022
Start-Process Outlook-Purview-Sensitivity.slnx

# 3. Press F5 to build, register, and launch Outlook with the add-in loaded
```

The add-in runs with `LoadBehavior=3` (load at startup) when used via F5 debugging.

## Deployment (for end users)

You have two deployment options:

### Option A: ClickOnce Setup (recommended for most users)

```powershell
# 1. Build and publish from Visual Studio
#    Right-click project → Publish → select "FolderProfile" → Publish
#    OR from command line:
msbuild /t:Publish /p:PublishProfile=FolderProfile /p:Configuration=Release

# 2. Distribute the contents of bin\Release\app.publish\ to users
# 3. Users run setup.exe to install
```

`setup.exe` installs the add-in, downloads the VSTO runtime if needed, and registers it with `LoadBehavior=3`.

### Option B: PowerShell Script (for IT admins / Intune)

```powershell
# After publishing (see Option A above), run:
.\Install-AddIn.ps1 -InstallPath .\bin\Release\app.publish

# To uninstall:
.\Install-AddIn.ps1 -Uninstall
```

The script:
- Copies files to `%LocalAppData%\Outlook-Purview-Sensitivity\`
- Registers the add-in in `HKCU\Software\Microsoft\Office\Outlook\AddIns\`
- Sets `LoadBehavior=3` so the add-in loads on startup
- Does NOT require admin rights (HKCU registration)

### Intune / SCCM Deployment

1. Publish the add-in using Option A
2. Package `bin\Release\app.publish\` and `Install-AddIn.ps1` together
3. Deploy `Install-AddIn.ps1 -InstallPath <path>` as a user-context PowerShell script
4. The add-in installs per-user without admin elevation

## Troubleshooting

### Column does not appear in Sent Items or subfolders

The add-in now automatically detects non-TableView folders and switches them to **Messages** view before adding the PurviewLabel column. If the column still doesn't appear:

1. Manually switch the folder to **Messages** view (View → Change View → Messages)
2. Restart Outlook to trigger the startup check
3. Check debug output for errors (use DebugView from Sysinternals)

### Add-in does not load on startup

1. Check registry: `HKCU\Software\Microsoft\Office\Outlook\AddIns\Outlook-Purview-Sensitivity`
   - Verify `LoadBehavior` is `3` (REG_DWORD)
   - Verify `Manifest` points to a valid `.vsto` file
2. Check Outlook's disabled items list:
   - File → Options → Add-ins → Manage: Disabled Items → Go
   - Re-enable if listed
3. Re-run `Install-AddIn.ps1 -InstallPath <path>` to re-register

### "Save failed" errors in debug output

This is expected for some folder types (PST files, delegate mailboxes, shared mailboxes where the user lacks write permission). The add-in logs these as informational messages and continues processing. Stamping will succeed on folders where the user has write access.

### Viewing debug output

Download DebugView from Sysinternals and run it as administrator:
```
https://learn.microsoft.com/en-us/sysinternals/downloads/debugview
```

Filter for `[PS]` (startup/heath checks) or `[CM]` (column manager operations).

### Add-in crashes or Outlook hangs

1. Restart Outlook
2. If persistent, uninstall with `.\Install-AddIn.ps1 -Uninstall`
3. Check Event Viewer → Windows Logs → Application for errors from `Outlook.exe` or `VSTO`
4. Ensure you have the 64-bit version of Outlook installed (32-bit not supported)

### Shared / delegate mailboxes

The add-in now handles these gracefully. If a `Save()` operation fails due to insufficient permissions, the error is logged and processing continues. The column will still be added to the view.

## Architecture

| File | Responsibility |
|------|---------------|
| `ThisAddIn.cs` | Startup, shutdown, explorer event wiring, health checks |
| `ColumnManager.cs` | User-defined property creation, column management, view switching, item stamping |
| `LabelReader.cs` | Reads `msip_labels` from MAPI `PropertyAccessor` |
| `LabelResolver.cs` | Parses label name from the `msip_labels` string |

## License

MIT

# Outlook-Purview-Sensitivity

Displays Microsoft Purview sensitivity labels (PII, Confidential, etc.) as a sortable column in the Outlook inbox message list.

Built as a VSTO add-in for classic Outlook (desktop), 64-bit.

## Features

- Adds a **PurviewLabel** user-defined field to any folder you open
- Reads the `msip_labels` MAPI property from each email to extract the label name
- Stamps the label into the custom field so it appears as a column
- No modification of email subjects, categories, or body content

## Prerequisites

- Visual Studio 2022 (Community or higher)
- .NET Framework 4.7.2
- Microsoft Outlook 2016+ (64-bit)
- VSTO 4.0 runtime (included with Visual Studio / Office)

## Build & Run

1. Clone the repo
2. Open `Outlook-Purview-Sensitivity.slnx` in Visual Studio
3. Press **F5** to build, register, and launch Outlook with the add-in loaded
4. Open your inbox — the **PurviewLabel** column will be added automatically

To deploy, publish via **Build → Publish** and distribute the ClickOnce output through Intune or your deployment tool.

## How It Works

1. On startup, the add-in adds a `PurviewLabel` user-defined property to the current folder
2. It reads each email's `msip_labels` MAPI header via `PropertyAccessor`
3. The label name is parsed from the `msip_labels` string (e.g., `PII High`, `General`)
4. The value is written to the `PurviewLabel` field so Outlook can display it

When you switch folders, the first ~50 labeled items in the view are stamped automatically.

## License

MIT

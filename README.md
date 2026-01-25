# OM Downloader 🚀

An automated Order Management (OM) report downloader built with Node.js and Playwright. This tool automates the process of logging into the Marico apps portal, selecting report parameters, and exporting data to Excel.

## ✨ Features

- **Automated Login**: Securely handles Microsoft-based authentication popups.
- **Micro-Interaction Automation**: Automatically selects required report parameters and triggers the export.
- **Multi-Level Notifications**:
  - 🔊 **Audio Alert**: System beep on failure.
  - 🍞 **Toast Notification**: Windows system notification.
  - 🖼️ **GUI Popup**: Persistent PowerShell-based dialog for critical failures.
- **Robustness**: Built-in 5-attempt retry mechanism with 10-second delays.
- **Headless Mode**: Option to run in the background or see the automation in action.
- **Smart Cleanup**: Automatically closes existing Edge instances to prevent session conflicts (non-headless mode).

## 📋 Prerequisites

- [Node.js](https://nodejs.org/) (v16 or higher recommended)
- [Microsoft Edge Browser](https://www.microsoft.com/edge)

## 🚀 Installation

1. Clone or download this repository.
2. Run the `setup.bat` file to install all dependencies:
   ```bash
   ./setup.bat
   ```
   *This will install necessary npm packages and set up the Playwright Edge driver.*

## ⚙️ Configuration

Before running, update the `config.json` file with your credentials and preferences:

```json
{
  "email": "your-email@marico.com",
  "password": "your-password",
  "reportUrl": "https://mblnet.maricoapps.biz/...",
  "headless": true,
  "outputFolder": "C:/Downloads/OM_Reports"
}
```

| Field | Description |
| :--- | :--- |
| `email` | Your company login email. |
| `password` | Your company login password. |
| `reportUrl` | The direct URL to the report management page. |
| `headless` | `true` to run in background; `false` to see browser window. |
| `outputFolder` | Path where reports should be saved (defaults to current directory if empty). |

## 🛠️ Usage

### Standard Run
To start the automation with your `config.json` settings:
```bash
node automate_om.js
```

### Headless Profile Run (Experimental)
To run using a temporary persistent profile:
```bash
node automate_om_headless.js
```

## 🔍 Troubleshooting

- **Login Issues**: If your account requires Multi-Factor Authentication (MFA), try running with `"headless": false` for the first time to complete the authentication manually.
- **Edge Not Found**: Ensure Microsoft Edge is installed and updated to the latest version.
- **Failed Attempts**: The script takes a screenshot (`debug_error.png`) on failure to help you identify what went wrong.

---

**Disclaimer**: This tool is designed for internal productivity. Please ensure its use complies with company IT policies.

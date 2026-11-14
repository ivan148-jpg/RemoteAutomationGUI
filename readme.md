# Remote Automation GUI (RAG)

A lightweight, portable Windows PowerShell graphical utility designed for system administrators to manage and execute remote operations, scripts, queries, and maintenance across distributed endpoints via WinRM.

---
## 🚀 Quick Start & Deployment

### Option 1: One-Click Windows Installer (Recommended)
For an automated installation that configures WinRM, installs required SQL modules, Sets execution policy to RemoteSigned and creates desktop shortcuts:

👉 **[Download InstallRAG.exe (Latest Release)]([https://github.com](https://github.com/ivan148-jpg/RemoteAutomationGUI/releases/download/1.0.1/InstallRAG.exe))**

1. Run **`InstallRAG.exe`** and follow the setup wizard.
   > *(Note: If Windows SmartScreen appears, click **More info** → **Run anyway**, as the installer is not signed with a paid digital certificate.)*
2. Once installation finishes, launch **`RAG.exe`** from your Desktop or Start Menu to start the application.
3. Configure TrustedHosts (See below: option2: step 2 )

---

### Option 2: Manual Setup from Source
Alternatively, you can download or clone the full repository and run the setup script manually using PowerShell:

### 1. Run Setup Script
Open PowerShell **as Administrator** in the project directory, and execute `SetupRAG.ps1`:

```powershell
# Unblock repository files if downloaded from the web
Get-ChildItem -Recurse | Unblock-File

# Run environment setup
powershell.exe -ExecutionPolicy Bypass -File .\SetupRAG.ps1
```

The setup script automatically handles:
* **Admin Elevation**: Prompts for elevation if not already running as Administrator.
* **WinRM Configuration**: Enables and configures WinRM (`winrm quickconfig -quiet`).
* **Execution Policy**: Sets execution policy to `RemoteSigned` for local machine scope.
* **Module Installation**: Checks for and installs the `SqlServer` PowerShell module.
* **TrustedHosts Import**: Reads `TrustedHosts_list.txt` in the root directory and configures initial remote endpoint permissions. (If empty, you can add trusted hosts later via the GUI and Trusted Host script).

### 2. Configure TrustedHosts
Because this utility relies on PowerShell Remoting (WinRM), target servers must be recognized as trusted endpoints:
1. **Populate Your Server List**: Add your target servers (one IP or hostname per line) into a `.txt` file inside the `Lists\` folder (e.g., `All_list.txt`).
2. **Select Servers in the GUI**: Launch `RAG.ps1`, choose your server list from the dropdown, and check your target servers.
3. **Run TrustedHosts Expansion Script**: Select `Trusted_Host_expand_list_from GUI` from the scripts dropdown menu and click **Run Script**. This will automatically append your selected servers to your local WinRM `TrustedHosts` configuration.

### 3. Launch the Application
Run the core GUI application using PowerShell:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\RAG.ps1
```

### 4. Setup Credentials (First-Run)
Select your target servers from the server list dropdown, then use the top control buttons to securely store your credentials:
* **Create and Encrypt Admin Credential** → Stores your Windows Administrator login securely using DPAPI.
* **Create and Encrypt SQL Credential** → Stores your SQL authentication parameters.
* **Create and Encrypt VNC Passwords** → Encrypts viewer credentials for remote VNC management.

> **Note**: Encrypted credentials are bound via DPAPI to the current Windows user profile on that specific machine for maximum security. Each operator should run credential creation once on their local machine.

---

## 📂 Project Structure

Place your modular files in their designated subdirectories:

```text
RemoteAutomationGUI/
│
├── .github/                # GitHub community configurations (FUNDING.yml)
├── Assets/                 # Icons and UI images (RAG.ico, Restart_icon.png)
├── Bat files/              # Batch (.bat/.cmd) scripts for remote execution
├── Lists/                  # Target server list text files (.txt, one IP/hostname per line)
├── Logs/                   # Auto-created directory for execution logs (Git Ignored)
├── Scripts/                # Custom PowerShell (.ps1) task scripts
├── Secure/                 # Auto-created directory for DPAPI encrypted credentials (Git Ignored)
├── Tools/                  # Executables and supporting utilities (e.g., UltraVNC viewer)
├── Trusted hosts/          # TrustedHosts configuration assets and output scripts
│
├── LICENSE                 # Open-source MIT license
├── RAG.ps1                 # Primary application GUI script
├── readme.md               # Project documentation
├── SetupRAG.ps1            # Automated environment configuration script
├── Skeleton_for_New_scripts.txt # Starter template for authoring custom task scripts
└── TrustedHosts_list.txt   # Initial IP list used during environment setup
```

---

## 🛠️ Key Features & Actions

* **File Operations**: Copy files instantly to remote shares or execute remote scripts, batches, MSIs, and executables (`.bat`, `.cmd`, `.ps1`, `.vbs`, `.exe`, `.msi`).
* **Database Management**: Execute structured SQL queries directly across selected servers.
* **System Monitoring & Tools**:
  * Real-time ICMP ping checks.
  * One-click RDP launcher (`mstsc.exe`).
  * Integrated UltraVNC connection handler.
  * Direct Windows Explorer shortcuts to remote shares (`C$` / `D$`).
  * Clean remote OS restarts with immediate visual feedback.
* **Safety First**: Built-in Dry Run Mode simulates all commands and file transfers without making actual changes to target infrastructure.

---

## 📐 High DPI Displays

If the GUI elements appear crowded or blurry on high-resolution displays:
1. Ensure your host machine is running .NET 4.7+ or PowerShell 7+, where Windows Forms native DPI scaling is enabled automatically.
2. If running via standard `powershell.exe` on legacy host environments, you can adjust high DPI scaling properties on your PowerShell shortcut properties menu under **Compatibility > Change high DPI settings**.

---

## ☕ Support & Donations

If this tool saves you time, simplifies your administration tasks, or helps your organization's workflow, consider supporting its ongoing development and maintenance!

[![Donate via PayPal](https://img.shields.io/badge/Donate-PayPal-00457C?style=for-the-badge&logo=paypal&logoColor=white)](https://paypal.me/IPavlovic836)
[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ffdd00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/IPavlovic836)

Your support is greatly appreciated!

---

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information.

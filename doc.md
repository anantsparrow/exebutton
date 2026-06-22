# Building the Outlook Phishing Reporter MSI Installer

This guide explains how to compile the C# Outlook VSTO Add-in and generate the MSI installer on a Windows machine.

---

## 📋 Prerequisites

Before building the installer, ensure the following software is installed on the Windows build machine:

1. **Visual Studio (2019 or 2022)**:
   * During installation, select the **Office/SharePoint development** workload.
   * This includes the MSBuild VSTO targets needed to build and package VSTO add-ins.
2. **WiX Toolset v3.11**:
   * Download and run the installer from the [official WiX website](https://wixtoolset.org/).
   * Verify that WiX is either added to your system `PATH` or resides at `C:\Program Files (x86)\WiX Toolset v3.11\bin\`.

---

## 🛠️ Step 1: Compile the C# Project in Release Mode

To package the files, you must first compile the production-ready Release binaries:

1. Copy the project folders to your Windows build host.
2. Open Visual Studio.
3. Open `PhishingReporterAddIn.csproj`.
4. In the toolbar, change the active configuration from **Debug** to **Release** (and make sure **Any CPU** is selected).
5. Build the project (`Ctrl + Shift + B` or select **Build > Build Solution** from the menu).
6. Verify that the compilation completed successfully. The binaries will be written to:
   `PhishingReporterAddIn\bin\Release\`
   * *Required output files*:
     * `PhishingReporterAddIn.dll`
     * `PhishingReporterAddIn.vsto`
     * `PhishingReporterAddIn.dll.manifest`

---

## 📦 Step 2: Build the MSI Package

The installer directory (`Setup/`) contains the compiler script:

1. Navigate to the `Setup/` folder.
2. Double-click **`build_wix.bat`** (or open a command prompt in that directory and run `build_wix.bat`).
3. The script executes the following stages:
   * **Compilation (`candle.exe`)**: Parses the `Setup.wxs` XML definitions and translates them into an intermediate object file `Setup.wixobj`.
   * **Linking (`light.exe`)**: Links the object file and bundles the built add-in files (Release binaries and default config) into the final MSI installer **`PhishingReporterSetup.msi`**.

If you want to run these commands manually, execute:
```cmd
"C:\Program Files (x86)\WiX Toolset v3.11\bin\candle.exe" Setup.wxs -o Setup.wixobj
"C:\Program Files (x86)\WiX Toolset v3.11\bin\light.exe" Setup.wixobj -out PhishingReporterSetup.msi
```

---

## 🚀 Step 3: Installation & Configuration

1. **Install**: Double-click the generated `PhishingReporterSetup.msi` to run the setup wizard.
2. **Settings**: The installer drops the default settings template at:
   `C:\ProgramData\PhishingReporter\config.json`
   To configure the REST API endpoint, open this JSON file and edit the `ApiUrl` value:
   ```json
   {
     "ApiUrl": "https://webhook.site/YOUR_WEBHOOK_GUID",
     "TimeoutSeconds": 30
   }
   ```
3. **Verify Outlook**: Launch Outlook. A new group named **Phishing Triage** will appear in your ribbon bar with the custom exclamation shield icon on the **Home** mail list and the **Read Message** window.

---

## 🔍 Troubleshooting

### Add-In is Not Loading in Outlook
If the button does not appear after installation:
1. Open Outlook and navigate to **File > Options > Add-ins**.
2. Look for "Report Phishing" in the list:
   * **Disabled Items**: If it is listed under "Disabled Items", select **Manage: Disabled Items** at the bottom, click **Go**, select the add-in, and click **Enable**.
   * **Inactive Add-ins**: If it is listed under "Inactive Application Add-ins", select **Manage: COM Add-ins**, click **Go**, and make sure the checkbox next to **Report Phishing** is checked.
3. Check the registry keys using `regedit.exe`. They should be registered under:
   * `HKLM\Software\Microsoft\Office\Outlook\Addins\PhishingReporter` (64-bit Office)
   * `HKLM\Software\Wow6432Node\Microsoft\Office\Outlook\Addins\PhishingReporter` (32-bit Office)
   Verify that:
   * `LoadBehavior` is set to `3` (Load on startup).
   * `Manifest` points to the correct location: `C:\Program Files\Phishing Triage\Phishing Reporter\PhishingReporterAddIn.vsto|vstolocal`.

### Administrative Privileges
Since the MSI registers the add-in machine-wide under `HKLM` and installs to `Program Files`, you must run the MSI installer with **Administrator Privileges**.
To install silently or as admin via CLI:
```cmd
msiexec /i PhishingReporterSetup.msi /quiet /qn
```

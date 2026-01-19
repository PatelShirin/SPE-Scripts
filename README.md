# SPE-Scripts Usage Guide.

This repository contains PowerShell scripts for Sitecore and related automation tasks. Follow these steps to ensure you can run the scripts successfully as a first-time user:

## Prerequisites

1. **PowerShell**
   - Ensure you have PowerShell 5.1 or later installed on your system.
   - On Windows, PowerShell is included by default. For other platforms, download from [Microsoft PowerShell](https://github.com/PowerShell/PowerShell).

2. **Execution Policy**
   - Scripts may be blocked by default. To allow script execution, run PowerShell as Administrator and execute:
     ```powershell
     Set-ExecutionPolicy RemoteSigned
     ```
   - This allows local scripts to run and requires remote scripts to be signed.

3. **Required Modules**
   - Some scripts may require additional modules (e.g., Sitecore PowerShell Extensions, PnP.PowerShell, etc.).
   - Install modules as needed using:
     ```powershell
     Install-Module -Name <ModuleName> -Scope CurrentUser
     ```
   - Refer to individual script headers for specific module requirements.

## Running Scripts

1. Open PowerShell and navigate to the script directory:
   ```powershell
   cd "<path-to-repo>\SPE-Scripts\PS-CustomScripts"
   ```
2. Run a script by executing:
   ```powershell
   .\<script-name>.ps1
   ```
   - Example:
     ```powershell
     .\Get-SitecoreSession.ps1
     ```

## Troubleshooting

- If you see a policy error, check your execution policy.
- If a module is missing, install it as described above.
- For script-specific instructions, see comments at the top of each script file.

## Contributing

- Please document any new scripts and specify prerequisites in the script header.
- **Submit pull requests for improvements or bug fixes. This is VERY IMPORTANT to ensure the quality and maintainability of these scripts.**

---


For further help, contact the repository maintainer or refer to the script comments.

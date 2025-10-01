# 🗺️ Excel Coordinate Distance and Travel Time Updater ⏱️

A PowerShell script to **automatically calculate the driving distance and travel time** between pairs of geographic coordinates (latitude and longitude) and update a specified Excel file. It uses the Open Source Routing Machine (OSRM) API for its routing calculations.

---

## ✨ Features

- **Coordinate-to-Distance:** Calculates distance and travel time between source and destination coordinates in an Excel file.
- **OSRM Integration:** Leverages the robust [Project OSRM API](https://project-osrm.org/) for fast, detailed routing information.
- **Configuration Driven:** All settings are managed through a single, easy-to-use JSON configuration file.
- **Flexible Output:** Updates specified columns in the Excel file with the calculated **distance (km)** and **duration (minutes)**.
- **Logging & Notifications:** Built-in support for detailed log files (CSV, JSON, TXT, XLSX) and email notifications on completion or error.

---

## 🚀 Getting Started

### Prerequisites

- **PowerShell 7.x or later:** This script is built for modern PowerShell.
- **OSRM Connectivity:** The machine running the script must have internet access to communicate with the OSRM routing service (e.g., `https://router.project-osrm.org/`).
- **Excel Module:** The script requires the `ImportExcel` PowerShell module to read and write data from the Excel file (`.xlsx`).
  ```powershell
  Install-Module -Name 'ImportExcel' -Scope 'AllUsers'
  ```
- **Email Modules (Optional):** If you plan to use the email notification feature, you must install the `MailKit` and `MimeKit` packages:
  ```powershell
  Install-Package -Name 'MailKit' -Source '[https://www.nuget.org/api/v2](https://www.nuget.org/api/v2)' -Scope 'AllUsers' -SkipDependencies
  Install-Package -Name 'MimeKit' -Source '[https://www.nuget.org/api/v2](https://www.nuget.org/api/v2)' -Scope 'AllUsers' -SkipDependencies
  ```

### Installation

1.  Clone this repository to your local machine:
    ```bash
    git clone https://github.com/DarkLite1/Update-Excel-with-distance.git
    cd Update-Excel-with-distance
    ```
2.  Review `Example.json` to understand the configuration options.
3.  Create your own configuration file (e.g., `MyConfig.json`) based on `Example.json`.

## 📝 Usage

Execute the script from PowerShell, providing the path to your configuration JSON file.

```powershell
& '.\Update Excel with distance.ps1' -ConfigurationJsonFile '.\MyConfig.json'
```

### Scheduled Task Example ⏱️

For automated, recurring executions (e.g., via Windows Task Scheduler), you'll typically invoke the script using `pwsh.exe -Command`. This allows you to pass parameters directly.

- Program: `pwsh.exe`
- Arguments:

```
-Command  "& 'C:\Update-Excel-with-distance\Update Excel with distance.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit $LASTEXITCODE"
```

## 📞 Support

If you encounter any issues or have questions, please open an issue on the GitHub repository.

## 📜 License

This software is provided for your use, free of charge, for both private and professional purposes.

**NO GUARANTEES ARE GIVEN:** This software is provided "as is," without any warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and non-infringement.

**YOUR RESPONSIBILITY:** In no event shall the author be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the software or the use or other dealings in the software. You are solely responsible for the code you deploy and its consequences.
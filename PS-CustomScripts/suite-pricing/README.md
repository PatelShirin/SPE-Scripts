# Suite Pricing Monthly Update Guide

This folder contains scripts and resources for updating suite plan pricing on a monthly basis.

## Overview

The monthly update process ensures that suite pricing data is refreshed and exported for business use. The main scripts involved are:

- `suite-plan-pricing-monthly-update.ps1`: Automates the monthly update of suite pricing.
- `suite-plan-pricing-excel-file-import.ps1`: Imports pricing data from Excel files.

## Monthly Update Steps

1. **Prepare Excel Files**

   - Place the latest monthly pricing Excel files in the `excel/<year>/<month>/` subfolder.

2. **Run the Import Script (Automatically Updates Pricing)**

   - Open PowerShell and navigate to this folder:
     ```powershell
     cd "<path-to-repo>\\SPE-Scripts\\PS-CustomScripts\\suite-pricing"
     ```
   - Run the import script:
     ```powershell
     .\\suite-plan-pricing-excel-file-import.ps1
     ```
   - Follow any prompts to select the correct Excel file if required.
   - The import script will automatically call the monthly update script (`suite-plan-pricing-monthly-update.ps1`) after importing the data, so no additional manual step is required.

3. **Export Updated Data**
   - The updated pricing data will be exported to the `export/<year>/<month>/` folder as a CSV file.
   - Example: `2026-market-rate-business-case-tracker-c2c-web.csv`

## Notes

- Ensure you have the necessary permissions and modules installed (see root README).
- Review script comments for any script-specific requirements or options.
- Back up previous data before running updates.

## Troubleshooting

- If you encounter errors, check the file paths and naming conventions.
- Ensure Excel files are not open in another program during import.
- Refer to script comments for additional help.

---

For questions or issues, contact the repository maintainer or refer to the root README for general setup instructions.

# XLS to XLSM

A template repository that converts legacy `.xls` files to macro-enabled `.xlsm` format using a GitHub Actions workflow and a PowerShell script powered by Excel COM automation.

## How to use

1. Create a new repo by clicking **`Use this template`** then **`Create a new repository`** in the top-right corner of the screen.
2. Add the `.xls` files you want to convert to the `input/` folder and commit them.
3. The workflow will run automatically on every push to `main` that changes files in `input/`. You can also trigger it manually from the **Actions** tab by selecting the **XLS to XLSM** workflow and clicking **Run workflow**.

Once the workflow finishes, the converted `.xlsm` files are uploaded as a build artifact named **XLSM-Files**, which you can download from the workflow run summary page.

> **Note:** The workflow installs Microsoft Office on the runner, which takes about 5 minutes.

## How it works

| File | Description |
|------|-------------|
| `input/` | Place your `.xls` files here |
| `scripts/Convert-XlsToXlsm.ps1` | PowerShell script that opens each `.xls` file via Excel COM and saves it as `.xlsm` |
| `.github/workflows/Build-VBA.yml` | GitHub Actions workflow that sets up Excel, runs the conversion script, and uploads the results |

The conversion script accepts two optional parameters:

```powershell
./scripts/Convert-XlsToXlsm.ps1 -InputDir "./input" -OutputDir "./output"
```

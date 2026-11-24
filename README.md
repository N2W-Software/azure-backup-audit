<a href="https://n2ws.com"><img src="docs/n2w-logo.png" alt="N2W" width="80" align="right"></a>

# Azure Backup Estate Audit

See what Azure Backup is protecting in your tenant, what it isn't, how many vaults and policies that takes, and what it costs at list price. Runs inside your own Azure portal or Cloud Shell, as you, read-only.

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2FN2W-Software%2Fazure-backup-audit%2Fmain%2Fworkbook%2Fazuredeploy.json)

## What it shows

| Tab | What you get |
|---|---|
| 1 · Estate | VMs, how many are backed up, which aren't (prod-tagged first), unattached disks, old snapshots, SQL in VMs |
| 2 · Vaults and plumbing | Vaults by region and redundancy, policies that repeat each other or protect nothing, failed jobs in the last 14 days, report Logic Apps, and which vaults send diagnostics to Log Analytics |
| 3 · Cost | Azure Backup at list price, line by line with the arithmetic shown, next to the same estate on N2W. Every price is a parameter you can change |
| 4 · Export summary | One line of counts and prices, no names, if you want an itemised receipt or a review |

## Is it safe to run?

- **Read-only.** It runs Azure Resource Graph queries and a few `GET` calls, as you. Reader on the subscriptions is enough.
- **Nothing leaves your tenant.** A workbook can't send data anywhere. The summary on tab 4 leaves only if you copy and paste it yourself.
- **You can read all of it first.** The workbook is one JSON file and the script is one PowerShell file. Every query is in [`build/build_workbook.py`](build/build_workbook.py).

## Three ways to run it

**1. Deploy to Azure** (button above). Creates one workbook resource in a resource group you choose, so you need Workbook Contributor there. Then open it under *Monitor → Workbooks*.

**2. Paste it, save nothing.** In the portal: *Monitor → Workbooks → New → Advanced editor* (`</>`), paste [`workbook/n2w-azure-backup-audit.workbook.json`](workbook/n2w-azure-backup-audit.workbook.json), *Apply*. It runs without being saved.

**3. The script**, for live prices, real vault storage and CSV files. In Azure Cloud Shell (PowerShell):

```powershell
curl -sLO https://raw.githubusercontent.com/N2W-Software/azure-backup-audit/main/script/Invoke-N2WBackupAudit.ps1
./Invoke-N2WBackupAudit.ps1
```

Options: `-SubscriptionId`, `-Region westeurope`, `-UsedRatio 70`, `-N2WTier Cold`, `-N2WServerMonthly 140`, `-Offline`. Run `Get-Help ./Invoke-N2WBackupAudit.ps1 -Full` for all of them.

### Running the script on your own machine (macOS or Linux)

The same script runs in PowerShell 7 on macOS and Linux. There's no separate shell version.

1. Install PowerShell 7:

   ```bash
   # macOS (Homebrew)
   brew install --cask powershell

   # Linux (snap; or use Microsoft's apt/dnf package repository)
   sudo snap install powershell --classic
   ```

   Other options: [PowerShell on macOS](https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-macos), [PowerShell on Linux](https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-linux).

2. Start PowerShell, install the two Azure modules the script uses, and sign in:

   ```powershell
   pwsh
   Install-Module Az.Accounts, Az.ResourceGraph -Scope CurrentUser
   Connect-AzAccount            # add -UseDeviceAuthentication on a machine without a browser
   ```

3. Download and run the script, still in `pwsh`:

   ```powershell
   curl -sLO https://raw.githubusercontent.com/N2W-Software/azure-backup-audit/main/script/Invoke-N2WBackupAudit.ps1
   ./Invoke-N2WBackupAudit.ps1
   ```

The account you sign in with needs Reader on the subscriptions you want to scan, same as in Cloud Shell. The CSV and summary files are written to the current folder, or to `-OutputFolder`.

## Workbook vs script

| | Workbook | Script |
|---|---|---|
| Where | Azure portal | Cloud Shell or local PowerShell with Az |
| Prices | Parameters, defaults are East US list on the release date | Live from Microsoft's Retail Prices API, your region |
| Vault storage | Estimate (one copy of protected data) unless you type the real figure | Measured from each vault |
| Diagnostic settings | Per vault, click a row | Counted for every vault |
| Output | Dashboards, Export to Excel on any grid | Report in the terminal, two CSVs and a summary file |

## How the cost is worked out

Azure Backup charges a protected-instance fee per VM, tiered by size (up to 50 GB: half a unit; then one unit per started 500 GB), plus vault storage by redundancy. Resource Graph sees provisioned disk size, not used data, so the fee shown is an upper bound unless you lower *Used ÷ provisioned*. Full formulas and price sources: [`docs/PRICING.md`](docs/PRICING.md).

The N2W side is a flat price per VM plus Blob storage in your own account at the tier you pick. Both sides leave out instant-restore snapshots, Log Analytics ingestion and storage transactions.

## Feedback

Questions or feedback: contact N2W at [n2ws.com](https://n2ws.com). This repository takes issues and pull requests from the N2W team only. Built by [N2W](https://n2ws.com). MIT licensed.

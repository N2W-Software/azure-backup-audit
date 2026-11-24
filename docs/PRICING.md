# Pricing model and sources

Same model in the workbook (`build/build_workbook.py`, `PRICES`) and the script (`$FallbackPrices` plus the live API).

## Azure Backup, per protected VM, per month

| Used data in the VM | Fee |
|---|---|
| up to 50 GB | half a unit ($5 at East US list) |
| over 50 GB, up to 500 GB | one unit ($10) |
| over 500 GB | one unit per started 500 GB ($20 for 501–1,000 GB, and so on) |

Source: Azure Backup pricing page ("Instance < or = 50 GB", "> 50 GB but < or = 500 GB", "> 500 GB: for each 500 GB increment"), unit price from the Retail Prices API meter *Backup / Azure VM / Azure VM Protected Instance* = $10.00, East US.

**Instance size is used data, not provisioned.** The pricing page: "the size calculation is based on actual (used) size of VM... the sum of all data in the VM, excluding temporary storage." Resource Graph only exposes provisioned managed-disk size (`properties.diskSizeGB`), so both tools multiply it by *Used ÷ provisioned* (default 100%, the upper bound).

## SQL Server in Azure VM

$25 per instance per month up to 500 GB of configured databases, then $25 per started 500 GB (meter *SQL Server in Azure VM Protected Instance*). Database size isn't in Resource Graph, so both tools charge every instance at the lowest tier. That understates the Azure Backup side.

## Vault storage, $/GB-month, East US list

| Redundancy | Azure Backup (Standard) |
|---|---|
| LRS | 0.0224 |
| ZRS | 0.028 |
| GRS | 0.0448 |
| RA-GRS | 0.0569 |

New Recovery Services vaults default to GRS. The workbook's price boxes default to GRS for the vault and Cool GRS for N2W (like for like); tab 3 lists your vaults' actual redundancy so you can switch both to LRS or ZRS; the script reads bytes per redundancy from each vault's `usages` endpoint (`LRSStorageUsage`, `GRSStorageUsage`, …) and prices each separately.

**Vault storage in the workbook is an estimate** (protected VMs' provisioned disk × *Used ÷ provisioned*). In testing it came out several times higher than the measured figure, which inflates the dollar saving. The workbook shows each vault's real figure on tab 2; for numbers anyone will quote, use the script or type the real figure.

## N2W side

| Item | Default | Note |
|---|---|---|
| License | $5 per VM per month, SQL included | Change it to your quoted price |
| Storage | Blob, General Block Blob v2, East US: Cool LRS 0.0152, Cool ZRS 0.019, Cool GRS 0.0334, Cold LRS 0.0036, Cold ZRS 0.004, Cold GRS 0.0081 (workbook offers all six) | Same GB as the vault, same redundancy family. Cool: 30-day minimum. Cold: 90-day minimum plus retrieval per GB |
| N2W server VM | $0 (left out) | Admin enters it; the tools say when it's excluded |

## Left out on both sides

Instant-restore snapshots (both products keep them), Log Analytics ingestion for Backup Reports, storage transactions, egress. Stopped-with-retained-data items aren't priced (conservative for Azure).

## Where the comparison is narrowest

- A tenant of mostly small VMs (≤50 GB) on LRS vaults is where the gap is narrowest: $5 instance fee against N2W's $5, and 0.0224 against 0.0152 on Cool.
- The **Cold** figure depends on 90-day-plus retention and rare restores. Both tools lead with Cool and show Cold in brackets.
- With *Used ÷ provisioned* at 100% the Azure fee is an upper bound. If your disks are half empty, lower it.

Prices retrieved 2 Nov 2026 from `prices.azure.com/api/retail/prices` (no authentication). Re-run `build/build_workbook.py` after changing `PRICES`.

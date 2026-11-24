#!/usr/bin/env python3
"""
Builds the N2W Azure Backup Estate Audit workbook.

Outputs:
  workbook/n2w-azure-backup-audit.workbook.json   gallery template (paste into Advanced editor)
  workbook/azuredeploy.json                       ARM template behind the Deploy-to-Azure button

Every price default lives in PRICES below. Change it there, re-run, commit both files.
Formulas are documented in docs/PRICING.md and mirrored in script/Invoke-N2WBackupAudit.ps1.
"""
import json, uuid, pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
NS = uuid.UUID("6f1c2a52-0d1e-4c55-9a51-2b6a3e0c0f17")
def gid(name): return str(uuid.uuid5(NS, name))

VERSION = "1.0.9"
PRICE_DATE = "2 Nov 2026"

# Azure Retail Prices API, East US, list (Consumption), retrieved 2026-11-02.
PRICES = {
    "PiUnit": 10,        # Backup / Azure VM Protected Instance, per unit per month (each 500 GB band)
    "PiSmall": 5,        # VMs up to 50 GB: half a unit (pricing page)
    "SqlFee": 25,        # Backup / SQL Server in Azure VM Protected Instance (up to 500 GB)
    "BackupStorage": {"LRS": 0.0224, "ZRS": 0.028, "GRS": 0.0448, "RA-GRS": 0.0569},   # Backup / Standard * Data Stored, $/GB-month
    "N2WStorage": {"Cool LRS": 0.0152, "Cool ZRS": 0.019, "Cool GRS": 0.0334, "Cold LRS": 0.0036, "Cold ZRS": 0.004, "Cold GRS": 0.0081},  # Blob (General Block Blob v2), $/GB-month
    "N2WLicense": 5,     # N2W flat per-VM per month, SQL included. CONFIRM against the locked pricing scenario.
}

ACTIVE_STATES = "'ProtectionConfigured','ConfiguringProtection','ProtectionError','ConfiguringProtectionFailed','IRPending'"
ITEMS = "microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems"

# ---------------------------------------------------------------- shared KQL
# One row per VM: provisioned disk GB, protection state, prod tag, fee band.
# Joins: 1 same-table (disks), 1 cross-table (RecoveryServicesResources). Within ARG limits.
VMP = f"""Resources
| where type =~ 'microsoft.compute/virtualmachines'
| extend vmId = tolower(id)
| extend isProd = tolower(tostring(tags)) contains tolower(strcat('"', '{{ProdTagKey}}', '":"', '{{ProdTagValue}}'))
| extend os = tostring(properties.storageProfile.osDisk.osType), vmSize = tostring(properties.hardwareProfile.vmSize)
| project vmId, id, name, subscriptionId, resourceGroup, location, os, vmSize, isProd
| join kind=leftouter (
    Resources
    | where type =~ 'microsoft.compute/disks' and isnotempty(managedBy)
    | summarize provGB = sum(toint(properties.diskSizeGB)) by vmId = tolower(tostring(managedBy))
  ) on vmId
| join kind=leftouter (
    RecoveryServicesResources
    | where type =~ '{ITEMS}'
    | where tostring(properties.backupManagementType) =~ 'AzureIaasVM'
    | where tostring(properties.currentProtectionState) in~ ({ACTIVE_STATES})
    | extend vmId = tolower(tostring(properties.sourceResourceId))
    | extend vaultPart = tostring(split(tolower(id), '/providers/microsoft.recoveryservices/vaults/')[1])
    | summarize vaults = make_set(tostring(split(vaultPart, '/')[0])), policies = make_set(tostring(properties.policyName)) by vmId
  ) on vmId
| extend provGB = iff(isnull(provGB), 0, provGB)
| extend protected = coalesce(array_length(vaults), 0) > 0
| extend vault = tostring(vaults[0]), policy = tostring(policies[0])
| extend usedGB = todouble(provGB) * {{UsedRatio}} / 100.0
| extend piUnits = iff(usedGB <= 50, 0.5, ceiling(usedGB / 500.0))
| extend piFee = iff(usedGB <= 50, todouble({{PiSmall}}), {{PiUnit}} * ceiling(usedGB / 500.0))
| extend band = case(usedGB <= 50, '1 · up to 50 GB', usedGB <= 500, '2 · 50–500 GB', usedGB <= 1000, '3 · 0.5–1 TB', usedGB <= 2000, '4 · 1–2 TB', '5 · over 2 TB')"""

SQLQ = f"""RecoveryServicesResources
| where type =~ '{ITEMS}'
| where tostring(properties.backupManagementType) =~ 'AzureWorkload' and tostring(properties.workloadType) =~ 'SQLDataBase'
| where tostring(properties.currentProtectionState) in~ ({ACTIVE_STATES})
| extend inst = strcat(tostring(split(tolower(id), '/protecteditems/')[0]), '|', tolower(tostring(properties.parentName)))
| summarize sqlInstances = dcount(inst), sqlDbs = count()"""

SQL_ESTATE = """Resources
| where type in~ ('microsoft.sqlvirtualmachine/sqlvirtualmachines', 'microsoft.sql/servers/databases')
| where not(type =~ 'microsoft.sql/servers/databases' and name =~ 'master')
| summarize sqlVms = countif(type =~ 'microsoft.sqlvirtualmachine/sqlvirtualmachines'), sqlDbPaas = countif(type =~ 'microsoft.sql/servers/databases')"""

DISKS = """Resources
| where type in~ ('microsoft.compute/disks', 'microsoft.compute/snapshots', 'microsoft.compute/restorepointcollections')
| extend isUn = type =~ 'microsoft.compute/disks' and tostring(properties.diskState) =~ 'Unattached'
| extend isSnap = type =~ 'microsoft.compute/snapshots'
| extend unGB = iff(isUn, toint(properties.diskSizeGB), 0)
| extend snapGB = iff(isSnap, toint(properties.diskSizeGB), 0)
| extend snapAt = iff(isSnap, todatetime(properties.timeCreated), datetime(null))
| summarize unattached = countif(isUn), unattachedGB = sum(unGB), snaps = countif(isSnap), snapGB = sum(snapGB), oldest = min(snapAt), rpc = countif(type =~ 'microsoft.compute/restorepointcollections')
| extend oldestDays = iff(isnull(oldest), 0, toint(datetime_diff('day', now(), oldest)))"""

VAULTS_RES = """Resources
| where type in~ ('microsoft.recoveryservices/vaults', 'microsoft.dataprotection/backupvaults')
    or (type =~ 'microsoft.logic/workflows' and tolower(tostring(tags)) contains '"usedbybackupreports":"true"')
| extend isRs = type =~ 'microsoft.recoveryservices/vaults', isBv = type =~ 'microsoft.dataprotection/backupvaults'
| extend isRpt = type =~ 'microsoft.logic/workflows'
| extend vloc = iff(isRpt, '', location)
| summarize rs = countif(isRs), bv = countif(isBv), rpt = countif(isRpt), locs = dcount(vloc)
| extend vaultRegions = locs - iff(rpt > 0, 1, 0)"""

RS_PLUMB = f"""RecoveryServicesResources
| where type in~ ('microsoft.recoveryservices/vaults/backuppolicies', 'microsoft.recoveryservices/vaults/backupjobs', '{ITEMS}')
| extend isPol = type =~ 'microsoft.recoveryservices/vaults/backuppolicies'
| extend polId = iff(isPol, tolower(id), '')
| join kind=leftouter (
    RecoveryServicesResources
    | where type =~ '{ITEMS}' and isnotempty(tostring(properties.policyId))
    | summarize usedBy = count() by polId = tolower(tostring(properties.policyId))
  ) on polId
| extend isDefault = isPol and name in~ ('DefaultPolicy', 'HourlyLogBackup', 'EnhancedPolicy')
| extend isUnused = isPol and isnull(usedBy) and not(isDefault)
| extend isUnusedDefault = isPol and isnull(usedBy) and isDefault
| extend isFailed = type =~ 'microsoft.recoveryservices/vaults/backupjobs' and tostring(properties.status) =~ 'Failed' and todatetime(properties.startTime) > ago(14d)
| extend isStopped = type =~ '{ITEMS}' and tostring(properties.currentProtectionState) =~ 'ProtectionStopped'
| summarize policies = countif(isPol), unusedPolicies = countif(isUnused), unusedDefaults = countif(isUnusedDefault), failedJobs14d = countif(isFailed), stoppedRetained = countif(isStopped)"""

DUP_GROUPS = """RecoveryServicesResources
| where type =~ 'microsoft.recoveryservices/vaults/backuppolicies'
| where tostring(properties.backupManagementType) =~ 'AzureIaasVM'
| extend p = properties
| extend vault = tostring(split(tostring(split(tolower(id), '/providers/microsoft.recoveryservices/vaults/')[1]), '/')[0])
| extend sig = strcat(tostring(p.schedulePolicy.scheduleRunFrequency), ' · every ', tostring(p.schedulePolicy.hourlySchedule['interval']), 'h · keep ',
    tostring(p.retentionPolicy.dailySchedule.retentionDuration['count']), 'd / ', tostring(p.retentionPolicy.weeklySchedule.retentionDuration['count']), 'w / ',
    tostring(p.retentionPolicy.monthlySchedule.retentionDuration['count']), 'm / ', tostring(p.retentionPolicy.yearlySchedule.retentionDuration['count']), 'y · instant restore ',
    tostring(p.instantRpRetentionRangeInDays), 'd')
| summarize policyCount = count(), vaultCount = dcount(vault), names = make_set(name) by sig
| where policyCount > 1"""

def N(name):
    """A parameter used as a number: 0 if it's empty, so an empty result never becomes a syntax error."""
    return f"coalesce(todouble('{{{name}}}'), 0.0)"

def dummy(tail):
    """ARG has no `print`; borrow one VM row so we can do arithmetic on parameters."""
    return "ResourceContainers\n| where type =~ 'microsoft.resources/subscriptions'\n| take 1\n" + tail

# ---------------------------------------------------------------- item helpers
def text(name, md, style=None):
    c = {"json": md}
    if style: c["style"] = style
    return {"type": 1, "content": c, "name": name}

def arg_query(name, query, viz="table", title=None, size=0, extra=None, nodata=None):
    c = {"version": "KqlItem/1.0", "query": query, "size": size, "queryType": 1,
         "resourceType": "microsoft.resourcegraph/resources", "crossComponentResources": ["{Subscription}"],
         "visualization": viz}
    if title: c["title"] = title
    if nodata: c["noDataMessage"] = nodata
    if extra: c.update(extra)
    return {"type": 3, "content": c, "name": name}

def tiles(name, query, title=None, money=False):
    num = {"unit": 0, "options": {"style": "decimal", "maximumFractionDigits": 0}}
    if money:
        num = {"unit": 0, "options": {"style": "decimal", "maximumFractionDigits": 0}}
    return arg_query(name, query, "tiles", title, 3, {"tileSettings": {
        "titleContent": {"columnMatch": "label", "formatter": 1},
        "leftContent": {"columnMatch": "value", "formatter": 12, "formatOptions": {"palette": "none"}, "numberFormat": num},
        "showBorder": True}})

def pack_tiles(rows):
    """rows: list of (label, expr). Produces label/value rows ordered as given."""
    packs = ",\n    ".join(f"pack('k', '{i:02d}', 'label', '{lbl}', 'value', todouble({expr}))" for i, (lbl, expr) in enumerate(rows))
    return f"\n| extend p = pack_array(\n    {packs})\n| mv-expand p\n| project k = tostring(p.k), label = tostring(p.label), value = todouble(p.value)\n| order by k asc"

def param(name, label, ptype, value=None, **kw):
    p = {"id": gid("p-" + name), "version": "KqlParameterItem/1.0", "name": name, "label": label, "type": ptype}
    if value is not None: p["value"] = value
    p.update(kw)
    return p

def qparam(name, label, query, hidden=True):
    """Text parameter whose value is the first cell of an ARG query."""
    return param(name, label, 1, query=query, queryType=1, resourceType="microsoft.resourcegraph/resources",
                 crossComponentResources=["{Subscription}"], isHiddenWhenLocked=hidden,
                 timeContext={"durationMs": 86400000})

def params_step(name, plist):
    return {"type": 9, "content": {"version": "KqlParameterItem/1.0", "parameters": plist, "style": "pills",
                                   "queryType": 1, "resourceType": "microsoft.resourcegraph/resources"}, "name": name}

def dropdown(name, label, options, default):
    data = json.dumps([{"value": str(v), "label": f"{k} · ${v}/GB"} for k, v in options.items()])
    return param(name, label, 2, value=str(default), jsonData=data, typeSettings={"showDefault": False})

REDUNDANCY_TOP = """Resources
| where type =~ 'microsoft.recoveryservices/vaults'
| extend r = tostring(properties.redundancySettings.standardTierStorageRedundancy)
| summarize g = countif(r =~ 'GeoRedundant'), z = countif(r =~ 'ZoneRedundant'), l = countif(r =~ 'LocallyRedundant')
| extend top = iff(g + z + l == 0, 'LRS', case(g >= l and g >= z, 'GRS', z >= l, 'ZRS', 'LRS'))"""

def smart_dropdown(name, label, options, pick):
    """Dropdown whose default follows the redundancy most of the tenant's vaults use. pick: KQL bool over k and top."""
    packs = ", ".join(f"pack('k', '{k}', 'v', '{v}')" for k, v in options.items())
    q = REDUNDANCY_TOP + f"\n| extend p = pack_array({packs})\n| mv-expand p\n| extend k = tostring(p.k)\n| project value = tostring(p.v), label = strcat(k, ' · $', tostring(p.v), '/GB'), selected = {pick}"
    return param(name, label, 2, query=q, queryType=1, resourceType="microsoft.resourcegraph/resources",
                 crossComponentResources=["{Subscription}"], typeSettings={"showDefault": False},
                 description="Defaults to the redundancy most of your vaults use; change it if yours differ.")

def group(name, items, tab=None, title=None, expandable=False, expanded=True):
    g = {"type": 12, "content": {"version": "NotebookGroup/1.0", "groupType": "editable", "items": items}, "name": name}
    if title: g["content"]["title"] = title
    if expandable:
        g["content"]["expandable"] = True
        g["content"]["expanded"] = expanded
    if tab:
        g["conditionalVisibility"] = {"parameterName": "Tab", "comparison": "isEqualTo", "value": tab}
    return g

RES = lambda col: {"columnMatch": col, "formatter": 13, "formatOptions": {"linkTarget": "Resource", "showIcon": True}}
BAR = lambda col, pal="blue": {"columnMatch": col, "formatter": 4, "formatOptions": {"palette": pal}}

# ---------------------------------------------------------------- top
header = text("header", f"""## N2W Azure Backup Estate Audit
Counts your VMs, what's backed up, your vaults and policies, and the reporting plumbing around them, then prices it at list.

**Read-only.** It runs as you, with Reader, using Azure Resource Graph queries you can read in the Advanced editor. It creates nothing and sends nothing outside your tenant. The only thing that ever leaves is the summary line on the last tab, and only if you paste it somewhere yourself.

<sub>v{VERSION} · prices: Azure list, East US, retrieved {PRICE_DATE} · source and license: github.com/N2W-Software/azure-backup-audit</sub>""")

p_top = params_step("scope", [
    param("Subscription", "Subscriptions", 6, value=["value::all"], isRequired=True, multiSelect=True, quote="'", delimiter=",",
          typeSettings={"additionalResourceOptions": ["value::all"], "includeAll": False, "showDefault": False}),
    param("ProdTagKey", "Prod tag key", 1, value="env", description="Tag key that marks production VMs (case-insensitive)"),
    param("ProdTagValue", "Prod tag value", 1, value="prod", description="Matches values that start with this, e.g. prod, production"),
])

P = PRICES
p_price = params_step("pricing", [
    param("UsedRatio", "Used ÷ provisioned disk, %", 1, value="100",
          description="Azure bills the instance fee on used data; Resource Graph only sees provisioned size. 100 = upper bound for the fee."),
    param("PiUnit", "Instance fee per 500 GB, $", 1, value=str(P["PiUnit"])),
    param("PiSmall", "Instance fee up to 50 GB, $", 1, value=str(P["PiSmall"])),
    param("SqlFee", "SQL-in-VM instance fee, $", 1, value=str(P["SqlFee"])),
    dropdown("BackupStorage", "Vault storage", P["BackupStorage"], P["BackupStorage"]["GRS"]),
    param("N2WLicense", "N2W per VM, $", 1, value=str(P["N2WLicense"])),
    dropdown("N2WStorage", "N2W storage tier", P["N2WStorage"], P["N2WStorage"]["Cool GRS"]),
    param("N2WServer", "N2W server VM, $/month", 1, value="0",
          description="The VM N2W runs on, in your subscription. 0 = left out; enter your size's monthly cost."),
    qparam("C_ProtVms", "Protected VMs", VMP + "\n| where protected\n| summarize n = count()\n| project v = tostring(n)"),
    qparam("C_PiFee", "Instance fees", VMP + "\n| where protected\n| summarize n = coalesce(round(sum(piFee), 0), 0.0)\n| project v = tostring(tolong(n))"),
    qparam("C_PiUnits", "Fee units", VMP + "\n| where protected\n| summarize n = coalesce(sum(piUnits), 0.0)\n| project v = iff(n == tolong(n), tostring(tolong(n)), tostring(n))"),
    qparam("C_ProtGB", "Protected GB (used, estimated)", VMP + "\n| where protected\n| summarize n = coalesce(round(sum(usedGB), 0), 0.0)\n| project v = tostring(tolong(n))"),
    qparam("C_Sql", "SQL instances", SQLQ + "\n| project v = tostring(coalesce(sqlInstances, 0))"),
])
VS = f"iff('{{VaultStorageTB}}' =~ 'auto', {N('C_ProtGB')} / 1024.0, {N('VaultStorageTB')})"
p_storage = params_step("storage", [
    param("VaultStorageTB", "Vault storage used, TB", 1, value="auto",
          description="auto = estimate from provisioned disk (often well off). Click each vault on tab 2, add up its storage, and type the total in TB."),
    qparam("C_VsTB", "Vault storage in use, TB", dummy(f"| project v = tostring(round({VS}, 2))")),
])

pricing_group = group("pricing-group", [
    text("pricing-note", f"""The price boxes above are the assumptions behind tab 3. Defaults are Azure list prices for East US on {PRICE_DATE}; type your negotiated rates and everything recalculates.

- **Instance fee:** up to 50 GB ${P['PiSmall']}; then ${P['PiUnit']} per started 500 GB. Billed on used data; this workbook sees provisioned size, so lower *Used ÷ provisioned* if your disks are half empty.
- **N2W storage tier:** Cool has a 30-day minimum, Cold a 90-day minimum plus per-GB retrieval. Compare like with like: GRS vaults against a GRS tier.
- **Vault storage** isn't in Resource Graph. The default is the protected VMs' provisioned disk × *Used ÷ provisioned*, which can be far off: in testing it was 5× the real figure. Click each vault on tab 2 to see its actual storage and type the total into *Vault storage used*. The script in the same repo reads it automatically."""),
], title="How the cost assumptions work", expandable=True, expanded=False)

tabs = {"type": 11, "content": {"version": "LinkItem/1.0", "style": "tabs", "links": [
    {"id": gid("tab-" + t), "cellValue": "Tab", "linkTarget": "parameter", "linkLabel": lbl, "subTarget": t, "style": "link"}
    for t, lbl in [("estate", "1 · Estate"), ("plumbing", "2 · Vaults and plumbing"), ("money", "3 · Cost"), ("export", "4 · Export summary")]]},
    "name": "tabs"}

# ---------------------------------------------------------------- tab 1 estate
estate = group("estate", [
    tiles("estate-tiles", VMP + """
| summarize vms = count(), prot = countif(protected), unprot = countif(not(protected)), unprotProd = countif(not(protected) and isProd)""" +
          pack_tiles([("Azure VMs", "vms"), ("Backed up by Azure Backup", "prot"), ("No backup", "unprot"), ("No backup, tagged prod", "unprotProd")])),
    tiles("estate-tiles-2", DISKS + pack_tiles([("Unattached disks", "unattached"), ("Unattached disk GB", "unattachedGB"),
                                                ("Standalone snapshots", "snaps"), ("Oldest snapshot, days", "oldestDays")])),
    tiles("estate-tiles-3", SQL_ESTATE + pack_tiles([("SQL Server VMs (registered)", "sqlVms"), ("Azure SQL databases (PaaS)", "sqlDbPaas")])),
    tiles("estate-tiles-4", SQLQ + pack_tiles([("SQL instances backed up in VMs", "sqlInstances"), ("SQL databases backed up", "sqlDbs")])),
    text("sql-note", "<sub>SQL Server VMs are counted when they're registered with the SQL IaaS Agent extension (the default for Azure Marketplace SQL images); SQL installed by hand on a plain VM won't show here.</sub>"),
    arg_query("coverage", VMP + """
| summarize Protected = countif(protected), Unprotected = countif(not(protected)) by subscriptionId
| extend Subscription = strcat('/subscriptions/', subscriptionId), Coverage = round(100.0 * Protected / (Protected + Unprotected), 0)
| project Subscription, Protected, Unprotected, Coverage
| order by Unprotected desc""", "table", "Backup coverage by subscription", 0,
              {"gridSettings": {"formatters": [RES("Subscription"), BAR("Coverage", "redGreen")],
                                "labelSettings": [{"columnId": "Coverage", "label": "Coverage, %"}]}}),
    arg_query("unprotected", VMP + """
| where not(protected)
| project VM = id, Prod = iff(isProd, 'yes', ''), Subscription = strcat('/subscriptions/', subscriptionId), ResourceGroup = resourceGroup, Location = location, OS = os, ProvisionedGB = provGB
| order by Prod desc, ProvisionedGB desc""", "table", "VMs with no Azure Backup (prod first)", 0,
              {"gridSettings": {"formatters": [RES("VM"), RES("Subscription")], "filter": True}},
              nodata="Every VM in scope has an active Azure Backup item."),
], tab="estate")

# ---------------------------------------------------------------- tab 2 plumbing
VAULT_ROWS_Q = f"""Resources
| where type in~ ('microsoft.recoveryservices/vaults', 'microsoft.dataprotection/backupvaults')
| extend vaultId = tolower(id)
| extend Kind = iff(type =~ 'microsoft.recoveryservices/vaults', 'Recovery Services', 'Backup vault')
| extend Redundancy = iff(type =~ 'microsoft.recoveryservices/vaults', tostring(properties.redundancySettings.standardTierStorageRedundancy), tostring(properties.storageSettings[0]['type']))
| join kind=leftouter (
    RecoveryServicesResources
    | where type in~ ('{ITEMS}', 'microsoft.dataprotection/backupvaults/backupinstances')
    | extend vaultId = iff(type =~ '{ITEMS}', tostring(split(tolower(id), '/backupfabrics/')[0]), tostring(split(tolower(id), '/backupinstances/')[0]))
    | summarize Items = count() by vaultId
  ) on vaultId
| project Vault = id, Kind, Location = location, Redundancy, Items = iff(isnull(Items), 0, Items), Subscription = strcat('/subscriptions/', subscriptionId)
| order by Location asc, Items desc"""

plumbing = group("plumbing", [
    text("plumbing-note", "Azure Backup needs one vault per region, a policy per vault, and a diagnostic setting per vault before Backup Reports can see it; emailed reports add a Logic App. This tab counts what that adds up to in your tenant."),
    tiles("plumb-tiles", VAULTS_RES + pack_tiles([("Recovery Services vaults", "rs"), ("Backup vaults", "bv"), ("Regions with a vault", "vaultRegions"),
                                                   ("Report Logic Apps (UsedByBackupReports)", "rpt")])),
    tiles("plumb-tiles-2", RS_PLUMB + pack_tiles([("Backup policies", "policies"), ("Your policies protecting nothing", "unusedPolicies"), ("Azure default policies, unused", "unusedDefaults"),
                                                  ("Failed backup jobs, last 14 days", "failedJobs14d"), ("Stopped, data still retained", "stoppedRetained")])),
    arg_query("vaults", VAULT_ROWS_Q, "table", "Vaults (click a row to see its storage and diagnostic settings)", 0,
              {"gridSettings": {"formatters": [RES("Vault"), RES("Subscription"), BAR("Items")], "filter": True},
               "exportFieldName": "Vault", "exportParameterName": "SelectedVault", "exportDefaultValue": ""}),
    {"type": 3, "name": "diag", "conditionalVisibility": {"parameterName": "SelectedVault", "comparison": "isNotEqualTo", "value": ""},
     "content": {"version": "KqlItem/1.0", "queryType": 12, "size": 4, "visualization": "table",
                 "title": "Diagnostic settings on the selected vault",
                 "noDataMessage": "No diagnostic settings: this vault sends nothing to Log Analytics, so Backup Reports can't see it.",
                 "query": json.dumps({"version": "ARMEndpoint/1.0", "data": None, "headers": [], "method": "GET",
                                      "path": "{SelectedVault}/providers/Microsoft.Insights/diagnosticSettings",
                                      "urlParams": [{"key": "api-version", "value": "2021-05-01-preview"}], "batchDisabled": False,
                                      "transformers": [{"type": "jsonpath", "settings": {"tablePath": "$.value", "columns": [
                                          {"path": "$.name", "columnid": "Setting"},
                                          {"path": "$.properties.workspaceId", "columnid": "Log Analytics workspace"},
                                          {"path": "$.properties.logAnalyticsDestinationType", "columnid": "Table type"},
                                          {"path": "$.properties.storageAccountId", "columnid": "Storage account"},
                                          {"path": "$.properties.eventHubAuthorizationRuleId", "columnid": "Event hub"}]}}]})}},
    {"type": 3, "name": "vault-usage", "conditionalVisibility": {"parameterName": "SelectedVault", "comparison": "isNotEqualTo", "value": ""},
     "content": {"version": "KqlItem/1.0", "queryType": 12, "size": 4, "visualization": "table",
                 "title": "Storage in the selected vault (add these up for Vault storage used)",
                 "noDataMessage": "No storage figures returned. Backup vaults (Data Protection) don't report storage here.",
                 "query": json.dumps({"version": "ARMEndpoint/1.0", "data": None, "headers": [], "method": "GET",
                                      "path": "{SelectedVault}/usages",
                                      "urlParams": [{"key": "api-version", "value": "2016-06-01"}], "batchDisabled": False,
                                      "transformers": [{"type": "jsonpath", "settings": {"tablePath": "$.value[?(@.unit == 'Bytes')]", "columns": [
                                          {"path": "$.name.localizedValue", "columnid": "Storage"},
                                          {"path": "$.currentValue", "columnid": "Size"}]}}]}),
                 "gridSettings": {"formatters": [{"columnMatch": "Size", "formatter": 0, "numberFormat": {"unit": 2, "options": {"style": "decimal", "maximumFractionDigits": 1}}}]}}},
    arg_query("dups", DUP_GROUPS + """
| project Schedule = sig, Policies = policyCount, Vaults = vaultCount, Names = strcat_array(names, ', ')
| order by Policies desc""", "table", "VM policies that repeat another policy's schedule and retention", 0,
              {"gridSettings": {"formatters": [BAR("Policies")]}},
              nodata="No repeated VM policies found."),
], tab="plumbing")

# ---------------------------------------------------------------- tab 3 money
AZ_STOR = f"{N('C_VsTB')} * 1024 * {N('BackupStorage')}"
N2_STOR = f"{N('C_VsTB')} * 1024 * {N('N2WStorage')}"
AZ_TOT = f"({N('C_PiFee')} + {N('C_Sql')} * {N('SqlFee')} + {AZ_STOR})"
N2_TOT = f"({N('C_ProtVms')} * {N('N2WLicense')} + {N2_STOR} + {N('N2WServer')})"

receipt_rows = [
    ("a1", "Azure Backup", "Protected-instance fees", "strcat('{C_ProtVms} VMs = {C_PiUnits} fee units × $', '{PiUnit}', ' (up to 50 GB counts as half)')", N("C_PiFee")),
    ("a2", "Azure Backup", "SQL Server in VM", "strcat('{C_Sql} instances × $', '{SqlFee}', ' (lowest tier)')", f"{N('C_Sql')} * {N('SqlFee')}"),
    ("a3", "Azure Backup", "Vault storage", "strcat('{C_VsTB} TB', iff('{VaultStorageTB}' =~ 'auto', ' (estimate)', ''), ' × $', '{BackupStorage}', '/GB')", AZ_STOR),
    ("a9", "Azure Backup", "Total per month", "'excludes instant-restore snapshots and Log Analytics ingestion'", AZ_TOT),
    ("n1", "N2W", "License, flat", "strcat('{C_ProtVms} VMs × $', '{N2WLicense}', ', SQL included')", f"{N('C_ProtVms')} * {N('N2WLicense')}"),
    ("n2", "N2W", "Storage in your Blob account", "strcat('{C_VsTB} TB × $', '{N2WStorage}', '/GB')", N2_STOR),
    ("n3", "N2W", "N2W server VM", "'as entered in Pricing'", N("N2WServer")),
    ("n9", "N2W", "Total per month", "'excludes snapshots and storage transactions, same as above'", N2_TOT),
]
receipt_q = dummy("| extend r = pack_array(\n    " + ",\n    ".join(
    f"pack('k', '{k}', 'side', '{side}', 'line', '{line}', 'math', {math}, 'usd', round(todouble({usd}), 0))"
    for k, side, line, math, usd in receipt_rows) +
    ")\n| mv-expand r\n| project k = tostring(r.k), Side = tostring(r.side), Line = tostring(r.line), Arithmetic = tostring(r.math), PerMonth = todouble(r.usd)\n| order by k asc\n| project-away k")

money = group("money", [
    dict(text("money-empty", "**No VMs are actively backed up by Azure Backup** in the selected subscriptions, so the instance fees below are $0 and any cost shown is storage still held in vaults. Tab 1 lists the unprotected VMs; to price protecting them on each side, use the N2W calculator with your VM count."),
         conditionalVisibility={"parameterName": "C_ProtVms", "comparison": "isEqualTo", "value": "0"}),
    text("money-note", "List-price estimate for what's backed up today, computed from the price boxes at the top. **Vault storage is an estimate until you type the real figure** (tab 2, click each vault). Every line shows its arithmetic. Type your own rates there and this tab recalculates."),
    tiles("money-tiles", dummy(f"| extend az = round({AZ_TOT}, 0), n2 = round({N2_TOT}, 0)" + pack_tiles([
        ("Azure Backup, per month", "az"), ("Same estate on N2W, per month", "n2"), ("Difference per month", "az - n2"),
        ("Difference, %", "iff(az > 0, round(100.0 * (az - n2) / az, 0), 0.0)")])), money=True),
    arg_query("receipt", receipt_q, "table", "The arithmetic", 0,
              {"gridSettings": {"formatters": [{"columnMatch": "PerMonth", "formatter": 0, "numberFormat": {"unit": 0, "options": {"style": "currency", "currency": "USD", "maximumFractionDigits": 0}}}],
                                "labelSettings": [{"columnId": "PerMonth", "label": "Per month"}]}}),
    arg_query("bands", VMP + """
| where protected
| summarize VMs = count(), FeeEach = max(piFee), FeeTotal = sum(piFee) by Band = band
| order by Band asc""", "table", "Instance fee by VM size (the size tax)", 1,
              {"gridSettings": {"formatters": [BAR("VMs")], "labelSettings": [{"columnId": "FeeEach", "label": "Fee per VM, $"}, {"columnId": "FeeTotal", "label": "Total, $"}]}}),
    arg_query("redundancy", """Resources
| where type =~ 'microsoft.recoveryservices/vaults'
| summarize Vaults = count() by Redundancy = tostring(properties.redundancySettings.standardTierStorageRedundancy)""",
              "table", "Your vaults' storage redundancy (pick the matching Vault storage price above)", 1),
    text("money-links", "Open this estate in the N2W calculator: [n2ws.com/azure-cost-savings-calculator](https://n2ws.com/azure-cost-savings-calculator?src=workbook). The link carries none of your numbers; type them in there.  \nN2W is listed on Azure Marketplace. Purchases there do not count toward an Azure consumption commitment (MACC)."),
], tab="money")

# ---------------------------------------------------------------- tab 4 export
E_VM = VMP + """
| summarize vms = count(), prot = countif(protected), unprot = countif(not(protected)), unprotProd = countif(not(protected) and isProd),
    regions = dcount(location), subs = dcount(subscriptionId),
    b1 = countif(protected and usedGB <= 50), b2 = countif(protected and usedGB > 50 and usedGB <= 500), b3 = countif(protected and usedGB > 500),
    units = sum(iff(protected, piUnits, 0.0))
| project s = strcat('subscriptions=', subs, ';regions=', regions, ';vms=', vms, ';vms_protected=', prot, ';vms_unprotected=', unprot,
    ';vms_unprotected_prod=', unprotProd, ';vm_band_le50=', b1, ';vm_band_50_500=', b2, ';vm_band_gt500=', b3, ';vm_fee_units=', units)"""
E_SQL = SQLQ + "\n| project s = strcat('sql_instances=', sqlInstances, ';sql_databases=', sqlDbs)"
E_SQE = SQL_ESTATE + "\n| project s = strcat('sql_vms_registered=', sqlVms, ';azure_sql_databases=', sqlDbPaas)"
E_RES = VAULTS_RES + "\n| project s = strcat('rs_vaults=', rs, ';backup_vaults=', bv, ';vault_regions=', vaultRegions, ';report_logic_apps=', rpt)"
E_RSP = RS_PLUMB + "\n| project s = strcat('policies=', policies, ';policies_unused=', unusedPolicies, ';policies_default_unused=', unusedDefaults, ';failed_jobs_14d=', failedJobs14d, ';stopped_retained=', stoppedRetained)"
E_DUP = DUP_GROUPS + "\n| summarize d = sum(policyCount)\n| project s = strcat('policies_repeated=', iff(isnull(d), 0, d))"
E_DSK = DISKS + "\n| project s = strcat('unattached_disks=', unattached, ';snapshots=', snaps, ';oldest_snapshot_days=', oldestDays)"
E_CST = dummy(f"""| extend az = round({AZ_TOT}, 0), n2 = round({N2_TOT}, 0)
| extend basis = iff('{{VaultStorageTB}}' =~ 'auto', 'estimate', 'entered')
| project s = strcat('vault_storage_gb=', tolong(round({N('C_VsTB')} * 1024, 0)), ';vault_storage_basis=', basis, ';used_ratio_pct={{UsedRatio}}',
    ';price_pi_unit={{PiUnit}};price_sql={{SqlFee}};price_vault_gb={{BackupStorage}};price_n2w_vm={{N2WLicense}};price_n2w_gb={{N2WStorage}};n2w_server={{N2WServer}}',
    ';azure_monthly_usd=', tolong(az), ';n2w_monthly_usd=', tolong(n2), ';date=', format_datetime(now(), 'yyyy-MM-dd'))""")

E_NAMES = ("E_Vm", "E_Sql", "E_Sqe", "E_Res", "E_Rsp", "E_Dup", "E_Dsk", "E_Cst")
export = group("export", [
    params_step("export-params", [
        qparam("E_Vm", "vm", E_VM), qparam("E_Sql", "sql", E_SQL), qparam("E_Sqe", "sqe", E_SQE), qparam("E_Res", "res", E_RES), qparam("E_Rsp", "rsp", E_RSP),
        qparam("E_Dup", "dup", E_DUP), qparam("E_Dsk", "dsk", E_DSK), qparam("E_Cst", "cst", E_CST),
    ]),
    text("export-text", """### Your summary
Counts and prices only: no VM names, no subscription IDs, no resource names. Read it before you share it."""),
    # The line is built by a query and shown with the text visualization. A text step that mixes query-derived
    # parameters into its markdown makes the portal show "may send data to an external URL" and UNTRUSTED_PARAM.
    arg_query("export-line", dummy("| project Summary = strcat('N2W-AUDIT/1;source=workbook;tool_version=" + VERSION + ";', " +
                                   ", ';', ".join("'{" + n + "}'" for n in E_NAMES) + ")"),
              "card", None, 3, {"textSettings": {"style": "plain"}}),
    text("export-after", "Select the whole line above (it starts with N2W-AUDIT/1 and ends with the date) and copy it. Paste it at **n2ws.com/azure-backup-audit** to get the itemised receipt by email, or to book a 20-minute review of your numbers with an N2W engineer. Nothing is sent unless you paste it."),
], tab="export")
# Show the summary only when every part is filled in, so nobody copies a line with gaps.
for _i in export["content"]["items"]:
    if _i["name"] in ("export-text", "export-line", "export-after"):
        _i["conditionalVisibilities"] = [{"parameterName": n, "comparison": "isNotEqualTo", "value": ""} for n in E_NAMES]
export["content"]["items"].insert(1, text("export-wait", "If no summary appears below after a few seconds, one of the counts couldn't be read. Click *Edit* at the top to see which value is unset, and send us a screenshot."))

workbook = {
    "version": "Notebook/1.0",
    "items": [header, p_top, p_price, p_storage, pricing_group, tabs, estate, plumbing, money, export],
    "fallbackResourceIds": ["Azure Monitor"],
    "$schema": "https://github.com/Microsoft/Application-Insights-Workbooks/blob/master/schema/workbook.json",
}

out = ROOT / "workbook" / "n2w-azure-backup-audit.workbook.json"
out.write_text(json.dumps(workbook, indent=2, ensure_ascii=False) + "\n")

arm = {
    "$schema": "https://schema.management.azure.com/schemas/2019-04-01/deploymentTemplate.json#",
    "contentVersion": "1.0.0.0",
    "metadata": {"description": "Deploys the N2W Azure Backup Estate Audit workbook (read-only queries). Creates one workbook resource; nothing else."},
    "parameters": {
        "workbookDisplayName": {"type": "string", "defaultValue": "N2W Azure Backup Estate Audit"},
        "workbookId": {"type": "string", "defaultValue": "[newGuid()]"},
        "location": {"type": "string", "defaultValue": "[resourceGroup().location]"},
    },
    "resources": [{
        "name": "[parameters('workbookId')]",
        "type": "microsoft.insights/workbooks",
        "location": "[parameters('location')]",
        "apiVersion": "2022-04-01",
        "kind": "shared",
        "tags": {"source": "n2w-azure-backup-audit", "version": VERSION},
        "properties": {
            "displayName": "[parameters('workbookDisplayName')]",
            "serializedData": json.dumps(workbook, ensure_ascii=False),
            "version": "Notebook/1.0",
            "sourceId": "Azure Monitor",
            "category": "workbook",
        },
    }],
    "outputs": {"workbookId": {"type": "string", "value": "[resourceId('microsoft.insights/workbooks', parameters('workbookId'))]"}},
}
(ROOT / "workbook" / "azuredeploy.json").write_text(json.dumps(arm, indent=2, ensure_ascii=False) + "\n")

# ---------------------------------------------------------------- script: same queries, generated
SCRIPT_QUERIES = {
    "vmp": VMP + "\n| project id, name, subscriptionId, resourceGroup, location, os, vmSize, isProd, provGB, protected, vault, policy",
    "sql": SQLQ, "sqlEstate": SQL_ESTATE, "disks": DISKS, "vaultsSummary": VAULTS_RES, "rsPlumb": RS_PLUMB,
    "dupGroups": DUP_GROUPS + "\n| project policyCount, vaultCount", "vaults": VAULT_ROWS_Q,
}
sp = ROOT / "script" / "Invoke-N2WBackupAudit.ps1"
src = sp.read_text()
a = src.index("# BEGIN GENERATED QUERIES"); a = src.index("\n", a) + 1
z = src.index("# END GENERATED QUERIES")
block = "$Q = @{}\n" + "".join(f"$Q['{k}'] = @'\n{v}\n'@\n" for k, v in SCRIPT_QUERIES.items())
assert "\n'@" not in "".join(SCRIPT_QUERIES.values())
sp.write_text(src[:a] + block + src[z:])
print("wrote", out.relative_to(ROOT), ", workbook/azuredeploy.json and the query block in", sp.relative_to(ROOT))

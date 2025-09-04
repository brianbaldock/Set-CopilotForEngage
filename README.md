# Set-EngageFeatureAccess

Manage **Viva Engage** feature access policies (e.g., **Copilot in Viva Engage** and **AI‑Powered Summarization**) with idempotent, automation‑friendly PowerShell.

> **Effective policy rules (per Microsoft):**  
> - A user or group can be in scope of **multiple** policies for the **same feature**.  
> - **Most restrictive wins** (Disabled overrides Enabled).  
> - **Direct user/group assignments** take precedence over **org‑wide** settings.  
> See: [Manage access policies in Viva](https://learn.microsoft.com/en-us/viva/manage-access-policies).

---

## Recommended model

- **Baseline:** one **org‑wide *Disabled*** policy *per feature* (conservative default).
- **Access:** one or more **Enable** policies targeted to specific **groups**.
- You **don’t need a “Disable” group** unless you choose a permissive org‑wide default and want carve‑outs.

### Suggested groups (examples)
- `GG-Engage-Copilot-Enabled`
- `GG-Engage-AISum-Enabled`

Keep memberships **mutually exclusive** per feature when possible; if a user lands in conflicting policies, *Disabled* wins by design.

---

## Quick start

> Requires: PowerShell 5.1+, `ExchangeOnlineManagement` 3.9.0+ (the module helper can auto‑install/update with switches).

```powershell
# 1) Baseline: Disable both features org‑wide
Set-EngageFeatureAccess -Mode Disable -Copilot -AISummarization `
  -Everyone -PolicyNamePrefix "All" -Confirm:$false -Verbose

# 2) Enable Copilot for one or more groups
Set-EngageFeatureAccess -Mode Enable -Copilot `
  -GroupIds "e25dc5ed-9ccf-4d04-bb06-1b77fda4e636" `
  -PolicyNamePrefix "Enable" -Confirm:$false -Verbose

# 3) Enable AI Summarization for one or more groups
Set-EngageFeatureAccess -Mode Enable -AISummarization `
  -GroupIds "c62c7ce6-5aef-4b67-b420-fa5deb75ecd6" `
  -PolicyNamePrefix "Enable" -Confirm:$false -Verbose

# 4) Verify policy layout
Get-VivaModuleFeaturePolicy -ModuleId VivaEngage
```

> Tip: Add `-AutoInstallEXO` / `-AutoUpdateEXO` to `Set-EngageFeatureAccess` if you want the helper to ensure the Exchange Online module is present/up‑to‑date.

---

## What “good” looks like

The view below shows **two org‑wide block policies** (one per feature) and **two group‑targeted enable policies**. With this layout, anyone **not** in an enable group stays blocked; group members are enabled.

![Policy layout screenshot](Images/policy-layout.png)

---

## Function reference

- **`Set-EngageFeatureAccess`** — Public: create/update policies for Copilot and/or AI Summarization (SupportsShouldProcess; non‑interactive).  
- **`Update-VivaPolicy`** — Internal: idempotent create/update for a single policy.  
- **`Resolve-VivaEngageFeatures`** — Internal: resolves feature IDs and caches results.  
- **`Get-EXOModule`, `Connect-EXOIfNeeded`** — Internal: module presence/connection helpers.

---

## Best practices

- Prefer **org‑wide Disabled** + **targeted Enable** groups for clarity and least‑privilege.
- Keep group memberships **mutually exclusive** per feature.
- Use `-PolicyNamePrefix` to make intent obvious (e.g., `All`, `Enable`).
- Periodically audit with:  
  ```powershell
  Get-VivaModuleFeaturePolicy -ModuleId VivaEngage |
    Sort-Object FeatureId, Name |
    Format-Table Name, FeatureId, IsFeatureEnabled, AccessControlList
  ```

---

## Troubleshooting

- **Module import/connect issues** → run with `-Verbose`; the module emits categorized errors
  (`ResourceUnavailable`, `ConnectionError`, etc.).  
- **Feature appears disabled unexpectedly for a user** → check for overlapping policies; remember
  *Disabled* beats *Enabled*, and **group/user assignment** beats **org‑wide**.

---

## References

- Microsoft Docs — **Manage access policies in Viva**:  
  https://learn.microsoft.com/en-us/viva/manage-access-policies

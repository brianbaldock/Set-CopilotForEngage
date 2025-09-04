# Set-EngageFeatureAccess

Check out the blog article here: [Admin Guide: Controlling Copilot in Viva Engage](https://blog.brianbaldock.net/admin-guide-controlling-copilot-in-viva-engage)

Manage **Viva Engage** feature access policies (e.g., **Copilot in Viva Engage** and **AI‑Powered Summarization**) with idempotent, automation‑friendly PowerShell.

> **Effective policy rules (per Microsoft):**  
>
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

> Requires: PowerShell 5.1+, `ExchangeOnlineManagement` 3.9.0+ (the module helper can auto‑install/update with switches) see step 2a

```powershell
# 1) Dot source the mini module
. .\Set-CopilotForEngage.ps1

# 2a) Baseline: Use this to disable both features org-wide and install/update the to the latest version of Exchange Online Managment PowerShell Module
Set-EngageFeatureAccess -Mode Disable -Copilot -AISummarization -Everyone -PolicyNamePrefix "All" -AutoInstallEXO -AutoUpdateEXO -Confirm:$false -Verbose

# 2b) Baseline: Disable both features org‑wide 
Set-EngageFeatureAccess -Mode Disable -Copilot -AISummarization `
  -Everyone -PolicyNamePrefix "All" -Confirm:$false -Verbose

# 2) Enable Copilot for one or more groups
Set-EngageFeatureAccess -Mode Enable -Copilot `
  -GroupIds "GROUP GUID HERE" `
  -PolicyNamePrefix "Enable" -Confirm:$false -Verbose

# 3) Enable AI Summarization for one or more groups
Set-EngageFeatureAccess -Mode Enable -AISummarization `
  -GroupIds "GROUP GUID HERE" `
  -PolicyNamePrefix "Enable" -Confirm:$false -Verbose

# 4) Verify policy layout
Get-VivaModuleFeaturePolicy -ModuleId VivaEngage
```

---

## What “good” looks like

The view below shows **two org‑wide block policies** (one per feature) and **two group‑targeted enable policies**. With this layout, anyone **not** in an enable group stays blocked; group members are enabled.

![Policy layout screenshot](Images/Policy%20Layout.png)

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
  **Failed to complete the request: Please confirm that you have the necessary permissions to manage <moduleId/featureId> and that the <moduleId/featureId> provided is valid.** → Validate that you have elevated privileges to run this script in your tenant

---

## References

- Microsoft Docs — **Manage access policies in Viva**:  
  <https://learn.microsoft.com/en-us/viva/manage-access-policies>

## Applies To

- Viva Engage and Copilot for Microsoft 365

## Author

|Author|Original Publish Date
|----|--------------------------
|Brian Baldock, Microsoft|September 4th, 2025

## Issues

Please report any issues you find to the [issues list](../../../../issues).

## Support Statement

The scripts, samples, and tools made available through the FastTrack Open Source initiative are provided as-is. These resources are developed in partnership with the community and do not represent official Microsoft software. As such, support is not available through premier or other Microsoft support channels. If you find an issue or have questions please reach out through the issues list and we'll do our best to assist, however there is no associated SLA.

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Legal Notices

Microsoft and any contributors grant you a license to the Microsoft documentation and other content in this repository under the [MIT License](https://opensource.org/licenses/MIT), see the [LICENSE](LICENSE) file, and grant you a license to any code in the repository under the [MIT License](https://opensource.org/licenses/MIT), see the [LICENSE-CODE](LICENSE-CODE) file.

Microsoft, Windows, Microsoft Azure and/or other Microsoft products and services referenced in the documentation may be either trademarks or registered trademarks of Microsoft in the United States and/or other countries. The licenses for this project do not grant you rights to use any Microsoft names, logos, or trademarks. Microsoft's general trademark guidelines can be found at <http://go.microsoft.com/fwlink/?LinkID=254653>.

Privacy information can be found at <https://privacy.microsoft.com/en-us/>

Microsoft and any contributors reserve all others rights, whether under their respective copyrights, patents,or trademarks, whether by implication, estoppel or otherwise.

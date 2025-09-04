#requires -Version 5.1
Set-StrictMode -Version Latest

#region Error helpers ---------------------------------------------------------
function New-TerminatingError {
    [CmdletBinding()] param(
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][System.Management.Automation.ErrorCategory]$Category,
        [string]$ErrorId = 'ModuleError',
        [object]$TargetObject
    )
    $ex = [System.Exception]::new($Message)
    $err = [System.Management.Automation.ErrorRecord]::new($ex, $ErrorId, $Category, $TargetObject)
    $PSCmdlet.ThrowTerminatingError($err)
}
#endregion

#region Utility ---------------------------------------------------------------
function Format-PolicyName {
    <#
    .SYNOPSIS
        Normalizes a policy name to allowed characters.
    .OUTPUTS
        [string]
    #>
    [CmdletBinding()] param([Parameter(Mandatory, ValueFromPipeline)][ValidateNotNullOrEmpty()][string]$Base)
    process {
        $name = $Base -replace '-', ' '
        $name = $name -replace '[^a-zA-Z0-9,. ]', ''
        $name = ($name -replace '\s+', ' ').Trim()
        return $name
    }
}

# Cache for feature lookup to avoid repeated service calls
$script:EngageFeatureCache = $null

function Get-EXOModule {
    <#
    .SYNOPSIS
        Ensures ExchangeOnlineManagement is installed, up-to-date, and imported.
    .DESCRIPTION
        Designed for non-interactive use. Prompts are disabled by default; use -AutoInstall / -AutoUpdate to allow actions.
    .PARAMETER MinVersion
        Minimum version to import. Default: 3.9.0.
    .PARAMETER AutoInstall
        Install module if missing.
    .PARAMETER AutoUpdate
        Update to a newer PSGallery version if available.
    .PARAMETER Repository
        PowerShellGet repository to use. Default PSGallery.
    .OUTPUTS
        [version]
    #>
    [CmdletBinding(SupportsShouldProcess)] param(
        [Version]$MinVersion = '3.9.0',
        [switch]$AutoInstall,
        [switch]$AutoUpdate,
        [ValidateNotNullOrEmpty()][string]$Repository = 'PSGallery'
    )

    $recommendedMin = [Version]'3.9.0'

    $installed = Get-Module ExchangeOnlineManagement -ListAvailable |
    Sort-Object Version -Descending | Select-Object -First 1

    $latest = $null
    try {
        $latest = Find-Module ExchangeOnlineManagement -Repository $Repository -ErrorAction Stop
    }
    catch {
        Write-Verbose "Could not query $Repository for ExchangeOnlineManagement: $_"
    }

    if (-not $installed) {
        if ($AutoInstall -and $PSCmdlet.ShouldProcess('ExchangeOnlineManagement', 'Install-Module')) {
            Install-Module ExchangeOnlineManagement -Repository $Repository -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            $installed = Get-Module ExchangeOnlineManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        }
        else {
            New-TerminatingError -Message 'ExchangeOnlineManagement is required but not installed. Rerun with -AutoInstall to install automatically.' -Category ResourceUnavailable -ErrorId 'EXO.NotInstalled'
        }
    }

    if ($latest -and ([Version]$installed.Version -lt [Version]$latest.Version)) {
        if ($AutoUpdate -and $PSCmdlet.ShouldProcess("ExchangeOnlineManagement $($installed.Version)", "Update-Module to $($latest.Version)")) {
            Update-Module ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction Stop
            $installed = Get-Module ExchangeOnlineManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        }
        else {
            Write-Warning "A newer ExchangeOnlineManagement $($latest.Version) is available (current: $($installed.Version))."
        }
    }

    if ([Version]$installed.Version -lt $recommendedMin) {
        Write-Warning ("ExchangeOnlineManagement {0} detected. For best compatibility, use {1}+." -f $installed.Version, $recommendedMin)
    }

    try {
        if ($MinVersion -and ([Version]$installed.Version -ge $MinVersion)) {
            Import-Module ExchangeOnlineManagement -MinimumVersion $MinVersion -ErrorAction Stop
        }
        else {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            if ($MinVersion -and ([Version]$installed.Version -lt $MinVersion)) {
                Write-Warning "Installed version $($installed.Version) is lower than requested MinimumVersion $MinVersion."
            }
        }
    }
    catch {
        New-TerminatingError -Message "Failed to import ExchangeOnlineManagement. $_" -Category ResourceUnavailable -ErrorId 'EXO.ImportFailed'
    }

    return [Version]$installed.Version
}

function Connect-EXOIfNeeded {
    [CmdletBinding()] param()
    try {
        if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
    }
    catch {
        New-TerminatingError -Message "Failed to connect to Exchange Online. $_" -Category ConnectionError -ErrorId 'EXO.ConnectFailed'
    }
}

function Resolve-VivaEngageFeatures {
    <#
    .SYNOPSIS
        Retrieves and caches Viva Engage feature descriptors.
    .OUTPUTS
        [hashtable]
    #>
    [CmdletBinding()] param()
    if ($script:EngageFeatureCache) { return $script:EngageFeatureCache }
    $all = Get-VivaModuleFeature -ModuleId VivaEngage -ErrorAction Stop
    $cache = @{
        CopilotInVivaEngage = $all | Where-Object { $_.Id -eq 'CopilotInVivaEngage' } | Select-Object -First 1
        AISummarization     = $all | Where-Object { $_.Id -eq 'AISummarization' } | Select-Object -First 1
    }
    $script:EngageFeatureCache = $cache
    return $cache
}

function Update-VivaPolicy {
    <#
    .SYNOPSIS
        Creates or updates a Viva Engage Feature Access policy (idempotent).
    .PARAMETER FeatureId
        Target feature Id.
    .PARAMETER PolicyName
        Policy display name.
    .PARAMETER IsEnabled
        Enables or disables the feature in the policy.
    .PARAMETER Everyone
        Apply tenant-wide.
    .PARAMETER GroupIds
        Apply to one or more groups (GUID or email).
    .PARAMETER UserIds
        Apply to one or more users (UPN).
    .PARAMETER UserOptInByDefault
        Sets IsUserOptedInByDefault when enabling user control.
    .OUTPUTS
        Policy object (selected properties)
    #>
    [OutputType([object])]
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$FeatureId,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$PolicyName,
        [Parameter(Mandatory)][bool]$IsEnabled,
        [switch]$Everyone,
        [string[]]$GroupIds,
        [string[]]$UserIds,
        [Nullable[bool]]$UserOptInByDefault
    )

    $PolicyName = Format-PolicyName $PolicyName

    # Pull once for the feature scope to reduce service calls
    $existingAll = Get-VivaModuleFeaturePolicy -ModuleId VivaEngage -FeatureId $FeatureId -ErrorAction Stop
    $existing = $existingAll | Where-Object { Test-PolicyNameMatch -InputObject $_ -PolicyName $PolicyName } | Select-Object -First 1

    $applyOptIn = $PSBoundParameters.ContainsKey('UserOptInByDefault') -and $IsEnabled
    if ($PSBoundParameters.ContainsKey('UserOptInByDefault') -and -not $IsEnabled) {
        Write-Warning 'Ignoring -UserOptInByDefault because IsEnabled is False.'
    }

    if ($existing) {
        if ($PSCmdlet.ShouldProcess("Update '$PolicyName'", "Update-VivaModuleFeaturePolicy IsFeatureEnabled:$IsEnabled")) {

            $setfeats = @{
                ModuleId         = 'VivaEngage'
                FeatureId        = $FeatureId
                PolicyId         = $existing.PolicyId
                IsFeatureEnabled = $IsEnabled
            }

            if ($PSBoundParameters.ContainsKey('UserOptInByDefault')) {
                $setfeats.IsUserControlEnabled = $true
                $setfeats.IsUserOptedInByDefault = $UserOptInByDefault
            }

            Update-VivaModuleFeaturePolicy @setfeats -Confirm:$false -ErrorAction Stop | Out-Null

            $updated = Get-VivaModuleFeaturePolicy -ModuleId VivaEngage -FeatureId $FeatureId -ErrorAction Stop |
            Where-Object { Test-PolicyNameMatch -InputObject $_ -PolicyName $PolicyName } |
            Select-Object -First 1
            return $updated
        }
        return $existing
    }


    $feats = @{ ModuleId = 'VivaEngage'; FeatureId = $FeatureId; Name = $PolicyName; IsFeatureEnabled = $IsEnabled }
    if ($Everyone) { $feats.Everyone = $true }
    if ($GroupIds) { $feats.GroupIds = $GroupIds }
    if ($UserIds) { $feats.UserIds = $UserIds }
    if ($applyOptIn) { $feats.IsUserOptedInByDefault = $UserOptInByDefault }

    if ($PSCmdlet.ShouldProcess("Create policy '$PolicyName' (FeatureId=$FeatureId, IsEnabled=$IsEnabled)")) {
        try {
            Add-VivaModuleFeaturePolicy @feats -Confirm:$false -ErrorAction Stop | Out-Null
        }
        catch {
            if ($_.Exception.Message -match 'already a tenant level policy') {
                Write-Verbose "Policy already exists at tenant scope; returning existing."
            }
            else {
                throw
            }
        }
        return (Get-VivaModuleFeaturePolicy -ModuleId VivaEngage -FeatureId $FeatureId -ErrorAction Stop |
            Where-Object { Test-PolicyNameMatch -InputObject $_ -PolicyName $PolicyName } |
            Select-Object -First 1)
    }
}

function Set-EngageFeatureAccess {
    <#
    .SYNOPSIS
        Enable or disable Copilot in Viva Engage and/or AI-Powered Summarization via Feature Access policies.
    .DESCRIPTION
        Creates or updates policies (idempotent) for the selected features and scope.
    .PARAMETER Mode
        'Disable' or 'Enable'. Default: Disable.
    .PARAMETER Copilot
        Target the "Copilot in Viva Engage" feature.
    .PARAMETER AISummarization
        Target the "AI-Powered Summarization" feature.
    .PARAMETER PolicyNamePrefix
        Prefix for generated policy names (default 'Engage').
    .PARAMETER Everyone
        Apply tenant-wide. (Parameter set: Everyone)
    .PARAMETER GroupIds
        One or more group emails or GUIDs. (Parameter set: Groups)
    .PARAMETER UserIds
        One or more user UPNs. (Parameter set: Users)
    .PARAMETER UserOptInByDefault
        Set IsUserOptedInByDefault when enabling user control.
    .PARAMETER AutoInstallEXO
        If specified, will install the ExchangeOnlineManagement module if missing.
    .PARAMETER AutoUpdateEXO
        If specified, will update the ExchangeOnlineManagement module if newer exists.
    .OUTPUTS
        A compact object per policy reflecting the resulting state.
    #>
    [OutputType([object])]
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Everyone')]
    param(
        [ValidateSet('Disable', 'Enable')][string]$Mode = 'Disable',

        [Parameter()][switch]$Copilot,
        [Parameter()][switch]$AISummarization,

        [string]$PolicyNamePrefix = 'Engage',

        # Scope parameter sets enforce mutual exclusivity without manual counting
        [Parameter(ParameterSetName = 'Everyone', Mandatory)][switch]$Everyone,
        [Parameter(ParameterSetName = 'Groups', Mandatory)][ValidateNotNullOrEmpty()][string[]]$GroupIds,
        [Parameter(ParameterSetName = 'Users', Mandatory)][ValidateNotNullOrEmpty()][string[]]$UserIds,

        [Nullable[bool]]$UserOptInByDefault,

        [switch]$AutoInstallEXO,
        [switch]$AutoUpdateEXO
    )

    if (-not ($Copilot -or $AISummarization)) {
        New-TerminatingError -Message 'Select at least one feature: -Copilot and/or -AISummarization.' -Category InvalidArgument -ErrorId 'Param.MissingFeature'
    }

    $exoVersion = Get-EXOModule -MinVersion '3.9.0' -AutoInstall:$AutoInstallEXO -AutoUpdate:$AutoUpdateEXO
    Write-Verbose ("ExchangeOnlineManagement loaded: {0}" -f $exoVersion)
    Connect-EXOIfNeeded

    $features = Resolve-VivaEngageFeatures
    $results = @()
    $isEnabled = ($Mode -eq 'Enable')

    if ($PSCmdlet.ShouldProcess('Set Engage Feature Access', "Mode: $Mode, Copilot: $Copilot, AISummarization: $AISummarization, Scope: $($PSCmdlet.ParameterSetName)")) {
        if ($Copilot) {
            if (-not $features.CopilotInVivaEngage) {
                New-TerminatingError -Message 'Feature not found: CopilotInVivaEngage.' -Category ObjectNotFound -ErrorId 'Feature.NotFound' -TargetObject 'CopilotInVivaEngage'
            }
            $name = Format-PolicyName "$PolicyNamePrefix, Copilot in Viva Engage"
            $feats = @{ FeatureId = $features.CopilotInVivaEngage.Id; PolicyName = $name; IsEnabled = $isEnabled; Everyone = $Everyone; GroupIds = $GroupIds; UserIds = $UserIds }
            if ($PSBoundParameters.ContainsKey('UserOptInByDefault')) { $feats.UserOptInByDefault = $UserOptInByDefault }
            $results += Update-VivaPolicy @feats
        }

        if ($AISummarization) {
            if (-not $features.AISummarization) {
                New-TerminatingError -Message 'Feature not found: AISummarization.' -Category ObjectNotFound -ErrorId 'Feature.NotFound' -TargetObject 'AISummarization'
            }
            $name = Format-PolicyName "$PolicyNamePrefix, AI Powered Summarization"
            $feats = @{ FeatureId = $features.AISummarization.Id; PolicyName = $name; IsEnabled = $isEnabled; Everyone = $Everyone; GroupIds = $GroupIds; UserIds = $UserIds }
            if ($PSBoundParameters.ContainsKey('UserOptInByDefault')) { $feats.UserOptInByDefault = $UserOptInByDefault }
            $results += Update-VivaPolicy @feats
        }
    }

    $results | Select-Object Name, FeatureId, IsFeatureEnabled, IsUserOptedInByDefault,
    @{ Name = 'Access'; Expression = { ($_.AccessControlList) -join ', ' } },
    PolicyId
}

function Test-PolicyNameMatch {
    [CmdletBinding()] param(
        [Parameter(Mandatory)][object]$InputObject,
        [Parameter(Mandatory)][string]$PolicyName
    )
    foreach ($prop in 'Name', 'DisplayName', 'PolicyName') {
        $p = $InputObject.PSObject.Properties[$prop]
        if ($p -and $p.Value -eq $PolicyName) { return $true }
    }
    return $false
}

# Quiet banner unless -InformationAction Continue
if ($MyInvocation.InvocationName -ne '.') {
    Write-Information "Loaded: Set-EngageFeatureAccess`nExample: Set-EngageFeatureAccess -Mode Disable -Copilot -AISummarization -Everyone -PolicyNamePrefix 'Global'" -InformationAction SilentlyContinue
}

param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteTitle,

    [Parameter(Mandatory)]
    [ValidatePattern('^[a-zA-Z0-9-]+$')]
    [string]$SiteAlias,

    [Parameter(Mandatory)]
    [Guid]$TenantId,

    [Parameter(Mandatory)]
    [Guid]$ClientId,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Thumbprint,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Owner,

    [Parameter(Mandatory)]
    [Guid]$SensitivityLabelId,

    [string]$Description = "",

    [string]$SourceSiteUrl =
        "https://securityconsultinggroup.sharepoint.com/sites/DMSAdministratorSiteTemplateOnly",

    [string]$LibraryTemplatePath =
        "C:\Users\YogeshVipradas\Documents\Scripts\SCG\Site-Libraries.xml",

    [string]$LibraryManifestPath =
        "C:\Users\YogeshVipradas\Documents\Scripts\SCG\Site-Libraries.manifest.json",

    [string[]]$RequiredHubContentTypeIds = @(
        "0x010100AFD7414EA532944AB25D361F1EB9E22E0201"
    ),

    [bool]$SyncAllCompatibleHubContentTypes = $true,

    [bool]$ExcludeDocumentSetContentTypes = $true,

    [bool]$EnableTaxonomy = $true,

    [ValidateSet("Blank", "Topic", "Showcase")]
    [string]$CommunicationSiteDesign = "Blank",

    [ValidateSet("Auto", "TopNavigationBar", "QuickLaunch")]
    [string]$SourceNavigationLocation = "Auto",

    [ValidateSet("TopNavigationBar", "QuickLaunch")]
    [string]$NavigationLocation = "TopNavigationBar"
)

$ErrorActionPreference = "Stop"

# ============================================================
# Configuration
# ============================================================

$AdminSiteUrl =
    "https://securityconsultinggroup-admin.sharepoint.com"

$TenantRootUrl =
    "https://securityconsultinggroup.sharepoint.com"

$TargetSiteUrl =
    "$TenantRootUrl/sites/$SiteAlias"

$OutputFolder =
    "C:\Users\YogeshVipradas\Documents\Scripts\SCG"

$PageTemplate = Join-Path `
    -Path $OutputFolder `
    -ChildPath "Communication-HomePage.pnp"

$NavigationTemplate = Join-Path `
    -Path $OutputFolder `
    -ChildPath "Communication-Navigation.pnp"

$LibraryImportScript = Join-Path `
    -Path $PSScriptRoot `
    -ChildPath "Import-SiteLibraries.ps1"

$Thumbprint =
    $Thumbprint.Replace(" ", "").ToUpperInvariant()

$Owner = $Owner.Trim()

# ============================================================
# Helper functions
# ============================================================

function Get-DmsDetailedError {
    param(
        [Parameter(Mandatory)]
        $ErrorRecord
    )

    $detail = [string]$ErrorRecord.Exception.Message

    if (
        $ErrorRecord.ErrorDetails -and
        -not [string]::IsNullOrWhiteSpace(
            [string]$ErrorRecord.ErrorDetails.Message
        )
    ) {
        $detail = [string]$ErrorRecord.ErrorDetails.Message
    }

    if ($ErrorRecord.Exception.InnerException) {
        $detail =
            "$detail | Inner: $($ErrorRecord.Exception.InnerException.Message)"
    }

    return $detail
}

function Get-DmsContentTypeId {
    param(
        [Parameter(Mandatory)]
        $ContentType
    )

    if (
        $ContentType.Id -and
        $ContentType.Id.PSObject.Properties.Name -contains "StringValue"
    ) {
        return [string]$ContentType.Id.StringValue
    }

    if (
        $ContentType.PSObject.Properties.Name -contains "ContentTypeId"
    ) {
        return [string]$ContentType.ContentTypeId
    }

    return [string]$ContentType.Id
}



function Get-DmsNavigationNodes {
    param(
        [Parameter(Mandatory)]
        [ValidateSet("TopNavigationBar", "QuickLaunch")]
        [string]$Location,

        [Parameter(Mandatory)]
        $Connection
    )

    $nodes = @(
        Get-PnPNavigationNode `
            -Location $Location `
            -Tree `
            -Connection $Connection `
            -ErrorAction SilentlyContinue |
            Where-Object {
                $null -ne $_ -and
                -not [string]::IsNullOrWhiteSpace([string]$_.Title)
            }
    )

    if ($nodes.Count -gt 0) {
        return $nodes
    }

    # Some sites return no objects when -Tree is used. Retry without
    # -Tree and then request each node individually with its children.
    $flatNodes = @(
        Get-PnPNavigationNode `
            -Location $Location `
            -Connection $Connection `
            -ErrorAction SilentlyContinue |
            Where-Object {
                $null -ne $_ -and
                -not [string]::IsNullOrWhiteSpace([string]$_.Title)
            }
    )

    if ($flatNodes.Count -eq 0) {
        return @()
    }

    $treeNodes = @(
        foreach ($flatNode in $flatNodes) {
            $treeNode = Get-PnPNavigationNode `
                -Id $flatNode.Id `
                -Tree `
                -Connection $Connection `
                -ErrorAction SilentlyContinue

            if ($treeNode) {
                $treeNode
            }
            else {
                $flatNode
            }
        }
    )

    return $treeNodes
}

function Convert-DmsNavigationUrl {
    param(
        [string]$Url,

        [Parameter(Mandatory)]
        [string]$SourceSiteUrl,

        [Parameter(Mandatory)]
        [string]$TargetSiteUrl
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return ""
    }

    if (
        $Url -ieq "http://linkless.header/" -or
        $Url -ieq "https://linkless.header/"
    ) {
        return ""
    }

    $sourceBase = $SourceSiteUrl.TrimEnd("/")
    $targetBase = $TargetSiteUrl.TrimEnd("/")

    if ($Url.StartsWith($sourceBase, [StringComparison]::OrdinalIgnoreCase)) {
        return $targetBase + $Url.Substring($sourceBase.Length)
    }

    try {
        $sourceUri = [Uri]$sourceBase
        $targetUri = [Uri]$targetBase

        if ($Url.StartsWith("/")) {
            $sourcePath = $sourceUri.AbsolutePath.TrimEnd("/")
            $targetPath = $targetUri.AbsolutePath.TrimEnd("/")

            if (
                $Url.StartsWith(
                    $sourcePath,
                    [StringComparison]::OrdinalIgnoreCase
                )
            ) {
                return $targetPath + $Url.Substring($sourcePath.Length)
            }
        }
    }
    catch {
        # Return the original URL when it cannot be parsed.
    }

    return $Url
}

function Copy-DmsNavigationTree {
    param(
        [Parameter(Mandatory)]
        [object[]]$Nodes,

        [Parameter(Mandatory)]
        [ValidateSet("TopNavigationBar", "QuickLaunch")]
        [string]$TargetLocation,

        [Parameter(Mandatory)]
        $TargetConnection,

        [Parameter(Mandatory)]
        [string]$SourceSiteUrl,

        [Parameter(Mandatory)]
        [string]$TargetSiteUrl,

        $Parent
    )

    $tenantHost =
        ([Uri]$TargetSiteUrl).Host

    foreach ($node in $Nodes) {
        if (
            $null -eq $node -or
            [string]::IsNullOrWhiteSpace([string]$node.Title)
        ) {
            continue
        }

        $mappedUrl = Convert-DmsNavigationUrl `
            -Url ([string]$node.Url) `
            -SourceSiteUrl $SourceSiteUrl `
            -TargetSiteUrl $TargetSiteUrl

        $parameters = @{
            Location    = $TargetLocation
            Title       = [string]$node.Title
            Connection  = $TargetConnection
            ErrorAction = "Stop"
        }

        if (-not [string]::IsNullOrWhiteSpace($mappedUrl)) {
            $parameters["Url"] = $mappedUrl

            try {
                $mappedUri = [Uri]$mappedUrl

                if (
                    $mappedUri.IsAbsoluteUri -and
                    $mappedUri.Host -ine $tenantHost
                ) {
                    $parameters["External"] = $true
                }
            }
            catch {
                # Relative URLs do not need the External switch.
            }
        }

        if ($null -ne $Parent) {
            $parameters["Parent"] = $Parent
        }

        Write-Host (
            "Adding navigation node: {0} -> {1}" -f
            $node.Title,
            $mappedUrl
        )

        $createdNode =
            Add-PnPNavigationNode @parameters

        $children = @($node.Children)

        if ($children.Count -gt 0) {
            Copy-DmsNavigationTree `
                -Nodes $children `
                -TargetLocation $TargetLocation `
                -TargetConnection $TargetConnection `
                -SourceSiteUrl $SourceSiteUrl `
                -TargetSiteUrl $TargetSiteUrl `
                -Parent $createdNode
        }
    }
}

function Test-DmsSiteExists {
    param(
        [Parameter(Mandatory)]
        $AdminConnection,

        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    try {
        $site = Get-PnPTenantSite `
            -Identity $SiteUrl `
            -Connection $AdminConnection `
            -ErrorAction Stop

        return ($null -ne $site)
    }
    catch {
        $detail = Get-DmsDetailedError -ErrorRecord $_

        if (
            $detail -match "does not exist" -or
            $detail -match "cannot be found" -or
            $detail -match "Cannot get site" -or
            $detail -match "File Not Found" -or
            $detail -match "404"
        ) {
            return $false
        }

        throw "Could not determine whether site '$SiteUrl' exists. $detail"
    }
}

function Test-DmsDeletedSiteExists {
    param(
        [Parameter(Mandatory)]
        $AdminConnection,

        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    try {
        $deletedSite = Get-PnPTenantDeletedSite `
            -Identity $SiteUrl `
            -Connection $AdminConnection `
            -ErrorAction Stop

        return ($null -ne $deletedSite)
    }
    catch {
        $detail = Get-DmsDetailedError -ErrorRecord $_

        if (
            $detail -match "does not exist" -or
            $detail -match "cannot be found" -or
            $detail -match "Cannot get site" -or
            $detail -match "Unable to find the deleted site" -or
            $detail -match "File Not Found" -or
            $detail -match "404"
        ) {
            return $false
        }

        throw "Could not check the deleted-sites recycle bin. $detail"
    }
}

function Enable-DmsSiteFeature {
    param(
        [Parameter(Mandatory)]
        $AdminConnection,

        [Parameter(Mandatory)]
        $TargetConnection,

        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [Guid]$FeatureId,

        [Parameter(Mandatory)]
        [string]$FeatureName,

        [int]$MaximumAttempts = 30,

        [int]$DelaySeconds = 10,

        [int]$StabilisationSeconds = 60
    )

    Write-Host ""
    Write-Host "Checking $FeatureName feature..."

    $enabled = Get-PnPFeature `
        -Identity $FeatureId `
        -Scope Site `
        -Connection $TargetConnection `
        -ErrorAction SilentlyContinue

    if (-not $enabled) {
        Write-Host (
            "Activating $FeatureName through a server-side " +
            "SharePoint site script..."
        )

        $siteScript = [ordered]@{
            '$schema' =
                "https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json"

            actions = @(
                [ordered]@{
                    verb      = "activateSPFeature"
                    featureId = $FeatureId.ToString()
                    scope     = "site"
                }
            )

            version = 1
        } | ConvertTo-Json -Depth 10

        try {
            Invoke-PnPSiteScript `
                -Script $siteScript `
                -WebUrl $SiteUrl `
                -Connection $AdminConnection `
                -ErrorAction Stop |
                Out-Host
        }
        catch {
            Write-Warning (
                "The server-side feature request returned: {0}" -f
                $_.Exception.Message
            )
        }
    }

    for (
        $attempt = 1;
        $attempt -le $MaximumAttempts;
        $attempt++
    ) {
        $enabled = Get-PnPFeature `
            -Identity $FeatureId `
            -Scope Site `
            -Connection $TargetConnection `
            -ErrorAction SilentlyContinue

        if ($enabled) {
            break
        }

        Write-Host (
            "{0} activation validation attempt {1} of {2}..." -f
            $FeatureName,
            $attempt,
            $MaximumAttempts
        )

        if ($attempt -lt $MaximumAttempts) {
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    if (-not $enabled) {
        Write-Host (
            "$FeatureName is not active after the site script. " +
            "Trying the PnP feature command..."
        )

        try {
            Enable-PnPFeature `
                -Identity $FeatureId `
                -Scope Site `
                -Force `
                -Connection $TargetConnection `
                -ErrorAction Stop |
                Out-Null
        }
        catch {
            Write-Warning (
                "PnP feature activation returned: {0}" -f
                $_.Exception.Message
            )
        }

        $enabled = Get-PnPFeature `
            -Identity $FeatureId `
            -Scope Site `
            -Connection $TargetConnection `
            -ErrorAction SilentlyContinue
    }

    if (-not $enabled) {
        Write-Host (
            "$FeatureName is not active after PnP. " +
            "Trying the SharePoint REST feature endpoint..."
        )

        try {
            Invoke-PnPSPRestMethod `
                -Method Post `
                -Url "/_api/site/features/add('$FeatureId')" `
                -Connection $TargetConnection `
                -ErrorAction Stop |
                Out-Null
        }
        catch {
            $detail = Get-DmsDetailedError -ErrorRecord $_

            throw @"
$FeatureName could not be enabled through the site script,
PnP feature command or SharePoint REST endpoint.

Target site:
$SiteUrl

Feature ID:
$FeatureId

Error:
$detail
"@
        }

        for ($attempt = 1; $attempt -le 12; $attempt++) {
            $enabled = Get-PnPFeature `
                -Identity $FeatureId `
                -Scope Site `
                -Connection $TargetConnection `
                -ErrorAction SilentlyContinue

            if ($enabled) {
                break
            }

            if ($attempt -lt 12) {
                Start-Sleep -Seconds 5
            }
        }
    }

    if (-not $enabled) {
        throw @"
$FeatureName is still not enabled on the target site.

Target site:
$SiteUrl

Feature ID:
$FeatureId

Activate it manually from:
Site settings > Site collection features > $FeatureName
"@
    }

    Write-Host "$FeatureName feature is enabled."

    if ($StabilisationSeconds -gt 0) {
        Write-Host (
            "Waiting {0} seconds for {1} to propagate..." -f
            $StabilisationSeconds,
            $FeatureName
        )

        Start-Sleep -Seconds $StabilisationSeconds
    }
}



function New-DmsLibraryTemplateWithoutDocumentSets {
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [string]$DestinationPath
    )

    $documentSetPrefix =
        "0x0120D520"

    [xml]$templateXml = Get-Content `
        -LiteralPath $SourcePath `
        -Raw `
        -Encoding UTF8

    $nodes = @(
        $templateXml.SelectNodes("//*")
    )

    $nodesToRemove = @()

    foreach ($node in $nodes) {
        if (-not $node.Attributes) {
            continue
        }

        foreach ($attribute in @($node.Attributes)) {
            $attributeValue =
                [string]$attribute.Value

            if (
                -not [string]::IsNullOrWhiteSpace($attributeValue) -and
                $attributeValue.IndexOf(
                    $documentSetPrefix,
                    [StringComparison]::OrdinalIgnoreCase
                ) -ge 0
            ) {
                $nodesToRemove += $node
                break
            }
        }
    }

    $nodesToRemove = @(
        $nodesToRemove |
            Select-Object -Unique
    )

    foreach ($node in $nodesToRemove) {
        if ($node.ParentNode) {
            $null =
                $node.ParentNode.RemoveChild($node)
        }
    }

    $settings = New-Object `
        -TypeName System.Xml.XmlWriterSettings

    $settings.Indent = $true
    $settings.Encoding =
        New-Object System.Text.UTF8Encoding($false)

    $writer = [System.Xml.XmlWriter]::Create(
        $DestinationPath,
        $settings
    )

    try {
        $templateXml.Save($writer)
    }
    finally {
        $writer.Dispose()
    }

    Write-Host (
        "Created library template without Document Set references."
    )
    Write-Host "Removed XML nodes: $($nodesToRemove.Count)"
    Write-Host "Filtered template: $DestinationPath"

    return $DestinationPath
}

function Get-DmsCompatibleHubContentTypeIds {
    param(
        [Parameter(Mandatory)]
        $Connection,

        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    Write-Host ""
    Write-Host "Retrieving compatible published hub content types..."

    $compatible = @(
        Get-PnPCompatibleHubContentTypes `
            -WebUrl $SiteUrl `
            -Connection $Connection `
            -ErrorAction Stop
    )

    $records = @(
        foreach ($contentType in $compatible) {
            $id = Get-DmsContentTypeId `
                -ContentType $contentType

            if (-not [string]::IsNullOrWhiteSpace($id)) {
                [PSCustomObject]@{
                    Name          = [string]$contentType.Name
                    Group         = [string]$contentType.Group
                    ContentTypeId = $id.Trim()
                }
            }
        }
    )

    if ($records.Count -gt 0) {
        $records |
            Sort-Object {
                $_.ContentTypeId.Length
            }, Name |
            Format-Table Name, Group, ContentTypeId -AutoSize |
            Out-Host
    }

    return @(
        $records |
            Select-Object -ExpandProperty ContentTypeId |
            Sort-Object Length |
            Select-Object -Unique
    )
}

function Confirm-DmsHubContentTypes {
    param(
        [Parameter(Mandatory)]
        $Connection,

        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [string[]]$ContentTypeIds,

        [int]$MaximumSyncAttempts = 20,

        [int]$SyncDelaySeconds = 60,

        [int]$MaximumValidationAttempts = 30,

        [int]$ValidationDelaySeconds = 10
    )

    $requiredIds = @(
        $ContentTypeIds |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace([string]$_)
            } |
            ForEach-Object {
                ([string]$_).Trim()
            } |
            Sort-Object Length |
            Select-Object -Unique
    )

    if ($requiredIds.Count -eq 0) {
        throw "No content type IDs were selected."
    }

    Write-Host ""
    Write-Host "Synchronising content types from the hub..."
    Write-Host "Target site: $SiteUrl"

    $syncSucceeded = $false
    $lastFailureText = ""

    for (
        $syncAttempt = 1;
        $syncAttempt -le $MaximumSyncAttempts;
        $syncAttempt++
    ) {
        Write-Host (
            "Content Type Hub sync attempt {0} of {1}..." -f
            $syncAttempt,
            $MaximumSyncAttempts
        )

        try {
            $syncResults = @(
                Add-PnPContentTypesFromContentTypeHub `
                    -ContentTypes $requiredIds `
                    -Site $SiteUrl `
                    -Connection $Connection `
                    -ErrorAction Stop
            )
        }
        catch {
            $lastFailureText =
                Get-DmsDetailedError -ErrorRecord $_

            if (
                $syncAttempt -lt $MaximumSyncAttempts -and
                $lastFailureText -match "feature.*not enabled"
            ) {
                Write-Warning $lastFailureText

                $DocumentSetFeatureStatus = Get-PnPFeature `
                    -Identity "3bae86a2-776d-499d-9db8-fa4cdc7884f8" `
                    -Scope Site `
                    -Connection $Connection `
                    -ErrorAction SilentlyContinue

                if ($DocumentSetFeatureStatus) {
                    Write-Host (
                        "The target site reports Document Sets as enabled, " +
                        "but the Content Type Hub service has not yet " +
                        "recognised the feature."
                    )
                }

                Write-Host (
                    "Waiting $SyncDelaySeconds seconds before retrying..."
                )

                Start-Sleep -Seconds $SyncDelaySeconds
                continue
            }

            throw "Content type synchronisation failed. $lastFailureText"
        }

        $attemptFailed = $false
        $retryForFeaturePropagation = $false
        $failureMessages = @()

        foreach ($result in $syncResults) {
            $propertyNames =
                @($result.PSObject.Properties.Name)

            $failedIds = @()

            if ($propertyNames -contains "FailedContentTypeIDs") {
                $failedIds =
                    @($result.FailedContentTypeIDs)
            }

            $failedErrors = @()

            if ($propertyNames -contains "FailedContentTypeErrors") {
                $failedErrors =
                    @($result.FailedContentTypeErrors)
            }

            $failedReason = ""

            if ($propertyNames -contains "FailedReason") {
                $failedReason =
                    [string]$result.FailedReason
            }

            $resultFailed =
                (
                    $propertyNames -contains "IsPassed" -and
                    $result.IsPassed -eq $false
                ) -or
                $failedIds.Count -gt 0 -or
                $failedErrors.Count -gt 0

            if ($resultFailed) {
                $attemptFailed = $true

                $combinedFailure = @(
                    $failedReason
                    $failedIds
                    $failedErrors
                ) -join " | "

                $failureMessages += $combinedFailure

                if (
                    $combinedFailure -match
                    "feature.*not enabled"
                ) {
                    $retryForFeaturePropagation = $true
                }
            }
        }

        if (-not $attemptFailed) {
            $syncSucceeded = $true
            break
        }

        $lastFailureText =
            $failureMessages -join " || "

        if (
            $retryForFeaturePropagation -and
            $syncAttempt -lt $MaximumSyncAttempts
        ) {
            Write-Warning $lastFailureText

            $DocumentSetFeatureStatus = Get-PnPFeature `
                -Identity "3bae86a2-776d-499d-9db8-fa4cdc7884f8" `
                -Scope Site `
                -Connection $Connection `
                -ErrorAction SilentlyContinue

            if ($DocumentSetFeatureStatus) {
                Write-Host (
                    "The target site reports Document Sets as enabled, " +
                    "but the Content Type Hub service has not yet " +
                    "recognised the feature."
                )
            }

            Write-Host (
                "Waiting $SyncDelaySeconds seconds before retrying..."
            )

            Start-Sleep -Seconds $SyncDelaySeconds
            continue
        }

        throw @"
The Content Type Hub rejected one or more content types.

Details:
$lastFailureText
"@
    }

    if (-not $syncSucceeded) {
        throw @"
Content type synchronisation did not succeed.

Last failure:
$lastFailureText
"@
    }

    Write-Host "Waiting for content types on the target site..."

    $missingIds = @($requiredIds)

    for (
        $attempt = 1;
        $attempt -le $MaximumValidationAttempts;
        $attempt++
    ) {
        $availableContentTypes = @(
            Get-PnPContentType `
                -InSiteHierarchy `
                -Connection $Connection `
                -ErrorAction Stop
        )

        $availableIds = @(
            $availableContentTypes |
                ForEach-Object {
                    Get-DmsContentTypeId `
                        -ContentType $_
                } |
                Where-Object {
                    -not [string]::IsNullOrWhiteSpace($_)
                }
        )

        $missingIds = @(
            $requiredIds |
                Where-Object {
                    $requiredId = $_

                    -not (
                        $availableIds |
                            Where-Object {
                                $_ -ieq $requiredId
                            }
                    )
                }
        )

        if ($missingIds.Count -eq 0) {
            Write-Host (
                "All {0} required content types are available." -f
                $requiredIds.Count
            ) -ForegroundColor Green

            return
        }

        Write-Host (
            "Content type validation attempt {0} of {1}. Missing: {2}" -f
            $attempt,
            $MaximumValidationAttempts,
            ($missingIds -join ", ")
        )

        if ($attempt -lt $MaximumValidationAttempts) {
            Start-Sleep -Seconds $ValidationDelaySeconds
        }
    }

    throw @"
The required hub content types did not become available.

Missing IDs:
$($missingIds -join ', ')
"@
}

function Set-DmsCommunicationSiteSensitivityLabel {
    param(
        [Parameter(Mandatory)]
        $AdminConnection,

        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [Guid]$SensitivityLabelId,

        [int]$MaximumAttempts = 30,

        [int]$DelaySeconds = 10
    )

    Write-Host ""
    Write-Host "Applying sensitivity label through the tenant admin endpoint..."
    Write-Host "Site: $SiteUrl"
    Write-Host "Label ID: $SensitivityLabelId"

    try {
        Set-PnPTenantSite `
            -Identity $SiteUrl `
            -SensitivityLabel $SensitivityLabelId.ToString() `
            -Connection $AdminConnection `
            -ErrorAction Stop

        Write-Host "Sensitivity label application requested."
    }
    catch {
        $detail =
            Get-DmsDetailedError -ErrorRecord $_

        throw @"
The Communication Site was created, but the sensitivity
label could not be applied through the SharePoint tenant
administration endpoint.

Site:
$SiteUrl

Label ID:
$SensitivityLabelId

Error:
$detail

Confirm that:
1. The app has SharePoint Sites.FullControl.All application
   permission with administrator consent.
2. The label is published and configured for Groups & Sites.
3. The supplied GUID is the sensitivity label GUID, not the
   label-policy GUID.
"@
    }

    Write-Host "Waiting for the sensitivity label to become visible..."

    $labelApplied = $false
    $appliedLabelId = ""

    for (
        $attempt = 1;
        $attempt -le $MaximumAttempts;
        $attempt++
    ) {
        try {
            $siteDetails = Get-PnPTenantSite `
                -Identity $SiteUrl `
                -Detailed `
                -Connection $AdminConnection `
                -ErrorAction Stop

            $appliedLabelId =
                [string]$siteDetails.SensitivityLabel

            $appliedGuid = [Guid]::Empty

            if (
                [Guid]::TryParse(
                    $appliedLabelId,
                    [ref]$appliedGuid
                ) -and
                $appliedGuid -eq $SensitivityLabelId
            ) {
                $labelApplied = $true

                Write-Host "Sensitivity label applied successfully."
                Write-Host "Applied label ID: $appliedLabelId"
                break
            }

            Write-Host (
                "Label validation attempt {0} of {1}: label not ready." -f
                $attempt,
                $MaximumAttempts
            )
        }
        catch {
            Write-Host (
                "Label validation attempt {0} of {1}: site details not ready." -f
                $attempt,
                $MaximumAttempts
            )
        }

        if ($attempt -lt $MaximumAttempts) {
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    if (-not $labelApplied) {
        Write-Warning @"
The sensitivity-label request completed, but the label was
not visible within the expected time.

Site:
$SiteUrl

Requested label:
$SensitivityLabelId

Last visible label:
$appliedLabelId

Provisioning will continue because sensitivity-label changes
can take several minutes to become visible.
"@
    }

    return $labelApplied
}

# ============================================================
# Validate environment and files
# ============================================================

if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    throw @"
PnP.PowerShell is not installed.

Install it in PowerShell 7 using:
Install-Module PnP.PowerShell -Scope CurrentUser
"@
}

Import-Module PnP.PowerShell `
    -ErrorAction Stop

if ($PSVersionTable.PSVersion -lt [Version]"7.4.0") {
    throw "PowerShell 7.4 or later is required."
}

if (-not $IsWindows) {
    throw "Windows is required for certificate thumbprint authentication."
}

$CertificatePath =
    "Cert:\CurrentUser\My\$Thumbprint"

$Certificate = Get-Item `
    -Path $CertificatePath `
    -ErrorAction SilentlyContinue

if (-not $Certificate) {
    throw "Certificate not found: $CertificatePath"
}

if (-not $Certificate.HasPrivateKey) {
    throw "The certificate does not contain an accessible private key."
}

if ($Certificate.NotAfter -le (Get-Date)) {
    throw "The certificate expired on $($Certificate.NotAfter)."
}

if ($Certificate.NotBefore -gt (Get-Date)) {
    throw "The certificate is not valid until $($Certificate.NotBefore)."
}

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item `
        -Path $OutputFolder `
        -ItemType Directory `
        -Force |
        Out-Null
}

foreach (
    $requiredFile in @(
        $LibraryImportScript,
        $LibraryTemplatePath,
        $LibraryManifestPath
    )
) {
    if (-not (Test-Path -LiteralPath $requiredFile)) {
        throw "Required file not found: $requiredFile"
    }
}

$EffectiveLibraryTemplatePath =
    $LibraryTemplatePath

if ($ExcludeDocumentSetContentTypes) {
    $FilteredLibraryTemplatePath = Join-Path `
        -Path $OutputFolder `
        -ChildPath "Site-Libraries.NoDocumentSets.xml"

    $EffectiveLibraryTemplatePath =
        New-DmsLibraryTemplateWithoutDocumentSets `
            -SourcePath $LibraryTemplatePath `
            -DestinationPath $FilteredLibraryTemplatePath
}

$null = Read-PnPSiteTemplate `
    -Path $EffectiveLibraryTemplatePath

$LibraryManifest = Get-Content `
    -LiteralPath $LibraryManifestPath `
    -Raw `
    -Encoding UTF8 |
    ConvertFrom-Json

$ManifestLibraryNames = @(
    $LibraryManifest.Libraries |
        ForEach-Object {
            [string]$_.Title
        }
)

if ($ManifestLibraryNames.Count -eq 0) {
    throw "The library manifest does not contain any libraries."
}

# ============================================================
# Display details
# ============================================================

Write-Host ""
Write-Host "========================================"
Write-Host "Communication site provisioning"
Write-Host "========================================"
Write-Host "Tenant ID: $TenantId"
Write-Host "Client ID: $ClientId"
Write-Host "Source site: $SourceSiteUrl"
Write-Host "Target title: $SiteTitle"
Write-Host "Target URL: $TargetSiteUrl"
Write-Host "Owner: $Owner"
Write-Host "Communication site design: $CommunicationSiteDesign"
Write-Host "Requested source navigation: $SourceNavigationLocation"
Write-Host "Target navigation location: $NavigationLocation"
Write-Host "Sensitivity label: $SensitivityLabelId"
Write-Host "Exclude Document Set content types: $ExcludeDocumentSetContentTypes"
Write-Host "Libraries: $($ManifestLibraryNames -join ', ')"
Write-Host ""

# ============================================================
# Connect to source and export homepage/navigation
# ============================================================

$SourceConnection = Connect-PnPOnline `
    -Url $SourceSiteUrl `
    -Tenant $TenantId `
    -ClientId $ClientId `
    -Thumbprint $Thumbprint `
    -ValidateConnection `
    -ReturnConnection `
    -ErrorAction Stop

$SourceWeb = Get-PnPWeb `
    -Includes WebTemplate,Configuration `
    -Connection $SourceConnection `
    -ErrorAction Stop

Write-Host "Connected to source site: $($SourceWeb.Title)"
Write-Host (
    "Source template: {0}#{1}" -f
    $SourceWeb.WebTemplate,
    $SourceWeb.Configuration
)

if ($SourceWeb.WebTemplate -ne "SITEPAGEPUBLISHING") {
    Write-Warning @"
The source site is not reported as a Communication Site.
For the closest result, use a Communication Site as the
source template, especially for top navigation.
"@
}

$HomePageUrl = Get-PnPHomePage `
    -Connection $SourceConnection `
    -ErrorAction Stop

$HomePageName = Split-Path `
    -Path $HomePageUrl `
    -Leaf

if ([string]::IsNullOrWhiteSpace($HomePageName)) {
    throw "The source homepage could not be identified."
}

$SourcePage = Get-PnPPage `
    -Identity $HomePageName `
    -Connection $SourceConnection `
    -ErrorAction Stop

if (-not $SourcePage) {
    throw "The source homepage could not be retrieved."
}

if (@($SourcePage.Sections).Count -eq 0) {
    throw "The source homepage contains no exportable sections."
}

if (Test-Path -LiteralPath $PageTemplate) {
    Remove-Item `
        -LiteralPath $PageTemplate `
        -Force
}

Export-PnPPage `
    -Identity $HomePageName `
    -Out $PageTemplate `
    -PersistBrandingFiles `
    -Force `
    -Connection $SourceConnection `
    -ErrorAction Stop

if (-not (Test-Path -LiteralPath $PageTemplate)) {
    throw "The homepage package was not created."
}

$ResolvedSourceNavigationLocation =
    $SourceNavigationLocation

$SkipNavigationCopy = $false

if ($SourceNavigationLocation -eq "Auto") {
    Write-Host "Checking source Top Navigation..."

    $SourceNavigation = @(
        Get-DmsNavigationNodes `
            -Location TopNavigationBar `
            -Connection $SourceConnection
    )

    if ($SourceNavigation.Count -gt 0) {
        $ResolvedSourceNavigationLocation =
            "TopNavigationBar"
    }
    else {
        Write-Host "No source Top Navigation found."
        Write-Host "Checking source Quick Launch..."

        $SourceNavigation = @(
            Get-DmsNavigationNodes `
                -Location QuickLaunch `
                -Connection $SourceConnection
        )

        $ResolvedSourceNavigationLocation =
            "QuickLaunch"
    }
}
else {
    $SourceNavigation = @(
        Get-DmsNavigationNodes `
            -Location $SourceNavigationLocation `
            -Connection $SourceConnection
    )
}

if ($SourceNavigation.Count -eq 0) {
    $SkipNavigationCopy = $true
    $ResolvedSourceNavigationLocation = "None"

    Write-Warning @"
The source site has no local Top Navigation or Quick Launch
nodes. Navigation copying will be skipped and the target
Communication Site will retain its default navigation.

This can occur when the source uses hub navigation, global
navigation, or has local navigation disabled.
"@
}
else {
    Write-Host (
        "Source navigation selected: {0}" -f
        $ResolvedSourceNavigationLocation
    )

    Write-Host (
        "Source navigation nodes: {0}" -f
        $SourceNavigation.Count
    )
}

# ============================================================
# Connect to admin and root
# ============================================================

$AdminConnection = Connect-PnPOnline `
    -Url $AdminSiteUrl `
    -Tenant $TenantId `
    -ClientId $ClientId `
    -Thumbprint $Thumbprint `
    -ValidateConnection `
    -ReturnConnection `
    -ErrorAction Stop

$RootConnection = Connect-PnPOnline `
    -Url $TenantRootUrl `
    -Tenant $TenantId `
    -ClientId $ClientId `
    -Thumbprint $Thumbprint `
    -ValidateConnection `
    -ReturnConnection `
    -ErrorAction Stop

# ============================================================
# Validate target URL
# ============================================================

Write-Host "Checking whether the target site already exists..."

if (
    Test-DmsSiteExists `
        -AdminConnection $AdminConnection `
        -SiteUrl $TargetSiteUrl
) {
    throw "A site already exists at '$TargetSiteUrl'."
}

Write-Host "Checking the deleted-sites recycle bin..."

if (
    Test-DmsDeletedSiteExists `
        -AdminConnection $AdminConnection `
        -SiteUrl $TargetSiteUrl
) {
    throw @"
A deleted site still exists at:
$TargetSiteUrl

Permanently remove or restore that site before provisioning.
"@
}

# ============================================================
# Create Communication Site
# ============================================================

Write-Host ""
Write-Host "Creating Communication Site..."

$NewSiteParameters = @{
    Type        = "CommunicationSite"
    Title       = $SiteTitle
    Url         = $TargetSiteUrl
    Owner       = $Owner
    SiteDesign  = $CommunicationSiteDesign
    Connection  = $AdminConnection
    ErrorAction = "Stop"
}

try {
    $NewSiteResult =
        New-PnPSite @NewSiteParameters
}
catch {
    $detail = Get-DmsDetailedError -ErrorRecord $_

    throw @"
Communication Site creation failed.

Target site:
$TargetSiteUrl

Owner:
$Owner

Error:
$detail
"@
}

Write-Host "Communication Site creation request completed."

# ============================================================
# Connect to target
# ============================================================

$TargetConnection = $null
$TargetWeb = $null

for ($attempt = 1; $attempt -le 30; $attempt++) {
    try {
        $TargetConnection = Connect-PnPOnline `
            -Url $TargetSiteUrl `
            -Tenant $TenantId `
            -ClientId $ClientId `
            -Thumbprint $Thumbprint `
            -ValidateConnection `
            -ReturnConnection `
            -ErrorAction Stop

        $TargetWeb = Get-PnPWeb `
            -Includes WebTemplate,Configuration `
            -Connection $TargetConnection `
            -ErrorAction Stop

        break
    }
    catch {
        Write-Host "Target connection attempt $attempt of 30..."

        if ($attempt -eq 30) {
            throw @"
The new Communication Site did not become available.

Target site:
$TargetSiteUrl

Error:
$($_.Exception.Message)
"@
        }

        Start-Sleep -Seconds 10
    }
}

Write-Host "Connected to target site: $($TargetWeb.Title)"

if ($TargetWeb.WebTemplate -ne "SITEPAGEPUBLISHING") {
    throw @"
The new site was created, but it is not reported as a
Communication Site.

Template:
$($TargetWeb.WebTemplate)#$($TargetWeb.Configuration)
"@
}

if (-not [string]::IsNullOrWhiteSpace($Description)) {
    Set-PnPWeb `
        -Description $Description `
        -Connection $TargetConnection `
        -ErrorAction Stop
}

# ============================================================
# Configure and verify site collection owner
# ============================================================

Write-Host ""
Write-Host "Configuring site collection owner: $Owner"

Set-PnPTenantSite `
    -Identity $TargetSiteUrl `
    -PrimarySiteCollectionAdmin $Owner `
    -Owners @($Owner) `
    -Connection $AdminConnection `
    -ErrorAction Stop

Add-PnPSiteCollectionAdmin `
    -Owners $Owner `
    -Connection $TargetConnection `
    -ErrorAction Stop

$SiteCollectionAdmins = @(
    Get-PnPSiteCollectionAdmin `
        -Connection $TargetConnection `
        -ErrorAction Stop
)

$OwnerVerified = $false

foreach ($admin in $SiteCollectionAdmins) {
    if (
        [string]$admin.Email -ieq $Owner -or
        [string]$admin.LoginName -like "*$Owner"
    ) {
        $OwnerVerified = $true
        break
    }
}

if (-not $OwnerVerified) {
    throw "The owner was not found in the site collection administrators."
}

Write-Host "Site collection owner configured successfully."

# ============================================================
# Apply sensitivity label
# ============================================================

$LabelApplied =
    Set-DmsCommunicationSiteSensitivityLabel `
        -AdminConnection $AdminConnection `
        -SiteUrl $TargetSiteUrl `
        -SensitivityLabelId $SensitivityLabelId

# ============================================================
# Enable required features
# ============================================================

if ($EnableTaxonomy) {
    Enable-DmsSiteFeature `
        -AdminConnection $AdminConnection `
        -TargetConnection $TargetConnection `
        -SiteUrl $TargetSiteUrl `
        -FeatureId "73ef14b1-13a9-416b-a9b5-ececa2b0604c" `
        -FeatureName "Taxonomy"
}

# ============================================================
# Select and synchronise content types
# ============================================================

$DocumentSetContentTypePrefix =
    "0x0120D520"

$ExplicitContentTypeIds = @(
    $RequiredHubContentTypeIds |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_)
        } |
        ForEach-Object {
            ([string]$_).Trim()
        }
)

if ($SyncAllCompatibleHubContentTypes) {
    $CompatibleContentTypeIds = @(
        Get-DmsCompatibleHubContentTypeIds `
            -Connection $TargetConnection `
            -SiteUrl $TargetSiteUrl
    )

    $ContentTypeIdsToSync = @(
        @(
            $CompatibleContentTypeIds
            $ExplicitContentTypeIds
        ) |
            Sort-Object Length |
            Select-Object -Unique
    )
}
else {
    $ContentTypeIdsToSync = @(
        $ExplicitContentTypeIds |
            Sort-Object Length |
            Select-Object -Unique
    )
}

if ($ExcludeDocumentSetContentTypes) {
    $ExcludedContentTypeIds = @(
        $ContentTypeIdsToSync |
            Where-Object {
                $_.StartsWith(
                    $DocumentSetContentTypePrefix,
                    [StringComparison]::OrdinalIgnoreCase
                )
            }
    )

    $ContentTypeIdsToSync = @(
        $ContentTypeIdsToSync |
            Where-Object {
                -not $_.StartsWith(
                    $DocumentSetContentTypePrefix,
                    [StringComparison]::OrdinalIgnoreCase
                )
            }
    )

    Write-Host ""
    Write-Host "Document Set content types are excluded."

    if ($ExcludedContentTypeIds.Count -gt 0) {
        Write-Host (
            "Excluded content type IDs: {0}" -f
            ($ExcludedContentTypeIds -join ", ")
        )
    }
}

if ($ContentTypeIdsToSync.Count -gt 0) {
    Confirm-DmsHubContentTypes `
        -Connection $TargetConnection `
        -SiteUrl $TargetSiteUrl `
        -ContentTypeIds $ContentTypeIdsToSync
}
else {
    Write-Host "No hub content types were selected for synchronisation."
}

# ============================================================
# Import document libraries
# ============================================================

Write-Host ""
Write-Host "Importing document libraries..."

try {
    $LibraryImportOutput = @(
        & $LibraryImportScript `
            -TenantId $TenantId `
            -ClientId $ClientId `
            -Thumbprint $Thumbprint `
            -TargetSiteUrl $TargetSiteUrl `
            -TemplatePath $EffectiveLibraryTemplatePath `
            -ManifestPath $LibraryManifestPath
    )
}
catch {
    throw @"
The Communication Site was created, but its document
libraries could not be imported.

Target site:
$TargetSiteUrl

Error:
$($_.Exception.Message)
"@
}

$LibraryImportResult = $LibraryImportOutput |
    Where-Object {
        $null -ne $_ -and
        ($_.PSObject.Properties.Name -contains "Success")
    } |
    Select-Object -Last 1

if (
    -not $LibraryImportResult -or
    $LibraryImportResult.Success -ne $true
) {
    throw "The document-library import did not report success."
}

$ImportedLibraryCount =
    [int]$LibraryImportResult.LibraryCount

$ImportedLibraryNames =
    @($LibraryImportResult.Libraries)

# ============================================================
# Apply homepage
# ============================================================

$TemporaryPageName =
    "ProvisioningTemporaryHome"

$TemporaryPageFile =
    "$TemporaryPageName.aspx"

$TemporaryPagePath =
    "SitePages/$TemporaryPageFile"

$TemporaryPage = Get-PnPPage `
    -Identity $TemporaryPageFile `
    -Connection $TargetConnection `
    -ErrorAction SilentlyContinue

if (-not $TemporaryPage) {
    Add-PnPPage `
        -Name $TemporaryPageName `
        -LayoutType Home `
        -Publish `
        -Connection $TargetConnection `
        -ErrorAction Stop |
        Out-Null
}

Set-PnPHomePage `
    -RootFolderRelativeUrl $TemporaryPagePath `
    -Connection $TargetConnection `
    -ErrorAction Stop

$ExistingHomePage = Get-PnPPage `
    -Identity $HomePageName `
    -Connection $TargetConnection `
    -ErrorAction SilentlyContinue

if ($ExistingHomePage) {
    Remove-PnPPage `
        -Identity $HomePageName `
        -Force `
        -Connection $TargetConnection `
        -ErrorAction Stop
}

Invoke-PnPSiteTemplate `
    -Path $PageTemplate `
    -Handlers Pages `
    -Connection $TargetConnection `
    -ErrorAction Stop `
    -Verbose

$CopiedHomePage = Get-PnPPage `
    -Identity $HomePageName `
    -Connection $TargetConnection `
    -ErrorAction SilentlyContinue

if (-not $CopiedHomePage) {
    throw "The homepage package did not create '$HomePageName'."
}

Set-PnPPage `
    -Identity $HomePageName `
    -Publish `
    -Connection $TargetConnection `
    -ErrorAction Stop

Set-PnPHomePage `
    -RootFolderRelativeUrl "SitePages/$HomePageName" `
    -Connection $TargetConnection `
    -ErrorAction Stop

$ConfiguredHomePage = Get-PnPHomePage `
    -Connection $TargetConnection `
    -ErrorAction Stop

if ($ConfiguredHomePage -notlike "*$HomePageName") {
    throw "The copied homepage was not configured as the homepage."
}

$TemporaryPage = Get-PnPPage `
    -Identity $TemporaryPageFile `
    -Connection $TargetConnection `
    -ErrorAction SilentlyContinue

if ($TemporaryPage) {
    Remove-PnPPage `
        -Identity $TemporaryPageFile `
        -Force `
        -Connection $TargetConnection `
        -ErrorAction Stop
}

# ============================================================
# Apply and validate navigation
# ============================================================

if ($SkipNavigationCopy) {
    Write-Host ""
    Write-Host "Navigation copying was skipped."

    $TargetNavigation = @(
        Get-PnPNavigationNode `
            -Location $NavigationLocation `
            -Tree `
            -Connection $TargetConnection `
            -ErrorAction SilentlyContinue |
            Where-Object {
                $null -ne $_ -and
                -not [string]::IsNullOrWhiteSpace($_.Title)
            }
    )
}
else {
    Write-Host ""
    Write-Host (
        "Mapping source {0} navigation to target {1} navigation..." -f
        $ResolvedSourceNavigationLocation,
        $NavigationLocation
    )

    $ExistingTargetNavigation = @(
        Get-PnPNavigationNode `
            -Location $NavigationLocation `
            -Connection $TargetConnection `
            -ErrorAction Stop
    )

    foreach ($existingNode in $ExistingTargetNavigation) {
        Remove-PnPNavigationNode `
            -Identity $existingNode.Id `
            -Force `
            -Connection $TargetConnection `
            -ErrorAction Stop
    }

    Copy-DmsNavigationTree `
        -Nodes $SourceNavigation `
        -TargetLocation $NavigationLocation `
        -TargetConnection $TargetConnection `
        -SourceSiteUrl $SourceSiteUrl `
        -TargetSiteUrl $TargetSiteUrl

    $TargetNavigation = @(
        Get-DmsNavigationNodes `
            -Location $NavigationLocation `
            -Connection $TargetConnection
    )

    if ($TargetNavigation.Count -eq 0) {
        throw @"
Source navigation was found, but no target navigation nodes
were created at:
$NavigationLocation
"@
    }
}

# ============================================================
# Final result
# ============================================================

Write-Host ""
Write-Host "========================================"
Write-Host "Communication Site provisioning completed"
Write-Host "========================================"
Write-Host "Target site: $TargetSiteUrl"
Write-Host "Owner: $Owner"
Write-Host "Sensitivity label visible: $LabelApplied"
Write-Host "Content types synchronised: $($ContentTypeIdsToSync.Count)"
Write-Host "Document Set content types excluded: $ExcludeDocumentSetContentTypes"
Write-Host "Libraries imported: $ImportedLibraryCount"
Write-Host "Library names: $($ImportedLibraryNames -join ', ')"
Write-Host "Homepage: $ConfiguredHomePage"
Write-Host "Source navigation location: $ResolvedSourceNavigationLocation"
Write-Host "Target navigation location: $NavigationLocation"
Write-Host "Navigation nodes: $($TargetNavigation.Count)"
Write-Host "Certificate expiry: $($Certificate.NotAfter)"

[PSCustomObject]@{
    Success                       = $true
    TenantId                      = $TenantId
    ClientId                      = $ClientId
    SourceSiteUrl                 = $SourceSiteUrl
    TargetSiteTitle               = $SiteTitle
    TargetSiteAlias               = $SiteAlias
    TargetSiteUrl                 = $TargetSiteUrl
    CommunicationSiteDesign       = $CommunicationSiteDesign
    Owner                         = $Owner
    SensitivityLabelId            = $SensitivityLabelId
    SensitivityLabelVisible       = $LabelApplied
    DocumentSetContentTypesExcluded = $ExcludeDocumentSetContentTypes
    EffectiveLibraryTemplatePath    = $EffectiveLibraryTemplatePath
    TaxonomyEnabled                 = $EnableTaxonomy
    SyncAllCompatibleContentTypes = $SyncAllCompatibleHubContentTypes
    ContentTypesSynchronised      = $ContentTypeIdsToSync
    LibrariesImported             = $ImportedLibraryCount
    LibraryNames                  = $ImportedLibraryNames
    Homepage                      = $ConfiguredHomePage
    SourceNavigationLocation      = $ResolvedSourceNavigationLocation
    NavigationLocation            = $NavigationLocation
    NavigationNodeCount           = $TargetNavigation.Count
    CertificateExpiry             = $Certificate.NotAfter
}

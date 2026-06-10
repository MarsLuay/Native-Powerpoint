param(
    [switch]$NoBuild,
    [string]$VaultPath
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PluginRoot = Split-Path -Parent $ScriptDir
$PluginManifest = Join-Path $PluginRoot "manifest.json"
$PluginId = "native-powerpoint-doc-editor"
$PluginName = "Native PowerPoint Doc Editor"
$ObsoletePluginIds = @("native-powerpoint", "docxidian")
$LogFile = Join-Path $env:TEMP "native-powerpoint-doc-editor-install.log"
$AutoSearchDepth = if ($env:NATIVE_POWERPOINT_AUTO_SEARCH_DEPTH) { [int]$env:NATIVE_POWERPOINT_AUTO_SEARCH_DEPTH } else { 6 }
$SelectedDirSearchDepth = if ($env:NATIVE_POWERPOINT_SELECTED_DIR_SEARCH_DEPTH) { [int]$env:NATIVE_POWERPOINT_SELECTED_DIR_SEARCH_DEPTH } else { 10 }

Set-Content -Path $LogFile -Value "" -Encoding UTF8

function Write-Log {
    param([string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"
    $line = "[$timestamp] $Message"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Add-NodeInstallPathsToProcessPath {
    $candidateDirs = @()

    if ($env:ProgramFiles) {
        $candidateDirs += (Join-Path $env:ProgramFiles "nodejs")
    }

    $programFilesX86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
    if ($programFilesX86) {
        $candidateDirs += (Join-Path $programFilesX86 "nodejs")
    }

    foreach ($dir in $candidateDirs) {
        if ((Test-Path $dir) -and ($env:Path -notlike "*$dir*")) {
            $env:Path = "$dir;$env:Path"
        }
    }
}

function Install-NodeRuntime {
    if ($env:OS -ne "Windows_NT") {
        throw "Node.js/npm is required. Install Node.js LTS from https://nodejs.org/ and run this installer again."
    }

    $winget = Get-Command "winget" -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw "Node.js/npm is required, but winget was not found. Install Node.js LTS from https://nodejs.org/ and run this installer again."
    }

    Write-Log "Node.js/npm was not found. Installing Node.js LTS with winget..."
    & $winget.Source install --id OpenJS.NodeJS.LTS --exact --source winget --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) {
        throw "winget failed to install Node.js LTS. Exit code: $LASTEXITCODE"
    }

    Add-NodeInstallPathsToProcessPath
}

function Get-NpmCommand {
    Add-NodeInstallPathsToProcessPath

    $npm = Get-Command "npm.cmd" -ErrorAction SilentlyContinue
    if (-not $npm) {
        $npm = Get-Command "npm" -ErrorAction SilentlyContinue
    }

    return $npm
}

function Get-ManifestValue {
    param(
        [string]$ManifestFile,
        [string]$Key
    )

    $manifest = Get-Content -Raw -Path $ManifestFile | ConvertFrom-Json
    $value = $manifest.$Key
    if ($null -eq $value) {
        return ""
    }

    return [string]$value
}

function Test-SkippedPath {
    param([string]$Path)

    return (
        $Path -match "\\node_modules\\" -or
        $Path -match "\\\.Trash\\" -or
        $Path -match "\\AppData\\Local\\Temp\\" -or
        $Path -match "\\AppData\\Local\\Microsoft\\Windows\\INetCache\\"
    )
}

function Add-VaultCandidate {
    param(
        [System.Collections.Generic.List[string]]$Vaults,
        [string]$Candidate
    )

    if (-not $Candidate) {
        return
    }

    $obsidianDir = Join-Path $Candidate ".obsidian"
    if ((Test-Path $obsidianDir) -and (-not (Test-SkippedPath $obsidianDir))) {
        $resolved = (Resolve-Path $Candidate).Path
        if (-not $Vaults.Contains($resolved)) {
            [void]$Vaults.Add($resolved)
        }
    }
}

function Find-VaultsUnder {
    param(
        [string]$Root,
        [int]$Depth = $SelectedDirSearchDepth
    )

    $vaults = [System.Collections.Generic.List[string]]::new()

    if (-not $Root) {
        return $vaults
    }

    $expanded = [Environment]::ExpandEnvironmentVariables($Root.Trim('"'))
    if (-not (Test-Path $expanded)) {
        return $vaults
    }

    try {
        $resolvedRoot = (Resolve-Path $expanded).Path
        $obsidianDirs = Get-ChildItem -Path $resolvedRoot -Force -Directory -Filter ".obsidian" -Recurse -Depth $Depth -ErrorAction SilentlyContinue
        foreach ($obsidianDir in $obsidianDirs) {
            if (Test-SkippedPath $obsidianDir.FullName) {
                continue
            }

            Add-VaultCandidate $vaults $obsidianDir.Parent.FullName
        }
    } catch {
        Write-Log "WARNING: Could not search for vaults under ${Root}: $($_.Exception.Message)"
    }

    return $vaults
}

function Resolve-InstallVault {
    param([string]$Candidate)

    if (-not $Candidate) {
        return $null
    }

    $expanded = [Environment]::ExpandEnvironmentVariables($Candidate.Trim('"'))
    if (-not (Test-Path $expanded)) {
        Write-Host "Folder not found: $expanded"
        return $null
    }

    $resolved = (Resolve-Path $expanded).Path
    $obsidianDir = Join-Path $resolved ".obsidian"
    if (Test-Path $obsidianDir) {
        return $resolved
    }

    Write-Log "Searching selected folder for existing Obsidian vaults: $resolved"
    $nestedVaults = @(Find-VaultsUnder $resolved $SelectedDirSearchDepth)
    if ($nestedVaults.Count -gt 0) {
        return Select-ObsidianVault $nestedVaults
    }

    Write-Host ""
    Write-Host "The selected folder does not contain an .obsidian folder:"
    Write-Host "  $resolved"
    $answer = Read-Host "Create .obsidian there and install into this folder? [y/N]"
    if ($answer -match "^(y|yes)$") {
        New-Item -ItemType Directory -Force -Path $obsidianDir | Out-Null
        return $resolved
    }

    return $null
}

function Select-FolderWithDialog {
    param([string]$Description)

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $Description
        $dialog.ShowNewFolderButton = $true

        $result = $dialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }
    } catch {
        Write-Log "WARNING: Windows folder picker is unavailable: $($_.Exception.Message)"
    }

    return ""
}

function Prompt-ForVaultFolder {
    param([string]$Reason)

    Write-Host ""
    Write-Host $Reason

    while ($true) {
        $candidate = Select-FolderWithDialog "Choose the Obsidian vault folder to install Native PowerPoint Doc Editor into."

        if (-not $candidate) {
            $candidate = Read-Host "Enter the path to your Obsidian vault folder, or leave blank to cancel"
        }

        if (-not $candidate) {
            throw "No Obsidian vault folder was selected."
        }

        $resolved = Resolve-InstallVault $candidate
        if ($resolved) {
            return $resolved
        }

        Write-Host "Choose another folder."
    }
}

function Find-ObsidianVaults {
    $vaults = [System.Collections.Generic.List[string]]::new()

    if ($VaultPath) {
        Add-VaultCandidate $vaults $VaultPath
        if ($vaults.Count -eq 0) {
            foreach ($vault in (Find-VaultsUnder $VaultPath $SelectedDirSearchDepth)) {
                Add-VaultCandidate $vaults $vault
            }
        }
        return $vaults
    }

    if ($env:OBSIDIAN_VAULT) {
        Add-VaultCandidate $vaults $env:OBSIDIAN_VAULT
        if ($vaults.Count -eq 0) {
            foreach ($vault in (Find-VaultsUnder $env:OBSIDIAN_VAULT $SelectedDirSearchDepth)) {
                Add-VaultCandidate $vaults $vault
            }
        }
        return $vaults
    }

    $current = Resolve-Path $PluginRoot
    while ($current) {
        Add-VaultCandidate $vaults $current.Path
        $parent = Split-Path -Parent $current.Path
        if (-not $parent -or $parent -eq $current.Path) {
            break
        }
        $current = Resolve-Path $parent
    }

    $profileRoot = if ($env:USERPROFILE) { $env:USERPROFILE } else { $HOME }
    $searchRootCandidates = @()
    if ($profileRoot) {
        $searchRootCandidates += $profileRoot
        $searchRootCandidates += (Join-Path $profileRoot "Documents")
        $searchRootCandidates += (Join-Path $profileRoot "Desktop")
    }
    $fileSystemDrives = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
        Where-Object { $_.Root -and (Test-Path $_.Root) } |
        ForEach-Object { $_.Root }

    $searchRootCandidates += $env:OneDrive
    $searchRootCandidates += $env:OneDriveConsumer
    $searchRootCandidates += $env:OneDriveCommercial
    if ($env:SystemDrive) {
        $searchRootCandidates += (Join-Path $env:SystemDrive "Users")
    }

    $searchRoots = $searchRootCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

    foreach ($root in $searchRoots) {
        foreach ($vault in (Find-VaultsUnder $root $AutoSearchDepth)) {
            Add-VaultCandidate $vaults $vault
        }
    }

    if ($vaults.Count -eq 0) {
        $broadSearchRoots = @($searchRoots + $fileSystemDrives) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
        foreach ($root in $broadSearchRoots) {
            Write-Log "Doing broader vault search under ${root}."
            foreach ($vault in (Find-VaultsUnder $root $SelectedDirSearchDepth)) {
                Add-VaultCandidate $vaults $vault
            }

            if ($vaults.Count -gt 0) {
                break
            }
        }
    }

    return $vaults
}

function Select-ObsidianVault {
    param([string[]]$Vaults)

    if ($Vaults.Count -eq 0) {
        return Prompt-ForVaultFolder "No Obsidian vault folders were found automatically."
    }

    if ($Vaults.Count -eq 1) {
        return $Vaults[0]
    }

    Write-Host ""
    Write-Host "Multiple Obsidian vaults were detected. Choose where to install Native PowerPoint Doc Editor:"
    for ($i = 0; $i -lt $Vaults.Count; $i++) {
        Write-Host ("  {0}) {1}" -f ($i + 1), $Vaults[$i])
    }
    Write-Host "  B) Browse for another folder"

    while ($true) {
        $answer = Read-Host "Enter a number"
        if ($answer -match "^(b|browse)$") {
            return Prompt-ForVaultFolder "Choose the Obsidian vault folder to install into."
        }

        $selected = 0
        if ([int]::TryParse($answer, [ref]$selected) -and $selected -ge 1 -and $selected -le $Vaults.Count) {
            return $Vaults[$selected - 1]
        }
        Write-Host "Please enter a number from 1 to $($Vaults.Count)."
    }
}

function Test-PackageScript {
    param(
        [string]$PackageFile,
        [string]$ScriptName
    )

    $package = Get-Content -Raw -Path $PackageFile | ConvertFrom-Json
    return $null -ne $package.scripts -and $null -ne $package.scripts.$ScriptName
}

function Invoke-LoggedCommand {
    param(
        [string]$FilePath,
        [string[]]$Arguments,
        [string]$WorkingDirectory
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = ($Arguments | ForEach-Object {
        if ($_ -match "\s") {
            '"' + ($_ -replace '"', '\"') + '"'
        } else {
            $_
        }
    }) -join " "
    $psi.WorkingDirectory = $WorkingDirectory
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false

    $process = [System.Diagnostics.Process]::Start($psi)
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    if ($stdout) {
        Add-Content -Path $LogFile -Value $stdout -Encoding UTF8
    }

    if ($stderr) {
        Add-Content -Path $LogFile -Value $stderr -Encoding UTF8
    }

    if ($process.ExitCode -ne 0) {
        throw "$FilePath $($psi.Arguments) failed with exit code $($process.ExitCode). See log: $LogFile"
    }
}

function Install-NodeDependencies {
    $npm = Get-NpmCommand
    if (-not $npm) {
        Install-NodeRuntime
        $npm = Get-NpmCommand
    }

    if (-not $npm) {
        throw "npm is still unavailable after attempting to install Node.js. Install Node.js LTS from https://nodejs.org/ and run this installer again."
    }

    if (-not (Test-Path (Join-Path $PluginRoot "node_modules"))) {
        if (Test-Path (Join-Path $PluginRoot "package-lock.json")) {
            Write-Log "Installing npm dependencies with npm ci."
            try {
                Invoke-LoggedCommand $npm.Source @("ci") $PluginRoot
            } catch {
                Write-Log "WARNING: npm ci failed. Trying npm install instead."
                Invoke-LoggedCommand $npm.Source @("install") $PluginRoot
            }
        } else {
            Write-Log "Installing npm dependencies with npm install."
            Invoke-LoggedCommand $npm.Source @("install") $PluginRoot
        }
    } else {
        Write-Log "npm dependencies already present."
    }

    return $npm
}

function Assert-ReleaseFilesPresent {
    $requiredFiles = @("manifest.json", "main.js", "styles.css")
    $missing = @()

    foreach ($fileName in $requiredFiles) {
        if (-not (Test-Path (Join-Path $PluginRoot $fileName))) {
            $missing += $fileName
        }
    }

    if ($missing.Count -gt 0) {
        throw "The plugin build is missing required file(s): $($missing -join ', '). See log: $LogFile"
    }
}

function Build-PluginIfPossible {
    if ($NoBuild -or (-not (Test-Path (Join-Path $PluginRoot "package.json")))) {
        Assert-ReleaseFilesPresent
        return
    }

    $npm = Get-NpmCommand
    if (-not $npm) {
        try {
            Install-NodeRuntime
            $npm = Get-NpmCommand
        } catch {
            if (Test-Path (Join-Path $PluginRoot "main.js")) {
                Write-Log "WARNING: npm was not found. Installing the already-built plugin files. $($_.Exception.Message)"
                Assert-ReleaseFilesPresent
                return
            }

            throw
        }
    }

    $npm = Install-NodeDependencies

    if (Test-PackageScript (Join-Path $PluginRoot "package.json") "build") {
        Write-Log "Building $PluginName."
        Invoke-LoggedCommand $npm.Source @("run", "build") $PluginRoot
    }

    Assert-ReleaseFilesPresent
}

function Get-EnabledPluginIds {
    param([string]$CommunityPluginsFile)

    if (-not (Test-Path $CommunityPluginsFile)) {
        return @()
    }

    $raw = (Get-Content -Raw -Path $CommunityPluginsFile).Trim()
    if (-not $raw) {
        return @()
    }

    if (-not $raw.StartsWith("[")) {
        throw "$CommunityPluginsFile must contain a JSON array."
    }

    $parsed = $raw | ConvertFrom-Json
    return @($parsed | Where-Object { $_ -is [string] })
}

function Set-EnabledPluginIds {
    param(
        [string]$CommunityPluginsFile,
        [string[]]$EnabledPluginIds
    )

    $communityPluginsDir = Split-Path -Parent $CommunityPluginsFile
    New-Item -ItemType Directory -Force -Path $communityPluginsDir | Out-Null

    $jsonItems = @($EnabledPluginIds) | ForEach-Object { "  " + ($_ | ConvertTo-Json -Compress) }
    if ($jsonItems.Count -eq 0) {
        "[]" | Set-Content -Encoding UTF8 -Path $CommunityPluginsFile
    } else {
        ("[`n" + ($jsonItems -join ",`n") + "`n]") | Set-Content -Encoding UTF8 -Path $CommunityPluginsFile
    }
}

function Install-Plugin {
    param([string]$Vault)

    $pluginsDir = Join-Path $Vault ".obsidian\plugins"
    $targetDir = Join-Path $pluginsDir $PluginId
    $communityPluginsFile = Join-Path $Vault ".obsidian\community-plugins.json"

    New-Item -ItemType Directory -Force -Path $targetDir | Out-Null

    foreach ($obsoleteId in $ObsoletePluginIds) {
        $obsoleteDir = Join-Path $pluginsDir $obsoleteId
        if (Test-Path $obsoleteDir) {
            Remove-Item -Recurse -Force $obsoleteDir
            Write-Log "Removed obsolete plugin install: $obsoleteDir"
        }
    }

    $files = @()
    $files += Get-Item (Join-Path $PluginRoot "manifest.json")
    $files += Get-ChildItem -Path $PluginRoot -File | Where-Object { $_.Extension -in @(".js", ".css") }

    foreach ($file in $files) {
        Copy-Item -Force -Path $file.FullName -Destination (Join-Path $targetDir $file.Name)
    }

    if ((-not (Test-Path (Join-Path $targetDir "manifest.json"))) -or (-not (Test-Path (Join-Path $targetDir "main.js")))) {
        throw "Installed plugin is missing manifest.json or main.js: $targetDir"
    }

    $enabled = Get-EnabledPluginIds $communityPluginsFile |
        Where-Object { $ObsoletePluginIds -notcontains $_ }

    if ($enabled -notcontains $PluginId) {
        $enabled = @($enabled + $PluginId)
    }

    Set-EnabledPluginIds $communityPluginsFile $enabled
    Write-Log "Installed $PluginId to $targetDir ($($files.Count) file(s); existing data.json preserved)."
}

try {
    if (-not (Test-Path $PluginManifest)) {
        throw "Could not find plugin manifest: $PluginManifest"
    }

    $manifestId = Get-ManifestValue $PluginManifest "id"
    $manifestName = Get-ManifestValue $PluginManifest "name"
    if ($manifestId) {
        $PluginId = $manifestId
    }
    if ($manifestName) {
        $PluginName = $manifestName
    }

    $vaults = Find-ObsidianVaults
    $selectedVault = Select-ObsidianVault $vaults

    Write-Log "Selected Obsidian vault: $selectedVault"
    Write-Log "Plugin root: $PluginRoot"

    Build-PluginIfPossible
    Install-Plugin $selectedVault

    Write-Host ""
    Write-Host "$PluginName was installed and enabled in:"
    Write-Host "  $selectedVault"
    Write-Host ""
    Write-Host "Reload Obsidian or disable/re-enable the plugin to pick up changes."
    Write-Host "Log: $LogFile"
} catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Host ""
    Write-Host "Install failed. Log: $LogFile"
    exit 1
}

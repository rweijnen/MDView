# MDView build and package script
param(
    [switch]$Release,
    [switch]$x86,
    [switch]$x64,
    [switch]$Clean,
    [switch]$Package,
    [string]$OutputDir = "dist",
    [string]$ZipFile
)

$ErrorActionPreference = "Stop"

if ($Clean) {
    Write-Host "Cleaning build artifacts..."
    cargo clean
    exit 0
}

# By default:
# - without -Package: build x64 only (preserves existing behavior)
# - with -Package and no explicit arch flags: build both x64 and x86
if (-not $x86 -and -not $x64) {
    $build_x64 = $true
    $build_x86 = $Package
} else {
    $build_x64 = $x64
    $build_x86 = $x86
}

$targets = @()
if ($build_x64) {
    $targets += [PSCustomObject]@{ Target = "x86_64-pc-windows-msvc"; Arch = "x64" }
}
if ($build_x86) {
    $targets += [PSCustomObject]@{ Target = "i686-pc-windows-msvc"; Arch = "x86" }
}

$profile = if ($Release) { "--release" } else { "" }
$profileDir = if ($Release) { "release" } else { "debug" }

$archs = $targets | ForEach-Object { $_.Arch }
Write-Host "Building MDView for: $($archs -join ', ')..."

# Clean output folder for deterministic package contents
if (Test-Path $OutputDir) {
    Remove-Item $OutputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $OutputDir | Out-Null

foreach ($item in $targets) {
    $target = $item.Target
    $arch = $item.Arch

    # Install i686 target if needed
    if ($arch -eq "x86") {
        $installedTargets = @(rustup target list --installed)
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to query installed Rust targets"
            exit 1
        }

        if ($installedTargets -notcontains $target) {
            Write-Host "Installing Rust target: $target"
            & rustup target add $target

            if ($LASTEXITCODE -ne 0) {
                Write-Error "Failed to install Rust target: $target"
                exit 1
            }
        }
    }

    $buildCmd = "cargo build --target $target $profile"
    Write-Host $buildCmd
    Invoke-Expression $buildCmd

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build failed"
        exit 1
    }

    $outDir = "target\$target\$profileDir"

    if ($arch -eq "x64") {
        Copy-Item "$outDir\mdview_wlx.dll" "$OutputDir\mdview.wlx64" -Force

        # Keep the x64 exe as mdview64.exe to match distributed release artifact naming
        Copy-Item "$outDir\mdview.exe" "$OutputDir\mdview64.exe" -Force

        if (Test-Path "$outDir\mdview.pdb") {
            Copy-Item "$outDir\mdview.pdb" "$OutputDir\mdview64.pdb" -Force
        }
        if (Test-Path "$outDir\mdview_wlx.pdb") {
            Copy-Item "$outDir\mdview_wlx.pdb" "$OutputDir\mdview_wlx64.pdb" -Force
        }
    } else {
        Copy-Item "$outDir\mdview_wlx.dll" "$OutputDir\mdview.wlx" -Force

        # Keep the x86 exe as mdview.exe (used by plugin installer / legacy naming)
        Copy-Item "$outDir\mdview.exe" "$OutputDir\mdview.exe" -Force

        if (Test-Path "$outDir\mdview.pdb") {
            Copy-Item "$outDir\mdview.pdb" "$OutputDir\mdview.pdb" -Force
        }
        if (Test-Path "$outDir\mdview_wlx.pdb") {
            Copy-Item "$outDir\mdview_wlx.pdb" "$OutputDir\mdview_wlx.pdb" -Force
        }
    }
}

# Include documentation and installer file for release package
Copy-Item "pluginst.inf" "$OutputDir\" -Force
Copy-Item "README.md" "$OutputDir\" -Force
Copy-Item "LICENSE" "$OutputDir\" -Force

Write-Host "\nBuild complete. Output files:" 
Get-ChildItem $OutputDir | Format-Table Name, Length

if ($Package) {
    if (-not $ZipFile) {
        $cargoToml = Get-Content Cargo.toml -Raw
        if ($cargoToml -match 'version\s*=\s*"([^"]+)"') {
            $ZipFile = "MDView-v$($matches[1]).zip"
        } else {
            $ZipFile = "MDView.zip"
        }
    }

    if (Test-Path $ZipFile) {
        Remove-Item $ZipFile -Force
    }

    Compress-Archive -Path "$OutputDir\*" -DestinationPath $ZipFile -CompressionLevel Optimal -Force
    Write-Host "\nPackage created: $ZipFile"
}

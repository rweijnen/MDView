# MDView build script
param(
    [switch]$Release,
    [switch]$x86,
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

if ($Clean) {
    Write-Host "Cleaning build artifacts..."
    cargo clean
    exit 0
}

$target = if ($x86) { "i686-pc-windows-msvc" } else { "x86_64-pc-windows-msvc" }
$profile = if ($Release) { "--release" } else { "" }
$profileDir = if ($Release) { "release" } else { "debug" }

Write-Host "Building MDView for $target..."

# Install target if needed
if ($x86) {
    rustup target add $target 2>$null
}

# Build
$buildCmd = "cargo build --target $target $profile"
Write-Host $buildCmd
Invoke-Expression $buildCmd

if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed"
    exit 1
}

# Copy outputs
$outDir = "target\$target\$profileDir"
$distDir = "dist"

if (!(Test-Path $distDir)) {
    New-Item -ItemType Directory -Path $distDir | Out-Null
}

$arch = if ($x86) { "32" } else { "64" }

Copy-Item "$outDir\mdview_wlx.dll" "$distDir\mdview.wlx$arch" -Force
Copy-Item "$outDir\mdview.exe" "$distDir\mdview$arch.exe" -Force

Write-Host ""
Write-Host "Build complete. Output files:"
Get-ChildItem $distDir | Format-Table Name, Length

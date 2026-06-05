param(
  [string]$PluginName = "research-figure-pipeline",
  [string]$DestinationRoot = (Join-Path $env:USERPROFILE "plugins"),
  [string]$MarketplacePath = (Join-Path $env:USERPROFILE ".agents\plugins\marketplace.json"),
  [switch]$SkipCodexInstall
)

$ErrorActionPreference = "Stop"

$SourceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Destination = Join-Path $DestinationRoot $PluginName

function Copy-PluginSource {
  param(
    [string]$From,
    [string]$To
  )

  $fromResolved = [System.IO.Path]::GetFullPath($From)
  $toResolved = [System.IO.Path]::GetFullPath($To)
  if ($fromResolved.TrimEnd('\') -ieq $toResolved.TrimEnd('\')) {
    return
  }

  New-Item -ItemType Directory -Path $To -Force | Out-Null
  $excludeDirs = @(".git", "__pycache__")
  Get-ChildItem -Path $From -Force | ForEach-Object {
    if ($_.PSIsContainer -and $excludeDirs -contains $_.Name) {
      return
    }
    Copy-Item -Path $_.FullName -Destination $To -Recurse -Force
  }

  Get-ChildItem -Path $To -Recurse -Directory -Force -Filter "__pycache__" |
    Remove-Item -Recurse -Force
  $gitDir = Join-Path $To ".git"
  if (Test-Path $gitDir) {
    Remove-Item -Path $gitDir -Recurse -Force
  }
}

function Ensure-Marketplace {
  param(
    [string]$Path,
    [string]$Name
  )

  $dir = Split-Path -Parent $Path
  New-Item -ItemType Directory -Path $dir -Force | Out-Null

  if (Test-Path $Path) {
    $raw = Get-Content -Path $Path -Raw
    $marketplace = if ($raw.Trim().Length -gt 0) { $raw | ConvertFrom-Json } else { $null }
  } else {
    $marketplace = $null
  }

  if ($null -eq $marketplace) {
    $marketplace = [pscustomobject]@{
      name = "personal"
      interface = [pscustomobject]@{ displayName = "Personal" }
      plugins = @()
    }
  }

  if (-not $marketplace.PSObject.Properties["name"]) {
    $marketplace | Add-Member -NotePropertyName name -NotePropertyValue "personal"
  }
  if (-not $marketplace.PSObject.Properties["interface"]) {
    $marketplace | Add-Member -NotePropertyName interface -NotePropertyValue ([pscustomobject]@{ displayName = "Personal" })
  }
  if (-not $marketplace.PSObject.Properties["plugins"]) {
    $marketplace | Add-Member -NotePropertyName plugins -NotePropertyValue @()
  }

  $entry = [pscustomobject]@{
    name = $Name
    source = [pscustomobject]@{
      source = "local"
      path = "./plugins/$Name"
    }
    policy = [pscustomobject]@{
      installation = "AVAILABLE"
      authentication = "ON_INSTALL"
    }
    category = "Productivity"
  }

  $plugins = @($marketplace.plugins | Where-Object { $_.name -ne $Name })
  $marketplace.plugins = @($plugins + $entry)
  $marketplace | ConvertTo-Json -Depth 10 | Set-Content -Path $Path -Encoding UTF8
}

Copy-PluginSource -From $SourceRoot -To $Destination
Ensure-Marketplace -Path $MarketplacePath -Name $PluginName

Write-Host "Plugin source: $Destination"
Write-Host "Marketplace: $MarketplacePath"

if (-not $SkipCodexInstall) {
  $codex = Get-Command codex -ErrorAction SilentlyContinue
  if ($null -eq $codex) {
    Write-Host "Codex CLI was not found. Run this after installing Codex CLI:"
    Write-Host "codex plugin add $PluginName@personal"
  } else {
    codex plugin add "$PluginName@personal"
  }
} else {
  Write-Host "Skipped Codex install. To install, run:"
  Write-Host "codex plugin add $PluginName@personal"
}

param(
  [string]$PythonVersion = "3.13.14",
  [string]$PandocVersion = "3.10",
  [switch]$WithPlantUml,
  [switch]$WithMermaid,
  [switch]$WithDiagrams
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Tools = Join-Path $Root "tools"
$PythonDir = Join-Path $Tools "python"
$PandocDir = Join-Path $Tools "pandoc"
$NodeDir = Join-Path $Tools "node"
$MermaidDir = Join-Path $Tools "mermaid"
$Temp = Join-Path $Tools "download"

if ($WithDiagrams) {
  $WithPlantUml = $true
  $WithMermaid = $true
}

function Get-CommandPath {
  param([string]$Name)
  $command = Get-Command $Name -ErrorAction SilentlyContinue
  if ($command) {
    return $command.Source
  }
  return $null
}

function Save-Url {
  param(
    [string]$Url,
    [string]$OutFile
  )

  $Curl = Get-CommandPath "curl.exe"
  Remove-Item $OutFile -Force -ErrorAction SilentlyContinue
  if ($Curl) {
    & $Curl -L --fail --retry 3 --connect-timeout 20 --output $OutFile $Url
    if ($LASTEXITCODE -ne 0) {
      throw "Download failed: $Url"
    }
    return
  }

  $PreviousProgressPreference = $ProgressPreference
  try {
    $ProgressPreference = "SilentlyContinue"
    Invoke-WebRequest -Uri $Url -OutFile $OutFile
  } finally {
    $ProgressPreference = $PreviousProgressPreference
  }
}

function Install-PortableNode {
  New-Item -ItemType Directory -Force -Path $Temp | Out-Null

  Write-Host "Resolving latest Node.js LTS..."
  $NodeIndex = Invoke-RestMethod -Uri "https://nodejs.org/dist/index.json"
  $NodeRelease = $NodeIndex |
    Where-Object { $_.lts -ne $false -and $_.files -contains "win-x64-zip" } |
    Select-Object -First 1

  if (-not $NodeRelease) {
    throw "Could not find a Windows x64 LTS Node.js release."
  }

  $NodeVersion = $NodeRelease.version
  $NodeZip = Join-Path $Temp "node.zip"
  $NodeUrl = "https://nodejs.org/dist/$NodeVersion/node-$NodeVersion-win-x64.zip"

  Write-Host "Downloading portable Node.js $NodeVersion..."
  Save-Url $NodeUrl $NodeZip

  Write-Host "Extracting Node.js..."
  $NodeExtract = Join-Path $Temp "node"
  Remove-Item $NodeExtract -Recurse -Force -ErrorAction SilentlyContinue
  Expand-Archive -Path $NodeZip -DestinationPath $NodeExtract -Force
  $ExtractedRoot = Get-ChildItem -Path $NodeExtract -Directory | Select-Object -First 1
  if (-not $ExtractedRoot) {
    throw "Node.js archive did not contain an extracted directory."
  }

  Remove-Item $NodeDir -Recurse -Force -ErrorAction SilentlyContinue
  New-Item -ItemType Directory -Force -Path $NodeDir | Out-Null
  Copy-Item (Join-Path $ExtractedRoot.FullName "*") $NodeDir -Recurse -Force
}

function Install-PortablePython {
  New-Item -ItemType Directory -Force -Path $Temp | Out-Null

  $PythonZip = Join-Path $Temp "python.zip"
  $PythonUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-embed-amd64.zip"

  Write-Host "Downloading portable Python $PythonVersion..."
  Save-Url $PythonUrl $PythonZip

  Write-Host "Extracting Python..."
  Remove-Item $PythonDir -Recurse -Force -ErrorAction SilentlyContinue
  New-Item -ItemType Directory -Force -Path $PythonDir | Out-Null
  Expand-Archive -Path $PythonZip -DestinationPath $PythonDir -Force

  $PythonExe = Join-Path $PythonDir "python.exe"
  if (-not (Test-Path $PythonExe)) {
    throw "python.exe was not found in the downloaded archive."
  }
}

function Write-MermaidLauncher {
  $Launcher = Join-Path $MermaidDir "mmdc.cmd"
  $Content = @"
@echo off
"%~dp0..\node\node.exe" "%~dp0node_modules\@mermaid-js\mermaid-cli\src\cli.js" %*
"@
  Set-Content -Path $Launcher -Value $Content -Encoding ascii
}

New-Item -ItemType Directory -Force -Path $Tools | Out-Null
New-Item -ItemType Directory -Force -Path $Temp | Out-Null

Install-PortablePython

New-Item -ItemType Directory -Force -Path $PandocDir | Out-Null

$PandocZip = Join-Path $Temp "pandoc.zip"
$PandocUrl = "https://github.com/jgm/pandoc/releases/download/$PandocVersion/pandoc-$PandocVersion-windows-x86_64.zip"

Write-Host "Downloading Pandoc $PandocVersion..."
Save-Url $PandocUrl $PandocZip

Write-Host "Extracting Pandoc..."
Expand-Archive -Path $PandocZip -DestinationPath $Temp -Force
$PandocExe = Get-ChildItem -Path $Temp -Filter "pandoc.exe" -Recurse | Select-Object -First 1
if (-not $PandocExe) {
  throw "pandoc.exe was not found in the downloaded archive."
}
Copy-Item $PandocExe.FullName (Join-Path $PandocDir "pandoc.exe") -Force

if ($WithPlantUml) {
  $PlantDir = Join-Path $Tools "plantuml"
  New-Item -ItemType Directory -Force -Path $PlantDir | Out-Null
  $PlantJar = Join-Path $PlantDir "plantuml.jar"
  Write-Host "Downloading PlantUML..."
  Save-Url "https://github.com/plantuml/plantuml/releases/latest/download/plantuml.jar" $PlantJar
}

if ($WithMermaid) {
  New-Item -ItemType Directory -Force -Path $MermaidDir | Out-Null

  $Npm = $null
  $LocalNpm = Join-Path $NodeDir "npm.cmd"
  if (Test-Path $LocalNpm) {
    $Npm = $LocalNpm
  } else {
    $Npm = Get-CommandPath "npm.cmd"
    if (-not $Npm) {
      $Npm = Get-CommandPath "npm"
    }
    if (-not $Npm) {
      Install-PortableNode
      $Npm = $LocalNpm
    }
  }

  if (-not (Test-Path $Npm) -and -not (Get-Command $Npm -ErrorAction SilentlyContinue)) {
    throw "npm was not found. Install Node.js or let this script download portable Node.js."
  }

  $env:PATH = "$NodeDir;$env:PATH"
  Write-Host "Installing Mermaid CLI locally..."
  & $Npm --prefix $MermaidDir install @mermaid-js/mermaid-cli
  Write-MermaidLauncher
}

Remove-Item $Temp -Recurse -Force

Write-Host ""
Write-Host "Tools installed."
Write-Host "Python: $PythonDir\python.exe"
Write-Host "Pandoc: $PandocDir\pandoc.exe"
if ($WithPlantUml) {
  Write-Host "PlantUML: $Tools\plantuml\plantuml.jar"
}
if ($WithMermaid) {
  Write-Host "Mermaid CLI: $MermaidDir\mmdc.cmd"
  if (Test-Path (Join-Path $NodeDir "node.exe")) {
    Write-Host "Portable Node.js: $NodeDir\node.exe"
  }
}

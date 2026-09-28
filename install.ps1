$ErrorActionPreference = "Stop"

$RepositoryRoot = (Resolve-Path -LiteralPath $PSScriptRoot).Path
$MarketplaceName = "grls-product-analyzer-marketplace"
$PluginName = "grls-product-analyzer"
$RequirementsPath = Join-Path $RepositoryRoot "plugins\grls-product-analyzer\requirements.txt"

Write-Host "Checking PDF and XLSB support"
& python -c "import pypdf, pyxlsb" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing document reader dependencies"
    & python -m pip install -r $RequirementsPath
    if ($LASTEXITCODE -ne 0) {
        throw "Could not install document readers from $RequirementsPath."
    }
}

Write-Host "Adding local plugin marketplace from $RepositoryRoot"
& codex plugin marketplace add $RepositoryRoot
if ($LASTEXITCODE -ne 0) {
    Write-Host "Marketplace may already be registered; continuing with plugin installation."
}

Write-Host "Installing $PluginName"
& codex plugin add "$PluginName@$MarketplaceName"
if ($LASTEXITCODE -ne 0) {
    throw "Could not install $PluginName from $MarketplaceName."
}

Write-Host "Installation complete. Restart ChatGPT desktop/Codex and start a new task."

# Launch Harbor Formulas VS Code Extension
$extensionPath = Join-Path $PSScriptRoot "apps\harbor-vscode"
$extensionPath = Resolve-Path $extensionPath

Write-Host "Launching Harbor Formulas extension from: $extensionPath"

# Open VS Code with extension development path
code --extensionDevelopmentPath=$extensionPath --new-window


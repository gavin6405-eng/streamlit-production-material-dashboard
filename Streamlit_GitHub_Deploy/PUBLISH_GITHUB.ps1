param([string]$RepoName = "production-material-dashboard")
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
foreach ($cmd in @("git", "gh")) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) { throw "Install Git and GitHub CLI first: https://cli.github.com/" }
}
gh auth status
if ($LASTEXITCODE -ne 0) { gh auth login --web; if ($LASTEXITCODE -ne 0) { throw "GitHub login failed" } }
if (Test-Path .git) { throw "This folder already has Git history. Use git push for updates; this script creates a NEW repository only." }
git init -b main
if ($LASTEXITCODE -ne 0) { throw "git init failed" }
git add .
if ($LASTEXITCODE -ne 0) { throw "git add failed" }
git -c user.name="Dashboard Publisher" -c user.email="dashboard@users.noreply.github.com" commit -m "Create Streamlit dashboard with persistent PostgreSQL storage"
if ($LASTEXITCODE -ne 0) { throw "git commit failed" }
gh repo create $RepoName --private --source=. --remote=origin --push
if ($LASTEXITCODE -ne 0) { throw "Repository create/push failed. Check repository name and GitHub permissions. Local files remain intact." }
gh repo view --web
Start-Process "https://share.streamlit.io/"
Write-Host "GitHub upload complete. On Streamlit select main / streamlit_app.py, configure Secrets, and Deploy."

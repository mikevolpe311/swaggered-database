# PowerShell helper to create a Python virtual environment for this project
# Usage (PowerShell):
#   .\create_venv.ps1
#
# What it does:
#  - Creates a virtual environment in .venv (if it doesn't exist)
#  - Upgrades pip/setuptools/wheel
#  - Installs packages from requirements.txt (if present)
#  - Shows activation command for current shell

param(
	[switch]$Force,
	[string]$Python = "python"
)

$venvPath = Join-Path -Path $PSScriptRoot -ChildPath '.venv'

if ((Test-Path $venvPath) -and (-not $Force)) {
	Write-Host ".venv already exists. Use -Force to recreate." -ForegroundColor Yellow
} else {
	if ((Test-Path $venvPath) -and $Force) {
		Write-Host "Removing existing .venv..." -ForegroundColor Yellow
		Remove-Item -Recurse -Force $venvPath
	}

	Write-Host "Creating virtual environment at $venvPath" -ForegroundColor Green
	& $Python -m venv $venvPath

	if ($LASTEXITCODE -ne 0) {
		Write-Error "Failed to create virtual environment. Make sure '$Python' is on your PATH and is Python 3.6+."
		exit 1
	}

	$pipPath = Join-Path -Path $venvPath -ChildPath 'Scripts\\pip.exe'
	$pythonPath = Join-Path -Path $venvPath -ChildPath 'Scripts\\python.exe'

	Write-Host "Upgrading pip, setuptools and wheel..." -ForegroundColor Green
	& $pythonPath -m pip install --upgrade pip setuptools wheel

	if (Test-Path (Join-Path $PSScriptRoot 'requirements.txt')) {
		Write-Host "Installing dependencies from requirements.txt..." -ForegroundColor Green
		& $pythonPath -m pip install -r (Join-Path $PSScriptRoot 'requirements.txt')
	} else {
		Write-Host "No requirements.txt found in project root. Skipping package install." -ForegroundColor Yellow
	}

	Write-Host "Virtual environment setup complete." -ForegroundColor Green
	Write-Host "Activate it with (PowerShell):" -ForegroundColor Cyan
	Write-Host "  .\\.venv\\Scripts\\Activate.ps1" -ForegroundColor Cyan
	Write-Host "Or (cmd.exe):" -ForegroundColor Cyan
	Write-Host "  .\\.venv\\Scripts\\activate.bat" -ForegroundColor Cyan
}

Write-Host "Done." -ForegroundColor Green
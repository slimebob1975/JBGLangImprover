# ===================== CONFIG (edit these) ==================================
# Base folders
$TempDir   = 'C:\...'
$DevRoot   = 'C:\...\JBGLangImprover'

# Base Python interpreter
$BasePython = 'C:\...\python.exe'

# Virtual environment
$VenvName  = 'langimprover_venv'
$VenvDir   = Join-Path $TempDir $VenvName
$VenvRequirements = '.\requirements.txt' 

# Executables inside the venv (once created)
$Py        = Join-Path $VenvDir 'Scripts\python.exe'
$Pip       = Join-Path $VenvDir 'Scripts\pip.exe'
$Uvicorn   = Join-Path $VenvDir 'Scripts\uvicorn.exe'
$Activate  = Join-Path $VenvDir 'Scripts\Activate.ps1'
$Deactivate= Join-Path $VenvDir 'Scripts\deactivate'

# Notebook + kernel
$App        = 'app.main:app'
$Port       = 8060
#
# ===================== END CONFIG (do not edit) =============================
#
# Helper: run in a directory (like a temporary cd with pushd/popd)
function Invoke-InDir {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )
    Push-Location $Path
    try   { & $ScriptBlock }
    finally { Pop-Location }
}

# Get current working directory
$CurrentDir = $PWD.Path

# --- Set up virtual environment ---
New-Item -ItemType Directory -Force -Path $TempDir | Out-Null

Set-Location $TempDir
if (-not (Test-Path $Py)) {
    Write-Host "Creating venv: $VenvDir" -ForegroundColor Yellow
    & $BasePython -m venv $VenvDir
}

# Activate the venv for this session
. $Activate

# --- Upgrade pip ---
& $Py -m pip install --upgrade pip

# --- Update web service code to latest version and install modules ---
Invoke-InDir -Path $DevRoot -ScriptBlock {
    git.exe pull
    & $Pip install -r $VenvRequirements
}

Invoke-InDir -Path $DevRoot -ScriptBlock {
    & $Pip install --upgrade uvicorn[standard]
}

# --- Launch uvicorn ---
Invoke-InDir -Path $DevRoot -ScriptBlock {
    & $Uvicorn $App --reload --host 127.0.0.1 --port $Port --log-level debug
}

# --- Deactivate and return to project folder ---
Set-Location $TempDir
& $Deactivate
Set-Location $CurrentDir

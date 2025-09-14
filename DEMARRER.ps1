# ================================
# DEMARRER.ps1 (Windows PowerShell)
# ================================
$ErrorActionPreference = "Stop"

$pyVer = "3.11.0"
$workDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pyDir = Join-Path $workDir "PythonEmbed"
$pyZip = Join-Path $workDir "python-$pyVer-embed-amd64.zip"
$pyUrl = "https://www.python.org/ftp/python/$pyVer/python-$pyVer-embed-amd64.zip"
$interfaceScript = Join-Path $workDir "Interface.py"
$reqFile = Join-Path $workDir "requirements.txt"

Write-Host "=== Vérification Python ==="
$python = (Get-Command python -ErrorAction SilentlyContinue)
if ($python) {
    Write-Host "Python système trouvé : $($python.Path)"
    $pythonPath = "python"
} else {
    Write-Host "Python non trouvé. Utilisation version portable..."
    if (!(Test-Path "$pyDir\python.exe")) {
        Write-Host "Téléchargement Python $pyVer..."
        if (Test-Path $pyZip) { Remove-Item $pyZip -Force }
        Invoke-WebRequest $pyUrl -OutFile $pyZip

        Write-Host "Décompression..."
        if (Test-Path $pyDir) { Remove-Item $pyDir -Recurse -Force }
        New-Item -ItemType Directory -Path $pyDir | Out-Null
        Expand-Archive -LiteralPath $pyZip -DestinationPath $pyDir
    } else {
        Write-Host "Version portable déjà présente."
    }
    $pythonPath = "$pyDir\python.exe"
}

Write-Host "`n=== Vérification pip ==="
try {
    & $pythonPath -m pip --version | Out-Null
} catch {
    Write-Host "Installation pip..."
    try {
        & $pythonPath -m ensurepip --default-pip
    } catch {
        Invoke-WebRequest "https://bootstrap.pypa.io/get-pip.py" -OutFile "$workDir\get-pip.py"
        & $pythonPath "$workDir\get-pip.py"
        Remove-Item "$workDir\get-pip.py" -Force
    }
}

Write-Host "Mise à jour pip/setuptools/wheel..."
& $pythonPath -m pip install --upgrade pip setuptools wheel

if (Test-Path $reqFile) {
    Write-Host "Installation des dépendances requirements.txt..."
    & $pythonPath -m pip install -r $reqFile
}

Write-Host "`n=== Lancement Interface.py ==="
& $pythonPath $interfaceScript

Write-Host "Nettoyage zip..."
if (Test-Path $pyZip) { Remove-Item $pyZip -Force }

Write-Host "`n=== Terminé ==="
Pause

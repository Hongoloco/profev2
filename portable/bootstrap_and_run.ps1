$ErrorActionPreference = 'Stop'

function Write-Step([string]$msg) {
    Write-Host "[planilla] $msg" -ForegroundColor Cyan
}

function Get-SystemPython {
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        return @('py', '-3')
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return @($python.Source)
    }

    return $null
}

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)] [string]$Exe,
        [Parameter(Mandatory = $false)] [string[]]$Args = @()
    )

    & $Exe @Args
    if ($LASTEXITCODE -ne 0) {
        throw "Fallo comando: $Exe $($Args -join ' ')"
    }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Join-Path $scriptDir 'app'
$runtimeDir = Join-Path $scriptDir 'runtime'
$venvDir = Join-Path $runtimeDir 'venv'
$venvPython = Join-Path $venvDir 'Scripts\python.exe'
$embedDir = Join-Path $runtimeDir 'python-embed'
$embedPython = Join-Path $embedDir 'python.exe'
$requirementsPath = Join-Path $appRoot 'requirements.txt'
$requirementsHashFile = Join-Path $runtimeDir 'requirements.sha256'
$appPath = Join-Path $appRoot 'webapp\app.py'
$templatePath = Join-Path $appRoot 'webapp\templates\index.html'

if (-not (Test-Path $appPath)) {
    throw "No se encontro webapp/app.py en $appRoot"
}
if (-not (Test-Path $requirementsPath)) {
    throw "No se encontro requirements.txt en $appRoot"
}
if (-not (Test-Path $templatePath)) {
    throw "No se encontro webapp/templates/index.html en $appRoot"
}

if (-not (Test-Path $runtimeDir)) {
    New-Item -ItemType Directory -Path $runtimeDir | Out-Null
}

$pythonExe = $null
$pythonArgsPrefix = @()

if (Test-Path $venvPython) {
    Write-Step 'Usando entorno virtual local existente.'
    $pythonExe = $venvPython
} else {
    $systemPython = Get-SystemPython
    if ($systemPython) {
        Write-Step 'Python del sistema detectado. Creando entorno virtual portable...'

        if ($systemPython.Count -eq 2) {
            Invoke-Checked -Exe $systemPython[0] -Args @($systemPython[1], '-m', 'venv', $venvDir)
        } else {
            Invoke-Checked -Exe $systemPython[0] -Args @('-m', 'venv', $venvDir)
        }

        if (-not (Test-Path $venvPython)) {
            throw 'No se pudo crear el entorno virtual local.'
        }

        $pythonExe = $venvPython
    } else {
        Write-Step 'No hay Python del sistema. Descargando Python embebido portable...'
        if (-not (Test-Path $embedDir)) {
            New-Item -ItemType Directory -Path $embedDir | Out-Null
        }

        $embedZip = Join-Path $runtimeDir 'python-3.11.9-embed-amd64.zip'
        $pythonUrl = 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip'

        if (-not (Test-Path $embedZip)) {
            Write-Step 'Descargando runtime Python...'
            Invoke-WebRequest -Uri $pythonUrl -OutFile $embedZip
        }

        Write-Step 'Extrayendo runtime Python...'
        Expand-Archive -Path $embedZip -DestinationPath $embedDir -Force

        if (-not (Test-Path $embedPython)) {
            throw 'No se encontro python.exe en runtime embebido.'
        }

        $pthFile = Get-ChildItem -Path $embedDir -Filter 'python*._pth' | Select-Object -First 1
        if ($pthFile) {
            $content = Get-Content -Path $pthFile.FullName -Raw
            $content = $content -replace '#import site', 'import site'
            Set-Content -Path $pthFile.FullName -Value $content -NoNewline
        }

        $getPipPath = Join-Path $runtimeDir 'get-pip.py'
        if (-not (Test-Path $getPipPath)) {
            Write-Step 'Descargando instalador de pip...'
            Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile $getPipPath
        }

        Write-Step 'Instalando pip en runtime portable...'
        Invoke-Checked -Exe $embedPython -Args @($getPipPath)

        $pythonExe = $embedPython
    }
}

Write-Step 'Verificando dependencias Python...'
$currentHash = (Get-FileHash -Path $requirementsPath -Algorithm SHA256).Hash
$storedHash = if (Test-Path $requirementsHashFile) { (Get-Content -Path $requirementsHashFile -Raw).Trim() } else { '' }

if ($currentHash -ne $storedHash) {
    Write-Step 'Instalando/actualizando dependencias...'
    Invoke-Checked -Exe $pythonExe -Args @('-m', 'pip', 'install', '--upgrade', 'pip')
    Invoke-Checked -Exe $pythonExe -Args @('-m', 'pip', 'install', '-r', $requirementsPath)
    Set-Content -Path $requirementsHashFile -Value $currentHash -NoNewline
} else {
    Write-Step 'Dependencias ya estan al dia.'
}

Write-Step 'Iniciando aplicacion web...'
Set-Location -Path $appRoot
& $pythonExe $appPath
exit $LASTEXITCODE

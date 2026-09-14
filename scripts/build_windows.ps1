<#
.SYNOPSIS
    Construye el ejecutable Windows de REMASEP Automation (Sprint 3.11, piloto).

.DESCRIPTION
    Build reproducible con PyInstaller, modo ONEDIR (carpeta, no un único
    .exe). No requiere privilegios de administrador. No toca Excel, no mata
    procesos, no borra carpetas fuera de dist/build de este proyecto.

.PARAMETER DebugBuild
    Construye la variante con consola (packaging/remasep_debug.spec) para
    diagnosticar errores de arranque, en vez de la variante windowed de
    piloto.

.EXAMPLE
    .\scripts\build_windows.ps1

.EXAMPLE
    .\scripts\build_windows.ps1 -DebugBuild
#>

[CmdletBinding()]
param(
    [switch]$DebugBuild
)

$ErrorActionPreference = "Stop"

# --- 1. validar Windows -----------------------------------------------------
if (-not $IsWindows -and [System.Environment]::OSVersion.Platform -ne "Win32NT") {
    Write-Error "Este build sólo corre en Windows (PyInstaller no cross-compila)."
    exit 1
}

$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

# --- 2. validar venv / Python ------------------------------------------------
$VenvPython = Join-Path $RepoRoot ".venv-win\Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    Write-Error @"
No se encontró .venv-win\Scripts\python.exe

Crea el entorno de build primero (una sola vez):
    python -m venv .venv-win
    .venv-win\Scripts\pip install -e .[build]
"@
    exit 1
}

Write-Host "Python de build: $VenvPython"
& $VenvPython -c "import PyInstaller" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Error "PyInstaller no está instalado en .venv-win. Ejecuta:`n    .venv-win\Scripts\pip install -e .[build]"
    exit 1
}

# --- 3. limpiar SOLO dist/build propios de este proyecto (nunca rutas ajenas)
$DistDir = Join-Path $RepoRoot "dist"
$BuildDir = Join-Path $RepoRoot "build"
foreach ($dir in @($DistDir, $BuildDir)) {
    if (Test-Path $dir) {
        Write-Host "Limpiando $dir"
        Remove-Item -Recurse -Force $dir
    }
}

# --- 4. ejecutar PyInstaller con el spec versionado -------------------------
$SpecFile = if ($DebugBuild) { "packaging\remasep_debug.spec" } else { "packaging\remasep.spec" }
$DistName = if ($DebugBuild) { "REMASEP-debug" } else { "REMASEP" }

Write-Host "Construyendo con $SpecFile ..."
& $VenvPython -m PyInstaller $SpecFile --noconfirm --distpath dist --workpath build
if ($LASTEXITCODE -ne 0) {
    Write-Error "PyInstaller terminó con error (código $LASTEXITCODE)."
    exit 1
}

# --- 5. verificar que el .exe exista -----------------------------------------
$ExePath = Join-Path $DistDir "$DistName\$DistName.exe"
if (-not (Test-Path $ExePath)) {
    Write-Error "Build terminó pero no se encontró $ExePath"
    exit 1
}

# --- 6. imprimir ruta final ---------------------------------------------------
$FullPath = (Resolve-Path $ExePath).Path
Write-Host ""
Write-Host "Build OK:" -ForegroundColor Green
Write-Host "  $FullPath"
Write-Host ""
Write-Host "Prueba manual (Fase 11 del sprint):"
Write-Host "  1. Copia toda la carpeta 'dist\$DistName\' a una ubicación fuera del repo"
Write-Host "     (p. ej. Escritorio\REMASEP_PILOT\)."
Write-Host "  2. Abre $DistName.exe con doble clic desde ahí."
Write-Host "  3. No debe requerir .venv-win activado ni Python del sistema."

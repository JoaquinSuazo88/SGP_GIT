# _actualizar_visor.ps1
# Lanzador del visor SGP con soporte para Ctrl+C limpio y commit Git al salir.

$Host.UI.RawUI.WindowTitle = "SGP Visor - Modo Vigilancia"
Set-Location $PSScriptRoot

# Raiz del proyecto (dos niveles arriba: MD/ -> Doc_Funcionales/ -> SGP-Produccion/)
$ProjectRoot = Split-Path (Split-Path $PSScriptRoot)

function Write-Header {
    Write-Host ""
    Write-Host "  =================================================" -ForegroundColor Cyan
    Write-Host "  SGP Upgrade - Generador de Documentacion HTML   " -ForegroundColor Cyan
    Write-Host "  =================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Cada vez que guardes un .md o agregues una nueva"
    Write-Host "  carpeta, el visor_documentos.html se actualiza."
    Write-Host ""
    Write-Host "  Presiona Ctrl+C para detener." -ForegroundColor Yellow
    Write-Host ""
}

function Invoke-GitPush {
    Write-Host ""
    Write-Host "  -------------------------------------------------"
    Write-Host "  Actualizacion del repositorio Git"
    Write-Host "  Carpeta: $ProjectRoot"
    Write-Host "  -------------------------------------------------"
    Write-Host ""

    # Cambiar al directorio raiz del proyecto para los comandos git
    Set-Location $ProjectRoot

    # git add
    Write-Host "  Agregando archivos al stage..." -ForegroundColor Yellow
    git add .
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR: git add fallo. Verifica que esta carpeta" -ForegroundColor Red
        Write-Host "  sea un repositorio Git inicializado." -ForegroundColor Red
        Set-Location $PSScriptRoot
        return
    }
    Write-Host "  OK." -ForegroundColor Green
    Write-Host ""

    # Pedir mensaje (no puede quedar vacio)
    do {
        $msg = Read-Host "  Mensaje del commit"
        if ($msg -eq "") {
            Write-Host "  El mensaje no puede estar vacio." -ForegroundColor Red
        }
    } while ($msg -eq "")

    # git commit
    Write-Host ""
    Write-Host "  Creando commit..." -ForegroundColor Yellow
    git commit -m "$msg"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR: No se pudo crear el commit." -ForegroundColor Red
        Write-Host "  Puede que no haya cambios pendientes para confirmar." -ForegroundColor Red
        Set-Location $PSScriptRoot
        return
    }

    # git push
    Write-Host ""
    Write-Host "  Subiendo cambios a origin/main..." -ForegroundColor Yellow
    git push -u origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR: git push fallo." -ForegroundColor Red
        Write-Host "  Verifica tu conexion y credenciales de Git." -ForegroundColor Red
        Set-Location $PSScriptRoot
        return
    }

    Write-Host ""
    Write-Host "  Repositorio actualizado correctamente." -ForegroundColor Green
    Set-Location $PSScriptRoot
}

# ── Ejecucion principal ──────────────────────────────────────
Write-Header

$continuar = $true

while ($continuar) {

    try {
        node _gen_html.js --watch
    }
    finally {
        # Este bloque siempre se ejecuta, incluso tras Ctrl+C
        Write-Host ""
        Write-Host "  =================================================" -ForegroundColor Cyan
        Write-Host "  Proceso detenido."
        Write-Host "  =================================================" -ForegroundColor Cyan
        Write-Host ""

        $respGit = Read-Host "  Deseas actualizar el repositorio Git? (s/n)"

        if ($respGit -match '^[sS]$') {
            Invoke-GitPush
        } else {
            Write-Host "  Git omitido." -ForegroundColor Gray
        }

        Write-Host ""
        $respWatch = Read-Host "  Deseas seguir vigilando modificaciones? (s/n)"

        if ($respWatch -match '^[sS]$') {
            Write-Host ""
            Write-Host "  Reiniciando modo vigilancia..." -ForegroundColor Cyan
            Write-Host ""
            $continuar = $true
        } else {
            $continuar = $false
        }
    }

}

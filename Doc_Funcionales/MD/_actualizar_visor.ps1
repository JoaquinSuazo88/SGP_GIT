# _actualizar_visor.ps1
# Lanzador del visor SGP — Node corre en segundo plano; teclas controlan el flujo.

$Host.UI.RawUI.WindowTitle = "SGP Visor - Modo Vigilancia"
Set-Location $PSScriptRoot

# Raiz del proyecto (dos niveles arriba: MD/ -> Doc_Funcionales/ -> SGP-Produccion/)
$ProjectRoot = Split-Path (Split-Path $PSScriptRoot)

# ── Encabezado ───────────────────────────────────────────────
function Write-Header {
    Write-Host ""
    Write-Host "  =================================================" -ForegroundColor Cyan
    Write-Host "  SGP Upgrade - Generador de Documentacion HTML   " -ForegroundColor Cyan
    Write-Host "  =================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Vigilando cambios en archivos .md ..."
    Write-Host ""
    Write-Host "  [G]      Gestionar Git y pausar vigilancia" -ForegroundColor Yellow
    Write-Host "  [Q]      Salir sin Git                    " -ForegroundColor Yellow
    Write-Host "  [Ctrl+C] Salir sin Git                    " -ForegroundColor Yellow
    Write-Host ""
}

# ── Git push ─────────────────────────────────────────────────
function Invoke-GitPush {
    Write-Host ""
    Write-Host "  -------------------------------------------------"
    Write-Host "  Actualizacion del repositorio Git"
    Write-Host "  Carpeta: $ProjectRoot"
    Write-Host "  -------------------------------------------------"
    Write-Host ""

    Set-Location $ProjectRoot

    Write-Host "  Agregando archivos al stage..." -ForegroundColor Yellow
    git add .
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR: git add fallo." -ForegroundColor Red
        Set-Location $PSScriptRoot
        return
    }
    Write-Host "  OK." -ForegroundColor Green
    Write-Host ""

    do {
        $msg = Read-Host "  Mensaje del commit"
        if ($msg -eq "") { Write-Host "  El mensaje no puede estar vacio." -ForegroundColor Red }
    } while ($msg -eq "")

    Write-Host ""
    Write-Host "  Creando commit..." -ForegroundColor Yellow
    git commit -m "$msg"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR: No se pudo crear el commit." -ForegroundColor Red
        Set-Location $PSScriptRoot
        return
    }

    Write-Host ""
    Write-Host "  Subiendo cambios a origin/main..." -ForegroundColor Yellow
    git push -u origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR: git push fallo." -ForegroundColor Red
        Set-Location $PSScriptRoot
        return
    }

    Write-Host ""
    Write-Host "  Repositorio actualizado correctamente." -ForegroundColor Green
    Set-Location $PSScriptRoot
}

# ── Iniciar Node como proceso hijo (hereda consola: output visible) ──
function Start-NodeWatch {
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo.FileName        = "node"
    $proc.StartInfo.Arguments       = "_gen_html.js --watch"
    $proc.StartInfo.WorkingDirectory = $PSScriptRoot
    $proc.StartInfo.UseShellExecute = $false   # hereda stdin/stdout/stderr del padre
    $proc.Start() | Out-Null
    return $proc
}

# ── Ejecucion principal ──────────────────────────────────────

# Ctrl+C se convierte en input (char 3) en vez de matar PowerShell.
# Tambien evita que Node reciba SIGINT por Ctrl+C (lo matamos manualmente).
[Console]::TreatControlCAsInput = $true

Write-Header

while ($true) {

    $nodeProc  = Start-NodeWatch
    $accion    = $null

    # Bucle de escucha: espera tecla mientras Node corre
    while (-not $nodeProc.HasExited) {
        if ([Console]::KeyAvailable) {
            $key  = [Console]::ReadKey($true)   # $true = no mostrar la tecla
            $char = [int]$key.KeyChar

            # G → gestionar Git
            if ($char -eq [int][char]'g' -or $char -eq [int][char]'G') {
                $accion = 'git'
                break
            }
            # Q o Ctrl+C (char 3) → salir
            if ($char -eq [int][char]'q' -or $char -eq [int][char]'Q' -or $char -eq 3) {
                $accion = 'salir'
                break
            }
        }
        Start-Sleep -Milliseconds 150
    }

    # Detener Node
    if (-not $nodeProc.HasExited) {
        $nodeProc.Kill()
        $nodeProc.WaitForExit(3000) | Out-Null
    }

    # Si Node salio solo (error o fin inesperado) ir a menu
    if ($null -eq $accion) { $accion = 'git' }

    # ── Salir sin Git ────────────────────────────────────────
    if ($accion -eq 'salir') {
        Write-Host ""
        Write-Host "  Proceso finalizado." -ForegroundColor Gray
        Write-Host ""
        break
    }

    # ── Menu Git ─────────────────────────────────────────────
    Write-Host ""
    Write-Host "  =================================================" -ForegroundColor Cyan
    Write-Host "  Vigilancia pausada."
    Write-Host "  =================================================" -ForegroundColor Cyan
    Write-Host ""

    # Restaurar Ctrl+C normal para que Read-Host funcione correctamente
    [Console]::TreatControlCAsInput = $false

    $respGit = Read-Host "  Deseas actualizar el repositorio Git? (s/n)"
    if ($respGit -match '^[sS]$') {
        Invoke-GitPush
    } else {
        Write-Host "  Git omitido." -ForegroundColor Gray
    }

    Write-Host ""
    $respWatch = Read-Host "  Deseas seguir vigilando modificaciones? (s/n)"

    # Volver a interceptar Ctrl+C antes del proximo ciclo
    [Console]::TreatControlCAsInput = $true

    if ($respWatch -match '^[sS]$') {
        Write-Host ""
        Write-Host "  Reiniciando modo vigilancia..." -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  [G] Gestionar Git y pausar   [Q] Salir" -ForegroundColor Yellow
        Write-Host ""
        # El while externo reinicia Node automaticamente
    } else {
        Write-Host ""
        Write-Host "  Proceso finalizado." -ForegroundColor Gray
        Write-Host ""
        break
    }
}

# Restaurar comportamiento normal de Ctrl+C al cerrar
[Console]::TreatControlCAsInput = $false

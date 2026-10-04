<#
.SYNOPSIS
    FLUXER XZ - PUBLIC RELEASE V26-0.1
    Lanzador, Instalador y Panel de Control Oficial Universal
.DESCRIPTION
    Herramienta agnóstica para Windows (Laptop, Sobremesa, ROG Ally).
    Gestiona el ciclo de vida completo de Fluxer XZ MCP:
    - Despliegue con splash screen e identidad visual oficial.
    - Registro de AppUserModelID (FluxerXZ.Security.V26) y notificaciones nativas con icono oficial.
    - Detección e instalación asistida de Node.js LTS.
    - Configuración no destructiva de clientes MCP (Antigravity y Claude Desktop).
    - Menú de mantenimiento: servidor, actualizaciones, logs, permisos y desinstalación limpia.
#>

[CmdletBinding()]
param(
    [switch]$NoSplash,
    [string]$Action = "",
    [string]$TargetDir = ""
)

$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ------------------------------------------------------------------------------
# CONSTANTES DE IDENTIDAD OFICIAL
# ------------------------------------------------------------------------------
$GLOBAL:AUMID          = "FluxerXZ.Security.V26"
$GLOBAL:BRAND_TITLE    = "Fluxer XZ · Seguridad"
$GLOBAL:DEFAULT_VER    = "V26-0.1"
$GLOBAL:REG_PATH       = "HKCU:\Software\FluxerXZ"

# ------------------------------------------------------------------------------
# RESOLUCION DE RUTAS Y ASSETS
# ------------------------------------------------------------------------------
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }

function Find-AssetFile([string]$fileName) {
    $searchPaths = @(
        (Join-Path $ScriptDir "docs\assets\$fileName"),
        (Join-Path $ScriptDir "assets\$fileName"),
        (Join-Path $ScriptDir "$fileName"),
        (Join-Path (Split-Path -Parent $ScriptDir) "assets\$fileName"),
        (Join-Path (Split-Path -Parent $ScriptDir) "docs\assets\$fileName"),
        (Join-Path $ScriptDir "..\docs\assets\$fileName")
    )
    foreach ($p in $searchPaths) {
        if (Test-Path -LiteralPath $p) { return (Resolve-Path $p).Path }
    }
    return $null
}

$SplashImage = Find-AssetFile "splash.jpg"
$ShieldImage = Find-AssetFile "shield.png"
$IconIco     = Find-AssetFile "icon.ico"

# ------------------------------------------------------------------------------
# PANTALLA DE CARGA VISUAL (SPLASH SCREEN OFICIAL)
# ------------------------------------------------------------------------------
function Show-OfficialSplash([string]$imagePath, [int]$durationMs = 2200) {
    if ($NoSplash -or -not $imagePath -or -not (Test-Path -LiteralPath $imagePath)) { return }
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop

        $img = [System.Drawing.Image]::FromFile($imagePath)
        $form = New-Object System.Windows.Forms.Form
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $form.TopMost = $true
        $form.ShowInTaskbar = $false

        $targetWidth = 620
        $targetHeight = [int]($img.Height * ($targetWidth / $img.Width))
        $form.ClientSize = New-Object System.Drawing.Size($targetWidth, $targetHeight)

        $picBox = New-Object System.Windows.Forms.PictureBox
        $picBox.Dock = [System.Windows.Forms.DockStyle]::Fill
        $picBox.Image = $img
        $picBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
        $form.Controls.Add($picBox)

        # Badge de version y estado
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "FLUXER XZ · PUBLIC RELEASE V26-0.1"
        $lbl.ForeColor = [System.Drawing.Color]::FromArgb(230, 255, 255, 255)
        $lbl.BackColor = [System.Drawing.Color]::FromArgb(210, 10, 15, 25)
        $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $lbl.AutoSize = $true
        $lbl.Location = New-Object System.Drawing.Point(14, $targetHeight - 34)
        $picBox.Controls.Add($lbl)

        $form.Show()
        $form.Refresh()

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while ($sw.ElapsedMilliseconds -lt $durationMs) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 40
        }

        $form.Close()
        $form.Dispose()
        $img.Dispose()
    } catch {
        # Si el entorno es headless o falla WinForms, continua silenciosamente
    }
}

# ------------------------------------------------------------------------------
# LECTURA DE VERSION
# ------------------------------------------------------------------------------
function Get-FluxerVersion([string]$dir) {
    if ($dir -and (Test-Path -LiteralPath (Join-Path $dir "VERSION"))) {
        return (Get-Content (Join-Path $dir "VERSION") -Raw -Encoding UTF8).Trim()
    }
    if ($dir -and (Test-Path -LiteralPath (Join-Path $dir "package.json"))) {
        try {
            $json = Get-Content (Join-Path $dir "package.json") -Raw | ConvertFrom-Json
            if ($json.version) { return "V26-$($json.version.Split('.')[1]).$($json.version.Split('.')[2])" }
        } catch {}
    }
    return $GLOBAL:DEFAULT_VER
}

# ------------------------------------------------------------------------------
# DETECCION DE INSTALACION EN EL SISTEMA
# ------------------------------------------------------------------------------
function Detect-InstalledLocation {
    # 1. ¿El script se esta ejecutando directamente dentro del repo instalado?
    if ((Test-Path -LiteralPath (Join-Path $ScriptDir "server.mjs")) -and (Test-Path -LiteralPath (Join-Path $ScriptDir "core\router.mjs"))) {
        return (Resolve-Path $ScriptDir).Path
    }

    # 2. Registro de Windows (HKCU:\Software\FluxerXZ)
    if (Test-Path -LiteralPath $GLOBAL:REG_PATH) {
        $regVal = (Get-ItemProperty -Path $GLOBAL:REG_PATH -Name "InstallPath" -ErrorAction SilentlyContinue).InstallPath
        if ($regVal -and (Test-Path -LiteralPath (Join-Path $regVal "server.mjs"))) {
            return $regVal
        }
    }

    # 3. Rutas estandar de Documents/Systems
    $docs = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    $candidate1 = Join-Path $docs "Systems\Flux-mcp-doc\Windows MCP\FLUXER XZ"
    if (Test-Path -LiteralPath (Join-Path $candidate1 "server.mjs")) { return $candidate1 }

    $candidate2 = Join-Path $docs "Systems\FLUXER XZ"
    if (Test-Path -LiteralPath (Join-Path $candidate2 "server.mjs")) { return $candidate2 }

    return $null
}

# ------------------------------------------------------------------------------
# DETECCION Y CONFIGURACION DE NODE.JS
# ------------------------------------------------------------------------------
function Get-NodeRuntime {
    $nodeCmd = (Get-Command "node" -ErrorAction SilentlyContinue)
    if ($nodeCmd) {
        try {
            $v = (& node -v).Trim()
            return @{ Path = $nodeCmd.Source; Version = $v; Installed = $true }
        } catch {}
    }

    # Buscar en rutas comunes
    $paths = @(
        "$env:ProgramFiles\nodejs\node.exe",
        "${env:ProgramFiles(x86)}\nodejs\node.exe",
        "$env:LOCALAPPDATA\Programs\node\node.exe"
    )
    foreach ($p in $paths) {
        if (Test-Path -LiteralPath $p) {
            try {
                $v = (& $p -v).Trim()
                return @{ Path = $p; Version = $v; Installed = $true }
            } catch {}
        }
    }

    return @{ Path = $null; Version = $null; Installed = $false }
}

function Ensure-NodeInstalled {
    $nodeInfo = Get-NodeRuntime
    if ($nodeInfo.Installed) {
        Write-Host " [OK] Node.js detectado: $($nodeInfo.Version) ($($nodeInfo.Path))" -ForegroundColor Green
        return $true
    }

    Write-Host ""
    Write-Host " [!] Node.js no se encuentra en el PATH del sistema." -ForegroundColor Yellow
    Write-Host "     Fluxer XZ requiere Node.js LTS (>= 22.0.0) para ejecutarse." -ForegroundColor Gray
    Write-Host ""

    # Buscar instalador local en USB
    $msiCandidates = @(
        (Join-Path $ScriptDir "node_installer.msi"),
        (Join-Path $ScriptDir "..\node_installer.msi"),
        "E:\Carpeta Personal\Systems\node_installer.msi",
        "E:\node_installer.msi"
    )
    $foundMsi = $null
    foreach ($m in $msiCandidates) {
        if (Test-Path -LiteralPath $m) { $foundMsi = (Resolve-Path $m).Path; break }
    }

    if ($foundMsi) {
        Write-Host " [1] Instalar Node.js LTS desde el archivo incluido en la USB:" -ForegroundColor Cyan
        Write-Host "     Ruta: $foundMsi" -ForegroundColor DarkGray
        Write-Host " [2] Instalar automaticamente mediante winget (winget install OpenJS.NodeJS.LTS)" -ForegroundColor Cyan
        Write-Host " [3] Abrir pagina oficial de descarga (https://nodejs.org)" -ForegroundColor Cyan
        Write-Host " [0] Cancelar" -ForegroundColor DarkGray
        Write-Host ""
        $choice = Read-Host " Seleccione una opcion [1-3]"

        switch ($choice) {
            "1" {
                Write-Host " [i] Iniciando instalador MSI de Node.js..." -ForegroundColor Cyan
                Start-Process msiexec.exe -ArgumentList "/i `"$foundMsi`"" -Wait
                $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
                $recheck = Get-NodeRuntime
                return $recheck.Installed
            }
            "2" {
                Write-Host " [i] Ejecutando winget install OpenJS.NodeJS.LTS..." -ForegroundColor Cyan
                winget install OpenJS.NodeJS.LTS --accept-package-agreements --accept-source-agreements
                $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
                $recheck = Get-NodeRuntime
                return $recheck.Installed
            }
            "3" {
                Start-Process "https://nodejs.org/en/download/"
                return $false
            }
            default { return $false }
        }
    } else {
        Write-Host " [1] Instalar via winget (winget install OpenJS.NodeJS.LTS)" -ForegroundColor Cyan
        Write-Host " [2] Abrir pagina oficial de descarga (https://nodejs.org)" -ForegroundColor Cyan
        Write-Host " [0] Cancelar" -ForegroundColor DarkGray
        $choice = Read-Host " Seleccione una opcion [1-2]"
        if ($choice -eq "1") {
            winget install OpenJS.NodeJS.LTS --accept-package-agreements --accept-source-agreements
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
            $recheck = Get-NodeRuntime
            return $recheck.Installed
        } elseif ($choice -eq "2") {
            Start-Process "https://nodejs.org/en/download/"
            return $false
        }
        return $false
    }
}

# ------------------------------------------------------------------------------
# REGISTRO DE APPOUSERMODELID Y ACCESOS DIRECTOS
# ------------------------------------------------------------------------------
function Register-OfficialIdentity([string]$installDir) {
    Write-Host " [*] Registrando AppUserModelID ($GLOBAL:AUMID)..." -ForegroundColor Cyan

    $iconPath = Join-Path $installDir "docs\assets\icon.ico"
    if (-not (Test-Path -LiteralPath $iconPath)) {
        $iconPath = Find-AssetFile "icon.ico"
    }

    # Registro HKCU para AUMID
    $aumidReg = "HKCU:\Software\Classes\AppUserModelId\$GLOBAL:AUMID"
    if (-not (Test-Path -LiteralPath $aumidReg)) {
        New-Item -Path $aumidReg -Force | Out-Null
    }
    Set-ItemProperty -Path $aumidReg -Name "DisplayName" -Value $GLOBAL:BRAND_TITLE -Force
    if ($iconPath) {
        Set-ItemProperty -Path $aumidReg -Name "IconUri" -Value $iconPath -Force
        Set-ItemProperty -Path $aumidReg -Name "IconPath" -Value $iconPath -Force
    }

    # Guardar InstallPath en HKCU:\Software\FluxerXZ
    if (-not (Test-Path -LiteralPath $GLOBAL:REG_PATH)) {
        New-Item -Path $GLOBAL:REG_PATH -Force | Out-Null
    }
    Set-ItemProperty -Path $GLOBAL:REG_PATH -Name "InstallPath" -Value $installDir -Force
    Set-ItemProperty -Path $GLOBAL:REG_PATH -Name "Version" -Value (Get-FluxerVersion $installDir) -Force

    # Crear acceso directo en Menu Inicio
    $startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
    $shortcutPath = Join-Path $startMenuDir "Fluxer XZ.lnk"
    
    $wshShell = New-Object -ComObject WScript.Shell
    $shortcut = $wshShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = (Join-Path $installDir "FLUXER_XZ_LAUNCHER.bat")
    $shortcut.WorkingDirectory = $installDir
    $shortcut.Description = "FLUXER XZ - Consola Oficial y Servidor MCP de Seguridad"
    if ($iconPath) {
        $shortcut.IconLocation = "$iconPath, 0"
    }
    $shortcut.Save()

    Write-Host " [OK] Identidad registrada y acceso directo en Inicio creado." -ForegroundColor Green
}

# ------------------------------------------------------------------------------
# ENVIO DE NOTIFICACION NATIVA DE PRUEBA
# ------------------------------------------------------------------------------
function Send-OfficialToast([string]$title, [string]$message, [string]$installDir) {
    $shieldPath = Join-Path $installDir "docs\assets\shield.png"
    if (-not (Test-Path -LiteralPath $shieldPath)) {
        $shieldPath = Find-AssetFile "shield.png"
    }
    $pngUri = if ($shieldPath) { "file:///$($shieldPath.Replace('\', '/'))" } else { "" }

    $safeTitle = [System.Security.SecurityElement]::Escape($title)
    $safeMsg = [System.Security.SecurityElement]::Escape($message)

    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
        $xmlText = "<toast duration=`"short`"><visual><binding template=`"ToastGeneric`"><text>$safeTitle</text><text>$safeMsg</text><image placement=`"appLogoOverride`" hint-crop=`"circle`" src=`"$pngUri`" /></binding></visual></toast>"
        $xmlDoc = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xmlDoc.LoadXml($xmlText)
        $toast = New-Object Windows.UI.Notifications.ToastNotification $xmlDoc
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($GLOBAL:AUMID)
        $notifier.Show($toast)
    } catch {}
}

# ------------------------------------------------------------------------------
# CONFIGURACION DINAMICA DE CLIENTES MCP (ANTIGRAVITY / CLAUDE)
# ------------------------------------------------------------------------------
function Configure-McpClient([string]$installDir) {
    Write-Host ""
    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host "   CONFIGURACION DE CLIENTES MCP (ASISTENCIA NO DESTRUCTIVA)" -ForegroundColor Cyan
    Write-Host "======================================================================" -ForegroundColor Cyan

    $serverMjsPath = Join-Path $installDir "server.mjs"
    $nodeCmd = "node"

    $mcpBlock = @{
        command = $nodeCmd
        args = @($serverMjsPath.Replace("\", "/"))
        env = @{
            NODE_ENV = "production"
        }
    }

    # 1. Antigravity MCP Config
    $homeDir = $env:USERPROFILE
    $antigravityPath = Join-Path $homeDir ".gemini\antigravity\mcp_config.json"
    $configList = @()

    if (Test-Path -LiteralPath (Split-Path -Parent $antigravityPath)) {
        $configList += @{ Name = "Google Antigravity"; Path = $antigravityPath }
    }

    # 2. Claude Desktop MCP Config
    $claudeDir = Join-Path $env:APPDATA "Claude"
    $claudePath = Join-Path $claudeDir "claude_desktop_config.json"
    if (Test-Path -LiteralPath $claudeDir) {
        $configList += @{ Name = "Claude Desktop"; Path = $claudePath }
    }

    if ($configList.Count -eq 0) {
        Write-Host " [i] No se detectaron clientes MCP preinstalados en las rutas estandar." -ForegroundColor DarkGray
        Write-Host "     Archivo de servidor listo en: $serverMjsPath" -ForegroundColor DarkGray
        return
    }

    foreach ($cfg in $configList) {
        Write-Host ""
        Write-Host " [*] Cliente detectado: $($cfg.Name) ($($cfg.Path))" -ForegroundColor Cyan
        $existingData = @{ mcpServers = @{} }
        if (Test-Path -LiteralPath $cfg.Path) {
            try {
                $raw = Get-Content -LiteralPath $cfg.Path -Raw -Encoding UTF8
                $parsed = $raw | ConvertFrom-Json -AsHashtable
                if ($parsed) { $existingData = $parsed }
                if (-not $existingData.ContainsKey("mcpServers")) { $existingData["mcpServers"] = @{} }
            } catch {
                Write-Host " [!] No se pudo parsear el archivo existente; se creara copia de seguridad." -ForegroundColor Yellow
            }
        }

        # Vista previa del bloque
        Write-Host "     Entrada propuesta:" -ForegroundColor Gray
        $jsonPreview = @{ Fluxer_XZ = $mcpBlock } | ConvertTo-Json -Depth 4
        Write-Host $jsonPreview -ForegroundColor DarkCyan

        $confirm = Read-Host "     ¿Desea registrar/actualizar Fluxer_XZ en $($cfg.Name)? [S/N] (Por defecto: S)"
        if ($confirm -eq "" -or $confirm -match "^[sSyY]") {
            # Crear respaldo
            if (Test-Path -LiteralPath $cfg.Path) {
                Copy-Item -LiteralPath $cfg.Path -Destination "$($cfg.Path).bak" -Force
            }
            $existingData["mcpServers"]["Fluxer_XZ"] = $mcpBlock
            $outJson = $existingData | ConvertTo-Json -Depth 10
            Set-Content -LiteralPath $cfg.Path -Value $outJson -Encoding UTF8 -Force
            Write-Host "     [OK] Registrado exitosamente en $($cfg.Name)." -ForegroundColor Green
        } else {
            Write-Host "     [-] Omitido por eleccion del usuario." -ForegroundColor DarkGray
        }
    }
}

# ------------------------------------------------------------------------------
# FLUJO DE INSTALACION LIMPIA (CUANDO NO ESTA INSTALADO)
# ------------------------------------------------------------------------------
function Run-InstallWorkflow {
    Write-Host ""
    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host "   FLUXER XZ - PUBLIC RELEASE V26-0.1" -ForegroundColor Cyan
    Write-Host "   Instalador Universal y Configurador Autonomo de Sistema" -ForegroundColor Cyan
    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host ""

    # 1. Verificar Node.js
    $nodeOk = Ensure-NodeInstalled
    if (-not $nodeOk) {
        Write-Host " [!] No se puede continuar la instalacion sin Node.js LTS activo." -ForegroundColor Red
        return
    }

    # 2. Determinar origen de los archivos de instalacion
    $zipCandidates = @(
        (Join-Path $ScriptDir "fluxer_xz_v26.0.1.zip"),
        (Join-Path (Split-Path -Parent $ScriptDir) "fluxer_xz_v26.0.1.zip"),
        "E:\Carpeta Personal\Systems\fluxer_xz_v26.0.1.zip",
        "E:\fluxer_xz_v26.0.1.zip"
    )
    $sourceZip = $null
    foreach ($z in $zipCandidates) {
        if (Test-Path -LiteralPath $z) { $sourceZip = (Resolve-Path $z).Path; break }
    }

    # 3. Preguntar directorio destino
    $docs = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    $defaultTarget = Join-Path $docs "Systems\Flux-mcp-doc\Windows MCP\FLUXER XZ"
    
    Write-Host " Directorio sugerido de instalacion:" -ForegroundColor Gray
    Write-Host " $defaultTarget" -ForegroundColor White
    Write-Host ""
    $userPath = Read-Host " Ingrese ruta destino o presione [ENTER] para aceptar la sugerida"
    $finalTarget = if ([string]::IsNullOrWhiteSpace($userPath)) { $defaultTarget } else { $userPath.Trim() }

    Write-Host ""
    Write-Host " [*] Preparando instalacion en: $finalTarget..." -ForegroundColor Cyan

    if (-not (Test-Path -LiteralPath $finalTarget)) {
        New-Item -ItemType Directory -Path $finalTarget -Force | Out-Null
    }

    # Extraccion o copia
    if ($sourceZip) {
        Write-Host " [*] Extrayendo paquete oficial ($sourceZip)..." -ForegroundColor Cyan
        $tarCmd = (Get-Command "tar.exe" -ErrorAction SilentlyContinue)
        if ($tarCmd) {
            & $tarCmd.Source -xf "$sourceZip" -C "$finalTarget"
        } else {
            Expand-Archive -LiteralPath $sourceZip -DestinationPath $finalTarget -Force
        }
    } else {
        # Copia directa de carpetas si se corre desde un clon
        Write-Host " [*] Copiando archivos de aplicacion..." -ForegroundColor Cyan
        $itemsToCopy = @("core", "tools", "updater", "docs", "tests", "package.json", "VERSION", "README.md", "LICENSE", "server.mjs", "server.js", "FLUXER_XZ_LAUNCHER.bat", "FLUXER_XZ_LAUNCHER.ps1", "install.bat")
        foreach ($item in $itemsToCopy) {
            $src = Join-Path $ScriptDir $item
            if (Test-Path -LiteralPath $src) {
                Copy-Item -Path $src -Destination $finalTarget -Recurse -Force
            }
        }
    }

    # Asegurar copia de launchers en la raiz de destino
    Copy-Item -Path (Join-Path $ScriptDir "FLUXER_XZ_LAUNCHER.bat") -Destination (Join-Path $finalTarget "FLUXER_XZ_LAUNCHER.bat") -Force -ErrorAction SilentlyContinue
    Copy-Item -Path (Join-Path $ScriptDir "FLUXER_XZ_LAUNCHER.ps1") -Destination (Join-Path $finalTarget "FLUXER_XZ_LAUNCHER.ps1") -Force -ErrorAction SilentlyContinue
    Copy-Item -Path (Join-Path $ScriptDir "install.bat") -Destination (Join-Path $finalTarget "install.bat") -Force -ErrorAction SilentlyContinue

    # Instalar dependencias si node_modules no esta presente
    if (-not (Test-Path -LiteralPath (Join-Path $finalTarget "node_modules"))) {
        Write-Host " [*] Instalando dependencias de produccion (npm install --omit=dev)..." -ForegroundColor Cyan
        Push-Location $finalTarget
        npm install --omit=dev --no-audit --no-fund
        Pop-Location
    }

    # Registrar identidad
    Register-OfficialIdentity $finalTarget

    # Configurar clientes MCP
    Configure-McpClient $finalTarget

    # Enviar notificacion de bienvenida
    Send-OfficialToast $GLOBAL:BRAND_TITLE "Instalacion completada con exito. Entorno seguro activo." $finalTarget

    Write-Host ""
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host "   [OK] FLUXER XZ V26-0.1 SE INSTALO CORRECTAMENTE EN ESTE EQUIPO" -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host " Se ha creado el acceso directo oficial en el Menu de Inicio." -ForegroundColor Gray
    Write-Host ""
    Read-Host " Presione [ENTER] para abrir el Panel de Control..."
    Run-DashboardMenu $finalTarget
}

# ------------------------------------------------------------------------------
# PANEL DE CONTROL INTERACTIVO (CUANDO YA ESTA INSTALADO)
# ------------------------------------------------------------------------------
function Run-DashboardMenu([string]$installDir) {
    $ver = Get-FluxerVersion $installDir
    $nodeInfo = Get-NodeRuntime

    while ($true) {
        Clear-Host
        Write-Host "======================================================================" -ForegroundColor Cyan
        Write-Host "   FLUXER XZ - PUBLIC RELEASE $ver" -ForegroundColor Cyan
        Write-Host "   Panel de Control Oficial y Herramientas del Sistema" -ForegroundColor Cyan
        Write-Host "======================================================================" -ForegroundColor Cyan
        Write-Host "   Ubicacion activa: $installDir" -ForegroundColor Gray
        Write-Host "   Runtime Node.js:  $($nodeInfo.Version) ($($nodeInfo.Path))" -ForegroundColor Gray
        Write-Host "   Identidad AUMID:  $GLOBAL:AUMID [Activo]" -ForegroundColor Gray
        Write-Host "======================================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  [1] Iniciar Servidor MCP en primer plano (Prueba / Depuracion)" -ForegroundColor White
        Write-Host "  [2] Probar Alerta de Seguridad (Notificacion Nativa con Icono)" -ForegroundColor White
        Write-Host "  [3] Buscar Actualizaciones Remotas (Modulo upd / GitHub)" -ForegroundColor White
        Write-Host "  [4] Inspeccionar Registro de Logs y Seguridad (data/logs)" -ForegroundColor White
        Write-Host "  [5] Gestionar Permisos y Politicas de Seguridad" -ForegroundColor White
        Write-Host "  [6] Reparar Accesos Directos, AUMID y Registro en Clientes MCP" -ForegroundColor White
        Write-Host "  [7] Ejecutar Suite de Pruebas Fisicas (tests/run_all_tests.mjs)" -ForegroundColor White
        Write-Host "  [8] Enviar Feedback / Reportar Incidencia" -ForegroundColor White
        Write-Host "  [9] Desinstalar Limpiamente de este Equipo" -ForegroundColor Yellow
        Write-Host "  [0] Salir" -ForegroundColor DarkGray
        Write-Host ""

        $opt = Read-Host " Seleccione una opcion [0-9]"

        switch ($opt) {
            "1" {
                Write-Host "`n [i] Iniciando servidor MCP stdio (Presione Ctrl+C para detener)..." -ForegroundColor Cyan
                Push-Location $installDir
                node server.mjs
                Pop-Location
                Write-Host "`n Servidor finalizado." -ForegroundColor Gray
                Read-Host " Presione [ENTER] para volver al menu..."
            }
            "2" {
                Write-Host "`n [*] Enviando notificacion nativa de prueba..." -ForegroundColor Cyan
                Send-OfficialToast $GLOBAL:BRAND_TITLE "Alerta de Seguridad: Accion verificada con icono oficial." $installDir
                Write-Host " [OK] Notificacion enviada al Centro de Notificaciones de Windows." -ForegroundColor Green
                Start-Sleep -Seconds 1
            }
            "3" {
                Write-Host "`n [*] Consultando actualizaciones remotas de Fluxer XZ..." -ForegroundColor Cyan
                Push-Location $installDir
                & node updater/updater.mjs --check
                Pop-Location
                Write-Host ""
                Read-Host " Presione [ENTER] para continuar..."
            }
            "4" {
                Write-Host "`n [*] Inspeccionando directorio de logs y auditoria..." -ForegroundColor Cyan
                $logDir = Join-Path $installDir "data\logs"
                if (Test-Path -LiteralPath $logDir) {
                    Get-ChildItem -Path $logDir | Select-Object Name, LastWriteTime, Length | Format-Table -AutoSize
                } else {
                    Write-Host " [i] Directorio data/logs aun no contiene archivos (sin actividad reciente)." -ForegroundColor DarkGray
                }
                Read-Host " Presione [ENTER] para continuar..."
            }
            "5" {
                Write-Host "`n [*] Politicas de Seguridad Fluxer XZ:" -ForegroundColor Cyan
                Write-Host '     - Acciones de riesgo BAJO: Auto-permitidas (files.read, system.info, etc.)' -ForegroundColor Gray
                Write-Host '     - Acciones de riesgo ALTO: Requieren confirmacion interactiva o elevacion' -ForegroundColor Gray
                Write-Host '     - Acciones DESTRUCTIVAS: Requieren confirmacion explicita obligatoria' -ForegroundColor Gray
                Write-Host '     - Rutas criticas bloqueadas: System32, Windows/System, SAM, etc.' -ForegroundColor Gray
                Write-Host ""
                Read-Host " Presione [ENTER] para continuar..."
            }
            "6" {
                Register-OfficialIdentity $installDir
                Configure-McpClient $installDir
                Write-Host " [OK] Reparacion finalizada." -ForegroundColor Green
                Start-Sleep -Seconds 2
            }
            "7" {
                Write-Host "`n [*] Ejecutando suite completa de pruebas fisicas..." -ForegroundColor Cyan
                Push-Location $installDir
                & node tests/run_all_tests.mjs
                Pop-Location
                Read-Host " Presione [ENTER] para continuar..."
            }
            "8" {
                Write-Host "`n [*] Generador de Reporte y Feedback:" -ForegroundColor Cyan
                $fbTitle = Read-Host " Titulo de la incidencia o sugerencia"
                $fbDesc  = Read-Host " Descripcion breve"
                if ($fbTitle) {
                    $fbCombined = "$fbTitle - $fbDesc"
                    $fbScript = 'import("./tools/feedback.mjs").then(async m => { const t = m.createFeedbackTool(); const r = await t.actions.submit_feedback({ type: "bug", description: process.argv[1], tool: "launcher" }); console.log("\n[OK] Feedback guardado:", r.feedbackId); if (r.githubIssueUrl) console.log("[URL] GitHub Issue URL:", r.githubIssueUrl); });'
                    Push-Location $installDir
                    & node --input-type=module -e $fbScript $fbCombined
                    Pop-Location
                }
                Read-Host " Presione [ENTER] para continuar..."
            }
            "9" {
                Write-Host "`n [!] ADVERTENCIA: DESINSTALACION DE FLUXER XZ" -ForegroundColor Red
                $confirm = Read-Host " ¿Esta seguro de que desea desinstalar Fluxer XZ de este equipo? [S/N]"
                if ($confirm -match "^[sSyY]") {
                    # Eliminar acceso directo
                    $lnk = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)) "Microsoft\Windows\Start Menu\Programs\Fluxer XZ.lnk"
                    if (Test-Path -LiteralPath $lnk) { Remove-Item -LiteralPath $lnk -Force }

                    # Limpiar registro
                    if (Test-Path -LiteralPath $GLOBAL:REG_PATH) { Remove-Item -LiteralPath $GLOBAL:REG_PATH -Recurse -Force -ErrorAction SilentlyContinue }
                    $aumidReg = "HKCU:\Software\Classes\AppUserModelId\$GLOBAL:AUMID"
                    if (Test-Path -LiteralPath $aumidReg) { Remove-Item -LiteralPath $aumidReg -Recurse -Force -ErrorAction SilentlyContinue }

                    Write-Host " [OK] Identidad, registros y accesos directos eliminados." -ForegroundColor Green
                    Write-Host " Nota: Los archivos de codigo en $installDir no fueron eliminados automaticamente por seguridad." -ForegroundColor Gray
                    Read-Host " Presione [ENTER] para salir..."
                    exit 0
                }
            }
            "0" {
                exit 0
            }
        }
    }
}

# ------------------------------------------------------------------------------
# PUNTO DE ENTRADA PRINCIPAL
# ------------------------------------------------------------------------------
# 1. Mostrar pantalla de bienvenida visual
if (-not $NoSplash -and $SplashImage) {
    Show-OfficialSplash $SplashImage 2200
}

# 2. Detectar si ya esta instalado
$installedPath = Detect-InstalledLocation

if ($installedPath) {
    Run-DashboardMenu $installedPath
} else {
    Run-InstallWorkflow
}

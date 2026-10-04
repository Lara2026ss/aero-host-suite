/**
 * FLUXER XZ - PUBLIC RELEASE V26-0.1
 * core/notify.mjs — Single Unified Notification and Security UI Module
 */
import path from "node:path";
import { exec, spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ASSETS_DIR = path.resolve(__dirname, "../docs/assets");
const ICON_ICO = path.join(ASSETS_DIR, "icon.ico");
const SHIELD_PNG = path.join(ASSETS_DIR, "shield.png");

export const AUMID = "FluxerXZ.Security.V26";
export const SECURITY_TITLE = "Fluxer XZ · Seguridad";

/**
 * Send a native Windows desktop toast notification with official AppUserModelID and Shield icon.
 * @param {string} title
 * @param {string} message
 * @param {object} options
 */
export async function sendNativeNotification(title = SECURITY_TITLE, message = "", options = {}) {
  const safeTitle = String(title || SECURITY_TITLE).replace(/'/g, "''").replace(/"/g, '""');
  const safeMessage = String(message || "").replace(/'/g, "''").replace(/"/g, '""');
  const pngUri = `file:///${SHIELD_PNG.replace(/\\/g, "/")}`;

  const psScript = `
    try {
      [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
      [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
      
      $xmlTemplate = @"
<toast duration="short">
    <visual>
        <binding template="ToastGeneric">
            <text>${safeTitle}</text>
            <text>${safeMessage}</text>
            <image placement="appLogoOverride" hint-crop="circle" src="${pngUri}" />
        </binding>
    </visual>
</toast>
"@
      $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
      $xml.LoadXml($xmlTemplate)
      $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
      $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("${AUMID}")
      $notifier.Show($toast)
      exit 0
    } catch {
      exit 1
    }
  `;

  return new Promise((resolve) => {
    const child = spawn("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", psScript], {
      windowsHide: true,
      stdio: "ignore"
    });
    child.on("close", (code) => {
      resolve(code === 0);
    });
    child.on("error", () => resolve(false));
  });
}

/**
 * Interactive security confirmation dialog for High and Destructive risk actions.
 * Presents a clear, friendly native dialog with 3 explicit choices:
 * - "Permitir una vez" (allow_once)
 * - "Permitir esta sesión" (allow_session)
 * - "Rechazar" (deny)
 * 
 * @param {object} details
 * @returns {Promise<"allow_once" | "allow_session" | "deny">}
 */
export async function promptSecurityConfirmation({ tool, action, resource, reason, risk = "high" }) {
  const isDestructive = risk === "destructive";
  const riskLabel = isDestructive ? "DESTRUCTIVO" : "ALTO";

  const psDialog = `
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "${SECURITY_TITLE}"
$form.Size = New-Object System.Drawing.Size(480, 260)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.TopMost = $true
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

if (Test-Path "${ICON_ICO.replace(/\\/g, "\\\\")}") {
    try { $form.Icon = New-Object System.Drawing.Icon("${ICON_ICO.replace(/\\/g, "\\\\")}") } catch {}
}

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Location = New-Object System.Drawing.Point(20, 15)
$lblTitle.Size = New-Object System.Drawing.Size(430, 24)
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = if ("${risk}" -eq "destructive") { [System.Drawing.Color]::DarkRed } else { [System.Drawing.Color]::DarkBlue }
$lblTitle.Text = "Solicitud de Autorización · Nivel ${riskLabel}"
$form.Controls.Add($lblTitle)

$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Location = New-Object System.Drawing.Point(20, 45)
$lblDesc.Size = New-Object System.Drawing.Size(430, 95)
$lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblDesc.Text = "La IA solicita ejecutar la acción '${tool}.${action}'.\`nRecurso: ${resource || 'Sistema local'}\`nMotivo: ${reason || 'Operación solicitada por el usuario'}\`n\`nSi rechaza esta solicitud, la acción se cancelará sin afectar el sistema."
$form.Controls.Add($lblDesc)

$btnOnce = New-Object System.Windows.Forms.Button
$btnOnce.Location = New-Object System.Drawing.Point(20, 160)
$btnOnce.Size = New-Object System.Drawing.Size(130, 35)
$btnOnce.Text = "Permitir una vez"
$btnOnce.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnOnce.Add_Click({ $form.Tag = "allow_once"; $form.Close() })
$form.Controls.Add($btnOnce)

$btnSession = New-Object System.Windows.Forms.Button
$btnSession.Location = New-Object System.Drawing.Point(160, 160)
$btnSession.Size = New-Object System.Drawing.Size(150, 35)
$btnSession.Text = "Permitir esta sesión"
$btnSession.Font = New-Object System.Drawing.Font("Segoe UI", 9)
if ("${risk}" -eq "destructive") {
    $btnSession.Enabled = $false
    $btnSession.Text = "(No disponible)"
}
$btnSession.Add_Click({ $form.Tag = "allow_session"; $form.Close() })
$form.Controls.Add($btnSession)

$btnDeny = New-Object System.Windows.Forms.Button
$btnDeny.Location = New-Object System.Drawing.Point(320, 160)
$btnDeny.Size = New-Object System.Drawing.Size(125, 35)
$btnDeny.Text = "Rechazar"
$btnDeny.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnDeny.ForeColor = [System.Drawing.Color]::Firebrick
$btnDeny.Add_Click({ $form.Tag = "deny"; $form.Close() })
$form.Controls.Add($btnDeny)

$form.ShowDialog() | Out-Null
Write-Output ($form.Tag ?? "deny")
  `;

  return new Promise((resolve) => {
    // If running in headless/CI test mode, check env FLUXER_TEST_HEADLESS_DECISION
    if (process.env.FLUXER_TEST_HEADLESS_DECISION) {
      return resolve(process.env.FLUXER_TEST_HEADLESS_DECISION);
    }

    exec(`powershell.exe -NoProfile -Command "${psDialog.replace(/"/g, '`"')}"`, (err, stdout) => {
      const decision = stdout.trim().toLowerCase();
      if (["allow_once", "allow_session", "deny"].includes(decision)) {
        resolve(decision);
      } else {
        resolve("deny");
      }
    });
  });
}

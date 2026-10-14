<#
.SYNOPSIS
    Aero Host Suite — tools/documents/office-com.ps1
    Windows Office COM Automation Engine (Word, PowerPoint, Excel).
    Executes in LIVE mode (GetActiveObject) or HEADLESS mode (New-Object).
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$PayloadFile
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"
Add-Type -ReferencedAssemblies @("Microsoft.CSharp", "System.Drawing") -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class LiveAccessibleHelper {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern IntPtr OpenDesktop(string lpszDesktop, uint dwFlags, bool fInherit, uint dwDesiredAccess);
    [DllImport("user32.dll")] public static extern bool SetThreadDesktop(IntPtr hDesktop);
    [DllImport("user32.dll")] public static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumWindowsProc lpfn, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

    public static dynamic GetApp(string targetClass, string childClass) {
        try {
            IntPtr hDesk = OpenDesktop("Default", 0, false, 0x01FF);
            if (hDesk != IntPtr.Zero) {
                SetThreadDesktop(hDesk);
            }
            dynamic appFound = null;
            EnumDesktopWindows(hDesk, (hWnd, lp) => {
                StringBuilder cls = new StringBuilder(256);
                GetClassName(hWnd, cls, 256);
                if (cls.ToString() == targetClass) {
                    IntPtr targetHwnd = hWnd;
                    if (!string.IsNullOrEmpty(childClass)) {
                        EnumChildWindows(hWnd, (chHwnd, lp2) => {
                            StringBuilder chCls = new StringBuilder(256);
                            GetClassName(chHwnd, chCls, 256);
                            if (chCls.ToString() == childClass) {
                                Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
                                uint OBJID_NATIVEOM = 0xFFFFFFF0;
                                object winObj;
                                int hr = AccessibleObjectFromWindow(chHwnd, OBJID_NATIVEOM, ref IID_IDispatch, out winObj);
                                if (hr == 0 && winObj != null) {
                                    dynamic win = winObj;
                                    appFound = win.Application;
                                    return false; // stop child enum
                                }
                            }
                            return true;
                        }, IntPtr.Zero);
                    } else {
                        Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
                        uint OBJID_NATIVEOM = 0xFFFFFFF0;
                        object winObj;
                        int hr = AccessibleObjectFromWindow(hWnd, OBJID_NATIVEOM, ref IID_IDispatch, out winObj);
                        if (hr == 0 && winObj != null) {
                            dynamic win = winObj;
                            appFound = win.Application;
                        }
                    }
                    if (appFound != null) return false; // stop window enum
                }
                return true;
            }, IntPtr.Zero);

            return appFound;
        } catch {}
        return null;
    }

    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBmp, uint nFlags);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    public static string CaptureWindow(string targetClass, string savePath) {
        try {
            IntPtr hDesk = OpenDesktop("Default", 0, false, 0x01FF);
            if (hDesk != IntPtr.Zero) SetThreadDesktop(hDesk);
            IntPtr targetHwnd = IntPtr.Zero;
            EnumDesktopWindows(hDesk, (hWnd, lp) => {
                StringBuilder cls = new StringBuilder(256);
                GetClassName(hWnd, cls, 256);
                if (cls.ToString() == targetClass) {
                    StringBuilder txt = new StringBuilder(256);
                    GetWindowText(hWnd, txt, 256);
                    bool vis = IsWindowVisible(hWnd);
                    bool iconic = IsIconic(hWnd);
                    if (txt.Length > 0) {
                        targetHwnd = hWnd;
                        return false; // Found titled document window!
                    }
                    if (vis || iconic) {
                        targetHwnd = hWnd;
                    }
                }
                return true;
            }, IntPtr.Zero);

            if (targetHwnd == IntPtr.Zero) return "Target window not found for class: " + targetClass;

            if (IsIconic(targetHwnd)) {
                ShowWindow(targetHwnd, 9); // SW_RESTORE
                System.Threading.Thread.Sleep(300);
            } else {
                RECT tempR;
                GetWindowRect(targetHwnd, out tempR);
                if ((tempR.Right - tempR.Left) < 300) {
                    ShowWindow(targetHwnd, 9);
                    System.Threading.Thread.Sleep(300);
                }
            }
            SetForegroundWindow(targetHwnd);
            System.Threading.Thread.Sleep(200);

            RECT rect;
            GetWindowRect(targetHwnd, out rect);
            int w = rect.Right - rect.Left;
            int h = rect.Bottom - rect.Top;
            if (w <= 0 || h <= 0) return "Invalid window size";

            using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb)) {
                using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp)) {
                    IntPtr hdc = g.GetHdc();
                    bool ok = PrintWindow(targetHwnd, hdc, 2);
                    g.ReleaseHdc(hdc);
                    if (!ok) {
                        hdc = g.GetHdc();
                        PrintWindow(targetHwnd, hdc, 0);
                        g.ReleaseHdc(hdc);
                    }
                }
                bmp.Save(savePath, System.Drawing.Imaging.ImageFormat.Png);
            }
            return "OK";
        } catch (Exception ex) {
            return "Capture error: " + ex.Message;
        }
    }
}
"@ -ErrorAction SilentlyContinue

function ConvertTo-CleanJson($obj) {
    return ($obj | ConvertTo-Json -Depth 10 -Compress)
}

function Get-RunningApp($progId, $docPath = "") {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        if ($null -ne $app) { return $app }
    } catch {}

    if ($docPath -and (Test-Path $docPath -ErrorAction SilentlyContinue)) {
        try {
            $bound = [System.Runtime.InteropServices.Marshal]::BindToMoniker($docPath)
            if ($null -ne $bound) {
                if ($bound.Application) { return $bound.Application }
                return $bound
            }
        } catch {}
    }

    # Accessible desktop fallback for sandboxed/cross-desktop sessions
    try {
        if ($progId -eq "Word.Application") {
            $app = [LiveAccessibleHelper]::GetApp("OpusApp", "_WwG")
            if ($null -ne $app) { return $app }
        } elseif ($progId -eq "PowerPoint.Application") {
            $app = [LiveAccessibleHelper]::GetApp("PPTFrameClass", "paneClassDC")
            if ($null -ne $app) { return $app }
        } elseif ($progId -eq "Excel.Application") {
            $app = [LiveAccessibleHelper]::GetApp("XLMAIN", "EXCEL7")
            if ($null -ne $app) { return $app }
        }
    } catch {}

    # COM instance fallback when Office process is running
    try {
        $procName = switch ($progId) {
            "Word.Application" { "WINWORD" }
            "PowerPoint.Application" { "POWERPNT" }
            "Excel.Application" { "EXCEL" }
            default { "" }
        }
        if ($procName -and (Get-Process $procName -ErrorAction SilentlyContinue)) {
            $app = New-Object -ComObject $progId
            if ($null -ne $app) { return $app }
        }
    } catch {}

    return $null
}

function Release-Com($comObj) {
    if ($null -ne $comObj) {
        try {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObj) | Out-Null
        } catch {
            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObj) | Out-Null
            } catch {}
        }
    }
}

function Invoke-GarbageCollection {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
}

function Invoke-ComWithRetry($scriptBlock, $maxTries = 3, $delayMs = 500) {
    $attempt = 0
    while ($attempt -lt $maxTries) {
        $attempt++
        try {
            return & $scriptBlock
        } catch {
            $hres = $_.Exception.HResult
            $isRetryable = ($hres -eq [int]0x8001010A -or $hres -eq -2147417846 -or $_.Exception.Message -match "0x8001010A|RPC_E_SERVERCALL_RETRYLATER|rejected by callee|servidor ocupado")
            if ($isRetryable -and $attempt -lt $maxTries) {
                Start-Sleep -Milliseconds $delayMs
                continue
            }
            throw $_
        }
    }
}

function Normalize-DocPath($p) {
    if (-not $p) { return "" }
    if ($p -match "^https?://") { return $p }
    try {
        return [System.IO.Path]::GetFullPath($p)
    } catch {
        return $p
    }
}

function Find-DocInCollection($app, $targetPath, $appType = "word") {
    if ($null -eq $app) { return $null }
    
    # Auto-spawn a blank visible document if app is running but has no documents open
    try {
        if ($appType -eq "word" -and $app.Documents.Count -eq 0) {
            $newDoc = $app.Documents.Add()
            $app.Visible = $true
            return $newDoc
        }
        if ($appType -eq "powerpoint" -and $app.Presentations.Count -eq 0) {
            $newPres = $app.Presentations.Add(1)
            $app.Visible = 1
            return $newPres
        }
        if ($appType -eq "excel" -and $app.Workbooks.Count -eq 0) {
            $newWb = $app.Workbooks.Add()
            $app.Visible = $true
            return $newWb
        }
    } catch {}

    # If no target path specified, use the currently active document
    if (-not $targetPath) {
        try {
            if ($appType -eq "word") { return $app.ActiveDocument }
            if ($appType -eq "powerpoint") { return $app.ActivePresentation }
            if ($appType -eq "excel") { return $app.ActiveWorkbook }
        } catch {}
    }
    
    $normTarget = Normalize-DocPath $targetPath
    $targetName = [System.IO.Path]::GetFileName($targetPath)
    
    $items = switch ($appType) {
        "word" { $app.Documents }
        "powerpoint" { $app.Presentations }
        "excel" { $app.Workbooks }
    }
    
    foreach ($item in $items) {
        try {
            $itemFull = Normalize-DocPath $item.FullName
            $itemName = $item.Name
            if ($itemFull -eq $normTarget -or $itemName -eq $targetName -or $itemName -eq $targetPath) {
                return $item
            }
        } catch {}
    }
    
    # Fallback to active document if any exists
    try {
        if ($appType -eq "word" -and $app.Documents.Count -gt 0) { return $app.ActiveDocument }
        if ($appType -eq "powerpoint" -and $app.Presentations.Count -gt 0) { return $app.ActivePresentation }
        if ($appType -eq "excel" -and $app.Workbooks.Count -gt 0) { return $app.ActiveWorkbook }
    } catch {}
    
    return $null
}

# --- Action Implementations ---

function Invoke-Status($params) {
    $appsToCheck = @(
        @{ name = "word"; progId = "Word.Application"; procName = "WINWORD" },
        @{ name = "powerpoint"; progId = "PowerPoint.Application"; procName = "POWERPNT" },
        @{ name = "excel"; progId = "Excel.Application"; procName = "EXCEL" }
    )
    
    $results = @{}
    foreach ($item in $appsToCheck) {
        $name = $item.name
        $installed = $false
        $type = [Type]::GetTypeFromProgID($item.progId)
        if ($null -ne $type) { $installed = $true }
        
        $procs = Get-Process -Name $item.procName -ErrorAction SilentlyContinue
        $running = ($procs.Count -gt 0)
        
        $liveCom = $false
        $comApp = Get-RunningApp $item.progId
        if ($null -ne $comApp) {
            $liveCom = $true
            Release-Com $comApp
            $comApp = $null
        }
        
        $results[$name] = @{
            installed = $installed
            running = $running
            processCount = $procs.Count
            liveComAccessible = $liveCom
        }
    }
    Invoke-GarbageCollection
    return @{ ok = $true; status = $results }
}

function Invoke-ListOpen($params) {
    $appFilter = if ($params.app) { $params.app.ToLower() } else { "all" }
    $openDocs = @()
    
    # Word
    if ($appFilter -eq "all" -or $appFilter -eq "word") {
        $word = Get-RunningApp "Word.Application"
        if ($null -ne $word) {
            try {
                $activeDocName = ""
                try { $activeDocName = $word.ActiveDocument.Name } catch {}
                foreach ($doc in $word.Documents) {
                    $openDocs += @{
                        app = "word"
                        title = $doc.Name
                        path = $doc.FullName
                        saved = $doc.Saved
                        active = ($doc.Name -eq $activeDocName)
                        paragraphs = $doc.Paragraphs.Count
                    }
                }
            } catch {}
            finally {
                Release-Com $word
            }
        }
    }
    
    # PowerPoint
    if ($appFilter -eq "all" -or $appFilter -eq "powerpoint") {
        $ppt = Get-RunningApp "PowerPoint.Application"
        if ($null -ne $ppt) {
            try {
                $activePresName = ""
                try { $activePresName = $ppt.ActivePresentation.Name } catch {}
                foreach ($pres in $ppt.Presentations) {
                    $openDocs += @{
                        app = "powerpoint"
                        title = $pres.Name
                        path = $pres.FullName
                        saved = $pres.Saved
                        active = ($pres.Name -eq $activePresName)
                        slides = $pres.Slides.Count
                    }
                }
            } catch {}
            finally {
                Release-Com $ppt
            }
        }
    }
    
    # Excel
    if ($appFilter -eq "all" -or $appFilter -eq "excel") {
        $excel = Get-RunningApp "Excel.Application"
        if ($null -ne $excel) {
            try {
                $activeWbName = ""
                try { $activeWbName = $excel.ActiveWorkbook.Name } catch {}
                foreach ($wb in $excel.Workbooks) {
                    $openDocs += @{
                        app = "excel"
                        title = $wb.Name
                        path = $wb.FullName
                        saved = $wb.Saved
                        active = ($wb.Name -eq $activeWbName)
                        sheets = $wb.Sheets.Count
                    }
                }
            } catch {}
            finally {
                Release-Com $excel
            }
        }
    }
    
    Invoke-GarbageCollection
    return @{ ok = $true; count = $openDocs.Count; documents = $openDocs }
}

function Invoke-GetActive($params) {
    $appType = if ($params.app) { $params.app.ToLower() } else { "word" }
    
    if ($appType -eq "word") {
        $word = Get-RunningApp "Word.Application"
        if ($null -eq $word) { return @{ ok = $false; error = "Word is not running or no active instance found" } }
        try {
            $doc = $word.ActiveDocument
            $res = @{
                ok = $true
                app = "word"
                title = $doc.Name
                path = $doc.FullName
                saved = $doc.Saved
                paragraphs = $doc.Paragraphs.Count
            }
            Release-Com $doc
            return $res
        } catch {
            return @{ ok = $false; error = "No active document in Word" }
        } finally {
            Release-Com $word
            Invoke-GarbageCollection
        }
    }
    
    if ($appType -eq "powerpoint") {
        $ppt = Get-RunningApp "PowerPoint.Application"
        if ($null -eq $ppt) { return @{ ok = $false; error = "PowerPoint is not running" } }
        try {
            $pres = $ppt.ActivePresentation
            $res = @{
                ok = $true
                app = "powerpoint"
                title = $pres.Name
                path = $pres.FullName
                saved = $pres.Saved
                slides = $pres.Slides.Count
            }
            Release-Com $pres
            return $res
        } catch {
            return @{ ok = $false; error = "No active presentation in PowerPoint" }
        } finally {
            Release-Com $ppt
            Invoke-GarbageCollection
        }
    }
    
    if ($appType -eq "excel") {
        $excel = Get-RunningApp "Excel.Application"
        if ($null -eq $excel) { return @{ ok = $false; error = "Excel is not running" } }
        try {
            $wb = $excel.ActiveWorkbook
            $res = @{
                ok = $true
                app = "excel"
                title = $wb.Name
                path = $wb.FullName
                saved = $wb.Saved
                sheets = $wb.Sheets.Count
            }
            Release-Com $wb
            return $res
        } catch {
            return @{ ok = $false; error = "No active workbook in Excel" }
        } finally {
            Release-Com $excel
            Invoke-GarbageCollection
        }
    }
    
    return @{ ok = $false; error = "Unsupported app: $appType" }
}

function Invoke-Create($params) {
    $appType = if ($params.app) { $params.app.ToLower() } else { "word" }
    $targetPath = if ($params.path) { Normalize-DocPath $params.path } else { "" }
    $visible = if ($null -ne $params.visible) { [bool]$params.visible } else { $true }
    $overwrite = if ($null -ne $params.overwrite) { [bool]$params.overwrite } else { $false }
    
    if ($targetPath -and (Test-Path $targetPath) -and (-not $overwrite)) {
        return @{ ok = $false; error = "File already exists at target path: $targetPath. Pass overwrite=true to overwrite." }
    }
    
    if ($targetPath) {
        $parentDir = [System.IO.Path]::GetDirectoryName($targetPath)
        if (-not (Test-Path $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }
    }
    
    if ($appType -eq "word") {
        $word = Get-RunningApp "Word.Application"
        $wasRunning = ($null -ne $word)
        if (-not $wasRunning) {
            $word = New-Object -ComObject Word.Application
        }
        $doc = $null
        try {
            $word.Visible = $visible
            
            # Smart reuse: If Word already has an empty, unsaved document open (like Documento1 with 0-1 chars)
            # and no specific template or forceNew requested, REUSE IT instead of opening another window!
            $reuseDoc = $false
            if ($wasRunning -and (-not $params.forceNew) -and (-not $params.template) -and (-not $targetPath)) {
                try {
                    if ($word.Documents.Count -gt 0) {
                        $active = $word.ActiveDocument
                        if ($active.Path -eq "" -and $active.Characters.Count -le 1) {
                            $doc = $active
                            $reuseDoc = $true
                        }
                    }
                } catch {}
            }
            
            if (-not $reuseDoc) {
                $doc = if ($params.template) { $word.Documents.Add($params.template) } else { $word.Documents.Add() }
            }
            $title = $doc.Name
            
            if ($targetPath) {
                $doc.SaveAs2($targetPath)
                $title = [System.IO.Path]::GetFileName($targetPath)
                if (-not $visible -and -not $wasRunning) {
                    try { $doc.Close([ref]-1) } catch {}
                    Release-Com $doc
                    $doc = $null
                    try { $word.Quit() } catch {}
                    Release-Com $word
                    $word = $null
                }
            }
            $finalPath = if ($doc -and $doc.FullName) { $doc.FullName } else { $targetPath }
            return @{ ok = $true; app = "word"; path = $finalPath; title = $title; liveMode = $wasRunning; reused = $reuseDoc }
        } catch {
            if (-not $wasRunning) {
                if ($null -ne $doc) { try { $doc.Close([ref]0) } catch {} }
                if ($null -ne $word) { try { $word.Quit() } catch {} }
            }
            throw $_
        } finally {
            if ($null -ne $doc) { Release-Com $doc }
            if ($null -ne $word) { Release-Com $word }
            Invoke-GarbageCollection
        }
    }
    
    if ($appType -eq "powerpoint") {
        $ppt = Get-RunningApp "PowerPoint.Application"
        $wasRunning = ($null -ne $ppt)
        if (-not $wasRunning) {
            $ppt = New-Object -ComObject PowerPoint.Application
        }
        $pres = $null
        try {
            $ppt.Visible = 1
            
            # Smart reuse for PowerPoint
            $reusePres = $false
            if ($wasRunning -and (-not $params.forceNew) -and (-not $params.template) -and (-not $targetPath)) {
                try {
                    if ($ppt.Presentations.Count -gt 0) {
                        $active = $ppt.ActivePresentation
                        if ($active.Path -eq "" -and $active.Slides.Count -le 1) {
                            $pres = $active
                            $reusePres = $true
                        }
                    }
                } catch {}
            }
            
            if (-not $reusePres) {
                $pres = if ($params.template) { $ppt.Presentations.Open($params.template) } else { $ppt.Presentations.Add(1) }
            }
            $title = $pres.Name
            
            if ($targetPath) {
                $pres.SaveAs($targetPath)
                $title = [System.IO.Path]::GetFileName($targetPath)
                if (-not $visible -and -not $wasRunning) {
                    try { $pres.Close() } catch {}
                    Release-Com $pres
                    $pres = $null
                    try { $ppt.Quit() } catch {}
                    Release-Com $ppt
                    $ppt = $null
                }
            }
            $finalPath = if ($pres -and $pres.FullName) { $pres.FullName } else { $targetPath }
            return @{ ok = $true; app = "powerpoint"; path = $finalPath; title = $title; liveMode = $wasRunning; reused = $reusePres }
        } catch {
            if (-not $wasRunning) {
                if ($null -ne $pres) { try { $pres.Close() } catch {} }
                if ($null -ne $ppt) { try { $ppt.Quit() } catch {} }
            }
            throw $_
        } finally {
            if ($null -ne $pres) { Release-Com $pres }
            if ($null -ne $ppt) { Release-Com $ppt }
            Invoke-GarbageCollection
        }
    }
    
    if ($appType -eq "excel") {
        $excel = Get-RunningApp "Excel.Application"
        $wasRunning = ($null -ne $excel)
        if (-not $wasRunning) {
            $excel = New-Object -ComObject Excel.Application
        }
        $wb = $null
        try {
            $excel.Visible = $visible
            
            # Smart reuse for Excel
            $reuseWb = $false
            if ($wasRunning -and (-not $params.forceNew) -and (-not $params.template) -and (-not $targetPath)) {
                try {
                    if ($excel.Workbooks.Count -gt 0) {
                        $active = $excel.ActiveWorkbook
                        if ($active.Path -eq "") {
                            $wb = $active
                            $reuseWb = $true
                        }
                    }
                } catch {}
            }
            
            if (-not $reuseWb) {
                $wb = if ($params.template) { $excel.Workbooks.Add($params.template) } else { $excel.Workbooks.Add() }
            }
            $title = $wb.Name
            
            if ($targetPath) {
                $wb.SaveAs($targetPath)
                $title = [System.IO.Path]::GetFileName($targetPath)
                if (-not $visible -and -not $wasRunning) {
                    try { $wb.Close($true) } catch {}
                    Release-Com $wb
                    $wb = $null
                    try { $excel.Quit() } catch {}
                    Release-Com $excel
                    $excel = $null
                }
            }
            $finalPath = if ($wb -and $wb.FullName) { $wb.FullName } else { $targetPath }
            return @{ ok = $true; app = "excel"; path = $finalPath; title = $title; liveMode = $wasRunning; reused = $reuseWb }
        } catch {
            if (-not $wasRunning) {
                if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
                if ($null -ne $excel) { try { $excel.Quit() } catch {} }
            }
            throw $_
        } finally {
            if ($null -ne $wb) { Release-Com $wb }
            if ($null -ne $excel) { Release-Com $excel }
            Invoke-GarbageCollection
        }
    }
    
    return @{ ok = $false; error = "Unsupported app: $appType" }
}

function Invoke-Open($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required to open a document" } }
    $resolvedPath = Normalize-DocPath $params.path
    if (-not (Test-Path $resolvedPath)) { return @{ ok = $false; error = "File not found: $resolvedPath" } }
    
    $ext = [System.IO.Path]::GetExtension($resolvedPath).ToLower()
    $visible = if ($null -ne $params.visible) { [bool]$params.visible } else { $true }
    $readOnly = if ($null -ne $params.readOnly) { [bool]$params.readOnly } else { $false }
    
    if ($ext -in @(".doc", ".docx", ".rtf", ".txt")) {
        $word = Get-RunningApp "Word.Application"
        $wasRunning = ($null -ne $word)
        if (-not $wasRunning) { $word = New-Object -ComObject Word.Application }
        $doc = $null
        try {
            $word.Visible = $visible
            $doc = $word.Documents.Open($resolvedPath, [ref]$false, [ref]$readOnly)
            return @{ ok = $true; app = "word"; path = $doc.FullName; title = $doc.Name; readOnly = $doc.ReadOnly }
        } catch {
            if (-not $wasRunning) {
                if ($null -ne $doc) { try { $doc.Close([ref]0) } catch {} }
                try { $word.Quit() } catch {}
            }
            throw $_
        } finally {
            if ($null -ne $doc) { Release-Com $doc }
            if ($null -ne $word) { Release-Com $word }
            Invoke-GarbageCollection
        }
    }
    
    if ($ext -in @(".ppt", ".pptx", ".potx")) {
        $ppt = Get-RunningApp "PowerPoint.Application"
        $wasRunning = ($null -ne $ppt)
        if (-not $wasRunning) { $ppt = New-Object -ComObject PowerPoint.Application }
        $pres = $null
        try {
            $ppt.Visible = 1
            $roState = if ($readOnly) { [Microsoft.Office.Core.MsoTriState]::msoTrue } else { [Microsoft.Office.Core.MsoTriState]::msoFalse }
            $pres = $ppt.Presentations.Open($resolvedPath, $roState, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)
            return @{ ok = $true; app = "powerpoint"; path = $pres.FullName; title = $pres.Name; readOnly = $pres.ReadOnly }
        } catch {
            if (-not $wasRunning) {
                if ($null -ne $pres) { try { $pres.Close() } catch {} }
                try { $ppt.Quit() } catch {}
            }
            throw $_
        } finally {
            if ($null -ne $pres) { Release-Com $pres }
            if ($null -ne $ppt) { Release-Com $ppt }
            Invoke-GarbageCollection
        }
    }
    
    if ($ext -in @(".xls", ".xlsx", ".csv")) {
        $excel = Get-RunningApp "Excel.Application"
        $wasRunning = ($null -ne $excel)
        if (-not $wasRunning) { $excel = New-Object -ComObject Excel.Application }
        $wb = $null
        try {
            $excel.Visible = $visible
            $wb = $excel.Workbooks.Open($resolvedPath, [Type]::Missing, $readOnly)
            return @{ ok = $true; app = "excel"; path = $wb.FullName; title = $wb.Name; readOnly = $wb.ReadOnly }
        } catch {
            if (-not $wasRunning) {
                if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
                try { $excel.Quit() } catch {}
            }
            throw $_
        } finally {
            if ($null -ne $wb) { Release-Com $wb }
            if ($null -ne $excel) { Release-Com $excel }
            Invoke-GarbageCollection
        }
    }
    
    return @{ ok = $false; error = "Unrecognized document extension: $ext" }
}

function Invoke-Save($params) {
    $targetPath = if ($params.path) { Normalize-DocPath $params.path } else { "" }
    $appType = if ($params.app) { $params.app.ToLower() } else { "" }
    
    $targetName = if ($targetPath) { [System.IO.Path]::GetFileName($targetPath) } else { "" }
    
    # Try Word
    if (-not $appType -or $appType -eq "word") {
        $word = Get-RunningApp "Word.Application"
        if ($null -ne $word) {
            try {
                $doc = Find-DocInCollection $word $targetPath "word"
                if ($null -ne $doc) {
                    if ([string]::IsNullOrWhiteSpace($doc.Path)) {
                        return @{ ok = $false; error = "Document is newly created and has no disk path. Use 'save_as' with targetPath instead of 'save'." }
                    }
                    $doc.Save()
                    return @{ ok = $true; app = "word"; path = $doc.FullName; title = $doc.Name; saved = $true }
                }
            } finally {
                Release-Com $word
                Invoke-GarbageCollection
            }
        }
    }
    
    # Try PowerPoint
    if (-not $appType -or $appType -eq "powerpoint") {
        $ppt = Get-RunningApp "PowerPoint.Application"
        if ($null -ne $ppt) {
            try {
                $pres = Find-DocInCollection $ppt $targetPath "powerpoint"
                if ($null -ne $pres) {
                    if ([string]::IsNullOrWhiteSpace($pres.Path)) {
                        return @{ ok = $false; error = "Presentation is newly created and has no disk path. Use 'save_as' with targetPath instead of 'save'." }
                    }
                    $pres.Save()
                    return @{ ok = $true; app = "powerpoint"; path = $pres.FullName; title = $pres.Name; saved = $true }
                }
            } finally {
                Release-Com $ppt
                Invoke-GarbageCollection
            }
        }
    }
    
    # Try Excel
    if (-not $appType -or $appType -eq "excel") {
        $excel = Get-RunningApp "Excel.Application"
        if ($null -ne $excel) {
            try {
                $wb = Find-DocInCollection $excel $targetPath "excel"
                if ($null -ne $wb) {
                    if ([string]::IsNullOrWhiteSpace($wb.Path)) {
                        return @{ ok = $false; error = "Workbook is newly created and has no disk path. Use 'save_as' with targetPath instead of 'save'." }
                    }
                    $wb.Save()
                    return @{ ok = $true; app = "excel"; path = $wb.FullName; title = $wb.Name; saved = $true }
                }
            } finally {
                Release-Com $excel
                Invoke-GarbageCollection
            }
        }
    }
    
    return @{ ok = $false; error = "Document not found in open instances to save" }
}

function Invoke-SaveAs($params) {
    if (-not $params.targetPath) { return @{ ok = $false; error = "targetPath is required for save_as" } }
    $targetPath = Normalize-DocPath $params.targetPath
    $sourcePath = if ($params.path) { Normalize-DocPath $params.path } else { "" }
    
    $parentDir = [System.IO.Path]::GetDirectoryName($targetPath)
    if (-not (Test-Path $parentDir)) { New-Item -ItemType Directory -Path $parentDir -Force | Out-Null }
    
    # Word
    $word = Get-RunningApp "Word.Application"
    if ($null -ne $word) {
        try {
            foreach ($doc in $word.Documents) {
                if (-not $sourcePath -or (Normalize-DocPath $doc.FullName) -eq $sourcePath -or $doc.Name -eq $sourcePath) {
                    $doc.SaveAs2($targetPath)
                    $res = @{ ok = $true; app = "word"; sourcePath = $doc.FullName; targetPath = $targetPath }
                    Release-Com $doc
                    return $res
                }
                Release-Com $doc
            }
        } finally {
            Release-Com $word
            Invoke-GarbageCollection
        }
    }
    
    # PowerPoint
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -ne $ppt) {
        try {
            foreach ($pres in $ppt.Presentations) {
                if (-not $sourcePath -or (Normalize-DocPath $pres.FullName) -eq $sourcePath -or $pres.Name -eq $sourcePath) {
                    $pres.SaveAs($targetPath)
                    $res = @{ ok = $true; app = "powerpoint"; sourcePath = $pres.FullName; targetPath = $targetPath }
                    Release-Com $pres
                    return $res
                }
                Release-Com $pres
            }
        } finally {
            Release-Com $ppt
            Invoke-GarbageCollection
        }
    }
    
    # Excel
    $excel = Get-RunningApp "Excel.Application"
    if ($null -ne $excel) {
        try {
            foreach ($wb in $excel.Workbooks) {
                if (-not $sourcePath -or (Normalize-DocPath $wb.FullName) -eq $sourcePath -or $wb.Name -eq $sourcePath) {
                    $wb.SaveAs($targetPath)
                    $res = @{ ok = $true; app = "excel"; sourcePath = $wb.FullName; targetPath = $targetPath }
                    Release-Com $wb
                    return $res
                }
                Release-Com $wb
            }
        } finally {
            Release-Com $excel
            Invoke-GarbageCollection
        }
    }
    
    return @{ ok = $false; error = "Target document not found among open instances" }
}

function Invoke-Close($params) {
    $targetPath = if ($params.path) { Normalize-DocPath $params.path } else { "" }
    $saveChanges = if ($null -ne $params.saveChanges) { [bool]$params.saveChanges } else { $true }
    $closeApp = if ($null -ne $params.closeApp) { [bool]$params.closeApp } else { $false }
    
    # Word
    $word = Get-RunningApp "Word.Application"
    if ($null -ne $word) {
        $closedDoc = $null
        try {
            foreach ($doc in $word.Documents) {
                if (-not $targetPath -or (Normalize-DocPath $doc.FullName) -eq $targetPath -or $doc.Name -eq $targetPath) {
                    $saveOption = if ($saveChanges) { -1 } else { 0 }
                    $closedDoc = $doc.Name
                    try { $doc.Close([ref]$saveOption) } catch {}
                    Release-Com $doc
                    break
                }
                Release-Com $doc
            }
            $remDocs = 0
            try { $remDocs = $word.Documents.Count } catch {}
            $shouldQuit = ($remDocs -eq 0 -or $closeApp)
            if ($shouldQuit) {
                try { $word.Quit() } catch {}
            }
        } finally {
            Release-Com $word
            Invoke-GarbageCollection
        }
        if ($null -ne $closedDoc) {
            return @{ ok = $true; app = "word"; closed = $closedDoc; saved = $saveChanges; appClosed = $shouldQuit }
        }
    }
    
    # PowerPoint
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -ne $ppt) {
        $closedPres = $null
        try {
            foreach ($pres in $ppt.Presentations) {
                if (-not $targetPath -or (Normalize-DocPath $pres.FullName) -eq $targetPath -or $pres.Name -eq $targetPath) {
                    if ($saveChanges) { try { $pres.Save() } catch {} }
                    $closedPres = $pres.Name
                    try { $pres.Close() } catch {}
                    Release-Com $pres
                    break
                }
                Release-Com $pres
            }
            $remPres = 0
            try { $remPres = $ppt.Presentations.Count } catch {}
            $shouldQuit = ($remPres -eq 0 -or $closeApp)
            if ($shouldQuit) {
                try { $ppt.Quit() } catch {}
            }
        } finally {
            Release-Com $ppt
            Invoke-GarbageCollection
        }
        if ($null -ne $closedPres) {
            return @{ ok = $true; app = "powerpoint"; closed = $closedPres; saved = $saveChanges; appClosed = $shouldQuit }
        }
    }
    
    # Excel
    $excel = Get-RunningApp "Excel.Application"
    if ($null -ne $excel) {
        $closedWb = $null
        try {
            foreach ($wb in $excel.Workbooks) {
                if (-not $targetPath -or (Normalize-DocPath $wb.FullName) -eq $targetPath -or $wb.Name -eq $targetPath) {
                    $closedWb = $wb.Name
                    try { $wb.Close([bool]$saveChanges) } catch {}
                    Release-Com $wb
                    break
                }
                Release-Com $wb
            }
            $remWbs = 0
            try { $remWbs = $excel.Workbooks.Count } catch {}
            $shouldQuit = ($remWbs -eq 0 -or $closeApp)
            if ($shouldQuit) {
                try { $excel.Quit() } catch {}
            }
        } finally {
            Release-Com $excel
            Invoke-GarbageCollection
        }
        if ($null -ne $closedWb) {
            return @{ ok = $true; app = "excel"; closed = $closedWb; saved = $saveChanges; appClosed = $shouldQuit }
        }
    }
    
    return @{ ok = $false; error = "Document not found to close" }
}

function Invoke-ExportPdf($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Document path is required for export_pdf" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    $outputPath = if ($params.outputPath) {
        Normalize-DocPath $params.outputPath
    } else {
        [System.IO.Path]::ChangeExtension($srcPath, ".pdf")
    }
    
    $parentDir = [System.IO.Path]::GetDirectoryName($outputPath)
    if (-not (Test-Path $parentDir)) { New-Item -ItemType Directory -Path $parentDir -Force | Out-Null }
    
    $ext = [System.IO.Path]::GetExtension($srcPath).ToLower()
    
    if ($ext -in @(".doc", ".docx", ".rtf", ".txt")) {
        $word = Get-RunningApp "Word.Application"
        $wasRunning = ($null -ne $word)
        if (-not $wasRunning) { $word = New-Object -ComObject Word.Application; $word.Visible = $false }
        
        $doc = $null
        foreach ($d in $word.Documents) {
            if ((Normalize-DocPath $d.FullName) -eq $srcPath) { $doc = $d; break }
        }
        $openedHere = $false
        if ($null -eq $doc) {
            $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$true)
            $openedHere = $true
        }
        
        try {
            # 17 = wdExportFormatPDF
            $doc.ExportAsFixedFormat($outputPath, 17)
            $fi = Get-Item $outputPath
            return @{ ok = $true; format = "pdf"; pdfPath = $outputPath; sizeBytes = $fi.Length }
        } finally {
            if ($openedHere) {
                try { $doc.Close([ref]0) } catch {}
                Release-Com $doc
            } elseif (-not $wasRunning) {
                Release-Com $doc
            }
            if (-not $wasRunning) {
                try { $word.Quit() } catch {}
                Release-Com $word
            }
            Invoke-GarbageCollection
        }
    }
    
    if ($ext -in @(".ppt", ".pptx")) {
        $ppt = Get-RunningApp "PowerPoint.Application"
        $wasRunning = ($null -ne $ppt)
        if (-not $wasRunning) { $ppt = New-Object -ComObject PowerPoint.Application }
        
        $pres = $null
        foreach ($p in $ppt.Presentations) {
            if ((Normalize-DocPath $p.FullName) -eq $srcPath) { $pres = $p; break }
        }
        $openedHere = $false
        if ($null -eq $pres) {
            $pres = $ppt.Presentations.Open($srcPath, 1, 0, 0)
            $openedHere = $true
        }
        
        try {
            # 2 = ppFixedFormatTypePDF
            $pres.ExportAsFixedFormat($outputPath, 2)
            $fi = Get-Item $outputPath
            return @{ ok = $true; format = "pdf"; pdfPath = $outputPath; sizeBytes = $fi.Length }
        } finally {
            if ($openedHere) {
                try { $pres.Close() } catch {}
                Release-Com $pres
            } elseif (-not $wasRunning) {
                Release-Com $pres
            }
            if (-not $wasRunning) {
                try { $ppt.Quit() } catch {}
                Release-Com $ppt
            }
            Invoke-GarbageCollection
        }
    }
    
    if ($ext -in @(".xls", ".xlsx")) {
        $excel = Get-RunningApp "Excel.Application"
        $wasRunning = ($null -ne $excel)
        if (-not $wasRunning) { $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false }
        
        $wb = $null
        foreach ($w in $excel.Workbooks) {
            if ((Normalize-DocPath $w.FullName) -eq $srcPath) { $wb = $w; break }
        }
        $openedHere = $false
        if ($null -eq $wb) {
            $wb = $excel.Workbooks.Open($srcPath, [Type]::Missing, $true)
            $openedHere = $true
        }
        
        try {
            # 0 = xlTypePDF
            $wb.ExportAsFixedFormat(0, $outputPath)
            $fi = Get-Item $outputPath
            return @{ ok = $true; format = "pdf"; pdfPath = $outputPath; sizeBytes = $fi.Length }
        } finally {
            if ($openedHere) {
                try { $wb.Close($false) } catch {}
                Release-Com $wb
            } elseif (-not $wasRunning) {
                Release-Com $wb
            }
            if (-not $wasRunning) {
                try { $excel.Quit() } catch {}
                Release-Com $excel
            }
            Invoke-GarbageCollection
        }
    }
    
    return @{ ok = $false; error = "Unsupported format for PDF export: $ext" }
}

# --- Word Content Actions ---

function Invoke-WordRead($params) {
    $srcPath = Normalize-DocPath $params.path
    $maxChars = if ($params.maxChars) { [int]$params.maxChars } else { 4000 }
    
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) {
        Start-Sleep -Milliseconds 250
        $word = Get-RunningApp "Word.Application" $srcPath
    }
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $wasRunning = $true
    } else {
        try { $word.Visible = $true } catch {}
    }
    
    $doc = Find-DocInCollection $word $srcPath "word"
    $openedHere = $false
    if ($null -eq $doc) {
        if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$true)
            $openedHere = $true
        } elseif ($word.Documents.Count -gt 0) {
            $doc = $word.ActiveDocument
        } else {
            return @{ ok = $false; error = "Document not open and file not found on disk: $srcPath" }
        }
    }
    
    try {
        $totalParagraphs = $doc.Paragraphs.Count
        $headings = @()
        $sampleText = ""
        $charCount = 0
        
        foreach ($para in $doc.Paragraphs) {
            $t = $para.Range.Text.Trim()
            if ($t) {
                $styleName = ""
                try { $styleName = $para.Style.NameLocal } catch {}
                $outline = 10
                try { $outline = [int]$para.OutlineLevel } catch {}
                if ($outline -lt 10 -or $styleName -match "Heading|T[ií]tulo|Title") {
                    $headings += @{ style = $styleName; text = $t; level = $outline }
                }
                if ($charCount -lt $maxChars) {
                    $sampleText += $t + "`n"
                    $charCount += $t.Length
                }
            }
        }
        
        $tablesCount = $doc.Tables.Count
        $commentsCount = $doc.Comments.Count
        
        $finalPath = if ($doc -and $doc.FullName) { $doc.FullName } else { $srcPath }
        return @{
            ok = $true
            path = $finalPath
            title = $doc.Name
            paragraphs = $totalParagraphs
            headings = $headings
            tables = $tablesCount
            comments = $commentsCount
            textSnippet = $sampleText.Trim()
            truncated = ($charCount -ge $maxChars)
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function ConvertTo-WordColor($color) {
    if ($null -eq $color) { return $null }
    if ($color -is [int]) { return $color }
    $c = $color.ToString().Trim().ToLower()
    if ($c.StartsWith("#")) {
        $hex = $c.TrimStart("#")
        if ($hex.Length -eq 6) {
            $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
            $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
            $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
            # Word wdColor uses BGR integer: R + (G * 256) + (B * 65536)
            return [int]($r + ($g * 256) + ($b * 65536))
        }
    }
    switch ($c) {
        "black" { return 0 }
        "blue" { return 16711680 }
        "red" { return 255 }
        "green" { return 32768 }
        "darkgreen" { return 197120 }
        "gray" { return 8421504 }
        "navy" { return 8388608 }
        "teal" { return 8421376 }
        "purple" { return 8388736 }
        "maroon" { return 128 }
        "guinda" { return 102 } # 0x000066 (RGB: 102, 0, 0)
        "ipn" { return 102 }
        "gold" { return 55295 } # RGB(255, 215, 0) -> BGR: 0 + 215*256 + 255*65536 = 16768768
        "orange" { return 33023 }
        "yellow" { return 65535 }
        "white" { return 16777215 }
        default {
            try { return [int]$color } catch { return $null }
        }
    }
}

function ConvertTo-WordAlignment($align) {
    if ($null -eq $align) { return $null }
    $a = $align.ToString().Trim().ToLower()
    switch ($a) {
        "center" { return 1 } # wdAlignParagraphCenter
        "centrado" { return 1 }
        "right" { return 2 }  # wdAlignParagraphRight
        "derecha" { return 2 }
        "justify" { return 3 } # wdAlignParagraphJustify
        "justificado" { return 3 }
        default { return 0 } # wdAlignParagraphLeft
    }
}

function Invoke-WordInsertText($params) {
    if (-not $params.text -and -not $params.content -and -not $params.pageBreak) { return @{ ok = $false; error = "Text to insert is required" } }
    $textToInsert = if ($null -ne $params.text) { [string]$params.text } else { [string]$params.content }
    $srcPath = Normalize-DocPath $params.path
    
    $word = Get-RunningApp "Word.Application" $srcPath
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) {
        if (-not $srcPath -or -not (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $true
            $wasRunning = $true
        } else {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
        }
    } else {
        $word.Visible = $true
    }
    
    $doc = Find-DocInCollection $word $srcPath "word"
    $openedHere = $false
    if ($null -eq $doc) {
        if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
            $openedHere = $true
        } elseif ($word.Documents.Count -gt 0) {
            $doc = $word.ActiveDocument
        } else {
            $doc = $word.Documents.Add()
            $openedHere = $false
        }
    }
    
    try {
        $pos = if ($params.position) { $params.position.ToLower() } else { "end" }
        $range = $null
        if ($pos -eq "start") {
            $range = $doc.Range(0, 0)
        } else {
            $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        }
        
        $startPos = $range.Start
        if ($textToInsert) {
            $range.Text = $textToInsert
            $range.InsertParagraphAfter()
        }
        if ($params.pageBreak) {
            $range.InsertBreak(7) # wdPageBreak
        }
        $endPos = $range.End
        
        # Apply formatting across the entire newly inserted range
        $appliedRange = $doc.Range($startPos, $endPos)
        
        if ($null -ne $params.bold) { $appliedRange.Bold = [int][bool]$params.bold }
        if ($null -ne $params.italic) { $appliedRange.Italic = [int][bool]$params.italic }
        if ($null -ne $params.underline) { $appliedRange.Underline = [int][bool]$params.underline }
        
        $fSize = if ($params.fontSize) { [float]$params.fontSize } elseif ($params.size) { [float]$params.size } else { $null }
        if ($null -ne $fSize) { $appliedRange.Font.Size = $fSize }
        
        $fName = if ($params.fontName) { [string]$params.fontName } elseif ($params.font) { [string]$params.font } else { $null }
        if ($null -ne $fName) { $appliedRange.Font.Name = $fName }
        
        $colorVal = ConvertTo-WordColor $params.color
        if ($null -ne $colorVal) { $appliedRange.Font.Color = $colorVal }
        
        $alignParam = if ($params.alignment) { $params.alignment } else { $params.align }
        $alignVal = ConvertTo-WordAlignment $alignParam
        if ($null -ne $alignVal) { $appliedRange.ParagraphFormat.Alignment = $alignVal }
        
        if ($null -ne $params.lineSpacing) {
            $appliedRange.ParagraphFormat.LineSpacingRule = 1
            $appliedRange.ParagraphFormat.LineSpacing = [float]($params.lineSpacing * 12)
        }
        if ($null -ne $params.spaceAfter) { $appliedRange.ParagraphFormat.SpaceAfter = [float]$params.spaceAfter }
        if ($null -ne $params.spaceBefore) { $appliedRange.ParagraphFormat.SpaceBefore = [float]$params.spaceBefore }
        
        # Apply Named Word Style
        if ($params.style) {
            $styleStr = $params.style.ToString().Trim()
            try { $appliedRange.ParagraphFormat.Style = $styleStr } catch {
                try {
                    $styleMap = @{
                        "title" = -63; "titulo" = -63;
                        "heading 1" = -2; "titulo 1" = -2;
                        "heading 2" = -3; "titulo 2" = -3;
                        "heading 3" = -4; "titulo 3" = -4;
                        "subtitle" = -75; "subtitulo" = -75;
                        "normal" = -1; "quote" = -29; "cita" = -29;
                    }
                    $low = $styleStr.ToLower()
                    if ($styleMap.ContainsKey($low)) {
                        $appliedRange.ParagraphFormat.Style = $styleMap[$low]
                    }
                } catch {}
            }
        }
        
        if ($params.bullet) {
            try { $appliedRange.ListFormat.ApplyBulletDefault() } catch {}
        }
        
        # NEVER force Save on unsaved Document1 to avoid "Save As" modal prompt
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        
        return @{
            ok = $true
            path = $doc.FullName
            title = $doc.Name
            insertedLength = $textToInsert.Length
            position = $pos
            totalParagraphs = $doc.Paragraphs.Count
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordReplaceText($params) {
    $search = if ($params.search) { $params.search } elseif ($params.find) { $params.find } elseif ($params.target) { $params.target } else { "" }
    if (-not $search) { return @{ ok = $false; error = "Search text is required" } }
    $replace = if ($null -ne $params.replace) { [string]$params.replace } elseif ($null -ne $params.replacement) { [string]$params.replacement } elseif ($null -ne $params.content) { [string]$params.content } else { "" }
    $srcPath = Normalize-DocPath $params.path
    
    $word = Get-RunningApp "Word.Application" $srcPath
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) {
        if (-not $srcPath -or -not (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $true
            $wasRunning = $true
        } else {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
        }
    } else {
        $word.Visible = $true
    }
    
    $doc = Find-DocInCollection $word $srcPath "word"
    $openedHere = $false
    if ($null -eq $doc) {
        if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
            $openedHere = $true
        } elseif ($word.Documents.Count -gt 0) {
            $doc = $word.ActiveDocument
        } else {
            return @{ ok = $false; error = "Document not open and file not found on disk: $srcPath" }
        }
    }
    
    try {
        $matchCase = if ($null -ne $params.matchCase) { [bool]$params.matchCase } else { $false }
        $matchWholeWord = if ($null -ne $params.matchWholeWord) { [bool]$params.matchWholeWord } else { $false }
        
        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $search
        $find.Replacement.Text = $replace
        $find.Forward = $true
        $find.Wrap = 1 # wdFindContinue
        
        # wdReplaceAll = 2
        $replaced = $find.Execute($search, [bool]$matchCase, [bool]$matchWholeWord, $false, $false, $false, $true, 1, $false, $replace, 2)
        
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        
        return @{
            ok = $true
            path = $doc.FullName
            title = $doc.Name
            search = $search
            replace = $replace
            success = $replaced
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordClear($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) {
        Start-Sleep -Milliseconds 250
        $word = Get-RunningApp "Word.Application" $srcPath
    }
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $wasRunning = $true
    } else {
        try { $word.Visible = $true } catch {}
    }
    
    $doc = Find-DocInCollection $word $srcPath "word"
    $openedHere = $false
    if ($null -eq $doc) {
        if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
            $openedHere = $true
        } elseif ($word.Documents.Count -gt 0) {
            $doc = $word.ActiveDocument
        } else {
            return @{ ok = $false; error = "Document not open and file not found on disk: $srcPath" }
        }
    }
    
    try {
        $doc.Content.Delete()
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        return @{
            ok = $true
            path = $doc.FullName
            title = $doc.Name
            cleared = $true
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordFormatText($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) {
        Start-Sleep -Milliseconds 250
        $word = Get-RunningApp "Word.Application" $srcPath
    }
    if ($null -eq $word) { return @{ ok = $false; error = "Word is not running" } }
    
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "No open document found" } }
    
    try {
        $search = if ($params.search) { [string]$params.search } elseif ($params.text) { [string]$params.text } elseif ($params.target) { [string]$params.target } else { "" }
        $all = [bool]$params.all
        $allOccurrences = if ($null -ne $params.allOccurrences) { [bool]$params.allOccurrences } else { $true }
        $matchCount = 0
        
        $applyStylesToRange = {
            param($targetRange)
            if ($null -ne $params.fontSize) { $targetRange.Font.Size = [float]$params.fontSize }
            if ($null -ne $params.fontName) { $targetRange.Font.Name = [string]$params.fontName }
            if ($null -ne $params.bold) { $targetRange.Font.Bold = [int][bool]$params.bold }
            if ($null -ne $params.italic) { $targetRange.Font.Italic = [int][bool]$params.italic }
            if ($null -ne $params.underline) { $targetRange.Font.Underline = if ($params.underline) { 1 } else { 0 } }
            if ($null -ne $params.color) {
                $c = ConvertTo-WordColor $params.color
                if ($null -ne $c) { $targetRange.Font.Color = $c }
            }
            if ($null -ne $params.alignment) {
                $alignVal = ConvertTo-WordAlignment $params.alignment
                if ($null -ne $alignVal) { $targetRange.ParagraphFormat.Alignment = $alignVal }
            }
            if ($null -ne $params.lineSpacing) {
                $targetRange.ParagraphFormat.LineSpacingRule = 1 # wdLineSpaceMultiple
                $targetRange.ParagraphFormat.LineSpacing = [float]($params.lineSpacing * 12)
            }
            if ($null -ne $params.spaceBefore) {
                $targetRange.ParagraphFormat.SpaceBefore = [float]$params.spaceBefore
            }
            if ($null -ne $params.spaceAfter) {
                $targetRange.ParagraphFormat.SpaceAfter = [float]$params.spaceAfter
            }
            if ($null -ne $params.style) {
                $sName = [string]$params.style
                try {
                    $targetRange.ParagraphFormat.Style = $sName
                } catch {
                    if ($sName -match "T[ií]tulo\s*(\d)") {
                        try { $targetRange.ParagraphFormat.Style = "Heading $($matches[1])" } catch {}
                    } elseif ($sName -match "Heading\s*(\d)") {
                        try { $targetRange.ParagraphFormat.Style = "Título $($matches[1])" } catch {}
                    }
                }
            }
        }
        
        if (-not $search -and $all) {
            & $applyStylesToRange $doc.Content
            $matchCount = 1
        } elseif ($search) {
            $find = $doc.Content.Find
            $find.ClearFormatting()
            $find.Text = $search
            $find.Forward = $true
            $find.Wrap = 0 # wdFindStop
            
            while ($find.Execute()) {
                & $applyStylesToRange $find.Parent
                $matchCount++
                if (-not $allOccurrences) { break }
            }
        } else {
            return @{ ok = $false; error = "Either 'search'/'text' or 'all: true' is required" }
        }
        
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        
        return @{
            ok = $true
            path = $doc.FullName
            title = $doc.Name
            matchesFormatted = $matchCount
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordFormatParagraphs($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) { return @{ ok = $false; error = "Word is not running" } }
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "No open document found" } }
    
    try {
        $totalParas = $doc.Paragraphs.Count
        $all = [bool]$params.all
        $startIdx = if ($all) { 1 } elseif ($params.start) { [int]$params.start } else { 1 }
        $endIdx = if ($all) { $totalParas } elseif ($params.end) { [int]$params.end } else { $startIdx }
        
        if ($startIdx -lt 1) { $startIdx = 1 }
        if ($endIdx -gt $totalParas) { $endIdx = $totalParas }
        
        $formattedCount = 0
        for ($i = $startIdx; $i -le $endIdx; $i++) {
            $p = $doc.Paragraphs.Item($i)
            if ($null -ne $params.fontSize) { $p.Range.Font.Size = [float]$params.fontSize }
            if ($null -ne $params.fontName) { $p.Range.Font.Name = [string]$params.fontName }
            if ($null -ne $params.bold) { $p.Range.Font.Bold = [int][bool]$params.bold }
            if ($null -ne $params.italic) { $p.Range.Font.Italic = [int][bool]$params.italic }
            if ($null -ne $params.underline) { $p.Range.Font.Underline = if ($params.underline) { 1 } else { 0 } }
            if ($null -ne $params.color) {
                $c = ConvertTo-WordColor $params.color
                if ($null -ne $c) { $p.Range.Font.Color = $c }
            }
            if ($null -ne $params.alignment) {
                $alignVal = ConvertTo-WordAlignment $params.alignment
                if ($null -ne $alignVal) { $p.Alignment = $alignVal }
            }
            if ($null -ne $params.lineSpacing) {
                $p.LineSpacingRule = 1 # wdLineSpaceMultiple
                $p.LineSpacing = [float]($params.lineSpacing * 12)
            }
            if ($null -ne $params.spaceBefore) {
                $p.SpaceBefore = [float]$params.spaceBefore
            }
            if ($null -ne $params.spaceAfter) {
                $p.SpaceAfter = [float]$params.spaceAfter
            }
            if ($null -ne $params.style) {
                $sName = [string]$params.style
                try {
                    $p.Format.Style = $sName
                } catch {
                    if ($sName -match "T[ií]tulo\s*(\d)") {
                        try { $p.Format.Style = "Heading $($matches[1])" } catch {}
                    } elseif ($sName -match "Heading\s*(\d)") {
                        try { $p.Format.Style = "Título $($matches[1])" } catch {}
                    }
                }
            }
            $formattedCount++
        }
        
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        
        return @{
            ok = $true
            path = $doc.FullName
            title = $doc.Name
            start = $startIdx
            end = $endIdx
            totalFormatted = $formattedCount
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordPageSetup($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) { return @{ ok = $false; error = "Word is not running" } }
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "No open document found" } }
    
    try {
        $ps = $doc.PageSetup
        
        $pSize = if ($params.paperSize) { $params.paperSize.ToString().ToLower() } else { "letter" }
        if ($pSize -match "letter|carta") { $ps.PaperSize = 2 }
        elseif ($pSize -match "a4") { $ps.PaperSize = 7 }
        elseif ($pSize -match "legal|oficio") { $ps.PaperSize = 3 }
        
        if ($params.orientation) {
            $ori = $params.orientation.ToString().ToLower()
            if ($ori -match "landscape|horizontal") { $ps.Orientation = 1 }
            else { $ps.Orientation = 0 }
        }
        
        $mType = if ($params.margins) { $params.margins.ToString().ToLower() } else { "normal" }
        if ($mType -match "narrow|estrecho") {
            $ps.TopMargin = 36.0; $ps.BottomMargin = 36.0; $ps.LeftMargin = 36.0; $ps.RightMargin = 36.0
        } elseif ($mType -match "moderate|moderado") {
            $ps.TopMargin = 72.0; $ps.BottomMargin = 72.0; $ps.LeftMargin = 54.0; $ps.RightMargin = 54.0
        } elseif ($mType -match "normal") {
            $ps.TopMargin = 72.0; $ps.BottomMargin = 72.0; $ps.LeftMargin = 72.0; $ps.RightMargin = 72.0
        }
        
        if ($null -ne $params.topMargin) { $ps.TopMargin = [float]$params.topMargin }
        if ($null -ne $params.bottomMargin) { $ps.BottomMargin = [float]$params.bottomMargin }
        if ($null -ne $params.leftMargin) { $ps.LeftMargin = [float]$params.leftMargin }
        if ($null -ne $params.rightMargin) { $ps.RightMargin = [float]$params.rightMargin }
        
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        
        return @{
            ok = $true
            path = $doc.FullName
            paperSize = $pSize
            margins = $mType
            pageWidth = [float]$ps.PageWidth
            pageHeight = [float]$ps.PageHeight
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordAddPageNumbers($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) { return @{ ok = $false; error = "Word is not running" } }
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "No open document found" } }
    
    try {
        $firstPage = if ($null -ne $params.firstPage) { [bool]$params.firstPage } else { $false }
        $sec = $doc.Sections.Item(1)
        if (-not $firstPage) {
            $sec.PageSetup.DifferentFirstPageHeaderFooter = -1
        }
        
        $alignInt = switch ($params.alignment) {
            "left" { 0 }
            "right" { 2 }
            default { 1 }
        }
        
        $footer = $sec.Footers.Item(1)
        $footer.PageNumbers.Add($alignInt, $firstPage) | Out-Null
        
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        
        return @{
            ok = $true
            path = $doc.FullName
            alignment = if ($alignInt -eq 2) { "right" } elseif ($alignInt -eq 0) { "left" } else { "center" }
            differentFirstPage = (-not $firstPage)
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordInsertPageBreak($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) { return @{ ok = $false; error = "Word is not running" } }
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "No open document found" } }
    
    try {
        $before = if ($params.beforeText) { [string]$params.beforeText } else { "" }
        $after = if ($params.afterText) { [string]$params.afterText } else { "" }
        $inserted = $false
        
        if ($before) {
            $find = $doc.Content.Find
            $find.ClearFormatting()
            $find.Text = $before
            $find.Forward = $true
            $find.Wrap = 0
            if ($find.Execute()) {
                $rng = $find.Parent
                $breakRng = $doc.Range($rng.Start, $rng.Start)
                $breakRng.InsertBreak(7)
                $inserted = $true
            }
        } elseif ($after) {
            $find = $doc.Content.Find
            $find.ClearFormatting()
            $find.Text = $after
            $find.Forward = $true
            $find.Wrap = 0
            if ($find.Execute()) {
                $rng = $find.Parent
                $breakRng = $doc.Range($rng.End, $rng.End)
                $breakRng.InsertBreak(7)
                $inserted = $true
            }
        } else {
            $endPos = $doc.Content.End - 1
            $breakRng = $doc.Range($endPos, $endPos)
            $breakRng.InsertBreak(7)
            $inserted = $true
        }
        
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        
        return @{
            ok = $inserted
            path = $doc.FullName
            inserted = $inserted
            totalParagraphs = $doc.Paragraphs.Count
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordAppendSection($params) {
    if (-not $params.heading) { return @{ ok = $false; error = "Heading is required" } }
    $srcPath = Normalize-DocPath $params.path
    
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) {
        Start-Sleep -Milliseconds 250
        $word = Get-RunningApp "Word.Application" $srcPath
    }
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $wasRunning = $true
    } else {
        try { $word.Visible = $true } catch {}
    }
    
    $doc = Find-DocInCollection $word $srcPath "word"
    $openedHere = $false
    if ($null -eq $doc) {
        if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
            $openedHere = $true
        } elseif ($word.Documents.Count -gt 0) {
            $doc = $word.ActiveDocument
        } else {
            return @{ ok = $false; error = "Document not open and file not found on disk: $srcPath" }
        }
    }
    
    try {
        $level = if ($params.headingLevel) { [int]$params.headingLevel } else { 1 }
        $headingStyleId = switch ($level) {
            2 { -3 } # wdStyleHeading2
            3 { -4 } # wdStyleHeading3
            default { -2 } # wdStyleHeading1
        }
        
        # Insert Heading paragraph
        $pHead = $doc.Paragraphs.Add()
        $pHead.Range.Text = $params.heading
        try {
            $pHead.Format.Style = $headingStyleId
        } catch {
            try { $pHead.Format.Style = "Heading $level" } catch {
                try { $pHead.Format.Style = "Título $level" } catch {}
            }
        }
        
        if ($null -ne $params.headingColor) {
            $hColor = ConvertTo-WordColor $params.headingColor
            if ($null -ne $hColor) { $pHead.Range.Font.Color = $hColor }
        }
        if ($null -ne $params.headingAlignment) {
            $hAlign = ConvertTo-WordAlignment $params.headingAlignment
            if ($null -ne $hAlign) { $pHead.Alignment = $hAlign }
        }
        
        $pHead.Range.InsertParagraphAfter()
        
        # Insert Body paragraph if present
        if ($params.body) {
            $pBody = $doc.Paragraphs.Add()
            $pBody.Range.Text = $params.body
            try { $pBody.Format.Style = -1 } catch {
                try { $pBody.Format.Style = "Normal" } catch {}
            }
            if ($null -ne $params.bodyColor) {
                $bColor = ConvertTo-WordColor $params.bodyColor
                if ($null -ne $bColor) { $pBody.Range.Font.Color = $bColor }
            }
            if ($null -ne $params.bodyAlignment) {
                $bAlign = ConvertTo-WordAlignment $params.bodyAlignment
                if ($null -ne $bAlign) { $pBody.Alignment = $bAlign }
            }
            if ($null -ne $params.fontSize) {
                $pBody.Range.Font.Size = [float]$params.fontSize
            }
            if ($null -ne $params.lineSpacing) {
                $pBody.LineSpacingRule = 1
                $pBody.LineSpacing = [float]($params.lineSpacing * 12)
            }
            $pBody.Range.InsertParagraphAfter()
        }
        
        if ($params.pageBreak) {
            $pBreak = $doc.Paragraphs.Add()
            $pBreak.Range.InsertBreak(7)
        }
        
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        return @{
            ok = $true
            path = $doc.FullName
            title = $doc.Name
            heading = $params.heading
            level = $level
            totalParagraphs = $doc.Paragraphs.Count
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordInsertTable($params) {
    $srcPath = Normalize-DocPath $params.path
    $rows = if ($params.rows) { [int]$params.rows } else { 2 }
    $cols = if ($params.cols) { [int]$params.cols } else { 2 }
    $data = $params.data
    
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) {
        Start-Sleep -Milliseconds 250
        $word = Get-RunningApp "Word.Application" $srcPath
    }
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $wasRunning = $true
    } else {
        try { $word.Visible = $true } catch {}
    }
    
    $doc = Find-DocInCollection $word $srcPath "word"
    $openedHere = $false
    if ($null -eq $doc) {
        if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
            $openedHere = $true
        } elseif ($word.Documents.Count -gt 0) {
            $doc = $word.ActiveDocument
        } else {
            return @{ ok = $false; error = "Document not open and file not found on disk: $srcPath" }
        }
    }
    
    try {
        $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        $table = $doc.Tables.Add($range, $rows, $cols)
        
        # Table Alignment (center, left, right)
        $tAlign = if ($params.alignment) { $params.alignment.ToString().ToLower() } else { "center" }
        if ($tAlign -match "center|centrado") { $table.Rows.Alignment = 1 }
        elseif ($tAlign -match "right|derecha") { $table.Rows.Alignment = 2 }
        else { $table.Rows.Alignment = 0 }
        
        # Grid borders by default or if requested
        try {
            $table.Borders.Enable = 1
        } catch {}
        
        try { $table.AutoFitBehavior(2) } catch {} # wdAutoFitWindow
        if ($null -ne $params.headerBgColor -or $null -ne $params.headerBg) {
            $rawHeaderBg = if ($params.headerBgColor) { $params.headerBgColor } else { $params.headerBg }
            $hBg = ConvertTo-WordColor $rawHeaderBg
            if ($null -ne $hBg) {
                try { $table.Rows.Item(1).Shading.BackgroundPatternColor = $hBg } catch {}
            }
        }
        if ($data -and $data.Count -gt 0) {
            for ($r = 0; $r -lt [Math]::Min($rows, $data.Count); $r++) {
                $rowArr = $data[$r]
                if ($r -gt 0 -and $params.zebra) {
                    $zCol = if ($params.zebraColor) { ConvertTo-WordColor $params.zebraColor } else { ConvertTo-WordColor "#F8FAFC" }
                    if ($r % 2 -eq 1 -and $null -ne $zCol) {
                        try { $table.Rows.Item($r + 1).Shading.BackgroundPatternColor = $zCol } catch {}
                    }
                }
                for ($c = 0; $c -lt [Math]::Min($cols, $rowArr.Count); $c++) {
                    $cell = $table.Cell($r + 1, $c + 1)
                    $cell.Range.Text = [string]$rowArr[$c]
                    $cell.VerticalAlignment = 1 # Center
                    if ($r -eq 0) {
                        if ($params.headerBold -ne $false) { $cell.Range.Bold = 1 }
                        if ($params.headerColor) {
                            $hCol = ConvertTo-WordColor $params.headerColor
                            if ($null -ne $hCol) { $cell.Range.Font.Color = $hCol }
                        }
                    }
                }
            }
        }
        
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        return @{ ok = $true; path = $doc.FullName; title = $doc.Name; tablesCount = $doc.Tables.Count; rows = $rows; cols = $cols }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-InsertShape($params) {
    $srcPath = Normalize-DocPath $params.path
    $shapeType = if ($params.shapeType) { $params.shapeType.ToLower() } else { "rectangle" }
    
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) {
        Start-Sleep -Milliseconds 250
        $word = Get-RunningApp "Word.Application" $srcPath
    }
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
        $wasRunning = $true
    } else {
        try { $word.Visible = $true } catch {}
    }
    
    $doc = Find-DocInCollection $word $srcPath "word"
    $openedHere = $false
    if ($null -eq $doc) {
        if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
            $openedHere = $true
        } elseif ($word.Documents.Count -gt 0) {
            $doc = $word.ActiveDocument
        } else {
            return @{ ok = $false; error = "Document not open and file not found on disk: $srcPath" }
        }
    }
    
    try {
        # Comprehensive shape mapping
        # msoShapeRectangle = 1, msoShapeRoundedRectangle = 5, msoShapeOval = 9, msoShapeIsoscelesTriangle = 7
        # msoShapeDiamond = 4, msoShapeHexagon = 10, msoShapeRightArrow = 13, msoShapeHeart = 21, msoShape5pointStar = 92
        $typeInt = switch ($shapeType) {
            "oval" { 9 }
            "circle" { 9 }
            "circulo" { 9 }
            "rounded_rectangle" { 5 }
            "rectangulo_redondeado" { 5 }
            "triangle" { 7 }
            "triangulo" { 7 }
            "diamond" { 4 }
            "diamante" { 4 }
            "rombo" { 4 }
            "hexagon" { 10 }
            "hexagono" { 10 }
            "arrow" { 13 }
            "flecha" { 13 }
            "arrow_right" { 13 }
            "heart" { 21 }
            "corazon" { 21 }
            "star" { 92 }
            "estrella" { 92 }
            default { 1 } # rectangle
        }
        
        # Dimensions
        $width = if ($params.width) { [float]$params.width } else { 220 }
        $height = if ($params.height) { [float]$params.height } else { 100 }
        $top = if ($params.top) { [float]$params.top } else { 120 }
        
        # Smart positioning (center, left, right)
        $pageWidth = 595.0
        try { $pageWidth = [float]$doc.PageSetup.PageWidth } catch {}
        
        $posParam = if ($params.position) { $params.position.ToString().ToLower() } elseif ($params.alignment) { $params.alignment.ToString().ToLower() } else { "" }
        $left = if ($null -ne $params.left) {
            [float]$params.left
        } elseif ($posParam -match "center|centrado") {
            [float](($pageWidth - $width) / 2)
        } elseif ($posParam -match "right|derecha") {
            [float]($pageWidth - $width - 54)
        } else {
            54.0 # Standard left margin
        }
        
        $shape = $doc.Shapes.AddShape($typeInt, $left, $top, $width, $height)
        
        $shapeText = if ($null -ne $params.text) { [string]$params.text } elseif ($null -ne $params.content) { [string]$params.content } else { "" }
        if ($shapeText) {
            $shape.TextFrame.TextRange.Text = $shapeText
            $fSize = if ($params.fontSize) { [float]$params.fontSize } elseif ($params.size) { [float]$params.size } else { $null }
            if ($null -ne $fSize) { $shape.TextFrame.TextRange.Font.Size = $fSize }
            if ($null -ne $params.bold) { $shape.TextFrame.TextRange.Font.Bold = [int][bool]$params.bold }
            if ($null -ne $params.fontName) { $shape.TextFrame.TextRange.Font.Name = [string]$params.fontName }
            
            $tColor = ConvertTo-WordColor $params.textColor
            if ($null -ne $tColor) { $shape.TextFrame.TextRange.Font.Color = $tColor }
            
            $alignVal = ConvertTo-WordAlignment $params.textAlignment
            if ($null -eq $alignVal) { $alignVal = ConvertTo-WordAlignment $params.alignment }
            if ($null -ne $alignVal) { $shape.TextFrame.TextRange.ParagraphFormat.Alignment = $alignVal }
        }
        
        $fillColParam = if ($params.fillColor) { $params.fillColor } elseif ($params.fill) { $params.fill } else { $null }
        if ($null -ne $fillColParam) {
            $fColor = ConvertTo-WordColor $fillColParam
            if ($null -ne $fColor) {
                $shape.Fill.Visible = -1
                $shape.Fill.Solid()
                $shape.Fill.ForeColor.RGB = $fColor
            }
        }
        
        $lineColParam = if ($params.lineColor) { $params.lineColor } elseif ($params.line) { $params.line } else { $null }
        if ($null -ne $lineColParam) {
            $lColor = ConvertTo-WordColor $lineColParam
            if ($null -ne $lColor) {
                $shape.Line.Visible = -1
                $shape.Line.ForeColor.RGB = $lColor
                if ($params.lineWidth) { $shape.Line.Weight = [float]$params.lineWidth }
            }
        }
        
        if ($params.save -and $doc.Path -ne "") {
            try { $doc.Save() } catch {}
        }
        return @{
            ok = $true
            path = $doc.FullName
            title = $doc.Name
            shapeName = $shape.Name
            shapeType = $shapeType
            left = $left
            top = $top
            width = $width
            height = $height
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-InsertImage($params) {
    if (-not $params.imagePath) { return @{ ok = $false; error = "imagePath is required" } }
    $imgPath = Normalize-DocPath $params.imagePath
    if (-not (Test-Path $imgPath -ErrorAction SilentlyContinue)) { return @{ ok = $false; error = "Image file not found: $imgPath" } }

    $srcPath = Normalize-DocPath $params.path
    $ext = [System.IO.Path]::GetExtension($srcPath).ToLower()
    if (-not $ext) { $ext = ".docx" }

    # Word
    if ($ext -in @(".doc", ".docx")) {
        $word = Get-RunningApp "Word.Application" $srcPath
        if ($null -eq $word) {
            Start-Sleep -Milliseconds 250
            $word = Get-RunningApp "Word.Application" $srcPath
        }
        $wasRunning = ($null -ne $word)
        if (-not $wasRunning) {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $true
            $wasRunning = $true
        } else {
            try { $word.Visible = $true } catch {}
        }
        
        $doc = Find-DocInCollection $word $srcPath "word"
        $openedHere = $false
        if ($null -eq $doc) {
            if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
                $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
                $openedHere = $true
            } elseif ($word.Documents.Count -gt 0) {
                $doc = $word.ActiveDocument
            } else {
                return @{ ok = $false; error = "Document not open and file not found on disk: $srcPath" }
            }
        }
        
        try {
            $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
            $shape = $doc.InlineShapes.AddPicture($imgPath, [ref]$false, [ref]$true, $range)
            if ($params.width) { $shape.Width = [float]$params.width }
            if ($params.height) { $shape.Height = [float]$params.height }
            if ($params.save -and $doc.Path -ne "") {
                try { $doc.Save() } catch {}
            }
            return @{
                ok = $true
                path = $doc.FullName
                title = $doc.Name
                imagePath = $imgPath
                type = "inline_shape"
                width = $shape.Width
                height = $shape.Height
            }
        } finally {
            Release-Com $doc
            Release-Com $word
            Invoke-GarbageCollection
        }
    }

    # PowerPoint
    if ($ext -in @(".ppt", ".pptx")) {
        $ppt = Get-RunningApp "PowerPoint.Application"
        $wasRunning = ($null -ne $ppt)
        if (-not $wasRunning) {
            $ppt = New-Object -ComObject PowerPoint.Application
            $ppt.Visible = 1
        }
        
        $pres = Get-PptPresentation $ppt $srcPath
        if ($null -eq $pres) {
            if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
                $pres = $ppt.Presentations.Open($srcPath)
            } else {
                return @{ ok = $false; error = "No open PowerPoint presentation found and file not on disk: $srcPath" }
            }
        }
        
        try {
            $slideIdx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
            if ($slideIdx -lt 1 -or $slideIdx -gt $pres.Slides.Count) {
                return @{ ok = $false; error = "Slide index $slideIdx out of range (1..$($pres.Slides.Count))" }
            }
            $slide = $pres.Slides.Item($slideIdx)
            $left = if ($null -ne $params.left) { [float]$params.left } else { 100 }
            $top = if ($null -ne $params.top) { [float]$params.top } else { 100 }
            $w = if ($params.width) { [float]$params.width } else { 300 }
            $h = if ($params.height) { [float]$params.height } else { 200 }
            
            $shape = $slide.Shapes.AddPicture($imgPath, 0, -1, $left, $top, $w, $h)
            if ($params.save -and $pres.Path -ne "") {
                try { $pres.Save() } catch {}
            }
            return @{
                ok = $true
                path = $pres.FullName
                imagePath = $imgPath
                slideIndex = $slideIdx
                shapeName = $shape.Name
                width = $shape.Width
                height = $shape.Height
            }
        } finally {
            Release-Com $pres
            Release-Com $ppt
            Invoke-GarbageCollection
        }
    }

    return @{ ok = $false; error = "Unsupported format for insert_image: $ext" }
}


# ── Word Ribbon Suite (Inicio, Insertar, Diseño, Disposición, Referencias, Correspondencia, Revisar, Vista) ──

function Invoke-WordApplyStyle($params) {
    $srcPath = Normalize-DocPath $params.path
    $styleName = if ($params.style) { [string]$params.style } else { "Heading 1" }
    
    $word = Get-RunningApp "Word.Application" $srcPath
    if ($null -eq $word) { $word = Get-RunningApp "Word.Application" $srcPath }
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $searchText = if ($params.search) { [string]$params.search } elseif ($params.text) { [string]$params.text } else { "" }
        if ($searchText) {
            $find = $doc.Content.Find
            $find.ClearFormatting()
            $find.Text = $searchText
            $find.Forward = $true
            $find.Wrap = 0
            if ($find.Execute()) {
                $rng = $find.Parent
                try { $rng.Style = $styleName } catch {
                    try {
                        $spanishMap = @{ "Heading 1" = "Título 1"; "Heading 2" = "Título 2"; "Heading 3" = "Título 3"; "Title" = "Título"; "Subtitle" = "Subtítulo"; "Normal" = "Normal" }
                        if ($spanishMap.ContainsKey($styleName)) { $rng.Style = $spanishMap[$styleName] }
                    } catch {}
                }
            }
        } else {
            try { $word.Selection.Style = $styleName } catch {}
        }
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; style = $styleName }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordInsertList($params) {
    $srcPath = Normalize-DocPath $params.path
    $items = $params.items
    $listType = if ($params.type) { $params.type.ToLower() } else { "bullet" }
    
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $inserted = 0
        if ($items -and $items.Count -gt 0) {
            foreach ($it in $items) {
                $p = $doc.Paragraphs.Add()
                $p.Range.Text = [string]$it
                if ($listType -eq "number" -or $listType -eq "numbered") {
                    try { $p.Range.ListFormat.ApplyNumberDefault() } catch {}
                } else {
                    try { $p.Range.ListFormat.ApplyBulletDefault() } catch {}
                }
                if ($params.lineSpacing) {
                    $p.LineSpacingRule = 1
                    $p.LineSpacing = [float]($params.lineSpacing * 12)
                }
                if ($null -ne $params.spaceAfter) {
                    $p.SpaceAfter = [float]$params.spaceAfter
                } else {
                    $p.SpaceAfter = 3
                }
                $p.Range.InsertParagraphAfter()
                $inserted++
            }
        }
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; insertedItems = $inserted; listType = $listType }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordHighlightText($params) {
    $srcPath = Normalize-DocPath $params.path
    $search = $params.search
    $colorName = if ($params.color) { $params.color.ToLower() } else { "yellow" }
    
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    $hlCode = switch ($colorName) {
        "yellow"      { 7 }
        "brightgreen" { 4 }
        "turquoise"   { 3 }
        "pink"        { 5 }
        "blue"        { 9 }
        "red"         { 6 }
        "gray"        { 16 }
        "none"        { 0 }
        default       { 7 }
    }
    
    try {
        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Text = $search
        $matches = 0
        while ($find.Execute()) {
            $find.Parent.HighlightColorIndex = $hlCode
            $matches++
        }
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; matches = $matches; color = $colorName }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordInsertCallout($params) {
    $srcPath = Normalize-DocPath $params.path
    $text = if ($params.text) { [string]$params.text } else { "" }
    $title = if ($params.title) { [string]$params.title } else { "" }
    
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        $table = $doc.Tables.Add($range, 1, 1)
        $table.Rows.Alignment = 1
        $table.AutoFitBehavior(2)
        
        $table.Borders.Enable = 0
        $leftBorder = $table.Borders.Item(-7)
        $leftBorder.LineStyle = 1
        $leftBorder.LineWidth = 24
        $borderColor = if ($params.borderColor) { ConvertTo-WordColor $params.borderColor } else { ConvertTo-WordColor "#6C1D45" }
        if ($null -ne $borderColor) { $leftBorder.Color = $borderColor }
        
        $bgColor = if ($params.bgColor) { ConvertTo-WordColor $params.bgColor } else { ConvertTo-WordColor "#F7F9FA" }
        if ($null -ne $bgColor) {
            $table.Rows.Item(1).Shading.BackgroundPatternColor = $bgColor
        }
        
        $table.TopPadding = 8
        $table.BottomPadding = 8
        $table.LeftPadding = 14
        $table.RightPadding = 14
        
        $cellRange = $table.Cell(1, 1).Range
        if ($title) {
            $cellRange.Text = "$title`r`n$text"
            $tPara = $cellRange.Paragraphs.Item(1)
            $tPara.Range.Bold = 1
            if ($null -ne $borderColor) { $tPara.Range.Font.Color = $borderColor }
            $tPara.Range.Font.Size = 11.5
            $tPara.SpaceAfter = 3
        } else {
            $cellRange.Text = $text
        }
        $cellRange.Paragraphs.Item($cellRange.Paragraphs.Count).SpaceAfter = 0
        
        $pAfter = $doc.Paragraphs.Add()
        $pAfter.SpaceAfter = 6
        
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; title = $title; textLength = $text.Length }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordInsertDivider($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $p = $doc.Paragraphs.Add()
        $p.Range.Text = " "
        $p.Range.Font.Size = 2
        $p.Format.Borders.Item(-3).LineStyle = 1
        $weight = if ($params.weight) { [int]$params.weight } else { 12 }
        $p.Format.Borders.Item(-3).LineWidth = $weight
        $divColor = if ($params.color) { ConvertTo-WordColor $params.color } else { ConvertTo-WordColor "#6C1D45" }
        if ($null -ne $divColor) { $p.Format.Borders.Item(-3).Color = $divColor }
        $p.SpaceBefore = if ($null -ne $params.spaceBefore) { [float]$params.spaceBefore } else { 6 }
        $p.SpaceAfter = if ($null -ne $params.spaceAfter) { [float]$params.spaceAfter } else { 12 }
        $p.Range.InsertParagraphAfter()
        
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordInsertHeaderFooter($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $sec = $doc.Sections.Item(1)
        try {
            $diffFirst = if ($null -ne $params.differentFirstPage -and -not [bool]$params.differentFirstPage) { 0 } else { -1 }
            $sec.PageSetup.DifferentFirstPageHeaderFooter = [int]$diffFirst
        } catch {}
        
        if ($params.headerText) {
            $hdr = $sec.Headers.Item(1)
            $hdr.Range.Text = [string]$params.headerText
            $hdr.Range.Font.Name = "Aptos"
            $hdr.Range.Font.Size = 9
            $hdr.Range.Font.Color = ConvertTo-WordColor "#7F8C8D"
            $hdr.Range.ParagraphFormat.Alignment = if ($params.headerAlign -match "center") { 1 } elseif ($params.headerAlign -match "right") { 2 } else { 0 }
        }
        
        if ($params.footerText) {
            $ftr = $sec.Footers.Item(1)
            $ftr.Range.Text = [string]$params.footerText
            $ftr.Range.Font.Name = "Aptos"
            $ftr.Range.Font.Size = 9
            $ftr.Range.Font.Color = ConvertTo-WordColor "#7F8C8D"
            $ftr.Range.ParagraphFormat.Alignment = if ($params.footerAlign -match "left") { 0 } elseif ($params.footerAlign -match "center") { 1 } else { 2 }
        }
        
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; headerText = $params.headerText; footerText = $params.footerText }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordInsertCoverPage($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $institution = if ($params.institution) { [string]$params.institution } else { "INSTITUTO POLITECNICO NACIONAL" }
        $school = if ($params.school) { [string]$params.school } else { "CENTRO DE ESTUDIOS CIENTIFICOS Y TECNOLOGICOS N. 8`r`n'NARCISO BASSOLS'" }
        $subject = if ($params.subject) { [string]$params.subject } else { "COMPUTACION BASICA 1" }
        $title = if ($params.title) { [string]$params.title } else { "INVESTIGACION: CIBERESPACIO" }
        $subtitle = if ($params.subtitle) { [string]$params.subtitle } else { "Servicios de Internet, Identidad Digital, Ciberseguridad e Inteligencia Artificial" }
        $group = if ($params.group) { [string]$params.group } else { "1IM12" }
        $team = if ($params.team) { $params.team } else { @("Lara Lozano Mauricio Leonardo") }
        $date = if ($params.date) { [string]$params.date } else { "16 de Octubre de 2024" }
        $primColor = if ($params.primaryColor) { ConvertTo-WordColor $params.primaryColor } else { ConvertTo-WordColor "#6C1D45" }
        
        $rng = $doc.Range(0, 0)
        
        $pInst = $doc.Paragraphs.Add($rng)
        $pInst.Range.Text = "$institution`r`n$school"
        $pInst.Range.Font.Name = "Aptos"
        $pInst.Range.Font.Bold = 1
        $pInst.Range.Font.Size = 14
        $pInst.Range.Font.Color = $primColor
        $pInst.Alignment = 1
        $pInst.SpaceBefore = 18
        $pInst.SpaceAfter = 8
        $pInst.Range.InsertParagraphAfter()
        
        $pSubj = $doc.Paragraphs.Add()
        $pSubj.Range.Text = "UNIDAD DE APRENDIZAJE: $subject"
        $pSubj.Range.Font.Name = "Aptos"
        $pSubj.Range.Font.Bold = 1
        $pSubj.Range.Font.Size = 11.5
        $pSubj.Range.Font.Color = ConvertTo-WordColor "#2B3A42"
        $pSubj.Alignment = 1
        $pSubj.SpaceAfter = 36
        $pSubj.Range.InsertParagraphAfter()
        
        $pTitle = $doc.Paragraphs.Add()
        $pTitle.Range.Text = $title
        $pTitle.Range.Font.Name = "Aptos Display"
        $pTitle.Range.Font.Bold = 1
        $pTitle.Range.Font.Size = 24
        $pTitle.Range.Font.Color = $primColor
        $pTitle.Alignment = 1
        $pTitle.SpaceAfter = 12
        $pTitle.Range.InsertParagraphAfter()
        
        $pSub = $doc.Paragraphs.Add()
        $pSub.Range.Text = $subtitle
        $pSub.Range.Font.Name = "Aptos"
        $pSub.Range.Font.Italic = 1
        $pSub.Range.Font.Size = 12.5
        $pSub.Range.Font.Color = ConvertTo-WordColor "#4A5568"
        $pSub.Alignment = 1
        $pSub.SpaceAfter = 70
        $pSub.Range.InsertParagraphAfter()
        
        $metaText = "DATOS DE ENTREGA:`r`nGrupo: $group`r`nModalidad: Equipo de 2`r`nIntegrantes:`r`n"
        foreach ($m in $team) {
            $metaText += " • $m`r`n"
        }
        $metaText += "`r`nFecha de Entrega: $date"
        
        $pMeta = $doc.Paragraphs.Add()
        $pMeta.Range.Text = $metaText
        $pMeta.Range.Font.Name = "Aptos"
        $pMeta.Range.Font.Size = 11
        $pMeta.Alignment = 1
        $pMeta.SpaceAfter = 24
        $pMeta.Range.InsertParagraphAfter()
        
        $pBrk = $doc.Paragraphs.Add()
        $pBrk.Range.InsertBreak(7)
        
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; title = $title; school = $school }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordSetPageBackground($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $color = if ($params.color) { ConvertTo-WordColor $params.color } else { 16777215 }
        $doc.Background.Fill.Visible = 1
        $doc.Background.Fill.ForeColor.RGB = $color
        $doc.ActiveWindow.View.DisplayBackgrounds = 1
        
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; color = $params.color }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordInsertSectionBreak($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    $breakType = switch ($params.type) {
        "continuous" { 3 }
        "even_page"  { 4 }
        "odd_page"   { 5 }
        default      { 2 }
    }
    
    try {
        $p = $doc.Paragraphs.Add()
        $p.Range.InsertBreak($breakType)
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; breakType = $params.type; totalSections = $doc.Sections.Count }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordInsertTOC($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $rng = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        $toc = $doc.TablesOfContents.Add($rng, $true, 1, 3)
        $toc.Update()
        
        if ($params.pageBreak -ne $false) {
            $p = $doc.Paragraphs.Add()
            $p.Range.InsertBreak(7)
        }
        
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; tocsCount = $doc.TablesOfContents.Count }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordUpdateTOC($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        foreach ($t in $doc.TablesOfContents) {
            try { $t.Update() } catch {}
        }
        try { $doc.Fields.Update() } catch {}
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; updated = $doc.TablesOfContents.Count }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordInsertFootnote($params) {
    $srcPath = Normalize-DocPath $params.path
    $search = $params.searchText
    $text = $params.footnoteText
    
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Text = $search
        if ($find.Execute()) {
            $fn = $doc.Footnotes.Add($find.Parent)
            $fn.Range.Text = $text
            if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
            return @{ ok = $true; path = $doc.FullName; search = $search; totalFootnotes = $doc.Footnotes.Count }
        }
        return @{ ok = $false; error = "Search text not found for footnote: $search" }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordReplacePlaceholders($params) {
    $srcPath = Normalize-DocPath $params.path
    $replacements = $params.replacements
    
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $count = 0
        if ($replacements) {
            foreach ($prop in $replacements.PSObject.Properties) {
                $k = $prop.Name
                $v = [string]$prop.Value
                $find = $doc.Content.Find
                $find.ClearFormatting()
                $find.Text = $k
                $find.Replacement.ClearFormatting()
                $find.Replacement.Text = $v
                $find.Execute([ref]$k, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$true, [ref]1, [ref]$false, [ref]$v, [ref]2) | Out-Null
                $count++
            }
        }
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; replacementsApplied = $count }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordGetStats($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $words = $doc.ComputeStatistics(1)
        $pages = $doc.ComputeStatistics(2)
        $chars = $doc.ComputeStatistics(0)
        $charsWithSpaces = $doc.ComputeStatistics(5)
        $paras = $doc.ComputeStatistics(3)
        $lines = $doc.ComputeStatistics(4)
        
        return @{
            ok = $true
            path = $doc.FullName
            title = $doc.Name
            pages = $pages
            words = $words
            paragraphs = $paras
            lines = $lines
            characters = $chars
            charactersWithSpaces = $charsWithSpaces
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordSetLanguage($params) {
    $srcPath = Normalize-DocPath $params.path
    $langId = if ($params.languageId) { [int]$params.languageId } else { 2058 }
    
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $doc.Content.LanguageID = $langId
        $doc.ShowSpellingErrors = $false
        $doc.ShowGrammaticalErrors = $false
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; languageId = $langId }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordSetView($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application" $srcPath
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) { return @{ ok = $false; error = "Document not found: $srcPath" } }
    
    try {
        $word.Visible = $true
        $win = $word.ActiveWindow
        $vType = switch ($params.viewType) {
            "web"     { 6 }
            "reading" { 7 }
            default   { 3 }
        }
        $win.View.Type = $vType
        if ($params.zoomPercent) {
            $win.View.Zoom.Percentage = [int]$params.zoomPercent
        }
        if ($null -ne $params.showRuler) {
            $word.DisplayRulers = [bool]$params.showRuler
        }
        return @{ ok = $true; path = $doc.FullName; viewType = $params.viewType; zoom = $win.View.Zoom.Percentage }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

# --- PowerPoint Helpers & Actions ---

function Get-PptPresentation($ppt, $srcPath, $autoCreate = $false) {
    $pres = $null
    if ($srcPath) {
        foreach ($p in $ppt.Presentations) {
            if ((Normalize-DocPath $p.FullName) -eq $srcPath -or $p.Name -eq $srcPath) { $pres = $p; break }
        }
        if ($null -eq $pres -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            try { $pres = $ppt.Presentations.Open($srcPath) } catch {}
        }
    }
    if ($null -eq $pres) {
        if ($ppt.Presentations.Count -gt 0) {
            try { $pres = $ppt.ActivePresentation } catch {}
            if ($null -eq $pres) {
                try { $pres = $ppt.Presentations.Item(1) } catch {}
            }
        } elseif ($autoCreate) {
            try { $pres = $ppt.Presentations.Add(1) } catch {}
        }
    }
    return $pres
}

function Invoke-PptReadSlides($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -eq $ppt) {
        if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $ppt = New-Object -ComObject PowerPoint.Application
            $ppt.Visible = 1
        } else {
            return @{ ok = $true; path = ""; title = "No open presentation"; slideCount = 0; slides = @() }
        }
    }
    
    $pres = Get-PptPresentation $ppt $srcPath $false
    if ($null -eq $pres) {
        return @{ ok = $true; path = ""; title = "No open presentation"; slideCount = 0; slides = @() }
    }
    
    try {
        $slidesData = @()
        foreach ($slide in $pres.Slides) {
            $slideNum = $slide.SlideNumber
            $title = ""
            $texts = @()
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasTextFrame -eq -1) {
                    $t = $shape.TextFrame.TextRange.Text.Trim()
                    if ($t) {
                        if (-not $title) { $title = $t }
                        $texts += $t
                    }
                }
            }
            $slidesData += @{
                slideNumber = $slideNum
                title = $title
                shapeCount = $slide.Shapes.Count
                texts = $texts
            }
        }
        
        return @{
            ok = $true
            path = $pres.FullName
            title = $pres.Name
            slideCount = $pres.Slides.Count
            slides = $slidesData
        }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

function Invoke-PptAddSlide($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    $wasRunning = ($null -ne $ppt)
    if (-not $wasRunning) {
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = 1
        $pres = if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) { $ppt.Presentations.Open($srcPath) } else { $ppt.Presentations.Add(1) }
    } else {
        $pres = Get-PptPresentation $ppt $srcPath
        if ($null -eq $pres) {
            $pres = $ppt.Presentations.Add(1)
        }
    }
    
    try {
        $layoutNum = 2 # ppLayoutText (title + content)
        $layoutParam = if ($params.layout) { $params.layout.ToString().ToLower() } else { "text" }
        if ($layoutParam -match "title|portada") { $layoutNum = 1 }
        elseif ($layoutParam -match "blank|blanco") { $layoutNum = 12 }
        elseif ($layoutParam -match "header|encabezado") { $layoutNum = 3 }
        elseif ($layoutParam -match "two|dos|comparison") { $layoutNum = 4 }
        
        $newIndex = if ($params.slideIndex) { [int]$params.slideIndex } else { $pres.Slides.Count + 1 }
        $slide = $pres.Slides.Add($newIndex, $layoutNum)
        
        $titleText = if ($params.title) { $params.title } elseif ($params.heading) { $params.heading } else { "" }
        if ($titleText -and $slide.Shapes.HasTitle) {
            $slide.Shapes.Title.TextFrame.TextRange.Text = $titleText
            if ($params.titleColor) {
                $c = ConvertTo-WordColor $params.titleColor
                if ($null -ne $c) { $slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = $c }
            }
        }
        
        $bodyText = if ($null -ne $params.body) { [string]$params.body } elseif ($null -ne $params.content) { [string]$params.content } elseif ($null -ne $params.text) { [string]$params.text } else { "" }
        if ($bodyText -and $slide.Shapes.Placeholders.Count -ge 2) {
            $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $bodyText
            if ($params.bodyColor) {
                $c = ConvertTo-WordColor $params.bodyColor
                if ($null -ne $c) { $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Font.Color.RGB = $c }
            }
        }
        
        if ($params.save -and $pres.Path -ne "") {
            try { $pres.Save() } catch {}
        }
        return @{
            ok = $true
            path = $pres.FullName
            title = $pres.Name
            slideNumber = $slide.SlideNumber
            totalSlides = $pres.Slides.Count
            slideTitle = $titleText
        }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

function Invoke-PptSetSlideText($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -eq $ppt) { return @{ ok = $false; error = "PowerPoint is not running" } }
    $pres = Get-PptPresentation $ppt $srcPath
    if ($null -eq $pres) { return @{ ok = $false; error = "No open PowerPoint presentation found" } }
    
    $slideIdx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
    $shapeIdx = if ($params.shapeIndex) { [int]$params.shapeIndex } else { 1 }
    
    try {
        if ($slideIdx -gt $pres.Slides.Count -or $slideIdx -lt 1) {
            return @{ ok = $false; error = "Slide index $slideIdx out of range (1..$($pres.Slides.Count))" }
        }
        $slide = $pres.Slides.Item($slideIdx)
        if ($shapeIdx -gt $slide.Shapes.Count -or $shapeIdx -lt 1) {
            return @{ ok = $false; error = "Shape index $shapeIdx out of range" }
        }
        $shape = $slide.Shapes.Item($shapeIdx)
        $textToSet = if ($null -ne $params.text) { [string]$params.text } else { [string]$params.content }
        $shape.TextFrame.TextRange.Text = $textToSet
        
        if ($params.fontSize) { $shape.TextFrame.TextRange.Font.Size = [float]$params.fontSize }
        if ($null -ne $params.bold) { $shape.TextFrame.TextRange.Font.Bold = [int][bool]$params.bold }
        if ($params.color) {
            $c = ConvertTo-WordColor $params.color
            if ($null -ne $c) { $shape.TextFrame.TextRange.Font.Color.RGB = $c }
        }
        
        if ($params.save -and $pres.Path -ne "") {
            try { $pres.Save() } catch {}
        }
        return @{ ok = $true; path = $pres.FullName; slideIndex = $slideIdx; shapeIndex = $shapeIdx; updatedText = $textToSet }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

function Invoke-PptDeleteSlide($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -eq $ppt) { return @{ ok = $false; error = "PowerPoint is not running" } }
    $pres = Get-PptPresentation $ppt $srcPath
    if ($null -eq $pres) { return @{ ok = $false; error = "No open PowerPoint presentation found" } }
    
    $slideIdx = [int]$params.slideIndex
    try {
        if ($slideIdx -gt $pres.Slides.Count -or $slideIdx -lt 1) {
            return @{ ok = $false; error = "Slide index $slideIdx out of bounds (1..$($pres.Slides.Count))" }
        }
        $pres.Slides.Item($slideIdx).Delete()
        if ($params.save -and $pres.Path -ne "") {
            try { $pres.Save() } catch {}
        }
        return @{ ok = $true; path = $pres.FullName; deletedIndex = $slideIdx; remainingSlides = $pres.Slides.Count }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

function Invoke-PptInsertShape($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -eq $ppt) {
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = 1
        $pres = if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) { $ppt.Presentations.Open($srcPath) } else { $ppt.Presentations.Add(1) }
    } else {
        $pres = Get-PptPresentation $ppt $srcPath
        if ($null -eq $pres) { $pres = $ppt.Presentations.Add(1) }
    }
    
    $slideIdx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
    if ($slideIdx -gt $pres.Slides.Count) {
        $pres.Slides.Add($pres.Slides.Count + 1, 12) | Out-Null
        $slideIdx = $pres.Slides.Count
    }
    $slide = $pres.Slides.Item($slideIdx)
    
    $shapeType = if ($params.shapeType) { $params.shapeType.ToLower() } else { "rectangle" }
    $typeInt = switch ($shapeType) {
        "oval" { 9 }
        "circle" { 9 }
        "circulo" { 9 }
        "rounded_rectangle" { 5 }
        "rectangulo_redondeado" { 5 }
        "triangle" { 7 }
        "triangulo" { 7 }
        "diamond" { 4 }
        "diamante" { 4 }
        "rombo" { 4 }
        "hexagon" { 10 }
        "hexagono" { 10 }
        "arrow" { 13 }
        "flecha" { 13 }
        "arrow_right" { 13 }
        "heart" { 21 }
        "corazon" { 21 }
        "star" { 92 }
        "estrella" { 92 }
        default { 1 }
    }
    
    $slideWidth = 960.0
    $slideHeight = 540.0
    try {
        $slideWidth = [float]$pres.PageSetup.SlideWidth
        $slideHeight = [float]$pres.PageSetup.SlideHeight
    } catch {}
    
    $width = if ($params.width) { [float]$params.width } else { 320 }
    $height = if ($params.height) { [float]$params.height } else { 120 }
    
    $posParam = if ($params.position) { $params.position.ToString().ToLower() } elseif ($params.alignment) { $params.alignment.ToString().ToLower() } else { "" }
    $left = if ($null -ne $params.left) {
        [float]$params.left
    } elseif ($posParam -match "center|centrado") {
        [float](($slideWidth - $width) / 2)
    } elseif ($posParam -match "right|derecha") {
        [float]($slideWidth - $width - 50)
    } else {
        50.0
    }
    
    $top = if ($null -ne $params.top) {
        [float]$params.top
    } elseif ($posParam -match "center|centrado") {
        [float](($slideHeight - $height) / 2)
    } else {
        100.0
    }
    
    try {
        $shape = $slide.Shapes.AddShape($typeInt, $left, $top, $width, $height)
        
        $shapeText = if ($null -ne $params.text) { [string]$params.text } elseif ($null -ne $params.content) { [string]$params.content } else { "" }
        if ($shapeText) {
            $shape.TextFrame.TextRange.Text = $shapeText
            if ($params.fontSize) { $shape.TextFrame.TextRange.Font.Size = [float]$params.fontSize }
            if ($null -ne $params.bold) { $shape.TextFrame.TextRange.Font.Bold = [int][bool]$params.bold }
            if ($params.fontName) { $shape.TextFrame.TextRange.Font.Name = [string]$params.fontName }
            if ($params.textColor) {
                $c = ConvertTo-WordColor $params.textColor
                if ($null -ne $c) { $shape.TextFrame.TextRange.Font.Color.RGB = $c }
            }
            if ($posParam -match "center|centrado" -or $params.textAlignment -match "center|centrado") {
                $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # ppAlignCenter
            }
        }
        
        $fillCol = if ($params.fillColor) { $params.fillColor } else { $params.fill }
        if ($fillCol) {
            $c = ConvertTo-WordColor $fillCol
            if ($null -ne $c) {
                $shape.Fill.Solid()
                $shape.Fill.ForeColor.RGB = $c
            }
        }
        
        $lineCol = if ($params.lineColor) { $params.lineColor } else { $params.line }
        if ($lineCol) {
            $c = ConvertTo-WordColor $lineCol
            if ($null -ne $c) {
                $shape.Line.Visible = -1
                $shape.Line.ForeColor.RGB = $c
                if ($params.lineWidth) { $shape.Line.Weight = [float]$params.lineWidth }
            }
        }
        
        if ($params.save -and $pres.Path -ne "") {
            try { $pres.Save() } catch {}
        }
        return @{
            ok = $true
            app = "powerpoint"
            path = $pres.FullName
            title = $pres.Name
            slideIndex = $slideIdx
            shapeType = $shapeType
            shapeName = $shape.Name
            left = $left
            top = $top
            width = $width
            height = $height
        }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

function Invoke-PptInsertTable($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -eq $ppt) {
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = 1
        $pres = if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) { $ppt.Presentations.Open($srcPath) } else { $ppt.Presentations.Add(1) }
    } else {
        $pres = Get-PptPresentation $ppt $srcPath
        if ($null -eq $pres) { $pres = $ppt.Presentations.Add(1) }
    }
    
    $slideIdx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
    if ($slideIdx -gt $pres.Slides.Count) {
        $pres.Slides.Add($pres.Slides.Count + 1, 12) | Out-Null
        $slideIdx = $pres.Slides.Count
    }
    $slide = $pres.Slides.Item($slideIdx)
    
    $rows = if ($params.rows) { [int]$params.rows } else { 3 }
    $cols = if ($params.cols) { [int]$params.cols } else { 3 }
    $data = if ($params.data) { $params.data } else { @() }
    if ($data.Count -gt $rows) { $rows = $data.Count }
    if ($data.Count -gt 0 -and $data[0].Count -gt $cols) { $cols = $data[0].Count }
    
    $slideWidth = 960.0
    $slideHeight = 540.0
    try {
        $slideWidth = [float]$pres.PageSetup.SlideWidth
        $slideHeight = [float]$pres.PageSetup.SlideHeight
    } catch {}
    
    $width = if ($params.width) { [float]$params.width } else { 700 }
    $height = if ($params.height) { [float]$params.height } else { [Math]::Min(300, $rows * 40) }
    $left = if ($null -ne $params.left) { [float]$params.left } else { [float](($slideWidth - $width) / 2) }
    $top = if ($null -ne $params.top) { [float]$params.top } else { 120.0 }
    
    try {
        $tableShape = $slide.Shapes.AddTable($rows, $cols, $left, $top, $width, $height)
        $table = $tableShape.Table
        
        if ($data -and $data.Count -gt 0) {
            for ($r = 0; $r -lt [Math]::Min($rows, $data.Count); $r++) {
                $rowArr = $data[$r]
                for ($c = 0; $c -lt [Math]::Min($cols, $rowArr.Count); $c++) {
                    $cell = $table.Cell($r + 1, $c + 1)
                    $cell.Shape.TextFrame.TextRange.Text = [string]$rowArr[$c]
                    if ($r -eq 0 -and $params.headerBold -ne $false) {
                        $cell.Shape.TextFrame.TextRange.Font.Bold = 1
                    }
                    if ($r -eq 0 -and $params.headerColor) {
                        $hCol = ConvertTo-WordColor $params.headerColor
                        if ($null -ne $hCol) {
                            $cell.Shape.Fill.Solid()
                            $cell.Shape.Fill.ForeColor.RGB = $hCol
                            $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 16777215
                        }
                    }
                }
            }
        }
        
        if ($params.save -and $pres.Path -ne "") {
            try { $pres.Save() } catch {}
        }
        return @{
            ok = $true
            app = "powerpoint"
            path = $pres.FullName
            title = $pres.Name
            slideIndex = $slideIdx
            rows = $rows
            cols = $cols
            left = $left
            top = $top
            width = $width
            height = $height
        }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

# --- Excel Helpers & Actions ---

function Get-ExcelWorkbook($excel, $srcPath, $autoCreate = $false) {
    $wb = $null
    if ($srcPath) {
        foreach ($w in $excel.Workbooks) {
            if ((Normalize-DocPath $w.FullName) -eq $srcPath -or $w.Name -eq $srcPath) { $wb = $w; break }
        }
        if ($null -eq $wb -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            try { $wb = $excel.Workbooks.Open($srcPath) } catch {}
        }
    }
    if ($null -eq $wb) {
        if ($excel.Workbooks.Count -gt 0) {
            try { $wb = $excel.ActiveWorkbook } catch {}
            if ($null -eq $wb) {
                try { $wb = $excel.Workbooks.Item(1) } catch {}
            }
        } elseif ($autoCreate) {
            try { $wb = $excel.Workbooks.Add() } catch {}
        }
    }
    return $wb
}

function Invoke-ExcelReadRange($params) {
    $srcPath = Normalize-DocPath $params.path
    $excel = Get-RunningApp "Excel.Application"
    if ($null -eq $excel) {
        if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
        } else {
            return @{ ok = $true; path = ""; sheet = ""; range = ""; rows = 0; cols = 0; values = @() }
        }
    }
    
    $wb = Get-ExcelWorkbook $excel $srcPath $false
    if ($null -eq $wb) {
        return @{ ok = $true; path = ""; sheet = ""; range = ""; rows = 0; cols = 0; values = @() }
    }
    
    try {
        $sheet = if ($params.sheet) { $wb.Sheets.Item($params.sheet) } else { $wb.ActiveSheet }
        $rangeStr = if ($params.range) { $params.range } else { $sheet.UsedRange.Address }
        $range = $sheet.Range($rangeStr)
        
        $data = @()
        for ($r = 1; $r -le $range.Rows.Count; $r++) {
            $rowVals = @()
            for ($c = 1; $c -le $range.Columns.Count; $c++) {
                $v = $range.Cells.Item($r, $c).Value2
                $rowVals += if ($null -ne $v) { $v } else { "" }
            }
            $data += ,$rowVals
        }
        
        return @{
            ok = $true
            path = $wb.FullName
            sheet = $sheet.Name
            range = $rangeStr
            rows = $range.Rows.Count
            cols = $range.Columns.Count
            values = $data
        }
    } finally {
        Release-Com $wb
        Release-Com $excel
        Invoke-GarbageCollection
    }
}

function Invoke-ExcelWriteRange($params) {
    $srcPath = Normalize-DocPath $params.path
    if (-not $params.data) { return @{ ok = $false; error = "data 2D array is required" } }
    
    $excel = Get-RunningApp "Excel.Application"
    $wasRunning = ($null -ne $excel)
    if (-not $wasRunning) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $wb = if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) { $excel.Workbooks.Open($srcPath) } else { $excel.Workbooks.Add() }
    } else {
        $wb = Get-ExcelWorkbook $excel $srcPath
        if ($null -eq $wb) { $wb = $excel.Workbooks.Add() }
    }
    
    try {
        $sheet = if ($params.sheet) { $wb.Sheets.Item($params.sheet) } else { $wb.ActiveSheet }
        $startCell = if ($params.startCell) { $params.startCell } else { "A1" }
        $start = $sheet.Range($startCell)
        $startRow = $start.Row
        $startCol = $start.Column
        
        $data = $params.data
        for ($r = 0; $r -lt $data.Count; $r++) {
            $rowArr = $data[$r]
            for ($c = 0; $c -lt $rowArr.Count; $c++) {
                $val = $rowArr[$c]
                $cell = $sheet.Cells.Item($startRow + $r, $startCol + $c)
                $valStr = [string]$val
                if ($valStr.StartsWith("=")) {
                    $cell.Formula = $valStr
                } else {
                    try {
                        $cell.Value2 = $val
                    } catch {
                        try {
                            $cell.Value = $val
                        } catch {
                            $cell.Value = [string]$val
                        }
                    }
                }
                
                if ($r -eq 0) {
                    if ($params.headerBold -ne $false) { $cell.Font.Bold = $true }
                    $headerBg = if ($params.headerBgColor) { $params.headerBgColor } else { $params.headerColor }
                    if ($headerBg) {
                        $cVal = ConvertTo-WordColor $headerBg
                        if ($null -ne $cVal) {
                            $cell.Interior.Color = $cVal
                            $cell.Font.Color = 16777215 # White font
                        }
                    }
                }
            }
        }
        
        $endRow = $startRow + $data.Count - 1
        $maxCols = 1
        foreach ($row in $data) { if ($row.Count -gt $maxCols) { $maxCols = $row.Count } }
        $endCol = $startCol + $maxCols - 1
        $dataRange = $sheet.Range($sheet.Cells.Item($startRow, $startCol), $sheet.Cells.Item($endRow, $endCol))
        
        if ($params.borders) {
            $dataRange.Borders.LineStyle = 1 # xlContinuous
        }
        if ($params.autoFit) {
            $dataRange.Columns.AutoFit()
        }
        
        if ($params.save -and $wb.Path -ne "") {
            try { $wb.Save() } catch {}
        }
        return @{ ok = $true; path = $wb.FullName; sheet = $sheet.Name; writtenRows = $data.Count }
    } finally {
        Release-Com $wb
        Release-Com $excel
        Invoke-GarbageCollection
    }
}

function Invoke-ExcelWriteCell($params) {
    $srcPath = Normalize-DocPath $params.path
    $excel = Get-RunningApp "Excel.Application"
    $wasRunning = ($null -ne $excel)
    if (-not $wasRunning) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $wb = if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) { $excel.Workbooks.Open($srcPath) } else { $excel.Workbooks.Add() }
    } else {
        $wb = Get-ExcelWorkbook $excel $srcPath
        if ($null -eq $wb) { $wb = $excel.Workbooks.Add() }
    }
    
    try {
        $sheet = if ($params.sheet) { $wb.Sheets.Item($params.sheet) } else { $wb.ActiveSheet }
        
        $cell = if ($params.cell) {
            $sheet.Range([string]$params.cell)
        } elseif ($params.row -and $params.col) {
            $sheet.Cells.Item([int]$params.row, [int]$params.col)
        } else {
            $sheet.Range("A1")
        }
        
        $val = if ($null -ne $params.value) { $params.value } elseif ($null -ne $params.formula) { $params.formula } else { $params.text }
        $valStr = [string]$val
        if ($valStr.StartsWith("=")) {
            $cell.Formula = $valStr
        } else {
            try {
                $cell.Value2 = $val
            } catch {
                try {
                    $cell.Value = $val
                } catch {
                    $cell.Value = [string]$val
                }
            }
        }
        
        if ($null -ne $params.bold) { $cell.Font.Bold = [bool]$params.bold }
        if ($params.fontSize) { $cell.Font.Size = [float]$params.fontSize }
        if ($params.color) {
            $c = ConvertTo-WordColor $params.color
            if ($null -ne $c) { $cell.Font.Color = $c }
        }
        $bgCol = if ($params.bgColor) { $params.bgColor } else { $params.fillColor }
        if ($bgCol) {
            $c = ConvertTo-WordColor $bgCol
            if ($null -ne $c) { $cell.Interior.Color = $c }
        }
        if ($params.numberFormat) {
            $cell.NumberFormat = [string]$params.numberFormat
        }
        
        if ($params.save -and $wb.Path -ne "") {
            try { $wb.Save() } catch {}
        }
        $addr = ""
        try { $addr = [string]$cell.Address($false, $false) } catch { $addr = [string]$cell.Address }
        return @{
            ok = $true
            path = $wb.FullName
            sheet = $sheet.Name
            cell = $addr
            value = $val
        }
    } finally {
        Release-Com $wb
        Release-Com $excel
        Invoke-GarbageCollection
    }
}

function Invoke-ExcelAddSheet($params) {
    $srcPath = Normalize-DocPath $params.path
    $excel = Get-RunningApp "Excel.Application"
    if ($null -eq $excel) { return @{ ok = $false; error = "Excel is not running" } }
    $wb = Get-ExcelWorkbook $excel $srcPath
    if ($null -eq $wb) { return @{ ok = $false; error = "No open Excel workbook found" } }
    
    try {
        $newSheet = $wb.Sheets.Add([Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count))
        if ($params.name) {
            $newSheet.Name = [string]$params.name
        }
        if ($params.save -and $wb.Path -ne "") {
            try { $wb.Save() } catch {}
        }
        return @{
            ok = $true
            path = $wb.FullName
            sheetName = $newSheet.Name
            totalSheets = $wb.Sheets.Count
        }
    } finally {
        Release-Com $wb
        Release-Com $excel
        Invoke-GarbageCollection
    }
}

# --- Advanced Office Live Functions ---

function Invoke-InspectStructure($params) {
    $srcPath = Normalize-DocPath $params.path
    $appParam = if ($params.app) { $params.app.ToLower() } else { "" }
    
    if (-not $appParam -and $srcPath) {
        $ext = [System.IO.Path]::GetExtension($srcPath).ToLower()
        if ($ext -in @(".docx", ".doc", ".dotx")) { $appParam = "word" }
        elseif ($ext -in @(".xlsx", ".xls", ".xlsm")) { $appParam = "excel" }
        elseif ($ext -in @(".pptx", ".ppt", ".potx")) { $appParam = "powerpoint" }
    }
    
    if (-not $appParam) {
        $w = Get-RunningApp "Word.Application"
        if ($null -ne $w -and $w.Documents.Count -gt 0) { $appParam = "word" }
        else {
            $x = Get-RunningApp "Excel.Application"
            if ($null -ne $x -and $x.Workbooks.Count -gt 0) { $appParam = "excel" }
            else {
                $p = Get-RunningApp "PowerPoint.Application"
                if ($null -ne $p -and $p.Presentations.Count -gt 0) { $appParam = "powerpoint" }
                else { $appParam = "word" }
            }
        }
    }
    
    if ($appParam -eq "word") {
        $word = Get-RunningApp "Word.Application"
        $wasRunning = ($null -ne $word)
        if (-not $wasRunning) {
            if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $doc = $word.Documents.Open($srcPath, [Type]::Missing, $true)
            } else {
                return @{ ok = $false; error = "Word is not running and no valid path provided" }
            }
        } else {
            $doc = Find-DocInCollection $word $srcPath "word"
            if ($null -eq $doc) {
                if ($word.Documents.Count -gt 0) { $doc = $word.ActiveDocument }
                elseif ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
                    $doc = $word.Documents.Open($srcPath, [Type]::Missing, $true)
                } else {
                    return @{ ok = $false; error = "No open Word document found" }
                }
            }
        }
        
        try {
            $pageCount = 1
            try { $pageCount = [int]$doc.ComputeStatistics(2) } catch {}
            $wordCount = 0
            try { $wordCount = [int]$doc.ComputeStatistics(0) } catch {}
            $paraCount = $doc.Paragraphs.Count
            
            $headings = @()
            $maxHeadings = 50
            for ($i = 1; $i -le [Math]::Min($paraCount, 500); $i++) {
                try {
                    $p = $doc.Paragraphs.Item($i)
                    $styleName = [string]$p.Style.NameLocal
                    $txt = $p.Range.Text.Trim()
                    if ($txt.Length -gt 0 -and ($styleName -match "Heading|Título|Header|Encabezado" -or $p.OutlineLevel -le 3)) {
                        $lvl = 1
                        if ($styleName -match "2|Level 2|Nivel 2") { $lvl = 2 }
                        elseif ($styleName -match "3|Level 3|Nivel 3") { $lvl = 3 }
                        $headings += @{ index = $i; level = $lvl; style = $styleName; text = if ($txt.Length -gt 120) { $txt.Substring(0, 120) + "..." } else { $txt } }
                        if ($headings.Count -ge $maxHeadings) { break }
                    }
                } catch {}
            }
            
            $tables = @()
            foreach ($tbl in $doc.Tables) {
                try {
                    $tables += @{
                        rows = $tbl.Rows.Count
                        columns = $tbl.Columns.Count
                    }
                } catch {}
            }
            
            $headerText = ""
            $footerText = ""
            try {
                $sec = $doc.Sections.Item(1)
                $headerText = $sec.Headers.Item(1).Range.Text.Trim()
                $footerText = $sec.Footers.Item(1).Range.Text.Trim()
            } catch {}
            
            return @{
                ok = $true
                app = "word"
                path = $doc.FullName
                title = $doc.Name
                pages = $pageCount
                words = $wordCount
                paragraphs = $paraCount
                sections = $doc.Sections.Count
                headings = $headings
                tablesCount = $doc.Tables.Count
                tables = $tables
                header = $headerText
                footer = $footerText
            }
        } finally {
            if (-not $wasRunning -and $null -ne $doc) { try { $doc.Close([Type]::Missing, [Type]::Missing, [Type]::Missing) } catch {} }
            if (-not $wasRunning -and $null -ne $word) { try { $word.Quit() } catch {} }
            Release-Com $doc
            Release-Com $word
            Invoke-GarbageCollection
        }
    } elseif ($appParam -eq "excel") {
        $excel = Get-RunningApp "Excel.Application"
        $wasRunning = ($null -ne $excel)
        if (-not $wasRunning) {
            if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $wb = $excel.Workbooks.Open($srcPath, [Type]::Missing, $true)
            } else {
                return @{ ok = $false; error = "Excel is not running and no valid path provided" }
            }
        } else {
            $wb = Get-ExcelWorkbook $excel $srcPath $false
            if ($null -eq $wb) {
                if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
                    $wb = $excel.Workbooks.Open($srcPath, [Type]::Missing, $true)
                } else {
                    return @{ ok = $false; error = "No open Excel workbook found" }
                }
            }
        }
        
        try {
            $sheetsList = @()
            foreach ($sh in $wb.Sheets) {
                $ur = $sh.UsedRange
                $headers = @()
                try {
                    $firstRow = $ur.Rows.Item(1)
                    for ($c = 1; $c -le [Math]::Min($firstRow.Columns.Count, 20); $c++) {
                        $v = $firstRow.Cells.Item(1, $c).Value2
                        $headers += if ($null -ne $v) { [string]$v } else { "" }
                    }
                } catch {}
                
                $sheetTables = @()
                try {
                    foreach ($lo in $sh.ListObjects) {
                        $sheetTables += @{ name = $lo.Name; range = $lo.Range.Address; style = $lo.TableStyle.Name }
                    }
                } catch {}
                
                $sheetCharts = @()
                try {
                    foreach ($co in $sh.ChartObjects()) {
                        $sheetCharts += @{ name = $co.Name; title = try { $co.Chart.ChartTitle.Text } catch { "" } }
                    }
                } catch {}
                
                $sheetsList += @{
                    name = $sh.Name
                    visible = ($sh.Visible -eq -1)
                    usedRange = $ur.Address
                    rows = $ur.Rows.Count
                    columns = $ur.Columns.Count
                    headers = $headers
                    tables = $sheetTables
                    charts = $sheetCharts
                }
            }
            
            $namedRanges = @()
            try {
                foreach ($nr in $wb.Names) {
                    $namedRanges += @{ name = $nr.Name; refersTo = $nr.RefersTo }
                }
            } catch {}
            
            return @{
                ok = $true
                app = "excel"
                path = $wb.FullName
                title = $wb.Name
                totalSheets = $wb.Sheets.Count
                sheets = $sheetsList
                namedRanges = $namedRanges
            }
        } finally {
            if (-not $wasRunning -and $null -ne $wb) { try { $wb.Close($false) } catch {} }
            if (-not $wasRunning -and $null -ne $excel) { try { $excel.Quit() } catch {} }
            Release-Com $wb
            Release-Com $excel
            Invoke-GarbageCollection
        }
    } elseif ($appParam -eq "powerpoint") {
        $ppt = Get-RunningApp "PowerPoint.Application"
        $wasRunning = ($null -ne $ppt)
        if (-not $wasRunning) {
            if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
                $ppt = New-Object -ComObject PowerPoint.Application
                $ppt.Visible = 1
                $pres = $ppt.Presentations.Open($srcPath, 1, 0, 0)
            } else {
                return @{ ok = $false; error = "PowerPoint is not running and no valid path provided" }
            }
        } else {
            $pres = Get-PptPresentation $ppt $srcPath $false
            if ($null -eq $pres) {
                if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) {
                    $pres = $ppt.Presentations.Open($srcPath, 1, 0, 0)
                } else {
                    return @{ ok = $false; error = "No open PowerPoint presentation found" }
                }
            }
        }
        
        try {
            $slidesInfo = @()
            foreach ($sl in $pres.Slides) {
                $title = ""
                $shapeTypes = @()
                $bodyPreview = ""
                foreach ($shp in $sl.Shapes) {
                    $shapeTypes += $shp.Type
                    if ($shp.HasTextFrame -eq -1) {
                        $t = $shp.TextFrame.TextRange.Text.Trim()
                        if ($t) {
                            if (-not $title) { $title = $t }
                            elseif (-not $bodyPreview) { $bodyPreview = if ($t.Length -gt 100) { $t.Substring(0, 100) + "..." } else { $t } }
                        }
                    }
                }
                
                $notes = ""
                try {
                    $notes = $sl.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text.Trim()
                } catch {}
                
                $slidesInfo += @{
                    slideNumber = $sl.SlideNumber
                    title = $title
                    shapeCount = $sl.Shapes.Count
                    preview = $bodyPreview
                    hasNotes = ($notes.Length -gt 0)
                }
            }
            
            $aspectRatio = "16:9"
            try {
                $ratio = [float]$pres.PageSetup.SlideWidth / [float]$pres.PageSetup.SlideHeight
                if ($ratio -lt 1.4) { $aspectRatio = "4:3" }
            } catch {}
            
            return @{
                ok = $true
                app = "powerpoint"
                path = $pres.FullName
                title = $pres.Name
                slideCount = $pres.Slides.Count
                aspectRatio = $aspectRatio
                slides = $slidesInfo
            }
        } finally {
            if (-not $wasRunning -and $null -ne $pres) { try { $pres.Close() } catch {} }
            if (-not $wasRunning -and $null -ne $ppt) { try { $ppt.Quit() } catch {} }
            Release-Com $pres
            Release-Com $ppt
            Invoke-GarbageCollection
        }
    } else {
        return @{ ok = $false; error = "Target app '$appParam' not supported for inspect_structure" }
    }
}

function Invoke-WordTableFormat($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application"
    if ($null -eq $word) { return @{ ok = $false; error = "Word is not running" } }
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) {
        if ($word.Documents.Count -gt 0) { $doc = $word.ActiveDocument }
        else { return @{ ok = $false; error = "No open Word document found" } }
    }
    
    $tableIdx = if ($params.tableIndex) { [int]$params.tableIndex } else { 1 }
    try {
        if ($doc.Tables.Count -lt $tableIdx -or $tableIdx -lt 1) {
            return @{ ok = $false; error = "Table index $tableIdx out of bounds (document has $($doc.Tables.Count) tables)" }
        }
        $tbl = $doc.Tables.Item($tableIdx)
        
        if ($params.headerBgColor) {
            $hBg = ConvertTo-WordColor $params.headerBgColor
            if ($null -ne $hBg) {
                $tbl.Rows.Item(1).Shading.BackgroundPatternColor = $hBg
                $tbl.Rows.Item(1).Range.Font.Color = 16777215
                $tbl.Rows.Item(1).Range.Font.Bold = 1
            }
        }
        if ($params.headerColor) {
            $hFg = ConvertTo-WordColor $params.headerColor
            if ($null -ne $hFg) { $tbl.Rows.Item(1).Range.Font.Color = $hFg }
        }
        
        if ($params.zebra -and $tbl.Rows.Count -gt 2) {
            $zCol = if ($params.zebraColor) { ConvertTo-WordColor $params.zebraColor } else { 16119285 }
            for ($r = 2; $r -le $tbl.Rows.Count; $r += 2) {
                if ($null -ne $zCol) { $tbl.Rows.Item($r).Shading.BackgroundPatternColor = $zCol }
            }
        }
        
        if ($params.alignment) {
            $tbl.Rows.Alignment = ConvertTo-WordAlignment $params.alignment
        }
        
        if ($params.columnWidths -and $params.columnWidths.Count -gt 0) {
            for ($c = 1; $c -le [Math]::Min($tbl.Columns.Count, $params.columnWidths.Count); $c++) {
                try { $tbl.Columns.Item($c).Width = [float]$params.columnWidths[$c - 1] } catch {}
            }
        }
        
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{
            ok = $true
            path = $doc.FullName
            tableIndex = $tableIdx
            rows = $tbl.Rows.Count
            columns = $tbl.Columns.Count
        }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordComment($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application"
    if ($null -eq $word) { return @{ ok = $false; error = "Word is not running" } }
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) {
        if ($word.Documents.Count -gt 0) { $doc = $word.ActiveDocument }
        else { return @{ ok = $false; error = "No open Word document found" } }
    }
    
    $commentText = if ($params.text) { [string]$params.text } else { [string]$params.comment }
    if (-not $commentText) { return @{ ok = $false; error = "Parameter 'text' or 'comment' is required" } }
    
    try {
        $rng = $doc.Content
        if ($params.search) {
            $find = $rng.Find
            $find.ClearFormatting()
            $find.Text = [string]$params.search
            if (-not $find.Execute()) {
                return @{ ok = $false; error = "Target search text not found: $($params.search)" }
            }
        }
        $comment = $doc.Comments.Add($rng, $commentText)
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; comment = $commentText; targetText = $rng.Text.Trim() }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-WordHyperlink($params) {
    $srcPath = Normalize-DocPath $params.path
    $word = Get-RunningApp "Word.Application"
    if ($null -eq $word) { return @{ ok = $false; error = "Word is not running" } }
    $doc = Find-DocInCollection $word $srcPath "word"
    if ($null -eq $doc) {
        if ($word.Documents.Count -gt 0) { $doc = $word.ActiveDocument }
        else { return @{ ok = $false; error = "No open Word document found" } }
    }
    
    $address = [string]$params.address
    if (-not $address) { return @{ ok = $false; error = "Parameter 'address' (URL) is required" } }
    
    try {
        $rng = $doc.Content
        if ($params.search) {
            $find = $rng.Find
            $find.ClearFormatting()
            $find.Text = [string]$params.search
            if (-not $find.Execute()) {
                return @{ ok = $false; error = "Target text for hyperlink not found: $($params.search)" }
            }
        }
        $screenTip = if ($params.screenTip) { [string]$params.screenTip } else { "" }
        $linkText = if ($params.text) { [string]$params.text } else { $rng.Text }
        $hl = $doc.Hyperlinks.Add($rng, $address, [Type]::Missing, $screenTip, $linkText)
        if ($params.save -and $doc.Path -ne "") { try { $doc.Save() } catch {} }
        return @{ ok = $true; path = $doc.FullName; address = $address; text = $linkText }
    } finally {
        Release-Com $doc
        Release-Com $word
        Invoke-GarbageCollection
    }
}

function Invoke-ExcelTable($params) {
    $srcPath = Normalize-DocPath $params.path
    $excel = Get-RunningApp "Excel.Application"
    $wasRunning = ($null -ne $excel)
    if (-not $wasRunning) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $wb = if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) { $excel.Workbooks.Open($srcPath) } else { $excel.Workbooks.Add() }
    } else {
        $wb = Get-ExcelWorkbook $excel $srcPath
        if ($null -eq $wb) { $wb = $excel.Workbooks.Add() }
    }
    
    try {
        $sheet = if ($params.sheet) { $wb.Sheets.Item($params.sheet) } else { $wb.ActiveSheet }
        $rangeStr = if ($params.range) { $params.range } else { $sheet.UsedRange.Address }
        $rng = $sheet.Range($rangeStr)
        
        $styleName = if ($params.style) { [string]$params.style } else { "TableStyleMedium9" }
        $tableName = if ($params.name) { [string]$params.name } else { "Table_" + [Guid]::NewGuid().ToString("N").Substring(0, 6) }
        $hasHeaders = if ($params.hasHeaders -ne $false) { 1 } else { 2 }
        
        $tableObj = $sheet.ListObjects.Add(1, $rng, [Type]::Missing, $hasHeaders)
        try { $tableObj.Name = $tableName } catch {}
        try { $tableObj.TableStyle = $styleName } catch {}
        if ($params.showTotals) {
            try { $tableObj.ShowTotals = $true } catch {}
        }
        
        if ($params.save -and $wb.Path -ne "") { try { $wb.Save() } catch {} }
        return @{
            ok = $true
            path = $wb.FullName
            sheet = $sheet.Name
            tableName = $tableObj.Name
            range = $tableObj.Range.Address
            style = $styleName
        }
    } finally {
        Release-Com $wb
        Release-Com $excel
        Invoke-GarbageCollection
    }
}

function Invoke-ExcelChart($params) {
    $srcPath = Normalize-DocPath $params.path
    $excel = Get-RunningApp "Excel.Application"
    $wasRunning = ($null -ne $excel)
    if (-not $wasRunning) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $wb = if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) { $excel.Workbooks.Open($srcPath) } else { $excel.Workbooks.Add() }
    } else {
        $wb = Get-ExcelWorkbook $excel $srcPath
        if ($null -eq $wb) { $wb = $excel.Workbooks.Add() }
    }
    
    try {
        $sheet = if ($params.sheet) { $wb.Sheets.Item($params.sheet) } else { $wb.ActiveSheet }
        $sourceRangeStr = if ($params.sourceRange) { $params.sourceRange } else { $sheet.UsedRange.Address }
        $sourceRng = $sheet.Range($sourceRangeStr)
        
        $chartTypeStr = if ($params.chartType) { $params.chartType.ToLower() } else { "column" }
        $xlChartType = switch ($chartTypeStr) {
            "line" { 4 }
            "pie"  { 5 }
            "bar"  { 57 }
            "area" { 1 }
            default { 51 }
        }
        
        $left = if ($params.left) { [float]$params.left } else { 350.0 }
        $top = if ($params.top) { [float]$params.top } else { 50.0 }
        $width = if ($params.width) { [float]$params.width } else { 450.0 }
        $height = if ($params.height) { [float]$params.height } else { 280.0 }
        
        $chartShape = $sheet.Shapes.AddChart2(201, $xlChartType, $left, $top, $width, $height)
        $chart = $chartShape.Chart
        $chart.SetSourceData($sourceRng)
        
        if ($params.title) {
            $chart.HasTitle = $true
            $chart.ChartTitle.Text = [string]$params.title
        }
        if ($null -ne $params.legend) {
            $chart.HasLegend = [bool]$params.legend
        }
        
        if ($params.save -and $wb.Path -ne "") { try { $wb.Save() } catch {} }
        return @{
            ok = $true
            path = $wb.FullName
            sheet = $sheet.Name
            chartType = $chartTypeStr
            title = if ($chart.HasTitle) { $chart.ChartTitle.Text } else { "" }
            left = $left
            top = $top
            width = $width
            height = $height
        }
    } finally {
        Release-Com $wb
        Release-Com $excel
        Invoke-GarbageCollection
    }
}

function Invoke-ExcelFormatRange($params) {
    $srcPath = Normalize-DocPath $params.path
    $excel = Get-RunningApp "Excel.Application"
    $wasRunning = ($null -ne $excel)
    if (-not $wasRunning) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $wb = if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) { $excel.Workbooks.Open($srcPath) } else { $excel.Workbooks.Add() }
    } else {
        $wb = Get-ExcelWorkbook $excel $srcPath
        if ($null -eq $wb) { $wb = $excel.Workbooks.Add() }
    }
    
    try {
        $sheet = if ($params.sheet) { $wb.Sheets.Item($params.sheet) } else { $wb.ActiveSheet }
        $rangeStr = if ($params.range) { $params.range } else { $sheet.UsedRange.Address }
        $rng = $sheet.Range($rangeStr)
        
        if ($null -ne $params.bold) { $rng.Font.Bold = [bool]$params.bold }
        if ($null -ne $params.italic) { $rng.Font.Italic = [bool]$params.italic }
        if ($params.fontSize) { $rng.Font.Size = [float]$params.fontSize }
        if ($params.fontName) { $rng.Font.Name = [string]$params.fontName }
        if ($params.color) {
            $c = ConvertTo-WordColor $params.color
            if ($null -ne $c) { $rng.Font.Color = $c }
        }
        if ($params.bgColor -or $params.fillColor) {
            $bg = if ($params.bgColor) { $params.bgColor } else { $params.fillColor }
            $c = ConvertTo-WordColor $bg
            if ($null -ne $c) { $rng.Interior.Color = $c }
        }
        
        if ($params.alignment) {
            $a = $params.alignment.ToString().ToLower()
            $hAlign = switch ($a) {
                "center" { -4108 }
                "centrado" { -4108 }
                "right" { -4152 }
                "derecha" { -4152 }
                default { -4131 }
            }
            $rng.HorizontalAlignment = $hAlign
        }
        if ($params.verticalAlignment) {
            $va = $params.verticalAlignment.ToString().ToLower()
            $vAlign = switch ($va) {
                "top" { -4160 }
                "bottom" { -4107 }
                default { -4108 }
            }
            $rng.VerticalAlignment = $vAlign
        }
        if ($null -ne $params.wrapText) {
            $rng.WrapText = [bool]$params.wrapText
        }
        
        if ($params.numberFormat) {
            $nf = $params.numberFormat.ToString().ToLower()
            $formatStr = switch ($nf) {
                "currency"   { "$#,##0.00" }
                "moneda"     { "$#,##0.00" }
                "percentage" { "0.0%" }
                "porcentaje" { "0.0%" }
                "date"       { "yyyy-mm-dd" }
                "fecha"      { "yyyy-mm-dd" }
                "integer"    { "#,##0" }
                "entero"     { "#,##0" }
                default      { [string]$params.numberFormat }
            }
            $rng.NumberFormat = $formatStr
        }
        
        if ($params.borders) {
            $rng.Borders.LineStyle = 1
        }
        
        if ($params.merge -eq $true) {
            $rng.Merge()
        } elseif ($params.merge -eq $false -or $params.unmerge -eq $true) {
            $rng.UnMerge()
        }
        
        if ($params.rowHeight) { $rng.RowHeight = [float]$params.rowHeight }
        if ($params.columnWidth) { $rng.ColumnWidth = [float]$params.columnWidth }
        if ($params.autoFit) { $rng.Columns.AutoFit() }
        
        if ($params.save -and $wb.Path -ne "") { try { $wb.Save() } catch {} }
        return @{
            ok = $true
            path = $wb.FullName
            sheet = $sheet.Name
            range = $rng.Address
            rows = $rng.Rows.Count
            columns = $rng.Columns.Count
        }
    } finally {
        Release-Com $wb
        Release-Com $excel
        Invoke-GarbageCollection
    }
}

function Invoke-ExcelSheetOps($params) {
    $srcPath = Normalize-DocPath $params.path
    $excel = Get-RunningApp "Excel.Application"
    if ($null -eq $excel) { return @{ ok = $false; error = "Excel is not running" } }
    $wb = Get-ExcelWorkbook $excel $srcPath
    if ($null -eq $wb) { return @{ ok = $false; error = "No open Excel workbook found" } }
    
    $op = if ($params.action) { $params.action.ToLower() } else { "rename" }
    $sheetTarget = if ($params.sheet) { $wb.Sheets.Item($params.sheet) } else { $wb.ActiveSheet }
    
    try {
        switch ($op) {
            "rename" {
                if (-not $params.newName) { return @{ ok = $false; error = "Parameter 'newName' required for rename" } }
                $sheetTarget.Name = [string]$params.newName
            }
            "delete" {
                if ($wb.Sheets.Count -le 1) { return @{ ok = $false; error = "Cannot delete the only worksheet in workbook" } }
                $excel.DisplayAlerts = $false
                $sheetTarget.Delete()
                $excel.DisplayAlerts = $true
            }
            "freeze" {
                $sheetTarget.Activate()
                $freezeCell = if ($params.freezeAt) { [string]$params.freezeAt } else { "B2" }
                $sheetTarget.Range($freezeCell).Select()
                $excel.ActiveWindow.FreezePanes = $true
            }
            "hide" {
                $sheetTarget.Visible = 0
            }
            "unhide" {
                $sheetTarget.Visible = -1
            }
            default {
                return @{ ok = $false; error = "Unknown sheet operation: $op" }
            }
        }
        
        if ($params.save -and $wb.Path -ne "") { try { $wb.Save() } catch {} }
        return @{
            ok = $true
            path = $wb.FullName
            operation = $op
            sheetName = $sheetTarget.Name
            totalSheets = $wb.Sheets.Count
        }
    } finally {
        Release-Com $wb
        Release-Com $excel
        Invoke-GarbageCollection
    }
}

function Invoke-ExcelCalculate($params) {
    $excel = Get-RunningApp "Excel.Application"
    if ($null -eq $excel) { return @{ ok = $false; error = "Excel is not running" } }
    
    try {
        $excel.Calculate()
        $wb = $excel.ActiveWorkbook
        $errorsFound = @()
        if ($null -ne $wb) {
            foreach ($sh in $wb.Sheets) {
                try {
                    $formulas = $sh.UsedRange.SpecialCells(-4123)
                    if ($null -ne $formulas) {
                        foreach ($cell in $formulas) {
                            $valStr = [string]$cell.Text
                            if ($valStr -match "#REF!|#VALUE!|#NAME\?|#DIV/0!|#N/A") {
                                $errorsFound += @{ sheet = $sh.Name; cell = $cell.Address; error = $valStr; formula = $cell.Formula }
                            }
                        }
                    }
                } catch {}
            }
        }
        return @{
            ok = $true
            calculated = $true
            errorsCount = $errorsFound.Count
            errors = $errorsFound
        }
    } finally {
        Release-Com $excel
        Invoke-GarbageCollection
    }
}

function Invoke-PptSlideOps($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -eq $ppt) { return @{ ok = $false; error = "PowerPoint is not running" } }
    $pres = Get-PptPresentation $ppt $srcPath
    if ($null -eq $pres) { return @{ ok = $false; error = "No open PowerPoint presentation found" } }
    
    $op = if ($params.action) { $params.action.ToLower() } else { "duplicate" }
    
    try {
        switch ($op) {
            "duplicate" {
                $idx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
                $dup = $pres.Slides.Item($idx).Duplicate()
                return @{ ok = $true; operation = "duplicate"; newSlideNumber = $dup.Item(1).SlideNumber; totalSlides = $pres.Slides.Count }
            }
            "move" {
                $from = [int]$params.fromIndex
                $to = [int]$params.toIndex
                $pres.Slides.Item($from).MoveTo($to)
                return @{ ok = $true; operation = "move"; from = $from; to = $to; totalSlides = $pres.Slides.Count }
            }
            "set_size" {
                $ratio = if ($params.aspectRatio) { $params.aspectRatio.ToString().ToLower() } else { "16:9" }
                if ($ratio -match "16:9|widescreen") {
                    $pres.PageSetup.SlideWidth = 960.0
                    $pres.PageSetup.SlideHeight = 540.0
                } else {
                    $pres.PageSetup.SlideWidth = 720.0
                    $pres.PageSetup.SlideHeight = 540.0
                }
                return @{ ok = $true; operation = "set_size"; width = $pres.PageSetup.SlideWidth; height = $pres.PageSetup.SlideHeight; ratio = $ratio }
            }
            default {
                return @{ ok = $false; error = "Unknown slide operation: $op" }
            }
        }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

function Invoke-PptNotes($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -eq $ppt) { return @{ ok = $false; error = "PowerPoint is not running" } }
    $pres = Get-PptPresentation $ppt $srcPath
    if ($null -eq $pres) { return @{ ok = $false; error = "No open PowerPoint presentation found" } }
    
    $idx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
    try {
        $slide = $pres.Slides.Item($idx)
        $notesShape = $slide.NotesPage.Shapes.Placeholders.Item(2)
        if ($params.notes -or $params.text) {
            $notesText = if ($params.notes) { [string]$params.notes } else { [string]$params.text }
            $notesShape.TextFrame.TextRange.Text = $notesText
            if ($params.save -and $pres.Path -ne "") { try { $pres.Save() } catch {} }
            return @{ ok = $true; slideIndex = $idx; notes = $notesText; updated = $true }
        } else {
            $existing = $notesShape.TextFrame.TextRange.Text.Trim()
            return @{ ok = $true; slideIndex = $idx; notes = $existing }
        }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

function Invoke-PptExportSlide($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -eq $ppt) { return @{ ok = $false; error = "PowerPoint is not running" } }
    $pres = Get-PptPresentation $ppt $srcPath
    if ($null -eq $pres) { return @{ ok = $false; error = "No open PowerPoint presentation found" } }
    
    $idx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
    $savePath = if ($params.outputPath) { $params.outputPath } else { "$env:TEMP\ppt_slide_$idx" + "_$([Guid]::NewGuid().ToString('N')).png" }
    $w = if ($params.width) { [int]$params.width } else { 1920 }
    $h = if ($params.height) { [int]$params.height } else { 1080 }
    
    try {
        $slide = $pres.Slides.Item($idx)
        $slide.Export($savePath, "PNG", $w, $h)
        return @{
            ok = $true
            slideIndex = $idx
            outputPath = $savePath
            width = $w
            height = $h
            sizeBytes = (Get-Item $savePath).Length
        }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

function Invoke-PptBackground($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -eq $ppt) { return @{ ok = $false; error = "PowerPoint is not running" } }
    $pres = Get-PptPresentation $ppt $srcPath
    if ($null -eq $pres) { return @{ ok = $false; error = "No open PowerPoint presentation found" } }
    
    $color = if ($params.color) { $params.color } else { "#FFFFFF" }
    $cVal = ConvertTo-WordColor $color
    $idx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
    $all = [bool]$params.all
    
    try {
        if ($all) {
            foreach ($s in $pres.Slides) {
                $s.FollowMasterBackground = 0
                $s.Background.Fill.Solid()
                $s.Background.Fill.ForeColor.RGB = $cVal
            }
        } else {
            $slide = $pres.Slides.Item($idx)
            $slide.FollowMasterBackground = 0
            $slide.Background.Fill.Solid()
            $slide.Background.Fill.ForeColor.RGB = $cVal
        }
        if ($params.save -and $pres.Path -ne "") { try { $pres.Save() } catch {} }
        return @{ ok = $true; color = $color; slideIndex = if ($all) { "all" } else { $idx } }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

function Invoke-PptChart($params) {
    $srcPath = Normalize-DocPath $params.path
    $ppt = Get-RunningApp "PowerPoint.Application"
    if ($null -eq $ppt) {
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = 1
        $pres = if ($srcPath -and (Test-Path $srcPath -ErrorAction SilentlyContinue)) { $ppt.Presentations.Open($srcPath) } else { $ppt.Presentations.Add(1) }
    } else {
        $pres = Get-PptPresentation $ppt $srcPath
        if ($null -eq $pres) { $pres = $ppt.Presentations.Add(1) }
    }
    
    $idx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
    if ($idx -gt $pres.Slides.Count) {
        $pres.Slides.Add($pres.Slides.Count + 1, 12) | Out-Null
        $idx = $pres.Slides.Count
    }
    $slide = $pres.Slides.Item($idx)
    
    $chartTypeStr = if ($params.chartType) { $params.chartType.ToLower() } else { "column" }
    $typeInt = switch ($chartTypeStr) {
        "line" { 4 }
        "pie"  { 5 }
        "bar"  { 57 }
        default { 51 }
    }
    
    $left = if ($params.left) { [float]$params.left } else { 100.0 }
    $top = if ($params.top) { [float]$params.top } else { 100.0 }
    $width = if ($params.width) { [float]$params.width } else { 600.0 }
    $height = if ($params.height) { [float]$params.height } else { 350.0 }
    
    try {
        $chartShape = $slide.Shapes.AddChart2(-1, $typeInt, $left, $top, $width, $height)
        if ($params.title) {
            $chartShape.Chart.HasTitle = $true
            $chartShape.Chart.ChartTitle.Text = [string]$params.title
        }
        if ($params.save -and $pres.Path -ne "") { try { $pres.Save() } catch {} }
        return @{
            ok = $true
            app = "powerpoint"
            slideIndex = $idx
            chartType = $chartTypeStr
            left = $left
            top = $top
            width = $width
            height = $height
        }
    } finally {
        Release-Com $pres
        Release-Com $ppt
        Invoke-GarbageCollection
    }
}

# --- Main Entry Point ---

try {
    $rawPayload = Get-Content -Path $PayloadFile -Raw -Encoding UTF8
    $request = $rawPayload | ConvertFrom-Json
    
    $action = $request.action
    $params = $request.params
    
    $result = Invoke-ComWithRetry {
        switch ($action) {
            "inspect_structure"   { Invoke-InspectStructure $params }
            "word_table_format"   { Invoke-WordTableFormat $params }
            "word_comment"        { Invoke-WordComment $params }
            "word_hyperlink"      { Invoke-WordHyperlink $params }
            "excel_table"         { Invoke-ExcelTable $params }
            "excel_chart"         { Invoke-ExcelChart $params }
            "excel_format_range"  { Invoke-ExcelFormatRange $params }
            "excel_sheet_ops"     { Invoke-ExcelSheetOps $params }
            "excel_calculate"     { Invoke-ExcelCalculate $params }
            "ppt_slide_ops"       { Invoke-PptSlideOps $params }
            "ppt_notes"           { Invoke-PptNotes $params }
            "ppt_export_slide"    { Invoke-PptExportSlide $params }
            "ppt_background"      { Invoke-PptBackground $params }
            "ppt_chart"           { Invoke-PptChart $params }
            "status"              { Invoke-Status $params }
            "list_open"           { Invoke-ListOpen $params }
            "get_active"          { Invoke-GetActive $params }
            "create"              { Invoke-Create $params }
            "open"                { Invoke-Open $params }
            "save"                { Invoke-Save $params }
            "save_as"             { Invoke-SaveAs $params }
            "close"               { Invoke-Close $params }
            "export_pdf"          { Invoke-ExportPdf $params }
            "read"                { Invoke-WordRead $params }
            "insert_text"         { Invoke-WordInsertText $params }
            "replace_text"        { Invoke-WordReplaceText $params }
            "clear_document"      { Invoke-WordClear $params }
            "append_section"      { Invoke-WordAppendSection $params }
            "insert_table"        {
                $app = if ($params.app) { $params.app.ToLower() } else { "word" }
                if ($app -eq "powerpoint" -or $app -eq "ppt") {
                    Invoke-PptInsertTable $params
                } else {
                    Invoke-WordInsertTable $params
                }
            }
            "insert_shape"        {
                $app = if ($params.app) { $params.app.ToLower() } else { "word" }
                if ($app -eq "powerpoint" -or $app -eq "ppt") {
                    Invoke-PptInsertShape $params
                } else {
                    Invoke-InsertShape $params
                }
            }
            "insert_image"        { Invoke-InsertImage $params }
            "read_slides"         { Invoke-PptReadSlides $params }
            "add_slide"           { Invoke-PptAddSlide $params }
            "set_slide_text"      { Invoke-PptSetSlideText $params }
            "delete_slide"        { Invoke-PptDeleteSlide $params }
            "read_range"          { Invoke-ExcelReadRange $params }
            "write_range"         { Invoke-ExcelWriteRange $params }
            "write_cell"          { Invoke-ExcelWriteCell $params }
            "add_sheet"           { Invoke-ExcelAddSheet $params }
            "format_text"         { Invoke-WordFormatText $params }
            "format_paragraphs"   { Invoke-WordFormatParagraphs $params }
            "page_setup"          { Invoke-WordPageSetup $params }
            "add_page_numbers"    { Invoke-WordAddPageNumbers $params }
            "insert_page_break"   { Invoke-WordInsertPageBreak $params }
            "apply_style"         { Invoke-WordApplyStyle $params }
            "insert_list"         { Invoke-WordInsertList $params }
            "highlight_text"      { Invoke-WordHighlightText $params }
            "insert_callout"      { Invoke-WordInsertCallout $params }
            "insert_divider"      { Invoke-WordInsertDivider $params }
            "insert_header_footer" { Invoke-WordInsertHeaderFooter $params }
            "insert_cover_page"   { Invoke-WordInsertCoverPage $params }
            "set_page_background" { Invoke-WordSetPageBackground $params }
            "insert_section_break" { Invoke-WordInsertSectionBreak $params }
            "insert_table_of_contents" { Invoke-WordInsertTOC $params }
            "update_toc"          { Invoke-WordUpdateTOC $params }
            "insert_footnote"     { Invoke-WordInsertFootnote $params }
            "replace_placeholders" { Invoke-WordReplacePlaceholders $params }
            "get_document_stats"  { Invoke-WordGetStats $params }
            "set_language"        { Invoke-WordSetLanguage $params }
            "set_view"            { Invoke-WordSetView $params }
            { $_ -in @("capture_preview", "screen_capture") } {
                $appType = if ($params.app) { $params.app.ToLower() } else { "word" }
                
                # Smart Page / Slide / Sheet Navigation before capturing
                if ($appType -eq "word") {
                    $wordApp = Get-RunningApp "Word.Application"
                    if ($null -ne $wordApp -and $wordApp.Documents.Count -gt 0) {
                        try {
                            $wordApp.Visible = $true
                            $wordApp.Activate()
                            if ($params.page) {
                                # 1 = wdGoToPage, 1 = wdGoToAbsolute
                                $rng = $wordApp.Selection.GoTo(1, 1, [int]$params.page)
                                if ($null -ne $rng) {
                                    $rng.Select()
                                    try { $wordApp.ActiveWindow.ScrollIntoView($rng, $true) } catch {}
                                }
                                Start-Sleep -Milliseconds 400
                            } elseif ($params.scrollToEnd) {
                                # 6 = wdStory, 0 = wdMove
                                $wordApp.Selection.EndKey(6, 0) | Out-Null
                                try { $wordApp.ActiveWindow.ScrollIntoView($wordApp.Selection.Range, $true) } catch {}
                                Start-Sleep -Milliseconds 400
                            }
                        } catch {}
                    }
                } elseif ($appType -eq "powerpoint") {
                    $pptApp = Get-RunningApp "PowerPoint.Application"
                    if ($null -ne $pptApp -and $pptApp.Presentations.Count -gt 0) {
                        try {
                            $slideNum = if ($params.slide) { [int]$params.slide } elseif ($params.slideIndex) { [int]$params.slideIndex } else { $null }
                            if ($null -ne $slideNum) {
                                $pptApp.ActiveWindow.View.GotoSlide($slideNum) | Out-Null
                                Start-Sleep -Milliseconds 250
                            }
                        } catch {}
                    }
                } elseif ($appType -eq "excel") {
                    $xlApp = Get-RunningApp "Excel.Application"
                    if ($null -ne $xlApp -and $xlApp.Workbooks.Count -gt 0) {
                        try {
                            if ($params.sheet) {
                                $xlApp.ActiveWorkbook.Sheets.Item($params.sheet).Activate() | Out-Null
                                Start-Sleep -Milliseconds 250
                            }
                        } catch {}
                    }
                }

                $targetClass = switch ($appType) {
                    "powerpoint" { "PPTFrameClass" }
                    "excel"      { "XLMAIN" }
                    default      { "OpusApp" }
                }
                $savePath = if ($params.outputPath) { $params.outputPath } else { "$env:TEMP\office_preview_$([Guid]::NewGuid().ToString('N')).png" }
                $capRes = [LiveAccessibleHelper]::CaptureWindow($targetClass, $savePath)
                if ($capRes -eq "OK") {
                    @{ ok = $true; app = $appType; previewPath = $savePath; sizeBytes = (Get-Item $savePath).Length; page = $params.page }
                } else {
                    @{ ok = $false; error = $capRes }
                }
            }
            default {
                @{ ok = $false; error = "Unrecognized document action: $action" }
            }
        }
    }
    
    $json = ConvertTo-CleanJson $result
    Write-Output $json
} catch {
    $errObj = @{
        ok = $false
        error = $_.Exception.Message
        stack = $_.ScriptStackTrace
    }
    Write-Output (ConvertTo-CleanJson $errObj)
}

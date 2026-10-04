<#
.SYNOPSIS
    FLUXER XZ — tools/documents/office-com.ps1
    Windows Office COM Automation Engine (Word, PowerPoint, Excel).
    Executes in LIVE mode (GetActiveObject) or HEADLESS mode (New-Object).
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$PayloadFile
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

function ConvertTo-CleanJson($obj) {
    return ($obj | ConvertTo-Json -Depth 10 -Compress)
}

function Get-RunningApp($progId, $docPath = "") {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        if ($null -ne $app) { return $app }
    } catch {}

    if ($docPath -and (Test-Path $docPath)) {
        try {
            $bound = [System.Runtime.InteropServices.Marshal]::BindToMoniker($docPath)
            if ($null -ne $bound) {
                if ($bound.Application) { return $bound.Application }
                return $bound
            }
        } catch {}
    }

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
    return [System.IO.Path]::GetFullPath($p)
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
                    Release-Com $doc
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
                    Release-Com $pres
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
                    Release-Com $wb
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
            $doc = if ($params.template) { $word.Documents.Add($params.template) } else { $word.Documents.Add() }
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
            return @{ ok = $true; app = "word"; path = $targetPath; title = $title; liveMode = $wasRunning }
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
            # In PPT Visible is an MsoTriState
            $ppt.Visible = 1 # msoTrue
            $pres = if ($params.template) { $ppt.Presentations.Open($params.template) } else { $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoTrue) }
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
            return @{ ok = $true; app = "powerpoint"; path = $targetPath; title = $title; liveMode = $wasRunning }
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
            $wb = $excel.Workbooks.Add()
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
            return @{ ok = $true; app = "excel"; path = $targetPath; title = $title; liveMode = $wasRunning }
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
    
    # Try Word
    if (-not $appType -or $appType -eq "word") {
        $word = Get-RunningApp "Word.Application"
        if ($null -ne $word) {
            try {
                foreach ($doc in $word.Documents) {
                    if (-not $targetPath -or (Normalize-DocPath $doc.FullName) -eq $targetPath -or $doc.Name -eq $targetPath) {
                        $doc.Save()
                        $res = @{ ok = $true; app = "word"; path = $doc.FullName; title = $doc.Name; saved = $true }
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
    }
    
    # Try PowerPoint
    if (-not $appType -or $appType -eq "powerpoint") {
        $ppt = Get-RunningApp "PowerPoint.Application"
        if ($null -ne $ppt) {
            try {
                foreach ($pres in $ppt.Presentations) {
                    if (-not $targetPath -or (Normalize-DocPath $pres.FullName) -eq $targetPath -or $pres.Name -eq $targetPath) {
                        $pres.Save()
                        $res = @{ ok = $true; app = "powerpoint"; path = $pres.FullName; title = $pres.Name; saved = $true }
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
    }
    
    # Try Excel
    if (-not $appType -or $appType -eq "excel") {
        $excel = Get-RunningApp "Excel.Application"
        if ($null -ne $excel) {
            try {
                foreach ($wb in $excel.Workbooks) {
                    if (-not $targetPath -or (Normalize-DocPath $wb.FullName) -eq $targetPath -or $wb.Name -eq $targetPath) {
                        $wb.Save()
                        $res = @{ ok = $true; app = "excel"; path = $wb.FullName; title = $wb.Name; saved = $true }
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
            $pres = $ppt.Presentations.Open($srcPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
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
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    $maxChars = if ($params.maxChars) { [int]$params.maxChars } else { 4000 }
    
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
        
        return @{
            ok = $true
            path = $srcPath
            title = $doc.Name
            paragraphs = $totalParagraphs
            headings = $headings
            tables = $tablesCount
            comments = $commentsCount
            textSnippet = $sampleText.Trim()
            truncated = ($charCount -ge $maxChars)
        }
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

function Invoke-WordInsertText($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    if (-not $params.text) { return @{ ok = $false; error = "Text to insert is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    
    $word = Get-RunningApp "Word.Application"
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) { $word = New-Object -ComObject Word.Application; $word.Visible = $false }
    
    $doc = $null
    foreach ($d in $word.Documents) {
        if ((Normalize-DocPath $d.FullName) -eq $srcPath) { $doc = $d; break }
    }
    $openedHere = $false
    if ($null -eq $doc) {
        $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
        $openedHere = $true
    }
    
    try {
        $pos = if ($params.position) { $params.position.ToLower() } else { "end" }
        $range = $null
        if ($pos -eq "start") {
            $range = $doc.Range(0, 0)
        } else {
            $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        }
        
        $range.Text = "`n" + $params.text
        if ($null -ne $params.bold) { $range.Bold = [int][bool]$params.bold }
        if ($null -ne $params.italic) { $range.Italic = [int][bool]$params.italic }
        if ($null -ne $params.fontSize) { $range.Font.Size = [float]$params.fontSize }
        
        $doc.Save()
        return @{
            ok = $true
            path = $srcPath
            insertedLength = $params.text.Length
            position = $pos
            totalParagraphs = $doc.Paragraphs.Count
        }
    } finally {
        if ($openedHere) {
            try { $doc.Close([ref]-1) } catch {}
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

function Invoke-WordReplaceText($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    if (-not $params.search) { return @{ ok = $false; error = "Search text is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    
    $word = Get-RunningApp "Word.Application"
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) { $word = New-Object -ComObject Word.Application; $word.Visible = $false }
    
    $doc = $null
    foreach ($d in $word.Documents) {
        if ((Normalize-DocPath $d.FullName) -eq $srcPath) { $doc = $d; break }
    }
    $openedHere = $false
    if ($null -eq $doc) {
        $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
        $openedHere = $true
    }
    
    try {
        $search = $params.search
        $replace = if ($params.replace) { $params.replace } else { "" }
        $matchCase = if ($null -ne $params.matchCase) { [bool]$params.matchCase } else { $false }
        $matchWholeWord = if ($null -ne $params.matchWholeWord) { [bool]$params.matchWholeWord } else { $false }
        
        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $search
        $find.Replacement.Text = $replace
        
        # 2 = wdReplaceAll, 1 = wdFindContinue
        $replaced = $find.Execute($search, [bool]$matchCase, [bool]$matchWholeWord, $false, $false, $false, $true, 1, $false, $replace, 2)
        
        $doc.Save()
        return @{
            ok = $true
            path = $srcPath
            search = $search
            replace = $replace
            success = $replaced
        }
    } finally {
        if ($openedHere) {
            try { $doc.Close([ref]-1) } catch {}
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

function Invoke-WordAppendSection($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    if (-not $params.heading) { return @{ ok = $false; error = "Heading is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    
    $word = Get-RunningApp "Word.Application"
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) { $word = New-Object -ComObject Word.Application; $word.Visible = $false }
    
    $doc = $null
    foreach ($d in $word.Documents) {
        if ((Normalize-DocPath $d.FullName) -eq $srcPath) { $doc = $d; break }
    }
    $openedHere = $false
    if ($null -eq $doc) {
        $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
        $openedHere = $true
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
        $pHead.Range.InsertParagraphAfter()
        
        # Insert Body paragraph if present
        if ($params.body) {
            $pBody = $doc.Paragraphs.Add()
            $pBody.Range.Text = $params.body
            try { $pBody.Format.Style = -1 } catch {
                try { $pBody.Format.Style = "Normal" } catch {}
            }
            $pBody.Range.InsertParagraphAfter()
        }
        
        $doc.Save()
        return @{
            ok = $true
            path = $srcPath
            heading = $params.heading
            level = $level
            totalParagraphs = $doc.Paragraphs.Count
        }
    } finally {
        if ($openedHere) {
            try { $doc.Close([ref]-1) } catch {}
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

function Invoke-WordInsertTable($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    
    $rows = if ($params.rows) { [int]$params.rows } else { 2 }
    $cols = if ($params.cols) { [int]$params.cols } else { 2 }
    $data = $params.data
    
    $word = Get-RunningApp "Word.Application"
    $wasRunning = ($null -ne $word)
    if (-not $wasRunning) { $word = New-Object -ComObject Word.Application; $word.Visible = $false }
    
    $doc = $null
    foreach ($d in $word.Documents) {
        if ((Normalize-DocPath $d.FullName) -eq $srcPath) { $doc = $d; break }
    }
    $openedHere = $false
    if ($null -eq $doc) {
        $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
        $openedHere = $true
    }
    
    try {
        $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        $table = $doc.Tables.Add($range, $rows, $cols)
        
        if ($data -and $data.Count -gt 0) {
            for ($r = 0; $r -lt [Math]::Min($rows, $data.Count); $r++) {
                $rowArr = $data[$r]
                for ($c = 0; $c -lt [Math]::Min($cols, $rowArr.Count); $c++) {
                    $table.Cell($r + 1, $c + 1).Range.Text = [string]$rowArr[$c]
                }
            }
        }
        
        $doc.Save()
        return @{ ok = $true; path = $srcPath; tablesCount = $doc.Tables.Count; rows = $rows; cols = $cols }
    } finally {
        if ($openedHere) {
            try { $doc.Close([ref]-1) } catch {}
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

function Invoke-InsertImage($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    if (-not $params.imagePath) { return @{ ok = $false; error = "imagePath is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "Document file not found: $srcPath" } }
    $imgPath = Normalize-DocPath $params.imagePath
    if (-not (Test-Path $imgPath)) { return @{ ok = $false; error = "Image file not found: $imgPath" } }

    $ext = [System.IO.Path]::GetExtension($srcPath).ToLower()

    # Word
    if ($ext -in @(".doc", ".docx")) {
        $word = Get-RunningApp "Word.Application"
        $wasRunning = ($null -ne $word)
        if (-not $wasRunning) { $word = New-Object -ComObject Word.Application; $word.Visible = $false }
        
        $doc = $null
        foreach ($d in $word.Documents) {
            if ((Normalize-DocPath $d.FullName) -eq $srcPath) { $doc = $d; break }
        }
        $openedHere = $false
        if ($null -eq $doc) {
            $doc = $word.Documents.Open($srcPath, [ref]$false, [ref]$false)
            $openedHere = $true
        }
        
        try {
            $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
            $shape = $doc.InlineShapes.AddPicture($imgPath, [ref]$false, [ref]$true, $range)
            if ($params.width) { $shape.Width = [float]$params.width }
            if ($params.height) { $shape.Height = [float]$params.height }
            $doc.Save()
            return @{
                ok = $true
                path = $srcPath
                imagePath = $imgPath
                type = "inline_shape"
                width = $shape.Width
                height = $shape.Height
            }
        } finally {
            if ($openedHere) {
                try { $doc.Close([ref]-1) } catch {}
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

    # PowerPoint
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
            $pres = $ppt.Presentations.Open($srcPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)
            $openedHere = $true
        }
        
        try {
            $slideIdx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
            if ($slideIdx -lt 1 -or $slideIdx -gt $pres.Slides.Count) {
                return @{ ok = $false; error = "Slide index $slideIdx out of range (1..$($pres.Slides.Count))" }
            }
            $slide = $pres.Slides.Item($slideIdx)
            $left = if ($params.left) { [float]$params.left } else { 100 }
            $top = if ($params.top) { [float]$params.top } else { 100 }
            $w = if ($params.width) { [float]$params.width } else { 300 }
            $h = if ($params.height) { [float]$params.height } else { 200 }
            
            $shape = $slide.Shapes.AddPicture($imgPath, 0, -1, $left, $top, $w, $h)
            $pres.Save()
            return @{
                ok = $true
                path = $srcPath
                imagePath = $imgPath
                slideIndex = $slideIdx
                shapeName = $shape.Name
                width = $shape.Width
                height = $shape.Height
            }
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

    return @{ ok = $false; error = "Unsupported format for insert_image: $ext" }
}

# --- PowerPoint Actions ---

function Invoke-PptReadSlides($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    
    $ppt = Get-RunningApp "PowerPoint.Application"
    $wasRunning = ($null -ne $ppt)
    if (-not $wasRunning) { $ppt = New-Object -ComObject PowerPoint.Application }
    
    $pres = $null
    foreach ($p in $ppt.Presentations) {
        if ((Normalize-DocPath $p.FullName) -eq $srcPath) { $pres = $p; break }
    }
    $openedHere = $false
    if ($null -eq $pres) {
        $pres = $ppt.Presentations.Open($srcPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
        $openedHere = $true
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
            path = $srcPath
            title = $pres.Name
            slideCount = $pres.Slides.Count
            slides = $slidesData
        }
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

function Invoke-PptAddSlide($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    
    $ppt = Get-RunningApp "PowerPoint.Application"
    $wasRunning = ($null -ne $ppt)
    if (-not $wasRunning) { $ppt = New-Object -ComObject PowerPoint.Application }
    
    $pres = $null
    foreach ($p in $ppt.Presentations) {
        if ((Normalize-DocPath $p.FullName) -eq $srcPath) { $pres = $p; break }
    }
    $openedHere = $false
    if ($null -eq $pres) {
        $pres = $ppt.Presentations.Open($srcPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)
        $openedHere = $true
    }
    
    try {
        $layoutNum = 2 # ppLayoutText (title + content)
        if ($params.layout -eq "title") { $layoutNum = 1 }
        elseif ($params.layout -eq "blank") { $layoutNum = 12 }
        
        $newIndex = $pres.Slides.Count + 1
        $slide = $pres.Slides.Add($newIndex, $layoutNum)
        
        if ($params.title -and $slide.Shapes.HasTitle) {
            $slide.Shapes.Title.TextFrame.TextRange.Text = $params.title
        }
        
        if ($params.body -and $slide.Shapes.Placeholders.Count -ge 2) {
            $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $params.body
        }
        
        $pres.Save()
        return @{
            ok = $true
            path = $srcPath
            slideNumber = $slide.SlideNumber
            totalSlides = $pres.Slides.Count
            title = $params.title
        }
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

function Invoke-PptSetSlideText($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    $slideIdx = if ($params.slideIndex) { [int]$params.slideIndex } else { 1 }
    $shapeIdx = if ($params.shapeIndex) { [int]$params.shapeIndex } else { 1 }
    
    $ppt = Get-RunningApp "PowerPoint.Application"
    $wasRunning = ($null -ne $ppt)
    if (-not $wasRunning) { $ppt = New-Object -ComObject PowerPoint.Application }
    
    $pres = $null
    foreach ($p in $ppt.Presentations) {
        if ((Normalize-DocPath $p.FullName) -eq $srcPath) { $pres = $p; break }
    }
    $openedHere = $false
    if ($null -eq $pres) {
        $pres = $ppt.Presentations.Open($srcPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)
        $openedHere = $true
    }
    
    try {
        if ($slideIdx -gt $pres.Slides.Count -or $slideIdx -lt 1) {
            return @{ ok = $false; error = "Slide index $slideIdx out of range (1..$($pres.Slides.Count))" }
        }
        $slide = $pres.Slides.Item($slideIdx)
        if ($shapeIdx -gt $slide.Shapes.Count -or $shapeIdx -lt 1) {
            return @{ ok = $false; error = "Shape index $shapeIdx out of range" }
        }
        $shape = $slide.Shapes.Item($shapeIdx)
        $shape.TextFrame.TextRange.Text = $params.text
        $pres.Save()
        return @{ ok = $true; path = $srcPath; slideIndex = $slideIdx; shapeIndex = $shapeIdx; updatedText = $params.text }
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

function Invoke-PptDeleteSlide($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    $slideIdx = [int]$params.slideIndex
    
    $ppt = Get-RunningApp "PowerPoint.Application"
    $wasRunning = ($null -ne $ppt)
    if (-not $wasRunning) { $ppt = New-Object -ComObject PowerPoint.Application }
    
    $pres = $null
    foreach ($p in $ppt.Presentations) {
        if ((Normalize-DocPath $p.FullName) -eq $srcPath) { $pres = $p; break }
    }
    $openedHere = $false
    if ($null -eq $pres) {
        $pres = $ppt.Presentations.Open($srcPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)
        $openedHere = $true
    }
    
    try {
        if ($slideIdx -gt $pres.Slides.Count -or $slideIdx -lt 1) {
            return @{ ok = $false; error = "Slide index $slideIdx out of bounds (1..$($pres.Slides.Count))" }
        }
        $pres.Slides.Item($slideIdx).Delete()
        $pres.Save()
        return @{ ok = $true; path = $srcPath; deletedIndex = $slideIdx; remainingSlides = $pres.Slides.Count }
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

# --- Excel Actions ---

function Invoke-ExcelReadRange($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    
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
            path = $srcPath
            sheet = $sheet.Name
            range = $rangeStr
            rows = $range.Rows.Count
            cols = $range.Columns.Count
            values = $data
        }
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

function Invoke-ExcelWriteRange($params) {
    if (-not $params.path) { return @{ ok = $false; error = "Path is required" } }
    $srcPath = Normalize-DocPath $params.path
    if (-not (Test-Path $srcPath)) { return @{ ok = $false; error = "File not found: $srcPath" } }
    if (-not $params.data) { return @{ ok = $false; error = "data 2D array is required" } }
    
    $excel = Get-RunningApp "Excel.Application"
    $wasRunning = ($null -ne $excel)
    if (-not $wasRunning) { $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false }
    
    $wb = $null
    foreach ($w in $excel.Workbooks) {
        if ((Normalize-DocPath $w.FullName) -eq $srcPath) { $wb = $w; break }
    }
    $openedHere = $false
    if ($null -eq $wb) {
        $wb = $excel.Workbooks.Open($srcPath, [Type]::Missing, $false)
        $openedHere = $true
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
                try {
                    $cell.Value2 = $val
                } catch {
                    try {
                        $cell.Formula = [string]$val
                    } catch {
                        $cell.Value2 = [string]$val
                    }
                }
            }
        }
        
        $wb.Save()
        return @{ ok = $true; path = $srcPath; sheet = $sheet.Name; writtenRows = $data.Count }
    } finally {
        if ($openedHere) {
            try { $wb.Close($true) } catch {}
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

# --- Main Entry Point ---

try {
    $rawPayload = Get-Content -Path $PayloadFile -Raw -Encoding UTF8
    $request = $rawPayload | ConvertFrom-Json
    
    $action = $request.action
    $params = $request.params
    
    $result = Invoke-ComWithRetry {
        switch ($action) {
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
            "append_section"      { Invoke-WordAppendSection $params }
            "insert_table"        { Invoke-WordInsertTable $params }
            "insert_image"        { Invoke-InsertImage $params }
            "read_slides"         { Invoke-PptReadSlides $params }
            "add_slide"           { Invoke-PptAddSlide $params }
            "set_slide_text"      { Invoke-PptSetSlideText $params }
            "delete_slide"        { Invoke-PptDeleteSlide $params }
            "read_range"          { Invoke-ExcelReadRange $params }
            "write_range"         { Invoke-ExcelWriteRange $params }
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

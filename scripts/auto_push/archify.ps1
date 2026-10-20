<#
.SYNOPSIS
    Archify: Herramienta de empaquetado, respaldos de facturas, logs y snapshots.
.DESCRIPTION
    Empaqueta y respalda facturas descargadas, capturas de pantalla de error,
    archivos temporales o genera snapshots completos del código fuente del proyecto.
.PARAMETER Mode
    Modo de empaquetado:
    - Invoices: Respalda facturas PDF, XML y capturas de error (error_*.png) de /descargas o raíz.
    - Full: Genera snapshot limpio del código fuente excluyendo node_modules, browser_data y .env.
    - All: Combina snapshot de código y respaldo de facturas en archivos independientes.
.PARAMETER Tag
    Etiqueta personalizada para el nombre del archivo.
.PARAMETER OutputDir
    Directorio de destino (por defecto: <repo_root>/archives).
.PARAMETER CleanAfterArchive
    Si se especifica, elimina los archivos temporales/facturas respaldadas tras generar el ZIP.
.PARAMETER DryRun
    Modo simulación, muestra los archivos que se incluirían sin crear el archivo ZIP.
.EXAMPLE
    .\scripts\auto_push\archify.ps1 -Mode Invoices
.EXAMPLE
    .\scripts\auto_push\archify.ps1 -Mode Full -Tag "version-1.1"
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, Mandatory = $false)]
    [ValidateSet("Invoices", "Full", "All")]
    [string]$Mode = "Invoices",

    [Parameter(Mandatory = $false)]
    [Alias("t", "Name")]
    [string]$Tag = "",

    [Parameter(Mandatory = $false)]
    [Alias("o")]
    [string]$OutputDir = "",

    [Parameter(Mandatory = $false)]
    [switch]$CleanAfterArchive,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Determinar raíz del proyecto
try {
    $repoRoot = (git rev-parse --show-toplevel 2>$null)
    if ($LASTEXITCODE -eq 0 -and $repoRoot -and (Test-Path $repoRoot.Trim())) {
        $repoRoot = $repoRoot.Trim()
    } else {
        $repoRoot = (Resolve-Path "$PSScriptRoot\..\..").Path
    }
} catch {
    $repoRoot = (Resolve-Path "$PSScriptRoot\..\..").Path
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $repoRoot "archives"
}
if (-not (Test-Path $OutputDir) -and -not $DryRun) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$gitBranch = "master"
try {
    $bOut = & git -C $repoRoot branch --show-current 2>&1
    if ($LASTEXITCODE -eq 0 -and $bOut) { $gitBranch = $bOut.ToString().Trim() }
} catch { $gitBranch = "master" }

$gitCommit = "initial"
try {
    $cOut = & git -C $repoRoot rev-parse --short HEAD 2>&1
    if ($LASTEXITCODE -eq 0 -and $cOut) { $gitCommit = $cOut.ToString().Trim() }
} catch { $gitCommit = "initial" }

$tagSuffix = if ($Tag) { "_$($Tag -replace '[^a-zA-Z0-9_\-]', '_')" } else { "" }

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " 📦 ARCHIFY: Empaquetado y Respaldos (Invoice Duplicate Verifier)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "📁 Directorio base:  $repoRoot" -ForegroundColor Gray
Write-Host "⚙️  Modo activo:      $Mode" -ForegroundColor White
Write-Host "🌿 Rama / Commit:    $gitBranch ($gitCommit)" -ForegroundColor Gray
Write-Host "🎯 Destino:          $OutputDir" -ForegroundColor Gray
if ($DryRun) {
    Write-Host "🔍 MODO DRY-RUN ACTIVO (Simulación sin escritura)" -ForegroundColor Magenta
}
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

function Create-ArchiveFromFiles {
    param(
        [string]$ZipPath,
        [System.Collections.Generic.List[PSObject]]$FileList,
        [hashtable]$ManifestData,
        [switch]$DeleteSources
    )

    if ($FileList.Count -eq 0) {
        Write-Host "ℹ️  No se encontraron archivos para incluir en el archivo ZIP." -ForegroundColor Yellow
        return
    }

    $totalBytes = 0
    foreach ($item in $FileList) { $totalBytes += $item.Length }

    Write-Host "`n📊 Archivos a empaquetar: $($FileList.Count) ($([math]::Round($totalBytes / 1MB, 2)) MB)" -ForegroundColor Green

    if ($DryRun) {
        Write-Host "`n📋 Vista previa (primeros 20 archivos):" -ForegroundColor Cyan
        $FileList | Select-Object -First 20 | ForEach-Object {
            Write-Host "   + $($_.RelativePath) ($([math]::Round($_.Length / 1KB, 1)) KB)" -ForegroundColor DarkGray
        }
        return
    }

    $manifestPath = Join-Path $env:TEMP "manifest_$timestamp.json"
    $ManifestData.filesCount = $FileList.Count
    $ManifestData.uncompressedBytes = $totalBytes
    $ManifestData.createdAt = (Get-Date).ToString("o")
    $ManifestData | ConvertTo-Json -Depth 5 | Out-File -FilePath $manifestPath -Encoding utf8

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression

    if (Test-Path $ZipPath) { Remove-Item -Path $ZipPath -Force }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $zip = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Create)

    try {
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip, $manifestPath, "manifest.json", [System.IO.Compression.CompressionLevel]::Optimal
        ) | Out-Null

        foreach ($file in $FileList) {
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $zip, $file.FullPath, $file.RelativePath.Replace('\', '/'), [System.IO.Compression.CompressionLevel]::Optimal
            ) | Out-Null
        }
    } finally {
        $zip.Dispose()
        if (Test-Path $manifestPath) { Remove-Item -Path $manifestPath -Force -ErrorAction SilentlyContinue }
    }

    $sw.Stop()
    $zipSize = (Get-Item $ZipPath).Length
    $ratio = [math]::Round((1 - ($zipSize / $totalBytes)) * 100, 1)

    Write-Host "`n============================================================" -ForegroundColor Green
    Write-Host " 🎉 ARCHIVO GENERADO CON ÉXITO" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "📁 Archivo:       $ZipPath" -ForegroundColor White
    Write-Host "📄 Archivos:      $($FileList.Count) empaquetados (+ manifest.json)" -ForegroundColor White
    Write-Host "💾 Tamaño orig:   $([math]::Round($totalBytes / 1MB, 2)) MB" -ForegroundColor White
    Write-Host "🗜️  Tamaño ZIP:    $([math]::Round($zipSize / 1MB, 2)) MB (Reducción: $ratio%)" -ForegroundColor White
    Write-Host "⏱️  Tiempo:        $([math]::Round($sw.Elapsed.TotalSeconds, 2)) segundos" -ForegroundColor White

    if ($DeleteSources -and -not $DryRun) {
        Write-Host "`n🧹 Limpiando archivos temporales procesados..." -ForegroundColor Yellow
        $cleaned = 0
        foreach ($file in $FileList) {
            if ($file.RelativePath -ne "facturas.xlsx" -and -not $file.RelativePath.StartsWith("src\")) {
                Remove-Item -Path $file.FullPath -Force -ErrorAction SilentlyContinue
                $cleaned++
            }
        }
        Write-Host "✅ Se eliminaron $cleaned archivos temporales respaldados." -ForegroundColor Green
    }
}

if ($Mode -eq "Invoices" -or $Mode -eq "All") {
    Write-Host "`n🔍 Buscando facturas, comprobantes descargados y capturas de error..." -ForegroundColor Yellow
    $zipName = "facturas_backup_${timestamp}_${gitBranch}${tagSuffix}.zip"
    $zipPath = Join-Path $OutputDir $zipName

    $invoiceFiles = [System.Collections.Generic.List[PSObject]]::new()
    $searchDirs = @($repoRoot, (Join-Path $repoRoot "descargas"), (Join-Path $repoRoot "downloads"))
    foreach ($dir in $searchDirs) {
        if (Test-Path $dir) {
            $files = Get-ChildItem -Path $dir -File | Where-Object {
                $_.Extension -match "\.(pdf|xml)$" -or
                $_.Name -match "^error_.*\.png$" -or
                $_.Name -match "^captura_.*\.png$" -or
                ($_.Extension -eq ".xlsx" -and $_.Name -ne "facturas.xlsx")
            }
            foreach ($f in $files) {
                $rel = $f.FullName.Substring($repoRoot.Length).TrimStart('\', '/')
                $invoiceFiles.Add([PSCustomObject]@{
                    FullPath     = $f.FullName
                    RelativePath = $rel
                    Length       = $f.Length
                })
            }
        }
    }

    $manifest = @{
        type        = "InvoicesAndLogsBackup"
        project     = "Auto_Download_Invoice"
        timestamp   = $timestamp
        gitBranch   = $gitBranch
        gitCommit   = $gitCommit
        cleaned     = [bool]$CleanAfterArchive
    }

    Create-ArchiveFromFiles -ZipPath $zipPath -FileList $invoiceFiles -ManifestData $manifest -DeleteSources:$CleanAfterArchive
}

if ($Mode -eq "Full" -or $Mode -eq "All") {
    Write-Host "`n🔍 Escaneando código fuente del proyecto para Snapshot..." -ForegroundColor Yellow
    $zipName = "snapshot_source_${timestamp}_${gitBranch}_${gitCommit}${tagSuffix}.zip"
    $zipPath = Join-Path $OutputDir $zipName

    $excludeDirs = @("node_modules", ".git", "browser_data", ".astro", "dist", "archives", "descargas", ".vscode", ".idea")
    $excludeFiles = @(".DS_Store", "Thumbs.db", "desktop.ini")

    $sourceFiles = [System.Collections.Generic.List[PSObject]]::new()
    $all = Get-ChildItem -Path $repoRoot -Recurse -File -Force

    foreach ($f in $all) {
        $rel = $f.FullName.Substring($repoRoot.Length).TrimStart('\', '/')
        $parts = $rel.Split([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)

        $skip = $false
        foreach ($p in $parts) {
            if ($excludeDirs -contains $p) { $skip = $true; break }
        }
        if ($skip) { continue }
        if ($excludeFiles -contains $f.Name) { continue }
        if ($f.Name.StartsWith(".env") -and -not $f.Name.EndsWith(".example")) { continue }
        if ($f.Extension -match "\.(zip|tar|gz|rar)$") { continue }
        if ($f.Name -match "^error_.*\.png$") { continue }

        $sourceFiles.Add([PSCustomObject]@{
            FullPath     = $f.FullName
            RelativePath = $rel
            Length       = $f.Length
        })
    }

    $manifest = @{
        type        = "FullSourceSnapshot"
        project     = "Auto_Download_Invoice"
        timestamp   = $timestamp
        gitBranch   = $gitBranch
        gitCommit   = $gitCommit
        tag         = $Tag
    }

    Create-ArchiveFromFiles -ZipPath $zipPath -FileList $sourceFiles -ManifestData $manifest
}


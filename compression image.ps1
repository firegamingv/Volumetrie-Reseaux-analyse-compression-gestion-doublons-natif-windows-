<#
.SYNOPSIS
    Compression d'images pour un dossier Pilotage donné,
    sur le CSV 'DonneeVolumetrique_IMG_VIDEOS.csv',
    avec skip des fichiers déjà traités, batchs pour gros volumes,
    gestion des verrouillages du rapport,
    affichage console détaillé par fichier, notification de skip, et vérification de la taille finale.

.PARAMETER PilotagePath
    Chemin complet vers le dossier Pilotage du service.
.PARAMETER TargetWidth
    Largeur maximale après redimensionnement (défaut : 1224).
.PARAMETER TargetHeight
    Hauteur maximale après redimensionnement (défaut : 1632).
.PARAMETER JpegQuality
    Qualité JPEG (1–100, défaut : 75).
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$PilotagePath,
    [int]$TargetWidth  = 1224,
    [int]$TargetHeight = 1632,
    [int]$JpegQuality  = 75
)

if (-not (Test-Path $PilotagePath)) {
    Throw "Le dossier Pilotage spécifié n'existe pas : $PilotagePath"
}

# Chemin du rapport
[string]$report = Join-Path $PilotagePath 'CompressionReport_Final.csv'

# Fonction de test de verrouillage
function Test-FileLocked {
    param([string]$Path)
    try { $fs = [IO.File]::Open($Path, 'Open','ReadWrite','None'); $fs.Close(); return $false } catch { return $true }
}

# Ajout sécurisé au rapport
function Append-ReportLine {
    param([string]$Line)
    try { Add-Content -Path $report -Value $Line -Encoding UTF8 } catch {
        Write-Host "Impossible d'écrire dans le rapport (peut-être ouvert): $report" -ForegroundColor Red
    }
}

# Initialisation ou nettoyage
if (Test-Path $report) {
    if (-not (Test-FileLocked $report)) {
        $kept = Get-Content $report | Where-Object { -not ($_ -match '^(Total;;|Summary;)') }
        try { Set-Content -Path $report -Value $kept -Encoding UTF8 } catch {
            Write-Host "Erreur nettoyage rapport: $report" -ForegroundColor Red
        }
    } else {
        Write-Host "Attention: rapport ouvert, synthèses précédentes conservées." -ForegroundColor Yellow
    }
} else {
    try { Set-Content -Path $report -Value 'File;Path;OriginalSize;CompressedSize;DeltaBytes' -Encoding UTF8 } catch {
        Write-Host "Erreur création rapport: $report" -ForegroundColor Red
    }
}

# Skip déjà traités
$done = @{}
Import-Csv -Path $report -Delimiter ';' -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.Path) { $done[$_.Path] = $true }
}

# Préparation CSV unique
[string]$csvPath = Join-Path $PilotagePath 'DonneeVolumetrique_IMG_VIDEOS.csv'
if (-not (Test-Path $csvPath)) { Throw "Fichier introuvable : $csvPath" }
[string[]]$headers = (Get-Content $csvPath -First 1) -split ';'

# Compteurs
[long]$totalOrig = 0; [long]$totalNew = 0
$totalEntries = (Get-Content $csvPath | Measure-Object).Count - 1
[int]$processed = 0

# Codecs
Add-Type -AssemblyName PresentationCore
$jpegCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object MimeType -eq 'image/jpeg'
$encParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
$encParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
    [System.Drawing.Imaging.Encoder]::Quality, [int64]$JpegQuality)

function Get-ResizedDimensions {
    param($w,$h,$maxW,$maxH)
    if ($w -le $maxW -and $h -le $maxH) { return @($w,$h) }
    $r = $w/$h; $rMax = $maxW/$maxH
    if ($r -gt $rMax) { $nw = $maxW; $nh = [int]($maxW/$r) } else { $nh = $maxH; $nw = [int]($maxH*$r) }
    return @([math]::Max($nw,1), [math]::Max($nh,1))
}

# Lecture par batch
[int]$batchSize = 10000
$sr = [System.IO.File]::OpenText($csvPath)
$sr.ReadLine() | Out-Null  # saute entête

while (-not $sr.EndOfStream) {
    $batch = for ($i=0; $i -lt $batchSize -and -not $sr.EndOfStream; $i++) { $sr.ReadLine() }
    $rows = $batch | ConvertFrom-Csv -Delimiter ';' -Header $headers

    foreach ($e in $rows) {
        $src = Join-Path $e.Chemin $e.Nom
        $base = [IO.Path]::GetFileNameWithoutExtension($src)
        # Skip conditions
        if ($done.ContainsKey($src)) {
            Write-Host "[SKIP] $base - déjà traité" -ForegroundColor Yellow
            $processed++
            continue
        }
        if (-not (Test-Path $src)) {
            Write-Host "[SKIP] $base - introuvable" -ForegroundColor Yellow
            $processed++
            continue
        }
        $ext = [IO.Path]::GetExtension($src).ToLower()
        if ($ext -notmatch '\.jpe?g|\.png|\.gif') {
            Write-Host "[SKIP] $base - format non pris en charge ($ext)" -ForegroundColor Yellow
            $processed++
            continue
        }

        # Compression
        $orig = (Get-Item $src).Length; $totalOrig += $orig
        try { $img = [System.Drawing.Image]::FromFile($src) } catch {
            $processed++
            continue
        }
        $nw, $nh = Get-ResizedDimensions $img.Width $img.Height $TargetWidth $TargetHeight
        $bmp = New-Object System.Drawing.Bitmap($nw, $nh)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.DrawImage($img, [System.Drawing.Rectangle]::new(0,0,$nw,$nh))

        $temp = Join-Path ([IO.Path]::GetDirectoryName($src)) ("${base}_compressed$ext")
        try {
            switch ($ext) {
                '\.jpe?g' { $bmp.Save($temp, $jpegCodec, $encParams) }
                '\.png'   {
                    $st = [IO.File]::OpenRead($src)
                    $bi = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bi.BeginInit(); $bi.StreamSource = $st; $bi.CacheOption = 'OnLoad'; $bi.EndInit(); $st.Close()
                    $pe = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
                    $pe.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bi))
                    $fs = [IO.File]::Open($temp, 'Create'); $pe.Save($fs); $fs.Close()
                }
                default { $bmp.Save($temp, $img.RawFormat) }
            }
        } catch {
            $g.Dispose(); $bmp.Dispose(); $img.Dispose(); $processed++
            continue
        } finally {
            $g.Dispose(); $bmp.Dispose(); $img.Dispose()
        }

        $new = (Get-Item $temp).Length
        $delta = $orig - $new
        # Vérification de la taille finale
        if ($new -gt $orig) {
            # Nouveau plus gros : pas de gain
            Remove-Item $temp -Force
            $totalNew += $orig
            $deltaToReport = 0
            Write-Host "[NO GAIN] $base – nouveau ($([math]::Round($new/1MB,2)) Mo) > origine ($([math]::Round($orig/1MB,2)) Mo)" -ForegroundColor Magenta
        } elseif ($new -eq $orig) {
            # Taille identique : pas de gain
            Remove-Item $temp -Force
            $totalNew += $orig
            $deltaToReport = 0
            Write-Host "[NO GAIN] $base – taille identique ($([math]::Round($orig/1MB,2)) Mo)" -ForegroundColor Yellow
        } else {
            # Gain réel
            Remove-Item $src -Force
            Rename-Item $temp -NewName "$base$ext" -Force
            $totalNew += $new
            $deltaToReport = $delta
            Write-Host "[COMPRESSED] $base – Δ $([math]::Round($delta/1MB,2)) Mo ($delta bytes)" -ForegroundColor Green
        }

        # Enregistrement dans le rapport
        $line = "${base};${src};${orig};$([int]($orig - $deltaToReport));${deltaToReport}"
        Append-ReportLine $line

        # Progression
        $processed++
        $pct = [math]::Round($processed / $totalEntries * 100, 1)
        Write-Progress -Activity 'Compression' -Status "$pct%" -PercentComplete $pct

        # Marquer comme traité
        $done[$src] = $true
    }
}
$sr.Close()

# Synthèse finale
$totalDelta = $totalOrig - $totalNew
$gain      = if ($totalOrig -gt 0) {[math]::Round(($totalDelta/$totalOrig)*100,2)} else {0}
Append-ReportLine "Total;;$totalOrig;$totalNew;$totalDelta"
Append-ReportLine "Summary;%%Gain;;;$gain%"
Write-Host "`nOrigine: $([math]::Round($totalOrig/1MB,2)) Mo, Compressé: $([math]::Round($totalNew/1MB,2)) Mo, Économisé: $([math]::Round($totalDelta/1MB,2)) Mo ($gain% gain)" -ForegroundColor Yellow

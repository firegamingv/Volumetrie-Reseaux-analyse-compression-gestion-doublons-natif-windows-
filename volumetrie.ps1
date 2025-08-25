param(
    [Parameter(Mandatory=$true)]
    [string]$BaseServicePath,

    [Parameter(Mandatory=$true)]
    [string[]]$SubFolders,

    [Parameter(Mandatory=$true)]
    [string]$PilotagePath
)

# création du dossier pilote si nécessaire
if (-not (Test-Path -LiteralPath $PilotagePath)) {
    New-Item -Path $PilotagePath -ItemType Directory | Out-Null
}

# 1) création d’un sous-dossier par service
foreach ($sub in $SubFolders) {
    $serviceFolder = Join-Path $PilotagePath $sub
    if (-not (Test-Path -LiteralPath $serviceFolder)) {
        New-Item -Path $serviceFolder -ItemType Directory | Out-Null
    }
}

# motifs et catégories
$Patterns = [ordered]@{
    PDF_ECM         = 'D5140|D4533'
    PDF             = '\.pdf$'
    Archive_OUTLOOK = '\.pst$'
    IMG_VIDEOS      = '\.(3gp|ai|asf|avi|bmp|cdr|drs|flac|flv|gif|jpe?g|m4[av]|mbt|mdl|mkv|mov|mp2|mp3|mp4|mpeg|mpg|mts|ogg|pak|pex|png|psd|raw|rec|rm|swf|tga|tiff?|tr2|ts|umx|unity3D|unr|utx|vob|vro|wad|wav|wbc|wma|wmv)$'
    ZIP_unzip       = '\.zip$'
}
$Categories = $Patterns.Keys + 'Autres','Doublons'

# 2) création des CSV (en-têtes) DANS CHAQUE sous-dossier
$outFiles = @{}
foreach ($sub in $SubFolders) {
    $serviceFolder = Join-Path $PilotagePath $sub
    foreach ($cat in $Categories) {
        $csv = Join-Path $serviceFolder "DonneeVolumetrique_$cat.csv"
        'Nom;Chemin;Taille (Mo);Extension;Proprietaire;DernierAcces;DerniereModif' |
            Set-Content -Path $csv -Encoding UTF8
        $outFiles["$sub|$cat"] = $csv
    }
}

# préparation des chemins sources
$SourcePaths = $SubFolders | ForEach-Object { Join-Path $BaseServicePath $_ }
foreach ($p in $SourcePaths) {
    if (-not (Test-Path $p)) { throw "Le dossier source n'existe pas : $p" }
}

# récupération de tous les fichiers (tous services mélangés)
$allFiles = Get-ChildItem -Path $SourcePaths -File -Recurse -Force `
               -ErrorAction SilentlyContinue -Exclude 'Thumbs.db','~$*'
$total   = $allFiles.Count
Write-Host "Total fichiers : $total"

# tampons et compteurs
$batchSize    = 10000
$processed    = 0
# structure : $categoryData[$service][$cat]
$categoryData = @{}
foreach ($sub in $SubFolders) {
    $categoryData[$sub] = @{}
    foreach ($cat in $Categories) {
        $categoryData[$sub][$cat] = @()
    }
}

# 3) boucle de traitement
$i = 0
foreach ($file in $allFiles) {
    $i++; $processed++

    # détermination du service à partir du chemin
    $relative = $file.FullName.Substring($BaseServicePath.Length + 1)
    $service  = $relative.Split('\')[0]

    # catégorie
    $cat = 'Autres'
    foreach ($kv in $Patterns.GetEnumerator()) {
        if ($file.Name -imatch $kv.Value) { $cat = $kv.Key; break }
    }

    # propriétaire
    try {
        $owner = (Get-Acl -LiteralPath $file.FullName -ErrorAction Stop).Owner -replace '^.*\\',''
    } catch {
        $owner = 'Inconnu'
    }

    # enregistrement en mémoire tampon
    $rec = [PSCustomObject]@{
        Nom           = $file.Name
        Chemin        = $file.DirectoryName
        'Taille (Mo)' = [math]::Round($file.Length/1MB,1)
        Extension     = $file.Extension
        Proprietaire  = $owner
        DernierAcces  = $file.LastAccessTime.ToShortDateString()
        DerniereModif = $file.LastWriteTime.ToShortDateString()
    }
    $categoryData[$service][$cat] += $rec

    # flush par lots
    if ($processed -ge $batchSize) {
        Write-Host "→ Écriture de $batchSize lignes…"
        foreach ($sub in $SubFolders) {
            foreach ($c in $Categories) {
                $data = $categoryData[$sub][$c]
                if ($data.Count -gt 0) {
                    $data |
                      Select-Object Nom,Chemin,'Taille (Mo)',Extension,Proprietaire,DernierAcces,DerniereModif |
                      ConvertTo-Csv -NoTypeInformation -Delimiter ';' |
                      Select-Object -Skip 1 |
                      Out-File -FilePath $outFiles["$sub|$c"] -Encoding UTF8 -Append

                    $categoryData[$sub][$c] = @()
                }
            }
        }
        $processed = 0
        [GC]::Collect()
    }

    Write-Progress -Activity 'Traitement fichiers' -Status "$i/\$total" `
                   -PercentComplete ([int](($i/$total)*100))
}

# dernier flush pour les restes
Write-Host "→ Écriture des lignes restantes…"
foreach ($sub in $SubFolders) {
    foreach ($c in $Categories) {
        $data = $categoryData[$sub][$c]
        if ($data.Count -gt 0) {
            $data |
              Select-Object Nom,Chemin,'Taille (Mo)',Extension,Proprietaire,DernierAcces,DerniereModif |
              ConvertTo-Csv -NoTypeInformation -Delimiter ';' |
              Select-Object -Skip 1 |
              Out-File -FilePath $outFiles["$sub|$c"] -Encoding UTF8 -Append
        }
    }
}

Write-Host "Traitement terminé."

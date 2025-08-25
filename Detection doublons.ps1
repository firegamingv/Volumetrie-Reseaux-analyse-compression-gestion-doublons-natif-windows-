[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ServiceName,
    [string]$PilotageRoot = 'a remplir',
    [int]$NameThreshold = 1,
    [double]$MaxSizeDiff = 1,
    [int]$StreamBufferSize = 2000,
    [int]$HashBatchSize = 100,
    [int]$PartialHashSize = 65536,     # 64KB pour hash partiel
    [int]$SmallFileThreshold = 10485760, # 10MB - seuil pour hash complet
    [switch]$SkipHashValidation,
    [switch]$NoConsolidation,
    [switch]$UseExtendedMetadata       # Utilise date de modification
)

#region Fonctions 
function Test-NetworkConnectivity {
    param($ServiceName)
    
    Write-Host "🔍 Test de connectivité réseau..." -ForegroundColor Cyan
    
    $tests = @{
        "Serveur atlas.edf.fr" = "\\atlas.edf.fr"
        "Services CO" = "\\atlas.edf.fr\CO"
        "Dossier services.006" = "\\atlas.edf.fr\CO\45dam-dpn\services.006"
        "Service spécifique" = "\\atlas.edf.fr\CO\45dam-dpn\services.006\$ServiceName"
    }
    
    $results = @{}
    foreach ($test in $tests.GetEnumerator()) {
        try {
            $accessible = Test-Path -Path $test.Value -ErrorAction SilentlyContinue
            $results[$test.Key] = $accessible
            $status = if ($accessible) { "✓" } else { "✗" }
            $color = if ($accessible) { "Green" } else { "Red" }
            Write-Host "  $status $($test.Key): $($test.Value)" -ForegroundColor $color
        }
        catch {
            $results[$test.Key] = $false
            Write-Host "   $($test.Key): $($test.Value) (Erreur: $($_.Exception.Message))" -ForegroundColor Red
        }
    }
    
    return $results
}

function Get-FastLevenshteinDistance {
    param([string]$s, [string]$t, [int]$maxDistance = 3)
    
    $lenS = $s.Length
    $lenT = $t.Length
    
    if ([Math]::Abs($lenS - $lenT) -gt $maxDistance) { return $maxDistance + 1 }
    if ($lenS -eq 0) { return $lenT }
    if ($lenT -eq 0) { return $lenS }
    if ($s -eq $t) { return 0 }
    
    if ($lenS -gt 0 -and $lenT -gt 0 -and $s[0] -ne $t[0] -and $maxDistance -eq 1) {
        return 2
    }
    
    $prevRow = 0..$lenT
    $currRow = New-Object int[] ($lenT + 1)
    
    for ($i = 1; $i -le $lenS; $i++) {
        $currRow[0] = $i
        $minInRow = $i
        
        for ($j = 1; $j -le $lenT; $j++) {
            $cost = if ($s[$i - 1] -eq $t[$j - 1]) { 0 } else { 1 }
            $currRow[$j] = [Math]::Min([Math]::Min($currRow[$j - 1] + 1, $prevRow[$j] + 1), $prevRow[$j - 1] + $cost)
            if ($currRow[$j] -lt $minInRow) { $minInRow = $currRow[$j] }
        }
        
        if ($minInRow -gt $maxDistance) { return $maxDistance + 1 }
        
        $temp = $prevRow; $prevRow = $currRow; $currRow = $temp
    }
    
    return $prevRow[$lenT]
}

function ConvertTo-CompactKey {
    param([string]$name, [double]$size)
    $nameKey = if ($name.Length -gt 3) { $name.Substring(0, 3).ToLower() } else { $name.ToLower() }
    $sizeKey = [Math]::Floor($size / 0.5) * 0.5
    return "$nameKey|$sizeKey"
}

function Get-PartialFileHash {
    param(
        [string]$FilePath,
        [int]$PartialSize = 65536,  # 64KB par défaut
        [string]$Algorithm = "MD5"
    )
    
    try {
        $fileInfo = Get-Item -LiteralPath $FilePath -ErrorAction Stop
        $fileSize = $fileInfo.Length
        
        # Si fichier très petit, hash complet
        if ($fileSize -le $PartialSize * 2) {
            return (Get-FileHash -LiteralPath $FilePath -Algorithm $Algorithm -ErrorAction Stop).Hash
        }
        
        # Hash partiel : début + milieu + fin
        $stream = [System.IO.File]::OpenRead($FilePath)
        $hasher = [System.Security.Cryptography.MD5]::Create()
        
        # Début du fichier
        $startBuffer = New-Object byte[] $PartialSize
        $bytesRead = $stream.Read($startBuffer, 0, $PartialSize)
        $hasher.TransformBlock($startBuffer, 0, $bytesRead, $null, 0) | Out-Null
        
        # Milieu du fichier
        if ($fileSize -gt $PartialSize * 4) {
            $middlePosition = [Math]::Floor($fileSize / 2) - [Math]::Floor($PartialSize / 2)
            $stream.Seek($middlePosition, [System.IO.SeekOrigin]::Begin) | Out-Null
            $middleBuffer = New-Object byte[] $PartialSize
            $bytesRead = $stream.Read($middleBuffer, 0, $PartialSize)
            $hasher.TransformBlock($middleBuffer, 0, $bytesRead, $null, 0) | Out-Null
        }
        
        # Fin du fichier
        $endPosition = [Math]::Max(0, $fileSize - $PartialSize)
        $stream.Seek($endPosition, [System.IO.SeekOrigin]::Begin) | Out-Null
        $endBuffer = New-Object byte[] $PartialSize
        $bytesRead = $stream.Read($endBuffer, 0, $PartialSize)
        $hasher.TransformFinalBlock($endBuffer, 0, $bytesRead) | Out-Null
        
        $hash = [System.BitConverter]::ToString($hasher.Hash) -replace '-', ''
        
        $stream.Close()
        $hasher.Dispose()
        
        return $hash
    }
    catch {
        if ($stream) { $stream.Close() }
        if ($hasher) { $hasher.Dispose() }
        return $null
    }
}

function Read-CsvStreaming {
    param($CsvPath, $BufferSize, $ServiceName, [switch]$IncludeExtendedMetadata)
    
    Write-Host "    • Lecture streaming : $([System.IO.Path]::GetFileName($CsvPath))" -ForegroundColor Gray
    
    $entries = @{}
    $processedCount = 0
    
    try {
        $reader = [System.IO.File]::OpenText($CsvPath)
        $headerLine = $reader.ReadLine()
        if (-not $headerLine) { return $entries }
        
        $sep = if ($headerLine -like '*;*') { ';' } else { ',' }
        $headers = $headerLine -split $sep
        
        $nameIndex = -1; $pathIndex = -1; $sizeIndex = -1; $dateIndex = -1
        for ($i = 0; $i -lt $headers.Length; $i++) {
            $header = $headers[$i].Trim().ToLower()
            if ($header -match 'nom' -and $nameIndex -eq -1) { $nameIndex = $i }
            if ($header -match 'chemin' -and $pathIndex -eq -1) { $pathIndex = $i }
            if ($header -match 'taille' -and $sizeIndex -eq -1) { $sizeIndex = $i }
            if ($header -match 'date|modif' -and $dateIndex -eq -1) { $dateIndex = $i }
        }
        
        if ($nameIndex -lt 0 -or $pathIndex -lt 0 -or $sizeIndex -lt 0) { 
            $reader.Close()
            return $entries 
        }
        
        while (($line = $reader.ReadLine()) -ne $null) {
            $cols = $line -split $sep
            if ($cols.Length -le [Math]::Max([Math]::Max($nameIndex, $pathIndex), $sizeIndex)) { continue }
            
            try {
                $name = $cols[$nameIndex].Trim('"').Trim()
                $path = $cols[$pathIndex].Trim('"').Trim()
                $sizeStr = $cols[$sizeIndex] -replace '[",]', '' -replace '\.', ','
                $size = [double]$sizeStr
                
                # Métadonnées étendues
                $dateModified = $null
                if ($IncludeExtendedMetadata -and $dateIndex -ge 0 -and $dateIndex -lt $cols.Length) {
                    try {
                        $dateStr = $cols[$dateIndex].Trim('"').Trim()
                        if ($dateStr -ne "") {
                            $dateModified = [DateTime]::Parse($dateStr)
                        }
                    } catch {
                        # Ignore les erreurs de parsing de date
                    }
                }
                
                if ($name -and $path -and $size -ge 0) {
                    $uniqueKey = "$name|$path|$size"
                    if (-not $entries.ContainsKey($uniqueKey)) {
                        $entry = @{
                            Nom = $name
                            Chemin = $path  
                            Taille = $size
                            CompactKey = ConvertTo-CompactKey $name $size
                        }
                        
                        if ($dateModified) {
                            $entry.DateModified = $dateModified
                        }
                        
                        $entries[$uniqueKey] = $entry
                    }
                    $processedCount++
                }
            } catch { continue }
            
            if ($processedCount % 20000 -eq 0) {
                [System.GC]::Collect()
            }
        }
        $reader.Close()
    }
    catch {
        if ($reader) { $reader.Close() }
        Write-Warning "Erreur lecture $CsvPath : $($_.Exception.Message)"
    }
    
    Write-Host "       $processedCount entrées lues" -ForegroundColor Green
    return $entries
}

function Find-DuplicatesUltraFast {
    param($AllEntries, $NameThreshold, $MaxSizeDiff, [switch]$UseExtendedMetadata)
    
    Write-Host "→ Recherche ultra-rapide des doublons avec métadonnées étendues..." -ForegroundColor Cyan
    
    $groups = @{}
    $totalEntries = $AllEntries.Count
    $processed = 0
    
    Write-Host "  • Regroupement intelligent par similarité..." -ForegroundColor Gray
    foreach ($entry in $AllEntries.Values) {
        $key = $entry.CompactKey
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = @()
        }
        $groups[$key] += $entry
        
        $processed++
        if ($processed % 50000 -eq 0) {
            Write-Progress -Activity "Regroupement" -Status "$processed/$totalEntries" -PercentComplete (($processed * 100) / $totalEntries)
            [System.GC]::Collect()
        }
    }
    
    Write-Host "  • $($groups.Count) groupes créés (réduction : $([Math]::Round((1 - ($groups.Count / $totalEntries)) * 100, 1))%)" -ForegroundColor Green
    
    $candidates = @()
    $comparisons = 0
    $groupsProcessed = 0
    
    foreach ($group in $groups.Values) {
        if ($group.Count -lt 2) { 
            $groupsProcessed++
            continue 
        }
        
        $sortedGroup = $group | Sort-Object Nom
        
        for ($i = 0; $i -lt $sortedGroup.Count - 1; $i++) {
            $e1 = $sortedGroup[$i]
            $maxJ = [Math]::Min($i + 20, $sortedGroup.Count - 1)
            
            for ($j = $i + 1; $j -le $maxJ; $j++) {
                $e2 = $sortedGroup[$j]
                $comparisons++
                
                # Pré-filtres ultra-rapides
                if ($e1.Chemin -eq $e2.Chemin) { continue }
                if ([Math]::Abs($e1.Taille - $e2.Taille) -gt $MaxSizeDiff) { continue }
                if ([Math]::Abs($e1.Nom.Length - $e2.Nom.Length) -gt $NameThreshold) { continue }
                
                # Comparaison des noms
                if ($e1.Nom -eq $e2.Nom) {
                    $distance = 0
                } else {
                    $distance = Get-FastLevenshteinDistance -s $e1.Nom -t $e2.Nom -maxDistance $NameThreshold
                }
                
                if ($distance -le $NameThreshold) {
                    # Validation avec métadonnées étendues
                    $isValidCandidate = $true
                    $confidence = "Standard"
                    
                    # Niveau 1: Nom + Taille exacte = Haute confiance
                    if ($distance -eq 0 -and [Math]::Abs($e1.Taille - $e2.Taille) -lt 0.001) {
                        $confidence = "Haute"
                    }
                    
                    # Niveau 2: Ajout de la date de modification si disponible
                    if ($UseExtendedMetadata -and $e1.DateModified -and $e2.DateModified) {
                        $timeDiff = [Math]::Abs(($e1.DateModified - $e2.DateModified).TotalSeconds)
                        $daysDiff = [Math]::Abs(($e1.DateModified.Date - $e2.DateModified.Date).TotalDays)
                        
                        # OBLIGATOIRE: Même jour requis
                        if ($daysDiff -gt 0) {
                            # Fichiers de jours différents = pas un doublon
                            $isValidCandidate = $false
                        } elseif ($timeDiff -lt 5) {  
                            # Même jour + même heure (5 sec) = Très haute confiance
                            $confidence = "Très Haute"
                        } elseif ($timeDiff -lt 3600) {  
                            # Même jour + même heure (1h) = Haute confiance
                            $confidence = "Haute"
                        } else {
                            # Même jour mais horaires différents = Moyenne confiance
                            $confidence = "Moyenne"
                        }
                    } elseif ($UseExtendedMetadata) {
                        # Si métadonnées étendues activées mais dates manquantes
                        # On ne peut pas valider le critère "même jour" = rejet
                        $isValidCandidate = $false
                    }
                    
                    if ($isValidCandidate) {
                        $candidates += @{
                            Entry1 = $e1
                            Entry2 = $e2
                            Distance = $distance
                            SizeDiff = [Math]::Abs($e1.Taille - $e2.Taille)
                            Confidence = $confidence
                            TimeDiff = if ($e1.DateModified -and $e2.DateModified) { 
                                [Math]::Abs(($e1.DateModified - $e2.DateModified).TotalSeconds) 
                            } else { $null }
                            DaysDiff = if ($e1.DateModified -and $e2.DateModified) { 
                                [Math]::Abs(($e1.DateModified.Date - $e2.DateModified.Date).TotalDays) 
                            } else { $null }
                        }
                    }
                }
            }
        }
        
        $groupsProcessed++
        if ($groupsProcessed % 100 -eq 0) {
            Write-Progress -Activity "Recherche doublons" -Status "$groupsProcessed/$($groups.Count) groupes" -PercentComplete (($groupsProcessed * 100) / $groups.Count)
        }
    }
    
    Write-Progress -Activity "Recherche doublons" -Completed
    Write-Host "  • $comparisons comparaisons effectuées" -ForegroundColor Green
    Write-Host "  • $($candidates.Count) candidats doublons trouvés" -ForegroundColor Yellow
    
    # Statistiques par niveau de confiance
    $confidenceStats = $candidates | Group-Object Confidence | Sort-Object Name
    foreach ($stat in $confidenceStats) {
        Write-Host "    - $($stat.Name): $($stat.Count) candidats" -ForegroundColor Gray
    }
    
    # Statistiques des critères temporels
    if ($UseExtendedMetadata) {
        $sameDayCount = ($candidates | Where-Object { $_.DaysDiff -eq 0 }).Count
        $totalWithDates = ($candidates | Where-Object { $_.DaysDiff -ne $null }).Count
        if ($totalWithDates -gt 0) {
            Write-Host "    - Même jour (critère obligatoire): $sameDayCount/$totalWithDates" -ForegroundColor Green
        }
    }
    
    return $candidates
}

function Resolve-FilePath {
    param($OriginalPath, $ServiceName, $PathCache, [switch]$Verbose)
    
    if ($PathCache.ContainsKey($OriginalPath)) {
        return $PathCache[$OriginalPath]
    }
    
    $cleanPath = $OriginalPath.Trim().Trim('"')
    
    $alternativePaths = @(
        $cleanPath,
        ($cleanPath -replace '^C:', 'remplir'),
        ($cleanPath -replace '^D:', 'remplir'),
        ($cleanPath -replace '^E:', 'remplir'),
        ($cleanPath -replace '^F:', 'remplir'),
        ($cleanPath -replace '^[A-Z]:', 'remplir'),
        (Join-Path "remplir\$ServiceName" ([System.IO.Path]::GetFileName($cleanPath))),
        ($cleanPath -replace '^[A-Z]:', 'remplir'),
        (Join-Path $PilotageRoot ([System.IO.Path]::GetFileName($cleanPath)))
    )
    
    foreach ($path in $alternativePaths) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        
        try {
            if (Test-Path -LiteralPath $path -PathType Leaf -ErrorAction SilentlyContinue) {
                $PathCache[$OriginalPath] = $path
                if ($Verbose) {
                    Write-Host "     Résolu: $($cleanPath) → $path" -ForegroundColor Green
                }
                return $path
            }
        }
        catch {
            continue
        }
    }
    
    if ($Verbose) {
        Write-Host "     NON RÉSOLU: $cleanPath" -ForegroundColor Red
    }
    
    $PathCache[$OriginalPath] = $null
    return $null
}

function Validate-DuplicatesAdvanced {
    param($Candidates, $ServiceName, [switch]$SkipHashValidation, $HashBatchSize, $PartialHashSize, $SmallFileThreshold)
    
    if ($SkipHashValidation) {
        Write-Host " Validation par métadonnées étendues avec critère MÊME JOUR strict..." -ForegroundColor Yellow
        $confirmed = @()
        $rejectedDifferentDays = 0
        
        foreach ($candidate in $Candidates) {
            # VÉRIFICATION STRICTE : Même jour obligatoire
            if ($UseExtendedMetadata -and $candidate.DaysDiff -ne $null) {
                if ($candidate.DaysDiff -gt 0) {
                    $rejectedDifferentDays++
                    continue  # REJET IMMÉDIAT si jours différents
                }
            } elseif ($UseExtendedMetadata) {
                # Si métadonnées étendues activées mais pas de date = rejet
                continue
            }
            
            # Accepte seulement les candidats avec haute confiance ET même jour
            if ($candidate.Confidence -in @("Haute", "Très Haute", "Moyenne")) {
                $confirmed += $candidate
            } elseif ($candidate.Distance -eq 0 -and $candidate.SizeDiff -eq 0) {
                # Nom et taille identiques, mais seulement si même jour validé
                if (-not $UseExtendedMetadata -or ($candidate.DaysDiff -ne $null -and $candidate.DaysDiff -eq 0)) {
                    $confirmed += $candidate
                }
            }
        }
        
        Write-Host "  • $($confirmed.Count) doublons confirmés par métadonnées (même jour)" -ForegroundColor Green
        if ($rejectedDifferentDays -gt 0) {
            Write-Host "  • $rejectedDifferentDays candidats rejetés (jours différents)" -ForegroundColor Red
        }
        return $confirmed
    }
    
    Write-Host "→ Validation avancée avec hash partiel intelligent..." -ForegroundColor Cyan
    
    # Extraction des fichiers uniques à valider
    $uniquePaths = @{}
    foreach ($candidate in $Candidates) {
        $uniquePaths[$candidate.Entry1.Chemin] = $candidate.Entry1
        $uniquePaths[$candidate.Entry2.Chemin] = $candidate.Entry2
    }
    
    Write-Host "  • $($uniquePaths.Count) fichiers uniques à valider" -ForegroundColor Gray
    
    # Cache de résolution et hash
    $pathCache = @{}
    $hashCache = @{}
    $pathsArray = $uniquePaths.Keys
    $resolvedCount = 0
    $hashCount = 0
    
    # Phase 1: Résolution des chemins
    Write-Host "  • Résolution des chemins..." -ForegroundColor Gray
    for ($i = 0; $i -lt $pathsArray.Count; $i += $HashBatchSize) {
        $batch = $pathsArray[$i..([Math]::Min($i + $HashBatchSize - 1, $pathsArray.Count - 1))]
        
        foreach ($path in $batch) {
            $resolved = Resolve-FilePath -OriginalPath $path -ServiceName $ServiceName -PathCache $pathCache -Verbose:($resolvedCount -lt 5)
            if ($resolved) { $resolvedCount++ }
        }
        
        $currentProgress = $i + $HashBatchSize
        if ($currentProgress -gt $pathsArray.Count) { $currentProgress = $pathsArray.Count }
        $progress = [Math]::Min([Math]::Round(($currentProgress * 100.0) / $pathsArray.Count, 1), 100)
        
        Write-Progress -Activity "Résolution chemins" -Status "$resolvedCount/$($pathsArray.Count) résolus" -PercentComplete $progress
        
        if ($i % ($HashBatchSize * 10) -eq 0) { [System.GC]::Collect() }
    }
    
    Write-Progress -Activity "Résolution chemins" -Completed
    Write-Host "    $resolvedCount/$($pathsArray.Count) chemins résolus" -ForegroundColor $(if($resolvedCount -gt ($pathsArray.Count * 0.3)){'Green'}else{'Yellow'})
    
    if ($resolvedCount -eq 0) {
        Write-Host "   Aucun fichier accessible - Validation par métadonnées uniquement" -ForegroundColor Yellow
        return Validate-DuplicatesAdvanced -Candidates $Candidates -ServiceName $ServiceName -SkipHashValidation:$true -HashBatchSize $HashBatchSize -PartialHashSize $PartialHashSize -SmallFileThreshold $SmallFileThreshold
    }
    
    # Phase 2: Calcul des hash adaptatifs
    Write-Host "  • Calcul des hash adaptatifs (partiel/complet selon taille)..." -ForegroundColor Gray
    $resolvedPaths = $pathCache.GetEnumerator() | Where-Object { $_.Value } | ForEach-Object { $_.Key }
    
    for ($i = 0; $i -lt $resolvedPaths.Count; $i += $HashBatchSize) {
        $batch = $resolvedPaths[$i..([Math]::Min($i + $HashBatchSize - 1, $resolvedPaths.Count - 1))]
        
        foreach ($originalPath in $batch) {
            $resolvedPath = $pathCache[$originalPath]
            try {
                $fileInfo = Get-Item -LiteralPath $resolvedPath -ErrorAction Stop
                $fileSize = $fileInfo.Length
                
                # Stratégie adaptative selon la taille
                if ($fileSize -le $SmallFileThreshold) {
                    # Petits fichiers: hash complet
                    $hash = (Get-FileHash -LiteralPath $resolvedPath -Algorithm MD5 -ErrorAction Stop).Hash
                    $hashType = "Complet"
                } else {
                    # Gros fichiers: hash partiel
                    $hash = Get-PartialFileHash -FilePath $resolvedPath -PartialSize $PartialHashSize
                    $hashType = "Partiel"
                }
                
                if ($hash) {
                    $hashCache[$originalPath] = @{
                        Hash = $hash
                        Type = $hashType
                        Size = $fileSize
                    }
                    $hashCount++
                }
            } catch {
                $hashCache[$originalPath] = $null
            }
        }
        
        $currentProgress = $i + $HashBatchSize
        if ($currentProgress -gt $resolvedPaths.Count) { $currentProgress = $resolvedPaths.Count }
        $progress = [Math]::Min([Math]::Round(($currentProgress * 100.0) / $resolvedPaths.Count, 1), 100)
        
        Write-Progress -Activity "Calcul hash adaptatif" -Status "$hashCount/$($resolvedPaths.Count) hash calculés" -PercentComplete $progress
        
        if ($i % ($HashBatchSize * 5) -eq 0) { [System.GC]::Collect() }
    }
    
    Write-Progress -Activity "Calcul hash adaptatif" -Completed
    Write-Host "    → $hashCount hash calculés" -ForegroundColor Green
    
    # Statistiques des types de hash
    $hashStats = $hashCache.Values | Where-Object { $_ } | Group-Object Type
    foreach ($stat in $hashStats) {
        Write-Host "      - Hash $($stat.Name): $($stat.Count) fichiers" -ForegroundColor Gray
    }
    
    # Phase 3: Validation finale avec niveaux de confiance ET critère même jour
    $confirmed = @()
    $validationStats = @{
        'HashMatch' = 0
        'MetadataOnly' = 0
        'Skipped' = 0
        'RejectedDifferentDays' = 0
    }
    
    foreach ($candidate in $Candidates) {
        # VÉRIFICATION PRÉALABLE : Critère même jour OBLIGATOIRE
        if ($UseExtendedMetadata -and $candidate.DaysDiff -ne $null -and $candidate.DaysDiff -gt 0) {
            $validationStats['RejectedDifferentDays']++
            continue  # REJET IMMÉDIAT pour jours différents
        }
        
        $hash1Info = $hashCache[$candidate.Entry1.Chemin]
        $hash2Info = $hashCache[$candidate.Entry2.Chemin]
        
        if ($hash1Info -and $hash2Info -and $hash1Info.Hash -eq $hash2Info.Hash) {
            # Hash identiques = doublon confirmé (ET même jour déjà vérifié)
            $candidate.Hash = $hash1Info.Hash
            $candidate.HashType = $hash1Info.Type
            $candidate.ValidationMethod = "Hash"
            $confirmed += $candidate
            $validationStats['HashMatch']++
        } elseif (-not $hash1Info -or -not $hash2Info) {
            # Pas de hash disponible, validation par métadonnées 
            $metadataValid = $candidate.Confidence -in @("Haute", "Très Haute", "Moyenne")
            
            if ($metadataValid) {
                $candidate.ValidationMethod = "Metadata"
                $confirmed += $candidate
                $validationStats['MetadataOnly']++
            } else {
                $validationStats['Skipped']++
            }
        } else {
            # Hash différents = pas de doublon
            $validationStats['Skipped']++
        }
    }
    
    # Rapport de validation
    Write-Host "  • Validation terminée:" -ForegroundColor Green
    Write-Host "    - Confirmés par hash: $($validationStats['HashMatch'])" -ForegroundColor Green
    Write-Host "    - Confirmés par métadonnées: $($validationStats['MetadataOnly'])" -ForegroundColor Yellow
    Write-Host "    - Rejetés (jours différents): $($validationStats['RejectedDifferentDays'])" -ForegroundColor Red
    Write-Host "    - Rejetés (autres raisons): $($validationStats['Skipped'])" -ForegroundColor Red
    
    return $confirmed
}
#endregion

#region Script Principal 
Write-Host " DÉTECTEUR DE DOUBLONS AVANCÉ " -ForegroundColor Magenta
Write-Host " Hash partiel intelligent + Métadonnées étendues" -ForegroundColor Gray
Write-Host " Hash partiel: $PartialHashSize bytes | Seuil petits fichiers: $([Math]::Round($SmallFileThreshold/1MB,1))MB | Critère: MÊME JOUR obligatoire" -ForegroundColor Gray

# Validation des paramètres et test de connectivité
$svcPath = Join-Path $PilotageRoot $ServiceName
if (-not (Test-Path $svcPath)) { 
    Write-Error "Service introuvable : $svcPath"
    exit 1 
}

# Test de connectivité réseau avant de commencer
if (-not $SkipHashValidation) {
    $networkTests = Test-NetworkConnectivity -ServiceName $ServiceName
    $hasNetworkAccess = $networkTests.Values -contains $true

    if (-not $hasNetworkAccess) {
        Write-Warning " Aucun accès réseau détecté. Basculement en mode métadonnées étendues."
        Write-Host "   ATTENTION: Mode métadonnées étendues REQUIS pour critère 'même jour'" -ForegroundColor Yellow
        $SkipHashValidation = $true
        $UseExtendedMetadata = $true
    }
}

# FORCAGE: Si critère "même jour" requis, métadonnées étendues OBLIGATOIRES
if (-not $UseExtendedMetadata) {
    Write-Host "  ℹ Activation OBLIGATOIRE des métadonnées étendues (critère 'même jour')" -ForegroundColor Cyan
    $UseExtendedMetadata = $true
}

# Vérification que les CSV contiennent bien des dates
Write-Host "   IMPORTANT: Vérification que les CSV contiennent des colonnes de dates..." -ForegroundColor Yellow

Write-Host " Recherche des CSV de volumétrie dans : $svcPath" -ForegroundColor Cyan
$csvs = Get-ChildItem $svcPath -Filter 'DonneeVolumetrique_*.csv' -File |
    Where-Object Name -notmatch 'Doublons'

if (-not $csvs) {
    Write-Warning "Aucun fichier CSV de volumétrie trouvé."
    exit 0
}

Write-Host "  • $($csvs.Count) fichiers CSV trouvés" -ForegroundColor Green

# PHASE 1: Lecture streaming avec métadonnées étendues
Write-Host " PHASE 1: Lecture streaming avec métadonnées étendues..." -ForegroundColor Cyan
$allEntries = @{}
$totalFiles = 0

foreach ($csv in $csvs) {
    $csvEntries = Read-CsvStreaming -CsvPath $csv.FullName -BufferSize $StreamBufferSize -ServiceName $ServiceName -IncludeExtendedMetadata:$UseExtendedMetadata
    
    foreach ($entry in $csvEntries.GetEnumerator()) {
        if (-not $allEntries.ContainsKey($entry.Key)) {
            $allEntries[$entry.Key] = $entry.Value
            $totalFiles++
        }
    }
    
    $csvEntries.Clear()
    $csvEntries = $null
    [System.GC]::Collect()
}

Write-Host "  • TOTAL: $totalFiles fichiers uniques collectés" -ForegroundColor Green

if ($totalFiles -eq 0) {
    Write-Warning "Aucune entrée valide trouvée."
    exit 0
}

# PHASE 2: Recherche avec métadonnées étendues
$candidates = Find-DuplicatesUltraFast -AllEntries $allEntries -NameThreshold $NameThreshold -MaxSizeDiff $MaxSizeDiff -UseExtendedMetadata:$UseExtendedMetadata

# Libération immédiate de la mémoire
$allEntries.Clear()
$allEntries = $null
[System.GC]::Collect()

if ($candidates.Count -eq 0) {
    Write-Host " Aucun doublon potentiel trouvé." -ForegroundColor Yellow
    exit 0
}

# PHASE 3: Validation avancée
$confirmedDuplicates = Validate-DuplicatesAdvanced -Candidates $candidates -ServiceName $ServiceName -SkipHashValidation:$SkipHashValidation -HashBatchSize $HashBatchSize -PartialHashSize $PartialHashSize -SmallFileThreshold $SmallFileThreshold

Write-Host "  • $($confirmedDuplicates.Count) doublons confirmés au total" -ForegroundColor Green

# PHASE 4: Export des résultats avec informations étendues
$outputDoublons = Join-Path $svcPath 'DonneeVolumetrique_ConfirmedDoublons_Advanced.csv'
Write-Host "→ Export des doublons avec métadonnées étendues..." -ForegroundColor Cyan

$headerLine = '"Nom1","Nom2","DistLevenshtein","SizeDiffMo","Confidence","ValidationMethod","TimeDiffSeconds","DaysDiff","HashType","Hash","Chemin1","Chemin2"'
$headerLine | Set-Content $outputDoublons -Encoding UTF8

$exportCount = 0
foreach ($duplicate in $confirmedDuplicates) {
    # VÉRIFICATION DEBUG: Alerte si des jours différents passent quand même
    if ($duplicate.DaysDiff -ne $null -and $duplicate.DaysDiff -gt 0) {
        Write-Warning "⚠ ERREUR DÉTECTÉE: Doublon avec jours différents confirmé - $($duplicate.Entry1.Nom) (écart: $($duplicate.DaysDiff) jours)"
        Write-Host "   Chemin1: $($duplicate.Entry1.Chemin)" -ForegroundColor Gray
        Write-Host "   Chemin2: $($duplicate.Entry2.Chemin)" -ForegroundColor Gray
        Write-Host "   Méthode validation: $($duplicate.ValidationMethod)" -ForegroundColor Gray
        continue  # Ne pas exporter ce doublon erroné
    }
    
    $hashValue = if ($duplicate.Hash) { $duplicate.Hash } else { "NO-HASH" }
    $hashType = if ($duplicate.HashType) { $duplicate.HashType } else { "None" }
    $validationMethod = if ($duplicate.ValidationMethod) { $duplicate.ValidationMethod } else { "Legacy" }
    $confidence = if ($duplicate.Confidence) { $duplicate.Confidence } else { "Standard" }
    $timeDiff = if ($duplicate.TimeDiff -ne $null) { [Math]::Round($duplicate.TimeDiff, 2) } else { "" }
    $daysDiff = if ($duplicate.DaysDiff -ne $null) { [Math]::Round($duplicate.DaysDiff, 0) } else { "" }
    
    $nom1Escaped = '"' + $duplicate.Entry1.Nom.Replace('"', '""') + '"'
    $nom2Escaped = '"' + $duplicate.Entry2.Nom.Replace('"', '""') + '"'
    $chemin1Escaped = '"' + $duplicate.Entry1.Chemin.Replace('"', '""') + '"'
    $chemin2Escaped = '"' + $duplicate.Entry2.Chemin.Replace('"', '""') + '"'
    $hashEscaped = '"' + $hashValue.Replace('"', '""') + '"'
    
    $line = "$nom1Escaped,$nom2Escaped,$($duplicate.Distance),$([Math]::Round($duplicate.SizeDiff,3)),""$confidence"",""$validationMethod"",$timeDiff,$daysDiff,""$hashType"",$hashEscaped,$chemin1Escaped,$chemin2Escaped"
    $line | Out-File $outputDoublons -Append -Encoding UTF8
    
    $exportCount++
    if ($exportCount % 1000 -eq 0) {
        Write-Progress -Activity "Export doublons" -Status "$exportCount/$($confirmedDuplicates.Count)" -PercentComplete (($exportCount * 100) / $confirmedDuplicates.Count)
        [System.GC]::Collect()
    }
}

Write-Progress -Activity "Export doublons" -Completed

# PHASE 5: Consolidation ultra-rapide avec statistiques avancées
if (-not $NoConsolidation -and $confirmedDuplicates.Count -gt 0) {
    Write-Host "→ PHASE 5: Consolidation avancée avec analyse d'impact..." -ForegroundColor Cyan
    
    $groupes = @{}
    foreach ($duplicate in $confirmedDuplicates) {
        $nom = $duplicate.Entry1.Nom
        $hash = if ($duplicate.Hash) { $duplicate.Hash } else { "NOHASH" }
        $cle = "$nom|$hash"
        
        if (-not $groupes.ContainsKey($cle)) {
            $groupes[$cle] = @{
                Nom = $nom
                Hash = $hash
                Chemins = @{}
                Taille = $duplicate.Entry1.Taille
                ValidationMethods = @{}
                Confidences = @{}
                HashTypes = @{}
            }
        }
        
        # Accumulation des informations
        $groupes[$cle].Chemins[$duplicate.Entry1.Chemin] = $true
        $groupes[$cle].Chemins[$duplicate.Entry2.Chemin] = $true
        
        # Statistiques de validation
        $validationMethod = if ($duplicate.ValidationMethod) { $duplicate.ValidationMethod } else { "Legacy" }
        $confidence = if ($duplicate.Confidence) { $duplicate.Confidence } else { "Standard" }
        $hashType = if ($duplicate.HashType) { $duplicate.HashType } else { "None" }
        
        if (-not $groupes[$cle].ValidationMethods.ContainsKey($validationMethod)) { 
            $groupes[$cle].ValidationMethods[$validationMethod] = 0 
        }
        $groupes[$cle].ValidationMethods[$validationMethod] += 1
        
        if (-not $groupes[$cle].Confidences.ContainsKey($confidence)) { 
            $groupes[$cle].Confidences[$confidence] = 0 
        }
        $groupes[$cle].Confidences[$confidence] += 1
        
        if (-not $groupes[$cle].HashTypes.ContainsKey($hashType)) { 
            $groupes[$cle].HashTypes[$hashType] = 0 
        }
        $groupes[$cle].HashTypes[$hashType] += 1
    }
    
    # Génération rapport consolidé avancé
    $rapportConsolide = Join-Path $svcPath "Doublons_Consolides_Advanced_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $consolidatedHeader = '"NomFichier","NombreEmplacements","TailleMB","EspaceEconomisableMB","MethodeValidation","NiveauConfiance","TypeHash","Hash","ListeChemins"'
    $consolidatedHeader | Set-Content $rapportConsolide -Encoding UTF8
    
    $totalEspaceEconomisable = 0
    $consolidationCount = 0
    $stats = @{
        'TotalGroups' = 0
        'HashValidated' = 0
        'MetadataValidated' = 0
        'HighConfidence' = 0
        'TotalSpaceSaved' = 0
    }
    
    foreach ($groupe in ($groupes.Values | Sort-Object { $_.Taille * ($_.Chemins.Keys.Count - 1) } -Descending)) {
        if ($groupe.Chemins.Keys.Count -lt 2) { continue }
        
        $nbEmplacements = $groupe.Chemins.Keys.Count
        $espaceEconomisable = [Math]::Round($groupe.Taille * ($nbEmplacements - 1), 2)
        $totalEspaceEconomisable += $espaceEconomisable
        
        # Méthode de validation dominante
        $mainValidationMethod = ($groupe.ValidationMethods.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
        $mainConfidence = ($groupe.Confidences.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key  
        $mainHashType = ($groupe.HashTypes.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
        
        $nomEscape = '"' + $groupe.Nom.Replace('"', '""') + '"'
        $hashEscape = '"' + $groupe.Hash.Replace('"', '""') + '"'
        $cheminsEscape = '"' + (($groupe.Chemins.Keys -join '; ').Replace('"', '""')) + '"'
        $methodEscape = '"' + $mainValidationMethod + '"'
        $confidenceEscape = '"' + $mainConfidence + '"'
        $hashTypeEscape = '"' + $mainHashType + '"'
        
        $ligne = "$nomEscape,$nbEmplacements,$($groupe.Taille),$espaceEconomisable,$methodEscape,$confidenceEscape,$hashTypeEscape,$hashEscape,$cheminsEscape"
        $ligne | Out-File $rapportConsolide -Append -Encoding UTF8
        
        # Statistiques
        $stats['TotalGroups']++
        if ($mainValidationMethod -eq "Hash") { $stats['HashValidated']++ }
        if ($mainValidationMethod -eq "Metadata") { $stats['MetadataValidated']++ }
        if ($mainConfidence -in @("Haute", "Très Haute")) { $stats['HighConfidence']++ }
        $stats['TotalSpaceSaved'] += $espaceEconomisable
        
        $consolidationCount++
        if ($consolidationCount % 500 -eq 0) { [System.GC]::Collect() }
    }
    
    Write-Host "  • Rapport consolidé généré : $([System.IO.Path]::GetFileName($rapportConsolide))" -ForegroundColor Green
    Write-Host "  • Statistiques de validation:" -ForegroundColor Cyan
    Write-Host "    - Validés par hash: $($stats['HashValidated'])/$($stats['TotalGroups'])" -ForegroundColor Green
    Write-Host "    - Validés par métadonnées: $($stats['MetadataValidated'])/$($stats['TotalGroups'])" -ForegroundColor Yellow  
    Write-Host "    - Haute confiance: $($stats['HighConfidence'])/$($stats['TotalGroups'])" -ForegroundColor Green
    Write-Host "  • Espace économisable total : $([Math]::Round($totalEspaceEconomisable, 2)) MB" -ForegroundColor Cyan
}

# Nettoyage final complet
$candidates = $null
$confirmedDuplicates = $null
$groupes = $null
[System.GC]::Collect()

# RÉSUMÉ FINAL AVANCÉ
Write-Host "`n RÉSUMÉ DÉTECTION AVANCÉE " -ForegroundColor Magenta
Write-Host "  • Fichiers analysés        : $totalFiles" -ForegroundColor White
Write-Host "  • Doublons confirmés       : $exportCount" -ForegroundColor Green

# Statistiques de validation
if ($exportCount -gt 0) {
    $validationStats = $confirmedDuplicates | Group-Object ValidationMethod
    foreach ($stat in $validationStats) {
        $method = if ($stat.Name) { $stat.Name } else { "Legacy" }
        Write-Host "    - Par $method : $($stat.Count)" -ForegroundColor Gray
    }
    
    $confidenceStats = $confirmedDuplicates | Group-Object Confidence  
    foreach ($stat in $confidenceStats) {
        $conf = if ($stat.Name) { $stat.Name } else { "Standard" }
        Write-Host "    - Confiance $conf : $($stat.Count)" -ForegroundColor Gray
    }
}

Write-Host "  • Mode de fonctionnement   : $(if($SkipHashValidation){'Métadonnées étendues uniquement'}else{'Hash partiel intelligent'})" -ForegroundColor Yellow
Write-Host "  • Fichier doublons         : $([System.IO.Path]::GetFileName($outputDoublons))" -ForegroundColor Cyan
if (-not $NoConsolidation -and $exportCount -gt 0) {
    Write-Host "  • Rapport consolidé        : $([System.IO.Path]::GetFileName($rapportConsolide))" -ForegroundColor Cyan
}

if ($exportCount -gt 0) {
    Write-Host "`n DÉTECTION AVANCÉE TERMINÉE AVEC SUCCÈS !" -ForegroundColor Green
    Write-Host "   Fiabilité optimisée par hash partiel + métadonnées étendues" -ForegroundColor Gray
    
    if (-not $SkipHashValidation -and $stats) {
        $hashValidationRate = [Math]::Round(($stats['HashValidated'] * 100.0) / $stats['TotalGroups'], 1)
        Write-Host "   Taux de validation par hash: $hashValidationRate%" -ForegroundColor Green
    }
} else {
    Write-Host "`n→ Aucun doublon trouvé." -ForegroundColor Yellow
}
#endregion

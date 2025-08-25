[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ServiceName,
    [string]$PilotageRoot = 'a remplir',
    [string]$DoublonsFile = '',
    [ValidateSet('Auto', 'Manual', 'Preview')][string]$Mode = 'Preview',
    [switch]$CreateBackup,
    [switch]$UseConsolidated
)

# Fonction pour créer un raccourci Windows robuste
function New-Shortcut {
    param(
        [string]$ShortcutPath,
        [string]$TargetPath,
        [string]$Description = "Raccourci vers fichier dédupliqué"
    )
    
    try {
        # Vérification que le fichier cible existe
        if (-not (Test-Path -LiteralPath $TargetPath -PathType Leaf)) {
            Write-Warning "Fichier cible inexistant: $TargetPath"
            return $false
        }
        
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
        $Shortcut.TargetPath = $TargetPath
        $Shortcut.Description = $Description
        $Shortcut.WorkingDirectory = Split-Path $TargetPath -Parent
        $Shortcut.Save()
        
        # Vérification que le raccourci a été créé
        Start-Sleep -Milliseconds 500  # Attendre que Windows finalise
        $shortcutCreated = Test-Path -LiteralPath $ShortcutPath
        
        # Libération COM
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Shortcut) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WScriptShell) | Out-Null
        
        if ($shortcutCreated) {
            Write-Host "      → Raccourci créé: $([System.IO.Path]::GetFileName($ShortcutPath))" -ForegroundColor Gray
            return $true
        } else {
            Write-Warning "Le raccourci n'a pas été créé correctement"
            return $false
        }
    }
    catch {
        Write-Warning "Erreur création raccourci $ShortcutPath : $($_.Exception.Message)"
        return $false
    }
}

# Fonction pour choisir le meilleur fichier parmi TOUS les emplacements
function Choose-BestFile {
    param($Chemins, $NomFichier)
    
    $scores = @{}
    $infos = @{}
    $cheminsCorrigés = @{}
    
    # Analyse de chaque chemin avec correction automatique
    foreach ($chemin in $Chemins) {
        $score = 0
        $cheminFinal = $chemin
        
        try {
            # Correction automatique du chemin si nécessaire
            if (-not (Test-Path -LiteralPath $chemin -PathType Leaf)) {
                $cheminAvecFichier = Join-Path $chemin $NomFichier
                if (Test-Path -LiteralPath $cheminAvecFichier -PathType Leaf) {
                    $cheminFinal = $cheminAvecFichier
                }
            }
            
            # Sauvegarde du chemin corrigé
            $cheminsCorrigés[$chemin] = $cheminFinal
            
            if (Test-Path -LiteralPath $cheminFinal -PathType Leaf) {
                $info = Get-Item -LiteralPath $cheminFinal
                $infos[$chemin] = $info
                
                # Critère 1: Chemin plus court = mieux organisé (+5 points)
                $score += (1000 - $chemin.Length) * 0.005
                
                # Critère 2: Éviter les mots suspects dans le chemin (-20 points)
                $badPatterns = @('backup', 'sauvegarde', 'temp', 'temporaire', 'old', 'ancien', 'archive', 'copie', 'copy', 'bak')
                foreach ($pattern in $badPatterns) {
                    if ($chemin -match $pattern) { $score -= 20 }
                }
                
                # Critère 3: Préférer les fichiers plus récents (+10 points max)
                $ageInDays = ((Get-Date) - $info.LastWriteTime).TotalDays
                if ($ageInDays -lt 365) { $score += (365 - $ageInDays) / 36.5 }
                
                # Critère 4: Éviter les chemins avec trop d'espaces ou caractères spéciaux
                $specialCharsCount = ($chemin -split '[^a-zA-Z0-9\\._-]').Count - 1
                $score -= $specialCharsCount * 0.5
                
                # Critère 5: Préférer les dossiers "principaux" (+10 points)
                if ($chemin -match '\\(Documents|Projets|Principal|Main|Original)\\') { $score += 10 }
                
                # Critère 6: Éviter les dossiers profonds (-1 point par niveau après 5)
                $depth = ($chemin -split '\\').Count
                if ($depth -gt 5) { $score -= ($depth - 5) }
                
                $scores[$chemin] = $score
            } else {
                $scores[$chemin] = -1000  # Fichier inaccessible
                $cheminsCorrigés[$chemin] = $chemin  # Pas de correction possible
            }
        }
        catch {
            $scores[$chemin] = -1000
            $cheminsCorrigés[$chemin] = $chemin
        }
    }
    
    # NOUVELLE LOGIQUE: Vérifier s'il y a le moindre -1000
    $fichiersInaccessibles = $scores.Values | Where-Object { $_ -eq -1000 }
    
    if ($fichiersInaccessibles.Count -gt 0) {
        # Si il y a AU MOINS UN fichier inaccessible, on ignore le groupe entier
        return @{
            KeepFile = $null
            RemoveFiles = @()
            Scores = $scores
            CheminsCorrigés = $cheminsCorrigés
            Reason = "Groupe entièrement ignoré - $($fichiersInaccessibles.Count) fichier(s) inaccessible(s) détecté(s)"
            ShouldSkip = $true
        }
    }
    
    # Si tous les fichiers sont accessibles, continuer normalement
    $bestPath = $scores.Keys | Sort-Object { $scores[$_] } -Descending | Select-Object -First 1
    $bestPathCorrigé = $cheminsCorrigés[$bestPath]
    
    # Récupération des autres chemins accessibles avec leurs corrections
    $othersToRemove = @()
    foreach ($chemin in $Chemins) {
        if ($chemin -ne $bestPath) {
            $othersToRemove += $cheminsCorrigés[$chemin]
        }
    }
    
    return @{
        KeepFile = $bestPathCorrigé
        RemoveFiles = $othersToRemove
        Scores = $scores
        CheminsCorrigés = $cheminsCorrigés
        Reason = "Meilleur score: $([Math]::Round($scores[$bestPath], 2)) - Tous fichiers accessibles"
        ShouldSkip = $false
    }
}

# Fonction de consolidation interne
function ConsolidateDuplicates {
    param($DoublonsList)
    
    Write-Host "→ Consolidation des doublons par fichier unique..." -ForegroundColor Cyan
    $groupes = @{}
    
    foreach ($doublon in $DoublonsList) {
        $cle = "$($doublon.Nom1)::$($doublon.Hash)"
        
        if (-not $groupes.ContainsKey($cle)) {
            $groupes[$cle] = @{
                Nom = $doublon.Nom1
                Hash = $doublon.Hash
                Chemins = New-Object System.Collections.ArrayList
            }
        }
        
        # Ajout des chemins uniques
        if ($doublon.Chemin1 -and $groupes[$cle].Chemins -notcontains $doublon.Chemin1) {
            [void]$groupes[$cle].Chemins.Add($doublon.Chemin1)
        }
        if ($doublon.Chemin2 -and $groupes[$cle].Chemins -notcontains $doublon.Chemin2) {
            [void]$groupes[$cle].Chemins.Add($doublon.Chemin2)
        }
    }
    
    # Filtrage des groupes avec au moins 2 fichiers
    $groupesValides = @{}
    foreach ($cle in $groupes.Keys) {
        if ($groupes[$cle].Chemins.Count -ge 2) {
            $groupesValides[$cle] = $groupes[$cle]
        }
    }
    
    Write-Host "  • $($groupesValides.Count) fichiers uniques avec doublons trouvés" -ForegroundColor Green
    return $groupesValides
}

# Validation des paramètres
$svcPath = Join-Path $PilotageRoot $ServiceName
if (-not (Test-Path $svcPath)) { 
    Write-Error "Service introuvable : $svcPath"
    exit 1 
}

# Détermination du fichier de doublons
if (-not $DoublonsFile) {
    $DoublonsFile = Join-Path $svcPath 'DonneeVolumetrique_ConfirmedDoublons.csv'
}

if (-not (Test-Path $DoublonsFile)) {
    Write-Error "Fichier de doublons introuvable : $DoublonsFile"
    Write-Host "Exécutez d'abord le script de détection de doublons." -ForegroundColor Yellow
    exit 1
}

# Lecture du fichier de doublons
Write-Host "→ Lecture du fichier de doublons : $DoublonsFile" -ForegroundColor Cyan
try {
    $doublons = Import-Csv -Path $DoublonsFile -Encoding UTF8
    Write-Host "  • $($doublons.Count) paires de doublons lues" -ForegroundColor Green
}
catch {
    Write-Error "Erreur lors de la lecture du fichier CSV : $($_.Exception.Message)"
    exit 1
}

if ($doublons.Count -eq 0) {
    Write-Host "→ Aucun doublon à traiter." -ForegroundColor Yellow
    exit 0
}

# Consolidation intelligente
$groupes = ConsolidateDuplicates -DoublonsList $doublons

# Préparation du rapport
$reportPath = Join-Path $svcPath "Gestion_Doublons_Intelligente_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$backupPath = Join-Path $svcPath "Backup_Metadata_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"

# Initialisation des statistiques
$stats = @{
    TotalGroups = $groupes.Count
    TotalFiles = ($groupes.Values | ForEach-Object { $_.Chemins.Count } | Measure-Object -Sum).Sum
    Processed = 0
    Success = 0
    Errors = 0
    FilesRemoved = 0
    SpaceSaved = 0
    Skipped = 0
    HashVerified = 0
    HashErrors = 0
}

$actions = @()
$errors = @()

# Analyse et traitement de chaque groupe de fichiers identiques
Write-Host "→ Analyse intelligente des doublons en mode $Mode..." -ForegroundColor Cyan

foreach ($cle in $groupes.Keys) {
    $groupe = $groupes[$cle]
    $stats.Processed++
    
    Write-Progress -Activity "Traitement intelligent des doublons" -Status "$($stats.Processed)/$($stats.TotalGroups) groupes" -PercentComplete (($stats.Processed / $stats.TotalGroups) * 100)
    
    $cheminsArray = $groupe.Chemins.ToArray()
    
    # Choix intelligent du meilleur fichier
    $decision = Choose-BestFile -Chemins $cheminsArray -NomFichier $groupe.Nom
    
    # NOUVELLE VÉRIFICATION: Si le groupe doit être ignoré
    if ($decision.ShouldSkip) {
        Write-Host "  [$($stats.Processed)] IGNORÉ - $($groupe.Nom)" -ForegroundColor DarkYellow
        Write-Host "    Raison: $($decision.Reason)" -ForegroundColor Gray
        
        # Affichage des scores pour diagnostic
        Write-Host "    Diagnostic des emplacements:" -ForegroundColor Gray
        foreach ($chemin in $cheminsArray) {
            $score = $decision.Scores[$chemin]
            $status = if ($score -eq -1000) { "INACCESSIBLE" } else { "OK" }
            $color = if ($score -eq -1000) { 'Red' } else { 'Green' }
            Write-Host "      $status ($score) - $chemin" -ForegroundColor $color
        }
        
        $action = @{
            Index = $stats.Processed
            FileName = $groupe.Nom
            KeepFile = "N/A"
            RemoveFiles = @()
            RemoveCount = 0
            Reason = $decision.Reason
            SizeMB = 0
            Status = "Groupe ignoré"
            Scores = $decision.Scores
        }
        
        $actions += $action
        $stats.Skipped += $cheminsArray.Count
        continue
    }
    
    $keepFile = $decision.KeepFile
    $removeFiles = $decision.RemoveFiles
    $reason = $decision.Reason
    
    # Calcul de l'espace économisé
    $groupSpaceSaved = 0
    foreach ($removeFile in $removeFiles) {
        try {
            if (Test-Path -LiteralPath $removeFile -PathType Leaf) {
                $removeFileInfo = Get-Item -LiteralPath $removeFile
                $fileSizeMB = [Math]::Round($removeFileInfo.Length / 1MB, 2)
                $groupSpaceSaved += $fileSizeMB
            }
        }
        catch {
            continue
        }
    }
    $stats.SpaceSaved += $groupSpaceSaved
    
    $action = @{
        Index = $stats.Processed
        FileName = $groupe.Nom
        KeepFile = $keepFile
        RemoveFiles = $removeFiles
        RemoveCount = $removeFiles.Count
        Reason = $reason
        SizeMB = $groupSpaceSaved
        Status = "Planifié"
        Scores = $decision.Scores
    }
    
    if ($Mode -eq 'Preview') {
        Write-Host "  [$($stats.Processed)] PREVIEW - Fichier: $($groupe.Nom)" -ForegroundColor Yellow
        Write-Host "    Garderait  : $keepFile" -ForegroundColor Green
        Write-Host "    Supprimerait: $($removeFiles.Count) copies ($groupSpaceSaved MB)" -ForegroundColor Red
        Write-Host "    Raison     : $reason" -ForegroundColor Gray
        foreach ($removeFile in ($removeFiles | Select-Object -First 3)) {
            Write-Host "      - $removeFile" -ForegroundColor DarkRed
        }
        if ($removeFiles.Count -gt 3) {
            Write-Host "      ... et $($removeFiles.Count - 3) autres" -ForegroundColor DarkRed
        }
        $action.Status = "Preview"
        $stats.Skipped += $removeFiles.Count
    }
    else {
        # Mode Auto ou Manual
        $proceed = $true
        
        if ($Mode -eq 'Manual') {
            Write-Host "`n  [$($stats.Processed)] Groupe de doublons détecté:" -ForegroundColor Cyan
            Write-Host "    Fichier    : $($groupe.Nom)" -ForegroundColor White
            Write-Host "    Emplacements: $($cheminsArray.Count)" -ForegroundColor White
            Write-Host "    Recommandation: Garder '$keepFile'" -ForegroundColor Green
            Write-Host "    Supprimer  : $($removeFiles.Count) copies" -ForegroundColor Yellow
            Write-Host "    Espace économisé: $groupSpaceSaved MB" -ForegroundColor Yellow
            Write-Host "    Raison     : $reason" -ForegroundColor Gray
            
            # Affichage des scores pour info
            Write-Host "    Scores des emplacements:" -ForegroundColor Gray
            foreach ($chemin in ($decision.Scores.Keys | Sort-Object { $decision.Scores[$_] } -Descending)) {
                $color = if ($chemin -eq $keepFile) { 'Green' } else { 'Red' }
                Write-Host "      $([Math]::Round($decision.Scores[$chemin], 1)) - $chemin" -ForegroundColor $color
            }
            
            do {
                $choice = Read-Host "    Action: [O]ui / [N]on / [Q]uitter"
                switch ($choice.ToUpper()) {
                    'O' { $proceed = $true; break }
                    'N' { $proceed = $false; $stats.Skipped += $removeFiles.Count; break }
                    'Q' { 
                        Write-Host "Arrêt demandé par l'utilisateur." -ForegroundColor Yellow
                        break 2
                    }
                    default { Write-Host "Choix invalide. Utilisez O, N ou Q." -ForegroundColor Red }
                }
            } while ($choice -notmatch '^[ONQ]$')
        }
        
        if ($proceed) {
            $groupErrors = 0
            $groupSuccess = 0
            
            # VÉRIFICATION HASH OBLIGATOIRE EN MODE AUTO
            if ($Mode -eq 'Auto') {
                Write-Host "    → Vérification hash de sécurité..." -ForegroundColor Cyan
                
                # Calcul du hash du fichier à garder
                try {
                    $keepFileHash = (Get-FileHash -LiteralPath $keepFile -Algorithm MD5 -ErrorAction Stop).Hash
                } catch {
                    $errors += "Impossible de calculer le hash du fichier à garder: $keepFile - $($_.Exception.Message)"
                    $action.Status = "Erreur hash fichier à garder"
                    $stats.Errors++
                    continue
                }
                
                # Vérification hash de chaque fichier à supprimer
                $hashErrors = @()
                $hashValid = @()
                
                foreach ($removeFile in $removeFiles) {
                    try {
                        if (Test-Path -LiteralPath $removeFile -PathType Leaf) {
                            $removeFileHash = (Get-FileHash -LiteralPath $removeFile -Algorithm MD5 -ErrorAction Stop).Hash
                            
                            if ($removeFileHash -eq $keepFileHash) {
                                $hashValid += $removeFile
                                Write-Host "      ✓ Hash confirmé: $([System.IO.Path]::GetFileName($removeFile))" -ForegroundColor Green
                            } else {
                                $hashErrors += "Hash différent pour $removeFile (attendu: $keepFileHash, trouvé: $removeFileHash)"
                                Write-Host "      ✗ Hash différent: $([System.IO.Path]::GetFileName($removeFile))" -ForegroundColor Red
                            }
                        } else {
                            $hashErrors += "Fichier inaccessible pour vérification hash: $removeFile"
                        }
                    } catch {
                        $hashErrors += "Erreur calcul hash pour $removeFile : $($_.Exception.Message)"
                    }
                }
                
                # Si des erreurs de hash, abandon du groupe
                if ($hashErrors.Count -gt 0) {
                    $errors += $hashErrors
                    $action.Status = "Erreur vérification hash - Groupe ignoré par sécurité"
                    $stats.Errors++
                    $stats.HashErrors += $hashErrors.Count
                    Write-Host "    ⚠ SÉCURITÉ: Groupe ignoré à cause d'erreurs de hash" -ForegroundColor Red
                    foreach ($error in $hashErrors) {
                        Write-Host "      - $error" -ForegroundColor Red
                    } 
                    continue
                }
                
                # Mise à jour de la liste : seuls les fichiers avec hash confirmé
                $removeFiles = $hashValid
                $stats.HashVerified += $hashValid.Count
                Write-Host "    ✅ Tous les hash confirmés - Traitement sécurisé autorisé" -ForegroundColor Green
            }
            
            # Traitement des fichiers (après vérification hash si mode Auto)
            foreach ($removeFile in $removeFiles) {
                try {
                    # Vérification finale de l'existence
                    if (-not (Test-Path -LiteralPath $keepFile -PathType Leaf)) {
                        $errors += "Fichier à garder inaccessible: $keepFile"
                        $groupErrors++
                        continue
                    }
                    
                    if (-not (Test-Path -LiteralPath $removeFile -PathType Leaf)) {
                        $errors += "Fichier à supprimer déjà absent: $removeFile"
                        continue
                    }
                    
                    Write-Host "      → Traitement: $([System.IO.Path]::GetFileName($removeFile))" -ForegroundColor Yellow
                    
                    # Nom du raccourci (même nom + .lnk)
                    $shortcutPath = "$removeFile.lnk"
                    
                    # Étape 1: Créer le raccourci Windows standard
                    $shortcutCreated = New-Shortcut -ShortcutPath $shortcutPath -TargetPath $keepFile
                    
                    if ($shortcutCreated) {
                        # Étape 2: Supprimer le fichier original
                        Remove-Item -LiteralPath $removeFile -Force
                        Write-Host "      → Fichier original supprimé" -ForegroundColor Gray
                        Write-Host "      ✓ Raccourci créé: $([System.IO.Path]::GetFileName($shortcutPath))" -ForegroundColor Green
                        
                        $groupSuccess++
                        $stats.FilesRemoved++
                        
                    } else {
                        $errors += "Impossible de créer le raccourci pour: $removeFile"
                        $groupErrors++
                    }
                    
                }
                catch {
                    $error = "Erreur lors du traitement de $removeFile : $($_.Exception.Message)"
                    $errors += $error
                    $groupErrors++
                    Write-Host "      ✗ $error" -ForegroundColor Red
                }
            }
            
            if ($groupErrors -eq 0) {
                $action.Status = "Succès total"
                $stats.Success++
                Write-Host "  ✓ [$($stats.Processed)] Traité: $($groupe.Nom) - $groupSuccess fichiers → raccourcis ($groupSpaceSaved MB économisés)" -ForegroundColor Green
            } elseif ($groupSuccess -gt 0) {
                $action.Status = "Succès partiel ($groupSuccess/$($removeFiles.Count))"
                $stats.Success++
                Write-Host "  ⚠ [$($stats.Processed)] Traité partiellement: $($groupe.Nom) - $groupSuccess/$($removeFiles.Count) fichiers traités" -ForegroundColor Yellow
            } else {
                $action.Status = "Échec total"
                $stats.Errors++
                Write-Host "  ✗ [$($stats.Processed)] Échec: $($groupe.Nom) - Aucun fichier traité" -ForegroundColor Red
            }
        } else {
            $action.Status = "Ignoré"
        }
    }
    
    $actions += $action
}

Write-Progress -Activity "Traitement intelligent des doublons" -Completed

# Génération du rapport
Write-Host "`n→ Génération du rapport..." -ForegroundColor Cyan
$report = @"
=== RAPPORT DE GESTION INTELLIGENTE DES DOUBLONS ===
Service: $ServiceName
Date: $(Get-Date)
Mode: $Mode

=== STATISTIQUES ===
Groupes de fichiers traités: $($stats.TotalGroups)
Fichiers analysés au total: $($stats.TotalFiles)
Groupes traités avec succès: $($stats.Success)
Erreurs de groupes: $($stats.Errors)
Fichiers individuels supprimés: $($stats.FilesRemoved)
Fichiers ignorés: $($stats.Skipped)
Espace économisé total: $([Math]::Round($stats.SpaceSaved, 2)) MB

=== DÉTAILS DES ACTIONS ===
$($actions | ForEach-Object { "[$($_.Index)] $($_.Status) - $($_.FileName): Gardé 1, Supprimé $($_.RemoveCount) ($($_.SizeMB) MB)" } | Out-String)

=== ERREURS ===
$($errors | ForEach-Object { "- $_" } | Out-String)
"@

$report | Set-Content $reportPath -Encoding UTF8

# Résumé final
Write-Host "`n→ RÉSUMÉ FINAL" -ForegroundColor Magenta
Write-Host "  • Groupes de fichiers   : $($stats.TotalGroups)" -ForegroundColor White
Write-Host "  • Fichiers analysés     : $($stats.TotalFiles)" -ForegroundColor White
Write-Host "  • Opérations réussies   : $($stats.Success)" -ForegroundColor Green
Write-Host "  • Erreurs               : $($stats.Errors)" -ForegroundColor $(if($stats.Errors -gt 0){'Red'}else{'White'})
Write-Host "  • Fichiers supprimés    : $($stats.FilesRemoved)" -ForegroundColor Yellow
Write-Host "  • Fichiers ignorés      : $($stats.Skipped)" -ForegroundColor Gray
if ($Mode -eq 'Auto') {
    Write-Host "  • Hash vérifiés (sécurité): $($stats.HashVerified)" -ForegroundColor Green
    Write-Host "  • Erreurs de hash       : $($stats.HashErrors)" -ForegroundColor $(if($stats.HashErrors -gt 0){'Red'}else{'Green'})
}
Write-Host "  • Espace économisé      : $([Math]::Round($stats.SpaceSaved, 2)) MB" -ForegroundColor Cyan
Write-Host "  • Rapport généré        : $reportPath" -ForegroundColor Gray

if ($Mode -eq 'Preview') {
    Write-Host "`n→ Mode PREVIEW - Aucune modification effectuée" -ForegroundColor Yellow
    Write-Host "  Relancez avec -Mode 'Auto' ou -Mode 'Manual' pour appliquer les changements." -ForegroundColor Yellow
} elseif ($stats.Success -gt 0) {
    Write-Host "`n→ Traitement terminé avec succès !" -ForegroundColor Green
    Write-Host "  Les fichiers supprimés ont été remplacés par des raccourcis transparents." -ForegroundColor Green
}

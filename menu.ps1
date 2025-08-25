# MenuLauncher.ps1

[string]$scriptVolPath = Join-Path $PSScriptRoot 'volumetriemenu.ps1'
[string]$scriptSec1   = Join-Path $PSScriptRoot 'compression image.ps1'
[string]$scriptSec2   = Join-Path $PSScriptRoot 'Detection doublon.ps1'
[string]$scriptSec3   = Join-Path $PSScriptRoot 'traitement de pdf.ps1'
[string]$scriptSec4   = Join-Path $PSScriptRoot 'Gestion doublons.ps1'

# Chemins communs
[string]$baseServicePath = 'a remplir'
[string]$pilotageRoot    = 'a remplir'

# Valeurs par défaut pour la compression d'images
[int]$defaultWidth   = 1224
[int]$defaultHeight  = 1632
[int]$defaultQuality = 75

function Prompt-Subfolders {
    param([string]$ScriptLabel)

    do {
        $input    = Read-Host "Entrez le(s) nom(s) de service pour '$ScriptLabel' (séparés par des virgules, la compression d'image ne peut traité uniquement 1 service à la fois)"
        $raw      = $input -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        $existent = @()
        $missing  = @()

        foreach ($svc in $raw) {
            if (Test-Path (Join-Path $baseServicePath $svc)) { $existent += $svc }
            else                                  { $missing  += $svc }
        }

        if ($missing) {
            Write-Warning "Ces services n'existent pas : $($missing -join ', ')"
            $retry = Read-Host 'Ressaisir la liste ? (O/N)'
        } else {
            $retry = 'N'
        }
    } while ($retry -match '^[Oo]')

    if (-not $existent) {
        Write-Warning "Aucun service valide pour '$ScriptLabel'."
        return $null
    }

    return $existent
}

function Launch-Menu {
    do {
        Clear-Host
            Write-Host ' '
            Write-Host ' '
            Write-Host ' '
            Write-Host ' '
            Write-Host ' '
            Write-Host ' '
            Write-Host ' '
            Write-Host '=== MENU PRINCIPAL ===' -ForegroundColor Cyan
            Write-Host '[1] Traitement volumétrie'
            Write-Host '[2] Compression des images'
            Write-Host '[3] Détection des doublons'
            Write-Host '[4] Gestion des doublons'
            Write-Host '[5] Compression PDF'
            Write-Host '[Q] Quitter'
        $choice = Read-Host 'Votre choix'

        switch ($choice) {
            '1' {
                $services = Prompt-Subfolders 'Traitement volumétrie'
                if ($services) {
                    Write-Host "→ Lancement volumétrie pour : $($services -join ', ')" -ForegroundColor Green
                    & $scriptVolPath -BaseServicePath $baseServicePath -SubFolders $services -PilotagePath $pilotageRoot
                }
                Read-Host 'Entrée pour revenir' | Out-Null
            }
            
            '2' {
                # On s'assure que $services est un tableau, même à 1 élément
                $services = @(Prompt-Subfolders 'Compression des images')
                
                if ($services.Count -eq 1) {
                    $svc     = $services[0]
                    $pathSvc = Join-Path $pilotageRoot $svc
                    Write-Host "→ Compression des images pour : $svc ($pathSvc)" -ForegroundColor Green
                    & $scriptSec1 `
                        -PilotagePath $pathSvc `
                        -TargetWidth  $defaultWidth `
                        -TargetHeight $defaultHeight `
                        -JpegQuality  $defaultQuality
                }
                else {
                    Write-Warning 'Veuillez sélectionner exactement UN service pour la compression des images.'
                }
                Read-Host 'Entrée pour revenir' | Out-Null
            }

            '3' {
                $services = Prompt-Subfolders 'Détection doublons'
                if ($services) {
                    foreach ($svc in $services) {
                        Write-Host "→ Détection doublons pour : $svc" -ForegroundColor Green
                        & $scriptSec2 -ServiceName $svc -PilotageRoot $pilotageRoot
                    }
                }
                Read-Host 'Entrée pour revenir' | Out-Null
            }

            '4' {
                $services = Prompt-Subfolders 'Gestion des doublons'
                if ($services) {
                    foreach ($svc in $services) {
                        Write-Host "`n→ Gestion des doublons pour : $svc" -ForegroundColor Green
                        
                        # Sous-menu pour le mode de gestion
                        do {
                            Write-Host "`n=== MODE DE GESTION DES DOUBLONS ===" -ForegroundColor Yellow
                            Write-Host "[P] Preview - Voir les actions sans les exécuter"
                            Write-Host "[A] Auto - Traitement automatisé"
                            Write-Host "[M] Manual - Validation manuelle de chaque doublon"
                            Write-Host "[R] Retour au menu principal"
                            $modeChoice = Read-Host "Mode de gestion"
                            
                            switch ($modeChoice.ToUpper()) {
                                'P' { 
                                    Write-Host "→ Mode Preview pour $svc" -ForegroundColor Cyan
                                    & $scriptSec4 -ServiceName $svc -PilotageRoot $pilotageRoot -Mode 'Preview'
                                    Read-Host 'Entrée pour continuer' | Out-Null
                                    break
                                }
                                'A' { 
                                    Write-Host "→ Mode Automatique pour $svc" -ForegroundColor Green
                                    $confirm = Read-Host "ATTENTION: Les fichiers seront supprimés et remplacés par des raccourcis. Continuer? (O/N)"
                                    if ($confirm -match '^[Oo]') {
                                        & $scriptSec4 -ServiceName $svc -PilotageRoot $pilotageRoot -Mode 'Auto' -CreateBackup
                                    }
                                    Read-Host 'Entrée pour continuer' | Out-Null
                                    break
                                }
                                'M' { 
                                    Write-Host "→ Mode Manuel pour $svc" -ForegroundColor Yellow
                                    & $scriptSec4 -ServiceName $svc -PilotageRoot $pilotageRoot -Mode 'Manual' -CreateBackup
                                    Read-Host 'Entrée pour continuer' | Out-Null
                                    break
                                }
                                'R' { break }
                                default { Write-Warning "Choix invalide" }
                            }
                        } while ($modeChoice.ToUpper() -ne 'R')
                    }
                }
                Read-Host 'Entrée pour revenir' | Out-Null
            }
            
            '5' {
                $services = Prompt-Subfolders 'Compression pdf'
                if ($services) {
                    foreach ($svc in $services) {
                        Write-Host "→ Lancement des compressions pdf pour : $svc" -ForegroundColor Green
                        & $scriptSec3 `
                            -ServiceName     $svc `
                            -BaseServicePath $baseServicePath `
                            -PilotagePath    $pilotageRoot
                    }
                }
                Read-Host 'Entrée pour revenir' | Out-Null
            }

            'Q' { break } 
            Default {
                Write-Warning "Choix invalide"
                Start-Sleep -Seconds 1
            }
        }
    } while ($true)
}

Launch-Menu

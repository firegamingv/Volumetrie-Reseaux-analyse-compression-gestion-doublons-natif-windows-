# Suite d'outils PowerShell pour la gestion de volumétrie et optimisation de l'espace disque

Cette suite d'outils PowerShell permet d'analyser, optimiser et gérer l'espace disque des services réseau. Elle comprend des fonctionnalités d'analyse volumétrique, de compression d'images, de détection et gestion de doublons.

## 📋 Table des matières

- [Vue d'ensemble](#vue-densemble)
- [Prérequis](#prérequis)
- [Installation](#installation)
- [Scripts disponibles](#scripts-disponibles)
- [Usage](#usage)
- [Configuration](#configuration)
- [Exemples](#exemples)
- [Dépannage](#dépannage)

## 🔍 Vue d'ensemble

La suite comprend 5 scripts PowerShell principaux :

- **menu.ps1** : Interface de menu principal pour lancer tous les outils
- **volumetriemenu.ps1** : Analyse volumétrique des services réseau
- **compression image.ps1** : Compression intelligente des images
- **Detection doublon.ps1** : Détection avancée des fichiers en double
- **Gestion doublons.ps1** : Gestion automatisée des doublons détectés

## ⚡ Prérequis

- PowerShell 5.1 ou supérieur
- Droits d'accès aux services réseau (``)
- Modules .NET requis :
  - `System.Drawing` (pour la compression d'images)
  - `PresentationCore` (pour les codecs d'images)
- Accès réseau aux dossiers de pilotage

## 🚀 Installation

1. Téléchargez tous les fichiers PowerShell dans un même dossier
2. Assurez-vous que la politique d'exécution PowerShell autorise les scripts :
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## 📁 Scripts disponibles

### 1. menu.ps1 - Interface principale
**Fonction :** Point d'entrée principal avec menu interactif
- Lance les différents outils de la suite
- Gestion des paramètres et validation des services
- Interface utilisateur conviviale

### 2. volumetriemenu.ps1 - Analyse volumétrique
**Fonction :** Analyse et catégorisation des fichiers par type
- **Catégories automatiques :**
  - PDF_ECM : Documents ECM (D5140, D4533)
  - PDF : Fichiers PDF standards
  - Archive_OUTLOOK : Archives PST
  - IMG_VIDEOS : Images et vidéos (30+ formats supportés)
  - ZIP_unzip : Archives ZIP
  - Autres : Tous les autres types
  - Doublons : Réservé pour les fichiers en double

**Paramètres :**
- `BaseServicePath` : Chemin de base des services
- `SubFolders` : Liste des services à analyser
- `PilotagePath` : Dossier de sortie des rapports

### 3. compression image.ps1 - Optimisation des images
**Fonction :** Compression intelligente avec préservation de la qualité
- **Formats supportés :** JPEG, PNG, GIF
- **Fonctionnalités :**
  - Redimensionnement adaptatif (défaut : 1224x1632)
  - Compression JPEG avec qualité ajustable (défaut : 75%)
  - Skip automatique des fichiers déjà traités
  - Vérification des gains d'espace
  - Traitement par batch pour les gros volumes
  - Rapport détaillé des compressions

**Paramètres :**
- `PilotagePath` : Dossier du service à traiter
- `TargetWidth` : Largeur maximale (défaut : 1224)
- `TargetHeight` : Hauteur maximale (défaut : 1632)
- `JpegQuality` : Qualité JPEG 1-100 (défaut : 75)

### 4. Detection doublon.ps1 - Détection avancée
**Fonction :** Détection ultra-performante des fichiers en double
- **Algorithmes avancés :**
  - Hash partiel intelligent pour gros fichiers
  - Distance de Levenshtein pour noms similaires
  - Métadonnées étendues avec critère "même jour"
  - Validation par hash MD5/SHA256

**Paramètres :**
- `ServiceName` : Nom du service à analyser
- `PilotageRoot` : Dossier racine de pilotage
- `NameThreshold` : Seuil de similarité des noms (défaut : 1)
- `MaxSizeDiff` : Différence de taille maximale en Mo (défaut : 1)
- `PartialHashSize` : Taille du hash partiel en bytes (défaut : 65536)
- `SmallFileThreshold` : Seuil pour hash complet en bytes (défaut : 10MB)
- `UseExtendedMetadata` : Utilise les métadonnées étendues (recommandé)
- `SkipHashValidation` : Mode métadonnées uniquement

### 5. Gestion doublons.ps1 - Traitement automatisé
**Fonction :** Suppression intelligente et création de raccourcis
- **Modes de fonctionnement :**
  - **Preview** : Simulation sans modification
  - **Manual** : Validation manuelle de chaque groupe
  - **Auto** : Traitement automatisé avec vérifications hash

**Fonctionnalités :**
- Sélection intelligente du "meilleur" fichier à conserver
- Création de raccourcis Windows (.lnk) transparents
- Validation hash obligatoire en mode automatique
- Critère "même jour" strict pour la sécurité
- Rapport détaillé des actions effectuées

**Paramètres :**
- `ServiceName` : Nom du service
- `PilotageRoot` : Dossier racine de pilotage
- `DoublonsFile` : Fichier CSV des doublons (optionnel)
- `Mode` : 'Preview', 'Manual' ou 'Auto'
- `CreateBackup` : Création de sauvegarde
- `UseConsolidated` : Utilise le rapport consolidé

## 🎯 Usage

### Lancement rapide
```powershell
# Depuis le dossier contenant les scripts
.\menu.ps1
```

### Usage direct des scripts

#### Analyse volumétrique
```powershell
.\volumetriemenu.ps1 -BaseServicePath "" -SubFolders @("Service1", "Service2") -PilotagePath "C:\Pilotage"
```

#### Compression d'images
```powershell
.\compression image.ps1 -PilotagePath "C:\Pilotage\MonService" -TargetWidth 1024 -TargetHeight 1024 -JpegQuality 80
```

#### Détection de doublons
```powershell
.\Detection doublon.ps1 -ServiceName "MonService" -PilotageRoot "C:\Pilotage" -UseExtendedMetadata
```

#### Gestion des doublons (Preview)
```powershell
.\Gestion doublons.ps1 -ServiceName "MonService" -PilotageRoot "C:\Pilotage" -Mode Preview
```

#### Gestion automatisée des doublons
```powershell
.\Gestion doublons.ps1 -ServiceName "MonService" -PilotageRoot "C:\Pilotage" -Mode Auto -CreateBackup
```

## ⚙️ Configuration

### Chemins par défaut
```powershell
# Chemin des services
$baseServicePath = ''

# Dossier de pilotage
$pilotageRoot = ''
```

### Paramètres de compression par défaut
```powershell
$defaultWidth   = 1224    # Largeur maximale
$defaultHeight  = 1632    # Hauteur maximale  
$defaultQuality = 75      # Qualité JPEG
```

## 📊 Exemples de workflow complet

### 1. Workflow d'analyse et optimisation complète
```powershell
# 1. Analyse volumétrique
.\menu.ps1
# Choisir option 1 : Traitement volumétrie

# 2. Compression des images  
# Choisir option 2 : Compression des images

# 3. Détection des doublons
# Choisir option 3 : Détection des doublons

# 4. Preview des actions de déduplication
# Choisir option 4 : Gestion des doublons -> Preview

# 5. Application automatique (optionnel)
# Choisir option 4 : Gestion des doublons -> Auto
```

### 2. Traitement d'urgence avec scripts individuels
```powershell
# Analyse rapide
.\volumetriemenu.ps1 -BaseServicePath "\\server\services" -SubFolders @("UrgentService") -PilotagePath "C:\Temp\Pilotage"

# Compression agressive
.\compression image.ps1 -PilotagePath "C:\Temp\Pilotage\UrgentService" -JpegQuality 60

# Détection et traitement des doublons
.\Detection doublon.ps1 -ServiceName "UrgentService" -UseExtendedMetadata
.\Gestion doublons.ps1 -ServiceName "UrgentService" -Mode Auto
```

## 📈 Résultats et rapports

### Fichiers générés par service

#### Analyse volumétrique :
- `DonneeVolumetrique_PDF_ECM.csv`
- `DonneeVolumetrique_PDF.csv`
- `DonneeVolumetrique_Archive_OUTLOOK.csv`
- `DonneeVolumetrique_IMG_VIDEOS.csv`
- `DonneeVolumetrique_ZIP_unzip.csv`
- `DonneeVolumetrique_Autres.csv`
- `DonneeVolumetrique_Doublons.csv`

#### Compression d'images :
- `CompressionReport_Final.csv` : Rapport détaillé des compressions

#### Détection de doublons :
- `DonneeVolumetrique_ConfirmedDoublons_Advanced.csv` : Doublons confirmés
- `Doublons_Consolides_Advanced_[timestamp].csv` : Rapport consolidé

#### Gestion des doublons :
- `Gestion_Doublons_Intelligente_[timestamp].log` : Journal des actions
- `Backup_Metadata_[timestamp].json` : Sauvegarde métadonnées (si activé)

## 🔧 Dépannage

### Problèmes courants

#### Erreur d'accès réseau
```
Solution : Vérifier la connectivité réseau et les droits d'accès
Test : Test-Path "\\atlas.edf.fr\CO\45dam-dpn\services.006"
```

#### Erreur de politique d'exécution
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### Fichiers verrouillés lors de la compression
```
Solution : Fermer les applications utilisant les fichiers
Le script skip automatiquement les fichiers verrouillés
```

#### Hash différents lors de la gestion des doublons
```
Cause : Fichiers modifiés entre détection et gestion
Solution : Relancer la détection puis la gestion
```

### Mode de débogage

#### Activation des logs détaillés
```powershell
# Pour la détection de doublons
.\Detection doublon.ps1 -ServiceName "Test" -Verbose

# Pour la gestion avec validation étendue  
.\Gestion doublons.ps1 -ServiceName "Test" -Mode Manual -Verbose
```

## 🔒 Sécurité

### Mesures de protection intégrées

1. **Validation hash obligatoire** en mode automatique
2. **Critère "même jour"** pour éviter les faux positifs
3. **Création de raccourcis** au lieu de suppressions définitives
4. **Mode Preview** pour tester sans risque
5. **Sauvegardes optionnelles** des métadonnées

### Recommandations

- Toujours tester en mode **Preview** d'abord
- Utiliser le mode **Manual** pour les données critiques
- Créer des sauvegardes avec `-CreateBackup` 
- Vérifier les rapports avant validation
- Tester sur un sous-ensemble de données d'abord

## 📞 Support

Pour toute question ou problème :
1. Consulter les fichiers de log générés
2. Vérifier la connectivité réseau
3. Tester en mode Preview d'abord
4. Utiliser le mode Verbose pour plus de détails

---

**Version :** 2.0  
**Dernière mise à jour :** 2025  
**Compatibilité :** PowerShell 5.1+, Windows Server 2016+

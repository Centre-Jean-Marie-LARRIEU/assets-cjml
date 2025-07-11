﻿# set_user_signatures.ps1 (v48.02 - Version Finale Signature, QR Code, Carte Numérique avec Mentions Légagles et Carte Imprimable + GA4 Tracking)
#
param(
    [string]$SingleUserEmail = "",
    [switch]$IncludeSuspended,
    [switch]$AddDigitalCard,
    [switch]$GeneratePrintQr,
    [switch]$GeneratePrintableCard,
    [switch]$GeneratePdfCard,
    [switch]$ShowHelp,
    [switch]$DebugMode
)

# NOUVEAU : Définir et afficher la version du script APRES le bloc param
$script:ScriptVersion = "v48.02 - Version Finale Signature, QR Code, Carte Numérique avec Mentions Légagles et Carte Imprimable + GA4 Tracking"
Write-Host "Démarrage du script : set_user_signatures.ps1 ($script:ScriptVersion)" -ForegroundColor Green

if ($ShowHelp) {
    $helpText = @"
NOM:
    set_user_signatures.ps1

SYNOPSIS:
    Met à jour les signatures email et peut générer une carte de visite numérique complète (vCard + QR Code interactif).
    Ce script a été optimisé pour la concision et la clarté.

SYNTAXE:
    .\set_user_signatures.ps1 [-SingleUserEmail <string>] [-IncludeSuspended] [-AddDigitalCard] [-GeneratePrintQr] [-GeneratePrintableCard] [-GeneratePdfCard] [-ShowHelp] [-DebugMode]

DESCRIPTION:
    Ce script automatise la mise à jour des signatures Gmail via GAM. Il est optimisé pour ne pas effectuer de mises à jour inutiles.

    - Mode standard : Met à jour la signature principale de l'utilisateur.

    - Mode Carte de Visite (-AddDigitalCard) : Génère une page web professionnelle (hébergée sur GitHub Pages) qui contient un lien de
      déchargement direct pour la vCard de l'utilisateur. Pour assurer la compatibilité, la vCard est encodée
      directement dans le lien de téléchargement (méthode Data-URL).
      Nouveauté : La carte inclus désormais un QR Code interactif qui peut être agrandi pour un partage facile,
      et le label de l'adresse (ex: 'Siège Social') est dynamiquement affiché.
      Cette version inclut le suivi Google Analytics 4 (GA4) pour les consultations de pages et les clics sur les boutons.

PARAMÈTRES:
    -SingleUserEmail <string>
        Spécifie l'adresse email d'un seul utilisateur à mettre à jour.

    -IncludeSuspended
        Commutateur. Si présent, le script mettra à jour TOUS les utilisateurs, y compris les comptes suspendus.

    -AddDigitalCard
        Commutateur. Si présent, active la génération de la carte de visite numérique avec QR Code et le suivi GA4.

    -GeneratePrintQr
        Commutateur. Si présent, génère des QR Codes haute résolution pour l'impression dans un dossier local.
        Ces fichiers NE sont PAS poussés vers GitHub.

    -GeneratePrintableCard
        Commutateur. Si présent, génère un fichier HTML de carte de visite optimisé pour l'impression locale.
        Ce fichier est destiné à être ouvert dans un navigateur puis "Imprimé au format PDF".

    -GeneratePdfCard
        Commutateur. Si présent, génère des fichiers PDF de cartes de visite imprimables (recto-verso) en utilisant wkhtmltopdf.
        Nécessite wkhtmltopdf installé et accessible via PATH.

    -ShowHelp
        Affiche ce message d'aide et quitte le script.

    -DebugMode
        Commutateur. Si présent, affiche des informations de débogage détaillées sur les variables HTML générées.
        À utiliser pour diagnostiquer les problèmes d'affichage.

EXEMPLES:
    # Affiche cette aide complète
    .\set_user_signatures.ps1 -ShowHelp

    # Met à jour la signature ET la carte de visite numérique pour un utilisateur spécifique
    # Le site web est configuré directement dans le script.
    .\set_user_signatures.ps1 -SingleUserEmail "s.gille@cjml.fr" -AddDigitalCard

    # Génère des QR Codes haute résolution pour tous les utilisateurs actifs (pour impression)
    # Note: Génère des fichiers .png dans C:\GAMWork\PrintQrCodes
    .\set_user_signatures.ps1 -GeneratePrintQr

    # Génère une carte de visite HTML imprimable pour un utilisateur spécifique
    # Note: Génère un fichier .html dans C:\GAMWork\PrintableCards à ouvrir dans un navigateur pour imprimer.
    .\set_user_signatures.ps1 -SingleUserEmail "s.gille@cjml.fr" -GeneratePrintableCard

    # Génère une carte de visite PDF pour un utilisateur (recto-verso sur 2 pages du PDF)
    # Note: Nécessite wkhtmltopdf.exe dans le PATH ou chemin complet configuré. Génère un .pdf dans C:\GAMWork\PdfCards
    .\set_user_signatures.ps1 -SingleUserEmail "s.gille@cjml.fr" -GeneratePdfCard

    # Met à jour les signatures, les cartes de visite ET génère des QR Codes/Cartes imprimables/PDF pour un utilisateur
    .\set_user_signatures.ps1 -SingleUserEmail "s.gille@cjml.fr" -AddDigitalCard -GeneratePrintQr -GeneratePrintableCard -GeneratePdfCard -DebugMode
"@
    Write-Host $helpText
    return
}

# --- Configuration Globale ---
$config = @{
    ProjectRoot           = "C:\GAMWork\Signatures"
    GamPath               = "C:\GAM7\gam.exe"
    SignatureTemplateName = "signature_template.html"
    DigitalCardTemplateName = "digital_card_template.html"
    PrintableCardTemplateName = "printable_business_card_template.html"
    PrintQrOutputFolder   = "C:\GAMWork\PrintQrCodes"
    PrintableCardOutputFolder = "C:\GAMWork\PrintableCards"
    PdfCardOutputFolder   = "C:\GAMWork\PdfCards"

    SignatureLogoUrl      = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/Logo-CJML.png"
    DigitalCardLogoUrl    = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/logo-horizontal.jpg"
    PrintLogoUrl          = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/Logo-CJML.png"

    # --- NOUVELLES LIGNES À AJOUTER POUR LES LOGOS PARTENAIRES ET SOCIAUX ---
    GcsmsLogoUrl          = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/logo-gcsms-pyrenees.png" # <--- NOUVELLE URL
    FacebookLogoUrl       = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/icon-facebook.png"    # <--- NOUVELLE URL
    LinkedinLogoUrl       = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/icon-linkedin.png"    # <--- NOUVELLE URL

    FacebookPageUrl       = "https://www.facebook.com/CentreJeanMarieLARRIEU" # <-- REMPLACEZ PAR L'URL RÉELLE DE LA PAGE FB DU CJML
    LinkedinCompanyUrl    = "https://www.linkedin.com/company/centre-jean-marie-larrieu" # <-- REMPLACEZ PAR L'URL RÉELLE DE LA PAGE LINKEDIN DU CJML
    # --- FIN NOUVELLES LIGNES ---
	
    OrgName               = "Centre Jean-Marie LARRIEU"
    DefaultPhoneNumberRaw = "+33562913250"
    DefaultPhoneNumberDisplay = "05 62 91 32 50"
    DefaultAddress        = @"
414 Rue du Layris
65710 CAMPAN
"@.Trim()

    WebsiteUrl            = "http://www.cjml.fr"
    MainDomain            = "cjml.fr"
    ExcludeDomain         = "eleves.cjml.fr"

    QrCodeDllPath         = Join-Path -Path $PSScriptRoot -ChildPath "QRCoder.dll"
    QrCodeColors = @{
        Blue  = [byte[]](6, 143, 208)
        White = [byte[]](255, 255, 255)
    }

    WkhtmltopdfPath       = "wkhtmltopdf.exe"
}

# Calculs de chemins et URLs basés sur la configuration
$config.SignatureTemplatePath = Join-Path -Path $config.ProjectRoot -ChildPath $config.SignatureTemplateName
$config.DigitalCardTemplatePath = Join-Path -Path $config.ProjectRoot -ChildPath $config.DigitalCardTemplateName
$config.PrintableCardTemplatePath = Join-Path -Path $config.ProjectRoot -ChildPath $config.PrintableCardTemplateName

# Calcul de websiteDisplayUrl
if (-not [string]::IsNullOrEmpty($config.WebsiteUrl)) {
    $config.WebsiteDisplayUrl = $config.WebsiteUrl -replace "^https?:\/\/(www\.)?", ""
    if ($config.WebsiteDisplayUrl -like "cjml.fr*") {
        $config.WebsiteDisplayUrl = "www." + $config.WebsiteDisplayUrl
    }
} else {
    $config.WebsiteDisplayUrl = ""
}

# GitHub Configuration (chargé séparément car dépend du fichier token)
$githubConfig = @{
    UserOrOrg = "Centre-Jean-Marie-LARRIEU"
    Repo = "assets-cjml"
    VcardFolderPath = "vcards"
    QrcodeFolderPath = "qrcodes"
    PagesBaseUrl = "https://ressources.cjml.fr"
}
try {
    $tokenPath = Join-Path -Path $config.ProjectRoot -ChildPath "github_token.txt"
    $githubConfig.Token = Get-Content -Path $tokenPath -Raw -ErrorAction Stop
} catch {
    Write-Warning "Fichier 'github_token.txt' introuvable dans $($config.ProjectRoot). La fonction -AddDigitalCard sera désactivée si utilisée."
    $githubConfig.Token = $null
}


# --- Initialisation de l'environnement ---
$originalEncoding = [Console]::OutputEncoding
chcp.com 65001 | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
    Add-Type -Path $config.QrCodeDllPath -ErrorAction Stop
} catch {
    Write-Error "Impossible de charger la bibliothèque QRCoder.dll. Assurez-vous que le fichier se trouve bien à l'emplacement : $($config.QrCodeDllPath)"
    exit 1
}

# --- Fonctions optimisées ---

function Invoke-GitPublish {
    param(
        [string]$FileName,
        [byte[]]$FileContentBytes,
        [string]$FolderPathInRepo,
        [hashtable]$GitHubConfig,
        [switch]$DebugMode
    )
    if ([string]::IsNullOrEmpty($GitHubConfig.UserOrOrg) -or [string]::IsNullOrEmpty($GitHubConfig.Repo) -or [string]::IsNullOrEmpty($FolderPathInRepo)) {
        Write-Error "Configuration GitHub incomplète (UserOrOrg, Repo ou FolderPathInRepo sont vides). Impossible d'appeler l'API GitHub."
        return $null
    }
    
    $apiUrl = "https://api.github.com/repos/$($GitHubConfig.UserOrOrg)/$($GitHubConfig.Repo)/contents/$FolderPathInRepo/$FileName"
    $headers = @{ "Authorization" = "Bearer $($GitHubConfig.Token)"; "Accept" = "application/vnd.github.com.v3+json" }

    if ($DebugMode) {
        Write-Host "    - DEBUG: Appel API GitHub pour $FileName - URL: $apiUrl" -ForegroundColor DarkCyan
        Write-Host "    - DEBUG: Dossier cible sur GitHub: $FolderPathInRepo" -ForegroundColor DarkCyan
    }

    $sha = $null
    try {
        $existingFile = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers -ErrorAction Stop
        if ($existingFile) {
            $sha = $existingFile.sha
            if (-not $FileContentBytes) {
                Write-Warning ("      Le contenu du fichier est null, ne peut pas calculer le SHA pour " + $FileName)
                $localSha = "INVALID_CONTENT_" + (Get-Random).ToString()
            } else {
                $header = "blob $($FileContentBytes.Length)`0"; $headerBytes = [System.Text.Encoding]::UTF8.GetBytes($header)
                $combinedBytes = $headerBytes + $FileContentBytes; $sha1 = New-Object System.Security.Cryptography.SHA1Managed
                $localSha = [System.BitConverter]::ToString($sha1.ComputeHash($combinedBytes)).Replace("-", "").ToLower()
            }

            if ($DebugMode) {
                Write-Host "    - DEBUG : Forçage de la mise à jour pour $FileName (DebugMode activé)." -ForegroundColor Cyan
                $localSha = "FORCE_UPDATE_" + (Get-Random).ToString()
            }

            if ($localSha -eq $sha) {
                Write-Host "    - Contenu identique pour '$FileName' sur GitHub. Aucune mise à jour nécessaire." -ForegroundColor Green
                return $existingFile
            }
            Write-Host "    - Fichier existant détecté sur GitHub. Préparation de la mise à jour." -ForegroundColor DarkGray
        }
    }
    catch [System.Net.WebException] {
        if ($_.Exception.Response.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) { Write-Host "    - Fichier absent sur GitHub. Préparation pour la création." -ForegroundColor DarkGray }
        else {
            $errorMessage = ($_.Exception).Message
            Write-Warning ("      Erreur web inattendue lors de la vérification de " + $FileName + ": " + $errorMessage)
        }
    }
    catch {
        $errorMessage = ($_.Exception).Message
        Write-Warning ("      Erreur inattendue lors de la vérification de " + $FileName + ": " + $errorMessage)
    }

    if (-not $FileContentBytes -or $FileContentBytes.Length -eq 0) {
        Write-Error "Le contenu du fichier pour l'upload GitHub est vide ou null. L'upload est annulé pour $FileName."
        return $null
    }

    $contentBase64 = [System.Convert]::ToBase64String($FileContentBytes)
    $body = @{ message = "Automated update of $FileName"; content = $contentBase64 }
    if ($sha) { $body.Add("sha", $sha) }

    try {
        $uploadResult = Invoke-RestMethod -Uri $apiUrl -Method Put -Headers $headers -Body ($body | ConvertTo-Json) -ContentType "application/json"
        Write-Host "    - '$FileName' $(if ($sha) {'mis à jour'} else {'créé'}) sur GitHub." -ForegroundColor Green
        return $uploadResult.content
    } catch {
        $detailedError = $_.Exception.Message
        if ($_.Exception.Response) {
            try {
                $responseBody = (New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() | ConvertFrom-Json
                $detailedError += " - GitHub API Response: $($responseBody.message)"
            } catch {
                # Could not parse response body
            }
        }
        Write-Error "Échec de l'upload sur GitHub pour $FileName. Erreur: $detailedError"
        return $null
    }
}

function Generate-QrCodeFile {
    param(
        [string]$QrDataUrl,
        [string]$OutputFileName,
        [string]$OutputFolder,
        [hashtable]$QrCodeColors,
        [int]$PixelsPerModule = 100
    )

    if ([string]::IsNullOrEmpty($OutputFolder)) {
        Write-Error "Le chemin de sortie du QR Code est vide. Impossible de générer le fichier."
        return $false
    }
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }

    try {
        $qrGenerator = New-Object QRCoder.QRCodeGenerator
        if ([string]::IsNullOrEmpty($QrDataUrl)) {
            Write-Error "L'URL des données du QR Code est vide. Impossible de générer le code."
            return $false
        }
        $qrCodeData = $qrGenerator.CreateQrCode($QrDataUrl, [QRCoder.QRCodeGenerator+ECCLevel]::Q)
        
        if ($qrCodeData -eq $null) {
            Write-Error "La génération des données QR Code a échoué pour l'URL: $QrDataUrl. Le QR CodeData est null."
            return $false
        }

        $qrCode = New-Object QRCoder.PngByteQRCode($qrCodeData)

        $qrCodeBytes = $qrCode.GetGraphic($PixelsPerModule, $QrCodeColors.Blue, $QrCodeColors.White)
        $outputPath = Join-Path -Path $OutputFolder -ChildPath $OutputFileName

        [System.IO.File]::WriteAllBytes($outputPath, $qrCodeBytes)
        Write-Host "    QR Code généré : $outputPath" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Échec de la génération du QR Code pour '$OutputFileName'. Erreur: $($_.Exception.Message)"
        return $false
    }
}

# La fonction Format-PhoneNumber doit être définie avant Get-UserPhoneData
function Format-PhoneNumber ($rawNumber) {
    if (-not [string]::IsNullOrEmpty($rawNumber)) {
        $formattedDisplayValue = $rawNumber
        if ($rawNumber -match '^\+33[1-9]\d{8}$') {
            $localNumber = $rawNumber -replace '^\+33', '0'
            $formattedDisplayValue = $localNumber -replace '^(0\d)(\d{2})(\d{2})(\d{2})(\d{2})$', '$1 $2 $3 $4 $5'
        } elseif ($rawNumber -match '^[0-9]{9,10}$' -and -not ($rawNumber -match '^\+')) {
            $formattedDisplayValue = $rawNumber -replace '^(0\d)(\d{2})(\d{2})(\d{2})(\d{2})$', '$1 $2 $3 $4 $5'
        }
        $rawNumberForLink = $rawNumber
        if ($rawNumberForLink -match '^0[1-9]\d{8}$') { $rawNumberForLink = "+33" + $rawNumberForLink.Substring(1) }
        return @{ Display = $formattedDisplayValue; Raw = $rawNumber; RawForLink = $rawNumberForLink }
    }
    return $null
}

function Get-UserPhoneData {
    param(
        [pscustomobject]$UserObject,
        [string]$DefaultPhoneNumberRaw,
        [string]$DefaultPhoneNumberDisplay
    )

    $phoneData = @{
        WorkPhoneDisplayForTemplates = "";
        MobilePhoneDisplayForTemplates = "";
        RawWorkPhone = "";
        RawMobilePhone = "";
        RawPrimaryDisplayPhone = ""; # Le numéro qui sera affiché en premier (work > mobile > standard)
        RawPrimaryDialPhone = ""; # Le numéro à composer pour la ligne principale
        UsedDefaultPhoneAsPrimary = $false; # Indique si le numéro principal est le standard par défaut
        HasMobilePhoneFromGam = $false; # Indique si un vrai numéro mobile de GAM est présent
        HasWorkPhoneFromGam = $false; # Indique si un vrai numéro de travail de GAM est présent
    }
    
    $gamWorkPhoneParsed = $null
    $gamMobilePhoneParsed = $null

    for ($i = 0; $i -lt 5; $i++) {
        $typeProperty = "phones.$i.type"
        $valueProperty = "phones.$i.value"

        if ($UserObject.PSObject.Properties.Name -contains $typeProperty -and $UserObject.PSObject.Properties.Name -contains $valueProperty) {
            $phoneType = $UserObject.$typeProperty
            $phoneValue = $UserObject.$valueProperty

            $formatted = Format-PhoneNumber($phoneValue)
            if ($formatted) {
                if ($phoneType -eq 'work' -and ($gamWorkPhoneParsed -eq $null)) {
                    $gamWorkPhoneParsed = $formatted
                    $phoneData.HasWorkPhoneFromGam = $true
                } elseif ($phoneType -eq 'mobile' -and ($gamMobilePhoneParsed -eq $null)) {
                    $gamMobilePhoneParsed = $formatted
                    $phoneData.HasMobilePhoneFromGam = $true
                }
            }
        }
    }
    
    if ($DebugMode) {
        Write-Host "    - DEBUG: Parsed GAM Phones (direct access):" -ForegroundColor DarkYellow
        Write-Host "      HasWorkPhoneFromGam: $($phoneData.HasWorkPhoneFromGam)" -ForegroundColor DarkYellow
        Write-Host "      gamWorkPhoneParsed: $($gamWorkPhoneParsed | ConvertTo-Json)" -ForegroundColor DarkYellow
        Write-Host "      HasMobilePhoneFromGam: $($phoneData.HasMobilePhoneFromGam)" -ForegroundColor DarkYellow
        Write-Host "      gamMobilePhoneParsed: $($gamMobilePhoneParsed | ConvertTo-Json)" -ForegroundColor DarkYellow
    }

    # Assignation des numéros Raw (ceux directement issus de GAM)
    if ($gamWorkPhoneParsed) {
        $phoneData.RawWorkPhone = $gamWorkPhoneParsed.RawForLink
        $phoneData.WorkPhoneDisplayForTemplates = $gamWorkPhoneParsed.Display
    }
    if ($gamMobilePhoneParsed) {
        $phoneData.RawMobilePhone = $gamMobilePhoneParsed.RawForLink
        $phoneData.MobilePhoneDisplayForTemplates = $gamMobilePhoneParsed.Display
    }

    # Détermination du numéro principal affiché (RawPrimaryDisplayPhone et RawPrimaryDialPhone)
    # Règle: mobile (si existe) > work (si existe) > standard (par défaut)
    if ($phoneData.HasMobilePhoneFromGam) { # Priorité 1: Mobile de GAM
        $phoneData.RawPrimaryDisplayPhone = $phoneData.MobilePhoneDisplayForTemplates
        $phoneData.RawPrimaryDialPhone = $phoneData.RawMobilePhone
    } elseif ($phoneData.HasWorkPhoneFromGam) { # Priorité 2: Work phone de GAM (si pas de mobile)
        $phoneData.RawPrimaryDisplayPhone = $phoneData.WorkPhoneDisplayForTemplates
        $phoneData.RawPrimaryDialPhone = $phoneData.RawWorkPhone
    } else { # Priorité 3: Aucun téléphone GAM trouvé, utiliser le standard par défaut
        $phoneData.RawPrimaryDisplayPhone = $DefaultPhoneNumberDisplay
        $phoneData.RawPrimaryDialPhone = $DefaultPhoneNumberRaw
        $phoneData.UsedDefaultPhoneAsPrimary = $true
    }

    # Les champs WorkPhoneHtmlForSignature et MobilePhoneHtmlForSignature ne sont plus générés ici,
    # mais directement dans le bloc de préparation de la signature (phoneBlockHtmlForSignatureFinal)
    # pour une flexibilité maximale dans l'ordre et les labels.

    return $phoneData
}


function Get-TemplateContent($TemplatePath) {
    $content = Get-Content -Path $TemplatePath -Encoding UTF8 -Raw
    return $content.TrimStart([char]65279, [char]22)
}

# --- Récupération et filtrage des utilisateurs ---
$usersToProcess = @()
if (-not $ShowHelp) {
    $fieldsToGet = 'primaryEmail,name,organizations,phones,addresses,suspended'
    if (-not [string]::IsNullOrEmpty($SingleUserEmail)) {
        Write-Host "--- MODE UTILISATEUR UNIQUE: Cible l'utilisateur '$SingleUserEmail' ---" -ForegroundColor Yellow
        $gamArgs = @(
            'print', 'users',
            'query', "email='$SingleUserEmail'",
            'fields', $fieldsToGet
        )
        Write-Host "DEBUG: GAM Command for single user: $($config.GamPath) $($gamArgs -join ' ')" -ForegroundColor DarkGray
        $gamOutput = & $config.GamPath $gamArgs | ConvertFrom-Csv
        if ($gamOutput) { $usersToProcess = $gamOutput }
        else { Write-Error "Impossible de récupérer les informations pour l'utilisateur '$SingleUserEmail'." }
    } else {
        $gamArgs = @('print', 'users');
        if (-not $IncludeSuspended) { $gamArgs += 'query', 'isSuspended=False' }
        $gamArgs += 'fields', $fieldsToGet;
        
        Write-Host "DEBUG: GAM Command for multiple users: $($config.GamPath) $($gamArgs -join ' ')" -ForegroundColor DarkGray
        $gamRawOutput = & $config.GamPath $gamArgs;
        $allGSuiteUsers = $gamRawOutput | ConvertFrom-Csv;
        $usersToProcess = $allGSuiteUsers | Where-Object { $_.primaryEmail -like "*@$($config.MainDomain)" -and $_.primaryEmail -notlike "*@$($config.ExcludeDomain)" }
    }
}

if ($usersToProcess.Count -eq 0 -and -not $ShowHelp) { Write-Host "Aucun utilisateur trouvé à traiter. Quitte le script."; exit 0 }
if (-not $ShowHelp) { Write-Host "Found $($usersToProcess.Count) user(s) to process." -ForegroundColor Cyan }

# --- Boucle principale de traitement des utilisateurs ---
foreach ($user in $usersToProcess) {
    if ($user -eq $null) { Write-Error "Skipping null user object."; continue }

    $primaryEmail_val = $user.primaryEmail
    $givenName_val = $user."name.givenName"
    $familyName_val = $user."name.familyName"
    $title_val = if ($user."organizations.0.title") { $user."organizations.0.title" } else { "" }
    $isSuspended = [string]::Equals($user.suspended, "True", [System.StringComparison]::OrdinalIgnoreCase)

    Write-Host "--- Processing user: $primaryEmail_val (Suspended: $isSuspended) ---" -ForegroundColor Cyan

    # Traitement de l'adresse
    $address_val = $config.DefaultAddress
    $addressLabelForCard = "Siège Social"

    $addressFound = $false
    for ($i = 0; $i -lt 5; $i++) {
        $typeProperty = "addresses.$i.type"
        $formattedProperty = "addresses.$i.formatted"

        if ($user.PSObject.Properties.Name -contains $typeProperty -and $UserObject.PSObject.Properties.Name -contains $formattedProperty) {
            if ($user.$typeProperty -eq 'work' -and -not [string]::IsNullOrEmpty($user.$formattedProperty)) {
                $address_val = $user.$formattedProperty.Trim()
                $addressLabelForCard = "Adresse du Bureau"
                $addressFound = $true
                break
            }
        }
    }

    $addressForSignature = ($address_val -replace "`r`n|`n", " - ").Trim()
    $addressForDigitalCard = ($address_val -replace "`r`n", "<br>").Trim()
    $address_url_maps = "https://www.google.com/maps/search/?api=1&query=" + [System.Net.WebUtility]::UrlEncode($addressForSignature)


    # Préparation des données téléphoniques
    $phoneData = Get-UserPhoneData -UserObject $user -DefaultPhoneNumberRaw $config.DefaultPhoneNumberRaw -DefaultPhoneNumberDisplay $config.DefaultPhoneNumberDisplay

    # --- DÉBOGAGE : Affichage des valeurs clés ---
    if ($DebugMode) {
        Write-Host "--- DÉBOGAGE: Données utilisateur et téléphone ---" -ForegroundColor Yellow
        Write-Host "  Email: $primaryEmail_val" -ForegroundColor DarkCyan
        Write-Host "  Nom Complet: $givenName_val $familyName_val" -ForegroundColor DarkCyan
        Write-Host "  Titre: $title_val" -ForegroundColor DarkCyan
        Write-Host "  Adresse Finale: $address_val (Label: $addressLabelForCard)" -ForegroundColor DarkCyan
        Write-Host "  URL Maps: $address_url_maps" -ForegroundColor DarkCyan
        $phoneData | Format-List -Force
        Write-Host "--- FIN DÉBOGAGE ---" -ForegroundColor Yellow
    }

    # --- Préparation des variables spécifiques à la carte numérique et à la signature ---

    $cardContactTextHtmlForDigitalCard = ""
    $phoneBlockHtmlForSignatureFinal = "" # Nouveau: Initialiser ici pour la signature

    $linkStyleGeneral = "color: #555555; text-decoration: underline;" # Style commun pour les liens de téléphone/email/adresse


    # Logique pour construire les lignes de téléphone pour la CARTE NUMÉRIQUE et la SIGNATURE
    # Ligne 1 : Mobile si présent, sinon Ligne directe si présente, sinon Téléphone du Centre
    if ($phoneData.HasMobilePhoneFromGam) { # Priorité 1: Mobile de GAM
        $cardContactTextHtmlForDigitalCard += @"
<div class="contact-item"><span class="label">Mobile</span><a href="tel:$($phoneData.RawMobilePhone)">$($phoneData.MobilePhoneDisplayForTemplates)</a></div>
"@
        $phoneBlockHtmlForSignatureFinal += "Mobile : <a href=`"tel:$($phoneData.RawMobilePhone)`" style=`"$linkStyleGeneral`">$($phoneData.MobilePhoneDisplayForTemplates)</a><br>"
    }

    if ($phoneData.HasWorkPhoneFromGam) { # Priorité 2: Ligne directe (si présente)
        $cardContactTextHtmlForDigitalCard += @"
<div class="contact-item"><span class="label">Ligne directe</span><a href="tel:$($phoneData.RawWorkPhone)">$($phoneData.WorkPhoneDisplayForTemplates)</a></div>
"@
        $phoneBlockHtmlForSignatureFinal += "Ligne directe : <a href=`"tel:$($phoneData.RawWorkPhone)`" style=`"$linkStyleGeneral`">$($phoneData.WorkPhoneDisplayForTemplates)</a><br>"
    }

    # Toujours ajouter le Téléphone du Centre si aucun work phone n'est présent (que le mobile soit là ou non)
    if (-not $phoneData.HasWorkPhoneFromGam) {
        # Vérifier si le mobile de GAM est le seul numéro et le standard n'est pas déjà affiché comme principal (pour la signature)
        # OU si le standard est le numéro principal (aucun work/mobile de GAM)
        # Évitons la duplication: si Mobile EST la ligne principale dans phoneData, ET c'est le seul numéro GAM.
        # On va l'ajouter si le mobile n'est PAS le numéro du centre.
        if ($phoneData.RawMobilePhone -ne $config.DefaultPhoneNumberRaw) { # Pour éviter de dupliquer si le mobile est aussi le standard
            $cardContactTextHtmlForDigitalCard += @"
<div class="contact-item"><span class="label">Téléphone (Centre)</span><a href="tel:$($config.DefaultPhoneNumberRaw)">$($config.DefaultPhoneNumberDisplay)</a></div>
"@
            # Pour la signature, n'ajouter que si le standard n'est pas déjà le mobile principal
            if ($phoneData.RawPrimaryDialPhone -ne $config.DefaultPhoneNumberRaw -or -not $phoneData.HasMobilePhoneFromGam) { # Si standard est principal OU si mobile est principal mais standard est un autre numero
                $phoneBlockHtmlForSignatureFinal += "Téléphone (Centre) : <a href=`"tel:$($config.DefaultPhoneNumberRaw)`" style=`"$linkStyleGeneral`">$($config.DefaultPhoneNumberDisplay)</a><br>"
            }
        }
    }


    # Email
    $cardContactTextHtmlForDigitalCard += @"
<div class="contact-item"><span class="label">Email</span><a href="mailto:$primaryEmail_val">$primaryEmail_val</a></div>
"@
    $phoneBlockHtmlForSignatureFinal += "<a href=`"mailto:$primaryEmail_val`" style=`"$linkStyleGeneral`">$primaryEmail_val</a><br>"


    # Site Web
    if (-not [string]::IsNullOrEmpty($config.WebsiteUrl)) {
        $cardContactTextHtmlForDigitalCard += @"
<div class="contact-item">
    <span class="label">Site Web</span>
    <a href="$($config.WebsiteUrl)" target="_blank" rel="noopener noreferrer" style="color: var(--primary-blue); text-decoration: underline;">$($config.WebsiteDisplayUrl)</a>
</div>
"@
        $phoneBlockHtmlForSignatureFinal += "<a href=`"$($config.WebsiteUrl)`" target=`"_blank`" rel=`"noopener noreferrer`" style=`"color: #FCB041; text-decoration: underline; font-weight: bold;`">$($config.OrgName)</a><br>"
    }
    # Adresse (toujours en dernière position)
    $cardContactTextHtmlForDigitalCard += @"
<div class="contact-item">
    <span class="label">$addressLabelForCard</span>
    <a href="https://www.google.com/maps/search/?api=1&query=$([System.Net.WebUtility]::UrlEncode($addressForSignature))" target="_blank" rel="noopener noreferrer">$addressForSignature</a>
</div>
"@
    $phoneBlockHtmlForSignatureFinal += "<a href=`"$address_url_maps`" target=`"_blank`" rel=`"noopener noreferrer`" style=`"$linkStyleGeneral`">$addressForSignature</a><br>"


    # --- DÉBUT MODIFICATIONS POUR LE SUIVI GA4 DES BOUTONS D'ACTION DE LA CARTE NUMÉRIQUE ---
    $actionButtonsHtmlForDigitalCard = ""
    # Bouton Appeler (Mobile) - Seulement si un mobile GAM est présent
    if ($phoneData.HasMobilePhoneFromGam -and (-not [string]::IsNullOrEmpty($phoneData.RawMobilePhone))) {
        $actionButtonsHtmlForDigitalCard += @"
<a href="tel:$($phoneData.RawMobilePhone)" class="button secondary trackable-action-button" data-analytics-label="Appeler (Mobile)">Appeler (Mobile)</a>
"@
    }
    # Bouton Appeler (Direct)
    if ($phoneData.HasWorkPhoneFromGam -and (-not [string]::IsNullOrEmpty($phoneData.RawWorkPhone))) {
        $actionButtonsHtmlForDigitalCard += @"
<a href="tel:$($phoneData.RawWorkPhone)" class="button secondary trackable-action-button" data-analytics-label="Appeler (Direct)">Appeler (Direct)</a>
"@
    }
    # Bouton Appeler le Centre - Seulement si le standard n'est pas déjà le mobile ou le work phone
    if ($config.DefaultPhoneNumberRaw -ne $phoneData.RawMobilePhone -and $config.DefaultPhoneNumberRaw -ne $phoneData.RawWorkPhone) {
        $actionButtonsHtmlForDigitalCard += @"
<a href="tel:$($config.DefaultPhoneNumberRaw)" class="button secondary trackable-action-button" data-analytics-label="Appeler (Centre)">Appeler le Centre</a>
"@
    }

    $actionButtonsHtmlForDigitalCard += @"
<a href="mailto:$primaryEmail_val" class="button secondary trackable-action-button" data-analytics-label="Envoyer Email">Envoyer un Email</a>
"@
    $actionButtonsHtmlForDigitalCard += @"
<a href="$address_url_maps" target="_blank" class="button secondary trackable-action-button" data-analytics-label="Itinéraire">Itinéraire</a>
"@
    # --- FIN MODIFICATIONS POUR LE SUIVI GA4 DES BOUTONS D'ACTION DE LA CARTE NUMÉRIQUE ---

    # Génération du nom de fichier pour la page de téléchargement et son URL finale
    $downloaderPageFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_').html"
    $downloaderPageUrl_final = "$($githubConfig.PagesBaseUrl)/$($githubConfig.VcardFolderPath)/$downloaderPageFileName"

    # --- Génération des QR Codes et Cartes imprimables (locales) ---
    if ($GeneratePrintQr -or $GeneratePrintableCard -or $GeneratePdfCard) {
        $qrPrintFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_')_print_qrcode.png"
        Write-Host "  - Génération du QR Code pour impression : '$qrPrintFileName' dans '$($config.PrintQrOutputFolder)'..." -ForegroundColor DarkYellow
        $printQrSuccess = Generate-QrCodeFile `
            -QrDataUrl $downloaderPageUrl_final `
            -OutputFileName $qrPrintFileName `
            -OutputFolder $config.PrintQrOutputFolder `
            -QrCodeColors $config.QrCodeColors `
            -PixelsPerModule 100
        if (-not $printQrSuccess) {
            Write-Warning "La génération du QR Code imprimable a échoué. La carte imprimable/PDF pourrait être incomplète."
        }
    }

    # Préparer le contenu HTML pour la carte imprimable/PDF
    $printableCardTemplateContent = Get-TemplateContent($config.PrintableCardTemplatePath)

    # --- Remplacements pour la carte imprimable / PDF (utilisation de la hashtable) ---
    $printableCardReplacements = @{
        '{{digital_card_logo_url_for_print}}' = $config.PrintLogoUrl
        '{{user_full_name}}'                  = "$givenName_val $familyName_val"
        '{{user_title}}'                      = $title_val
        '{{contact_list_html}}'               = $cardContactTextHtmlForDigitalCard # <- UTILISE LA MÊME VARIABLE ICI
        '{{address_label}}'                   = $addressLabelForCard # Maintenu au cas où ce soit utilisé pour un label spécifique
        '{{address_text_print}}'              = ($address_val -replace "`r`n", "<br>") # Maintenu pour backward compatibility si besoin
        '{{website_url}}'                     = $config.WebsiteUrl
        '{{website_display_url}}'             = $config.WebsiteDisplayUrl
        '{{qrcode_print_url}}'                = (Join-Path -Path $config.PrintQrOutputFolder -ChildPath $qrPrintFileName)
    }

    $finalPrintableCardHtml = $printableCardTemplateContent
    foreach ($key in $printableCardReplacements.Keys) {
        $finalPrintableCardHtml = $finalPrintableCardHtml -replace $key, $printableCardReplacements[$key]
    }
    # --- FIN Remplacements pour la carte imprimable / PDF ---

    # Génération du fichier HTML pour la carte imprimable
    if ($GeneratePrintableCard) {
        $printableCardFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_')_print_card.html"
        Write-Host "  - Génération de la carte de visite HTML imprimable : '$printableCardFileName' dans '$($config.PrintableCardOutputFolder)'..." -ForegroundColor DarkYellow

        if (-not (Test-Path $config.PrintableCardOutputFolder)) {
            New-Item -ItemType Directory -Path $config.PrintableCardOutputFolder -Force | Out-Null
        }

        try {
            $outputPath = Join-Path -Path $config.PrintableCardOutputFolder -ChildPath $printableCardFileName
            [System.IO.File]::WriteAllText($outputPath, $finalPrintableCardHtml, [System.Text.Encoding]::UTF8)
            Write-Host "    Carte de visite HTML imprimable générée : $outputPath" -ForegroundColor Green
        } catch {
            Write-Error "Échec de la génération du carte de visite HTML imprimable pour '$printableCardFileName'. Erreur: $($_.Exception.Message)"
        }
    }

    # NOUVEAU BLOC : Génération PDF de la carte imprimable
    if ($GeneratePdfCard) {
        $pdfFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_')_card.pdf"
        $pdfOutputPath = Join-Path -Path $config.PdfCardOutputFolder -ChildPath $pdfFileName
        $tempHtmlForPdfPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "$($pdfFileName).html"

        Write-Host "  - Génération PDF de la carte de visite pour : '$primaryEmail_val'..." -ForegroundColor DarkYellow

        if (-not (Test-Path $config.PdfCardOutputFolder)) {
            New-Item -ItemType Directory -Path $config.PdfCardOutputFolder -Force | Out-Null
        }

        try {
            [System.IO.File]::WriteAllText($tempHtmlForPdfPath, $finalPrintableCardHtml, [System.Text.Encoding]::UTF8)

            $wkhtmltopdfArgs = @(
                '--page-width', '85mm',
                '--page-height', '55mm',
                '--margin-top', '0mm',
                '--margin-bottom', '0mm',
                '--margin-left', '0mm',
                '--margin-right', '0mm',
                '--no-stop-slow-scripts',
                '--enable-local-file-access',
                '--print-media-type',
                '--dpi', '300',
                $tempHtmlForPdfPath,
                $pdfOutputPath
            )

            Write-Host "    Appel wkhtmltopdf : $($config.WkhtmltopdfPath) $($wkhtmltopdfArgs -join ' ')" -ForegroundColor DarkGray
            
            $wkhtmltopdfResult = & $config.WkhtmltopdfPath $wkhtmltopdfArgs 2>&1
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "    Carte de visite PDF générée : $pdfOutputPath" -ForegroundColor Green
            } else {
                Write-Error "Échec de la génération du PDF pour '$pdfFileName'. Erreur: $($wkhtmltopdfResult | Out-String)"
            }

        } catch {
            Write-Error "Erreur lors de l'appel de wkhtmltopdf pour '$pdfFileName'. Erreur: $($_.Exception.Message)"
        } finally {
            Remove-Item -Path $tempHtmlForPdfPath -ErrorAction SilentlyContinue
        }
    }


    # --- Logique pour la Carte de Visite Numérique (et son QR Code) vers GitHub ---
    $qrCodeImageUrl_raw_for_digital_card = ""
    $qrCodeImageUrl_pages_for_digital_card = ""

    if ($AddDigitalCard -and $githubConfig.Token) {
        Write-Host "  - Démarrage de l'upload de la Carte de Visite Numérique vers GitHub pour $primaryEmail_val..." -ForegroundColor Cyan

        $tempQrCodeBaseFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_').png"
        $tempQrCodeFullFileName = "temp_web_qr_$tempQrCodeBaseFileName"
        $tempQrPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath $tempQrCodeFullFileName

        Write-Host "    Tentative de génération temporaire du QR Code web: $($tempQrPath)" -ForegroundColor DarkGray

        $qrCodeGenerationSuccess = Generate-QrCodeFile `
            -QrDataUrl $downloaderPageUrl_final `
            -OutputFileName $tempQrCodeFullFileName `
            -OutputFolder ([System.IO.Path]::GetTempPath()) `
            -QrCodeColors $config.QrCodeColors `
            -PixelsPerModule 20

        if ($qrCodeGenerationSuccess) {
            Write-Host "    Lecture du QR Code temporaire depuis: $($tempQrPath)" -ForegroundColor DarkGray
            $qrCodeBytesForWeb = $null
            try {
                $qrCodeBytesForWeb = [System.IO.File]::ReadAllBytes($tempQrPath)
            } catch {
                Write-Error "Impossible de lire le QR Code temporaire '$tempQrPath'. Erreur: $($_.Exception.Message)"
            } finally {
                Remove-Item -Path $tempQrPath -ErrorAction SilentlyContinue
            }

            if ($qrCodeBytesForWeb -ne $null -and $qrCodeBytesForWeb.Length -gt 0) {
                $uploadResultQrCode = Invoke-GitPublish `
                    -FileName $tempQrCodeBaseFileName `
                    -FileContentBytes $qrCodeBytesForWeb `
                    -FolderPathInRepo $githubConfig.QrcodeFolderPath `
                    -GitHubConfig $githubConfig `
                    -DebugMode:$DebugMode

                if ($uploadResultQrCode) {
                    $qrCodeImageUrl_raw_for_digital_card = $uploadResultQrCode.download_url
                    Write-Host "    QR Code web URL (pour signature mail - raw.githubusercontent.com) : $qrCodeImageUrl_raw_for_digital_card" -ForegroundColor Green

                    $qrCodeImageUrl_pages_for_digital_card = "$($githubConfig.PagesBaseUrl)/$($githubConfig.QrcodeFolderPath)/$tempQrCodeBaseFileName"
                    Write-Host "    QR Code web URL (pour carte numérique - github.io) : $qrCodeImageUrl_pages_for_digital_card" -ForegroundColor Green

                } else {
                    Write-Warning "Échec de l'upload du QR Code pour la carte numérique. Il pourrait ne pas s'afficher."
                }
            } else {
                Write-Warning "Contenu du QR Code web vide. L'upload vers GitHub sera ignoré."
            }
        } else {
            Write-Warning "Échec de la génération du QR Code web temporaire. L'upload vers GitHub sera ignoré."
        }

        # --- IMPORTANT : Mise à jour du HTML de la carte numérique avec la bonne URL QR code et les données téléphone/boutons ---
        $vcfContent = "BEGIN:VCARD`nVERSION:3.0`nN:$($familyName_val);$($givenName_val);;;`nFN:$($givenName_val) $($familyName_val)`nORG:$($config.OrgName)"
        if (-not [string]::IsNullOrEmpty($title_val)) { $vcfContent += "`nTITLE:$title_val" }
        
        # LOGIQUE VCF: Utilise les numéros bruts.
        if (-not [string]::IsNullOrEmpty($phoneData.RawWorkPhone)) { $vcfContent += "`nTEL;type=WORK,voice:$($phoneData.RawWorkPhone)" }
        if (-not [string]::IsNullOrEmpty($phoneData.RawMobilePhone)) { $vcfContent += "`nTEL;type=CELL,voice:$($phoneData.RawMobilePhone)" }
        # Ajouter le standard à la vCard si c'est le numéro principal par défaut, et qu'il n'y a pas de work/mobile de GAM
        if ($phoneData.UsedDefaultPhoneAsPrimary -and -not $phoneData.HasWorkPhoneFromGam -and -not $phoneData.HasMobilePhoneFromGam) {
            $vcfContent += "`nTEL;type=WORK,voice:$($config.DefaultPhoneNumberRaw)"
        }
        
        $vcfContent += "`nEMAIL;type=INTERNET;type=WORK;type=pref:$primaryEmail_val"
        $vcfContent += "`nADR;type=WORK:;;$($address_val -replace "`r`n|`n", '\n');;;;"
        $vcfContent += "`nEND:VCARD"
        
        $vcfEncodedForUrl = [System.Net.WebUtility]::UrlEncode($vcfContent).Replace("+", "%20")
        $vcfDataUrl = "data:text/vcard;charset=utf-8,$vcfEncodedForUrl"
        $vcardDownloadName = "$($givenName_val)_$($familyName_val).vcf".Replace(" ", "_")
        
        $cardTemplateContent_digital = Get-TemplateContent($config.DigitalCardTemplatePath)

        $replacements = @{
            '{{logo_url}}'             = $config.DigitalCardLogoUrl
            '{{user_full_name}}'       = "$givenName_val $familyName_val"
            '{{user_title}}'           = $title_val
            '{{contact_list_html}}'    = $cardContactTextHtmlForDigitalCard
            '{{action_buttons_html}}'  = $actionButtonsHtmlForDigitalCard # <-- Cette variable est maintenant correctement formatée pour le suivi GA4
            '{{vcf_url}}'              = $vcfDataUrl
            '{{vcf_download_name}}'    = $vcardDownloadName
            '{{qrcode_image_url}}'     = $qrCodeImageUrl_pages_for_digital_card
            '{{digital_card_page_url}}'= $downloaderPageUrl_final
            '{{address_label}}'        = $addressLabelForCard
            '{{address_texte}}'        = $addressForDigitalCard
            '{{website_html_for_card}}'= ""
        }

        $downloaderPageContent = $cardTemplateContent_digital
        foreach ($key in $replacements.Keys) {
            $downloaderPageContent = $downloaderPageContent -replace $key, $replacements[$key]
        }

        $downloaderPageBytes = [System.Text.Encoding]::UTF8.GetBytes($downloaderPageContent)
        $uploadResultDownloader = Invoke-GitPublish `
            -FileName $downloaderPageFileName `
            -FileContentBytes $downloaderPageBytes `
            -FolderPathInRepo $githubConfig.VcardFolderPath `
            -GitHubConfig $githubConfig `
            -DebugMode:$DebugMode

        if ($uploadResultDownloader) {
            Write-Host "    Digital Card page public URL: $downloaderPageUrl_final" -ForegroundColor Green
        } else {
            Write-Warning "Échec de l'upload de la page de la carte numérique."
        }
    }

    # --- LOGIQUE POUR LE BLOC QR CODE DANS LA SIGNATURE MAIL ---
    $digital_card_html_block = ""
    if ($AddDigitalCard -and (-not [string]::IsNullOrEmpty($qrCodeImageUrl_raw_for_digital_card))) {
        $digital_card_html_block = @"
<table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="padding-top:10px;"><tr>
<td style="width:100%; text-align: right; vertical-align: middle;">
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="display: inline-block;">
        <tr>
            <td style="padding-right:15px; vertical-align:middle; text-align: right;">
                <p style="margin:0;padding:0;font-size:9pt;color:#555555;line-height:1.3">Scannez-moi ou <a href="$downloaderPageUrl_final" target="_blank" style="color:#068FD0;text-decoration:underline">cliquez ici</a><br>pour ma carte de visite numérique.</p>
            </td>
            <td style="width:96px; vertical-align:middle;">
                <a href="$downloaderPageUrl_final" target="_blank" style="text-decoration:none;"><img src="$qrCodeImageUrl_raw_for_digital_card" width="80" style="width:80px;height:80px;display:block;border:0;" alt="QR Code"/></a>
            </td>
        </tr>
    </table>
</td>
</tr></table>
"@
    } else {
        if ($AddDigitalCard) {
            Write-Warning "Le bloc QR Code pour la signature ne sera pas généré (URL QR code manquante)."
        } else {
            Write-Host "Le bloc QR Code pour la signature est désactivé (-AddDigitalCard non spécifié)." -ForegroundColor DarkGray
        }
    }

    # --- Préparation de la SIGNATURE GMAIL ---
    $logPhoneLines = @();

    # Construire phoneBlockHtmlForSignatureFinal de manière explicite
    $phoneBlockHtmlForSignatureFinal = ""
    # Ligne mobile (si présente)
    if (-not [string]::IsNullOrEmpty($phoneData.RawMobilePhone)) {
        $phoneBlockHtmlForSignatureFinal += "Mobile : <a href=`"tel:$($phoneData.RawMobilePhone)`" style=`"color: #555555; text-decoration: underline;`">$($phoneData.MobilePhoneDisplayForTemplates)</a><br>"
    }
    # Ligne directe (si présente)
    if (-not [string]::IsNullOrEmpty($phoneData.RawWorkPhone)) {
        $phoneBlockHtmlForSignatureFinal += "Ligne directe : <a href=`"tel:$($phoneData.RawWorkPhone)`" style=`"color: #555555; text-decoration: underline;`">$($phoneData.WorkPhoneDisplayForTemplates)</a><br>"
    }
    # Standard (si pas de work phone)
    if (-not $phoneData.HasWorkPhoneFromGam) {
        # Ajouter le standard SEULEMENT si ce n'est PAS déjà le numéro mobile (pour éviter la duplication si le mobile est le seul numéro GAM)
        if ($config.DefaultPhoneNumberRaw -ne $phoneData.RawMobilePhone) {
            $phoneBlockHtmlForSignatureFinal += "Téléphone (Centre) : <a href=`"tel:$($config.DefaultPhoneNumberRaw)`" style=`"$linkStyleGeneral`">$($config.DefaultPhoneNumberDisplay)</a><br>"
        }
    }


    # Logs pour la console (rappel, ceci n'affecte pas le HTML, juste l'affichage dans la console)
    if ($phoneData.RawPrimaryDisplayPhone) {
        $logPhoneLines += "Ligne principale (affichée) : $($phoneData.RawPrimaryDisplayPhone)"
    }
    if ($phoneData.MobilePhoneDisplayForTemplates -and ($phoneData.RawMobilePhone -ne $phoneData.RawPrimaryDialPhone)) {
        $logPhoneLines += "Mobile (affiché) : $($phoneData.MobilePhoneDisplayForTemplates)"
    }
    if ($phoneData.UsedDefaultPhoneAsPrimary -eq $false -and $phoneData.HasMobilePhoneFromGam -eq $true -and $phoneData.HasWorkPhoneFromGam -eq $false) {
        if ($config.DefaultPhoneNumberRaw -ne $phoneData.RawMobilePhone) {
            $logPhoneLines += "Téléphone (standard - ajouté) : $($config.DefaultPhoneNumberDisplay)"
        }
    }
    if ($phoneData.UsedDefaultPhoneAsPrimary -eq $true) {
        $logPhoneLines += "Téléphone (standard - principal) : $($config.DefaultPhoneNumberDisplay)"
    }

    Write-Host "  - Prénom      : $givenName_val" -ForegroundColor Gray
    Write-Host "  - Nom         : $familyName_val" -ForegroundColor Gray
    Write-Host "  - Titre       : $(if ([string]::IsNullOrEmpty($title_val)) { '(aucun)' } else { $title_val })" -ForegroundColor Gray
    Write-Host "  - Adresse     : $addressForSignature" -ForegroundColor Gray
    if ($logPhoneLines.Count -gt 0) { foreach($line in $logPhoneLines){ Write-Host "  - Téléphone   : $line" -ForegroundColor Gray } } else { Write-Host "  - Téléphone   : (aucun)" -ForegroundColor Gray }

    $functionLineConditional = ""; if ($title_val -ne "") { $functionLineConditional = "<span style=`"font-size: 10pt; color: #555555;`">" + $title_val.Trim() + "</span>" }

    $htmlTemplateContent = Get-TemplateContent($config.SignatureTemplatePath)

    $signatureReplacements = @{
        '{{digital_card_html_block}}' = $digital_card_html_block
        '{{givenName}}'               = $givenName_val
        '{{familyName}}'              = $familyName_val
        '{{functionLineConditional}}' = $functionLineConditional
        '{{primaryEmail}}'            = $primaryEmail_val
        '{{phoneBlock}}'              = $phoneBlockHtmlForSignatureFinal
        '{{address_texte}}'           = $addressForSignature
        '{{address_url_maps}}'        = $address_url_maps
        '{{logo_url}}'                = $config.SignatureLogoUrl
        '{{website_url}}'             = $config.WebsiteUrl
        '{{website_display_url}}'     = $config.WebsiteDisplayUrl
        '{{org_name}}'                = $config.OrgName
		# --- NOUVELLES LIGNES À AJOUTER ICI ---
        '{{gcsms_logo_url}}'          = $config.GcsmsLogoUrl
        '{{facebook_logo_url}}'       = $config.FacebookLogoUrl
        '{{linkedin_logo_url}}'       = $config.LinkedinLogoUrl
        '{{facebook_page_url}}'       = $config.FacebookPageUrl
        '{{linkedin_company_url}}'    = $config.LinkedinCompanyUrl
        # --- FIN DES NOUVELLES LIGNES ---
    }

    $finalSignatureHtml = $htmlTemplateContent
    foreach ($key in $signatureReplacements.Keys) {
        $finalSignatureHtml = $finalSignatureHtml -replace $key, $signatureReplacements[$key]
    }

    if ($DebugMode) { Write-Host "Debug: finalSignatureHtml ($($finalSignatureHtml.Length) chars):`n$finalSignatureHtml" -ForegroundColor DarkYellow }

    $tempSignaturePath = Join-Path -Path $config.ProjectRoot -ChildPath "temp_sig_$($primaryEmail_val.Replace('@','_')).html"

    Write-Host "  - Vérification de la signature actuelle sur Google..." -ForegroundColor DarkGray
    $currentSignatureHtml = & $config.GamPath user "$primaryEmail_val" print signature | Out-String
    $newSigNormalized = ($finalSignatureHtml -replace '\s' -replace ' ', '').Trim()
    $currentSigNormalized = ($currentSignatureHtml -replace '\s' -replace ' ', '').Trim()

    if ($newSigNormalized -eq $currentSigNormalized) {
        Write-Host "  - La signature est déjà à jour. Mise à jour ignorée." -ForegroundColor Green
    } else {
        Write-Host "  - Signature mise à jour détectée. Application en cours..." -ForegroundColor DarkCyan
        [System.IO.File]::WriteAllText($tempSignaturePath, $finalSignatureHtml, [System.Text.Encoding]::UTF8)
        Write-Host "Application de la signature pour $primaryEmail_val..." -ForegroundColor DarkCyan
        & $config.GamPath user "$primaryEmail_val" signature file "$tempSignaturePath" html
        Remove-Item -Path $tempSignaturePath -ErrorAction SilentlyContinue
    }
}

[Console]::OutputEncoding = $originalEncoding
Write-Host "Processus d'application des signatures terminé." -ForegroundColor Green
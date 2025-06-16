# set_user_signatures.ps1 (v46.1 - Carte de Visite Imprimable)
#
param(
    [string]$SingleUserEmail = "",
    [switch]$IncludeSuspended,
    [switch]$AddDigitalCard,
    [switch]$GeneratePrintQr,
    [switch]$GeneratePrintableCard,
    [switch]$ShowHelp,
    [switch]$DebugMode 
)

# NOUVEAU : Définir et afficher la version du script APRES le bloc param
$script:ScriptVersion = "v45.41"
Write-Host "Démarrage du script : set_user_signatures.ps1 ($script:ScriptVersion)" -ForegroundColor Green

if ($ShowHelp) {
    $helpText = @"
NOM:
    set_user_signatures.ps1

SYNOPSIS:
    Met à jour les signatures email et peut générer une carte de visite numérique complète (vCard + QR Code interactif).

SYNTAXE:
    .\set_user_signatures.ps1 [-SingleUserEmail <string>] [-IncludeSuspended] [-AddDigitalCard] [-GeneratePrintQr] [-GeneratePrintableCard] [-ShowHelp] [-DebugMode]

DESCRIPTION:
    Ce script automatise la mise à jour des signatures Gmail via GAM. Il est optimisé pour ne pas effectuer de mises à jour inutiles.

    - Mode standard : Met à jour la signature principale de l'utilisateur.
    
    - Mode Carte de Visite (-AddDigitalCard) : Génère une page web professionnelle (hébergée sur GitHub Pages) qui contient un lien de
      déchargement direct pour la vCard de l'utilisateur. Pour assurer la compatibilité, la vCard est encodée
      directement dans le lien de téléchargement (méthode Data-URL).
      Nouveauté : La carte inclut désormais un QR Code interactif qui peut être agrandi pour un partage facile,
      et le label de l'adresse (ex: 'Siège Social') est dynamiquement affiché.

PARAMÈTRES:
    -SingleUserEmail <string>
        Spécifie l'adresse email d'un seul utilisateur à mettre à jour.

    -IncludeSuspended
        Commutateur. Si présent, le script mettra à jour TOUS les utilisateurs, y compris les comptes suspendus.

    -AddDigitalCard
        Commutateur. Si présent, active la génération de la carte de visite numérique avec QR Code.
    
    -GeneratePrintQr
        Commutateur. Si présent, génère des QR Codes haute résolution pour l'impression dans un dossier local.
        Ces fichiers NE sont PAS poussés vers GitHub.
    
    -GeneratePrintableCard
        Commutateur. Si présent, génère un fichier HTML de carte de visite optimisé pour l'impression locale.
        Ce fichier est destiné à être ouvert dans un navigateur puis "Imprimé au format PDF".

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
    .\set_user_signatures.ps1 -GeneratePrintQr

    # Génère une carte de visite HTML imprimable pour un utilisateur spécifique
    .\set_user_signatures.ps1 -SingleUserEmail "s.gille@cjml.fr" -GeneratePrintableCard

    # Met à jour les signatures, les cartes de visite ET génère des QR Codes/Cartes imprimables pour un utilisateur
    .\set_user_signatures.ps1 -SingleUserEmail "s.gille@cjml.fr" -AddDigitalCard -GeneratePrintQr -GeneratePrintableCard

    # Débogage : Exécute le script avec les paramètres habituels et affiche les infos de débogage
    .\set_user_signatures.ps1 -SingleUserEmail "s.gille@cjml.fr" -AddDigitalCard -DebugMode
"@
    Write-Host $helpText
    return
}

# --- Configuration ---
$projectRoot = "C:\GAMWork\Signatures"
$printQrOutputFolder = "C:\GAMWork\PrintQrCodes"
$printableCardOutputFolder = "C:\GAMWork\PrintableCards"

$gamPath = "C:\GAM7\gam.exe"
$signatureTemplatePath = Join-Path -Path $projectRoot -ChildPath "signature_template.html"
$digitalCardTemplatePath = Join-Path -Path $projectRoot -ChildPath "digital_card_template.html"
$printableCardTemplatePath = Join-Path -Path $projectRoot -ChildPath "printable_business_card_template.html"

$signatureLogoUrl = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/Logo-CJML.png"
$digitalCardLogoUrl = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/logo-horizontal.jpg"
$printLogoUrl = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/logo-horizontal.jpg" 

$orgName = "Centre Jean-Marie LARRIEU"

$defaultPhoneNumberRaw = "+33562913250" 
$defaultPhoneNumberDisplay = "05 62 91 32 50" 
$defaultAddress = @"
414 Rue du Layris
65710 CAMPAN
"@

$WebsiteUrl = "http://www.cjml.fr"
$websiteDisplayUrl = ""
if (-not [string]::IsNullOrEmpty($WebsiteUrl)) {
    $websiteDisplayUrl = $WebsiteUrl -replace "^https?:\/\/(www\.)?", ""
    if ($websiteDisplayUrl -like "cjml.fr*") {
        $websiteDisplayUrl = "www." + $websiteDisplayUrl
    }
}

# --- CONFIGURATION GITHUB & QR CODE ---
try {
    $tokenPath = Join-Path -Path $projectRoot -ChildPath "github_token.txt"
    $githubToken = Get-Content -Path $tokenPath -Raw -ErrorAction Stop
} catch {
    Write-Warning "Fichier 'github_token.txt' introuvable dans $projectRoot. La fonction -AddDigitalCard sera désactivée si utilisée."
    $githubToken = $null
}
$githubUserOrOrg = "Centre-Jean-Marie-LARRIEU"
$githubRepo = "assets-cjml"
$vcardFolderPath = "vcards"
$qrcodeFolderPath = "qrcodes"
$qrCodeDllPath = Join-Path -Path $projectRoot -ChildPath "QRCoder.dll"
$githubPagesBaseUrl = "https://Centre-Jean-Marie-LARRIEU.github.io/assets-cjml"
$qrCodeBlue = [byte[]](6, 143, 208)
$qrCodeWhite = [byte[]](255, 255, 255)
# --- FIN CONFIGURATION ---

$mainDomain = "cjml.fr"
$excludeDomain = "eleves.cjml.fr"

$originalEncoding = [Console]::OutputEncoding 
chcp.com 65001 | Out-Null 
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8 


try {
    Add-Type -Path $qrCodeDllPath -ErrorAction Stop
} catch {
    Write-Error "Impossible de charger la bibliothèque QRCoder.dll. Assurez-vous que le fichier se trouve bien à l'emplacement : $qrCodeDllPath"
    exit 1
}

function Publish-FileToGitHub {
    param([string]$FileName, [byte[]]$FileContentBytes, [string]$FolderPathInRepo)
    $apiUrl = "https://api.github.com/repos/$githubUserOrOrg/$githubRepo/contents/$FolderPathInRepo/$FileName"
    $headers = @{ "Authorization" = "Bearer $githubToken"; "Accept" = "application/vnd.github.com.v3+json" } 
    
    $sha = $null
    try {
        $existingFile = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers -ErrorAction Stop
        if ($existingFile) {
            $sha = $existingFile.sha
            $header = "blob $($FileContentBytes.Length)`0"; $headerBytes = [System.Text.Encoding]::UTF8.GetBytes($header)
            $combinedBytes = $headerBytes + $FileContentBytes; $sha1 = New-Object System.Security.Cryptography.SHA1Managed
            $localSha = [System.BitConverter]::ToString($sha1.ComputeHash($combinedBytes)).Replace("-", "").ToLower()
            # DÉCOMMENTEZ LA LIGNE SUIVANTE POUR FORCER LA MISE À JOUR LORS DU DÉBOGAGE. RECOMMENTEZ APRÈS.
            # if ($DebugMode) { # Force pour tout si DebugMode est actif.
                # Write-Host "    - DEBUG : Forçage de la mise à jour pour $FileName" -ForegroundColor Cyan
                # $localSha = "FORCE_UPDATE_" + (Get-Random).ToString() 
            # }
            
            if ($localSha -eq $sha) {
                Write-Host "    - Contenu identique pour '$FileName' sur GitHub. Aucune mise à jour nécessaire." -ForegroundColor Green
                return $existingFile 
            }
            Write-Host "    - Fichier existant détecté sur GitHub. Préparation de la mise à jour." -ForegroundColor DarkGray
        }
    }
    catch [System.Net.WebException] {
        if ($_.Exception.Response.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) { Write-Host "    - Fichier absent sur GitHub. Préparation pour la création." -ForegroundColor DarkGray }
        else { Write-Warning "      Erreur web inattendue : $($_.Exception.Message)" }
    }
    catch { Write-Warning "      Erreur inattendue : $($_.Exception.Message)" }
    $contentBase64 = [System.Convert]::ToBase64String($FileContentBytes)
    $body = @{ message = "Automated update of $FileName"; content = $contentBase64 }
    if ($sha) { $body.Add("sha", $sha) }
    try {
        $uploadResult = Invoke-RestMethod -Uri $apiUrl -Method Put -Headers $headers -Body ($body | ConvertTo-Json) -ContentType "application/json"
        return $uploadResult.content
    } catch {
        Write-Error "Échec de l'upload sur GitHub pour $FileName. Erreur: $($_.Exception.Message)"
        return $null
    }
}

function Generate-PrintQrCode {
    param([string]$QrDataUrl, [string]$OutputFileName, [string]$OutputFolder)

    Write-Host "  - Génération du QR Code pour impression : '$OutputFileName' dans '$OutputFolder'..." -ForegroundColor DarkYellow
    
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }

    try {
        $pixelsPerModule = 100 
        $qrGenerator = New-Object QRCoder.QRCodeGenerator
        $qrCodeData = $qrGenerator.CreateQrCode($QrDataUrl, [QRCoder.QRCodeGenerator+ECCLevel]::Q) 
        $qrCode = New-Object QRCoder.PngByteQRCode($qrCodeData)
        
        $qrCodeBytes = $qrCode.GetGraphic($pixelsPerModule, $qrCodeBlue, $qrCodeWhite)
        $outputPath = Join-Path -Path $OutputFolder -ChildPath $OutputFileName
        
        [System.IO.File]::WriteAllBytes($outputPath, $qrCodeBytes)
        Write-Host "    QR Code imprimable généré : $outputPath" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Échec de la génération du QR Code imprimable pour '$OutputFileName'. Erreur: $($_.Exception.Message)"
        return $false
    }
}

function Generate-PrintableBusinessCard {
    param([hashtable]$UserData, [string]$OutputFileName, [string]$OutputFolder, [string]$TemplatePath, [string]$PrintLogoUrl, [string]$WebsiteUrl, [string]$WebsiteDisplayUrl)

    Write-Host "  - Génération de la carte de visite HTML imprimable : '$OutputFileName' dans '$OutputFolder'..." -ForegroundColor DarkYellow
    
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Nul
    }

    try {
        $cardTemplateContent = Get-Content -Path $TemplatePath -Encoding UTF8 -Raw
        $cardTemplateContent = $cardTemplateContent.TrimStart([char]65279, [char]22)

        $qrPrintFileName = "$($UserData.primaryEmail_val -replace '[^a-zA-Z0-9]','_')_print_qrcode.png"
        $qrCodePrintPath = Join-Path -Path $printQrOutputFolder -ChildPath $qrPrintFileName
        
        if (-not (Test-Path $qrCodePrintPath)) {
            Write-Warning "Le QR code imprimable nécessaire pour la carte n'existe pas. Tentative de génération..."
            $downloaderPageUrl_final_for_print = "$githubPagesBaseUrl/$vcardFolderPath/$($UserData.primaryEmail_val -replace '[^a-zA-Z0-9]','_').html" 
            if (-not (Generate-PrintQrCode -QrDataUrl $downloaderPageUrl_final_for_print -OutputFileName $qrPrintFileName -OutputFolder $printQrOutputFolder)) {
                throw "Impossible de générer le QR code imprimable nécessaire pour la carte."
            }
        }
        
        # Préparer les données téléphoniques pour le template imprimable (texte brut)
        $finalWorkPhoneDisplayForPrint = $UserData.phoneData.WorkPhoneDisplayForTemplates
        if([string]::IsNullOrEmpty($finalWorkPhoneDisplayForPrint)) { $finalWorkPhoneDisplayForPrint = "N/A" } 

        $finalMobilePhoneDisplayForPrint = $UserData.phoneData.MobilePhoneDisplayForTemplates
        if([string]::IsNullOrEmpty($finalMobilePhoneDisplayForPrint)) { $finalMobilePhoneDisplayForPrint = "N/A" } 
        
        $finalCardHtml = $cardTemplateContent 
        $finalCardHtml = $finalCardHtml -replace '\{\{digital_card_logo_url_for_print\}\}', $PrintLogoUrl 
        $finalCardHtml = $finalCardHtml -replace '\{\{user_full_name\}\}', "$($UserData.givenName_val) $($UserData.familyName_val)"
        $finalCardHtml = $finalCardHtml -replace '\{\{user_title\}\}', $UserData.title_val
        $finalCardHtml = $finalCardHtml -replace '\{\{phone_work_display\}\}', $finalWorkPhoneDisplayForPrint
        $finalCardHtml = $finalCardHtml -replace '\{\{phone_mobile_display\}\}', $finalMobilePhoneDisplayForPrint
        $finalCardHtml = $finalCardHtml -replace '\{\{primary_email_raw\}\}', $UserData.primaryEmail_val
        $finalCardHtml = $finalCardHtml -replace '\{\{address_label\}\}', $UserData.addressLabelForCard
        $finalCardHtml = $finalCardHtml -replace '\{\{address_text_print\}\}', ($UserData.address_val -replace "`r`n", "<br>") 
        $finalCardHtml = $finalCardHtml -replace '\{\{website_url\}\}', $WebsiteUrl
        $finalCardHtml = $finalCardHtml -replace '\{\{website_display_url\}\}', $websiteDisplayUrl
        $finalCardHtml = $finalCardHtml -replace '\{\{qrcode_print_url\}\}', $qrCodePrintPath


        $outputPath = Join-Path -Path $OutputFolder -ChildPath $OutputFileName
        
        [System.IO.File]::WriteAllText($outputPath, $finalCardHtml, [System.Text.Encoding]::UTF8)
        Write-Host "    Carte de visite imprimable générée : $outputPath" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Échec de la génération du carte de visite imprimable pour '$OutputFileName'. Erreur: $($_.Exception.Message)"
        return $false
    }
}


# --- Récupération des utilisateurs ---
$usersToProcess = @()
if ($ShowHelp) {
    # If showing help, don't try to get users
} else { # Only try to get users if not showing help
    $fieldsToGet = 'primaryEmail,name,organizations,phones,suspended,addresses'
    if (-not [string]::IsNullOrEmpty($SingleUserEmail)) {
        Write-Host "--- MODE UTILISATEUR UNIQUE: Cible l'utilisateur '$SingleUserEmail' ---" -ForegroundColor Yellow
        $gamOutput = & $gamPath print users query "email='$SingleUserEmail'" fields $fieldsToGet
        if ($gamOutput -and $gamOutput.Count -gt 1) { $usersToProcess = $gamOutput | ConvertFrom-Csv }
        else { Write-Error "Impossible de récupérer les informations pour l'utilisateur '$SingleUserEmail'." }
    } else {
        $gamArguments = @('print', 'users'); if (-not $IncludeSuspended) { $gamArguments += 'query', 'isSuspended=False' }
        $gamArguments += 'fields', $fieldsToGet; $gamRawOutput = & $gamPath $gamArguments; $allGSuiteUsers = $gamRawOutput | ConvertFrom-Csv
        $usersToProcess = $allGSuiteUsers | Where-Object { $_.primaryEmail -like "*@$mainDomain" -and $_.primaryEmail -notlike "*@$excludeDomain" }
    }
}

if ($ShowHelp) { # Exit if ShowHelp was requested
    return
}
if ($usersToProcess.Count -eq 0) { Write-Host "Aucun utilisateur trouvé à traiter. Quitte le script."; exit 0 }
Write-Host "Found $($usersToProcess.Count) user(s) to process." -ForegroundColor Cyan

foreach ($user in $usersToProcess) {
    if ($user -eq $null) { Write-Error "Skipping null user object."; continue } 
    
    $primaryEmail_val = if ($user.primaryEmail) { $user.primaryEmail } else { "" }
    $givenName_val = if ($user."name.givenName") { $user."name.givenName" } else { "" } 
    $familyName_val = if ($user."name.familyName") { $user."name.familyName" } else { "" }
    $title_val = if ($user."organizations.0.title") { $user."organizations.0.title" } else { "" }
    
    $userRawAddress = "" 
    $address_val = $defaultAddress 

    for ($i = 0; $i -lt 5; $i++) {
        $typeProperty = "addresses.$i.type"; $formattedProperty = "addresses.$i.formatted"
        if (($user.PSObject.Properties.Name -contains $typeProperty) -and ($user.$typeProperty -eq 'work')) {
            if (($user.PSObject.Properties.Name -contains $formattedProperty) -and (-not [string]::IsNullOrEmpty($user.$formattedProperty))) {
                $address_val = $user.$formattedProperty 
                $userRawAddress = $address_val 
                break
            }
        }
    }
    
    $addressForSignature = ($address_val -replace "`r`n", " - " -replace "`n", " - ").Trim()
    $addressForDigitalCard = ($address_val -replace "`r`n", "<br>").Trim()
    $address_url_maps = "https://www.google.com/maps/search/?api=1&query=" + [System.Net.WebUtility]::UrlEncode($addressForSignature)

    if ([string]::Compare($userRawAddress.Trim(), ($defaultAddress -replace "`r`n", "`n").Trim(), $true) -eq 0 -or [string]::IsNullOrEmpty($userRawAddress)) {
        $addressLabelForCard = "Siège Social"
    } else {
        $addressLabelForCard = "Adresse du Bureau"
    }

    Write-Host "--- Processing user: $primaryEmail_val (Suspended: $($user.suspended)) ---" -ForegroundColor Cyan
    
    # --- DÉBUT : Préparation Robuste et Finale des Numéros de Téléphone pour TOUS les Templates ---
    $phonesByType = @{ work = @(); mobile = @() } 
    for ($i = 0; $i -lt 5; $i++) {
        $phoneValueProperty = "phones.$i.value"; $phoneTypeProperty = "phones.$i.type"
        if ($user.PSObject.Properties.Name -contains $phoneValueProperty -and -not [string]::IsNullOrEmpty($user.$phoneValueProperty)) {
            $phoneValue = $user.$phoneValueProperty; $formattedDisplayValue = $phoneValue
            # Formatage du numéro de GAM: +33 -> 0 et espace si 10 chiffres
            if ($phoneValue -match '^\+33[1-9]\d{8}$') { 
                $localNumber = $phoneValue -replace '^\+33', '0'
                $formattedDisplayValue = $localNumber -replace '^(0\d)(\d{2})(\d{2})(\d{2})(\d{2})$', '$1 $2 $3 $4 $5' 
            } elseif ($phoneValue -match '^[0-9]{9,10}$' -and -not ($phoneValue -match '^\+')) { # Gère les 9/10 chiffres sans préfixe +33
                $formattedDisplayValue = $phoneValue -replace '^(0\d)(\d{2})(\d{2})(\d{2})(\d{2})$', '$1 $2 $3 $4 $5'
            }

            $phoneType = if ($user.PSObject.Properties.Name -contains $phoneTypeProperty) { $user.$phoneTypeProperty.ToLower() } else { "work" }
            if ($phonesByType.ContainsKey($phoneType)) { $phonesByType[$phoneType] += @{ Display = $formattedDisplayValue; Raw = $phoneValue } }
        }
    }

    $phoneData = @{
        WorkPhoneHtmlForSignature = ""; 
        MobilePhoneHtmlForSignature = ""; 
        WorkPhoneDisplayForTemplates = ""; 
        MobilePhoneDisplayForTemplates = ""; 

        RawWorkPhone = ""; 
        RawMobilePhone = ""; 
        RawDefaultPhone = $defaultPhoneNumberRaw; # Le "+33..."
        FormattedDefaultPhoneDisplay = $defaultPhoneNumberDisplay; # Le "05 62..."

        HasWorkPhoneFromGam = $false; # Indique si un vrai numéro de travail de GAM est présent
        HasMobilePhoneFromGam = $false; # Indique si un vrai numéro mobile de GAM est présent
        UsedDefaultPhoneAsWorkPhone = $false; # Indique si le standard remplace la ligne directe
    }

    $linkStyleSignature = "color: #555555; text-decoration: underline;" 
    
    # 1. Traitement des numéros de Ligne directe (provenant de GAM)
    if ($phonesByType.work.Count -gt 0) {
        $phoneData.HasWorkPhoneFromGam = $true
        $phoneData.WorkPhoneDisplayForTemplates = ($phonesByType.work | ForEach-Object { $_.Display }) -join ', '
        $phoneData.RawWorkPhone = ($phonesByType.work | Select-Object -First 1).Raw # Garde le format RAW de GAM
        # S'assure que le RawWorkPhone est au format international pour tel: si ce n'est pas déjà le cas (+33 ou 0033)
        if ($phoneData.RawWorkPhone -match '^0[1-9]\d{8}$') { $phoneData.RawWorkPhone = "+33" + $phoneData.RawWorkPhone.Substring(1) } 
        
        $phoneData.WorkPhoneHtmlForSignature = "Ligne directe : <a href=`"tel:$($phoneData.RawWorkPhone)`" style=`"$linkStyleSignature`">$($phoneData.WorkPhoneDisplayForTemplates)</a><br>"
    }

    # 2. Traitement des numéros Mobile (provenant de GAM)
    if ($phonesByType.mobile.Count -gt 0) {
        $phoneData.HasMobilePhoneFromGam = $true
        $phoneData.MobilePhoneDisplayForTemplates = ($phonesByType.mobile | ForEach-Object { $_.Display }) -join ', '
        $phoneData.RawMobilePhone = ($phonesByType.mobile | Select-Object -First 1).Raw 
        # S'assure que le RawMobilePhone est au format international pour tel: si ce n'est pas déjà le cas
        if ($phoneData.RawMobilePhone -match '^0[6-7]\d{8}$') { $phoneData.RawMobilePhone = "+33" + $phoneData.RawMobilePhone.Substring(1) } 
        
        $phoneData.MobilePhoneHtmlForSignature = "Mobile : <a href=`"tel:$($phoneData.RawMobilePhone)`" style=`"$linkStyleSignature`">$($phoneData.MobilePhoneDisplayForTemplates)</a><br>"
    }

    # 3. Logique pour le numéro de téléphone du Standard (le standard doit être inclus si AUCUNE vraie ligne directe n'est trouvée)
    # Ceci est la logique qui décide si le standard remplace la ligne directe.
    if (-not $phoneData.HasWorkPhoneFromGam) { # Condition : Pas de ligne de travail de GAM
        $phoneData.WorkPhoneDisplayForTemplates = $phoneData.FormattedDefaultPhoneDisplay # Standard formaté pour affichage
        $phoneData.RawWorkPhone = $phoneData.RawDefaultPhone # Standard brut pour tel: et vcf
        $phoneData.WorkPhoneHtmlForSignature = "Téléphone : <a href=`"tel:$($phoneData.RawDefaultPhone)`" style=`"$linkStyleSignature`">$($phoneData.WorkPhoneDisplayForTemplates)</a><br>"
        
        $phoneData.HasWorkPhoneFromGam = $true # On le marque comme "trouvé" (le standard) pour le reste du script
        $phoneData.UsedDefaultPhoneForWork = $true # Indique que c'est le standard
    }
    # --- FIN : Préparation Robuste et Finale des Numéros de Téléphone pour TOUS les Templates ---

    # --- DÉBOGAGE : Affichage des valeurs de phoneData ---
    if ($DebugMode) {
        Write-Host "--- DÉBOGAGE : phoneData après traitement ---" -ForegroundColor Yellow
        $phoneData | Format-List -Force
        Write-Host "--- FIN DÉBOGAGE phoneData ---" -ForegroundColor Yellow
    }
    # --- FIN DÉBOGAGE ---


    # --- DÉBUT : VARIABLES HTML POUR LA CARTE NUMÉRIQUE ET LA SIGNATURE ---
    # Ces variables sont préparées une seule fois, puis utilisées dans les blocs respectifs
    $cardContactTextHtmlForDigitalCard = ""
    # Ligne directe ou Téléphone standard
    if($phoneData.HasWorkPhoneFromGam) { 
         $label = if($phoneData.UsedDefaultPhoneForWork){"Téléphone"}else{"Ligne directe"}
         $cardContactTextHtmlForDigitalCard += "<div class=`"contact-item`"><span class=`"label`">$label</span><a href=`"tel:$($phoneData.RawWorkPhone)`">$($phoneData.WorkPhoneDisplayForTemplates)</a></div>"
    }
    # Mobile
    if($phoneData.HasMobilePhoneFromGam) { 
        $cardContactTextHtmlForDigitalCard += "<div class=`"contact-item`"><span class=`"label`">Mobile</span><a href=`"tel:$($phoneData.RawMobilePhone)`">$($phoneData.MobilePhoneDisplayForTemplates)</a></div>"
    }
    # Email
    $cardContactTextHtmlForDigitalCard += "<div class=`"contact-item`"><span class=`"label`">Email</span><a href=`"mailto:$primaryEmail_val`">$primaryEmail_val</a></div>"

    # AJOUT DU SITE WEB À LA CARTE NUMÉRIQUE
    if (-not [string]::IsNullOrEmpty($WebsiteUrl)) {
        $cardContactTextHtmlForDigitalCard += @"
<div class="contact-item">
    <span class="label">Site Web</span>
    <a href="$WebsiteUrl" target="_blank" rel="noopener noreferrer" style="color: var(--primary-blue); text-decoration: underline;">$websiteDisplayUrl</a>
</div>
"@
    }

    $actionButtonsHtmlForDigitalCard = ""
    # Bouton Appeler (Mobile)
    if ($phoneData.HasMobilePhoneFromGam) { $actionButtonsHtmlForDigitalCard += "<a href=`"tel:$($phoneData.RawMobilePhone)`" class=`"button secondary`">Appeler (Mobile)</a>" }
    
    # Bouton Appeler (Direct) ou Appeler le Centre
    if ($phoneData.HasWorkPhoneFromGam) {
        if ($phoneData.UsedDefaultPhoneForWork) { 
            # Le bouton "Appeler le Centre" doit apparaître s'il n'y a pas d'autre numéro de mobile valide
            if (-not $phoneData.HasMobilePhoneFromGam) { 
                 $actionButtonsHtmlForDigitalCard += "<a href=`"tel:$($phoneData.RawDefaultPhone)`" class=`"button secondary`">Appeler le Centre</a>"
            }
        } else { # C'est une vraie ligne directe
            $actionButtonsHtmlForDigitalCard += "<a href=`"tel:$($phoneData.RawWorkPhone)`" class=`"button secondary`">Appeler (Direct)</a>"
        }
    }
    $actionButtonsHtmlForDigitalCard += "<a href=`"mailto:$primaryEmail_val`" class=`"button secondary`">Envoyer un Email</a>"
    $actionButtonsHtmlForDigitalCard += "<a href=`"$address_url_maps`" target=`"_blank`" class=`"button secondary`">Itinéraire</a>"

    $websiteHtmlForDigitalCard = "" # Cette variable n'est plus directement utilisée ici, le HTML est intégré ci-dessus
    # --- FIN : VARIABLES HTML POUR LA CARTE NUMÉRIQUE ET LA SIGNATURE ---


    $digital_card_html_block = "" # Variable pour le bloc QR code dans la signature

    $downloaderPageFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_').html"
    $downloaderPageUrl_final = "$githubPagesBaseUrl/$vcardFolderPath/$downloaderPageFileName"

    # --- DÉBUT : LOGIQUE DE GÉNÉRATION DES QR CODES ET CARTES IMPRIMABLES (Dépend de $downloaderPageUrl_final) ---
    if ($GeneratePrintQr -or $GeneratePrintableCard) {
        $qrPrintFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_')_print_qrcode.png"
        Generate-PrintQrCode -QrDataUrl $downloaderPageUrl_final -OutputFileName $qrPrintFileName -OutputFolder $printQrOutputFolder
    }

    $currentUserData = @{
        primaryEmail_val = $primaryEmail_val;
        givenName_val = $givenName_val;
        familyName_val = $familyName_val;
        title_val = $title_val;
        address_val = $address_val; 
        addressLabelForCard = $addressLabelForCard;
        phonesByType = $phonesByType; 
        downloaderPageFileName = $downloaderPageFileName; 
        downloaderPageUrl_final = $downloaderPageUrl_final;
        phoneData = $phoneData; # Passe TOUTES LES DONNÉES DE TÉLÉPHONE PRÉPARÉES
    }

    if ($GeneratePrintableCard) {
        $printableCardFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_')_print_card.html"
        Generate-PrintableBusinessCard -UserData $currentUserData -OutputFileName $printableCardFileName -OutputFolder $printableCardOutputFolder -TemplatePath $printableCardTemplatePath -PrintLogoUrl $printLogoUrl -WebsiteUrl $WebsiteUrl -WebsiteDisplayUrl $websiteDisplayUrl
    }
    # --- FIN : LOGIQUE DE GÉNÉRATION DES QR CODES ET CARTES IMPRIMABLES ---


    # --- DÉBUT : LOGIQUE POUR LA CARTE DE VISITE NUMÉRIQUE (ET SON QR CODE) VERS GITHUB ---
    $qrCodeImageUrl_raw_for_digital_card = "" 
    if ($AddDigitalCard -and $githubToken) {
        Write-Host "  - Démarrage de l'upload de la Carte de Visite Numérique vers GitHub pour $primaryEmail_val..." -ForegroundColor Cyan
        
        $qrGenerator = New-Object QRCoder.QRCodeGenerator; 
        $qrCodeData = $qrGenerator.CreateQrCode($downloaderPageUrl_final, [QRCoder.QRCodeGenerator+ECCLevel]::Q)
        $qrCode = New-Object QRCoder.PngByteQRCode($qrCodeData); 
        $qrCodeBytesSmall = $qrCode.GetGraphic(20, $qrCodeBlue, $qrCodeWhite) 

        $qrCodeWebFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_')_qrcode.png"
        $uploadResultQrCode = Publish-FileToGitHub -FileName $qrCodeWebFileName -FileContentBytes $qrCodeBytesSmall -FolderPathInRepo $qrcodeFolderPath
        if ($uploadResultQrCode) {
            $qrCodeImageUrl_raw_for_digital_card = $uploadResultQrCode.download_url
            Write-Host "    QR Code web URL (pour carte numérique) : $qrCodeImageUrl_raw_for_digital_card" -ForegroundColor Green
        } else {
            Write-Warning "Échec de l'upload du QR Code pour la carte numérique. Il pourrait ne pas s'afficher."
        }
        
        $vcfContent = "BEGIN:VCARD`nVERSION:3.0`nN:$($familyName_val);$($givenName_val);;;`nFN:$($givenName_val) $($familyName_val)`nORG:$orgName"
        if (-not [string]::IsNullOrEmpty($title_val)) { $vcfContent += "`nTITLE:$($title_val)" }
        
        # LOGIQUE VCF: Utilise les numéros bruts. Standard en dernier recours.
        if ($phoneData.HasWorkPhoneFromGam -and -not $phoneData.UsedDefaultPhoneForWork) { $vcfContent += "`nTEL;type=WORK,voice:$($phoneData.RawWorkPhone)" }
        if ($phoneData.HasMobilePhoneFromGam) { $vcfContent += "`nTEL;type=CELL,voice:$($phoneData.RawMobilePhone)" }
        if ($phoneData.UsedDefaultPhoneForWork -and -not $phoneData.HasMobilePhoneFromGam) { $vcfContent += "`nTEL;type=WORK,voice:$($phoneData.RawDefaultPhone)" } 
        
        $vcfContent += "`nEMAIL;type=INTERNET;type=WORK;type=pref:$($primaryEmail_val)"
        $vcfContent += "`nADR;type=WORK:;;$($address_val -replace "`r`n|`n", '\n');;;;"
        $vcfContent += "`nEND:VCARD"
        
        $vcfEncodedForUrl = [System.Net.WebUtility]::UrlEncode($vcfContent).Replace("+", "%20")
        $vcfDataUrl = "data:text/vcard;charset=utf-8,$vcfEncodedForUrl"
        $vcardDownloadName = "$($givenName_val)_$($familyName_val).vcf".Replace(" ", "_")
        
        $cardTemplateContent_digital = Get-Content -Path $digitalCardTemplatePath -Encoding UTF8 -Raw
        $cardTemplateContent_digital = $cardTemplateContent_digital.TrimStart([char]65279, [char]22)

        # Remplir le template HTML de la carte de visite interactive
        $downloaderPageContent = $cardTemplateContent_digital 
        $downloaderPageContent = $downloaderPageContent -replace '\{\{logo_url\}\}', $digitalCardLogoUrl
        $downloaderPageContent = $downloaderPageContent -replace '\{\{user_full_name\}\}', "$givenName_val $familyName_val"
        $downloaderPageContent = $downloaderPageContent -replace '\{\{user_title\}\}', $title_val
        $downloaderPageContent = $downloaderPageContent -replace '\{\{contact_list_html\}\}', $cardContactTextHtmlForDigitalCard 
        $downloaderPageContent = $downloaderPageContent -replace '\{\{action_buttons_html\}\}', $actionButtonsHtmlForDigitalCard
        $downloaderPageContent = $downloaderPageContent -replace '\{\{vcf_url\}\}', $vcfDataUrl
        $downloaderPageContent = $downloaderPageContent -replace '\{\{vcf_download_name\}\}', $vcardDownloadName
        $downloaderPageContent = $downloaderPageContent -replace '\{\{qrcode_image_url\}\}', $qrCodeImageUrl_raw_for_digital_card 
        $downloaderPageContent = $downloaderPageContent -replace '\{\{digital_card_page_url\}\}', $downloaderPageUrl_final
        $downloaderPageContent = $downloaderPageContent -replace '\{\{address_label\}\}', $addressLabelForCard
        $downloaderPageContent = $downloaderPageContent -replace '\{\{address_texte\}\}', $addressForDigitalCard
        $downloaderPageContent = $downloaderPageContent -replace '\{\{website_html_for_card\}\}', $websiteHtmlForDigitalCard
                                
        # Uploader le fichier HTML de la carte de visite
        $downloaderPageBytes = [System.Text.Encoding]::UTF8.GetBytes($downloaderPageContent)
        $uploadResultDownloader = Publish-FileToGitHub -FileName $downloaderPageFileName -FileContentBytes $downloaderPageBytes -FolderPathInRepo $vcardFolderPath

        if ($uploadResultDownloader) {
            Write-Host "    Digital Card page public URL: $downloaderPageUrl_final" -ForegroundColor Green
        } else {
            Write-Warning "Échec de l'upload de la page de la carte numérique."
        }
    }
    # --- FIN : LOGIQUE POUR LA CARTE DE VISITE NUMÉRIQUE (ET SON QR CODE) VERS GITHUB ---
    
    # --- DÉBUT : LOGIQUE POUR LE BLOC QR CODE DANS LA SIGNATURE MAIL ---
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
        Write-Warning "Le bloc QR Code pour la signature ne sera pas généré (pas de carte numérique activée ou URL QR code manquante)."
        $digital_card_html_block = ""
    }
    # --- FIN : LOGIQUE POUR LE BLOC QR CODE DANS LA SIGNATURE MAIL ---


    # --- DÉBUT : Préparation des données pour la SIGNATURE GMAIL ---
    $logPhoneLines = @(); 
    
    if ($DebugMode) { Write-Host "--- DÉBOGAGE FINAL SIGNATURE HTML ---" -ForegroundColor Yellow }

    # Préparation des lignes HTML des téléphones pour la signature
    $phoneBlockHtmlForSignatureFinal = ""
    # Ligne directe
    if(-not [string]::IsNullOrEmpty($phoneData.WorkPhoneHtmlForSignature)) { 
        $phoneBlockHtmlForSignatureFinal += $phoneData.WorkPhoneHtmlForSignature
    }
    # Mobile
    if(-not [string]::IsNullOrEmpty($phoneData.MobilePhoneHtmlForSignature)) { 
        $phoneBlockHtmlForSignatureFinal += $phoneData.MobilePhoneHtmlForSignature
    }
    # Standard (si AUCUNE ligne directe GAM ET AUCUN mobile GAM n'ont été trouvés)
    if (-not $phoneData.HasWorkPhoneFromGam -and -not $phoneData.HasMobilePhoneFromGam) {
        $phoneBlockHtmlForSignatureFinal += $phoneData.DefaultPhoneHtml
    }

    # Log les téléphones pour la console
    if ($DebugMode) { 
        Write-Host "Debug: phoneData.WorkPhoneHtmlForSignature: $($phoneData.WorkPhoneHtmlForSignature)" -ForegroundColor DarkCyan
        Write-Host "Debug: phoneData.MobilePhoneHtmlForSignature: $($phoneData.MobilePhoneHtmlForSignature)" -ForegroundColor DarkCyan
        Write-Host "Debug: phoneData.HasWorkPhoneFromGam: $($phoneData.HasWorkPhoneFromGam)" -ForegroundColor DarkCyan
        Write-Host "Debug: phoneData.HasMobilePhoneFromGam: $($phoneData.HasMobilePhoneFromGam)" -ForegroundColor DarkCyan
        Write-Host "Debug: phoneData.UsedDefaultPhoneForWork: $($phoneData.UsedDefaultPhoneForWork)" -ForegroundColor DarkCyan
        Write-Host "Debug: phoneData.WorkPhoneDisplayForTemplates (Debug): $($phoneData.WorkPhoneDisplayForTemplates)" -ForegroundColor DarkCyan
        Write-Host "Debug: phoneData.MobilePhoneDisplayForTemplates (Debug): $($phoneData.MobilePhoneDisplayForTemplates)" -ForegroundColor DarkCyan
        Write-Host "Debug: final phoneBlockHtmlForSignatureFinal content: $($phoneBlockHtmlForSignatureFinal)" -ForegroundColor DarkCyan
    }


    Write-Host "  - Prénom      : $givenName_val" -ForegroundColor Gray; Write-Host "  - Nom         : $familyName_val" -ForegroundColor Gray; Write-Host "  - Titre       : $(if ([string]::IsNullOrEmpty($title_val)) { '(aucun)' } else { $title_val })" -ForegroundColor Gray; Write-Host "  - Adresse     : $addressForSignature" -ForegroundColor Gray
    # Mise à jour du log console pour les téléphones
    if (-not [string]::IsNullOrEmpty($phoneData.WorkPhoneDisplayForTemplates)) { $logPhoneLines += "Ligne directe : $($phoneData.WorkPhoneDisplayForTemplates)" } 
    if (-not [string]::IsNullOrEmpty($phoneData.MobilePhoneDisplayForTemplates)) { $logPhoneLines += "Mobile : $($phoneData.MobilePhoneDisplayForTemplates)" }
    # Ajoute le standard au log si utilisé comme fallback et non déjà présent
    if ($phoneData.UsedDefaultPhoneForWork -and -not [string]::IsNullOrEmpty($phoneData.WorkPhoneDisplayForTemplates) -and -not ($logPhoneLines -like "*Téléphone :*")) {
        $logPhoneLines += "Téléphone : $($phoneData.FormattedDefaultPhoneDisplay)"
    }


    if ($logPhoneLines.Count -gt 0) { foreach($line in $logPhoneLines){ Write-Host "  - Téléphone   : $line" -ForegroundColor Gray } } else { Write-Host "  - Téléphone   : (aucun)" -ForegroundColor Gray }
    
    $functionLineConditional = ""; if ($title_val -ne "") { $functionLineConditional = "<span style=`"font-size: 10pt; color: #555555;`">" + $title_val.Trim() + "</span>" }
    
    $htmlTemplateContent = Get-Content -Path $signatureTemplatePath -Encoding UTF8 -Raw
    $charsToTrim = @([char]65279, [char]22); $htmlTemplateContent = $htmlTemplateContent.TrimStart($charsToTrim)
    
    $finalSignatureHtml = $htmlTemplateContent.Replace("{{digital_card_html_block}}", $digital_card_html_block)
    $finalSignatureHtml = $finalSignatureHtml -replace "{{givenName}}", $givenName_val
    $finalSignatureHtml = $finalSignatureHtml -replace "{{familyName}}", $familyName_val
    $finalSignatureHtml = $finalSignatureHtml -replace "{{functionLineConditional}}", $functionLineConditional
    $finalSignatureHtml = $finalSignatureHtml -replace "{{primaryEmail}}", $primaryEmail_val
    $finalSignatureHtml = $finalSignatureHtml -replace "{{phoneBlock}}", $phoneBlockHtmlForSignatureFinal 
    $finalSignatureHtml = $finalSignatureHtml -replace "{{address_texte}}", $addressForSignature
    $finalSignatureHtml = $finalSignatureHtml -replace "{{address_url_maps}}", $address_url_maps
    $finalSignatureHtml = $finalSignatureHtml -replace "{{logo_url}}", $signatureLogoUrl
    $finalSignatureHtml = $finalSignatureHtml -replace "{{website_url}}", $WebsiteUrl
    $finalSignatureHtml = $finalSignatureHtml -replace "{{website_display_url}}", $websiteDisplayUrl
    $finalSignatureHtml = $finalSignatureHtml -replace "{{org_name}}", $orgName

    if ($DebugMode) { Write-Host "Debug: finalSignatureHtml ($($finalSignatureHtml.Length) chars):`n$finalSignatureHtml" -ForegroundColor DarkYellow }


    $tempSignaturePath = Join-Path -Path $projectRoot -ChildPath "temp_sig_$($primaryEmail_val.Replace('@','_')).html"
    
    Write-Host "  - Vérification de la signature actuelle sur Google..." -ForegroundColor DarkGray
    $currentSignatureHtml = & $gamPath user "$primaryEmail_val" print signature | Out-String
    $newSigNormalized = $finalSignatureHtml -replace '\s' -replace ' ', ' ' 
    $currentSigNormalized = $currentSignatureHtml -replace '\s' -replace ' ', ' ' 

    if ($newSigNormalized -eq $currentSigNormalized) {
        Write-Host "  - La signature est déjà à jour. Mise à jour ignorée." -ForegroundColor Green
    } else {
        Write-Host "  - Signature mise à jour détectée. Application en cours..." -ForegroundColor DarkCyan
        $encoding = New-Object System.Text.UTF8Encoding($false); [System.IO.File]::WriteAllText($tempSignaturePath, $finalSignatureHtml, $encoding)
        
        Write-Host "Application de la signature pour $primaryEmail_val..." -ForegroundColor DarkCyan
        & $gamPath user "$primaryEmail_val" signature file "$tempSignaturePath" html
        Remove-Item -Path $tempSignaturePath -ErrorAction SilentlyContinue
    }
}

[Console]::OutputEncoding = $originalEncoding
Write-Host "Processus d'application des signatures terminé." -ForegroundColor Green
﻿# set_user_signatures.ps1 (v45.23 - Carte de Visite Numérique mise en forme compact et colorisation des titres)
#
param(
    [string]$SingleUserEmail = "",
    [switch]$IncludeSuspended,
    [switch]$AddDigitalCard,
    [switch]$ShowHelp
)

if ($ShowHelp) {
    $helpText = @"
NOM:
    set_user_signatures.ps1

SYNOPSIS:
    Met à jour les signatures email et peut générer une carte de visite numérique complète (vCard + QR Code interactif).

SYNTAXE:
    .\set_user_signatures.ps1 [-SingleUserEmail <string>] [-IncludeSuspended] [-AddDigitalCard] [-ShowHelp]

DESCRIPTION:
    Ce script automatise la mise à jour des signatures Gmail via GAM. Il est optimisé pour ne pas effectuer de mises à jour inutiles.

    - Mode standard : Met à jour la signature principale de l'utilisateur.
    
    - Mode Carte de Visite (-AddDigitalCard) : Génère une page web professionnelle (hébergée sur GitHub Pages) qui contient un lien de
      téléchargement direct pour la vCard de l'utilisateur. Pour assurer la compatibilité, la vCard est encodée
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
    
    -ShowHelp
        Affiche ce message d'aide et quitte le script.

EXEMPLES:
    # Affiche cette aide complète
    .\set_user_signatures.ps1 -ShowHelp
    
    # Met à jour la signature ET la carte de visite numérique pour un utilisateur spécifique
    # Le site web est configuré directement dans le script.
    .\set_user_signatures.ps1 -SingleUserEmail "s.gille@cjml.fr" -AddDigitalCard

    # Met à jour les signatures et cartes de visite pour tous les utilisateurs actifs
    # Le site web est configuré directement dans le script.
    .\set_user_signatures.ps1 -AddDigitalCard

    # Met à jour les signatures et cartes de visite pour tous les utilisateurs (actifs et suspendus)
    # Le site web est configuré directement dans le script.
    .\set_user_signatures.ps1 -AddDigitalCard -IncludeSuspended
"@
    Write-Host $helpText
    return
}

# --- Configuration ---
# MODIFIÉ : Définition d'un dossier racine pour le projet pour plus de portabilité
$projectRoot = "C:\GAMWork\Signatures"

$gamPath = "C:\GAM7\gam.exe"
$signatureTemplatePath = Join-Path -Path $projectRoot -ChildPath "signature_template.html"
$digitalCardTemplatePath = Join-Path -Path $projectRoot -ChildPath "digital_card_template.html"
# Logo pour la signature d'email (le carré)
$signatureLogoUrl = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/Logo-CJML.png"
# NOUVEAU : URL du logo rectangulaire pour la carte de visite
$digitalCardLogoUrl = "https://raw.githubusercontent.com/Centre-Jean-Marie-LARRIEU/assets-cjml/main/logo-horizontal.jpg"
$orgName = "Centre Jean-Marie LARRIEU" # Nom de l'organisation pour la signature

$defaultPhoneNumber = "05 62 91 32 50"
$defaultAddress = @"
414 Rue du Layris
65710 CAMPAN
"@

# NOUVEAU: Définition du site web en dur dans la configuration
$WebsiteUrl = "http://www.cjml.fr" # L'URL complète du site web
$websiteDisplayUrl = "" # Sera généré ci-dessous
if (-not [string]::IsNullOrEmpty($WebsiteUrl)) {
    # Enlève "http://", "https://" et "www." pour l'affichage
    $websiteDisplayUrl = $WebsiteUrl -replace "^https?:\/\/(www\.)?", ""
    # Si le domaine commence par cjml.fr, on ajoute www. pour l'affichage comme demandé
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

chcp.com 65001 | Out-Null
$originalEncoding = [Console]::OutputEncoding
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

# --- Récupération des utilisateurs ---
$usersToProcess = @()
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

if ($usersToProcess.Count -eq 0) { Write-Host "Aucun utilisateur trouvé à traiter. Quitte le script."; exit 0 }
Write-Host "Trouvé $($usersToProcess.Count) utilisateur(s) à traiter." -ForegroundColor Cyan

foreach ($user in $usersToProcess) {
    if ($user -eq $null) { Write-Error "Ignorer l'objet utilisateur nul."; continue }
    
    $primaryEmail_val = if ($user.primaryEmail) { $user.primaryEmail } else { "" }
    $givenName_val = if ($user."name.givenName") { $user."name.givenName" } else { "" }
    $familyName_val = if ($user."name.familyName") { $user."name.familyName" } else { "" }
    $title_val = if ($user."organizations.0.title") { $user."organizations.0.title" } else { "" }
    
    # Nouvelles variables pour l'adresse et le label
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
    
    # Formate l'adresse pour la signature et la carte numérique
    $addressForSignature = ($address_val -replace "`r`n", " - " -replace "`n", " - ").Trim()
    $addressForDigitalCard = ($address_val -replace "`r`n", "<br>").Trim()
    $address_url_maps = "https://www.google.com/maps/search/?api=1&query=" + [System.Net.WebUtility]::UrlEncode($addressForSignature)

    # NOUVEAU: Label pour l'adresse sur la carte de visite numérique
    # Utiliser un string comparateur pour une comparaison plus fiable des adresses
    if ([string]::Compare($userRawAddress.Trim(), ($defaultAddress -replace "`r`n", "`n").Trim(), $true) -eq 0 -or [string]::IsNullOrEmpty($userRawAddress)) {
        $addressLabelForCard = "Siège Social"
    } else {
        $addressLabelForCard = "Adresse du Bureau"
    }

    Write-Host "--- Traitement de l'utilisateur : $primaryEmail_val (Suspended: $($user.suspended)) ---"
    
    $phonesByType = @{ work = @(); mobile = @() }
    for ($i = 0; $i -lt 5; $i++) {
        $phoneValueProperty = "phones.$i.value"; $phoneTypeProperty = "phones.$i.type"
        if ($user.PSObject.Properties.Name -contains $phoneValueProperty -and -not [string]::IsNullOrEmpty($user.$phoneValueProperty)) {
            $phoneValue = $user.$phoneValueProperty; $formattedDisplayValue = $phoneValue
            if ($phoneValue -match '^\+33[1-9]\d{8}$') { $localNumber = $phoneValue -replace '^\+33', '0'; $formattedDisplayValue = $localNumber -replace '^(0\d)(\d{2})(\d{2})(\d{2})(\d{2})$', '$1 $2 $3 $4 $5' }
            $phoneType = if ($user.PSObject.Properties.Name -contains $phoneTypeProperty) { $user.$phoneTypeProperty.ToLower() } else { "work" }
            if ($phonesByType.ContainsKey($phoneType)) { $phonesByType[$phoneType] += @{ Display = $formattedDisplayValue; Raw = $phoneValue } }
        }
    }
    
    $digital_card_html_block = "" # Réinitialise pour chaque utilisateur
    if ($AddDigitalCard -and $githubToken) {
        Write-Host "  - Génération de la Carte de Visite Numérique pour $primaryEmail_val..." -ForegroundColor Cyan
        
        $vcfContent = "BEGIN:VCARD`nVERSION:3.0`nN:$($familyName_val);$($givenName_val);;;`nFN:$($givenName_val) $($familyName_val)`nORG:$orgName"
        if (-not [string]::IsNullOrEmpty($title_val)) { $vcfContent += "`nTITLE:$($title_val)" }
        $vcfContent += "`nEMAIL;type=INTERNET;type=WORK;type=pref:$($primaryEmail_val)"
        if ($phonesByType['work'].Count -gt 0) { foreach ($phone in $phonesByType['work']) { $vcfContent += "`nTEL;type=WORK,voice:$($phone.Raw)" } }
        if ($phonesByType['mobile'].Count -gt 0) { foreach ($phone in $phonesByType['mobile']) { $vcfContent += "`nTEL;type=CELL,voice:$($phone.Raw)" } }
        if ($phonesByType['work'].Count -eq 0 -and $phonesByType['mobile'].Count -eq 0) { $vcfContent += "`nTEL;type=WORK,voice:$($defaultPhoneNumber -replace '[^0-9+]')" }
        $vcfContent += "`nADR;type=WORK:;;$($address_val -replace "`r`n|`n", '\n');;;;"
        $vcfContent += "`nEND:VCARD"
        
        $vcfEncodedForUrl = [System.Net.WebUtility]::UrlEncode($vcfContent).Replace("+", "%20")
        $vcfDataUrl = "data:text/vcard;charset=utf-8,$vcfEncodedForUrl"
        $vcardDownloadName = "$($givenName_val)_$($familyName_val).vcf".Replace(" ", "_")
        
        # Lecture du template HTML pour la carte de visite
        $cardTemplateContent = Get-Content -Path $digitalCardTemplatePath -Encoding UTF8 -Raw
        $cardTemplateContent = $cardTemplateContent.TrimStart([char]65279, [char]22)

        # --- ORDRE DES OPÉRATIONS POUR LES URLS DE LA CARTE ET DU QR CODE ---

        # 1. Définir le nom du fichier HTML de la carte de visite
        $downloaderPageFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_').html"
        # Définir l'URL finale de la page AVANT de la générer/uploader
        $downloaderPageUrl_final = "$githubPagesBaseUrl/$vcardFolderPath/$downloaderPageFileName"

        # 2. Générer l'URL du QR code à partir de l'URL de la page
        $qrGenerator = New-Object QRCoder.QRCodeGenerator; $qrCodeData = $qrGenerator.CreateQrCode($downloaderPageUrl_final, [QRCoder.QRCodeGenerator+ECCLevel]::Q)
        $qrCode = New-Object QRCoder.PngByteQRCode($qrCodeData); $qrCodeBytes = $qrCode.GetGraphic(20, $qrCodeBlue, $qrCodeWhite)
        $qrCodeFileName = "$($primaryEmail_val -replace '[^a-zA-Z0-9]','_')_qrcode.png"

        # 3. Uploader l'image du QR code et obtenir son URL brute de téléchargement
        $uploadResultQrCode = Publish-FileToGitHub -FileName $qrCodeFileName -FileContentBytes $qrCodeBytes -FolderPathInRepo $qrcodeFolderPath
        if ($uploadResultQrCode) {
            $qrCodeImageUrl_raw = $uploadResultQrCode.download_url
            Write-Host "    QR Code raw URL: $qrCodeImageUrl_raw" -ForegroundColor Green
        } else {
            Write-Warning "Échec de l'upload de l'image QR Code. La carte numérique pourrait ne pas s'afficher correctement."
            $qrCodeImageUrl_raw = "" # S'assurer que la variable est vide en cas d'échec
        }

        # 4. Préparer le HTML pour les informations de contact (téléphone et email UNIQUEMENT)
        $cardContactTextHtml = ""
        if ($phonesByType['work'].Count -gt 0) { $workNumbers = ($phonesByType['work'] | ForEach-Object { "<a href=`"tel:$($_.Raw -replace '[^0-9+]')`">$($_.Display)</a>" }) -join ', '; $cardContactTextHtml += "<div class=`"contact-item`"><span class=`"label`">Ligne directe</span>$workNumbers</div>" }
        else { $defaultPhoneFormatted = $defaultPhoneNumber -replace '^(0\d)(\d{2})(\d{2})(\d{2})(\d{2})$', '$1 $2 $3 $4 $5'; $cardContactTextHtml += "<div class=`"contact-item`"><span class=`"label`">Téléphone</span><a href=`"tel:$($defaultPhoneNumber -replace '[^0-9+]')`">$defaultPhoneFormatted</a></div>" }
        if ($phonesByType['mobile'].Count -gt 0) { $mobileNumbers = ($phonesByType['mobile'] | ForEach-Object { "<a href=`"tel:$($_.Raw -replace '[^0-9+]')`">$($_.Display)</a>" }) -join ', '; $cardContactTextHtml += "<div class=`"contact-item`"><span class=`"label`">Mobile</span>$mobileNumbers</div>" }
        $cardContactTextHtml += "<div class=`"contact-item`"><span class=`"label`">Email</span><a href=`"mailto:$primaryEmail_val`">$primaryEmail_val</a></div>"
        # REMARQUE : La ligne pour l'adresse a été retirée d'ici car l'adresse est gérée par son propre bloc dans le template HTML.
        
        $actionButtonsHtml = ""
        if ($phonesByType['mobile'].Count -gt 0) { $mobilePhone = $phonesByType['mobile'][0]; $mobileTelLink = "tel:$($mobilePhone.Raw -replace '[^0-9+]')"; $actionButtonsHtml += "<a href=`"$mobileTelLink`" class=`"button secondary`">Appeler (Mobile)</a>" }
        if ($phonesByType['work'].Count -gt 0) { $workPhone = $phonesByType['work'][0]; $workTelLink = "tel:$($workPhone.Raw -replace '[^0-9+]')"; $actionButtonsHtml += "<a href=`"$workTelLink`" class=`"button secondary`">Appeler (Direct)</a>" }
        if ($phonesByType['work'].Count -eq 0 -and $phonesByType['mobile'].Count -eq 0) { $actionButtonsHtml += "<a href=`"tel:$($defaultPhoneNumber -replace '[^0-9+]')`" class=`"button secondary`">Appeler le Centre</a>" }
        $actionButtonsHtml += "<a href=`"mailto:$primaryEmail_val`" class=`"button secondary`">Envoyer un Email</a>"
        $actionButtonsHtml += "<a href=`"$address_url_maps`" target=`"_blank`" class=`"button secondary`">Itinéraire</a>"

        # NOUVEAU: Bloc HTML conditionnel pour le site web sur la carte de visite
        $websiteHtmlForCard = ""
        if (-not [string]::IsNullOrEmpty($WebsiteUrl)) {
            $websiteHtmlForCard = @"
<div class="contact-item">
    <span class="label">Site Web</span>
    <a href="$WebsiteUrl" target="_blank" rel="noopener noreferrer" style="color: var(--primary-blue); text-decoration: underline;">$websiteDisplayUrl</a>
</div>
"@
        }

        # 5. Maintenant que toutes les URLs et blocs HTML conditionnels sont définis, Remplir le template HTML de la carte de visite
        # Utilisation de réaffectations ligne par ligne pour plus de robusteté
        $downloaderPageContent = $cardTemplateContent 
        $downloaderPageContent = $downloaderPageContent -replace '\{\{logo_url\}\}', $digitalCardLogoUrl
        $downloaderPageContent = $downloaderPageContent -replace '\{\{user_full_name\}\}', "$givenName_val $familyName_val"
        $downloaderPageContent = $downloaderPageContent -replace '\{\{user_title\}\}', $title_val
        $downloaderPageContent = $downloaderPageContent -replace '\{\{contact_list_html\}\}', $cardContactTextHtml
        $downloaderPageContent = $downloaderPageContent -replace '\{\{action_buttons_html\}\}', $actionButtonsHtml
        $downloaderPageContent = $downloaderPageContent -replace '\{\{vcf_url\}\}', $vcfDataUrl
        $downloaderPageContent = $downloaderPageContent -replace '\{\{vcf_download_name\}\}', $vcardDownloadName
        $downloaderPageContent = $downloaderPageContent -replace '\{\{qrcode_image_url\}\}', $qrCodeImageUrl_raw
        $downloaderPageContent = $downloaderPageContent -replace '\{\{digital_card_page_url\}\}', $downloaderPageUrl_final
        $downloaderPageContent = $downloaderPageContent -replace '\{\{address_label\}\}', $addressLabelForCard
        $downloaderPageContent = $downloaderPageContent -replace '\{\{address_texte\}\}', $addressForDigitalCard
        $downloaderPageContent = $downloaderPageContent -replace '\{\{website_html_for_card\}\}', $websiteHtmlForCard # NOUVEAU REMPLACEMENT
                                

        # 6. Uploader le fichier HTML de la carte de visite
        $downloaderPageBytes = [System.Text.Encoding]::UTF8.GetBytes($downloaderPageContent)
        $uploadResultDownloader = Publish-FileToGitHub -FileName $downloaderPageFileName -FileContentBytes $downloaderPageBytes -FolderPathInRepo $vcardFolderPath

        if ($uploadResultDownloader) {
            Write-Host "    Digital Card page public URL: $downloaderPageUrl_final" -ForegroundColor Green
            # MODIFIÉ : Bloc HTML pour la signature email, avec le QR code aligné à droite
            $digital_card_html_block = @"
<table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="padding-top:10px;"><tr>
<td style="width:100%; text-align: right; vertical-align: middle;"> <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="display: inline-block;"> <tr>
            <td style="padding-right:15px; vertical-align:middle; text-align: right;"> <p style="margin:0;padding:0;font-size:9pt;color:#555555;line-height:1.3">Scannez-moi ou <a href="$downloaderPageUrl_final" target="_blank" style="color:#068FD0;text-decoration:underline">cliquez ici</a><br>pour ma carte de visite numérique.</p>
            </td>
            <td style="width:96px; vertical-align:middle;"> <a href="$downloaderPageUrl_final" target="_blank" style="text-decoration:none;"><img src="$qrCodeImageUrl_raw" width="80" style="width:80px;height:80px;display:block;border:0;" alt="QR Code"/></a>
            </td>
        </tr>
    </table>
</td>
</tr></table>
"@
        } else {
            Write-Warning "Échec de l'upload de la page de la carte numérique. Le bloc QR Code ne sera pas ajouté à la signature."
            $digital_card_html_block = ""
        }
    }
    
    $logPhoneLines = @(); $phoneBlockHtml = ""; $linkStyle = "color: #555555; text-decoration: underline;"
    if ($phonesByType['work'].Count -gt 0) {
        $phoneBlockHtml += "Ligne directe : "; $phoneLinks = @(); foreach ($phone in $phonesByType['work']) { $rawNumberForTel = ($phone.Raw -replace '[^0-9+]'); $logPhoneLines += "Ligne directe : " + $phone.Display; $phoneLinks += "<a href=`"tel:$rawNumberForTel`" style=`"$linkStyle`">$($phone.Display)</a>" }; $phoneBlockHtml += $phoneLinks -join ', '; $phoneBlockHtml += "<br>"
    } else {
        $defaultPhoneFormatted = $defaultPhoneNumber -replace '^(0\d)(\d{2})(\d{2})(\d{2})(\d{2})$', '$1 $2 $3 $4 $5'
        $defaultTelLink = "tel:$($defaultPhoneNumber -replace '[^0-9+]')"; $logPhoneLines += "Téléphone : $defaultPhoneFormatted"; $phoneBlockHtml += "Téléphone : <a href=`"$defaultTelLink`" style=`"$linkStyle`">$defaultPhoneFormatted</a><br>"
    }
    if ($phonesByType['mobile'].Count -gt 0) {
        $phoneBlockHtml += "Mobile : "; $phoneLinks = @(); foreach ($phone in $phonesByType['mobile']) { $rawNumberForTel = ($phone.Raw -replace '[^0-9+]'); $logPhoneLines += "Mobile : " + $phone.Display; $phoneLinks += "<a href=`"tel:$rawNumberForTel`" style=`"$linkStyle`">$($phone.Display)</a>" }; $phoneBlockHtml += $phoneLinks -join ', '; $phoneBlockHtml += "<br>"
    }

    Write-Host "  - Prénom      : $givenName_val" -ForegroundColor Gray; Write-Host "  - Nom         : $familyName_val" -ForegroundColor Gray; Write-Host "  - Titre       : $(if ([string]::IsNullOrEmpty($title_val)) { '(aucun)' } else { $title_val })" -ForegroundColor Gray; Write-Host "  - Adresse     : $addressForSignature" -ForegroundColor Gray
    if ($logPhoneLines.Count -gt 0) { foreach($line in $logPhoneLines){ Write-Host "  - Téléphone   : $line" -ForegroundColor Gray } } else { Write-Host "  - Téléphone   : (aucun)" -ForegroundColor Gray }
    
    $functionLineConditional = ""; if ($title_val -ne "") { $functionLineConditional = "<span style=`"font-size: 10pt; color: #555555;`">" + $title_val.Trim() + "</span>" }
    
    $htmlTemplateContent = Get-Content -Path $signatureTemplatePath -Encoding UTF8 -Raw
    $charsToTrim = @([char]65279, [char]22); $htmlTemplateContent = $htmlTemplateContent.TrimStart($charsToTrim)
    
    # MODIFIÉ : Construction de $finalSignatureHtml avec des réaffectations ligne par ligne pour plus de robustesse
    $finalSignatureHtml = $htmlTemplateContent.Replace("{{digital_card_html_block}}", $digital_card_html_block)
    $finalSignatureHtml = $finalSignatureHtml -replace "{{givenName}}", $givenName_val
    $finalSignatureHtml = $finalSignatureHtml -replace "{{familyName}}", $familyName_val
    $finalSignatureHtml = $finalSignatureHtml -replace "{{functionLineConditional}}", $functionLineConditional
    $finalSignatureHtml = $finalSignatureHtml -replace "{{primaryEmail}}", $primaryEmail_val
    $finalSignatureHtml = $finalSignatureHtml -replace "{{phoneBlock}}", $phoneBlockHtml
    $finalSignatureHtml = $finalSignatureHtml -replace "{{address_texte}}", $addressForSignature
    $finalSignatureHtml = $finalSignatureHtml -replace "{{address_url_maps}}", $address_url_maps
    $finalSignatureHtml = $finalSignatureHtml -replace "{{logo_url}}", $signatureLogoUrl
    $finalSignatureHtml = $finalSignatureHtml -replace "{{website_url}}", $WebsiteUrl
    $finalSignatureHtml = $finalSignatureHtml -replace "{{website_display_url}}", $websiteDisplayUrl
    $finalSignatureHtml = $finalSignatureHtml -replace "{{org_name}}", $orgName


    # La seule chose qui change est que le fichier temporaire sera aussi créé dans le dossier du projet
    $tempSignaturePath = Join-Path -Path $projectRoot -ChildPath "temp_sig_$($primaryEmail_val.Replace('@','_')).html"
    
    Write-Host "  - Vérification de la signature actuelle sur Google..." -ForegroundColor DarkGray
    $currentSignatureHtml = & $gamPath user $primaryEmail_val print signature | Out-String
    $newSigNormalized = $finalSignatureHtml -replace '\s'
    $currentSigNormalized = $currentSignatureHtml -replace '\s'

    if ($newSigNormalized -eq $currentSigNormalized) {
        Write-Host "  - La signature est déjà à jour. Mise à jour ignorée." -ForegroundColor Green
        continue
    }

    $encoding = New-Object System.Text.UTF8Encoding($false); [System.IO.File]::WriteAllText($tempSignaturePath, $finalSignatureHtml, $encoding)
    
    Write-Host "Application de la signature pour $primaryEmail_val..." -ForegroundColor DarkCyan
    & $gamPath user "$primaryEmail_val" signature file "$tempSignaturePath" html
    Remove-Item -Path $tempSignaturePath -ErrorAction SilentlyContinue
}

[Console]::OutputEncoding = $originalEncoding
Write-Host "Processus d'application des signatures terminé." -ForegroundColor Green
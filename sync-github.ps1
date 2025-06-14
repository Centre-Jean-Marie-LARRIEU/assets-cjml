# sync-github.ps1 (Version 2.2 - Améliorations de robustesse et gestion de branche)

# Définit l'encodage de la console en UTF-8 pour un affichage correct des caractères
chcp.com 65001 | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Fonction pour vérifier et configurer l'utilisateur Git
function Set-GitConfig {
    # Vérifie si user.name est configuré
    $userName = git config user.name
    if ([string]::IsNullOrWhiteSpace($userName)) {
        Write-Host "ATTENTION : Votre nom d'utilisateur Git n'est pas configuré." -ForegroundColor Yellow
        Write-Host "Cela peut causer des avertissements de Git." -ForegroundColor Yellow
        $inputName = Read-Host -Prompt "Entrez votre nom pour Git (ex: 'John Doe')"
        if (![string]::IsNullOrWhiteSpace($inputName)) {
            git config --global user.name "$inputName"
            Write-Host "Votre nom Git a été configuré globalement." -ForegroundColor Green
        }
    }

    # Vérifie si user.email est configuré
    $userEmail = git config user.email
    if ([string]::IsNullOrWhiteSpace($userEmail)) {
        Write-Host "ATTENTION : Votre adresse email Git n'est pas configurée." -ForegroundColor Yellow
        Write-Host "Cela peut causer des avertissements de Git." -ForegroundColor Yellow
        $inputEmail = Read-Host -Prompt "Entrez votre email pour Git (ex: 'john.doe@example.com')"
        if (![string]::IsNullOrWhiteSpace($inputEmail)) {
            git config --global user.email "$inputEmail"
            Write-Host "Votre email Git a été configuré globalement." -ForegroundColor Green
        }
    }
}

# Appelle la fonction de configuration Git au démarrage
Set-GitConfig

# Boucle principale du menu
while ($true) {
    Clear-Host
    Write-Host "----------------------------------------------------------------"
    Write-Host "  Outil de Gestion du Dépôt GitHub pour les Signatures" -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------"
    Write-Host "Que souhaitez-vous faire ?"
    Write-Host
    Write-Host "  [1] Synchroniser (envoyer) les modifications locales vers GitHub"
    Write-Host "  [2] Restaurer un fichier à une version précédente"
    Write-Host "  [Q] Quitter"
    Write-Host
    
    $choice = Read-Host -Prompt "Votre choix"
    $choice = $choice.Trim().ToLower() # Rend la saisie insensible à la casse et retire les espaces

    switch ($choice) {
        "1" {
            Clear-Host
            $commitMessage = Read-Host -Prompt "Veuillez entrer un message pour décrire vos modifications (ex: 'Mise à jour du logo')"
            if ([string]::IsNullOrWhiteSpace($commitMessage)) {
                Write-Host "Le message de commit ne peut pas être vide. Opération annulée." -ForegroundColor Red
                Read-Host -Prompt "Appuyez sur Entrée pour continuer..." # Ajout pour ne pas fermer directement
                break # Retourne au menu
            }

            Write-Host "`n--- Préparation de la synchronisation vers GitHub ---" -ForegroundColor Cyan
            
            # 1. Assurer que la branche locale est 'main' si elle était 'master'
            Write-Host "1. Vérification et ajustement de la branche locale..." -ForegroundColor Green
            try {
                $currentBranch = git rev-parse --abbrev-ref HEAD
                if ($currentBranch -eq "master") {
                    Write-Host "  Renommage de la branche 'master' en 'main'..." -ForegroundColor Yellow
                    git branch -M main
                    if ($LASTEXITCODE -ne 0) { throw "Échec du renommage de la branche." }
                }
            } catch {
                Write-Host "Une erreur est survenue lors de la vérification/renommage de la branche : $($_.Exception.Message)" -ForegroundColor Red
                Read-Host -Prompt "Appuyez sur Entrée pour continuer..."
                break
            }

            # 2. Ajout de tous les fichiers modifiés
            Write-Host "2. Ajout de tous les fichiers modifiés..." -ForegroundColor Green
            git add .
            if ($LASTEXITCODE -ne 0) { Write-Host "Erreur lors de l'ajout des fichiers. Opération annulée." -ForegroundColor Red; Read-Host -Prompt "Appuyez sur Entrée pour continuer..."; break }

            # 3. Création du 'commit'
            Write-Host "3. Création du 'commit' avec le message : '$commitMessage'..." -ForegroundColor Green
            git commit -m "$commitMessage"
            if ($LASTEXITCODE -ne 0) {
                Write-Host "Erreur lors de la création du commit. Assurez-vous qu'il y a des changements à committer." -ForegroundColor Red
                Read-Host -Prompt "Appuyez sur Entrée pour continuer..."
                break
            }

            # 4. Envoi des modifications vers GitHub
            Write-Host "4. Envoi des modifications vers GitHub (branche 'main')..." -ForegroundColor Green
            git push -u origin main
            if ($LASTEXITCODE -eq 0) {
                Write-Host "`n--- Synchronisation terminée avec succès ! ---" -ForegroundColor Green
            } else {
                Write-Host "`n--- ERREUR : La synchronisation vers GitHub a échoué ! ---" -ForegroundColor Red
                Write-Host "Vérifiez votre connexion internet, vos identifiants GitHub ou les permissions du dépôt." -ForegroundColor Red
            }
            break
        }
        "2" {
            Clear-Host
            Write-Host "----------------------------------------------------------------"
            Write-Host "Historique des 15 dernières versions enregistrées :" -ForegroundColor Yellow
            Write-Host "----------------------------------------------------------------"
            git log --oneline -n 15
            Write-Host "----------------------------------------------------------------`n"
            
            $commitHash = Read-Host -Prompt "Veuillez copier-coller l'identifiant de la version à restaurer (les 7 premiers caractères suffisent)"
            if ([string]::IsNullOrWhiteSpace($commitHash)) {
                Write-Host "Aucun identifiant fourni. Opération annulée." -ForegroundColor Red; break
            }
            
            $fileName = Read-Host -Prompt "Entrez le nom exact du fichier à restaurer (ex: set_user_signatures.ps1)"
            if ([string]::IsNullOrWhiteSpace($fileName)) {
                Write-Host "Nom de fichier vide. Opération annulée." -ForegroundColor Red; break
            }
            
            # Vérifie si le fichier existe localement avant de tenter la restauration
            if (-not (Test-Path $fileName -PathType Leaf)) {
                Write-Host "Le fichier '$fileName' n'existe pas dans le répertoire actuel. Impossible de le restaurer." -ForegroundColor Red
                Write-Host "Assurez-vous d'être dans le bon répertoire ou que le nom est correct." -ForegroundColor Yellow
                break
            }

            Write-Host "`nRestauration du fichier '$fileName' à la version '$commitHash'..." -ForegroundColor Cyan
            try {
                git checkout $commitHash -- $fileName
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "`n--- Opération terminée avec succès ! ---" -ForegroundColor Green
                    Write-Host "Le fichier '$fileName' a été restauré à la version '$commitHash'." -ForegroundColor White
                    Write-Host "Il est maintenant dans un état 'staged' (prêt à être commit). " -ForegroundColor Yellow
                    Write-Host "ASTUCE : Pour **finaliser** cette restauration et l'enregistrer dans l'historique," -ForegroundColor Yellow
                    Write-Host "utilisez l'option [1] du menu pour 'Synchroniser' vos modifications." -ForegroundColor Yellow
                } else {
                    Write-Host "La restauration a échoué pour une raison inconnue. Vérifiez l'identifiant et le nom du fichier." -ForegroundColor Red
                }
            } catch {
                Write-Error "Une erreur PowerShell inattendue s'est produite lors de la restauration : $($_.Exception.Message)"
            }
            break
        }
        "q" {
            Write-Host "Au revoir !"
            return
        }
        default {
            Write-Host "Choix invalide. Veuillez réessayer." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    
    # Pause après les opérations pour que l'utilisateur puisse lire le message
    if ($choice -ne "q") { # Ne pas demander si on quitte
        Write-Host ""
        Read-Host -Prompt "Appuyez sur Entrée pour revenir au menu principal"
    }
}
# sync-github.ps1 (Version 2.1 - Correction de l'encodage)

# Définit l'encodage de la console en UTF-8 pour un affichage correct des caractères
chcp.com 65001 | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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

    switch ($choice) {
        "1" {
            Clear-Host
            $commitMessage = Read-Host -Prompt "Veuillez entrer un message pour décrire vos modifications (ex: 'Mise à jour du logo')"
            if ([string]::IsNullOrWhiteSpace($commitMessage)) {
                Write-Host "Le message de commit ne peut pas être vide. Opération annulée." -ForegroundColor Red; break
            }
            Write-Host "`n--- Préparation de la synchronisation vers GitHub ---" -ForegroundColor Cyan
            Write-Host "1. Ajout de tous les fichiers modifiés..." -ForegroundColor Green
            git add .
            Write-Host "2. Création du 'commit' avec le message : '$commitMessage'..." -ForegroundColor Green
            git commit -m "$commitMessage"
            Write-Host "3. Envoi des modifications vers GitHub..." -ForegroundColor Green
            git push -u origin main # On spécifie la branche 'main' pour plus de clarté
            Write-Host "`n--- Synchronisation terminée ! ---" -ForegroundColor Cyan
            break
        }
        "2" {
            Clear-Host
            Write-Host "----------------------------------------------------------------"
            Write-Host "Historique des 15 dernières versions enregistrées :" -ForegroundColor Yellow
            Write-Host "----------------------------------------------------------------"
            git log --oneline -n 15
            Write-Host "----------------------------------------------------------------`n"
            $commitHash = Read-Host -Prompt "Veuillez copier-coller l'identifiant de la version à restaurer"
            if ([string]::IsNullOrWhiteSpace($commitHash)) {
                Write-Host "Aucun identifiant fourni. Opération annulée." -ForegroundColor Red; break
            }
            $fileName = Read-Host -Prompt "Entrez le nom exact du fichier à restaurer (ex: set_user_signatures.ps1)"
            if ([string]::IsNullOrWhiteSpace($fileName) -or -not (Test-Path $fileName)) {
                Write-Host "Nom de fichier '$fileName' invalide ou fichier non trouvé. Opération annulée." -ForegroundColor Red; break
            }
            Write-Host "`nRestauration du fichier '$fileName' à la version '$commitHash'..." -ForegroundColor Cyan
            try {
                git checkout $commitHash -- $fileName
                Write-Host "`n--- Opération terminée avec succès ! ---" -ForegroundColor Green
                Write-Host "ASTUCE : Si cette restauration vous convient, utilisez l'option [1] du menu pour la sauvegarder." -ForegroundColor Yellow
            } catch {
                Write-Error "Une erreur s'est produite lors de la restauration."
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
    
    if ($choice -in @("1", "2")) {
        Write-Host ""
        Read-Host -Prompt "Appuyez sur Entrée pour fermer cette fenêtre"
        return
    }
}
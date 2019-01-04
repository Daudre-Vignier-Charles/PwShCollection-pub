INSTALLATION :

- Lancer INSTALL.BAT en tant qu'administrateur.

DESINSTALLATION :

- Lancer UNINSTALL.BAT en tant qu'administrateur.

UTILISATION :

- Télécharger sur GLPI la liste des ordinateurs au format CSV.
- Déplacer le fichier CSV quelque part dans C: (la liste doit être accessible à touts les utilisateurs).
- Ouvrir la console PowerShell en tant qu'administrateur AD (!Prenom).
- Entrer "test_computerIDs.ps1" dans la console.
- Choisir un fichier CSV précedement téléchargé sur GLPI.
- Le scan va maintenant commencer.
- À la fin du scan, choisir un fichier CSV pour enregistrer les résultats.

AVERTISSMENT :

- En cas d'édition des scripts BATCH ou powershell,
  il faut veiller à conserver l'encodage windows1252.
  Les scripts PowerShell en UTF-X ne sont pas gérés par
  PowerShell.
- Le fichier base.code-workspace est un fichier de
  configuration pour Microsoft Visual Studio Code.
  Il est possible d'ouvrir ce fichier directement avec VSCode,
  le dossier de travail, l'encodage, etc, seront automatiquement
  définis.
#!/usr/bin/env python3
import os
import sys
import shutil
import subprocess
from pathlib import Path

def create_app_bundle():
    home = Path.home()
    app_name = "LettreMotivationAI"
    
    # Paths
    app_dir = home / "Documents" / f"{app_name}_Partage"
    app_bundle = app_dir / f"{app_name}.app"
    templates_dir = app_dir / "templates"
    
    # Create directories
    os.makedirs(app_dir, exist_ok=True)
    os.makedirs(templates_dir, exist_ok=True)
    
    # Copy application files
    src_dir = Path(__file__).parent
    if (src_dir / "dist" / f"{app_name}.app").exists():
        if app_bundle.exists():
            shutil.rmtree(app_bundle)
        shutil.copytree(src_dir / "dist" / f"{app_name}.app", app_bundle)
    
    # Copy templates
    if (src_dir / "templates").exists():
        for template in (src_dir / "templates").glob("*.txt"):
            shutil.copy2(template, templates_dir)
    
    # Set permissions
    executable = app_bundle / "Contents" / "MacOS" / app_name
    if executable.exists():
        os.chmod(executable, 0o755)
    
    # Create README
    readme = app_dir / "LISEZMOI.txt"
    with open(readme, "w", encoding="utf-8") as f:
        f.write("""Générateur de Lettre de Motivation - Guide d'installation
Version Universelle (Compatible Intel et Apple Silicon)

1. Installation
- Double-cliquez sur l'application "LettreMotivationAI"
- Si un message de sécurité apparaît :
  1. Faites un clic droit sur l'application
  2. Sélectionnez "Ouvrir"
  3. Cliquez sur "Ouvrir" dans la fenêtre de confirmation

Si le problème persiste :
  1. Allez dans le menu Apple () > Préférences Système
  2. Cliquez sur "Sécurité et confidentialité"
  3. Dans l'onglet "Général", cliquez sur "Ouvrir quand même"

2. Utilisation
- Remplissez les champs demandés (nom, prénom, poste, entreprise, etc.)
- Cliquez sur "Générer" pour créer votre lettre
- La lettre sera sauvegardée au format Word et PDF dans le dossier que vous choisirez

3. Templates personnalisés
- Vous pouvez ajouter vos propres templates dans le dossier "templates"
- Les templates doivent être au format .txt
- Variables disponibles :
  {nom} - Votre nom
  {prenom} - Votre prénom
  {poste} - Le poste visé
  {entreprise} - L'entreprise
  {recruteur} - Le nom du recruteur
  {contenu} - Le contenu généré

Cette version est compatible avec :
- MacBook Pro/Air/iMac avec processeur Intel (2006-2020)
- MacBook Pro/Air/iMac avec processeur Apple Silicon M1/M2/M3 (2020+)

En cas de problème :
1. Assurez-vous que tous les fichiers sont au même endroit (app, dossier templates)
2. Vérifiez que vous avez les droits d'exécution sur l'application
3. Si l'application ne s'ouvre pas, essayez de la réinstaller

Bon usage !""")
    
    print(f"\nApplication installée dans : {app_dir}")
    print("Pour lancer l'application :")
    print("1. Ouvrez le dossier Documents")
    print(f"2. Ouvrez le dossier {app_name}_Partage")
    print(f"3. Faites un clic droit sur {app_name}.app")
    print('4. Sélectionnez "Ouvrir"')
    print('5. Cliquez sur "Ouvrir" dans la fenêtre de confirmation')

if __name__ == "__main__":
    create_app_bundle()

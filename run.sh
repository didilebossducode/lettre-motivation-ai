#!/bin/bash

# Vérifier si Python 3 est installé
if ! command -v python3 &> /dev/null; then
    echo "Python 3 n'est pas installé. Installation nécessaire."
    exit 1
fi

# Installer pip si nécessaire
python3 -m ensurepip --upgrade

# Installer les dépendances
python3 -m pip install -r requirements.txt

# Lancer l'application
python3 app.py

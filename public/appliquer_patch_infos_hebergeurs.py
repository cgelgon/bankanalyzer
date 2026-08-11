#!/usr/bin/env python3
"""
Patch : remplace les placeholders [À VÉRIFIER] concernant les hébergeurs
(Vercel, Railway) par les informations réelles, dans mentions-legales.html
et politique-confidentialite.html.
"""
import os
import shutil
from datetime import datetime


def patch(fichier, remplacements):
    if not os.path.exists(fichier):
        print(f"ERREUR : {fichier} introuvable dans le dossier courant.")
        return

    with open(fichier, "r", encoding="utf-8") as f:
        contenu = f.read()

    modifie = False
    for ancien, nouveau in remplacements:
        if ancien not in contenu:
            print(f"ATTENTION ({fichier}) : ancre introuvable, ignorée :")
            print(f"   {ancien[:70]}...")
            continue
        if contenu.count(ancien) > 1:
            print(f"ERREUR ({fichier}) : ancre ambiguë (plusieurs occurrences), ignorée :")
            print(f"   {ancien[:70]}...")
            continue
        contenu = contenu.replace(ancien, nouveau)
        modifie = True

    if not modifie:
        print(f"Aucune modification appliquée sur {fichier}.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{fichier}.avant_patch_infos_hebergeurs_{timestamp}.bak"
    shutil.copy(fichier, backup_path)
    print(f"Sauvegarde créée : {backup_path}")

    with open(fichier, "w", encoding="utf-8") as f:
        f.write(contenu)
    print(f"PATCH APPLIQUE AVEC SUCCES sur {fichier}.")


patch("mentions-legales.html", [
    (
        "[À VÉRIFIER — adresse actuelle sur vercel.com, section legal/terms]",
        "440 N Barranca Avenue #4133, Covina, CA 91723, États-Unis",
    ),
    (
        "[À VÉRIFIER — adresse actuelle sur railway.com, section legal/terms]",
        "548 Market St PMB 68956, San Francisco, CA 94104, États-Unis — "
        "infrastructure hébergée en région UE. Représentant UE (RGPD) : "
        "DP-Dock GmbH, Ballindamm 39, 20095 Hambourg, Allemagne "
        "(railway-corp@gdpr-rep.com)",
    ),
])

patch("politique-confidentialite.html", [
    (
        "[À vérifier : région d'hébergement (UE ou hors UE) pour confirmer "
        "si un mécanisme de transfert est nécessaire.]",
        "Infrastructure hébergée en région UE. Railway Corporation "
        "(société américaine) a désigné un représentant UE au titre du "
        "RGPD : DP-Dock GmbH (Hambourg, Allemagne).",
    ),
])

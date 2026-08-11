#!/usr/bin/env python3
"""
Patch : ajoute une case à cocher de renonciation au droit de rétractation
dans la modale d'upgrade Pro, et bloque le passage à Stripe Checkout si
elle n'est pas cochée.
"""
import os
import shutil
from datetime import datetime

FICHIER = "index.html"

ANCRE1 = r"""    '<p>Cette analyse utilise une fonctionnalité Pro (plusieurs fichiers, patrimoine ou plusieurs devises). Débloquez tout pour 9€/mois.</p>'+
    '<button class="upgrade-btn" onclick="demarrerCheckout()">S\'abonner — 9€/mois</button>'+"""

NOUVEAU1 = r"""    '<p>Cette analyse utilise une fonctionnalité Pro (plusieurs fichiers, patrimoine ou plusieurs devises). Débloquez tout pour 9€/mois.</p>'+
    '<label style="display:flex;align-items:flex-start;gap:8px;text-align:left;font-size:12px;color:var(--text2);margin-bottom:14px;cursor:pointer"><input type="checkbox" id="cguCheckbox" style="margin-top:2px"><span>J\'accepte les <a href="/cgv.html" target="_blank" style="color:inherit;text-decoration:underline">CGV</a> et je demande l\'exécution immédiate du service, renonçant à mon droit de rétractation de 14 jours.</span></label>'+
    '<button class="upgrade-btn" onclick="demarrerCheckout()">S\'abonner — 9€/mois</button>'+"""

ANCRE2 = r"""async function demarrerCheckout(){
  const email=(document.getElementById('emailInput').value||'').trim();"""

NOUVEAU2 = r"""async function demarrerCheckout(){
  const cguCheckbox=document.getElementById('cguCheckbox');
  if(!cguCheckbox||!cguCheckbox.checked){
    alert('Merci d\'accepter les CGV pour continuer.');
    return;
  }
  const email=(document.getElementById('emailInput').value||'').trim();"""


def main():
    if not os.path.exists(FICHIER):
        print(f"ERREUR : {FICHIER} introuvable dans le dossier courant.")
        return

    with open(FICHIER, "r", encoding="utf-8") as f:
        contenu = f.read()

    erreurs = []
    for nom, ancre in [("modale (HTML)", ANCRE1), ("fonction demarrerCheckout (JS)", ANCRE2)]:
        if ancre not in contenu:
            erreurs.append(f"Ancre introuvable : {nom}")
        elif contenu.count(ancre) > 1:
            erreurs.append(f"Ancre ambiguë (plusieurs occurrences) : {nom}")

    if erreurs:
        print("ERREUR : patch annulé, aucune modification appliquée.")
        for e in erreurs:
            print(f"  - {e}")
        print("Le fichier a peut-être changé depuis la dernière vérification.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{FICHIER}.avant_patch_case_retractation_{timestamp}.bak"
    shutil.copy(FICHIER, backup_path)
    print(f"Sauvegarde créée : {backup_path}")

    nouveau_contenu = contenu.replace(ANCRE1, NOUVEAU1).replace(ANCRE2, NOUVEAU2)

    with open(FICHIER, "w", encoding="utf-8") as f:
        f.write(nouveau_contenu)

    print("PATCH APPLIQUE AVEC SUCCES.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Patch : ajoute un pied de page avec liens vers les mentions légales,
la politique de confidentialité et les CGV, juste avant </body>.
"""
import os
import shutil
from datetime import datetime

FICHIER = "index.html"

ANCRE = "</script>\n</body>\n</html>"

FOOTER = """</script>
<footer style="text-align:center;padding:32px 16px;font-size:13px;color:#86868b;">
  <a href="/mentions-legales.html" style="color:#86868b;text-decoration:none;margin:0 8px;">Mentions légales</a>
  ·
  <a href="/politique-confidentialite.html" style="color:#86868b;text-decoration:none;margin:0 8px;">Politique de confidentialité</a>
  ·
  <a href="/cgv.html" style="color:#86868b;text-decoration:none;margin:0 8px;">CGV</a>
</footer>
</body>
</html>"""

def main():
    if not os.path.exists(FICHIER):
        print(f"ERREUR : {FICHIER} introuvable dans le dossier courant.")
        print("Vérifie que tu es bien dans le dossier public/ avant de lancer ce script.")
        return

    with open(FICHIER, "r", encoding="utf-8") as f:
        contenu = f.read()

    if ANCRE not in contenu:
        print("ERREUR : l'ancre attendue est introuvable dans index.html.")
        print("Le fichier a peut-être déjà été modifié. Aucune modification appliquée.")
        return

    if contenu.count(ANCRE) > 1:
        print("ERREUR : l'ancre apparaît plusieurs fois, patch annulé par sécurité.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{FICHIER}.avant_patch_footer_legal_{timestamp}.bak"
    shutil.copy(FICHIER, backup_path)
    print(f"Sauvegarde créée : {backup_path}")

    nouveau_contenu = contenu.replace(ANCRE, FOOTER)
    with open(FICHIER, "w", encoding="utf-8") as f:
        f.write(nouveau_contenu)

    print("PATCH APPLIQUE AVEC SUCCES.")

if __name__ == "__main__":
    main()

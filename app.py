import os, io, json, re
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, request, jsonify
from flask_cors import CORS
import pypdf
import anthropic
import openpyxl

app = Flask(__name__)
CORS(app)

MOIS_MAP = {
    'janvier': 1, 'fevrier': 2, 'mars': 3, 'avril': 4, 'mai': 5, 'juin': 6,
    'juillet': 7, 'aout': 8, 'septembre': 9, 'octobre': 10, 'novembre': 11, 'decembre': 12
}


def extract_text_from_pdf(file_bytes):
    text = ''
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    for page in reader.pages:
        t = page.extract_text()
        if t:
            text += t + chr(10)
    return text


def extract_text_from_csv(file_bytes):
    return file_bytes.decode('utf-8', errors='ignore')


def extract_text_from_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active
    lines = []
    for row in ws.iter_rows(values_only=True):
        lines.append(' | '.join([str(c) if c else '' for c in row]))
    return chr(10).join(lines)


def get_periode(text):
    mois = r'(janvier|fevrier|mars|avril|mai|juin|juillet|aout|septembre|octobre|novembre|decembre)'
    m = re.search(mois + r'\s*\d{4}', text.lower())
    if m:
        return m.group(0).capitalize()
    m2 = re.search(r'(\d{1,2})[/-](\d{4})', text)
    if m2:
        return m2.group(0)
    return 'Periode inconnue'


def periode_sort_key(periode):
    """Retourne une cle (annee, mois) triable a partir d'une chaine periode. (0,0) si non reconnue."""
    if not periode:
        return (0, 0)
    p_low = periode.lower()
    for nom, num in MOIS_MAP.items():
        m = re.search(nom + r'\s*(\d{4})', p_low)
        if m:
            return (int(m.group(1)), num)
    m2 = re.search(r'(\d{1,2})[/-](\d{4})', periode)
    if m2:
        return (int(m2.group(2)), int(m2.group(1)))
    return (0, 0)


def analyse_releve(client, text, nom_banque, langue='français'):
    prompt = (
        'INSTRUCTION ABSOLUE: Tu dois repondre UNIQUEMENT en ' + langue + ', y compris le score_detail, le commentaire, les titres et details des actions. Aucun mot en francais si la langue demandee est differente. Analyse ce releve bancaire (' + nom_banque + ').' + chr(10) +
        'Retourne UNIQUEMENT ce JSON sans markdown:' + chr(10) +
        '{"totalRecettes":0,"totalDepenses":0,"soldeDepart":0,"soldeArrivee":0,' +
        '"recettes":[{"label":"cat","montant":0}],' +
        '"depenses":[{"label":"cat","montant":0}],' +
        '"top5depenses":[{"libelle":"desc","montant":0,"date":"JJ/MM"}],' +
        '"score":7,"score_detail":"phrase"}' + chr(10) +
        'REGLES:' + chr(10) +
        '1. Cherche EN PREMIER les totaux recapitulatifs (Total operations entrantes/sortantes, Solde initial, Solde final)' + chr(10) +
        '2. Utilise ces totaux pour totalRecettes et totalDepenses' + chr(10) +
        '3. soldeDepart = solde debut du releve, soldeArrivee = solde fin' + chr(10) +
        '4. top5depenses = 5 plus grosses transactions sortantes individuelles avec libelle et date' + chr(10) +
        '5. Montants entiers positifs, max 5 recettes, max 7 depenses' + chr(10) +
        chr(10) + 'Releve:' + chr(10) + text[:20000]
    )
    msg = client.messages.create(
        model='claude-sonnet-4-6',
        max_tokens=4000,
        system='Tu es un expert-comptable. Tu reponds UNIQUEMENT avec du JSON valide, sans aucun texte avant ou apres, sans markdown. IMPORTANT: toutes les valeurs textuelles du JSON (labels de categories, commentaire, score_detail, titres et details des actions) doivent etre redigees dans la langue specifiee dans le prompt utilisateur.',
        messages=[{'role': 'user', 'content': prompt}]
    )
    raw = msg.content[0].text.replace('```json', '').replace('```', '').strip()
    print('CLAUDE RESPONSE:', raw[:200])
    return json.loads(raw)


def get_conseil_global(client, comptes, total_r, total_d, periode, langue='francais'):
    comptes_str = chr(10).join(['- ' + c['nom'] + ' (' + c.get('periode', '') + '): recettes ' + str(c['totalRecettes']) + 'EUR, depenses ' + str(c['totalDepenses']) + 'EUR' for c in comptes])
    net = total_r - total_d
    taux = round(net / total_r * 100) if total_r else 0
    langue_map = {'francais': 'French', 'english': 'English', 'espanol': 'Spanish', 'deutsch': 'German', 'italiano': 'Italian', 'portugues': 'Portuguese', 'chinese': 'Chinese', 'arabic': 'Arabic'}
    langue_name = langue_map.get(langue, 'French')
    prompt = (
        'You are a senior financial expert. Write ALL text EXCLUSIVELY in ' + langue_name + '.' + chr(10) +
        'Financial data for ' + periode + ':' + chr(10) +
        comptes_str + chr(10) +
        'TOTAL: income=' + str(total_r) + 'EUR expenses=' + str(total_d) + 'EUR net=' + str(net) + 'EUR savings_rate=' + str(taux) + '%' + chr(10) +
        'SCORING (be strict):' + chr(10) +
        '- 9-10: savings>30% AND positive net AND diversified income' + chr(10) +
        '- 7-8: savings 10-30% AND positive net' + chr(10) +
        '- 5-6: savings 0-10% OR slightly negative' + chr(10) +
        '- 3-4: savings -30% to 0%' + chr(10) +
        '- 1-2: savings < -30% OR deficit > 20% of income' + chr(10) +
        'Current savings rate=' + str(taux) + '% -> apply strictly.' + chr(10) +
        'PHRASE_CHOC: One single punchy sentence (max 15 words) that hits hard. Examples:' + chr(10) +
        '- "At this rate, your savings will be gone in 3 months."' + chr(10) +
        '- "You save the equivalent of a new iPhone every month."' + chr(10) +
        '- "Your biggest hidden expense: 4 forgotten subscriptions."' + chr(10) +
        '- "Warning: 2 of your 3 accounts are in the red."' + chr(10) +
        'Make it personal, specific with real numbers from the data, and impactful.' + chr(10) + chr(10) +
        'Return ONLY valid JSON, ALL text in ' + langue_name + ':' + chr(10) +
        '{"score":0,"score_detail":"sentence","phrase_choc":"impactful sentence with real numbers",' +
        '"actions":[{"priorite":1,"titre":"title","detail":"detail with numbers"},' +
        '{"priorite":2,"titre":"title","detail":"detail with numbers"},' +
        '{"priorite":3,"titre":"title","detail":"detail with numbers"}],' +
        '"commentaire":"2 sentences"}'
    )
    msg = client.messages.create(
        model='claude-sonnet-4-6',
        max_tokens=1000,
        messages=[{'role': 'user', 'content': prompt}]
    )
    raw = msg.content[0].text.replace('```json', '').replace('```', '').strip()
    return json.loads(raw)


@app.route('/analyze', methods=['POST'])
def analyze():
    langue = request.form.get('langue', 'français')
    files = request.files.getlist('files')
    if not files:
        f = request.files.get('file')
        if f:
            files = [f]
        else:
            return jsonify({'error': 'Aucun fichier recu'}), 400
    files = files[:60]

    client = anthropic.Anthropic()

    # Etape 1 : extraction du texte de chaque fichier (rapide, local, sequentiel)
    fichiers_prepares = []
    fichiers_ignores = []
    for file in files:
        filename = file.filename.lower()
        file_bytes = file.read()
        if filename.endswith('.pdf'):
            text = extract_text_from_pdf(file_bytes)
        elif filename.endswith('.csv'):
            text = extract_text_from_csv(file_bytes)
        elif filename.endswith(('.xlsx', '.xls')):
            text = extract_text_from_excel(file_bytes)
        else:
            fichiers_ignores.append({'nom': file.filename, 'raison': 'Format non supporte'})
            continue
        if not text.strip():
            fichiers_ignores.append({'nom': file.filename, 'raison': 'Aucun texte extrait (PDF scanne/image ?)'})
            continue
        nom_banque = file.filename.replace('.pdf', '').replace('.csv', '').replace('.xlsx', '')[:30]
        periode = get_periode(text)
        fichiers_prepares.append({'nom': nom_banque, 'nomFichier': file.filename, 'periode': periode, 'text': text})

    if not fichiers_prepares:
        return jsonify({'error': 'Aucun releve analyse'}), 500

    # Etape 2 : appels IA en parallele (par lots pour respecter les limites de debit de l'API)
    comptes = []
    TAILLE_LOT = 8

    def analyser_un_fichier(f):
        derniere_erreur = None
        for tentative in range(2):
            try:
                data = analyse_releve(client, f['text'], f['nom'], langue)
                data['nom'] = f['nom']
                data['periode'] = f['periode']
                return data
            except Exception as e:
                derniere_erreur = e
        raise derniere_erreur

    for i in range(0, len(fichiers_prepares), TAILLE_LOT):
        lot = fichiers_prepares[i:i + TAILLE_LOT]
        with ThreadPoolExecutor(max_workers=TAILLE_LOT) as executor:
            futures = {executor.submit(analyser_un_fichier, f): f for f in lot}
            for future in as_completed(futures):
                f = futures[future]
                try:
                    comptes.append(future.result())
                except Exception as e:
                    print('ERREUR releve', f['nom'], str(e))
                    err_msg = str(e)[:150]
                    fichiers_ignores.append({'nom': f['nomFichier'], 'raison': "Echec de l'analyse IA : " + err_msg})
                    continue

    if not comptes:
        return jsonify({'error': 'Aucun releve analyse', 'fichiersIgnores': fichiers_ignores}), 500

    # --- Detection automatique multi-mois vs multi-comptes ---
    periodes_uniques = sorted(
        set(c.get('periode', 'Periode inconnue') for c in comptes),
        key=periode_sort_key
    )
    is_multi_mois = len(periodes_uniques) > 1

    evolution = []
    if is_multi_mois:
        premiere_periode = periodes_uniques[0]
        derniere_periode = periodes_uniques[-1]
        comptes_premiere = [c for c in comptes if c.get('periode', 'Periode inconnue') == premiere_periode]
        comptes_derniere = [c for c in comptes if c.get('periode', 'Periode inconnue') == derniere_periode]
        solde_depart = sum(c.get('soldeDepart', 0) for c in comptes_premiere)
        solde_arrivee = sum(c.get('soldeArrivee', 0) for c in comptes_derniere)
        periode_label = premiere_periode + ' -> ' + derniere_periode

        # Point de depart : solde d'ouverture du tout premier mois, avant tout mouvement
        evolution.append({
            'periode': 'Debut ' + premiere_periode,
            'totalRecettes': 0,
            'totalDepenses': 0,
            'net': 0,
            'soldeArrivee': solde_depart
        })
        for p in periodes_uniques:
            comptes_p = [c for c in comptes if c.get('periode', 'Periode inconnue') == p]
            tr_p = sum(c.get('totalRecettes', 0) for c in comptes_p)
            td_p = sum(c.get('totalDepenses', 0) for c in comptes_p)
            sa_p = sum(c.get('soldeArrivee', 0) for c in comptes_p)
            evolution.append({
                'periode': p,
                'totalRecettes': tr_p,
                'totalDepenses': td_p,
                'net': tr_p - td_p,
                'soldeArrivee': sa_p
            })
    else:
        solde_depart = sum(c.get('soldeDepart', 0) for c in comptes)
        solde_arrivee = sum(c.get('soldeArrivee', 0) for c in comptes)
        periode_label = periodes_uniques[0] if periodes_uniques else 'Periode inconnue'

    # Vue d'ensemble = toujours la somme de TOUS les mois/comptes envoyes
    total_r = sum(c.get('totalRecettes', 0) for c in comptes)
    total_d = sum(c.get('totalDepenses', 0) for c in comptes)

    all_rec = {}
    all_dep = {}
    for c in comptes:
        for r in c.get('recettes', []):
            all_rec[r['label']] = all_rec.get(r['label'], 0) + r['montant']
        for d in c.get('depenses', []):
            all_dep[d['label']] = all_dep.get(d['label'], 0) + d['montant']

    rec_global = sorted([{'label': k, 'montant': v} for k, v in all_rec.items()], key=lambda x: -x['montant'])[:5]
    dep_global = sorted([{'label': k, 'montant': v} for k, v in all_dep.items()], key=lambda x: -x['montant'])[:7]

    try:
        conseil = get_conseil_global(client, comptes, total_r, total_d, periode_label, langue)
    except Exception:
        conseil = {'score': 5, 'score_detail': 'Analyse partielle', 'actions': [], 'commentaire': 'Analyse disponible.'}

    all_top5 = []
    for c in comptes:
        all_top5.extend(c.get('top5depenses', []))
    all_top5 = sorted(all_top5, key=lambda x: -x.get('montant', 0))[:5]

    result = {
        'periode': periode_label,
        'isMultiMois': is_multi_mois,
        'evolution': evolution,
        'fichiersIgnores': fichiers_ignores,
        'totalRecettes': total_r,
        'totalDepenses': total_d,
        'soldeDepart': solde_depart,
        'soldeArrivee': solde_arrivee,
        'top5depenses': all_top5,
        'recettes': rec_global,
        'depenses': dep_global,
        'phrase_choc': conseil.get('phrase_choc', ''),
        'score': conseil.get('score', 5),
        'score_detail': conseil.get('score_detail', ''),
        'actions': conseil.get('actions', []),
        'commentaire': conseil.get('commentaire', ''),
        'comptes': comptes
    }
    print('PHRASE_CHOC:', result.get('phrase_choc', 'VIDE'))
    return jsonify(result)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5001)))

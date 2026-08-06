import os, io, json, re, unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, request, jsonify
from flask_cors import CORS
import pypdf
import anthropic
import openpyxl

app = Flask(__name__)
CORS(app)

MOIS_MAP = {
    # francais
    'janvier': 1, 'fevrier': 2, 'mars': 3, 'avril': 4, 'mai': 5, 'juin': 6,
    'juillet': 7, 'aout': 8, 'septembre': 9, 'octobre': 10, 'novembre': 11, 'decembre': 12,
    # english
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
    # espanol
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
    'julio': 7, 'agosto': 8, 'septiembre': 9, 'setiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12,
    # deutsch
    'januar': 1, 'februar': 2, 'marz': 3, 'juni': 6, 'juli': 7,
    'oktober': 10, 'dezember': 12,
    # italiano
    'gennaio': 1, 'febbraio': 2, 'aprile': 4, 'maggio': 5, 'giugno': 6, 'luglio': 7,
    'settembre': 9, 'ottobre': 10, 'dicembre': 12,
    # portugues
    'janeiro': 1, 'fevereiro': 2, 'marco': 3, 'maio': 5, 'junho': 6, 'julho': 7,
    'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12,
    # arabe
    'يناير': 1, 'فبراير': 2, 'مارس': 3, 'أبريل': 4, 'مايو': 5, 'يونيو': 6,
    'يوليو': 7, 'أغسطس': 8, 'سبتمبر': 9, 'أكتوبر': 10, 'نوفمبر': 11, 'ديسمبر': 12,
}


def sans_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def to_num(v):
    """Convertit une valeur (potentiellement une chaine renvoyee par l'IA) en nombre, sans jamais planter."""
    if isinstance(v, (int, float)):
        return v
    if v is None:
        return 0.0
    try:
        s = str(v).strip().replace('\u202f', '').replace(' ', '').replace(',', '.')
        return float(s)
    except (TypeError, ValueError):
        return 0.0


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
    # Format chinois : "2026年7月" ou "7月2026年"
    m_cn = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月', text)
    if m_cn:
        return m_cn.group(2) + '/' + m_cn.group(1)
    m_cn2 = re.search(r'(\d{1,2})\s*月\s*(\d{4})\s*年', text)
    if m_cn2:
        return m_cn2.group(1) + '/' + m_cn2.group(2)

    texte_sans_accents = sans_accents(text.lower())
    noms_mois = '|'.join(sorted(MOIS_MAP.keys(), key=len, reverse=True))
    m = re.search('(' + noms_mois + r')\s*\d{4}', texte_sans_accents)
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
    p_low = sans_accents(periode.lower())
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
        '{"compte":"nom de la banque et/ou du compte tel qu\'il apparait EXPLICITEMENT sur le releve (ex: REVOLUT, BNP PARIBAS - Compte Principal, Compte Booster)",' +
        '"devise":"code devise ISO du releve tel qu\'indique dessus (EUR, JOD, USD, GBP, etc.)",' +
        '"totalRecettes":0,"totalDepenses":0,"soldeDepart":0,"soldeArrivee":0,' +
        '"totalRecettesOfficiel":0,"totalDepensesOfficiel":0,' +
        '"recettes":[{"label":"cat","montant":0,"transactions":[{"libelle":"desc","montant":0,"date":"JJ/MM"}]}],' +
        '"depenses":[{"label":"cat","montant":0,"transactions":[{"libelle":"desc","montant":0,"date":"JJ/MM"}]}],' +
        '"top5depenses":[{"libelle":"desc","montant":0,"date":"JJ/MM"}],' +
        '"score":7,"score_detail":"phrase"}' + chr(10) +
        'REGLES:' + chr(10) +
        '1. Cherche EN PREMIER les totaux recapitulatifs (Total operations entrantes/sortantes, Solde initial, Solde final)' + chr(10) +
        '2. Utilise ces totaux pour totalRecettes et totalDepenses' + chr(10) +
        '3. soldeDepart = solde debut du releve, soldeArrivee = solde fin' + chr(10) +
        '4. top5depenses = 5 plus grosses transactions sortantes individuelles avec libelle et date' + chr(10) +
        '5. Montants entiers positifs, max 5 recettes, max 7 depenses' + chr(10) +
        '6b. Si le releve affiche un recapitulatif officiel imprime des totaux (ex: "Total des operations", "TOTAL", "Amount of Transactions", "Number/Amount of Transactions Debit/Credit"), rapporte ces totaux exacts dans totalRecettesOfficiel et totalDepensesOfficiel (arrondis a l\'entier). Si aucun recapitulatif officiel n\'est visible sur le releve, mets exactement les memes valeurs que totalRecettes et totalDepenses.' + chr(10) +
        '6. Pour CHAQUE categorie de recettes et de depenses, liste dans "transactions" jusqu\'a 5 transactions individuelles les plus importantes qui la composent (libelle, montant, date)' + chr(10) +
        '7. Pour "compte", identifie le nom de la banque et/ou du compte TEL QU\'IL APPARAIT sur le releve (logo, en-tete, intitule de compte). Si tu ne trouves rien de clair, mets "' + nom_banque + '"' + chr(10) +
        chr(10) + 'Releve:' + chr(10) + text[:20000]
    )
    msg = client.messages.create(
        model='claude-sonnet-4-6',
        max_tokens=4000,
        system='Tu es un expert-comptable. Tu reponds UNIQUEMENT avec du JSON valide, sans aucun texte avant ou apres, sans markdown, sans phrase d\'introduction ni de raisonnement visible (meme sur des releves complexes ou volumineux, va directement au JSON final). IMPORTANT: toutes les valeurs textuelles du JSON (labels de categories, commentaire, score_detail, titres et details des actions) doivent etre redigees dans la langue specifiee dans le prompt utilisateur.',
        messages=[{'role': 'user', 'content': prompt}]
    )
    raw = msg.content[0].text.replace('```json', '').replace('```', '').strip() if msg.content else ''
    stop_reason = getattr(msg, 'stop_reason', None)
    print('CLAUDE RESPONSE:', raw[:200], '| stop_reason:', stop_reason, '| nb_blocks:', len(msg.content), '| longueur_texte_source:', len(text))
    if not raw:
        raise ValueError('Reponse vide de Claude (stop_reason=' + str(stop_reason) + ', longueur texte source=' + str(len(text)) + ' caracteres)')
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # L'IA a parfois ajoute une phrase de raisonnement avant/apres le JSON malgre la consigne.
        # On tente de recuperer uniquement le bloc JSON (premiere { a derniere }).
        start = raw.find('{')
        end = raw.rfind('}')
        if start != -1 and end != -1 and end > start:
            try:
                return json.loads(raw[start:end + 1])
            except json.JSONDecodeError:
                pass
        raise


def get_conseil_global(client, comptes, total_r, total_d, periode, langue='francais', patrimoine=None, devise='EUR'):
    comptes_str = chr(10).join(['- ' + c['nom'] + ' (' + c.get('periode', '') + '): recettes ' + str(c['totalRecettes']) + devise + ', depenses ' + str(c['totalDepenses']) + devise for c in comptes])
    net = total_r - total_d
    taux = round(net / total_r * 100) if total_r else 0
    langue_map = {'francais': 'French', 'english': 'English', 'espanol': 'Spanish', 'deutsch': 'German', 'italiano': 'Italian', 'portugues': 'Portuguese', 'chinese': 'Chinese', 'arabic': 'Arabic'}
    langue_name = langue_map.get(langue, 'French')

    patrimoine_str = ''
    if patrimoine:
        patrimoine_str = (
            chr(10) + 'ADDITIONAL CONTEXT - USER-DECLARED NET WORTH (beyond the bank flows above):' + chr(10) +
            '- Owned assets: ' + str(round(patrimoine['totalActifs'])) + devise + chr(10) +
            '- Savings/investments: ' + str(round(patrimoine['totalEpargnes'])) + devise + chr(10) +
            '- Remaining debts/loans: ' + str(round(patrimoine['totalDettes'])) + devise + chr(10) +
            '- Estimated net worth: ' + str(round(patrimoine['patrimoineNet'])) + devise + chr(10) +
            'Factor this into your score and advice: a tight monthly cash flow matters less if net worth is solid, and vice versa. Mention net worth explicitly if it materially changes the picture.' + chr(10)
        )

    prompt = (
        'You are a senior financial expert. Write ALL text EXCLUSIVELY in ' + langue_name + '. The currency of ALL amounts is ' + devise + ' - use this currency code (' + devise + ') everywhere you mention a monetary amount, NEVER write EUR or any other currency code.' + chr(10) +
        'Financial data for ' + periode + ':' + chr(10) +
        comptes_str + chr(10) +
        'TOTAL: income=' + str(total_r) + devise + ' expenses=' + str(total_d) + devise + ' net=' + str(net) + devise + ' savings_rate=' + str(taux) + '%' + chr(10) +
        patrimoine_str +
        'SCORING (be strict):' + chr(10) +
        '- 9-10: savings>30% AND positive net AND diversified income' + chr(10) +
        '- 7-8: savings 10-30% AND positive net' + chr(10) +
        '- 5-6: savings 0-10% OR slightly negative' + chr(10) +
        '- 3-4: savings -30% to 0%' + chr(10) +
        '- 1-2: savings < -30% OR deficit > 20% of income' + chr(10) +
        'Current savings rate=' + str(taux) + '% -> apply strictly.' + chr(10) +
        'PHRASE_CHOC: One single punchy sentence (max 15 words) that hits hard, using the ' + devise + ' currency code. Examples:' + chr(10) +
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
    try:
        return _analyze_impl()
    except Exception as e:
        import traceback
        print('ERREUR FATALE /analyze:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': "Erreur serveur : " + str(e)[:300]}), 500


def _analyze_impl():
    langue = request.form.get('langue', 'français')
    files = request.files.getlist('files')

    mode_devise = request.form.get('modeDevise', 'unique')
    devise_unique = (request.form.get('deviseUnique') or 'EUR').upper().strip()
    devise_reference = (request.form.get('deviseReference') or 'EUR').upper().strip()
    devise_principale = devise_unique if mode_devise == 'unique' else devise_reference

    devises_raw = request.form.get('devises', '')
    taux_par_devise = {}
    if devises_raw:
        try:
            liste_taux = json.loads(devises_raw)
            for item in liste_taux:
                code = (item.get('code') or '').upper().strip()
                taux = to_num(item.get('taux'))
                if code and taux > 0:
                    taux_par_devise[code] = taux
        except (json.JSONDecodeError, AttributeError, TypeError):
            taux_par_devise = {}

    def convertir_montant_patrimoine(montant, devise_entree, taux_ligne):
        m = to_num(montant)
        devise_entree = (devise_entree or devise_principale).upper().strip()
        if devise_entree == devise_principale:
            return m
        t = to_num(taux_ligne)
        if t > 0:
            return m * t
        if devise_entree in taux_par_devise:
            return m * taux_par_devise[devise_entree]
        return m  # aucun taux disponible : on garde la valeur brute (imprecis mais on ne bloque pas)

    patrimoine_raw = request.form.get('patrimoine', '')
    patrimoine_resume = None
    if patrimoine_raw:
        try:
            patrimoine = json.loads(patrimoine_raw)
            emprunts = patrimoine.get('emprunts', []) or []
            epargnes = patrimoine.get('epargnes', []) or []
            actifs = patrimoine.get('actifs', []) or []
            if emprunts or epargnes or actifs:
                total_actifs = sum(convertir_montant_patrimoine(a.get('valeur'), a.get('devise'), a.get('tauxConversion')) for a in actifs)
                total_epargnes = sum(convertir_montant_patrimoine(e.get('montant'), e.get('devise'), e.get('tauxConversion')) for e in epargnes)
                total_dettes = sum(convertir_montant_patrimoine(e.get('capitalRestantDu') or e.get('montantInitial'), e.get('devise'), e.get('tauxConversion')) for e in emprunts)
                patrimoine_resume = {
                    'totalActifs': total_actifs,
                    'totalEpargnes': total_epargnes,
                    'totalDettes': total_dettes,
                    'patrimoineNet': total_actifs + total_epargnes - total_dettes,
                    'emprunts': emprunts,
                    'epargnes': epargnes,
                    'actifs': actifs
                }
        except (json.JSONDecodeError, AttributeError, TypeError):
            patrimoine_resume = None

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

    avertissements = []
    SEUIL_ECART = 0.03  # 3% d'ecart tolere avant d'avertir

    def analyser_un_fichier(f):
        derniere_erreur = None
        for tentative in range(2):
            try:
                data = analyse_releve(client, f['text'], f['nom'], langue)
                nom_detecte = (data.get('compte') or '').strip()
                data['nom'] = nom_detecte if nom_detecte else f['nom']
                data['periode'] = f['periode']

                tr = to_num(data.get('totalRecettes', 0))
                td = to_num(data.get('totalDepenses', 0))
                tr_off = to_num(data.get('totalRecettesOfficiel', tr))
                td_off = to_num(data.get('totalDepensesOfficiel', td))
                ecarts = []
                if tr_off > 0 and abs(tr - tr_off) / tr_off > SEUIL_ECART:
                    ecarts.append('recettes calculees=' + str(round(tr)) + ' vs officiel=' + str(round(tr_off)))
                if td_off > 0 and abs(td - td_off) / td_off > SEUIL_ECART:
                    ecarts.append('depenses calculees=' + str(round(td)) + ' vs officiel=' + str(round(td_off)))
                if ecarts:
                    avertissements.append({
                        'nom': data['nom'] + ' (' + f['periode'] + ')',
                        'raison': 'Ecart avec le recapitulatif officiel du releve : ' + ' ; '.join(ecarts)
                    })
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

    # --- Devise : mode choisi explicitement par l'utilisateur ---
    def convertir_compte_devise(compte, taux, nouvelle_devise):
        c = dict(compte)
        for champ in ['totalRecettes', 'totalDepenses', 'soldeDepart', 'soldeArrivee']:
            if champ in c:
                c[champ] = to_num(c[champ]) * taux
        for cle in ['recettes', 'depenses']:
            nouvelles = []
            for item in c.get(cle, []):
                item2 = dict(item)
                item2['montant'] = to_num(item2.get('montant', 0)) * taux
                item2['transactions'] = [
                    dict(t, montant=to_num(t.get('montant', 0)) * taux) for t in item.get('transactions', [])
                ]
                nouvelles.append(item2)
            c[cle] = nouvelles
        c['top5depenses'] = [
            dict(t, montant=to_num(t.get('montant', 0)) * taux) for t in c.get('top5depenses', [])
        ]
        c['devise'] = nouvelle_devise
        return c

    if mode_devise == 'unique':
        # L'utilisateur affirme que tout est deja dans une seule et meme devise : on force cette devise partout
        for c in comptes:
            c['devise'] = devise_unique
        devise_principale = devise_unique
    else:
        devise_principale = devise_reference
        comptes_convertis = []
        comptes_hors_devise = []
        for c in comptes:
            dv = (c.get('devise') or 'EUR').upper().strip()
            if dv == devise_principale:
                comptes_convertis.append(c)
            elif dv in taux_par_devise:
                comptes_convertis.append(convertir_compte_devise(c, taux_par_devise[dv], devise_principale))
            else:
                comptes_hors_devise.append(c)
        for c in comptes_hors_devise:
            fichiers_ignores.append({
                'nom': c.get('nom', '?') + ' (' + c.get('periode', '') + ')',
                'raison': 'Devise ' + (c.get('devise') or '?') + ' - aucun taux de conversion fourni vers ' + devise_principale
            })
        comptes = comptes_convertis

    if not comptes:
        return jsonify({'error': 'Aucun compte ne correspond a la devise choisie, aucune analyse coherente possible', 'fichiersIgnores': fichiers_ignores}), 500

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
        solde_depart = sum(to_num(c.get('soldeDepart', 0)) for c in comptes_premiere)
        solde_arrivee = sum(to_num(c.get('soldeArrivee', 0)) for c in comptes_derniere)
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
            tr_p = sum(to_num(c.get('totalRecettes', 0)) for c in comptes_p)
            td_p = sum(to_num(c.get('totalDepenses', 0)) for c in comptes_p)
            sa_p = sum(to_num(c.get('soldeArrivee', 0)) for c in comptes_p)
            evolution.append({
                'periode': p,
                'totalRecettes': tr_p,
                'totalDepenses': td_p,
                'net': tr_p - td_p,
                'soldeArrivee': sa_p
            })
    else:
        solde_depart = sum(to_num(c.get('soldeDepart', 0)) for c in comptes)
        solde_arrivee = sum(to_num(c.get('soldeArrivee', 0)) for c in comptes)
        periode_label = periodes_uniques[0] if periodes_uniques else 'Periode inconnue'

    # Vue d'ensemble = toujours la somme de TOUS les mois/comptes envoyes
    total_r = sum(to_num(c.get('totalRecettes', 0)) for c in comptes)
    total_d = sum(to_num(c.get('totalDepenses', 0)) for c in comptes)

    all_rec = {}
    all_rec_tx = {}
    all_dep = {}
    all_dep_tx = {}
    for c in comptes:
        for r in c.get('recettes', []):
            all_rec[r['label']] = all_rec.get(r['label'], 0) + to_num(r.get('montant', 0))
            all_rec_tx.setdefault(r['label'], []).extend(r.get('transactions', []))
        for d in c.get('depenses', []):
            all_dep[d['label']] = all_dep.get(d['label'], 0) + to_num(d.get('montant', 0))
            all_dep_tx.setdefault(d['label'], []).extend(d.get('transactions', []))

    def normaliser_montants(transactions):
        out = []
        for t in transactions:
            t2 = dict(t)
            t2['montant'] = to_num(t2.get('montant', 0))
            out.append(t2)
        return out

    def top_transactions(liste):
        return normaliser_montants(sorted(liste, key=lambda x: -to_num(x.get('montant', 0)))[:8])

    rec_global = sorted(
        [{'label': k, 'montant': v, 'transactions': top_transactions(all_rec_tx.get(k, []))} for k, v in all_rec.items()],
        key=lambda x: -to_num(x['montant'])
    )[:5]
    dep_global = sorted(
        [{'label': k, 'montant': v, 'transactions': top_transactions(all_dep_tx.get(k, []))} for k, v in all_dep.items()],
        key=lambda x: -to_num(x['montant'])
    )[:7]

    try:
        conseil = get_conseil_global(client, comptes, total_r, total_d, periode_label, langue, patrimoine_resume, devise_principale)
    except Exception:
        conseil = {'score': 5, 'score_detail': 'Analyse partielle', 'actions': [], 'commentaire': 'Analyse disponible.'}

    all_top5 = []
    for c in comptes:
        all_top5.extend(c.get('top5depenses', []))
    all_top5 = sorted(all_top5, key=lambda x: -to_num(x.get('montant', 0)))[:5]
    all_top5 = normaliser_montants(all_top5)

    result = {
        'periode': periode_label,
        'devise': devise_principale,
        'isMultiMois': is_multi_mois,
        'evolution': evolution,
        'fichiersIgnores': fichiers_ignores,
        'avertissements': avertissements,
        'patrimoine': patrimoine_resume,
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

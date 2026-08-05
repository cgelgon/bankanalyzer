import os, io, json, re
from flask import Flask, request, jsonify
from flask_cors import CORS
import pypdf
import anthropic
import openpyxl

app = Flask(__name__)
CORS(app)

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
    import re
    m = re.search(mois + r'\s*\d{4}', text.lower())
    if m:
        return m.group(0).capitalize()
    m2 = re.search(r'(\d{1,2})[/-](\d{4})', text)
    if m2:
        return m2.group(0)
    return 'Periode inconnue'

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
        max_tokens=2000,
        system='Tu es un expert-comptable. Tu reponds UNIQUEMENT avec du JSON valide, sans aucun texte avant ou apres, sans markdown. IMPORTANT: toutes les valeurs textuelles du JSON (labels de categories, commentaire, score_detail, titres et details des actions) doivent etre redigees dans la langue specifiee dans le prompt utilisateur.',
        messages=[{'role': 'user', 'content': prompt}]
    )
    raw = msg.content[0].text.replace('```json','').replace('```','').strip()
    print('CLAUDE RESPONSE:', raw[:200])
    return json.loads(raw)

def get_conseil_global(client, comptes, total_r, total_d, periode, langue='francais'):
    comptes_str = chr(10).join(['- ' + c['nom'] + ': recettes ' + str(c['totalRecettes']) + 'EUR, depenses ' + str(c['totalDepenses']) + 'EUR' for c in comptes])
    net = total_r - total_d
    taux = round(net/total_r*100) if total_r else 0
    langue_map = {'francais':'French','english':'English','espanol':'Spanish','deutsch':'German','italiano':'Italian','portugues':'Portuguese','chinese':'Chinese','arabic':'Arabic'}
    langue_name = langue_map.get(langue, 'French')
    prompt = (
        'You are a senior financial expert. Write ALL text fields EXCLUSIVELY in ' + langue_name + '. No other language allowed.' + chr(10) +
        'Financial data for ' + periode + ':' + chr(10) +
        comptes_str + chr(10) +
        'TOTAL: income=' + str(total_r) + 'EUR expenses=' + str(total_d) + 'EUR net=' + str(net) + 'EUR savings_rate=' + str(taux) + '%' + chr(10) +
        'SCORING RULES (be strict and objective):' + chr(10) +
        '- score 9-10: savings rate > 30% AND positive net AND diversified income' + chr(10) +
        '- score 7-8: savings rate 10-30% AND positive net' + chr(10) +
        '- score 5-6: savings rate 0-10% OR slightly negative net' + chr(10) +
        '- score 3-4: savings rate -30% to 0% OR net deficit < 20% of income' + chr(10) +
        '- score 1-2: savings rate < -30% OR net deficit > 20% of income' + chr(10) +
        'Current savings rate is ' + str(taux) + '% -> apply scoring rules strictly.' + chr(10) +
        'Return ONLY valid JSON, ALL text in ' + langue_name + ', no markdown:' + chr(10) +
        '{"score":0,"score_detail":"sentence in ' + langue_name + '","actions":[' +
        '{"priorite":1,"titre":"in ' + langue_name + '","detail":"concrete with numbers in ' + langue_name + '"},' +
        '{"priorite":2,"titre":"in ' + langue_name + '","detail":"concrete with numbers in ' + langue_name + '"},' +
        '{"priorite":3,"titre":"in ' + langue_name + '","detail":"concrete with numbers in ' + langue_name + '"}],' +
        '"commentaire":"2 sentences in ' + langue_name + '"}'
    )
    msg = client.messages.create(
        model='claude-sonnet-4-6',
        max_tokens=800,
        messages=[{'role': 'user', 'content': prompt}]
    )
    raw = msg.content[0].text.replace('```json','').replace('```','').strip()
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
    files = files[:5]
    client = anthropic.Anthropic()
    comptes = []
    periode = 'Periode inconnue'
    for i, file in enumerate(files):
        filename = file.filename.lower()
        file_bytes = file.read()
        if filename.endswith('.pdf'):
            text = extract_text_from_pdf(file_bytes)
        elif filename.endswith('.csv'):
            text = extract_text_from_csv(file_bytes)
        elif filename.endswith(('.xlsx','.xls')):
            text = extract_text_from_excel(file_bytes)
        else:
            continue
        if not text.strip():
            continue
        if i == 0:
            periode = get_periode(text)
        nom_banque = file.filename.replace('.pdf','').replace('.csv','').replace('.xlsx','')[:30]
        try:
            data = analyse_releve(client, text, nom_banque, langue)
            data['nom'] = nom_banque
            data['periode'] = get_periode(text)
            comptes.append(data)
        except Exception as e:
            print('ERREUR releve', nom_banque, str(e))
            continue
    if not comptes:
        return jsonify({'error': 'Aucun releve analyse'}), 500
    total_r = sum(c.get('totalRecettes', 0) for c in comptes)
    total_d = sum(c.get('totalDepenses', 0) for c in comptes)
    all_rec = {}
    all_dep = {}
    for c in comptes:
        for r in c.get('recettes', []):
            all_rec[r['label']] = all_rec.get(r['label'], 0) + r['montant']
        for d in c.get('depenses', []):
            all_dep[d['label']] = all_dep.get(d['label'], 0) + d['montant']
    rec_global = sorted([{'label':k,'montant':v} for k,v in all_rec.items()], key=lambda x:-x['montant'])[:5]
    dep_global = sorted([{'label':k,'montant':v} for k,v in all_dep.items()], key=lambda x:-x['montant'])[:7]
    try:
        conseil = get_conseil_global(client, comptes, total_r, total_d, periode, langue)
    except:
        conseil = {'score':5,'score_detail':'Analyse partielle','actions':[],'commentaire':'Analyse disponible.'}
    # Soldes : somme des soldes de depart et arrivee de tous les comptes
    solde_depart = sum(c.get('soldeDepart', 0) for c in comptes)
    solde_arrivee = sum(c.get('soldeArrivee', 0) for c in comptes)
    # Top5 : fusionner et prendre les 5 plus grosses
    all_top5 = []
    for c in comptes:
        all_top5.extend(c.get('top5depenses', []))
    all_top5 = sorted(all_top5, key=lambda x: -x.get('montant', 0))[:5]
    result = {
        'periode': periode,
        'totalRecettes': total_r,
        'totalDepenses': total_d,
        'soldeDepart': solde_depart,
        'soldeArrivee': solde_arrivee,
        'top5depenses': all_top5,
        'recettes': rec_global,
        'depenses': dep_global,
        'score': conseil.get('score', 5),
        'score_detail': conseil.get('score_detail', ''),
        'actions': conseil.get('actions', []),
        'commentaire': conseil.get('commentaire', ''),
        'comptes': comptes
    }
    return jsonify(result)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5001)))
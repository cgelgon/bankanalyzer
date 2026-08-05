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

def analyse_releve(client, text, nom_banque):
    prompt = ('Tu es un expert-comptable francais. Analyse ce releve bancaire (' + nom_banque + ').\n'
        'Retourne UNIQUEMENT ce JSON sans markdown:\n'
        '{"totalRecettes":0,"totalDepenses":0,'
        '"soldeDepart":0,"soldeArrivee":0,'
        '"recettes":[{"label":"cat","montant":0,"transactions":[{"libelle":"desc","montant":0,"date":"JJ/MM"}]}],'
        '"depenses":[{"label":"cat","montant":0,"transactions":[{"libelle":"desc","montant":0,"date":"JJ/MM"}]}],'
        '"top5depenses":[{"libelle":"desc","montant":0,"date":"JJ/MM"}],'
        '"score":7,"score_detail":"phrase"}\n'
        'REGLES IMPORTANTES:\n'
        '1. Cherche EN PREMIER les totaux recapitulatifs du releve (ex: Total operations entrantes, Solde final, etc)\n'
        '2. Utilise ces totaux comme totalRecettes et totalDepenses - ne les recalcule pas\n'
        '3. Si pas de totaux recapitulatifs, additionne toutes les lignes\n'
        '4. soldeDepart = solde au debut du releve, soldeArrivee = solde a la fin\n'
        '5. top5depenses = les 5 plus grosses transactions sortantes individuelles\n'
        '6. Pour chaque categorie, liste les transactions individuelles qui la composent\n'
        '7. Montants entiers positifs, max 5 recettes, max 7 depenses\n\n'
        'Releve:\n' + text[:20000])
    msg = client.messages.create(
        model='claude-sonnet-4-6',
        max_tokens=4000,
        messages=[{'role': 'user', 'content': prompt}]
    )
    raw = msg.content[0].text.replace('```json','').replace('```','').strip()
    return json.loads(raw)

def get_conseil_global(client, comptes, total_r, total_d, periode):
    comptes_str = chr(10).join(['- ' + c['nom'] + ': recettes ' + str(c['totalRecettes']) + 'EUR, depenses ' + str(c['totalDepenses']) + 'EUR' for c in comptes])
    net = total_r - total_d
    taux = round(net/total_r*100) if total_r else 0
    prompt = ('Tu es un expert-comptable. Situation consolidee de ' + str(len(comptes)) + ' compte(s) pour ' + periode + ':\n'
        + comptes_str + '\n'
        'TOTAL: recettes ' + str(total_r) + 'EUR, depenses ' + str(total_d) + 'EUR, net ' + str(net) + 'EUR, taux epargne ' + str(taux) + '%\n\n'
        'Retourne UNIQUEMENT ce JSON sans markdown:\n'
        '{"score":7,"score_detail":"phrase","actions":['
        '{"priorite":1,"titre":"titre","detail":"detail concret et chiffre"},'
        '{"priorite":2,"titre":"titre","detail":"detail concret et chiffre"},'
        '{"priorite":3,"titre":"titre","detail":"detail concret et chiffre"}],'
        '"commentaire":"2 phrases analyse globale."}')
    msg = client.messages.create(
        model='claude-sonnet-4-6',
        max_tokens=800,
        messages=[{'role': 'user', 'content': prompt}]
    )
    raw = msg.content[0].text.replace('```json','').replace('```','').strip()
    return json.loads(raw)

@app.route('/analyze', methods=['POST'])
def analyze():
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
            data = analyse_releve(client, text, nom_banque)
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
        conseil = get_conseil_global(client, comptes, total_r, total_d, periode)
    except:
        conseil = {'score':5,'score_detail':'Analyse partielle','actions':[],'commentaire':'Analyse disponible.'}
    result = {
        'periode': periode,
        'totalRecettes': total_r,
        'totalDepenses': total_d,
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
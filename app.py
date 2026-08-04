import os, io, json, re
from flask import Flask, request, jsonify
from flask_cors import CORS
import pypdf
import anthropic
import openpyxl

app = Flask(__name__)
CORS(app)

def extract_text_from_pdf(file_bytes):
    text = ""
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    for page in reader.pages:
        t = page.extract_text()
        if t:
            text += t + "\n"
    return text

def extract_text_from_csv(file_bytes):
    return file_bytes.decode("utf-8", errors="ignore")

def extract_text_from_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active
    lines = []
    for row in ws.iter_rows(values_only=True):
        lines.append(" | ".join([str(c) if c else "" for c in row]))
    return "\n".join(lines)

def analyse_chunk(client, chunk, idx, total):
    prompt = (
        "Tu es un expert-comptable. Extrait toutes les transactions de ce bloc "
        f"({idx}/{total}) d'un releve bancaire. "
        "Retourne UNIQUEMENT ce JSON sans markdown:\n"
        '{"recettes":[{"label":"cat","montant":0}],'
        '"depenses":[{"label":"cat","montant":0}]}\n'
        "Regle: regroupe par categorie logique, montants entiers positifs, "
        "ignore virements entre comptes du meme titulaire.\n\n"
        "Bloc:\n" + chunk
    )
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1500,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = msg.content[0].text.replace("```json","").replace("```","").strip()
    return json.loads(raw)

def consolide(chunks_data, periode):
    recettes = {}
    depenses = {}
    for d in chunks_data:
        for r in d.get("recettes", []):
            recettes[r["label"]] = recettes.get(r["label"], 0) + r["montant"]
        for d2 in d.get("depenses", []):
            depenses[d2["label"]] = depenses.get(d2["label"], 0) + d2["montant"]
    rec = sorted([{"label":k,"montant":v} for k,v in recettes.items()], key=lambda x:-x["montant"])[:5]
    dep = sorted([{"label":k,"montant":v} for k,v in depenses.items()], key=lambda x:-x["montant"])[:7]
    total_r = sum(r["montant"] for r in rec)
    total_d = sum(d["montant"] for d in dep)
    return rec, dep, total_r, total_d

def get_conseil(client, total_r, total_d, periode):
    net = total_r - total_d
    taux = round(net/total_r*100) if total_r else 0
    prompt = (
        f"Releve bancaire - periode: {periode}, recettes: {total_r}EUR, "
        f"depenses: {total_d}EUR, resultat net: {net}EUR, taux epargne: {taux}%.\n"
        "Retourne UNIQUEMENT ce JSON sans markdown:\n"
        '{"score":7,"score_detail":"phrase","actions":['
        '{"priorite":1,"titre":"titre","detail":"detail"},'
        '{"priorite":2,"titre":"titre","detail":"detail"},'
        '{"priorite":3,"titre":"titre","detail":"detail"}],'
        '"commentaire":"2 phrases analyse."}'
    )
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=800,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = msg.content[0].text.replace("```json","").replace("```","").strip()
    return json.loads(raw)

@app.route("/analyze", methods=["POST"])
def analyze():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "Aucun fichier recu"}), 400
    filename = file.filename.lower()
    file_bytes = file.read()
    if filename.endswith(".pdf"):
        text = extract_text_from_pdf(file_bytes)
    elif filename.endswith(".csv"):
        text = extract_text_from_csv(file_bytes)
    elif filename.endswith((".xlsx",".xls")):
        text = extract_text_from_excel(file_bytes)
    else:
        return jsonify({"error": "Format non supporte"}), 400
    if not text.strip():
        return jsonify({"error": "Impossible de lire le fichier"}), 400

    # Detection periode
    periode = "Periode inconnue"
    m = re.search(r'(janvier|fevrier|mars|avril|mai|juin|juillet|aout|septembre|octobre|novembre|decembre|june|july)\s*\d{4}', text.lower())
    if m:
        periode = m.group(0).capitalize()
    else:
        m2 = re.search(r'\d{1,2}/\d{4}|\d{4}', text)
        if m2:
            periode = m2.group(0)

    # Decoupage en chunks de 5000 chars
    chunk_size = 5000
    chunks = [text[i:i+chunk_size] for i in range(0, min(len(text), 60000), chunk_size)]

    client = anthropic.Anthropic()

    # Analyse de chaque chunk
    chunks_data = []
    for i, chunk in enumerate(chunks):
        try:
            data = analyse_chunk(client, chunk, i+1, len(chunks))
            chunks_data.append(data)
        except:
            continue

    if not chunks_data:
        return jsonify({"error": "Aucune transaction extraite"}), 500

    # Consolidation
    rec, dep, total_r, total_d = consolide(chunks_data, periode)

    # Score et conseils
    try:
        conseil = get_conseil(client, total_r, total_d, periode)
    except:
        conseil = {"score":5,"score_detail":"Analyse partielle","actions":[],"commentaire":"Analyse disponible."}

    result = {
        "periode": periode,
        "totalRecettes": total_r,
        "totalDepenses": total_d,
        "recettes": rec,
        "depenses": dep,
        "score": conseil.get("score", 5),
        "score_detail": conseil.get("score_detail", ""),
        "actions": conseil.get("actions", []),
        "commentaire": conseil.get("commentaire", "")
    }
    return jsonify(result)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(__import__("os").environ.get("PORT", 5001)))
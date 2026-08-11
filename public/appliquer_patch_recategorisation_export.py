# -*- coding: utf-8 -*-
"""
Ajout de 2 fonctionnalites au frontend (public/index.html) :

  1. RECATEGORISATION MANUELLE : un petit menu deroulant apparait a cote
     de chaque transaction dans le detail (deja cliquable via "detail").
     Choisir une autre categorie deplace la transaction, recalcule les
     montants des 2 categories concernees, et met a jour les graphiques
     + tableaux immediatement. Le total (RECETTES/DEPENSES en haut) ne
     bouge JAMAIS -- seule la repartition interne change, garanti par
     construction (la transaction change juste de case, son montant
     n'est jamais modifie).

  2. EXPORT EXCEL : un bouton telecharge un fichier .xlsx avec toutes les
     transactions (categorie, date, libelle, montant), y compris les
     corrections manuelles faites avant l'export. Genere entierement
     dans le navigateur (bibliotheque SheetJS, gratuite, ajoutee en une
     ligne) -- pas d'appel serveur supplementaire.

CE QUI N'EST PAS COUVERT PAR CE PREMIER JET (limitations connues) :
  - Le Top 5 depenses et les Prelevements/charges fixes ne se mettent
    PAS a jour apres une recategorisation (ce sont des listes calculees
    a part cote serveur, independantes de recettes[]/depenses[]).
  - Les traductions du bouton export et du menu ne sont ajoutees qu'en
    francais et anglais ; les autres langues afficheront un texte de
    repli en francais/anglais plutot qu'une vraie traduction.
  - Rien n'est sauvegarde : fermer l'onglet perd les corrections. Normal
    pour ce premier jet (pas de backend pour stocker les modifications).

UTILISATION : identique aux patchs backend.
    python3 appliquer_patch_recategorisation_export.py
    (a lancer depuis le dossier public/, sur index.html)
"""

import shutil
from datetime import datetime

CHEMIN_HTML = "index.html"

MARQUEUR_PREREQUIS = "function buildTable(items,colors,id,T){"

# =====================================================================
# 1) Ajout de SheetJS dans le <head>
# =====================================================================

ANCIEN_SCRIPT_HEAD = '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>'
NOUVEAU_SCRIPT_HEAD = (
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>\n'
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>'
)

# =====================================================================
# 2) Nouveau CSS, ajoute juste avant la fermeture </style>
# =====================================================================

ANCIEN_FIN_STYLE = ".evo-empty{font-size:13px;color:var(--text3);text-align:center;padding:1.5rem 0}"
NOUVEAU_FIN_STYLE = (
    ".evo-empty{font-size:13px;color:var(--text3);text-align:center;padding:1.5rem 0}\n"
    ".export-btn{background:#1d1d1f;color:#fff;border:none;padding:8px 16px;border-radius:20px;"
    "font-size:13px;font-weight:600;cursor:pointer}\n"
    ".export-btn:hover{background:#0071e3}\n"
    ".drill-row{align-items:center}\n"
    ".drill-right{display:flex;align-items:center;gap:8px;flex-shrink:0}\n"
    ".drill-amount{font-weight:500}\n"
    ".recat-select{font-size:11px;padding:3px 6px;border-radius:6px;border:1px solid var(--border);"
    "background:#fff;color:var(--text2);max-width:150px}"
)

# =====================================================================
# 3) Traductions (FR + EN, cles utilisees avec repli si absentes ailleurs)
# =====================================================================

ANCIEN_TR_FR = "chgTitle:'💳 Prélèvements & charges fixes',chgSeen:'vu sur {n} mois',chgPct:'des dépenses'},"
NOUVEAU_TR_FR = (
    "chgTitle:'💳 Prélèvements & charges fixes',chgSeen:'vu sur {n} mois',chgPct:'des dépenses',"
    "exportBtn:'⬇️ Exporter en Excel',recatNouvelle:'Nouvelle catégorie',"
    "recatPrompt:'Nom de la nouvelle catégorie :'},"
)

ANCIEN_TR_EN = "chgTitle:'💳 Recurring charges & subscriptions',chgSeen:'seen over {n} months',chgPct:'of expenses'},"
NOUVEAU_TR_EN = (
    "chgTitle:'💳 Recurring charges & subscriptions',chgSeen:'seen over {n} months',chgPct:'of expenses',"
    "exportBtn:'⬇️ Export to Excel',recatNouvelle:'New category',"
    "recatPrompt:'Name of the new category:'},"
)

# =====================================================================
# 4) dernierResultat global (a cote de cR,cD,cS,cT)
# =====================================================================

ANCIEN_LET_CHARTS = "let cR,cD,cS,cT;"
NOUVEAU_LET_CHARTS = "let cR,cD,cS,cT;\nlet dernierResultat=null;"

# =====================================================================
# 5) buildTable : ajout du parametre type, filtre des categories vides,
#    menu de recategorisation par transaction
# =====================================================================

ANCIEN_BUILDTABLE = """function buildTable(items,colors,id,T){
  const total=items.reduce((s,i)=>s+i.montant,0);
  let h='<table><thead><tr><th>'+T.cat+'</th><th style="text-align:right">'+T.mt+'</th><th style="text-align:right">%</th></tr></thead><tbody>';
  items.forEach((item,i)=>{
    const did='drill-'+id+'-'+i;
    const has=item.transactions&&item.transactions.length>0;
    h+='<tr id="row-'+id+'-'+i+'"'+(has?' class="has-detail" onclick="toggleDrill(\\''+did+'\\')"':'')+'><td><span class="dot" style="background:'+colors[i]+'"></span>'+item.label;
    if(has)h+='<span class="drill-toggle">▾ detail</span><div class="drill-content" id="'+did+'">'+item.transactions.map(t=>'<div class="drill-row"><span>'+(t.date||'')+' '+t.libelle+'</span><span>'+fmt(t.montant)+'</span></div>').join('')+'</div>';
    h+='</td><td style="text-align:right;font-weight:500">'+fmt(item.montant)+'</td><td style="text-align:right;color:#aeaeb2">'+Math.round(item.montant/total*100)+'%</td></tr>';
  });
  h+='</tbody></table>';
  document.getElementById(id).innerHTML=h;
}"""

NOUVEAU_BUILDTABLE = """function categoriesVisibles(items){
  return (items||[]).filter(it=>it.transactions&&it.transactions.length?true:Math.abs(it.montant)>0.01);
}

function echapperHtml(s){
  return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function buildTable(itemsBruts,colors,id,T,type){
  const items=categoriesVisibles(itemsBruts);
  const total=items.reduce((s,i)=>s+i.montant,0);
  const optionsCategories=items.map(it=>it.label);
  let h='<table><thead><tr><th>'+T.cat+'</th><th style="text-align:right">'+T.mt+'</th><th style="text-align:right">%</th></tr></thead><tbody>';
  items.forEach((item,i)=>{
    const did='drill-'+id+'-'+i;
    const has=item.transactions&&item.transactions.length>0;
    h+='<tr id="row-'+id+'-'+i+'"'+(has?' class="has-detail" onclick="toggleDrill(\\''+did+'\\')"':'')+'><td><span class="dot" style="background:'+colors[i]+'"></span>'+item.label;
    if(has){
      h+='<span class="drill-toggle">▾ detail</span><div class="drill-content" id="'+did+'" onclick="event.stopPropagation()">';
      h+=item.transactions.map((t,ti)=>{
        let optsHtml='';
        optionsCategories.forEach(l=>{optsHtml+='<option value="'+echapperHtml(l)+'"'+(l===item.label?' selected':'')+'>'+echapperHtml(l)+'</option>';});
        optsHtml+='<option value="__nouvelle__">'+(type?(T.recatNouvelle||'Nouvelle categorie'):'')+'…</option>';
        const select=type?('<select class="recat-select" onchange="changerCategorieTransaction(\\''+type+'\\','+i+','+ti+',this.value,\\''+id+'\\')">'+optsHtml+'</select>'):'';
        return '<div class="drill-row"><span>'+(t.date||'')+' '+echapperHtml(t.libelle)+'</span><span class="drill-right"><span class="drill-amount">'+fmt(t.montant)+'</span>'+select+'</span></div>';
      }).join('');
      h+='</div>';
    }
    h+='</td><td style="text-align:right;font-weight:500">'+fmt(item.montant)+'</td><td style="text-align:right;color:#aeaeb2">'+(total?Math.round(item.montant/total*100):0)+'%</td></tr>';
  });
  h+='</tbody></table>';
  document.getElementById(id).innerHTML=h;
}

function changerCategorieTransaction(type,catIndex,txIndex,nouveauLabelBrut,tableId){
  const T=TR[document.getElementById('langue').value]||TR.francais;
  let nouveauLabel=nouveauLabelBrut;
  if(nouveauLabel==='__nouvelle__'){
    nouveauLabel=(prompt(T.recatPrompt||'Nom de la nouvelle categorie :')||'').trim();
    if(!nouveauLabel){rafraichirApresRecategorisation();return;}
  }
  const categories=dernierResultat[type];
  const items=categoriesVisibles(categories);
  const catSource=items[catIndex];
  const transaction=catSource.transactions[txIndex];
  if(!catSource||!transaction||catSource.label===nouveauLabel){rafraichirApresRecategorisation();return;}

  catSource.transactions.splice(txIndex,1);
  catSource.montant=Math.round((catSource.montant-transaction.montant)*100)/100;

  let catCible=categories.find(c=>c.label===nouveauLabel);
  if(!catCible){
    catCible={label:nouveauLabel,montant:0,transactions:[]};
    categories.push(catCible);
  }
  catCible.transactions.push(transaction);
  catCible.montant=Math.round((catCible.montant+transaction.montant)*100)/100;

  rafraichirApresRecategorisation();
}

function rafraichirApresRecategorisation(){
  const langue=document.getElementById('langue').value;
  const T=TR[langue]||TR.francais;
  const recVisibles=categoriesVisibles(dernierResultat.recettes);
  const depVisibles=categoriesVisibles(dernierResultat.depenses);
  if(cR)cR.destroy();
  if(cD)cD.destroy();
  cR=new Chart(document.getElementById('chartR'),{type:'doughnut',data:{labels:recVisibles.map(r=>r.label),datasets:[{data:recVisibles.map(r=>r.montant),backgroundColor:CR.slice(0,recVisibles.length),borderWidth:3,borderColor:'#fff'}]},options:{responsive:true,onClick:(evt,elements)=>onChartSliceClick(elements,'tbl-r'),plugins:{legend:{position:'bottom',labels:{font:{size:11},padding:12}}},animation:{duration:400}}});
  cD=new Chart(document.getElementById('chartD'),{type:'doughnut',data:{labels:depVisibles.map(d=>d.label),datasets:[{data:depVisibles.map(d=>d.montant),backgroundColor:CD.slice(0,depVisibles.length),borderWidth:3,borderColor:'#fff'}]},options:{responsive:true,onClick:(evt,elements)=>onChartSliceClick(elements,'tbl-d'),plugins:{legend:{position:'bottom',labels:{font:{size:11},padding:12}}},animation:{duration:400}}});
  buildTable(dernierResultat.recettes,CR,'tbl-r',T,'recettes');
  buildTable(dernierResultat.depenses,CD,'tbl-d',T,'depenses');
}

function exporterExcel(){
  if(!dernierResultat||typeof XLSX==='undefined')return;
  const lignes=[];
  ['recettes','depenses'].forEach(type=>{
    const libelleType=type==='recettes'?'Recette':'Depense';
    (dernierResultat[type]||[]).forEach(cat=>{
      (cat.transactions||[]).forEach(t=>{
        lignes.push({
          Date:t.date||'',
          'Libelle':t.libelle||'',
          'Montant (EUR)':type==='depenses'?-Math.abs(t.montant):Math.abs(t.montant),
          Type:libelleType,
          'Categorie':cat.label
        });
      });
    });
  });
  if(!lignes.length)return;
  const ws=XLSX.utils.json_to_sheet(lignes);
  ws['!cols']=[{wch:12},{wch:38},{wch:14},{wch:10},{wch:28}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Transactions');
  const suffixe=(dernierResultat.periode||'export').replace(/[^0-9a-zA-Z]+/g,'_');
  XLSX.writeFile(wb,'bankanalyzer_'+suffixe+'.xlsx');
}"""

# =====================================================================
# 6) renderResult : memoriser dernierResultat, ajouter le bouton export,
#    passer le parametre 'type' aux 2 appels buildTable
# =====================================================================

ANCIEN_DEBUT_RENDERRESULT = """function renderResult(data){
  const langue=document.getElementById('langue').value;
  const T=TR[langue]||TR.francais;
  deviseActuelle=data.devise||'EUR';"""

NOUVEAU_DEBUT_RENDERRESULT = """function renderResult(data){
  dernierResultat=data;
  const langue=document.getElementById('langue').value;
  const T=TR[langue]||TR.francais;
  deviseActuelle=data.devise||'EUR';"""

ANCIEN_TABLES_GRID = """    '<div class="tables-grid anim">'+
    '<div class="card"><div class="card-title">'+T.dr+'</div><div id="tbl-r"></div></div>'+
    '<div class="card"><div class="card-title">'+T.dd+'</div><div id="tbl-d"></div></div>'+
    '</div>'+"""

NOUVEAU_TABLES_GRID = """    '<div class="anim" style="display:flex;justify-content:flex-end;margin-bottom:.75rem;">'+
    '<button class="export-btn" onclick="exporterExcel()">'+(T.exportBtn||'⬇️ Export')+'</button>'+
    '</div>'+
    '<div class="tables-grid anim">'+
    '<div class="card"><div class="card-title">'+T.dr+'</div><div id="tbl-r"></div></div>'+
    '<div class="card"><div class="card-title">'+T.dd+'</div><div id="tbl-d"></div></div>'+
    '</div>'+"""

ANCIEN_APPEL_BUILDTABLE = """  buildTable(data.recettes,CR,'tbl-r',T);
  buildTable(data.depenses,CD,'tbl-d',T);"""

NOUVEAU_APPEL_BUILDTABLE = """  buildTable(data.recettes,CR,'tbl-r',T,'recettes');
  buildTable(data.depenses,CD,'tbl-d',T,'depenses');"""


def main():
    with open(CHEMIN_HTML, "r", encoding="utf-8") as f:
        contenu = f.read()

    if MARQUEUR_PREREQUIS not in contenu:
        print("ERREUR -- ce ne semble pas etre le bon index.html (buildTable introuvable). Rien modifie.")
        return

    verifs = [
        ("script Chart.js dans le head", ANCIEN_SCRIPT_HEAD),
        ("fin du bloc CSS (.evo-empty)", ANCIEN_FIN_STYLE),
        ("traductions francaises (chgTitle)", ANCIEN_TR_FR),
        ("traductions anglaises (chgTitle)", ANCIEN_TR_EN),
        ("declaration let cR,cD,cS,cT", ANCIEN_LET_CHARTS),
        ("fonction buildTable", ANCIEN_BUILDTABLE),
        ("debut de renderResult", ANCIEN_DEBUT_RENDERRESULT),
        ("bloc tables-grid", ANCIEN_TABLES_GRID),
        ("appels a buildTable en fin de renderResult", ANCIEN_APPEL_BUILDTABLE),
    ]
    erreurs = []
    for nom, texte in verifs:
        n = contenu.count(texte)
        if n != 1:
            erreurs.append("Anchor '%s' introuvable ou different de ce qui est attendu (trouve %d fois)." % (nom, n))

    if erreurs:
        print("ERREUR -- rien n'a ete modifie dans index.html :")
        for e in erreurs:
            print(" -", e)
        print("\nColle-moi ce message et je corrige le script.")
        return

    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    sauvegarde = "index.html.avant_patch_recategorisation_%s.bak" % horodatage
    shutil.copyfile(CHEMIN_HTML, sauvegarde)

    contenu = contenu.replace(ANCIEN_SCRIPT_HEAD, NOUVEAU_SCRIPT_HEAD, 1)
    contenu = contenu.replace(ANCIEN_FIN_STYLE, NOUVEAU_FIN_STYLE, 1)
    contenu = contenu.replace(ANCIEN_TR_FR, NOUVEAU_TR_FR, 1)
    contenu = contenu.replace(ANCIEN_TR_EN, NOUVEAU_TR_EN, 1)
    contenu = contenu.replace(ANCIEN_LET_CHARTS, NOUVEAU_LET_CHARTS, 1)
    contenu = contenu.replace(ANCIEN_BUILDTABLE, NOUVEAU_BUILDTABLE, 1)
    contenu = contenu.replace(ANCIEN_DEBUT_RENDERRESULT, NOUVEAU_DEBUT_RENDERRESULT, 1)
    contenu = contenu.replace(ANCIEN_TABLES_GRID, NOUVEAU_TABLES_GRID, 1)
    contenu = contenu.replace(ANCIEN_APPEL_BUILDTABLE, NOUVEAU_APPEL_BUILDTABLE, 1)

    with open(CHEMIN_HTML, "w", encoding="utf-8") as f:
        f.write(contenu)

    print("PATCH APPLIQUE AVEC SUCCES.")
    print("Sauvegarde : %s" % sauvegarde)
    print("Verifie avec : grep -n 'exporterExcel\\|changerCategorieTransaction' index.html")


if __name__ == "__main__":
    main()

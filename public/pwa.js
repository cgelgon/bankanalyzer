/* BankAnalyzer — appli installable (PWA)
   - enregistre le service worker
   - Android / ordinateur (Chrome, Edge) : bannière avec un bouton "Installer"
   - iPhone / iPad (Safari) : bannière avec le mini-tuto en 2 étapes
   - jamais affichée dans l'appli installée, ni pendant une analyse, ni pendant
     30 jours après avoir été fermée
   - mesure anonyme (sans email) : bannière vue, installation, ouverture en mode appli */
(function () {
  'use strict';
  var API = 'https://bankanalyzer-production.up.railway.app/pwa-evenement';
  var CLE_FERME = 'bankanalyzer_pwa_ferme';
  var TXT = {
    francais: { titre: "Installe l'appli Banky 📲", texte: "Ton bilan en 1 tap depuis ton écran d'accueil. Gratuit, sans passer par un store.", installer: 'Installer', ios1: 'Appuie sur {i} Partager', ios2: "Puis « Sur l'écran d'accueil » ➕", fermer: 'Fermer' },
    english: { titre: 'Get the Banky app 📲', texte: 'Your report in 1 tap from your home screen. Free, no app store needed.', installer: 'Install', ios1: 'Tap {i} Share', ios2: 'Then “Add to Home Screen” ➕', fermer: 'Close' },
    espanol: { titre: 'Instala la app de Banky 📲', texte: 'Tu balance en 1 toque desde tu pantalla de inicio. Gratis, sin pasar por una tienda.', installer: 'Instalar', ios1: 'Toca {i} Compartir', ios2: 'Luego «Añadir a pantalla de inicio» ➕', fermer: 'Cerrar' },
    deutsch: { titre: 'Installiere die Banky-App 📲', texte: 'Deine Übersicht mit 1 Tipp vom Startbildschirm. Kostenlos, ohne App-Store.', installer: 'Installieren', ios1: 'Tippe auf {i} Teilen', ios2: 'Dann „Zum Home-Bildschirm“ ➕', fermer: 'Schließen' },
    italiano: { titre: "Installa l'app di Banky 📲", texte: 'Il tuo bilancio in 1 tocco dalla schermata Home. Gratis, senza passare da uno store.', installer: 'Installa', ios1: 'Tocca {i} Condividi', ios2: 'Poi «Aggiungi alla schermata Home» ➕', fermer: 'Chiudi' },
    portugues: { titre: 'Instale o app do Banky 📲', texte: 'Seu balanço em 1 toque na tela inicial. Grátis, sem passar por loja de apps.', installer: 'Instalar', ios1: 'Toque em {i} Compartilhar', ios2: 'Depois “Adicionar à Tela de Início” ➕', fermer: 'Fechar' },
    chinese: { titre: '安装Banky应用 📲', texte: '在主屏幕一键打开你的报告。免费，无需应用商店。', installer: '安装', ios1: '点击 {i} 分享', ios2: '然后选择“添加到主屏幕” ➕', fermer: '关闭' },
    arabic: { titre: 'ثبّت تطبيق Banky 📲', texte: 'تقريرك بلمسة واحدة من شاشتك الرئيسية. مجاني ودون المرور بمتجر تطبيقات.', installer: 'تثبيت', ios1: 'اضغط على {i} مشاركة', ios2: 'ثم «إضافة إلى الشاشة الرئيسية» ➕', fermer: 'إغلاق' }
  };
  var ICONE_PARTAGE = '<svg width="16" height="20" viewBox="0 0 16 20" aria-hidden="true" style="vertical-align:-4px"><path d="M8 1v12M4 5l4-4 4 4" stroke="#0071e3" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"/><path d="M5 8H2.5A1.5 1.5 0 0 0 1 9.5v8A1.5 1.5 0 0 0 2.5 19h11a1.5 1.5 0 0 0 1.5-1.5v-8A1.5 1.5 0 0 0 13.5 8H11" stroke="#0071e3" stroke-width="1.8" fill="none" stroke-linecap="round"/></svg>';

  function langue() { try { return localStorage.getItem('bankanalyzer_langue') || 'francais'; } catch (e) { return 'francais'; } }
  function T() { return TXT[langue()] || TXT.francais; }
  function echapper(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  var ua = navigator.userAgent || '';
  var estIOS = /iPhone|iPad|iPod/.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  var estAndroid = /Android/i.test(ua);
  function plateforme() { return estIOS ? 'ios' : (estAndroid ? 'android' : 'ordinateur'); }
  function enModeAppli() { return (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || navigator.standalone === true; }

  function suivre(type) {
    try {
      fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify({ type: type, plateforme: plateforme() }), keepalive: true }).catch(function () {});
    } catch (e) {}
  }

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () { navigator.serviceWorker.register('/sw.js').catch(function () {}); });
  }

  if (enModeAppli()) {
    try { if (!sessionStorage.getItem('bk_appli_ouverte')) { sessionStorage.setItem('bk_appli_ouverte', '1'); suivre('ouverture_appli'); } } catch (e) {}
    return;
  }

  var invitation = null;
  var banniere = null;

  function fermeeRecemment() { try { return Date.now() - (+localStorage.getItem(CLE_FERME) || 0) < 30 * 864e5; } catch (e) { return false; } }
  function analyseEnCours() {
    var r = document.getElementById('result'), l = document.getElementById('loading');
    return !!((r && r.style.display === 'block') || (l && l.style.display === 'block'));
  }
  function retirer() { if (banniere) { banniere.remove(); banniere = null; } }
  function fermer() { retirer(); try { localStorage.setItem(CLE_FERME, String(Date.now())); } catch (e) {} }

  function ajouterStyle() {
    if (document.getElementById('bk-pwa-style')) return;
    var s = document.createElement('style');
    s.id = 'bk-pwa-style';
    s.textContent =
      '.bk-pwa{position:fixed;left:12px;right:12px;bottom:calc(12px + env(safe-area-inset-bottom,0px));z-index:2147483000;max-width:460px;margin:0 auto;background:#fff;color:#1d1d1f;border-radius:18px;box-shadow:0 12px 40px rgba(0,0,0,.18);padding:14px 14px 14px 12px;display:flex;gap:12px;align-items:flex-start;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;animation:bkPwaIn .35s ease-out}' +
      '@keyframes bkPwaIn{from{transform:translateY(20px);opacity:0}to{transform:none;opacity:1}}' +
      '@media(prefers-reduced-motion:reduce){.bk-pwa{animation:none}}' +
      '.bk-pwa img{width:48px;height:48px;border-radius:12px;flex-shrink:0}' +
      '.bk-pwa-txt{flex:1;min-width:0}' +
      '.bk-pwa-txt b{display:block;font-size:15px;margin-bottom:2px}' +
      '.bk-pwa-txt>span{display:block}.bk-pwa-txt>span,.bk-pwa-txt li{font-size:13px;color:#6e6e73;line-height:1.4}' +
      '.bk-pwa-txt ol{margin:4px 0 0;padding-left:18px}' +
      '.bk-pwa-btn{display:block;margin-top:9px;font:inherit;font-size:14px;font-weight:700;color:#fff;background:linear-gradient(135deg,#0071e3,#00c7ff);border:none;border-radius:100px;padding:9px 18px;cursor:pointer}' +
      '.bk-pwa-txt svg{display:inline-block;vertical-align:-4px;max-width:none}' +
      '.bk-pwa-x{font:inherit;font-size:16px;line-height:1;color:#9a9a9f;background:none;border:none;padding:6px;cursor:pointer;flex-shrink:0;align-self:flex-start}' +
      '.bk-pwa-x:hover,.bk-pwa-x:focus-visible{color:#1d1d1f}';
    document.head.appendChild(s);
  }

  function afficher(mode) {
    if (banniere || fermeeRecemment() || analyseEnCours()) return;
    ajouterStyle();
    var t = T();
    banniere = document.createElement('div');
    banniere.className = 'bk-pwa';
    banniere.setAttribute('role', 'dialog');
    banniere.setAttribute('aria-label', t.titre);
    if (langue() === 'arabic') banniere.dir = 'rtl';
    var corps = mode === 'ios'
      ? '<ol><li>' + echapper(t.ios1).replace('{i}', ICONE_PARTAGE) + '</li><li>' + echapper(t.ios2) + '</li></ol>'
      : '<span>' + echapper(t.texte) + '</span>';
    banniere.innerHTML =
      '<img src="/app/icon-192.png" alt="">' +
      '<div class="bk-pwa-txt"><b>' + echapper(t.titre) + '</b>' + corps +
      (mode === 'ios' ? '' : '<button type="button" class="bk-pwa-btn">' + echapper(t.installer) + '</button>') + '</div>' +
      '<button type="button" class="bk-pwa-x" aria-label="' + echapper(t.fermer) + '">✕</button>';
    banniere.querySelector('.bk-pwa-x').addEventListener('click', fermer);
    var btn = banniere.querySelector('.bk-pwa-btn');
    if (btn) {
      btn.addEventListener('click', function () {
        if (!invitation) { retirer(); return; }
        invitation.prompt();
        invitation.userChoice.then(function (choix) {
          suivre(choix && choix.outcome === 'accepted' ? 'acceptee' : 'refusee');
          if (!choix || choix.outcome !== 'accepted') fermer(); else retirer();
          invitation = null;
        });
      });
    }
    document.body.appendChild(banniere);
    suivre('banniere_vue');
  }

  // La bannière s'efface d'elle-même si une analyse démarre
  setInterval(function () { if (banniere && analyseEnCours()) retirer(); }, 1000);

  window.addEventListener('beforeinstallprompt', function (e) {
    e.preventDefault();
    invitation = e;
    setTimeout(function () { afficher('bouton'); }, 3000);
  });
  window.addEventListener('appinstalled', function () { suivre('installee'); retirer(); });

  if (estIOS && /Safari/.test(ua) && !/CriOS|FxiOS|EdgiOS|OPiOS/.test(ua)) {
    setTimeout(function () { afficher('ios'); }, 4000);
  }
})();

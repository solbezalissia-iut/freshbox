/**
 * FreshBox — enregistrement des saisies et rappels par e-mail
 * À coller dans : Google Sheets > Extensions > Apps Script
 */

const FUSEAU = "Europe/Paris";
const COLONNES = ["saisi_le", "boite", "produit", "date_peremption", "date_estimee", "email",
                  "rappel_j2", "rappel_j1", "rappel_j0", "envoye_j2", "envoye_j1", "envoye_j0", "statut"];

// Reçoit les données envoyées par la page FreshBox
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Clic "Mangé" ou "Jeté"
  if (d.evenement) {
    const res = feuille_(ss, "Resultats", ["saisi_le", "evenement", "boite", "produit", "email"]);
    res.appendRow([d.saisi_le, d.evenement, d.boite, d.produit, d.contact]);
    // On arrête les rappels pour cet aliment
    const sh = feuille_(ss, "Saisies", COLONNES);
    const v = sh.getDataRange().getDisplayValues();
    for (let i = v.length - 1; i >= 1; i--) {
      if (v[i][1] == d.boite && v[i][2] == d.produit && v[i][12] === "en cours") {
        sh.getRange(i + 1, 13).setValue(d.evenement === "mange" ? "mangé" : "jeté");
        break;
      }
    }
    return ContentService.createTextOutput("ok");
  }

  // Nouvel aliment
  const sh = feuille_(ss, "Saisies", COLONNES);
  sh.appendRow([
    d.saisi_le, d.boite, d.produit, "'" + d.date_peremption, d.date_estimee ? "oui" : "non", d.contact,
    d.rappel_j2 ? "oui" : "non", d.rappel_j1 ? "oui" : "non", d.rappel_j0 ? "oui" : "non",
    "", "", "", "en cours"
  ]);

  // E-mail de confirmation immédiat
  MailApp.sendEmail({
    to: d.contact,
    name: "FreshBox",
    subject: "🟢 FreshBox : " + d.produit + " (boîte n°" + d.boite + ") est enregistré",
    htmlBody: "<p>C'est noté !</p><p><b>" + d.produit + "</b> (boîte n°" + d.boite + ") est à consommer avant le <b>"
      + joliDate_(d.date_peremption) + "</b>" + (d.date_estimee ? " (date estimée)" : "") + ".</p>"
      + "<p>Tu recevras un e-mail de rappel quand la date approchera.</p><p>— FreshBox</p>"
  });

  return ContentService.createTextOutput("ok");
}

// Envoie les rappels du jour (lancé automatiquement chaque matin)
function envoyerRappels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = feuille_(ss, "Saisies", COLONNES);
  const v = sh.getDataRange().getDisplayValues();
  const aujourdhui = Utilities.formatDate(new Date(), FUSEAU, "yyyy-MM-dd");

  for (let i = 1; i < v.length; i++) {
    const [ , boite, produit, date, , email, r2, r1, r0, e2, e1, e0, statut] = v[i];
    if (statut !== "en cours" || !email || !date) continue;
    const j = joursEntre_(aujourdhui, date);

    let envoi = null;
    if (j === 2 && r2 === "oui" && !e2) envoi = { col: 10, emoji: "🟠", texte: "est à consommer bientôt (dans 2 jours)" };
    if (j === 1 && r1 === "oui" && !e1) envoi = { col: 11, emoji: "🟠", texte: "est à consommer demain" };
    if (j === 0 && r0 === "oui" && !e0) envoi = { col: 12, emoji: "🔴", texte: "est à consommer aujourd'hui" };
    if (!envoi) continue;

    MailApp.sendEmail({
      to: email,
      name: "FreshBox",
      subject: envoi.emoji + " FreshBox : " + produit + " (boîte n°" + boite + ") " + envoi.texte,
      htmlBody: "<p>" + envoi.emoji + " <b>" + produit + "</b> (boîte n°" + boite + ") " + envoi.texte
        + " : date limite le " + joliDate_(date) + ".</p>"
        + "<p>Une idée : pense à l'utiliser dans ton prochain repas (omelette, poêlée, soupe, gratin…).</p><p>— FreshBox</p>"
    });
    sh.getRange(i + 1, envoi.col).setValue(Utilities.formatDate(new Date(), FUSEAU, "dd/MM HH:mm"));
  }
}

// À lancer UNE fois : programme l'envoi des rappels tous les jours vers 9 h
function installer() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("envoyerRappels").timeBased().everyDays(1).atHour(9).inTimezone(FUSEAU).create();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  feuille_(ss, "Saisies", COLONNES);
  feuille_(ss, "Resultats", ["saisi_le", "evenement", "boite", "produit", "email"]);
}

// --- Outils ---
function feuille_(ss, nom, entetes) {
  let sh = ss.getSheetByName(nom);
  if (!sh) { sh = ss.insertSheet(nom); sh.appendRow(entetes); sh.setFrozenRows(1); }
  return sh;
}
function joursEntre_(a, b) {
  const pa = a.split("-").map(Number), pb = b.split("-").map(Number);
  return Math.round((Date.UTC(pb[0], pb[1] - 1, pb[2]) - Date.UTC(pa[0], pa[1] - 1, pa[2])) / 86400000);
}
function joliDate_(s) {
  const p = s.split("-").map(Number);
  return Utilities.formatDate(new Date(p[0], p[1] - 1, p[2], 12), FUSEAU, "dd/MM/yyyy");
}

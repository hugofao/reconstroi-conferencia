/** @OnlyCurrentDoc */
// Recebe as inscrições do site e escreve-as na folha de cálculo (uma linha por inscrição).
const ABA = "Inscrições";
const CABECALHO = ["Ref", "Data/hora (Lisboa)", "Nome", "Email", "Telemóvel", "Faixa etária", "Igreja / comunidade", "Domingo (culto)", "Almoço", "Observações"];

function limpar(v) {
  v = String(v == null ? "" : v).trim().slice(0, 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // evita fórmulas maliciosas
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.website || !String(d.name || "").trim()) return ContentService.createTextOutput("ignorado");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(ABA) || ss.insertSheet(ABA);
    if (sh.getLastRow() === 0) {
      sh.appendRow(CABECALHO);
      sh.getRange(1, 1, 1, CABECALHO.length).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
    const ref = String(d.id || "").slice(0, 12).toUpperCase();
    if (ref && sh.getLastRow() > 1) {
      const refs = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat();
      if (refs.indexOf(ref) !== -1) return ContentService.createTextOutput("repetido");
    }
    const quando = Utilities.formatDate(new Date(), "Europe/Lisbon", "yyyy-MM-dd HH:mm");
    sh.appendRow([ref, quando, limpar(d.name), limpar(d.email), "'" + limpar(d.phone), limpar(d.age), limpar(d.church), limpar(d.sunday), limpar(d.lunch), limpar(d.notes)]);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return ContentService.createTextOutput("ok"); }

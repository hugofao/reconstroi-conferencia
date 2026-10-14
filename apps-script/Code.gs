/** @OnlyCurrentDoc */
// Recebe as inscrições do site e escreve-as na folha de cálculo (uma linha por inscrição).
const ABA = "Inscrições";
const CABECALHO = ["Ref", "Data/hora (Lisboa)", "Nome", "Email", "Telemóvel", "Faixa etária", "Igreja / comunidade", "Domingo (culto)", "Almoço", "Observações"];

function limpar(v) {
  v = String(v == null ? "" : v).trim().slice(0, 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // evita fórmulas maliciosas
}

// Devolve a folha a que o script está ligado (criado via Extensões → Apps Script, dentro da folha).
function folha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Este script não está ligado a nenhuma folha. Abre a folha Google e usa Extensões → Apps Script.");
  return ss;
}

// Corre isto no editor (menu de funções → "ondeEscreve" → Executar) para ver o link da folha.
function ondeEscreve() {
  const ss = folha();
  Logger.log("Folha: " + ss.getName() + "\nLink: " + ss.getUrl() + "\nSeparador: " + ABA);
}

// Corre isto no editor (→ "testar" → Executar): pede autorização e escreve uma linha de teste.
function testar() {
  const r = doPost({ postData: { contents: JSON.stringify({ id: "TESTE" + Date.now(), name: "Teste do editor", email: "teste@exemplo.pt", phone: "", age: "40", church: "", notes: "", sunday: "Não", lunch: "Não" }) } });
  Logger.log("Resultado: " + r.getContent() + "\nLink da folha: " + folha().getUrl());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.website || !String(d.name || "").trim()) return ContentService.createTextOutput("ignorado");
    const ss = folha();
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

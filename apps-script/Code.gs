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
    try { enviarEmail(ref, d); } catch (err) { console.error("Email não enviado: " + err); } // nunca impede o registo
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function esc(v) {
  return String(v == null ? "" : v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/\n/g, "<br>");
}

// Email de confirmação com o resumo da inscrição e uma mensagem de agradecimento.
function enviarEmail(ref, d) {
  const email = String(d.email || "").trim();
  if (!/^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]+$/.test(email) || email.length > 120) return;
  const nome = String(d.name || "").trim().slice(0, 100);
  const primeiro = nome.split(/\s+/)[0] || nome;
  const sim = (v) => String(v) === "Sim";
  const linhas = [
    ["Referência", ref],
    ["Nome", nome],
    ["Email", email],
    ["Telemóvel", String(d.phone || "").trim() || "—"],
    ["Idade", String(d.age || "").trim()],
    ["Igreja / comunidade", String(d.church || "").trim() || "—"],
    ["Almoço de sábado (5 €)", sim(d.lunch) ? "Sim, quero reservar" : "Não"],
    ["Celebração e Encerramento Domingo", sim(d.sunday) ? "Sim, quero participar" : "Não"],
    ["Observações", String(d.notes || "").trim() || "—"],
  ].map(([k, v]) => [k, v.slice(0, 500)]);

  const agradecimento =
    "Querido " + primeiro + ",\n\n" +
    "Obrigado por te inscreveres na conferência [re]constrói. A tua presença é uma bênção para nós. " +
    "Acreditamos que, quando os homens se reúnem para buscar a Deus, se fortalecem os lares, as famílias e as gerações que vêm depois. " +
    "\"Se o Senhor não edificar a casa, em vão trabalham os que a edificam.\" (Salmo 127:1)\n\n" +
    "Estamos a preparar tudo com muito carinho e vamos orar por ti até ao dia. Vem de coração aberto: vamos [re]construir juntos.";

  const evento = "28 e 29 de Novembro de 2026 · AD Benfica\n" +
    "Sábado de manhã: 09:30 às 13:00 · Sábado de tarde: 14:30 às 17:00 · Domingo: 10:00 às 11:30";

  const texto = agradecimento + "\n\n— O TEU REGISTO —\n" + linhas.map(([k, v]) => k + ": " + v).join("\n") +
    "\n\n— QUANDO E ONDE —\n" + evento +
    "\n\nGuarda este email: a referência " + ref + " identifica a tua inscrição.\n\nAté 28 de Novembro!\nEquipa [re]constrói · AD Benfica";

  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;margin:0 auto;color:#222;line-height:1.55">' +
    '<div style="background:#1a1410;color:#fff;padding:22px 24px;border-radius:10px 10px 0 0"><div style="font-size:24px;font-weight:800">[re]constrói</div>' +
    '<div style="color:#f26b1d;font-size:14px;margin-top:2px">Conferência de Homens · 28 e 29 Nov 2026 · AD Benfica</div></div>' +
    '<div style="border:1px solid #e5ddd5;border-top:0;padding:24px;border-radius:0 0 10px 10px">' +
    '<p style="font-size:18px;margin:0 0 12px"><b>Querido ' + esc(primeiro) + ',</b></p>' +
    '<p>Obrigado por te inscreveres na conferência <b>[re]constrói</b>. A tua presença é uma bênção para nós. ' +
    'Acreditamos que, quando os homens se reúnem para buscar a Deus, se fortalecem os lares, as famílias e as gerações que vêm depois.</p>' +
    '<p style="border-left:4px solid #f26b1d;padding-left:12px;color:#555"><i>"Se o Senhor não edificar a casa, em vão trabalham os que a edificam."</i> — Salmo 127:1</p>' +
    '<p>Estamos a preparar tudo com muito carinho e vamos orar por ti até ao dia. Vem de coração aberto: vamos [re]construir juntos.</p>' +
    '<h3 style="margin:24px 0 8px;color:#f26b1d">O teu registo</h3>' +
    '<table style="border-collapse:collapse;width:100%;font-size:14px">' +
    linhas.map(([k, v]) => '<tr><td style="padding:6px 10px;border-bottom:1px solid #eee;color:#777;width:42%">' + esc(k) + '</td><td style="padding:6px 10px;border-bottom:1px solid #eee"><b>' + esc(v) + '</b></td></tr>').join("") +
    '</table>' +
    '<h3 style="margin:24px 0 8px;color:#f26b1d">Quando e onde</h3>' +
    '<p style="margin:0">' + esc(evento).replace(/ · Sábado de tarde/, "<br>Sábado de tarde") + '</p>' +
    '<p style="margin-top:22px;font-size:13px;color:#777">Guarda este email: a referência <b>' + esc(ref) + '</b> identifica a tua inscrição.</p>' +
    '<p style="margin-top:18px"><b>Até 28 de Novembro!</b><br>Equipa [re]constrói · AD Benfica</p>' +
    '</div></div>';

  MailApp.sendEmail({
    to: email,
    subject: "[re]constrói · A tua inscrição está confirmada (Ref. " + ref + ")",
    body: texto,
    htmlBody: html,
    name: "[re]constrói · AD Benfica",
  });
}

// Corre no editor ("testarEmail" → Executar): envia um email de exemplo para a tua própria conta.
function testarEmail() {
  const eu = Session.getEffectiveUser().getEmail();
  enviarEmail("TESTE0001", { name: "Hugo Oliveira", email: eu, phone: "", age: "40", church: "AD Benfica", notes: "Teste de email", sunday: "Sim", lunch: "Sim" });
  Logger.log("Email de teste enviado para " + eu);
}

function doGet() { return ContentService.createTextOutput("ok"); }

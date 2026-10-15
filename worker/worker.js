// Cloudflare Worker: recebe a inscrição (já cifrada no browser), guarda-a no relay (ntfy)
// e "acorda" o GitHub Action para a processar de imediato.
// Variável necessária: GH_TOKEN (secret)
// Rota /email (envio de emails por SMTP, chamada só pelo Apps Script). Secrets: EMAIL_SECRET,
// SMTP_HOST, SMTP_PORT (587 = STARTTLS, 465 = TLS), SMTP_USER, SMTP_PASS, SMTP_FROM (endereço remetente)
import { WorkerMailer } from "worker-mailer";
const REPO = "hugofao/reconstroi-conferencia";
// endereços do site autorizados a enviar inscrições (acrescentar aqui o novo endereço do site)
const ORIGINS = ["https://hugofao.github.io"];
const NTFY_TOPIC = "reconstroi-b9eaf7b17932dcd721ba53a6"; // já é público no site

const corsFor = (req) => ({ "Access-Control-Allow-Origin": ORIGINS.includes(req.headers.get("Origin")) ? req.headers.get("Origin") : ORIGINS[0], "Vary": "Origin", "Access-Control-Allow-Methods": "POST, OPTIONS", "Access-Control-Allow-Headers": "Content-Type" });

const EMAIL_RE = /^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]+$/;

async function enviarEmail(req, env) {
  if (!env.EMAIL_SECRET || req.headers.get("x-secret") !== env.EMAIL_SECRET) return new Response("no", { status: 403 });
  for (const k of ["SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASS", "SMTP_FROM"]) if (!env[k]) return new Response("falta " + k, { status: 500 });
  let m;
  try { m = await req.json(); } catch (_) { return new Response("json", { status: 400 }); }
  const to = String(m.to || "").trim(), nome = String(m.name || "").slice(0, 100).replace(/[\r\n"<>]/g, "");
  if (!EMAIL_RE.test(to) || to.length > 120 || !m.subject || String(m.subject).length > 200 || !m.text || String(m.html || "").length > 60000 || String(m.text).length > 20000)
    return new Response("bad", { status: 400 });
  const port = Number(env.SMTP_PORT);
  try {
    await WorkerMailer.send(
      { host: env.SMTP_HOST, port, secure: port === 465, startTls: port !== 465, credentials: { username: env.SMTP_USER, password: env.SMTP_PASS }, authType: ["plain", "login"] },
      { from: { name: "[re]constrói · AD Benfica", email: env.SMTP_FROM }, to: { name: nome || to, email: to }, subject: String(m.subject), text: String(m.text), html: m.html ? String(m.html) : undefined },
    );
    return new Response("ok", { status: 200 });
  } catch (e) {
    return new Response("smtp: " + String(e && e.message || e).slice(0, 300), { status: 502 });
  }
}

export default {
  async fetch(req, env, ctx) {
    if (req.method === "POST" && new URL(req.url).pathname === "/email") return enviarEmail(req, env);
    const cors = corsFor(req);
    if (req.method === "GET") return new Response(env.GH_TOKEN ? "token: sim" : "token: NAO", { status: 200 });
    if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });
    if (req.method !== "POST" || !ORIGINS.includes(req.headers.get("Origin"))) return new Response("no", { status: 403, headers: cors });
    const body = (await req.text()).trim();
    if (body.length < 100 || body.length > 6000 || !/^[A-Za-z0-9+/=]+$/.test(body)) return new Response("bad", { status: 400, headers: cors });

    const r = await fetch("https://ntfy.sh/" + NTFY_TOPIC, { method: "POST", body });
    if (!r.ok) return new Response("relay", { status: 502, headers: cors });

    // acordar o processamento em segundo plano, sem atrasar a resposta (se falhar, o agendamento apanha-a)
    ctx.waitUntil(fetch(`https://api.github.com/repos/${REPO}/dispatches`, {
      method: "POST",
      headers: { Authorization: "Bearer " + env.GH_TOKEN, Accept: "application/vnd.github+json", "User-Agent": "reconstroi-worker", "Content-Type": "application/json" },
      body: JSON.stringify({ event_type: "inscricao" }),
    }).catch(() => {}));
    return new Response("ok", { status: 200, headers: cors });
  },
};

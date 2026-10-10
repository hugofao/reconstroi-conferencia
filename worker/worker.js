// Cloudflare Worker: recebe a inscrição (já cifrada no browser), guarda-a no relay (ntfy)
// e "acorda" o GitHub Action para a processar de imediato.
// Variável necessária: GH_TOKEN (secret)
const REPO = "hugofao/reconstroi-conferencia";
const ORIGIN = "https://hugofao.github.io";
const NTFY_TOPIC = "reconstroi-b9eaf7b17932dcd721ba53a6"; // já é público no site

const cors = { "Access-Control-Allow-Origin": ORIGIN, "Access-Control-Allow-Methods": "POST, OPTIONS", "Access-Control-Allow-Headers": "Content-Type" };

export default {
  async fetch(req, env) {
    if (req.method === "GET") return new Response(env.GH_TOKEN ? "token: sim" : "token: NAO", { status: 200 });
    if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });
    if (req.method !== "POST" || req.headers.get("Origin") !== ORIGIN) return new Response("no", { status: 403, headers: cors });
    const body = (await req.text()).trim();
    if (body.length < 100 || body.length > 6000 || !/^[A-Za-z0-9+/=]+$/.test(body)) return new Response("bad", { status: 400, headers: cors });

    const r = await fetch("https://ntfy.sh/" + NTFY_TOPIC, { method: "POST", body });
    if (!r.ok) return new Response("relay", { status: 502, headers: cors });

    // acordar o processamento (se falhar, o agendamento apanha-a na mesma)
    const g = await fetch(`https://api.github.com/repos/${REPO}/dispatches`, {
      method: "POST",
      headers: { Authorization: "Bearer " + env.GH_TOKEN, Accept: "application/vnd.github+json", "User-Agent": "reconstroi-worker", "Content-Type": "application/json" },
      body: JSON.stringify({ event_type: "inscricao" }),
    });
    return new Response("ok:" + g.status, { status: 200, headers: cors });
  },
};

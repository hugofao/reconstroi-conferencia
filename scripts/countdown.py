"""Email countdown aos inscritos: lê o Excel cifrado e envia pelo Worker (/countdown), com pedidos assinados.
Modo "simular": só conta destinatários. Modo "enviar": envia e regista quem já recebeu (HMAC, sem expor emails).
Os logs são públicos: nunca escrever emails ou nomes."""
import base64, datetime, hashlib, hmac, io, json, os, re, sys, time, urllib.request, urllib.error
from zoneinfo import ZoneInfo
import msoffcrypto, openpyxl
from cryptography.hazmat.primitives import serialization, hashes
from cryptography.hazmat.primitives.asymmetric import padding

WORKER = "https://reconstroi-conferencia.hugofao.workers.dev/countdown"
DATA = datetime.date(2026, 11, 28)
EMAIL_RE = re.compile(r'^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]+$')
CAMPANHAS = {
    "countdown": {"html": "countdown/email.html", "txt": "countdown/email.txt", "enviados": "countdown/enviados.txt", "min_dias": 1,
                  "assunto": lambda dias, faltam: f"[RE] CONSTRÓI - {faltam} - A obra está quase a começar"},
    "dia": {"html": "countdown/dia.html", "txt": "countdown/dia.txt", "enviados": "countdown/enviados-dia.txt", "min_dias": 0,
            "assunto": lambda dias, faltam: "[RE] CONSTRÓI - Já só faltam algumas horas!"},
}

def esc(v):
    return v.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

def main():
    pedido = json.load(open("countdown/pedido.json"))
    modo = pedido.get("modo")
    camp = CAMPANHAS[pedido.get("campanha", "countdown")]
    ENVIADOS = camp["enviados"]
    excluir_refs = {str(r).upper() for r in pedido.get("excluir_refs", [])}
    pw = os.environ["XLSX_PASSWORD"]
    priv = serialization.load_pem_private_key(os.environ["RSA_PRIVATE_KEY"].encode(), None)
    f = msoffcrypto.OfficeFile(open("inscricoes.xlsx", "rb")); f.load_key(password=pw)
    b = io.BytesIO(); f.decrypt(b); ws = openpyxl.load_workbook(b).active
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0]]
    excluir_mails = {str(r[3] or "").strip().lower() for r in rows if str(r[0]).upper() in excluir_refs}
    tag = lambda m: hmac.new(pw.encode(), m.encode(), hashlib.sha256).hexdigest()[:24]
    ja = set(open(ENVIADOS).read().split()) if os.path.exists(ENVIADOS) else set()
    dest, vistos, invalidos, excluidos, repetidos = [], set(), 0, 0, 0
    for r in rows:
        mail = str(r[3] or "").strip().lstrip("'")
        low = mail.lower()
        if not EMAIL_RE.match(mail) or len(mail) > 120: invalidos += 1; continue
        if low in excluir_mails: excluidos += 1; continue
        if low in vistos: repetidos += 1; continue
        vistos.add(low); dest.append((mail, str(r[2] or "").strip().lstrip("'")))
    pendentes = [d for d in dest if tag(d[0].lower()) not in ja]
    hoje = datetime.datetime.now(ZoneInfo("Europe/Lisbon")).date()
    dias = (DATA - hoje).days
    faltam = "Falta 1 dia" if dias == 1 else f"Faltam {dias} dias"
    res = [f"Data (Lisboa): {hoje} | {faltam}", f"Modo: {modo} | Campanha: {pedido.get('campanha', 'countdown')}",
           f"Linhas no Excel: {len(rows)}", f"Emails únicos válidos: {len(dest)}",
           f"Excluídos (refs {', '.join(sorted(excluir_refs)) or '-'}): {excluidos}", f"Repetidos: {repetidos}",
           f"Sem email válido: {invalidos}", f"Já tinham recebido antes: {len(dest) - len(pendentes)}", f"Por enviar: {len(pendentes)}"]
    if modo == "almoco":
        from collections import Counter
        val = lambda r: str(r[9] if len(r) > 9 and r[9] is not None else "").strip()
        valores = Counter(val(r) or "(vazio)" for r in rows)
        sim = [r for r in rows if val(r).lower() in ("sim", "s", "yes", "true")]
        sim_unicos = {str(r[3] or "").strip().lower() for r in sim}
        sim_sem_teste = sim_unicos - excluir_mails
        res = [f"Data (Lisboa): {hoje}", "Modo: almoco", f"Linhas no Excel: {len(rows)}",
               "Valores na coluna Almoço: " + ", ".join(f"{k}={v}" for k, v in valores.most_common()),
               f"Linhas com almoço Sim: {len(sim)}", f"Pessoas únicas (por email) com Sim: {len(sim_unicos)}",
               f"Linhas com Sim sem o email de teste: {sum(1 for r in sim if str(r[3] or '').strip().lower() not in excluir_mails)}",
               f"Pessoas únicas (nome + email) com Sim, sem o email de teste: {len({(' '.join(str(r[2] or '').lower().split()), str(r[3] or '').strip().lower()) for r in sim if str(r[3] or '').strip().lower() not in excluir_mails})}",
               f"Únicas com Sim sem o email de teste ({', '.join(sorted(excluir_refs)) or '-'}): {len(sim_sem_teste)}"]
    if modo == "teste":
        alvo = str(pedido.get("teste_sha256", "")).lower()
        teste = [(str(r[3] or "").strip().lstrip("'"), str(r[2] or "").strip().lstrip("'")) for r in rows
                 if hashlib.sha256(str(r[3] or "").strip().lstrip("'").lower().encode()).hexdigest() == alvo][:1]
        res.append(f"Teste: endereço de teste {'encontrado' if teste else 'NÃO encontrado'} no Excel")
        pendentes = teste
    if modo in ("enviar", "teste") and pendentes:
        if dias < camp["min_dias"]: sys.exit("A conferência já começou.")
        html0, txt0 = open(camp["html"]).read(), open(camp["txt"]).read()
        ok, falhas = 0, []
        for i, (mail, nome) in enumerate(pendentes, 1):
            primeiro = (nome.split() or [""])[0][:40]
            fill = lambda t, h: t.replace("{{FALTAM}}", faltam).replace("{{NOME}}" if primeiro else " {{NOME}}", (esc(primeiro) if h else primeiro) if primeiro else "")
            body = json.dumps({"to": mail, "name": nome[:100], "subject": camp["assunto"](dias, faltam),
                               "html": fill(html0, True), "text": fill(txt0, False), "ts": int(time.time() * 1000)}).encode()
            sig = base64.b64encode(priv.sign(body, padding.PSS(mgf=padding.MGF1(hashes.SHA256()), salt_length=32), hashes.SHA256())).decode()
            req = urllib.request.Request(WORKER, data=body, method="POST", headers={"Content-Type": "application/json", "x-sig": sig, "User-Agent": "reconstroi-countdown"})
            try:
                with urllib.request.urlopen(req, timeout=60) as r: code, txt = r.status, ""
            except urllib.error.HTTPError as e: code, txt = e.code, e.read().decode()[:200]
            except Exception as e: code, txt = 0, type(e).__name__
            if code == 200:
                ok += 1
                if modo == "enviar":
                    ja.add(tag(mail.lower())); open(ENVIADOS, "w").write("\n".join(sorted(ja)) + "\n")
            else:
                falhas.append(f"#{i}: HTTP {code} {txt}")
                if code == 403 and i == 1: res.append("Worker recusou a assinatura (ainda não atualizado?). Parado."); break
            time.sleep(2)
        res += [f"Enviados agora: {ok}", f"Falhas: {len(falhas)}"] + falhas
    out = "\n".join(res); print(out)
    open("countdown/resultado.txt", "w").write(out + "\n")

if __name__ == "__main__":
    main()

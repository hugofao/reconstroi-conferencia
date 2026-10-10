"""Lê inscrições cifradas do relay, decifra e acrescenta ao Excel cifrado com password."""
import base64, io, json, os, sys, urllib.request, datetime
import msoffcrypto, openpyxl
from cryptography.hazmat.primitives import serialization, hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

XLSX = "inscricoes.xlsx"
HEAD = ["Ref", "Data/hora (UTC)", "Nome", "Email", "Telemóvel", "Idade", "Igreja / comunidade", "Domingo (culto)", "Observações"]
KEYS = ["name", "email", "phone", "age", "church", "sunday", "notes"]

def decrypt(msg, priv):
    env = json.loads(base64.b64decode(msg))
    ck = priv.decrypt(base64.b64decode(env["k"]), padding.OAEP(mgf=padding.MGF1(hashes.SHA256()), algorithm=hashes.SHA256(), label=None))
    return json.loads(AESGCM(ck).decrypt(base64.b64decode(env["iv"]), base64.b64decode(env["d"]), None))

def clean(v):
    v = str(v or "").strip()[:500]
    return "'" + v if v and v[0] in "=+-@" else v

def load(pw):
    if not os.path.exists(XLSX):
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Inscrições"; ws.append(HEAD); return wb
    f = msoffcrypto.OfficeFile(open(XLSX, "rb")); f.load_key(password=pw)
    b = io.BytesIO(); f.decrypt(b); return openpyxl.load_workbook(b)

def save(wb, pw):
    ws = wb.active
    for col, w in zip("ABCDEFGHI", [10, 20, 30, 30, 16, 8, 26, 16, 40]): ws.column_dimensions[col].width = w
    for c in ws[1]: c.font = openpyxl.styles.Font(bold=True)
    ws.freeze_panes = "A2"
    b = io.BytesIO(); wb.save(b); b.seek(0)
    out = io.BytesIO(); msoffcrypto.OfficeFile(b).encrypt(pw, out)
    open(XLSX, "wb").write(out.getvalue())

def main(messages):
    pw = os.environ["XLSX_PASSWORD"]
    priv = serialization.load_pem_private_key(os.environ["RSA_PRIVATE_KEY"].encode(), None)
    wb = load(pw); ws = wb.active
    seen = {r[0] for r in ws.iter_rows(min_row=2, values_only=True)}
    added = 0
    for m in messages:
        try:
            d = decrypt(m, priv)
            ref = str(d["id"])[:12].upper()
            if ref in seen or not str(d.get("name", "")).strip(): continue
            ts = datetime.datetime.utcfromtimestamp(int(d["ts"]) / 1000).strftime("%Y-%m-%d %H:%M")
            ws.append([ref, ts] + [clean(d.get(k)) for k in KEYS]); seen.add(ref); added += 1
        except Exception:
            continue  # lixo / spam / não cifrado para nós
    if added or not os.path.exists(XLSX): save(wb, pw)
    print(f"novas inscrições: {added}")

if __name__ == "__main__":
    if len(sys.argv) > 1:  # modo teste: ficheiro com uma mensagem por linha
        msgs = open(sys.argv[1]).read().split()
    else:
        url = f"https://ntfy.sh/{os.environ['NTFY_TOPIC']}/json?poll=1&since=all"
        msgs = []
        for line in urllib.request.urlopen(url, timeout=30).read().decode().splitlines():
            j = json.loads(line)
            if j.get("event") == "message": msgs.append(j["message"])
    main(msgs)

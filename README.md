# [re]constrói — Conferência de Homens

Site: `index.html` (GitHub Pages). Inscrições: o formulário cifra os dados no browser; um GitHub Action
(`.github/workflows/inscricoes.yml`, de 20 em 20 min) decifra e acrescenta ao `inscricoes.xlsx`, cifrado com password (Secrets do repositório: `XLSX_PASSWORD`, `RSA_PRIVATE_KEY`, `NTFY_TOPIC`).

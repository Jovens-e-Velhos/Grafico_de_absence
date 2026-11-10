# Acompanhamento de faltas — Weekly Meeting

Painel que lê uma ou mais planilhas de presença da Weekly Meeting e mostra os itens abaixo.
Formatos aceitos: `.xlsx`, `.xlsm`, `.xls`, `.ods`, `.csv` e `.zip` contendo qualquer um desses (as planilhas de dentro do .zip são carregadas uma a uma).

- **Faltas por Líder Direto, Supervisor e Gerente** (com equipe e % de faltas). Cada nível usa a coluna de mesmo nome na planilha; a aba "Gerente" aparece quando a planilha tiver uma coluna "Gerente"
- **Faltas justificadas**: contagem por "Justificativa Base" (exceto "Outros")
- **Justificadas via "Outros"**: motivos informados em "Outro (detalhado)"
- **Faltas injustificadas**: "Justificativa Base" = "Outros (detalhe ao lado)" e "Report" = "Absence", com o detalhe informado
- **Evolução semanal** quando mais de uma planilha é carregada (cada arquivo = uma semana, lida do nome, ex.: `WM43`)

Os arquivos são processados **somente no navegador** — nada é enviado a servidores.

## Regras de classificação

| Report1 | Justificativa Base | Report | Classificação |
|---|---|---|---|
| Yes | – | – | Presente |
| Absence | qualquer opção ≠ "Outros" | – | Justificada |
| Absence | "Outros (detalhe ao lado)" | Absence | Injustificada |
| Absence | "Outros (detalhe ao lado)" | ≠ Absence | Justificada (Outros) |
| Absence | vazio | – | Ausente sem justificativa base |

## Estrutura

```
index.html        página publicada no GitHub Pages
assets/app.js     leitura das planilhas (SheetJS) + gráficos (ECharts)
assets/style.css  identidade visual (DSV Corporate Visual Identity)
app.py, processamento.py, iniciar.bat   versão local alternativa em Python/Streamlit
```

## Publicar no GitHub Pages

1. Crie um repositório no GitHub (de preferência **privado** ou interno da organização).
2. Envie os arquivos desta pasta. O `.gitignore` já impede que planilhas `.xlsx`/`.csv` sejam enviadas.
   ```
   git init
   git add .
   git commit -m "Painel de faltas da Weekly Meeting"
   git branch -M main
   git remote add origin https://github.com/<usuario>/<repositorio>.git
   git push -u origin main
   ```
3. No repositório: **Settings → Pages → Build and deployment → Source: Deploy from a branch**, branch `main`, pasta `/ (root)`.
4. Em ~1 minuto o painel fica em `https://<usuario>.github.io/<repositorio>/`.

> Observação: em contas gratuitas, o site do GitHub Pages é público mesmo com repositório privado.
> Como nenhum dado fica no site (as planilhas são carregadas por quem abre a página), isso não expõe dados dos colaboradores.

## Rodar localmente

- **Versão web**: abra `index.html` no navegador.
- **Versão Python**: `pip install -r requirements.txt` e depois `iniciar.bat` (ou `streamlit run app.py`).

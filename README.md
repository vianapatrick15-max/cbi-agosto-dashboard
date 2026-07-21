# CBI · Agosto Solidário — Dashboard de Performance

Tracker ao vivo da captação do **Agosto Solidário (CBI of Miami)**. Mesmo padrão
dos dashboards SINAREM / CBI Diagnóstico: `refresh.py` lê as planilhas via service
account, gera `data.json` granular e injeta em `template.html` → `index.html`
(GitHub Pages). O filtro de datas recalcula tudo no navegador.

## Seções
1. **Investimento, leads e CPL** — investido, leads total, CPL total, leads pagos,
   leads orgânicos, CPL pago.
2. **Perfil dos leads** — distribuição por Tipo do Contato (Psicopedagogo,
   Psicólogo, Professor, Terapeuta Ocupacional, etc.).
3. **Evolução por dia** — inscritos acumulados (período completo) + barras de
   inscritos/dia com linha de investimento.
4. **Anúncios · performance** — só anúncios pagos (Meta), leads atribuídos por
   UTM Content ↔ Ad Name; CPM, CTR, CPC, connect rate, CPL.

## Fontes
- **Tráfego (Meta):** planilha `1puw54S9fYWrGwdgp1jV2KGOo619oWQx0bNp6mecX9-c`
  (primeira aba). Cabeçalhos resolvidos por candidatos EN+PT.
- **Leads:** planilha `1vcpyCCE0d8zvoSfZqEacvRcJ3yQQgZdogO7CJu32MwA`,
  aba `[CBI] Leads Agosto Solidário` (HubSpot, um lead por linha).

Ambas precisam estar compartilhadas com a service account
`ga4-reader@n8n-tathi.iam.gserviceaccount.com` (Leitor).

## Rodar local
```bash
pip install -r requirements.txt
python refresh.py            # usa ~/.claude/skills/ga4/credentials/ga4-instituto-andhela.json
```

## Deploy
GitHub Pages (branch principal, raiz). Secret do Action:
`GOOGLE_SHEETS_CREDENTIALS_JSON` = JSON da service account. Refresh horário
(cron `17 * * * *`) + `workflow_dispatch`.

Sem PII no `data.json` (só contagens). Guard anti-wipe se o spend cair >50%.

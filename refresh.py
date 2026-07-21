#!/usr/bin/env python3
"""
Dashboard CBI of Miami — Agosto Solidário (captação de leads / lançamento).

Le duas planilhas (service account) e emite dados GRANULARES (por dia / anuncio /
lead) para que o filtro de datas do front recalcule tudo no navegador.

FONTES:
  - Trafego (Meta): planilha 1puw54S9... , aba unica (gid=0).
      entrega por linha (dia/campanha/adset/ad): spend, impressoes, alcance,
      cliques no link, LPV. Cabecalhos resolvidos por candidatos EN+PT.
  - Leads: planilha 1vcpyCCE... (a MESMA da Aristo/Diagnostico),
      aba `[CBI] Leads Agosto Solidário` — um lead por linha (HubSpot).
      Atribuicao ad-level por UTM Content <-> Ad Name. Guarda tambem o
      "Tipo do Contato" (profissao) para o recorte de perfil.

Tracker ao vivo, sem meta. Nenhum dado pessoal (e-mail) vai pro data.json.
"""
import os, json, datetime as dt
from pathlib import Path

# ----------------------------- FONTES -----------------------------
TRAF_SID  = "1puw54S9fYWrGwdgp1jV2KGOo619oWQx0bNp6mecX9-c"
TRAF_TAB  = None          # None = primeira aba (gid=0) -> Meta
GADS_TAB  = "Google_ads"  # aba de Google Ads (gid=721710950), nivel campanha
LEADS_SID = "1vcpyCCE0d8zvoSfZqEacvRcJ3yQQgZdogO7CJu32MwA"
TAB_LEADS = "[CBI] Leads Agosto Solidário"

OUT = Path(__file__).parent / "data.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]
LOCAL_CRED = os.path.expanduser("~/.claude/skills/ga4/credentials/ga4-instituto-andhela.json")

# Classificacao de origem por UTM Source. Agora ha verba de Meta E Google na
# planilha de trafego, entao ambos contam como PAGO. dgen_ads/search_ads = Google.
META_SOURCES   = {"ig", "fb", "facebook", "instagram", "meta", "meta_ads", "meta-ads",
                  "facebook_ads", "instagram_ads", "ig_ads"}
GOOGLE_SOURCES = {"dgen_ads", "search_ads", "google_ads", "google-ads", "google", "gads"}
PAID_SOURCES   = META_SOURCES | GOOGLE_SOURCES

def channel_of(src):
    s = (src or "").strip().lower()
    if s in META_SOURCES:   return "meta"
    if s in GOOGLE_SOURCES: return "google"
    return "org"

# Candidatos de cabecalho (case-insensitive) para a planilha de trafego.
H_DAY   = ["day", "date", "dia", "data", "reporting starts", "data de início dos relatórios"]
H_CAMP  = ["campaign name", "campaign", "campanha", "nome da campanha"]
H_ADSET = ["ad set name", "adset name", "ad set", "adset", "conjunto de anúncios",
           "nome do conjunto de anúncios"]
H_AD    = ["ad name", "ad", "anúncio", "nome do anúncio", "nome do anuncio"]
H_SPEND = ["amount spent", "amount spent (brl)", "spend", "cost (spend)", "cost", "custo",
           "valor gasto", "valor usado", "valor usado (brl)", "investimento", "gasto"]
H_IMP   = ["impressions", "impressões", "impressoes"]
H_REACH = ["reach", "alcance"]
H_CLICK = ["link clicks", "cliques no link", "clicks", "cliques", "cliques (todos)"]
H_LPV   = ["landing page views", "visualizações da página de destino",
           "visualizacoes da pagina de destino", "lpv"]

try:
    from zoneinfo import ZoneInfo
    TODAY = dt.datetime.now(ZoneInfo("America/Sao_Paulo")).date()
except Exception:
    TODAY = (dt.datetime.utcnow() - dt.timedelta(hours=3)).date()

# ----------------------------- HELPERS -----------------------------
def get_client():
    import gspread
    from google.oauth2.service_account import Credentials
    raw = os.environ.get("GOOGLE_SHEETS_CREDENTIALS_JSON")
    if raw:
        creds = Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
    else:
        path = os.environ.get("GOOGLE_SHEETS_CREDENTIALS_PATH", LOCAL_CRED)
        creds = Credentials.from_service_account_file(path, scopes=SCOPES)
    return gspread.authorize(creds)

def num(x):
    """Numero em formato BR (1.234,56) ou US. Robusto a R$, %, espacos."""
    if x is None: return 0.0
    s = str(x).strip().replace("R$", "").replace("%", "").replace("\xa0", " ").strip()
    if s in ("", "-", "—"): return 0.0
    # BR: ponto = milhar, virgula = decimal
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    s = s.replace(" ", "")
    try: return float(s)
    except: return 0.0

def inum(x): return int(round(num(x)))

def daykey(s):
    s = (s or "").strip()
    if not s: return ""
    d = s.split(" ")[0]
    if "-" in d and len(d) >= 10: return d[:10]
    try:
        dd, mm, yy = d.split("/")
        if len(yy) == 2: yy = "20" + yy
        return f"{yy}-{mm.zfill(2)}-{dd.zfill(2)}"
    except: return ""

def ws_by_title(sheet, title):
    for w in sheet.worksheets():
        if w.title.strip() == title.strip():
            return w
    raise KeyError(title)

def hidx(header, candidates):
    """Indice da 1a coluna cujo nome bate (case-insensitive, exato) um candidato."""
    low = {c.strip().lower(): i for i, c in enumerate(header)}
    for cand in candidates:
        if cand in low:
            return low[cand]
    # fallback: contains
    for cand in candidates:
        for name, i in low.items():
            if cand in name:
                return i
    return None

def norm(s):
    return (s or "").strip().lower()

# ----------------------------- LEITURA -----------------------------
gc = get_client()

# ---- Trafego (granular por dia/anuncio) ----
traf = gc.open_by_key(TRAF_SID)
tw = ws_by_title(traf, TRAF_TAB) if TRAF_TAB else traf.get_worksheet(0)
g = tw.get_all_values()
h = g[0]
cDay, cCamp, cAdset, cAd = hidx(h, H_DAY), hidx(h, H_CAMP), hidx(h, H_ADSET), hidx(h, H_AD)
cSp, cImp, cReach = hidx(h, H_SPEND), hidx(h, H_IMP), hidx(h, H_REACH)
cClk, cLpv = hidx(h, H_CLICK), hidx(h, H_LPV)
if cDay is None or cSp is None:
    raise SystemExit(f"[ERRO] trafego: nao achei Day/Spend. Cabecalhos: {h}")

def cell(r, i): return r[i] if (i is not None and len(r) > i) else ""

deliv = []
spend_days = set()
for r in g[1:]:
    if cDay is None or len(r) <= cDay or not str(r[cDay]).strip(): continue
    day = daykey(r[cDay])
    if not day: continue
    deliv.append({
        "d": day,
        "camp": cell(r, cCamp).strip(),
        "adset": cell(r, cAdset).strip(),
        "ad": cell(r, cAd).strip() or "(sem anúncio)",
        "s": round(num(cell(r, cSp)), 2),
        "i": inum(cell(r, cImp)),
        "r": inum(cell(r, cReach)),
        "c": inum(cell(r, cClk)),
        "lpv": inum(cell(r, cLpv)),
        "ch": "meta",
    })
    spend_days.add(day)

# ---- Trafego Google Ads (nivel campanha; sem LPV/alcance na fonte) ----
try:
    gw = ws_by_title(traf, GADS_TAB)
    gg = gw.get_all_values()
    gh = gg[0]
    gDay, gCamp = hidx(gh, H_DAY), hidx(gh, H_CAMP)
    gSp, gImp, gClk = hidx(gh, H_SPEND), hidx(gh, H_IMP), hidx(gh, H_CLICK)
    for r in gg[1:]:
        if gDay is None or len(r) <= gDay or not str(r[gDay]).strip(): continue
        day = daykey(r[gDay])
        if not day: continue
        camp = cell(r, gCamp).strip()
        deliv.append({
            "d": day,
            "camp": camp,
            "adset": "",
            "ad": camp or "(Google Ads)",
            "s": round(num(cell(r, gSp)), 2),
            "i": inum(cell(r, gImp)),
            "r": 0,
            "c": inum(cell(r, gClk)),
            "lpv": 0,
            "ch": "google",
        })
        spend_days.add(day)
except KeyError:
    pass  # aba de Google ainda nao existe nessa planilha

# ---- Leads (um por linha, sem PII) ----
leadsheet = gc.open_by_key(LEADS_SID)
lv = ws_by_title(leadsheet, TAB_LEADS).get_all_values()
lh = {x.strip(): i for i, x in enumerate(lv[0])}
iD    = lh.get("Data de conversão recente")
iSrc  = lh.get("UTM Source")
iCont = lh.get("UTM Content")
iTipo = lh.get("Tipo do Contato")
leads = []
lead_days = set()
for r in lv[1:]:
    if not any(str(x).strip() for x in r): continue
    if iD is None or len(r) <= iD or not str(r[iD]).strip(): continue
    day = daykey(r[iD])
    if not day: continue
    src = norm(cell(r, iSrc))
    ch  = channel_of(src)
    leads.append({
        "d": day,
        "ct": norm(cell(r, iCont)) or "(sem atribuição)",
        "src": src or "(sem origem)",
        "tp": cell(r, iTipo).strip(),
        "p": 1 if ch in ("meta", "google") else 0,
        "ch": ch,
    })
    lead_days.add(day)

# ----------------------------- JANELA -----------------------------
all_days = {d for d in (spend_days | lead_days) if d}
data_from = min(all_days) if all_days else TODAY.isoformat()
data_to   = max(all_days | {TODAY.isoformat()})
spend_from = min(spend_days) if spend_days else data_from
spend_to   = max(spend_days) if spend_days else data_from

camp_name = deliv[0]["camp"] if deliv else "Agosto Solidário"

data = {
    "updated_at": dt.datetime(TODAY.year, TODAY.month, TODAY.day, 12, 0).isoformat(),
    "today": TODAY.isoformat(),
    "data_from": data_from, "data_to": data_to,
    "spend_from": spend_from, "spend_to": spend_to,
    "campaign": camp_name,
    "deliv": deliv, "leads": leads,
}

# ----------------------------- GUARD anti-wipe -----------------------------
def total_spend(records): return sum(x["s"] for x in records)
if OUT.exists():
    try:
        prev = json.loads(OUT.read_text())
        p, n = total_spend(prev.get("deliv", [])), total_spend(deliv)
        if p >= 300 and n < 0.5 * p:
            raise SystemExit(f"[GUARD] spend {p:.0f}->{n:.0f} (fonte em rewrite); mantendo snapshot anterior.")
    except SystemExit: raise
    except Exception: pass

OUT.write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")))
base = Path(__file__).parent
tpl = (base / "template.html").read_text()
(base / "index.html").write_text(tpl.replace("__DATA__", json.dumps(data, ensure_ascii=False)))

# ----------------------------- RESUMO -----------------------------
sp_meta = sum(x["s"] for x in deliv if x.get("ch") != "google")
sp_gads = sum(x["s"] for x in deliv if x.get("ch") == "google")
l_meta = sum(1 for x in leads if x.get("ch") == "meta")
l_gads = sum(1 for x in leads if x.get("ch") == "google")
print(f"CBI Agosto Solidario | dados {data_from}..{data_to} | spend {spend_from}..{spend_to}")
print(f"deliv {len(deliv)} regs | leads {len(leads)}")
print(f"spend R$ {total_spend(deliv):,.2f}  (Meta R$ {sp_meta:,.2f} | Google R$ {sp_gads:,.2f})")
lp = sum(1 for x in leads if x['p'] == 1)
print(f"leads pagos {lp} (Meta {l_meta} | Google {l_gads}) | organicos {len(leads)-lp}")
print(f"OK -> {OUT} ({OUT.stat().st_size//1024} KB)")

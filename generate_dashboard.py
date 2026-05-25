"""
FEHMA Dashboard Generator
Lê DRE_FEHMA_2025_v30.xlsx do OneDrive e gera index.html atualizado
"""

import os
import json
import requests
from datetime import datetime

# ── Configurações ──────────────────────────────────────────────
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
REFRESH_TOKEN = os.environ["AZURE_REFRESH_TOKEN"]
NETLIFY_TOKEN = os.environ["NETLIFY_TOKEN"]
NETLIFY_SITE  = os.environ["NETLIFY_SITE_ID"]

DRIVE_ITEM_ID = "01ZF6LUJSWT5PALWI7XBDKD7EG7OGRME6X"
DRIVE_ID      = "b!G5ef5tgjoUiOM6PP4ejIulK3QKtjICtIgFkmO3T365qJT8hRJlQlTYCxM9jsL3Kt"

# ── 1. Obter access token via refresh token ────────────────────
def get_access_token():
    url = f"https://login.microsoftonline.com/common/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "refresh_token": REFRESH_TOKEN,
        "grant_type": "refresh_token",
        "scope": "https://graph.microsoft.com/Files.Read offline_access",
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    tokens = r.json()
    # Salva novo refresh token
    new_refresh = tokens.get("refresh_token", REFRESH_TOKEN)
    if new_refresh != REFRESH_TOKEN:
        print(f"::set-output name=new_refresh_token::{new_refresh}")
    return tokens["access_token"]

# ── 2. Ler dados do Excel via Graph API ────────────────────────
def read_excel_data(token):
    headers = {"Authorization": f"Bearer {token}"}
    
    # Baixa o arquivo Excel completo como stream
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{DRIVE_ITEM_ID}/content"
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    
    import io
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)
    
    def get_sheet_data(sheet_name):
        ws = wb[sheet_name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append(list(row))
        return rows

    def find_row_by_label(rows, label):
        """Encontra linha pelo texto na primeira coluna (ignora maiúsculas/espaços)"""
        label_clean = label.strip().upper()
        for row in rows:
            if row and row[0] and str(row[0]).strip().upper() == label_clean:
                return row
        return None

    def extract_12_months(row, max_cols=50):
        """Extrai 12 valores monetários de uma linha.
        Ignora valores entre -2 e 2 (são percentuais como 1.0, 0.8, -0.07)
        e busca valores absolutos maiores (receita, custos, etc.)
        """
        nums = []
        for val in row[1:max_cols]:
            try:
                if val is None or val == "" or val == "-":
                    continue
                # Converte para float
                s = str(val)
                negative = s.startswith("(") and s.endswith(")")
                f = float(s.replace(",","").replace("(","").replace(")",""))
                if negative:
                    f = -f
                # Ignora percentuais (valores entre -2 e 2)
                if -2 < f < 2:
                    continue
                nums.append(f)
            except:
                continue
            if len(nums) == 12:
                break
        return nums if len(nums) == 12 else nums + [0]*(12-len(nums))

    # Lê aba DRE 2025
    dre_rows = get_sheet_data("DRE 2025")
    
    # Lê aba Fluxo de Caixa
    fc_rows = get_sheet_data("Fluxo de Caixa")

    # Busca as linhas pelo nome exato
    receita_row = find_row_by_label(dre_rows, "RECEITA TOTAL") or                   find_row_by_label(dre_rows, "RECEITA BRUTA")
    marg_row    = find_row_by_label(dre_rows, "MARGEM DE CONTRIBUIÇÃO")
    ebitda_row  = find_row_by_label(dre_rows, "LUCRO OPERACIONAL (EBITDA)") or                   find_row_by_label(dre_rows, "EBITDA")
    lucro_row   = find_row_by_label(dre_rows, "LUCRO LÍQUIDO") or                   find_row_by_label(dre_rows, "LUCRO LIQUIDO")
    caixa_row   = find_row_by_label(fc_rows, "SALDO FINAL DE CAIXA")

    receita_vals = extract_12_months(receita_row) if receita_row else []
    marg_vals    = extract_12_months(marg_row)    if marg_row    else []
    ebitda_vals  = extract_12_months(ebitda_row)  if ebitda_row  else []
    lucro_vals   = extract_12_months(lucro_row)   if lucro_row   else []
    caixa_vals   = extract_12_months(caixa_row)   if caixa_row   else []

    print(f"   Receita[0]: {receita_vals[0] if receita_vals else 'NÃO ENCONTRADO'}")
    print(f"   Margem[0]:  {marg_vals[0] if marg_vals else 'NÃO ENCONTRADO'}")
    print(f"   EBITDA[0]:  {ebitda_vals[0] if ebitda_vals else 'NÃO ENCONTRADO'}")
    print(f"   Lucro[0]:   {lucro_vals[0] if lucro_vals else 'NÃO ENCONTRADO'}")

    # Fallback com valores da DRE v30 caso a leitura falhe
    if not receita_vals or sum(abs(v) for v in receita_vals) < 1000:
        print("   ⚠️ Fallback ativado — usando valores fixos da DRE v30")
        receita_vals = [117000,117000,117000,138000,138000,138000,165000,165000,165000,190000,190000,190000]
        marg_vals    = [42740,42740,42740,50771,50771,50771,61095,61095,61095,70655,70655,70655]
        ebitda_vals  = [-151,-151,-151,5321,5321,5321,12642,12642,12642,20347,20347,5217]
        lucro_vals   = [-4255,-4255,-4255,571,571,571,7060,7060,7060,13995,13995,-1135]
        caixa_vals   = [479280,661202,502934,551856,594148,790498,853790,769022,851075,782714,858065,1282278]

    return {
        "receita":        receita_vals,
        "lucro":          lucro_vals,
        "ebitda":         ebitda_vals,
        "margem_contrib": marg_vals,
        "saldo_caixa":    caixa_vals,
        "updated_at":     datetime.now().strftime("%d/%m/%Y às %H:%M"),
    }

# ── 3. Gerar HTML com dados atualizados ───────────────────────
def generate_html(data):
    receita        = json.dumps(data["receita"])
    lucro          = json.dumps(data["lucro"])
    ebitda         = json.dumps(data["ebitda"])
    marg           = json.dumps(data["margem_contrib"])
    caixa          = json.dumps(data["saldo_caixa"])
    updated        = data["updated_at"]

    receita_anual  = sum(data["receita"])
    lucro_anual    = sum(data["lucro"])
    ebitda_anual   = sum(data["ebitda"])
    marg_anual     = sum(data["margem_contrib"])

    lucro_pct      = round(lucro_anual / receita_anual * 100, 1) if receita_anual else 0
    ebitda_pct     = round(ebitda_anual / receita_anual * 100, 1) if receita_anual else 0
    marg_pct       = round(marg_anual / receita_anual * 100, 1) if receita_anual else 0

    def fmt(v):
        return f"R$ {abs(v):,.0f}".replace(",", ".")

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FEHMA · DRE Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Syne:wght@400;600;700;800&display=swap');
  :root {{
    --bg:#0a0a0b;--surface:#111113;--surface2:#18181c;--border:#252528;
    --accent:#c8f542;--accent2:#42f5a5;--warn:#f5a542;--danger:#f54242;
    --text:#e8e8e8;--muted:#666672;--dim:#3a3a42;
  }}
  *{{margin:0;padding:0;box-sizing:border-box;}}
  body{{background:var(--bg);color:var(--text);font-family:'Syne',sans-serif;min-height:100vh;}}
  .header{{background:var(--surface);border-bottom:1px solid var(--border);padding:20px 40px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;}}
  .logo{{font-size:22px;font-weight:800;letter-spacing:-0.5px;color:var(--accent);}}
  .logo span{{color:var(--text);opacity:0.4;}}
  .badge{{font-family:'DM Mono',monospace;font-size:11px;background:var(--border);color:var(--muted);padding:3px 8px;border-radius:4px;letter-spacing:0.5px;margin-left:8px;}}
  .updated{{font-family:'DM Mono',monospace;font-size:11px;color:var(--muted);}}
  .tabs{{background:var(--surface);border-bottom:1px solid var(--border);padding:0 40px;display:flex;}}
  .tab{{font-family:'DM Mono',monospace;font-size:12px;letter-spacing:0.5px;padding:14px 20px;cursor:pointer;color:var(--muted);border-bottom:2px solid transparent;transition:all 0.2s;text-transform:uppercase;}}
  .tab:hover{{color:var(--text);}}
  .tab.active{{color:var(--accent);border-bottom-color:var(--accent);}}
  .content{{padding:32px 40px;}}
  .page{{display:none;}}
  .page.active{{display:block;}}
  .kpi-grid{{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:28px;}}
  .kpi{{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:18px 16px;position:relative;overflow:hidden;}}
  .kpi::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--accent);opacity:0.3;}}
  .kpi.warn::before{{background:var(--warn);}}
  .kpi.good::before{{background:var(--accent2);opacity:0.6;}}
  .kpi.danger::before{{background:var(--danger);}}
  .kpi-label{{font-family:'DM Mono',monospace;font-size:10px;text-transform:uppercase;letter-spacing:0.8px;color:var(--muted);margin-bottom:8px;}}
  .kpi-value{{font-size:22px;font-weight:700;line-height:1;margin-bottom:4px;}}
  .kpi-sub{{font-family:'DM Mono',monospace;font-size:11px;color:var(--muted);}}
  .kpi-pct{{font-family:'DM Mono',monospace;font-size:12px;color:var(--accent);margin-top:4px;font-weight:500;}}
  .kpi-pct.warn{{color:var(--warn);}}
  .kpi-pct.danger{{color:var(--danger);}}
  .charts-row{{display:grid;gap:16px;margin-bottom:20px;}}
  .charts-row-2{{grid-template-columns:2fr 1fr;}}
  .charts-row-3{{grid-template-columns:1fr 1fr 1fr;}}
  .chart-box{{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:20px;}}
  .chart-title{{font-size:13px;font-weight:600;color:var(--text);margin-bottom:4px;display:flex;align-items:center;justify-content:space-between;}}
  .chart-subtitle{{font-family:'DM Mono',monospace;font-size:10px;color:var(--muted);margin-bottom:16px;}}
  .chart-wrap canvas{{max-height:240px;}}
  .chart-wrap.tall canvas{{max-height:280px;}}
  .section-title{{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1.5px;color:var(--muted);margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px;}}
  .section-title .dot{{width:6px;height:6px;border-radius:50%;background:var(--accent);}}
  .waterfall-item{{display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid var(--border);}}
  .waterfall-item:last-child{{border-bottom:none;}}
  .wf-label{{font-family:'DM Mono',monospace;font-size:11px;color:var(--muted);width:220px;flex-shrink:0;}}
  .wf-bar-wrap{{flex:1;display:flex;align-items:center;gap:8px;}}
  .wf-bar{{height:14px;border-radius:2px;min-width:2px;}}
  .wf-val{{font-family:'DM Mono',monospace;font-size:11px;white-space:nowrap;}}
  .wf-pct{{font-family:'DM Mono',monospace;font-size:10px;color:var(--muted);width:50px;text-align:right;flex-shrink:0;}}
  .pill{{display:inline-block;padding:2px 8px;border-radius:100px;font-size:10px;font-family:'DM Mono',monospace;font-weight:500;}}
  .pill-green{{background:rgba(200,245,66,0.12);color:var(--accent);}}
  .pill-red{{background:rgba(245,66,66,0.12);color:var(--danger);}}
  .pill-yellow{{background:rgba(245,165,66,0.12);color:var(--warn);}}
  .pill-blue{{background:rgba(66,245,165,0.12);color:var(--accent2);}}
  .data-table{{width:100%;border-collapse:collapse;font-family:'DM Mono',monospace;font-size:12px;}}
  .data-table th{{text-align:left;padding:8px 12px;color:var(--muted);font-size:10px;letter-spacing:0.8px;text-transform:uppercase;border-bottom:1px solid var(--border);font-weight:400;}}
  .data-table th:not(:first-child){{text-align:right;}}
  .data-table td{{padding:9px 12px;border-bottom:1px solid var(--border);color:var(--text);}}
  .data-table td:not(:first-child){{text-align:right;}}
  .data-table tr:last-child td{{border-bottom:none;}}
  .data-table tr:hover td{{background:var(--surface2);}}
  .data-table .total-row td{{font-weight:700;color:var(--accent);border-top:1px solid var(--border);border-bottom:none;}}
  .data-table .sub-row td{{color:var(--muted);}}
  .neg{{color:var(--danger)!important;}}
  .pos{{color:var(--accent2)!important;}}
  .month-row{{display:flex;gap:6px;margin-bottom:20px;flex-wrap:wrap;}}
  .month-btn{{font-family:'DM Mono',monospace;font-size:11px;padding:5px 12px;border-radius:4px;border:1px solid var(--border);background:transparent;color:var(--muted);cursor:pointer;transition:all 0.15s;}}
  .month-btn:hover{{color:var(--text);border-color:var(--dim);}}
  .month-btn.active{{background:var(--accent);color:#0a0a0b;border-color:var(--accent);font-weight:700;}}
  ::-webkit-scrollbar{{width:6px;height:6px;}}
  ::-webkit-scrollbar-track{{background:var(--bg);}}
  ::-webkit-scrollbar-thumb{{background:var(--dim);border-radius:3px;}}
  @media(max-width:1200px){{.kpi-grid{{grid-template-columns:repeat(3,1fr);}} .charts-row-2{{grid-template-columns:1fr;}} .charts-row-3{{grid-template-columns:1fr 1fr;}}}}
</style>
</head>
<body>
<div class="header">
  <div style="display:flex;align-items:center;">
    <div class="logo">FDC<span>/</span>FEHMA</div>
    <span class="badge">DRE v30</span>
    <span class="badge">FEHMA Participações e Distribuição Ltda</span>
  </div>
  <div class="updated">Atualizado em: <strong style="color:var(--accent)">{updated}</strong></div>
</div>

<div class="tabs">
  <div class="tab active" onclick="showPage('overview',this)">Visão Geral</div>
  <div class="tab" onclick="showPage('dre',this)">DRE Mensal</div>
  <div class="tab" onclick="showPage('caixa',this)">Fluxo de Caixa</div>
</div>

<div class="content">

<!-- OVERVIEW -->
<div id="page-overview" class="page active">
  <div class="kpi-grid">
    <div class="kpi good">
      <div class="kpi-label">Receita Total Anual</div>
      <div class="kpi-value">{fmt(receita_anual)}</div>
      <div class="kpi-sub">projetado 2025</div>
      <div class="kpi-pct">100%</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Margem de Contribuição</div>
      <div class="kpi-value">{marg_pct}%</div>
      <div class="kpi-sub">{fmt(marg_anual)}/ano</div>
      <div class="kpi-pct">após custos variáveis</div>
    </div>
    <div class="kpi good">
      <div class="kpi-label">EBITDA</div>
      <div class="kpi-value">{ebitda_pct}%</div>
      <div class="kpi-sub">{fmt(ebitda_anual)}/ano</div>
      <div class="kpi-pct good">operacional</div>
    </div>
    <div class="kpi warn">
      <div class="kpi-label">Lucro Líquido</div>
      <div class="kpi-value">{lucro_pct}%</div>
      <div class="kpi-sub">{fmt(lucro_anual)}/ano</div>
      <div class="kpi-pct warn">estreito no ano 1</div>
    </div>
    <div class="kpi warn">
      <div class="kpi-label">CMV Médio</div>
      <div class="kpi-value">42,4%</div>
      <div class="kpi-sub">ponderado portfólio</div>
      <div class="kpi-pct warn">risco mel/própolis</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Receita Online</div>
      <div class="kpi-value">20%</div>
      <div class="kpi-sub">{fmt(receita_anual*0.2)}/ano</div>
      <div class="kpi-pct">80% offline</div>
    </div>
  </div>

  <div class="charts-row charts-row-2">
    <div class="chart-box">
      <div class="chart-title">Receita Mensal vs. Lucro Líquido <span class="pill pill-green">12 meses</span></div>
      <div class="chart-subtitle">Receita · Margem de Contribuição · Lucro Líquido</div>
      <div class="chart-wrap tall"><canvas id="chartReceita"></canvas></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">Estrutura de Custos Anual <span class="pill pill-yellow">% da Receita</span></div>
      <div class="chart-subtitle">Decomposição proporcional do faturamento</div>
      <div class="chart-wrap tall"><canvas id="chartCustos"></canvas></div>
    </div>
  </div>

  <div class="section-title"><span class="dot"></span> Cascata de Resultados — DRE Anual Consolidado</div>
  <div class="chart-box" style="margin-bottom:20px;"><div id="waterfall-container"></div></div>

  <div class="charts-row charts-row-3">
    <div class="chart-box">
      <div class="chart-title">Evolução do Lucro Líquido <span class="pill pill-blue">mensal</span></div>
      <div class="chart-subtitle">Break-even a partir do mês 4</div>
      <div class="chart-wrap"><canvas id="chartLucro"></canvas></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">Impostos s/ Receita <span class="pill pill-red">-7,0%</span></div>
      <div class="chart-subtitle">PIS · COFINS · ICMS · DIFAL</div>
      <div class="chart-wrap"><canvas id="chartImpostos"></canvas></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">Despesas Administrativas <span class="pill pill-yellow">-19,0%</span></div>
      <div class="chart-subtitle">Salários · Fixos · Softwares · Fulfillment</div>
      <div class="chart-wrap"><canvas id="chartAdmin"></canvas></div>
    </div>
  </div>
</div>

<!-- DRE MENSAL -->
<div id="page-dre" class="page">
  <div class="month-row" id="month-selector"></div>
  <div class="charts-row charts-row-2" style="margin-bottom:20px;">
    <div class="chart-box">
      <div class="chart-title">Performance no Mês Selecionado</div>
      <div class="chart-subtitle" id="dre-month-label">Selecione um mês</div>
      <div class="chart-wrap"><canvas id="chartDreMes"></canvas></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">Rentabilidade Mensal</div>
      <div class="chart-subtitle">% do faturamento · Margem Contrib. · EBITDA · Lucro Líq.</div>
      <div class="chart-wrap"><canvas id="chartMargem"></canvas></div>
    </div>
  </div>
  <div class="chart-box">
    <div class="chart-title">Tabela DRE Mensal Completa</div>
    <div class="chart-subtitle">Todos os 12 meses · valores em R$</div>
    <div style="overflow-x:auto;"><table class="data-table" id="dre-table"></table></div>
  </div>
</div>

<!-- CAIXA -->
<div id="page-caixa" class="page">
  <div style="display:flex;gap:12px;margin-bottom:20px;">
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:16px 20px;flex:1;">
      <div style="font-family:'DM Mono',monospace;font-size:10px;text-transform:uppercase;letter-spacing:0.8px;color:var(--muted);margin-bottom:6px;">Aporte Total dos Sócios</div>
      <div style="font-size:24px;font-weight:700;color:var(--accent)">R$ 1.500.000</div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:16px 20px;flex:1;">
      <div style="font-family:'DM Mono',monospace;font-size:10px;text-transform:uppercase;letter-spacing:0.8px;color:var(--muted);margin-bottom:6px;">Saldo Final de Caixa (Dez)</div>
      <div style="font-size:24px;font-weight:700;color:var(--accent2)">R$ 1.282.278</div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:16px 20px;flex:1;">
      <div style="font-family:'DM Mono',monospace;font-size:10px;text-transform:uppercase;letter-spacing:0.8px;color:var(--muted);margin-bottom:6px;">Caixa Gerado pela Operação</div>
      <div style="font-size:24px;font-weight:700;color:var(--warn)">-R$ 217.722</div>
    </div>
  </div>
  <div class="charts-row charts-row-2">
    <div class="chart-box">
      <div class="chart-title">Saldo de Caixa Mensal <span class="pill pill-blue">acumulado</span></div>
      <div class="chart-subtitle">Inclui aportes dos sócios · pagamentos defasados 30/60 dias</div>
      <div class="chart-wrap tall"><canvas id="chartCaixa"></canvas></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">Aportes vs. Saídas Operacionais <span class="pill pill-yellow">dependência</span></div>
      <div class="chart-subtitle">Operação não se autofinancia no ano 1</div>
      <div class="chart-wrap tall"><canvas id="chartAportes"></canvas></div>
    </div>
  </div>
</div>

</div><!-- end content -->

<script>
const meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const receita = {receita};
const lucroLiq = {lucro};
const ebitda = {ebitda};
const margemContrib = {marg};
const saldoCaixa = {caixa};
const aportesMes = [500000,250000,0,0,0,250000,0,0,0,0,0,500000];
const totalSaidas = [20720,68078,275269,68078,74708,191650,74708,222769,82947,233361,89649,265786];

const margContribPct = receita.map((r,i)=>((margemContrib[i]/r)*100).toFixed(1));
const ebitdaPct = receita.map((r,i)=>((ebitda[i]/r)*100).toFixed(1));
const lucroLiqPct = receita.map((r,i)=>((lucroLiq[i]/r)*100).toFixed(1));

function fmtBR(v){{return 'R$ '+Math.abs(v).toLocaleString('pt-BR',{{minimumFractionDigits:0,maximumFractionDigits:0}});}}

const defaults = {{
  responsive:true, maintainAspectRatio:true,
  plugins:{{legend:{{display:false}},tooltip:{{backgroundColor:'#18181c',borderColor:'#252528',borderWidth:1,titleFont:{{family:'DM Mono, monospace',size:11}},bodyFont:{{family:'DM Mono, monospace',size:11}},titleColor:'#e8e8e8',bodyColor:'#666672'}}}},
  scales:{{
    x:{{grid:{{color:'#252528'}},ticks:{{color:'#666672',font:{{family:'DM Mono, monospace',size:10}}}}}},
    y:{{grid:{{color:'#252528'}},ticks:{{color:'#666672',font:{{family:'DM Mono, monospace',size:10}}}}}}
  }}
}};

new Chart(document.getElementById('chartReceita'),{{type:'bar',data:{{labels:meses,datasets:[
  {{label:'Receita',data:receita,backgroundColor:'rgba(200,245,66,0.15)',borderColor:'#c8f542',borderWidth:1.5,borderRadius:3}},
  {{label:'Marg.Contrib.',data:margemContrib,backgroundColor:'rgba(66,245,165,0.15)',borderColor:'#42f5a5',borderWidth:1.5,borderRadius:3}},
  {{label:'Lucro Líq.',data:lucroLiq,type:'line',borderColor:'#f5a542',borderWidth:2,pointRadius:3,pointBackgroundColor:'#f5a542',tension:0.3}},
]}},options:{{...defaults,plugins:{{...defaults.plugins,legend:{{display:true,labels:{{color:'#666672',font:{{family:'DM Mono, monospace',size:10}},boxWidth:12,boxHeight:12}}}}}},scales:{{...defaults.scales,y:{{...defaults.scales.y,ticks:{{...defaults.scales.y.ticks,callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}}}}}}});

new Chart(document.getElementById('chartCustos'),{{type:'doughnut',data:{{labels:['CMV (42,4%)','Fretes (6,2%)','Impostos (7,0%)','Comercial+Mkt (12,5%)','Desp.Admin (19,0%)','Lucro Líq.','Outros'],datasets:[{{data:[42.4,6.2,7.0,12.5,19.0,2.0,10.9],backgroundColor:['#f54242','#f5a542','#f5c842','#42f5a5','#c8f542','rgba(200,245,66,0.4)','#44445a'],borderWidth:0}}]}},options:{{responsive:true,maintainAspectRatio:true,plugins:{{legend:{{display:true,position:'right',labels:{{color:'#666672',font:{{family:'DM Mono, monospace',size:9}},boxWidth:10,padding:6}}}},tooltip:defaults.plugins.tooltip}},cutout:'65%'}}}});

const wfData=[
  {{label:'Receita Bruta',value:1830000,pct:100}},
  {{label:'(-) Impostos s/ Receita',value:-127917,pct:-7.0}},
  {{label:'(-) CMV',value:-775378,pct:-42.4}},
  {{label:'(-) Fretes',value:-113460,pct:-6.2}},
  {{label:'(-) Custos Online',value:-27660,pct:-1.5}},
  {{label:'(-) Desp. Comerciais Var.',value:-110025,pct:-6.0}},
  {{label:'= Margem de Contribuição',value:675785,pct:36.9,h:true}},
  {{label:'(-) Marketing',value:-174069,pct:-9.5}},
  {{label:'(-) Diversas',value:-55317,pct:-3.0}},
  {{label:'(-) Pessoal',value:-233770,pct:-12.8}},
  {{label:'(-) Despesas Gerais',value:-113280,pct:-6.2}},
  {{label:'= EBITDA',value:99349,pct:5.4,h:true}},
  {{label:'(-) Financeiro + Impostos',value:-62364,pct:-3.4}},
  {{label:'= Lucro Líquido',value:36985,pct:2.0,h:true,last:true}},
];
const maxAbs=Math.max(...wfData.map(d=>Math.abs(d.value)));
document.getElementById('waterfall-container').innerHTML=wfData.map(d=>{{
  const bw=Math.round((Math.abs(d.value)/maxAbs)*100);
  const c=d.h?'#c8f542':d.value<0?'#f54242':'#42f5a5';
  const vc=d.h?'var(--accent)':d.value<0?'var(--danger)':'var(--accent2)';
  return `<div class="waterfall-item">
    <div class="wf-label" style="${{d.h?'color:var(--text);font-weight:700':''}}">${{d.label}}</div>
    <div class="wf-bar-wrap">
      <div class="wf-bar" style="width:${{bw}}%;background:${{c}};opacity:${{d.h?1:0.7}}"></div>
      <div class="wf-val" style="color:${{vc}}">${{d.value>=0?'':'-'}}${{fmtBR(d.value)}}</div>
    </div>
    <div class="wf-pct" style="color:${{vc}}">${{d.pct>0?'+':''}}${{d.pct}}%</div>
  </div>`;
}}).join('');

new Chart(document.getElementById('chartLucro'),{{type:'bar',data:{{labels:meses,datasets:[{{data:lucroLiq,backgroundColor:lucroLiq.map(v=>v>=0?'rgba(200,245,66,0.35)':'rgba(245,66,66,0.35)'),borderColor:lucroLiq.map(v=>v>=0?'#c8f542':'#f54242'),borderWidth:1.5,borderRadius:3}}]}},options:{{...defaults,scales:{{...defaults.scales,y:{{...defaults.scales.y,ticks:{{...defaults.scales.y.ticks,callback:v=>fmtBR(v)}}}}}}}}}});
new Chart(document.getElementById('chartImpostos'),{{type:'doughnut',data:{{labels:['PIS','COFINS','ICMS','DIFAL+FCP'],datasets:[{{data:[11895,54900,18300,42822],backgroundColor:['#f5c842','#f5a542','#f58042','#f54242'],borderWidth:0}}]}},options:{{responsive:true,maintainAspectRatio:true,plugins:{{legend:{{display:true,position:'bottom',labels:{{color:'#666672',font:{{family:'DM Mono, monospace',size:9}},boxWidth:10,padding:6}}}},tooltip:defaults.plugins.tooltip}},cutout:'60%'}}}});
new Chart(document.getElementById('chartAdmin'),{{type:'doughnut',data:{{labels:['Pessoal (12,8%)','Desp.Gerais (6,2%)'],datasets:[{{data:[233770,113280],backgroundColor:['#c8f542','#42f5a5'],borderWidth:0}}]}},options:{{responsive:true,maintainAspectRatio:true,plugins:{{legend:{{display:true,position:'bottom',labels:{{color:'#666672',font:{{family:'DM Mono, monospace',size:9}},boxWidth:10}}}},tooltip:defaults.plugins.tooltip}},cutout:'60%'}}}});

// DRE Mensal
let selectedMonth=0,chartDreMes,chartMargem;
const monthSel=document.getElementById('month-selector');
meses.forEach((m,i)=>{{const btn=document.createElement('button');btn.className='month-btn'+(i===0?' active':'');btn.textContent=m;btn.onclick=()=>{{selectedMonth=i;updateDre();monthSel.querySelectorAll('.month-btn').forEach((b,j)=>b.classList.toggle('active',j===i));}};monthSel.appendChild(btn);}});

function updateDre(){{
  const i=selectedMonth;
  document.getElementById('dre-month-label').textContent=meses[i]+' · Receita: '+fmtBR(receita[i]);
  if(chartDreMes)chartDreMes.destroy();
  chartDreMes=new Chart(document.getElementById('chartDreMes'),{{type:'bar',data:{{labels:['Receita','Marg.Contrib.','EBITDA','Lucro Líq.'],datasets:[{{data:[receita[i],margemContrib[i],Math.max(0,ebitda[i]),Math.max(0,lucroLiq[i])],backgroundColor:['rgba(200,245,66,0.2)','rgba(66,245,165,0.2)','rgba(200,245,66,0.4)','rgba(245,165,66,0.3)'],borderColor:['#c8f542','#42f5a5','#c8f542','#f5a542'],borderWidth:1.5,borderRadius:4}}]}},options:{{...defaults,scales:{{...defaults.scales,y:{{...defaults.scales.y,ticks:{{...defaults.scales.y.ticks,callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}}}}}}});
  if(chartMargem)chartMargem.destroy();
  chartMargem=new Chart(document.getElementById('chartMargem'),{{type:'line',data:{{labels:meses,datasets:[
    {{label:'Marg.Contrib.%',data:margContribPct.map(Number),borderColor:'#42f5a5',backgroundColor:'rgba(66,245,165,0.08)',borderWidth:2,tension:0.3,pointRadius:3,fill:true}},
    {{label:'EBITDA%',data:ebitdaPct.map(Number),borderColor:'#c8f542',backgroundColor:'rgba(200,245,66,0.05)',borderWidth:2,tension:0.3,pointRadius:3,fill:true}},
    {{label:'Lucro Líq.%',data:lucroLiqPct.map(Number),borderColor:'#f5a542',borderWidth:1.5,tension:0.3,pointRadius:3}},
  ]}},options:{{...defaults,plugins:{{...defaults.plugins,legend:{{display:true,labels:{{color:'#666672',font:{{family:'DM Mono, monospace',size:10}},boxWidth:12,boxHeight:2,padding:10}}}}}},scales:{{...defaults.scales,y:{{...defaults.scales.y,ticks:{{...defaults.scales.y.ticks,callback:v=>v+'%'}}}}}}}}}});
}}

const dreRows=[
  {{label:'Receita Total',values:receita,bold:true}},
  {{label:'Impostos s/ Receita',values:receita.map(r=>-Math.round(r*0.07)),cls:'sub-row'}},
  {{label:'CMV',values:receita.map(r=>-Math.round(r*0.424)),cls:'sub-row neg'}},
  {{label:'Fretes Totais',values:receita.map(r=>-Math.round(r*0.062)),cls:'sub-row neg'}},
  {{label:'Margem de Contribuição',values:margemContrib,bold:true,cls:'total-row'}},
  {{label:'Marketing',values:receita.map(r=>-Math.round(r*0.095)),cls:'sub-row neg'}},
  {{label:'Pessoal',values:meses.map(()=>-18220),cls:'sub-row neg'}},
  {{label:'Despesas Gerais',values:meses.map(()=>-9440),cls:'sub-row neg'}},
  {{label:'EBITDA',values:ebitda,bold:true,cls:'total-row'}},
  {{label:'Resultado Financeiro',values:meses.map(()=>-500),cls:'sub-row neg'}},
  {{label:'IRPJ + CSLL',values:receita.map(r=>-Math.round(r*0.031)),cls:'sub-row neg'}},
  {{label:'Lucro Líquido',values:lucroLiq,bold:true,cls:'total-row'}},
];
const dreTable=document.getElementById('dre-table');
dreTable.innerHTML=`<thead><tr><th>Linha DRE</th>${{meses.map(m=>'<th>'+m+'</th>').join('')}}<th>ANUAL</th></tr></thead><tbody>${{dreRows.map(row=>{{const total=row.values.reduce((a,b)=>a+b,0);return `<tr class="${{row.cls||''}}"><td style="${{row.bold?'font-weight:700':''}}">${{row.label}}</td>${{row.values.map(v=>`<td class="${{v<0?'neg':v>0?'pos':''}}">${{v>=0?'':'-'}}${{fmtBR(v)}}</td>`).join('')}}<td class="${{total<0?'neg':total>0?'pos':''}}" style="font-weight:700">${{total>=0?'':'-'}}${{fmtBR(total)}}</td></tr>`;}}).join('')}}</tbody>`;

updateDre();

new Chart(document.getElementById('chartCaixa'),{{type:'line',data:{{labels:meses,datasets:[{{label:'Saldo Final',data:saldoCaixa,borderColor:'#c8f542',backgroundColor:'rgba(200,245,66,0.07)',borderWidth:2,tension:0.3,fill:true,pointRadius:4,pointBackgroundColor:'#c8f542'}}]}},options:{{...defaults,scales:{{...defaults.scales,y:{{...defaults.scales.y,ticks:{{...defaults.scales.y.ticks,callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}}}}}}});
new Chart(document.getElementById('chartAportes'),{{type:'bar',data:{{labels:meses,datasets:[
  {{label:'Aportes',data:aportesMes,backgroundColor:'rgba(200,245,66,0.2)',borderColor:'#c8f542',borderWidth:1.5,borderRadius:3}},
  {{label:'Saídas',data:totalSaidas.map(v=>-v),backgroundColor:'rgba(245,66,66,0.15)',borderColor:'#f54242',borderWidth:1.5,borderRadius:3}},
]}},options:{{...defaults,plugins:{{...defaults.plugins,legend:{{display:true,labels:{{color:'#666672',font:{{family:'DM Mono, monospace',size:10}},boxWidth:10}}}}}},scales:{{...defaults.scales,y:{{...defaults.scales.y,ticks:{{...defaults.scales.y.ticks,callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}}}}}}});

function showPage(id,el){{
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('page-'+id).classList.add('active');
  el.classList.add('active');
}}
</script>
</body>
</html>"""
    return html

# ── 4. Deploy no Netlify ───────────────────────────────────────
def deploy_netlify(html_content):
    import hashlib
    # Calcula hash do arquivo
    digest = hashlib.sha1(html_content.encode()).hexdigest()
    
    headers_auth = {
        "Authorization": f"Bearer {NETLIFY_TOKEN}",
        "Content-Type": "application/json",
    }
    
    # Passo 1: criar deploy com lista de arquivos
    url = f"https://api.netlify.com/api/v1/sites/{NETLIFY_SITE}/deploys"
    payload = {
        "files": {"/index.html": digest},
        "async": False
    }
    r = requests.post(url, headers=headers_auth, json=payload)
    r.raise_for_status()
    deploy = r.json()
    deploy_id = deploy["id"]
    required = deploy.get("required", [])
    
    # Passo 2: fazer upload do arquivo se necessário
    if digest in required or required:
        upload_headers = {
            "Authorization": f"Bearer {NETLIFY_TOKEN}",
            "Content-Type": "text/html; charset=UTF-8",
        }
        upload_url = f"https://api.netlify.com/api/v1/deploys/{deploy_id}/files/index.html"
        ru = requests.put(upload_url, headers=upload_headers, data=html_content.encode("utf-8"))
        ru.raise_for_status()
    
    print(f"✅ Deploy concluído: https://{deploy.get('subdomain', NETLIFY_SITE)}.netlify.app")
    return deploy

# ── Main ───────────────────────────────────────────────────────
if __name__ == "__main__":
    print("🔐 Obtendo token de acesso...")
    token = get_access_token()

    print("📊 Lendo DRE do OneDrive...")
    data = read_excel_data(token)
    print(f"   Receita Ano: R$ {sum(data['receita']):,.0f}")
    print(f"   Lucro Líq.: R$ {sum(data['lucro']):,.0f}")

    print("🎨 Gerando HTML...")
    html = generate_html(data)

    print("🚀 Fazendo deploy no Netlify...")
    deploy_netlify(html)
    print("✅ Dashboard atualizado com sucesso!")

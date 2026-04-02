"""
=============================================================================
CALCULADOR DE PRODUTIVIDADE LOGÍSTICA — 13 HORAS
=============================================================================
Sistema: Fusion
Referência: Relatório de Entregas exportado do Fusion (.xlsm)
Autor: Gerado via Claude (Anthropic)

LÓGICA EM CASCATA (3 etapas obrigatórias):
  Etapa 1 → Rotas válidas:     Medição Produtividade Final = SIM
  Etapa 2 → Processo correto:  Check-In = SIM  E  Check-Out = SIM
  Etapa 3 → No prazo:          Hora Chegada Cliente ≤ 13h (Resultado Produtividade = 1)

  % Produtividade = Etapa 3 (numerador) ÷ Etapa 2 (denominador)

ARQUIVOS NECESSÁRIOS (mesma pasta do script):
  - Arquivo .xlsm exportado do Fusion (nome configurável abaixo)
  - Aba "Relatório"  → base de entregas
  - Aba "Veiculos"   → de/para Placa → Transportador

SAÍDA:
  - dashboard_produtividade.html  → dashboard interativo completo
  - calculos_produtividade.xlsx   → planilha com todos os cálculos

=============================================================================
COMO USAR EM UMA NOVA CONVERSA COM O CLAUDE:
  1. Exporte o relatório do Fusion no mesmo formato .xlsm
  2. Atualize a variável ARQUIVO_EXCEL abaixo com o novo nome
  3. Cole este script no Claude e peça para executar
  4. O dashboard e a planilha serão gerados automaticamente
=============================================================================
"""

import pandas as pd
import json
from datetime import date
import warnings
warnings.filterwarnings('ignore')

# ─── CONFIGURAÇÕES ────────────────────────────────────────────────────────────

ARQUIVO_EXCEL = "Cópia_de_Produtividade_13_Horas_Fusion.xlsm"  # ← altere aqui
META           = 60       # meta de produtividade em %
MIN_ENTREGAS   = 10       # mínimo de entregas para entrar nos rankings
MES_REFERENCIA = "Março/2026"  # ← altere aqui

# ─── LEITURA DOS DADOS ───────────────────────────────────────────────────────

print(f"[1/6] Lendo arquivo: {ARQUIVO_EXCEL}")
df   = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Relatório', header=0)
veic = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Veiculos',  header=0)
df.columns   = df.columns.str.strip()
veic.columns = veic.columns.str.strip()

total_base = len(df)
print(f"      Base total: {total_base} registros")

# ─── FUNIL 3 ETAPAS ──────────────────────────────────────────────────────────

print("[2/6] Aplicando funil em cascata...")

# ETAPA 1 — Rotas válidas
etapa1 = df[df['Medição Produtividade Final'] == 'SIM'].copy()

# ETAPA 2 — Processo completo (Check-In E Check-Out)
etapa2 = etapa1[
    (etapa1['Check-In']  == 'SIM') &
    (etapa1['Check-Out'] == 'SIM')
].copy()

# ETAPA 3 — Chegou até as 13h
etapa3 = etapa2[etapa2['Resultado Produtividade'] == 1].copy()

# Resultados gerais
denominador  = len(etapa2)
numerador    = len(etapa3)
pct_geral    = round(numerador / denominador * 100, 1)
excl_rota    = total_base - len(etapa1)
excl_check   = len(etapa1) - len(etapa2)
excl_prazo   = len(etapa2) - len(etapa3)

print(f"      Etapa 1 (rota válida):    {len(etapa1):>6} | excluídos: {excl_rota}")
print(f"      Etapa 2 (check completo): {len(etapa2):>6} | excluídos: {excl_check}")
print(f"      Etapa 3 (até 13h):        {len(etapa3):>6} | excluídos: {excl_prazo}")
print(f"      % PRODUTIVIDADE GERAL:    {pct_geral}%  (meta: {META}%)")

# ─── FUNÇÕES AUXILIARES ──────────────────────────────────────────────────────

def calcular_grupo(df_grupo, col_grupo):
    """Calcula produtividade por grupo (gestor, rota, motorista, etc.)"""
    return df_grupo.groupby(col_grupo).apply(lambda x: pd.Series({
        'total': len(x),
        'sim':   int((x['Resultado Produtividade'] == 1).sum()),
        'nao':   int((x['Resultado Produtividade'] != 1).sum()),
        'pct':   round((x['Resultado Produtividade'] == 1).sum() / len(x) * 100, 1)
    })).reset_index()

def cor_pct(p):
    """Retorna cor semântica baseada na % de produtividade"""
    if p < META:       return 'vermelho'
    if p < META + 15:  return 'amarelo'
    return 'verde'

# ─── CÁLCULOS POR DIMENSÃO ───────────────────────────────────────────────────

print("[3/6] Calculando por dimensão...")

# Por Gestor
gestores = calcular_grupo(etapa2, 'Gestor').sort_values('pct', ascending=False)

# Por Data
etapa2['dt'] = pd.to_datetime(etapa2['Data Saída'], format='%d-%m-%Y', errors='coerce')
diario = etapa2.groupby(etapa2['dt'].dt.date).apply(lambda x: pd.Series({
    'total': len(x),
    'sim':   int((x['Resultado Produtividade'] == 1).sum()),
    'pct':   round((x['Resultado Produtividade'] == 1).sum() / len(x) * 100, 1)
})).reset_index()
diario.columns = ['data', 'total', 'sim', 'pct']
diario['data_fmt'] = pd.to_datetime(diario['data']).dt.strftime('%d/%m')
diario = diario.sort_values('data').dropna(subset=['data'])

# Por Rota
rotas = calcular_grupo(etapa2, 'Rota')
rotas = rotas[rotas['total'] >= MIN_ENTREGAS].sort_values('pct')

# Por Motorista
motoristas = calcular_grupo(etapa2, 'Motorista')
motoristas = motoristas[motoristas['total'] >= MIN_ENTREGAS]
mot_piores  = motoristas.sort_values('pct').head(15)
mot_melhores = motoristas.sort_values('pct', ascending=False).head(15)

# Por Transportador
etapa2_t = etapa2.merge(veic, on='Placa', how='left')
transp = calcular_grupo(etapa2_t, 'Nome fantasia')
transp = transp[transp['total'] >= 5].sort_values('pct')

# Sem check completo
sem_check = etapa1[
    ~((etapa1['Check-In'] == 'SIM') & (etapa1['Check-Out'] == 'SIM'))
]
sem_check_mot = sem_check.groupby('Motorista').apply(lambda x: pd.Series({
    'sem_check': len(x),
    'rotas':     ', '.join(x['Rota'].dropna().unique()[:3]),
    'status':    x['Status da Entrega'].value_counts().index[0] if len(x) > 0 else ''
})).reset_index().sort_values('sem_check', ascending=False)

# Programa de Excelência
excel_ouro  = mot_melhores[mot_melhores['pct'] >= 90].copy()
excel_prata = mot_melhores[(mot_melhores['pct'] >= 85) & (mot_melhores['pct'] < 90)].copy()
excel_bronze = mot_melhores[(mot_melhores['pct'] >= 80) & (mot_melhores['pct'] < 85)].copy()

print(f"      Gestores analisados:    {len(gestores)}")
print(f"      Dias analisados:        {len(diario)}")
print(f"      Rotas analisadas:       {len(rotas)}")
print(f"      Motoristas analisados:  {len(motoristas)}")
print(f"      Transportadores:        {len(transp)}")
print(f"      Sem check completo:     {len(sem_check_mot)} motoristas / {len(sem_check)} entregas")
print(f"      Excelência Ouro (≥90%): {len(excel_ouro)}")
print(f"      Excelência Prata (85-90%): {len(excel_prata)}")

# ─── EXPORTAR PLANILHA ───────────────────────────────────────────────────────

print("[4/6] Exportando planilha de cálculos...")

with pd.ExcelWriter('calculos_produtividade.xlsx', engine='openpyxl') as writer:

    # Aba Resumo
    resumo = pd.DataFrame({
        'Indicador': [
            'Período de referência',
            'Meta de produtividade',
            'Base total de registros',
            'Etapa 1 — Rotas válidas (Medição Final = SIM)',
            'Etapa 2 — Com Check-In E Check-Out',
            'Etapa 3 — Chegaram até 13h (numerador)',
            '% Produtividade Geral',
            'Excluídos por rota não produtividade',
            'Excluídos por ausência de check',
            'Chegaram após 13h',
            'Data de geração'
        ],
        'Valor': [
            MES_REFERENCIA,
            f'{META}%',
            total_base,
            len(etapa1),
            denominador,
            numerador,
            f'{pct_geral}%',
            excl_rota,
            excl_check,
            excl_prazo,
            date.today().strftime('%d/%m/%Y')
        ]
    })
    resumo.to_excel(writer, sheet_name='Resumo', index=False)

    # Aba Gestores
    gestores.to_excel(writer, sheet_name='Por Gestor', index=False)

    # Aba Diário
    diario[['data_fmt', 'total', 'sim', 'pct']].rename(
        columns={'data_fmt':'Data','total':'Total','sim':'SIM','pct':'% Prod'}
    ).to_excel(writer, sheet_name='Diário', index=False)

    # Aba Rotas
    rotas.to_excel(writer, sheet_name='Por Rota', index=False)

    # Aba Motoristas — Piores
    mot_piores.to_excel(writer, sheet_name='Motoristas Críticos', index=False)

    # Aba Motoristas — Melhores
    mot_melhores.to_excel(writer, sheet_name='Motoristas Destaque', index=False)

    # Aba Transportadores
    transp.to_excel(writer, sheet_name='Transportadores', index=False)

    # Aba Sem Check
    sem_check_mot.to_excel(writer, sheet_name='Sem Check', index=False)

    # Aba Excelência
    excel_todos = pd.concat([
        excel_ouro.assign(tier='Ouro ≥ 90%'),
        excel_prata.assign(tier='Prata 85-90%'),
        excel_bronze.assign(tier='Bronze 80-85%')
    ], ignore_index=True)
    excel_todos.to_excel(writer, sheet_name='Programa Excelência', index=False)

print("      calculos_produtividade.xlsx gerado!")

# ─── MONTAR JSON PARA O DASHBOARD ────────────────────────────────────────────

print("[5/6] Preparando dados para o dashboard...")

dados_dashboard = {
    'meta': META,
    'mes': MES_REFERENCIA,
    'funil': {
        'base_total': total_base,
        'etapa1': len(etapa1),
        'etapa2': denominador,
        'etapa3': numerador,
        'excl_rota': excl_rota,
        'excl_check': excl_check,
        'excl_prazo': excl_prazo,
        'pct_geral': pct_geral
    },
    'gestores': gestores.to_dict('records'),
    'diario': diario[['data_fmt','total','sim','pct']].rename(
        columns={'data_fmt':'d'}).to_dict('records'),
    'rotas_piores': rotas.head(10).to_dict('records'),
    'rotas_melhores': rotas.sort_values('pct',ascending=False).head(5).to_dict('records'),
    'motoristas_piores': mot_piores.to_dict('records'),
    'motoristas_melhores': mot_melhores.head(10).to_dict('records'),
    'transportadores': transp.to_dict('records'),
    'sem_check': sem_check_mot.to_dict('records'),
    'excelencia': {
        'ouro': excel_ouro.to_dict('records'),
        'prata': excel_prata.to_dict('records'),
        'bronze': excel_bronze.to_dict('records')
    }
}

# Salvar JSON (referência para o dashboard)
with open('dados_dashboard.json', 'w', encoding='utf-8') as f:
    json.dump(dados_dashboard, f, ensure_ascii=False, indent=2, default=str)

print("      dados_dashboard.json gerado!")
print("[6/6] Concluído!")
print(f"\n{'='*50}")
print(f"RESULTADO FINAL: {pct_geral}% de produtividade")
print(f"Meta: {META}% | Diferença: +{round(pct_geral-META,1)}pp")
print(f"{'='*50}")
print(f"\nArquivos gerados:")
print(f"  calculos_produtividade.xlsx")
print(f"  dados_dashboard.json")
print(f"\nPróximo passo: gerar o dashboard HTML com o Claude")
print(f"usando os dados de dados_dashboard.json")

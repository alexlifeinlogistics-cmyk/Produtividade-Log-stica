# 📦 Produtividade Logística — 13 Horas

Sistema de análise de produtividade de entregas baseado no relatório exportado do sistema **Fusion**.

---

## 📋 O que este projeto faz

Calcula a **% de produtividade logística** — proporção de entregas realizadas até as 13h — usando uma lógica em cascata de 3 etapas:

```
Base total do Fusion
    ↓
Etapa 1: Rotas válidas       → Medição Produtividade Final = SIM
    ↓
Etapa 2: Processo correto    → Check-In = SIM  E  Check-Out = SIM
    ↓
Etapa 3: No prazo            → Hora Chegada Cliente ≤ 13h

% Produtividade = Etapa 3 ÷ Etapa 2
```

---

## 📁 Estrutura do projeto

```
produtividade-logistica/
│
├── calcular_produtividade.py   # Script principal de cálculo
├── dashboard_produtividade.html # Dashboard interativo (gerado)
├── README.md                   # Este arquivo
│
└── dados/
    └── [arquivo .xlsm do Fusion]  # ← colocar aqui
```

---

## ⚙️ Como usar

### 1. Requisitos

```bash
pip install pandas openpyxl
```

### 2. Exportar o relatório do Fusion

- Acesse o Fusion → Relatório de Entregas
- Exporte no formato **.xlsm**
- Certifique-se de que o arquivo contém as abas:
  - `Relatório` — base de entregas
  - `Veiculos` — de/para Placa → Transportador

### 3. Configurar o script

Abra `calcular_produtividade.py` e edite as variáveis no topo:

```python
ARQUIVO_EXCEL  = "nome_do_arquivo.xlsm"   # nome do arquivo exportado
META           = 60                        # meta de produtividade em %
MIN_ENTREGAS   = 10                        # mínimo de entregas para rankings
MES_REFERENCIA = "Abril/2026"             # mês de referência
```

### 4. Executar

```bash
python calcular_produtividade.py
```

### 5. Resultado

O script gera dois arquivos:
- `calculos_produtividade.xlsx` — planilha com todos os cálculos por dimensão
- `dados_dashboard.json` — dados estruturados para o dashboard

---

## 📊 Colunas obrigatórias no Relatório Fusion

| Coluna | Uso |
|--------|-----|
| `Medição Produtividade Final` | SIM/NÃO — define quais rotas entram no cálculo |
| `Check-In` | SIM/NÃO — confirma início do processo |
| `Check-Out` | SIM/NÃO — confirma fim do processo |
| `Resultado Produtividade` | 1/0 — 1 = chegou até 13h |
| `Hora Chegada Cliente` | hora numérica (ex: 11, 13, 15) |
| `Gestor` | nome do gestor responsável |
| `Motorista` | nome do motorista |
| `Rota` | nome da rota |
| `Data Saída` | data no formato dd-mm-yyyy |
| `Status da Entrega` | Entregue, Devolvido, Faturado, etc. |
| `Placa` | placa do veículo (para cruzar com aba Veiculos) |

---

## 📈 Dimensões de análise

O script calcula a produtividade por:

- **Gestor** — Marcos, Terres, Aury
- **Evolução diária** — % por data de saída
- **Rota** — ranking das piores e melhores rotas
- **Motorista** — ranking dos piores e melhores motoristas
- **Transportador** — cruzado via aba Veículos
- **Sem check** — motoristas que não registraram Check-In/Out
- **Programa de Excelência** — tiers Ouro (≥90%), Prata (85-90%), Bronze (80-85%)

---

## 🔄 Fluxo para nova análise mensal

```
1. Exportar novo .xlsm do Fusion
2. Colocar na pasta do projeto
3. Atualizar ARQUIVO_EXCEL e MES_REFERENCIA no script
4. Rodar: python calcular_produtividade.py
5. Abrir nova conversa no Claude e pedir para gerar o dashboard
   com os dados do dados_dashboard.json gerado
```

---

## ⚠️ Regras de negócio importantes

| Regra | Comportamento |
|-------|--------------|
| Rota com `Medição Final = NÃO` | **Excluída** do cálculo (nem entra no denominador) |
| Sem Check-In ou Check-Out | **Excluído do denominador** (não penaliza, mas não conta) |
| Status `Devolvido` com hora ≤ 13h | Entra como **NÃO** (resultado = 0) |
| Status `Faturado` sem hora | Entra como **NÃO** no denominador |
| Clientes sem rota (`NaN`) | Entram se `Medição Final = SIM` |

---

## 📬 Relatório Executivo

O dashboard (`dashboard_produtividade.html`) inclui uma aba **"📊 Relatório Executivo"** com layout corporativo branco, pronta para copiar e colar em e-mail para gerência e diretoria.

Conteúdo do relatório executivo:
- Resultado geral em destaque
- KPIs principais
- Gráfico de evolução diária
- Desempenho por gestor com meta
- Tabela de rotas críticas
- Tabela de motoristas críticos
- Programa de Excelência (Ouro / Prata / Bronze)
- Alerta de motoristas sem check-in/out

---

## 🛠️ Manutenção da base de rotas

A aba `Base Dados` do arquivo .xlsm define quais rotas devem ou não entrar no cálculo. Para alterar:

| Rota | `Produtividade` |
|------|----------------|
| Rota ativa | `Sim` |
| Rota desativada / noturna / excluída | `Não` |

Essa configuração é refletida no campo `Medição Produtividade Final` do Relatório.

---

*Gerado com suporte do Claude (Anthropic) · Atualizado em Abril/2026*

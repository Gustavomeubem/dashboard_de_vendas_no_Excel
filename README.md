# 🎮 Xbox Game Pass — Sales Dashboard

Dashboard interativo de análise de vendas de assinaturas do **Xbox Game Pass**, construído inteiramente em Microsoft Excel. Apresenta indicadores-chave (KPIs), gráficos dinâmicos e respostas a perguntas de negócio sobre uma base de 295 assinaturas registradas em 2024.

![Preview](docs/preview.png)

## 📋 Sobre o Projeto

Este projeto faz parte de um desafio de **análise de dados e visualização** em Excel. O objetivo é transformar uma base bruta de assinaturas em um dashboard executivo que responda perguntas de negócio reais sobre faturamento, mix de produtos, comportamento de renovação e sazonalidade de vendas.

A base contém informações de assinantes do Xbox Game Pass nos três planos (**Core**, **Standard** e **Ultimate**), com três periodicidades (**Monthly**, **Quarterly** e **Annual**), incluindo add-ons opcionais (**EA Play Season Pass** e **Minecraft Season Pass**) e cupons de desconto.

## 🗂️ Estrutura do Arquivo Excel

O workbook contém quatro abas:

| Aba | Conteúdo |
|---|---|
| **Assets** | Paleta de cores oficial Xbox (#9BC848, #22C55E, #2AE6B1, #5BF6A8) e logos usados como referência visual |
| **Bases** | Base de dados original com 295 registros e 13 colunas |
| **Cálculos** | Tabelas auxiliares com fórmulas SUMIF, SUMIFS, COUNTIF e SUMPRODUCT que alimentam o dashboard |
| **Dashboard** | Visualização final com KPIs, 4 gráficos e respostas às perguntas de negócio |

## 📊 Dados Utilizados

Cada linha da aba **Bases** representa uma assinatura, com as seguintes colunas:

- `Subscriber ID` — Identificador único do assinante
- `Name` — Nome do cliente
- `Plan` — Plano contratado (Core / Standard / Ultimate)
- `Start Date` — Data de início da assinatura
- `Auto Renewal` — Renovação automática (Yes / No)
- `Subscription Price` — Valor mensal da assinatura
- `Subscription Type` — Periodicidade (Monthly / Quarterly / Annual)
- `EA Play Season Pass` — Add-on EA Play (Yes / No)
- `EA Play Season Pass Price` — Valor do add-on EA Play
- `Minecraft Season Pass` — Add-on Minecraft (Yes / No)
- `Minecraft Season Pass Price` — Valor do add-on Minecraft
- `Coupon Value` — Desconto aplicado
- `Total Value` — Valor final da venda

**Período coberto:** Janeiro a Dezembro de 2024  
**Total de registros:** 295 assinaturas

## 📈 Perguntas de Negócio Respondidas

O dashboard responde a quatro perguntas estratégicas:

1. **Qual o faturamento total de vendas de planos anuais?**  
   → `$1.754` (soma de Total Value onde Subscription Type = Annual)

2. **Qual o faturamento de planos anuais separado por Auto Renewal?**  
   → Com renovação automática: `$1.537` | Sem renovação: `$217`

3. **Qual o total de vendas do add-on EA Play Season Pass?**  
   → `$2.940`, concentrados 100% no plano Ultimate

4. **Qual o total de vendas do add-on Minecraft Season Pass?**  
   → `$3.880` (Ultimate: $1.960 | Standard: $1.920 | Core: $0)

## 🎯 KPIs do Dashboard

- **Total Subscribers** — Número total de assinantes (295)
- **Faturamento Total** — Receita consolidada do ano ($7.633)
- **Ticket Médio** — Valor médio por assinatura ($25,87)
- **Cupons Aplicados** — Total de descontos concedidos ($2.122)
- **% Auto Renewal** — Taxa de renovação automática (50,2%)
- **Faturamento Anual (Annual)** — Resposta direta à P1 ($1.754)

## 📉 Gráficos Inclusos

1. **Subscribers por Plano** — Gráfico de colunas mostrando a distribuição de assinantes
2. **Faturamento por Plano** — Gráfico de rosca com participação % de cada plano na receita
3. **Faturamento por Tipo de Assinatura** — Gráfico de barras horizontais (Monthly / Quarterly / Annual)
4. **Evolução Mensal do Faturamento** — Gráfico de linha mostrando sazonalidade ao longo de 2024

## 🛠️ Tecnologias e Recursos

- **Microsoft Excel** (compatível com Excel 2016+ e LibreOffice Calc)
- Fórmulas nativas: `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTA`, `AVERAGE`, `SUMPRODUCT`, `MONTH`
- Formatação condicional e tema personalizado Xbox (verde sobre fundo escuro)
- Gráficos dinâmicos com referências cruzadas entre abas (`'Cálculos'!G5`, etc.)

## ▶️ Como Reproduzir

### Opção 1: Apenas visualizar
1. Baixe o arquivo `dashboard_xbox_game_pass.xlsx` deste repositório
2. Abra no Microsoft Excel (recomendado) ou LibreOffice Calc
3. Navegue até a aba **Dashboard** para ver a visualização final

### Opção 2: Reconstruir do zero
Os dados de entrada estão na aba **Bases** do arquivo `base.xlsx` (arquivo original sem dashboard).

Para reconstruir as análises:

1. Abra o arquivo `base.xlsx`
2. Na aba **Cálculos**, recrie as tabelas-resumo usando as fórmulas abaixo (referenciando `'Bases'!A2:M296`):

```excel
# KPIs Gerais
=COUNTA('Bases'!C2:C296)                    # Total de subscribers
=SUM('Bases'!M2:M296)                       # Faturamento total
=AVERAGE('Bases'!M2:M296)                   # Ticket médio
=COUNTIF('Bases'!E2:E296,"Yes")             # Auto renewal Yes

# Resposta P1
=SUMIF('Bases'!G2:G296,"Annual",'Bases'!M2:M296)

# Resposta P2
=SUMIFS('Bases'!M2:M296,'Bases'!G2:G296,"Annual",'Bases'!E2:E296,"Yes")

# Faturamento por mês (exemplo janeiro)
=SUMPRODUCT((MONTH('Bases'!D2:D296)=1)*'Bases'!M2:M296)
```

3. Na aba **Dashboard**, crie células com referências às fórmulas calculadas (ex: `='Cálculos'!G5`) e insira os gráficos apontando para as tabelas auxiliares.

### Opção 3: Reproduzir programaticamente
O dashboard foi construído com Python usando `openpyxl`. Os scripts utilizados estão disponíveis no diretório `scripts/`:

```bash
pip install openpyxl pandas
python scripts/build_dashboard.py
python scripts/build_dashboard_step2.py
python scripts/build_dashboard_step3.py
python scripts/build_dashboard_step4.py
```

## 🎨 Paleta de Cores

| Cor | Hex | Uso |
|---|---|---|
| 🟢 Xbox Green | `#9BC848` | Cor principal, KPIs e barras |
| 🟢 Verde Escuro | `#22C55E` | Cor secundária, séries de gráficos |
| 🟢 Verde Menta | `#2AE6B1` | Highlights, KPIs secundários |
| 🟢 Verde Claro | `#5BF6A8` | Menus, fatias de gráficos |
| ⚫ BG Escuro | `#0E1116` | Fundo principal |
| ⚫ BG Painel | `#1A1F26` | Fundo dos cards e gráficos |

## 📁 Estrutura do Repositório

```
.
├── dashboard_xbox_game_pass.xlsx   # Arquivo final com dashboard
├── base.xlsx                        # Base de dados original (sem dashboard)
├── README.md                        # Este arquivo
├── docs/
│   └── preview.png                  # Screenshot do dashboard
└── scripts/                         # Scripts Python usados na construção
    ├── build_dashboard.py
    ├── build_dashboard_step2.py
    ├── build_dashboard_step3.py
    └── build_dashboard_step4.py
```

## 📝 Observações

- Todos os cálculos no dashboard são **dinâmicos** (via fórmulas Excel), permitindo atualização automática caso novos registros sejam adicionados à aba **Bases**.
- O dashboard foi otimizado para impressão em formato **A3 Landscape**.
- Os dados são fictícios e foram fornecidos para fins de exercício acadêmico.

## 📄 Licença

Este projeto está disponível sob a licença MIT — sinta-se à vontade para usar, modificar e distribuir.

---

**Autor:** Seu Nome  
**Data:** Maio/2026  
**Desafio:** Criar um dashboard de vendas no Excel a partir de dados reais

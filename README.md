# Consultor Fundamentalista B3

Robô em Python que coleta indicadores fundamentalistas de ações e FIIs da B3, classifica cada ativo com regras próprias e alimenta um dashboard em Power BI.

![Dashboard](imagens/dashboard.png)

## O que ele faz

- Coleta preço, P/L, P/VP, ROE, margem líquida, EV/EBITDA e dividend yield de mais de 110 ativos.
- O dividend yield vem de três fontes em ordem de prioridade: Status Invest, histórico de dividendos do Yahoo Finance e, por último, o campo de informações do Yahoo Finance, com normalização.
- Classifica cada ativo (Forte compra, Boa compra, Qualidade, Neutro, Aguardar, Turnaround).
- Salva o resultado no Excel (abas `Dados_Atuais`, `Oportunidades` e `Historico`) e no SQL Server.
- Acumula o histórico a cada execução e envia um resumo por e-mail.
- Tem tentativas automáticas com espera crescente (retry com backoff) nas chamadas de rede.

## Regras de classificação

**Ações**

| Classificação | Regra |
|---|---|
| Forte compra | 0 < P/L < 12, DY > 5%, margem > 10%, dívida líquida/EBITDA < 2,5 e ROE > 15% |
| Boa compra | Mesmas regras, sem exigir ROE |
| Aguardar | P/L < 12 com dívida > 4, ou P/L > 22, ou EV/EBITDA > 15 |
| Turnaround | P/L menor ou igual a 0 |
| Qualidade | ROE > 20% e P/L < 18 |
| Neutro | Demais casos |

**FIIs**

| Classificação | Regra |
|---|---|
| Forte compra | DY > 10% e P/VP entre 0,70 e 1,05 |
| Boa compra | DY > 8% e P/VP até 1,20 |
| Saudável | DY > 6% |

## Dashboard

O arquivo `Consultor_B3.pbix` lê o `Painel_Consultor.xlsx`. Depois de rodar o robô, abra o Power BI e clique em **Atualizar**. O fundo da página está em `dashboard/fundo_dashboard_b3_v6.png`.

## Tecnologias

- Python (yfinance, pandas, requests, pyodbc)
- SQL Server (histórico na tabela dbo.historico_financas)
- Excel (saída do robô e fonte de dados do dashboard)
- Power BI (dashboard)

## Como rodar

pip install -r requirements.txt
cp .env.example .env    # no Windows: copy .env.example .env

# preencha o .env com seus dados
python analise.py


O SQL Server e o e-mail são opcionais para testar a coleta, mas o script espera as variáveis do `.env`.

## Estrutura


analise.py               script principal (coleta, classificação, Excel, SQL e e-mail)
Painel_Consultor.xlsx    saída do robô, usada pelo dashboard
dashboard/               arquivo .pbix e imagem de fundo
imagens/                 print do dashboard
requirements.txt         dependências
.env.example             modelo das variáveis de ambiente


## Limitações conhecidas

- A coluna de dívida líquida/EBITDA está vindo zerada na coleta atual, então as regras que dependem dela ainda não filtram nada.
- O dividend yield tem teto de 25% (`LIMITE_DY`), e vários ativos empatam nesse valor.
- Os dados vêm de fontes públicas e podem ter falhas ou atrasos.

## Aviso

Projeto de estudo. Nada aqui é recomendação de investimento.

"""
Consultor Fundamentalista B3

Busca indicadores de ações e FIIs, classifica cada ativo, salva no Excel e no
SQL Server e manda um resumo por e-mail.

Para rodar: copie o .env.example para .env, preencha e execute analise.py.
"""

import os
import re
import time
import functools
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import requests
import yfinance as yf
import pandas as pd
import pyodbc
from dotenv import load_dotenv

load_dotenv()


DIRETORIO_PROJETO = os.path.dirname(os.path.abspath(__file__))
CAMINHO_EXCEL = os.path.join(DIRETORIO_PROJETO, "Painel_Consultor.xlsx")

LIMITE_DY = 25.0
DELAY_REQ = 0.4
DIAS_RETENCAO_EXCEL = 180

RETRY_TENTATIVAS = 3
RETRY_DELAY_INICIAL = 1.0
RETRY_FATOR_BACKOFF = 2.0

EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE", "")
EMAIL_SENHA_APP = os.getenv("EMAIL_SENHA_APP", "")
EMAIL_DESTINATARIO = os.getenv("EMAIL_DESTINATARIO", EMAIL_REMETENTE)

SQL_SERVER = os.getenv("SQL_SERVER", r"localhost\SQLEXPRESS")
SQL_DATABASE = os.getenv("SQL_DATABASE", "FinancasDB")
CONN_STR = (
    "Driver={SQL Server};"
    f"Server={SQL_SERVER};"
    f"Database={SQL_DATABASE};"
    "Trusted_Connection=yes;"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9",
}

UNITS_NAO_FII = {"SANB11", "BPAC11", "KLBN11", "IGTI11", "ALUP11", "ENGI11"}

LISTA_ATIVOS = [
    "ABEV3",  "ALOS3",  "AURE3",  "B3SA3",  "BBAS3",
    "BBDC3",  "BBDC4",  "BBSE3",  "BEEF3",  "BRAP4",
    "BRKM5",  "CMIG3",  "CMIG4",  "COGN3",  "CSMG3",
    "CSNA3",  "CXSE3",  "CYRE3",  "DIRR3",  "EGIE3",
    "EZTC3",  "FLRY3",  "GGBR4",  "GOAU4",  "ITSA4",
    "ITUB3",  "ITUB4",  "JHSF3",  "LAVV3",  "MDIA3",
    "MEAL3",  "MGLU3",  "MOVI3",  "MRVE3",  "OIBR3",
    "PETR3",  "PETR4",  "PLPL3",  "PRIO3",  "PSSA3",
    "RAIZ4",  "RANI3",  "RAPT4",  "RENT3",  "SANB11",
    "SAPR4",  "SBFG3",  "SBSP3",  "SIMH3",  "SLCE3",
    "SMTO3",  "SUZB3",  "TIMS3",  "UGPA3",  "USIM5",
    "VALE3",  "VIVA3",  "WEGE3",  "YDUQ3",
    # Removidos por 404 (Quote not found) em 03/09/2026: ELET3, ELET6, NTCO3,
    # CPLE6, CIEL3, RRRP3, PETZ3, ARZZ3, AZUL4, GOLL4, CCRO3, ODPV3, PARD3,
    # BRFS3, JBSS3, MRFG3, NEOE3, STBP3, TRPL4.
    "RADL3",  "VBBR3",  "TOTS3",
    "LREN3",  "HAPV3",  "ASAI3",  "CPFE3",
    "ENEV3",  "EQTL3",  "CVCB3",
    "VAMO3",  "POMO4",  "TUPY3",
    "RAIL3",
    "MULT3",  "CSAN3",  "VIVT3",  "GRND3",
    "LWSA3",  "ANIM3",  "SEER3",  "VULC3",  "HYPE3",
    "RDOR3",  "QUAL3",
    "AGRO3",  "PCAR3",
    "BPAC11", "KLBN11", "IGTI11", "ALUP11", "ENGI11",
    # HGPO11 removido: "Preço inválido ou zerado".
    "ALZR11", "BRCR11", "BTLG11", "CPTS11", "DEVA11",
    "HCTR11", "HGLG11", "HGRU11", "IRDM11",
    "KNCR11", "KNIP11", "KNRI11", "MXRF11", "PORD11",
    "RECR11", "SNCI11", "TAEE11", "TGAR11", "TORD11",
    "TRXF11", "VGHF11", "VINO11", "VISC11", "VRTA11",
    "XPML11",
]


def safe_float(value, default: float = 0.0) -> float:
    try:
        result = float(value)
        return result if pd.notna(result) else default
    except (TypeError, ValueError):
        return default


def normalizar_dy(valor_bruto) -> float:
    dy = safe_float(valor_bruto)
    if dy <= 0:
        return 0.0
    while dy > LIMITE_DY:
        dy /= 100
    if dy < 0.25:
        dy *= 100
    return round(dy, 2)


def normalizar_percentual(valor_bruto) -> float:
    v = safe_float(valor_bruto)
    if v == 0:
        return 0.0
    if abs(v) > 100:
        v /= 100
    elif abs(v) < 1:
        v *= 100
    return round(v, 2)


def get_status(is_fii, pl, dy_pct, margem, divida, ev_ebitda, p_vpa, roe) -> str:
    if is_fii:
        if dy_pct > 10 and 0.70 <= p_vpa <= 1.05:
            return "FORTE COMPRA (FII Barato + Alto Rendimento)"
        # p_vpa 0 quer dizer que não veio dado, então não pode contar como preço justo
        elif dy_pct > 8 and 0 < p_vpa <= 1.20:
            return "BOA COMPRA (FII Rendimento alto e preco justo)"
        elif dy_pct > 6:
            return "FII Saudavel (Rendimento em linha com mercado)"
        elif dy_pct == 0:
            return "FII sem rendimento informado"
        else:
            return "FII com Baixo Rendimento"
    else:
        if 0 < pl < 12 and dy_pct > 5 and margem > 10 and divida < 2.5 and roe > 15:
            return "FORTE COMPRA (Barata + Dividendos + ROE alto)"
        elif 0 < pl < 12 and dy_pct > 5 and margem > 10 and divida < 2.5:
            return "BOA COMPRA (Barata e Paga Dividendos)"
        elif 0 < pl < 12 and divida > 4.0:
            return "AGUARDAR (Barata mas Muito Endividada)"
        elif pl > 22 or ev_ebitda > 15:
            return "AGUARDAR (Preco Esticado)"
        elif pl <= 0:
            return "TURNAROUND / Sem Lucro Recente"
        elif roe > 20 and pl < 18:
            return "QUALIDADE (ROE alto, preco razoavel)"
        else:
            return "NEUTRO (Acompanhar evolucao)"


def retry_com_backoff(tentativas=RETRY_TENTATIVAS, delay_inicial=RETRY_DELAY_INICIAL,
                       fator=RETRY_FATOR_BACKOFF, excecoes=(Exception,)):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            delay = delay_inicial
            ultimo_erro = None
            for tentativa in range(1, tentativas + 1):
                try:
                    return func(*args, **kwargs)
                except excecoes as e:
                    ultimo_erro = e
                    if tentativa < tentativas:
                        time.sleep(delay)
                        delay *= fator
            raise ultimo_erro
        return wrapper
    return decorator


_cache_dy: dict = {}


@retry_com_backoff(excecoes=(requests.exceptions.RequestException,))
def _requisitar_status_invest(url: str) -> requests.Response:
    return requests.get(url, headers=HEADERS, timeout=8)


def dy_status_invest(ticker: str) -> float | None:
    if ticker in _cache_dy:
        return _cache_dy[ticker]
    try:
        secao = "fundos-imobiliarios" if ticker.endswith("11") else "acoes"
        url = f"https://statusinvest.com.br/{secao}/{ticker.lower()}"
        resp = _requisitar_status_invest(url)
        if resp.status_code != 200:
            return None

        padrao = re.search(r'"dy"\s*:\s*"?([\d]+[,.][\d]+)"?', resp.text, re.IGNORECASE)
        if not padrao:
            padrao = re.search(
                r'class="[^"]*dy[^"]*"[^>]*>[\s\S]{0,200}?([\d]+[,.][\d]+)\s*%',
                resp.text, re.IGNORECASE,
            )
        if padrao:
            valor = float(padrao.group(1).replace(",", "."))
            if 0 < valor <= LIMITE_DY:
                _cache_dy[ticker] = valor
                return valor
    except Exception:
        pass
    return None


@retry_com_backoff(excecoes=(Exception,))
def _obter_info_yf(ticker: str) -> dict:
    return yf.Ticker(f"{ticker}.SA").info


@retry_com_backoff(excecoes=(Exception,))
def _obter_dividendos_yf(ticker: str):
    return yf.Ticker(f"{ticker}.SA").dividends


def dy_via_historico_yf(ticker: str, preco_atual: float) -> float:
    try:
        hist = _obter_dividendos_yf(ticker)
        if hist.empty or preco_atual <= 0:
            return 0.0

        cutoff = pd.Timestamp.now(tz="America/Sao_Paulo") - pd.DateOffset(months=12)
        hist.index = (
            hist.index.tz_localize("America/Sao_Paulo")
            if hist.index.tz is None
            else hist.index.tz_convert("America/Sao_Paulo")
        )

        dividendos_12m = hist[hist.index >= cutoff].sum()
        if dividendos_12m <= 0:
            return 0.0

        return round(min((dividendos_12m / preco_atual) * 100, LIMITE_DY), 2)
    except Exception:
        return 0.0


def obter_dy(ticker: str, preco_atual: float, info: dict) -> tuple[float, str]:
    dy = dy_status_invest(ticker)
    if dy and dy > 0:
        return dy, "StatusInvest"

    dy = dy_via_historico_yf(ticker, preco_atual)
    if dy > 0:
        return dy, "yf-historico"

    return normalizar_dy(info.get("dividendYield")), "yf-info"


def coletar_ativo(ticker: str) -> dict:
    info = _obter_info_yf(ticker)

    if not info or not info.get("longName"):
        raise ValueError("Sem dados retornados pela API.")

    moeda = info.get("currency", "BRL")
    if moeda != "BRL":
        raise ValueError(f"Moeda {moeda} - pode ser ADR ou ticker errado.")

    preco = safe_float(info.get("currentPrice") or info.get("regularMarketPrice"))
    if preco <= 0:
        raise ValueError("Preco invalido ou zerado.")

    is_fii = ticker.endswith("11") and ticker not in UNITS_NAO_FII
    tipo = "FII" if is_fii else "Acao"

    nome = info.get("longName", ticker)
    pl = safe_float(info.get("trailingPE"))
    margem = normalizar_percentual(info.get("profitMargins"))
    # TODO: vem sempre 0, esse campo parece não existir no retorno do yfinance
    divida = safe_float(info.get("debtToEbitda"))
    ev_ebitda = safe_float(info.get("enterpriseToEbitda"))
    p_vpa = safe_float(info.get("priceToBook"))
    roe = normalizar_percentual(info.get("returnOnEquity"))

    dy_pct, fonte_dy = obter_dy(ticker, preco, info)
    status = get_status(is_fii, pl, dy_pct, margem, divida, ev_ebitda, p_vpa, roe)

    return {
        "Ticker": ticker,
        "Tipo_Ativo": tipo,
        "Empresa": nome,
        "Preco_Atual": preco,
        "P_L": pl,
        "Dividend_Yield_Porcento": dy_pct,
        "Fonte_DY": fonte_dy,
        "Status_Consultor": status,
        "Margem_Liquida_Porcento": margem,
        "Divida_Liquida_EBITDA": divida,
        "EV_EBITDA": ev_ebitda,
        "P_VPA": p_vpa,
        "ROE_Porcento": roe,
    }


def coletar_todos(lista_ativos: list[str], data_hoje: str) -> tuple[list[dict], list[dict], dict]:
    dados_consolidados = []
    erros = []
    contagem_fontes = {"StatusInvest": 0, "yf-historico": 0, "yf-info": 0}

    print(f"\n{'='*80}")
    print(f"  CONSULTOR FUNDAMENTALISTA B3 - {data_hoje}")
    print(f"  Ativos na fila : {len(lista_ativos)}")
    print(f"  Fonte DY       : Status Invest -> yf-historico -> yf-info")
    print(f"{'='*80}\n")

    for i, ticker in enumerate(lista_ativos, 1):
        try:
            resultado = coletar_ativo(ticker)
            resultado["Data_Analise"] = data_hoje
            dados_consolidados.append(resultado)
            contagem_fontes[resultado["Fonte_DY"]] += 1

            print(
                f"  OK [{i:>2}/{len(lista_ativos)}] {ticker:<8} | "
                f"R$ {resultado['Preco_Atual']:>8.2f} | "
                f"DY: {resultado['Dividend_Yield_Porcento']:>5.2f}% "
                f"[{resultado['Fonte_DY']:<14}] | "
                f"P/L: {resultado['P_L']:>6.1f} | "
                f"ROE: {resultado['ROE_Porcento']:>5.1f}% | "
                f"{resultado['Status_Consultor']}"
            )
        except Exception as e:
            erros.append({"Ticker": ticker, "Erro": str(e)})
            print(f"  ERRO [{i:>2}/{len(lista_ativos)}] {ticker:<8} | {str(e)}")

        time.sleep(DELAY_REQ)

    print(f"\n{'='*80}")
    print(f"  {len(dados_consolidados)} processados | {len(erros)} erros")
    print(
        f"  Fontes de DY -> StatusInvest: {contagem_fontes['StatusInvest']} | "
        f"yf-historico: {contagem_fontes['yf-historico']} | "
        f"yf-info: {contagem_fontes['yf-info']}"
    )
    if contagem_fontes["StatusInvest"] == 0 and len(dados_consolidados) > 0:
        print("  AVISO: Status Invest não retornou nenhum DY nesta rodada - "
              "verifique se o layout do site mudou.")
    print(f"{'='*80}\n")

    return dados_consolidados, erros, contagem_fontes


def salvar_sql(dados_consolidados: list[dict]) -> None:
    if not dados_consolidados:
        return

    conexao = None
    try:
        conexao = pyodbc.connect(CONN_STR)
        cursor = conexao.cursor()

        for col_sql in [
            "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo.historico_financas') AND name='ROE_Porcento') "
            "ALTER TABLE dbo.historico_financas ADD ROE_Porcento FLOAT",
            "IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id=OBJECT_ID('dbo.historico_financas') AND name='Fonte_DY') "
            "ALTER TABLE dbo.historico_financas ADD Fonte_DY VARCHAR(20)",
        ]:
            cursor.execute(col_sql)

        cursor.execute("""
            IF NOT EXISTS (
                SELECT 1 FROM sys.indexes
                WHERE name = 'IX_historico_financas_Ticker_Data'
                AND object_id = OBJECT_ID('dbo.historico_financas')
            )
            CREATE INDEX IX_historico_financas_Ticker_Data
            ON dbo.historico_financas (Ticker, Data_Analise)
        """)

        comando_sql = """
            INSERT INTO dbo.historico_financas (
                Data_Analise, Ticker, Tipo_Ativo, Empresa, Preco_Atual, P_L,
                Dividend_Yield_Porcento, Fonte_DY, Status_Consultor,
                Margem_Liquida_Porcento, Divida_Liquida_EBITDA, EV_EBITDA, P_VPA, ROE_Porcento
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        rows = [(
            d["Data_Analise"],            d["Ticker"],
            d["Tipo_Ativo"],              d["Empresa"],
            d["Preco_Atual"],             d["P_L"],
            d["Dividend_Yield_Porcento"], d["Fonte_DY"],
            d["Status_Consultor"],        d["Margem_Liquida_Porcento"],
            d["Divida_Liquida_EBITDA"],   d["EV_EBITDA"],
            d["P_VPA"],                   d["ROE_Porcento"],
        ) for d in dados_consolidados]

        cursor.executemany(comando_sql, rows)
        conexao.commit()
        print(f"[SQL] OK - {len(rows)} registros salvos (índice Ticker+Data garantido).")

    except Exception as e:
        if conexao:
            conexao.rollback()
        print(f"[SQL] ERRO - Rollback. Detalhe: {e}")
    finally:
        if conexao:
            conexao.close()


def salvar_excel(dados_consolidados: list[dict], erros: list[dict]) -> None:
    try:
        df = pd.DataFrame(dados_consolidados)

        df_historico_antigo = pd.DataFrame()
        if os.path.exists(CAMINHO_EXCEL):
            try:
                df_historico_antigo = pd.read_excel(CAMINHO_EXCEL, sheet_name="Historico")
            except ValueError:
                df_historico_antigo = pd.DataFrame()
            except PermissionError:
                print("[EXCEL] AVISO - Não foi possível ler o histórico anterior "
                      "(arquivo aberto?). Prosseguindo sem ele.")

        df_historico = pd.concat([df_historico_antigo, df], ignore_index=True)

        if "Data_Analise" in df_historico.columns and not df_historico.empty:
            datas = pd.to_datetime(df_historico["Data_Analise"], errors="coerce")
            limite = datetime.datetime.now() - datetime.timedelta(days=DIAS_RETENCAO_EXCEL)
            antes = len(df_historico)
            df_historico = df_historico[datas >= limite]
            removidas = antes - len(df_historico)
            if removidas > 0:
                print(f"[EXCEL] {removidas} registros antigos removidos do histórico do "
                      f"Excel (retenção de {DIAS_RETENCAO_EXCEL} dias). Continuam no SQL Server.")

        with pd.ExcelWriter(CAMINHO_EXCEL, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Dados_Atuais")

            df_op = df[df["Status_Consultor"].str.contains("FORTE COMPRA|BOA COMPRA", na=False)]
            df_op.sort_values("Dividend_Yield_Porcento", ascending=False).to_excel(
                writer, index=False, sheet_name="Oportunidades"
            )

            df_historico.to_excel(writer, index=False, sheet_name="Historico")

            if erros:
                pd.DataFrame(erros).to_excel(writer, index=False, sheet_name="Erros_Log")
                print(f"[EXCEL] {len(erros)} erros registrados na aba Erros_Log.")

        print(f"[EXCEL] OK - {len(dados_consolidados)} ativos salvos em: {CAMINHO_EXCEL}")
        print(f"[EXCEL] Historico no arquivo: {len(df_historico)} registros "
              f"({df_historico['Data_Analise'].nunique()} rodadas, últimos {DIAS_RETENCAO_EXCEL} dias).")

    except PermissionError:
        print("[EXCEL] ERRO - Feche o arquivo antes de rodar!")
    except Exception as e:
        print(f"[EXCEL] ERRO - {e}")


def enviar_alerta_email(dados: list[dict], data_analise: str) -> None:
    if not EMAIL_REMETENTE or not EMAIL_SENHA_APP:
        print("[EMAIL] AVISO - EMAIL_REMETENTE/EMAIL_SENHA_APP não configurados no "
              ".env. Pulando envio.")
        return

    oportunidades = [
        d for d in dados
        if "FORTE COMPRA" in d["Status_Consultor"] or "BOA COMPRA" in d["Status_Consultor"]
    ]

    if not oportunidades:
        print("[EMAIL] Nenhuma oportunidade encontrada - e-mail não enviado.")
        return

    fortes = [d for d in oportunidades if "FORTE COMPRA" in d["Status_Consultor"]]
    boas = [d for d in oportunidades if "BOA COMPRA" in d["Status_Consultor"]]
    linhas = []

    if fortes:
        linhas.append("🔥 FORTE COMPRA")
        linhas.append("-" * 50)
        for d in sorted(fortes, key=lambda x: x["Dividend_Yield_Porcento"], reverse=True):
            linhas.append(
                f"  {d['Ticker']:<8} | R$ {d['Preco_Atual']:>8.2f} | "
                f"DY: {d['Dividend_Yield_Porcento']:.2f}% | "
                f"P/L: {d['P_L']:.1f} | ROE: {d['ROE_Porcento']:.1f}%"
            )
        linhas.append("")

    if boas:
        linhas.append("✅ BOA COMPRA")
        linhas.append("-" * 50)
        for d in sorted(boas, key=lambda x: x["Dividend_Yield_Porcento"], reverse=True):
            linhas.append(
                f"  {d['Ticker']:<8} | R$ {d['Preco_Atual']:>8.2f} | "
                f"DY: {d['Dividend_Yield_Porcento']:.2f}% | "
                f"P/L: {d['P_L']:.1f} | ROE: {d['ROE_Porcento']:.1f}%"
            )

    corpo = (
        f"Consultor Fundamentalista B3 - {data_analise}\n"
        f"{'='*50}\n\n"
        + "\n".join(linhas)
        + f"\n\n{'='*50}\n"
        f"Total: {len(fortes)} Forte Compra | {len(boas)} Boa Compra\n"
        f"Gerado automaticamente pelo Consultor B3"
    )

    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_REMETENTE
        msg["To"] = EMAIL_DESTINATARIO
        msg["Subject"] = (
            f"📈 B3 - {len(fortes)} Forte Compra | {len(boas)} Boa Compra | {data_analise[:10]}"
        )
        msg.attach(MIMEText(corpo, "plain", "utf-8"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as servidor:
            servidor.login(EMAIL_REMETENTE, EMAIL_SENHA_APP)
            servidor.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIO, msg.as_bytes())

        print(f"[EMAIL] OK - Alerta enviado - {len(oportunidades)} oportunidades.")

    except smtplib.SMTPAuthenticationError:
        print("[EMAIL] ERRO - Autenticação falhou. Verifique o .env (e-mail e Senha de App).")
    except Exception as e:
        print(f"[EMAIL] ERRO - {e}")


def main() -> None:
    data_hoje = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    dados_consolidados, erros, contagem_fontes = coletar_todos(LISTA_ATIVOS, data_hoje)

    if dados_consolidados:
        salvar_sql(dados_consolidados)
        salvar_excel(dados_consolidados, erros)
        enviar_alerta_email(dados_consolidados, data_hoje)
    else:
        print("[MAIN] Nenhum dado coletado - pulando persistência e envio de e-mail.")


if __name__ == "__main__":
    main()
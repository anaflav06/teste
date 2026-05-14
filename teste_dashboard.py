import io
import math
import base64
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    import requests
except Exception:
    requests = None

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None


# =========================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================
st.set_page_config(
    page_title="Fechamento de Entregadores",
    page_icon="🚚",
    layout="wide"
)



# =========================================================
# PERSISTÊNCIA EM JSON NO GITHUB
# =========================================================
# Estes arquivos ficam na raiz do repositório:
# - database_motoristas.json
# - database_placas.json
# - database_ceps.json
#
# Para salvar no GitHub, configure no Streamlit Cloud em Settings > Secrets:
# GITHUB_TOKEN = "seu_token"
# GITHUB_REPO = "anaflav06/dashboard-gds"
# GITHUB_BRANCH = "main"
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
GITHUB_REPO = st.secrets.get("GITHUB_REPO", "")
GITHUB_BRANCH = st.secrets.get("GITHUB_BRANCH", "main")

DATABASE_MOTORISTAS_JSON = "database_motoristas.json"
DATABASE_PLACAS_JSON = "database_placas.json"
DATABASE_CEPS_JSON = "database_ceps.json"


def github_json_ativo() -> bool:
    return bool(GITHUB_TOKEN and GITHUB_REPO and GITHUB_BRANCH and requests is not None)


def _github_headers() -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _github_url(arquivo: str) -> str:
    return f"https://api.github.com/repos/{GITHUB_REPO}/contents/{arquivo}"


@st.cache_data(show_spinner=False, ttl=30)
def carregar_json_github(arquivo: str, padrao_json: str = "[]"):
    """Carrega um JSON da raiz do GitHub. Se não conseguir, retorna o padrão."""
    padrao = json.loads(padrao_json)
    if not github_json_ativo():
        return padrao

    try:
        resp = requests.get(
            _github_url(arquivo),
            headers=_github_headers(),
            params={"ref": GITHUB_BRANCH},
            timeout=15,
        )
        if resp.status_code == 404:
            return padrao
        if not resp.ok:
            return padrao

        conteudo_base64 = resp.json().get("content", "")
        conteudo = base64.b64decode(conteudo_base64).decode("utf-8")
        if not conteudo.strip():
            return padrao
        return json.loads(conteudo)
    except Exception:
        return padrao


def salvar_json_github(arquivo: str, dados, mensagem_commit: str) -> bool:
    """Salva um JSON no GitHub. Retorna False sem quebrar a dashboard se falhar."""
    if not github_json_ativo():
        return False

    try:
        url = _github_url(arquivo)
        headers = _github_headers()

        sha_atual = None
        resp_get = requests.get(url, headers=headers, params={"ref": GITHUB_BRANCH}, timeout=15)
        if resp_get.ok:
            sha_atual = resp_get.json().get("sha")

        conteudo = json.dumps(dados, ensure_ascii=False, indent=2)
        payload = {
            "message": mensagem_commit,
            "content": base64.b64encode(conteudo.encode("utf-8")).decode("ascii"),
            "branch": GITHUB_BRANCH,
        }
        if sha_atual:
            payload["sha"] = sha_atual

        resp_put = requests.put(url, headers=headers, json=payload, timeout=20)
        if resp_put.ok:
            carregar_json_github.clear()
            return True
        return False
    except Exception:
        return False


def _normalizar_db_placas(dados) -> Dict[str, list]:
    if isinstance(dados, list):
        return {"placas_extra": dados, "placas_excluidas": []}
    if not isinstance(dados, dict):
        return {"placas_extra": [], "placas_excluidas": []}
    return {
        "placas_extra": dados.get("placas_extra", []) if isinstance(dados.get("placas_extra", []), list) else [],
        "placas_excluidas": dados.get("placas_excluidas", []) if isinstance(dados.get("placas_excluidas", []), list) else [],
    }


def _normalizar_db_motoristas(dados) -> Dict[str, list]:
    if isinstance(dados, list):
        return {"cnpj_extra": dados, "cnpj_excluidos": []}
    if not isinstance(dados, dict):
        return {"cnpj_extra": [], "cnpj_excluidos": []}
    return {
        "cnpj_extra": dados.get("cnpj_extra", []) if isinstance(dados.get("cnpj_extra", []), list) else [],
        "cnpj_excluidos": dados.get("cnpj_excluidos", []) if isinstance(dados.get("cnpj_excluidos", []), list) else [],
    }


def _normalizar_db_ceps(dados) -> list:
    if isinstance(dados, dict):
        dados = dados.get("reajustes_cep", [])
    return dados if isinstance(dados, list) else []


def carregar_db_placas() -> Dict[str, list]:
    return _normalizar_db_placas(carregar_json_github(DATABASE_PLACAS_JSON, "{}"))


def salvar_db_placas(dados: Dict[str, list]) -> bool:
    return salvar_json_github(DATABASE_PLACAS_JSON, _normalizar_db_placas(dados), "Atualiza database_placas.json")


def carregar_db_motoristas() -> Dict[str, list]:
    return _normalizar_db_motoristas(carregar_json_github(DATABASE_MOTORISTAS_JSON, "{}"))


def salvar_db_motoristas(dados: Dict[str, list]) -> bool:
    return salvar_json_github(DATABASE_MOTORISTAS_JSON, _normalizar_db_motoristas(dados), "Atualiza database_motoristas.json")


def carregar_db_ceps() -> list:
    return _normalizar_db_ceps(carregar_json_github(DATABASE_CEPS_JSON, "[]"))


def salvar_db_ceps(dados: list) -> bool:
    return salvar_json_github(DATABASE_CEPS_JSON, _normalizar_db_ceps(dados), "Atualiza database_ceps.json")


# =========================================================
# CSS
# =========================================================
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 2.0rem;
            padding-bottom: 2rem;
            max-width: 1700px;
        }

        [data-testid="stSidebar"] {
            background: #f4f7fb;
            border-right: 1px solid #e5e7eb;
        }

        .main-header {
            background: linear-gradient(90deg, #06142f 0%, #08204d 45%, #0b2b69 100%);
            border-radius: 22px;
            padding: 28px 34px;
            margin-bottom: 26px;
            box-shadow: 0 12px 28px rgba(2, 6, 23, 0.20);
            color: white;
        }

        .main-header-title {
            font-size: 2.35rem;
            font-weight: 850;
            line-height: 1.1;
            margin-bottom: 8px;
        }

        .main-header-subtitle {
            font-size: 1.05rem;
            color: rgba(255,255,255,0.78);
        }

        .kpi-card {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 18px;
            padding: 18px 18px;
            min-height: 100px;
            box-shadow: 0 8px 20px rgba(15, 23, 42, 0.05);
        }

        .kpi-title {
            font-size: 14px;
            color: #64748b;
            margin-bottom: 10px;
            font-weight: 650;
        }

        .kpi-value {
            font-size: 31px;
            font-weight: 850;
            color: #111827;
            line-height: 1;
        }

        .section-heading {
            font-size: 25px;
            font-weight: 850;
            color: #1f2937;
            margin: 10px 0 16px 0;
        }

        .small-note {
            color: #6b7280;
            font-size: 13px;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid #e5e7eb;
            border-radius: 14px;
            overflow: hidden;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# BASES INTERNAS - VALORES POR CEP E PLACAS
# =========================================================
# Regra do CEP: a tabela usa o começo do CEP com 0 na frente.
# Exemplo: código 10 = CEP iniciando com 010, código 45 = CEP iniciando com 045.
BASE_VALORES_CEP = [
    ("10", 6.00, 6.00),
    ("11", 6.00, 6.00),
    ("12", 6.00, 6.00),
    ("13", 6.00, 6.00),
    ("14", 5.00, 5.00),
    ("15", 5.00, 5.00),
    ("28", 8.50, 8.50),
    ("29", 8.50, 8.50),
    ("30", 6.00, 6.00),
    ("39", 8.50, 8.50),
    ("40", 5.00, 5.00),
    ("41", 5.00, 5.00),
    ("42", 5.00, 5.00),
    ("43", 5.00, 5.00),
    ("44", 8.50, 8.50),
    ("45", 5.00, 6.00),
    ("46", 5.00, 6.00),
    ("47", 6.00, 6.00),
    ("48", 10.00, 10.00),
    ("49", 8.50, 8.50),
    ("50", 5.00, 5.00),
    ("51", 8.50, 8.50),
    ("52", 8.50, 8.50),
    ("53", 7.00, 7.00),
    ("54", 5.00, 5.00),
    ("55", 7.00, 7.00),
    ("56", 7.00, 8.50),
    ("57", 7.00, 7.00),
    ("58", 8.50, 8.50),
    ("60", 7.00, 7.00),
    ("61", 7.00, 7.00),
    ("62", 7.00, 7.00),
    ("63", 7.00, 7.00),
    ("67", 7.00, 8.50),
    ("68", 10.00, 10.00),
    ("69", 10.00, 10.00),
    ("80", 8.50, 8.50),
    ("81", 8.50, 8.50),
    ("83", 8.50, 8.50),
    ("84", 8.50, 8.50),
    ("90", 6.00, 6.00),
    ("91", 6.00, 6.00),
    ("92", 6.00, 6.00),
    ("93", 7.00, 7.00),
    ("94", 7.00, 7.00),
    ("95", 6.00, 6.00),
    ("96", 6.00, 6.00),
    ("97", 6.00, 6.00),
    ("98", 6.00, 6.00),
]

BASE_PLACAS_VEICULOS = [
    ("BZF2B49", "MOTO"),
    ("CFI2B71", "CARRO"),
    ("CLJ3F62", "CARRO"),
    ("CRJ6318", "CARRO"),
    ("CWG8E34", "CARRO"),
    ("DIZ4F95", "CARRO"),
    ("DDE3J85", "CARRO"),
    ("DMZ6317", "CARRO"),
    ("DNO1673", "CARRO"),
    ("DOM9J45", "CARRO"),
    ("DOW2G90", "MOTO"),
    ("DQV4H21", "CARRO"),
    ("DRE1E06", "CARRO"),
    ("DUH0E74", "CARRO"),
    ("DVV6I21", "MOTO"),
    ("DWU6B82", "MOTO"),
    ("DZI3H80", "MOTO"),
    ("EFC4F64", "CARRO"),
    ("EFT8H38", "CARRO"),
    ("EKG6B25", "MOTO"),
    ("EQJ5241", "CARRO"),
    ("EQJ5C41", "CARRO"),
    ("ERK2A88", "CARRO"),
    ("EUY6558", "CARRO"),
    ("EXR4J85", "MOTO"),
    ("FCA3I52", "CARRO"),
    ("FCB2C33", "CARRO"),
    ("FHB9F46", "CARRO"),
    ("FHX3979", "MOTO"),
    ("FJO2D53", "CARRO"),
    ("FNU1I85", "CARRO"),
    ("FQO4A14", "CARRO"),
    ("FQZ3147", "CARRO"),
    ("FVN6D97", "MOTO"),
    ("FYJ6835", "CARRO"),
    ("GCJ6D46", "MOTO"),
    ("GDQ9J06", "MOTO"),
    ("GDS0105", "CARRO"),
    ("GFT7E48", "CARRO"),
    ("HOB9J03", "CARRO"),
    ("IMX4J12", "CARRO"),
    ("IOG5494", "CARRO"),
    ("JKA9677", "MOTO"),
    ("KNX7947", "CARRO"),
    ("KPQ2848", "MOTO"),
    ("KXF7G43", "CARRO"),
    ("LKW8B95", "CARRO"),
    ("QQW8C60", "CARRO"),
    ("RFE2A05", "CARRO"),
    ("RNH2H82", "CARRO"),
    ("RTN9B33", "CARRO"),
    ("SVZ1J77", "MOTO"),
    ("SWY0H59", "MOTO"),
    ("TKV4H26", "MOTO"),
    ("GYT1T09", "CARRO"),
    ("FAS1E74", "CARRO"),
    
]




# =========================================================
# BASE INTERNA - CNPJ DOS MOTORISTAS
# =========================================================
BASE_CNPJ_MOTORISTAS = {
    "ADRIANO APARECIDO FERREIRA DE OLIVEIRA": "64.053.611/0001-80",
    "ALDACIR MARCON": "45.916.701/0001-03",
    "ANDREIA ALVES DE SOUSA": "54.921.728/0001-85",
    "BIANCA ZANARDIR GOMES DA SILVA": "425.879.028-19",
    "CAIQUE ADRIAN FERREIRA DA SILVA BRITO": "51.019.384/0001-25",
    "CAIQUE XAVIER TORRES DA SILVA": "63.428.066/0001-05",
    "CARLOS EDUARDO ZACHARIAS": "091.386.348-30",
    "CICERO DONESI": "52.140.817/0001-69",
    "CLEBER OTAVIO OLIVEIRA DA SILVA": "57.554.463/0001-12",
    "DAVI DE SOUZA LAURINDO": "37.627.765/0001-66",
    "DANIEL PEREIRA SILVA": "50.603.845/0001-40",
    "DIEGO DA SILVA MONTEIRO": "54.794.473/0001-37",
    "EVANDRO RODRIGUES BENTO": "55.267.940/0001-33",
    "EDIPO SOARES DE CARVALHO": "355.106.038-01",
    "FELIPE AUGUSTO": "SEM CNPJ",
    "FELIPE DA SILVA ABRANTES": "405.031.978-07",
    "FERNANDO LOURENÇO DE PAULA": "34.605.101/0001-08",
    "GERSON CANDIDO DA SILVA": "29.882.972/0001-39",
    "GERSON LANDIM": "SEM CNPJ",
    "GUSTAVO NOVAES SILVA": "64.317.201/0001-08",
    "ISAQUE FERNANDO DA COSTA MARTINS SILVA": "62.346.255/0001-68",
    "GUILHERME SOARES SILVA": "42.690.386/0001-50",
    "JOEL ELIAS FILHO": "49.631.479/0001-53",
    "JEFFERSON NOVAIS DE OLIVEIRA": "30.226.280/0001-11",
    "JOSE LUIS SABATINI": "42.891.659/0001-25",
    "JAIMILTON ALVES MOREIRA RODRIGUES": "55.127.216/0001-04",
    "JUAN OLIVEIRA DE MENEZES": "44.839.742/0001-80",
    "KELVIN MESSIAS DOS SANTOS OLIVEIRA": "59.462.329/0001-17",
    "MARCELO SANTOS PRAÇA": "40.218.730/0001-88",
    "MARIA NATALIA DA SILVA FIGUEIREDO BARBOSA": "47.702.239/0001-77",
    "MAURO ALGAVES": "34.247.737/0001-25",
    "MARCIO LEITE DE OLIVEIRA": "52.484.794/0001-00",
    "MATHEUS PINTO DE OLIVEIRA": "61.632.284/0001-23",
    "PAULO VICTOR SOUSA LOPES": "50.787.799/0001-86",
    "RENATA FERNANDES TERENSE SANDRI": "30.990.911/0001-74",
    "RICARDO RAPOSO": "63.975.972/0001-11",
    "RODOLFO LEONARDO DE MAGALHAES BEZZERA": "35.773.299/0001-00",
    "ROBSON RIBEIRO DA SILVA": "64.967.271/0001-01",
    "RODNEI FOGACA DA SILVA FILHO": "32.839.434/0001-76",
    "SANDRO COSTA DA SILVA": "63.073.440/0001-99",
    "SILAS MAIA DE JESUS": "42.786.379/0001-57",
    "THATIANA PEREIRA RIBEIRO POFFO": "37.631.336/001-62",
    "VINICIUS DE SOUZA RODRIGUES": "57.405.700/0001-02",
    "VIVIANE DA COSTA SILA": "46.450.740/0001-20",
    "WALLACE ANDRE": "SEM CNPJ",
    "WALLACE OLIVEIRA DA SILVA": "42.473.134/0001-70",
    "WALLAS AMARANTE ALVES DOS SANTOS": "45.765.135/0001-86",
    "WELLINGTON DIAS DA SILVA": "24.805.136/0001-37",
    "WENDHELL ANDERSON DA SIVA CAVALCANTE": "65.538.430/0001-07",
    "WILLIAM SOUZA SANTOS": "18.766.429/0001-50",
    "JESSICA PEREIRA DOS SANTOS": "48.847.009/0001-69",
    "JOSE CLAUDINEY DA SILVA": "18.036.024/0001-66",
    "JOSE LUIZ CANDIDO": "18.036.024/0001-66",
    "JOSE WILSON DA SILVA": "63.728.550/0001-41",
    "JULIANA CARVALHO LEBRE": "55.478.354/0001-38",
    "KAUA FERNANDES BRANDÃO SILVA": "54.476.632/0001-87",
    "LUCAS FAGUNDES DE LIMA": "26.985.251/0001-66",
    "LUIZ CARLOS MARQUES": "42.929.833/0001-81",
    "MARCOS ANTONIO ARAUJO GOMES": "31.851.508/0001-27",
    "MARCOS VINICIUS DOS SANTOS FERREIRA": "61.588.016/0001-51",
    "MATHEUS LUIZ FERREIRA": "65.274.932/0001-78",
    "MAURICIO RAMOS": "43.102.624/0001-22",
    "PAULO EDSON FERNANDES": "60.580.076/0001-65",
    "RAPHAEL BARBOSA TEIXEIRA": "37.385.753/0001-72",
    "REGINALDO GONCALVES DOS SANTOS": "26.561.607/0001-34",
    "SAMIRA MARIA RODRIGUES SOBRINHO SILVA": "52.959.490/0001-51",
    "SEBASTIÃO ROCHA DA SILVA": "59.196.445/0001-31",
    "SILVIO PEDRO ROBERTO": "59.989.591/0001-14",
    "SIRLENE ALMEIDA DA COSTA": "21.123.958/0001-40",
    "VANESSA OLIVEIRA SANTOS MARANGONI": "64.379.294/0001-97",
    "VINICIUS ALVES FERREIRA DE ARAUJO": "64.983.884/0001-24",
    "WESLEY SILVA DE PAULA": "33.908.396/0001-29",
    "YAEL DE CASTRO FERREIRA GOMES": "33.908.396/0001-29",
}

# Arquivo local para motoristas/CNPJ cadastrados pela dashboard.
# Esses cadastros sobrescrevem a base interna quando o nome for igual.
ARQUIVO_MOTORISTAS_CNPJ_EXTRA = Path("motoristas_cnpj_extra.csv")
COLUNAS_MOTORISTAS_CNPJ_EXTRA = ["Nome Motorista", "CNPJ"]
ARQUIVO_MOTORISTAS_CNPJ_EXCLUIDOS = Path("motoristas_cnpj_excluidos.csv")
COLUNAS_MOTORISTAS_CNPJ_EXCLUIDOS = ["Nome Motorista"]

# =========================================================
# CADASTRO LOCAL DE PLACAS / TIPO DE VEÍCULO
# =========================================================
ARQUIVO_MOTORISTAS_EXTRA = Path("motoristas_placas_extra.csv")
COLUNAS_MOTORISTAS_EXTRA = ["Motorista Cadastro", "Placa", "Tipo Veículo"]
ARQUIVO_PLACAS_EXCLUIDAS = Path("placas_excluidas.csv")
COLUNAS_PLACAS_EXCLUIDAS = ["Placa"]


def normalizar_placa(valor) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(valor or "").upper())


def salvar_motoristas_extra_df(df_extra: pd.DataFrame) -> None:
    df = df_extra.copy() if df_extra is not None else pd.DataFrame(columns=COLUNAS_MOTORISTAS_EXTRA)
    for col in COLUNAS_MOTORISTAS_EXTRA:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUNAS_MOTORISTAS_EXTRA].copy()
    df["Motorista Cadastro"] = ""
    df["Placa"] = df["Placa"].apply(normalizar_placa)
    df["Tipo Veículo"] = df["Tipo Veículo"].astype(str).str.strip().str.upper()
    df = df[(df["Placa"] != "") & (df["Tipo Veículo"].isin(["MOTO", "CARRO"]))].copy()
    df = df.drop_duplicates(subset=["Placa"], keep="last").sort_values("Placa")
    df.to_csv(ARQUIVO_MOTORISTAS_EXTRA, index=False, encoding="utf-8-sig")

    db = carregar_db_placas()
    db["placas_extra"] = df.to_dict(orient="records")
    salvar_db_placas(db)


def carregar_placas_excluidas_csv() -> pd.DataFrame:
    """Carrega placas removidas manualmente da base aplicada."""
    db = carregar_db_placas()
    if db.get("placas_excluidas"):
        df = pd.DataFrame(db.get("placas_excluidas", []))
    elif ARQUIVO_PLACAS_EXCLUIDAS.exists():
        try:
            df = pd.read_csv(ARQUIVO_PLACAS_EXCLUIDAS, dtype=str).fillna("")
        except Exception:
            df = pd.DataFrame(columns=COLUNAS_PLACAS_EXCLUIDAS)
    else:
        df = pd.DataFrame(columns=COLUNAS_PLACAS_EXCLUIDAS)

    if "Placa" not in df.columns:
        df["Placa"] = ""

    df = df[["Placa"]].copy()
    df["Placa"] = df["Placa"].apply(normalizar_placa)
    df = df[df["Placa"] != ""].drop_duplicates(subset=["Placa"], keep="last")
    return df.reset_index(drop=True)


def preparar_placas_excluidas_para_cache(df_excluidas: pd.DataFrame) -> Tuple[str, ...]:
    if df_excluidas.empty or "Placa" not in df_excluidas.columns:
        return tuple()
    placas = (
        df_excluidas["Placa"]
        .astype(str)
        .apply(normalizar_placa)
    )
    placas = [p for p in placas if p]
    return tuple(sorted(set(placas)))


def salvar_placas_excluidas(df_excluidas: pd.DataFrame) -> None:
    df = df_excluidas.copy() if df_excluidas is not None else pd.DataFrame(columns=COLUNAS_PLACAS_EXCLUIDAS)
    if "Placa" not in df.columns:
        df["Placa"] = ""
    df = df[["Placa"]].copy()
    df["Placa"] = df["Placa"].apply(normalizar_placa)
    df = df[df["Placa"] != ""].drop_duplicates(subset=["Placa"], keep="last").sort_values("Placa")
    df.to_csv(ARQUIVO_PLACAS_EXCLUIDAS, index=False, encoding="utf-8-sig")

    db = carregar_db_placas()
    db["placas_excluidas"] = df.to_dict(orient="records")
    salvar_db_placas(db)


def excluir_placa_base(placa: str) -> Tuple[bool, str]:
    placa = normalizar_placa(placa)
    if not placa or len(placa) < 7:
        return False, "Informe uma placa válida para excluir."

    # Remove do cadastro manual, se existir.
    df_extra = carregar_motoristas_extra_csv()
    removida_manual = False
    if not df_extra.empty:
        qtd_antes = len(df_extra)
        df_extra = df_extra[df_extra["Placa"].apply(normalizar_placa) != placa].copy()
        removida_manual = len(df_extra) < qtd_antes
        salvar_motoristas_extra_df(df_extra)

    # Também grava na lista de excluídas para remover placa que venha da base interna.
    df_excluidas = carregar_placas_excluidas_csv()
    if placa not in set(df_excluidas["Placa"].astype(str).apply(normalizar_placa)):
        df_excluidas = pd.concat([df_excluidas, pd.DataFrame([{"Placa": placa}])], ignore_index=True)
        salvar_placas_excluidas(df_excluidas)

    if removida_manual:
        return True, f"Placa {placa} removida do cadastro manual e bloqueada na base aplicada."
    return True, f"Placa {placa} bloqueada/removida da base aplicada."


def remover_placa_da_lista_excluidas(placa: str) -> None:
    placa = normalizar_placa(placa)
    if not placa or not ARQUIVO_PLACAS_EXCLUIDAS.exists():
        return
    df = carregar_placas_excluidas_csv()
    if df.empty:
        return
    df = df[df["Placa"].apply(normalizar_placa) != placa].copy()
    salvar_placas_excluidas(df)


def carregar_motoristas_extra_csv() -> pd.DataFrame:
    """Carrega placas/tipos de veículo cadastrados manualmente na própria dashboard."""
    db = carregar_db_placas()
    if db.get("placas_extra"):
        df = pd.DataFrame(db.get("placas_extra", []))
    elif ARQUIVO_MOTORISTAS_EXTRA.exists():
        try:
            df = pd.read_csv(ARQUIVO_MOTORISTAS_EXTRA, dtype=str).fillna("")
        except Exception:
            df = pd.DataFrame(columns=COLUNAS_MOTORISTAS_EXTRA)
    else:
        df = pd.DataFrame(columns=COLUNAS_MOTORISTAS_EXTRA)

    for col in COLUNAS_MOTORISTAS_EXTRA:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUNAS_MOTORISTAS_EXTRA].copy()
    # A partir desta versão, o cadastro manual usa somente Placa + Tipo Veículo.
    # Mantemos a coluna Motorista Cadastro vazia apenas por compatibilidade com arquivos antigos.
    df["Motorista Cadastro"] = ""
    df["Placa"] = df["Placa"].apply(normalizar_placa)
    df["Tipo Veículo"] = df["Tipo Veículo"].astype(str).str.strip().str.upper()
    df = df[(df["Placa"] != "") & (df["Tipo Veículo"].isin(["MOTO", "CARRO"]))].copy()
    return df.drop_duplicates(subset=["Placa"], keep="last").reset_index(drop=True)


def salvar_motorista_extra(placa: str, tipo_veiculo: str) -> Tuple[bool, str]:
    placa = normalizar_placa(placa)
    tipo_veiculo = limpar_texto(tipo_veiculo).upper()

    if not placa or len(placa) < 7:
        return False, "Informe uma placa válida."
    if tipo_veiculo not in ["MOTO", "CARRO"]:
        return False, "Selecione MOTO ou CARRO."

    df = carregar_motoristas_extra_csv()
    novo = pd.DataFrame([{
        "Motorista Cadastro": "",
        "Placa": placa,
        "Tipo Veículo": tipo_veiculo,
    }])

    df = pd.concat([df, novo], ignore_index=True)
    df = df.drop_duplicates(subset=["Placa"], keep="last")
    df = df.sort_values(["Placa"]).reset_index(drop=True)
    salvar_motoristas_extra_df(df)
    remover_placa_da_lista_excluidas(placa)
    return True, f"Placa {placa} cadastrada/atualizada como {tipo_veiculo}."


def preparar_motoristas_extra_para_cache(df_extra: pd.DataFrame) -> Tuple[Tuple[str, str, str], ...]:
    if df_extra.empty:
        return tuple()
    df = df_extra.copy()
    for col in COLUNAS_MOTORISTAS_EXTRA:
        if col not in df.columns:
            df[col] = ""
    df["Motorista Cadastro"] = ""
    df["Placa"] = df["Placa"].apply(normalizar_placa)
    df["Tipo Veículo"] = df["Tipo Veículo"].astype(str).str.strip().str.upper()
    df = df[(df["Placa"] != "") & (df["Tipo Veículo"].isin(["MOTO", "CARRO"]))].copy()
    df = df.drop_duplicates(subset=["Placa"], keep="last")
    return tuple(df[COLUNAS_MOTORISTAS_EXTRA].itertuples(index=False, name=None))


def montar_base_placas_final(
    motoristas_extra_tuple: Tuple[Tuple[str, str, str], ...] = tuple(),
    placas_excluidas_tuple: Tuple[str, ...] = tuple(),
) -> pd.DataFrame:
    base = carregar_base_placas_interna().copy()
    base["Origem"] = "Base interna"

    extras = []
    for _motorista, placa, tipo in motoristas_extra_tuple:
        extras.append({
            "Placa": normalizar_placa(placa),
            "Tipo Veículo": limpar_texto(tipo).upper(),
            "Motorista Cadastro": "",
            "Origem": "Cadastro manual",
        })

    if extras:
        df_extra = pd.DataFrame(extras)
        base = pd.concat([base, df_extra], ignore_index=True)

    base["Placa"] = base["Placa"].apply(normalizar_placa)
    base["Tipo Veículo"] = base["Tipo Veículo"].astype(str).str.strip().str.upper()
    base["Motorista Cadastro"] = ""
    base["Origem"] = base["Origem"].fillna("Base interna").astype(str)
    base = base[(base["Placa"] != "") & (base["Tipo Veículo"].isin(["MOTO", "CARRO"]))].copy()

    # O cadastro manual deve sobrescrever a base interna quando a placa for igual.
    base = base.drop_duplicates(subset=["Placa"], keep="last").reset_index(drop=True)

    placas_excluidas = {normalizar_placa(p) for p in placas_excluidas_tuple if normalizar_placa(p)}
    if placas_excluidas:
        base = base[~base["Placa"].isin(placas_excluidas)].copy()

    return base.reset_index(drop=True)


# =========================================================
# CADASTRO LOCAL DE REAJUSTES DE VALORES POR CEP
# =========================================================
ARQUIVO_REAJUSTES_CEP = Path("reajustes_valores_cep.csv")
COLUNAS_REAJUSTES_CEP = ["CEP Prefixo", "Tipo Veículo", "Valor CEP"]


def normalizar_prefixo_cep(valor) -> str:
    """Aceita 10, 010 ou CEP completo; retorna prefixo no formato 010."""
    digitos = re.sub(r"\D", "", str(valor or ""))
    if not digitos:
        return ""
    if len(digitos) >= 8:
        return digitos.zfill(8)[:3]
    if len(digitos) == 2:
        return f"0{digitos}"
    if len(digitos) == 1:
        return f"00{digitos}"
    return digitos[:3].zfill(3)


def salvar_reajustes_cep_df(df_reajustes: pd.DataFrame) -> None:
    df = df_reajustes.copy() if df_reajustes is not None else pd.DataFrame(columns=COLUNAS_REAJUSTES_CEP)
    for col in COLUNAS_REAJUSTES_CEP:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUNAS_REAJUSTES_CEP].copy()
    df["CEP Prefixo"] = df["CEP Prefixo"].apply(normalizar_prefixo_cep)
    df["Tipo Veículo"] = df["Tipo Veículo"].astype(str).str.strip().str.upper()
    df["Valor CEP"] = df["Valor CEP"].apply(to_float)
    df = df[(df["CEP Prefixo"] != "") & (df["Tipo Veículo"].isin(["MOTO", "CARRO"])) & (df["Valor CEP"] > 0)].copy()
    df = df.drop_duplicates(subset=["CEP Prefixo", "Tipo Veículo"], keep="last")
    df = df.sort_values(["CEP Prefixo", "Tipo Veículo"]).reset_index(drop=True)
    df.to_csv(ARQUIVO_REAJUSTES_CEP, index=False, encoding="utf-8-sig")
    salvar_db_ceps(df.to_dict(orient="records"))


def carregar_reajustes_cep_csv() -> pd.DataFrame:
    """Carrega reajustes de valor por CEP cadastrados manualmente na dashboard."""
    dados_json = carregar_db_ceps()
    if dados_json:
        df = pd.DataFrame(dados_json)
    elif ARQUIVO_REAJUSTES_CEP.exists():
        try:
            df = pd.read_csv(ARQUIVO_REAJUSTES_CEP, dtype=str).fillna("")
        except Exception:
            df = pd.DataFrame(columns=COLUNAS_REAJUSTES_CEP)
    else:
        df = pd.DataFrame(columns=COLUNAS_REAJUSTES_CEP)

    for col in COLUNAS_REAJUSTES_CEP:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUNAS_REAJUSTES_CEP].copy()
    df["CEP Prefixo"] = df["CEP Prefixo"].apply(normalizar_prefixo_cep)
    df["Tipo Veículo"] = df["Tipo Veículo"].astype(str).str.strip().str.upper()
    df["Valor CEP"] = df["Valor CEP"].apply(to_float)
    df = df[(df["CEP Prefixo"] != "") & (df["Tipo Veículo"].isin(["MOTO", "CARRO"]))].copy()
    return df.drop_duplicates(subset=["CEP Prefixo", "Tipo Veículo"], keep="last").reset_index(drop=True)


def salvar_reajuste_cep(prefixo_cep: str, tipo_veiculo: str, valor_cep) -> Tuple[bool, str]:
    prefixo_cep = normalizar_prefixo_cep(prefixo_cep)
    tipo_veiculo = limpar_texto(tipo_veiculo).upper()
    valor = to_float(valor_cep)

    if not prefixo_cep or len(prefixo_cep) != 3:
        return False, "Informe um prefixo de CEP válido. Exemplo: 010, 045 ou 056."
    if tipo_veiculo not in ["MOTO", "CARRO"]:
        return False, "Selecione MOTO ou CARRO."
    if valor <= 0:
        return False, "Informe um valor maior que zero."

    df = carregar_reajustes_cep_csv()
    novo = pd.DataFrame([{
        "CEP Prefixo": prefixo_cep,
        "Tipo Veículo": tipo_veiculo,
        "Valor CEP": valor,
    }])

    df = pd.concat([df, novo], ignore_index=True)
    df = df.drop_duplicates(subset=["CEP Prefixo", "Tipo Veículo"], keep="last")
    df = df.sort_values(["CEP Prefixo", "Tipo Veículo"]).reset_index(drop=True)
    salvar_reajustes_cep_df(df)
    return True, f"Valor do CEP {prefixo_cep} / {tipo_veiculo} reajustado para {moeda(valor)}."


def salvar_novo_cep(prefixo_cep: str, valor_moto, valor_carro) -> Tuple[bool, str]:
    """Cadastra um novo CEP/prefixo com valores para MOTO e CARRO.

    Usa o mesmo arquivo dos reajustes, pois qualquer prefixo salvo nele
    passa a compor a base aplicada no fechamento.
    """
    prefixo_cep = normalizar_prefixo_cep(prefixo_cep)
    valor_moto_float = to_float(valor_moto)
    valor_carro_float = to_float(valor_carro)

    if not prefixo_cep or len(prefixo_cep) != 3:
        return False, "Informe um prefixo de CEP válido. Exemplo: 010, 045 ou 056."
    if valor_moto_float <= 0:
        return False, "Informe um valor de MOTO maior que zero."
    if valor_carro_float <= 0:
        return False, "Informe um valor de CARRO maior que zero."

    df = carregar_reajustes_cep_csv()
    novos = pd.DataFrame([
        {"CEP Prefixo": prefixo_cep, "Tipo Veículo": "MOTO", "Valor CEP": valor_moto_float},
        {"CEP Prefixo": prefixo_cep, "Tipo Veículo": "CARRO", "Valor CEP": valor_carro_float},
    ])

    df = pd.concat([df, novos], ignore_index=True)
    df = df.drop_duplicates(subset=["CEP Prefixo", "Tipo Veículo"], keep="last")
    df = df.sort_values(["CEP Prefixo", "Tipo Veículo"]).reset_index(drop=True)
    salvar_reajustes_cep_df(df)
    return True, f"CEP {prefixo_cep} cadastrado: MOTO {moeda(valor_moto_float)} | CARRO {moeda(valor_carro_float)}."


def preparar_reajustes_cep_para_cache(df_reajustes: pd.DataFrame) -> Tuple[Tuple[str, str, float], ...]:
    if df_reajustes.empty:
        return tuple()
    df = df_reajustes.copy()
    for col in COLUNAS_REAJUSTES_CEP:
        if col not in df.columns:
            df[col] = ""
    df["CEP Prefixo"] = df["CEP Prefixo"].apply(normalizar_prefixo_cep)
    df["Tipo Veículo"] = df["Tipo Veículo"].astype(str).str.strip().str.upper()
    df["Valor CEP"] = df["Valor CEP"].apply(to_float)
    df = df[(df["CEP Prefixo"] != "") & (df["Tipo Veículo"].isin(["MOTO", "CARRO"])) & (df["Valor CEP"] > 0)].copy()
    df = df.drop_duplicates(subset=["CEP Prefixo", "Tipo Veículo"], keep="last")
    return tuple(df[COLUNAS_REAJUSTES_CEP].itertuples(index=False, name=None))


def montar_base_cep_final(reajustes_cep_tuple: Tuple[Tuple[str, str, float], ...] = tuple()) -> pd.DataFrame:
    base = carregar_base_cep_interna().copy()
    extras = []
    for prefixo, tipo, valor in reajustes_cep_tuple:
        extras.append({
            "CEP Prefixo": normalizar_prefixo_cep(prefixo),
            "Tipo Veículo": limpar_texto(tipo).upper(),
            "Valor CEP": to_float(valor),
        })

    if extras:
        df_extra = pd.DataFrame(extras)
        base = pd.concat([base, df_extra], ignore_index=True)

    base["CEP Prefixo"] = base["CEP Prefixo"].apply(normalizar_prefixo_cep)
    base["Tipo Veículo"] = base["Tipo Veículo"].astype(str).str.strip().str.upper()
    base["Valor CEP"] = base["Valor CEP"].apply(to_float)
    base = base[(base["CEP Prefixo"] != "") & (base["Tipo Veículo"].isin(["MOTO", "CARRO"])) & (base["Valor CEP"] > 0)].copy()

    # O reajuste manual sobrescreve o valor interno quando o CEP Prefixo + Tipo Veículo forem iguais.
    return base.drop_duplicates(subset=["CEP Prefixo", "Tipo Veículo"], keep="last").reset_index(drop=True)


@st.cache_data(show_spinner=False)
def carregar_base_cep_interna() -> pd.DataFrame:
    linhas = []
    for prefixo, valor_moto, valor_carro in BASE_VALORES_CEP:
        prefixo_cep = f"0{str(prefixo).zfill(2)}"
        linhas.append({"CEP Prefixo": prefixo_cep, "Tipo Veículo": "MOTO", "Valor CEP": float(valor_moto)})
        linhas.append({"CEP Prefixo": prefixo_cep, "Tipo Veículo": "CARRO", "Valor CEP": float(valor_carro)})
    return pd.DataFrame(linhas)


@st.cache_data(show_spinner=False)
def carregar_base_placas_interna() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Placa": re.sub(r"[^A-Z0-9]", "", placa.upper()),
                "Tipo Veículo": tipo.upper().strip(),
                "Motorista Cadastro": "",
            }
            for placa, tipo in BASE_PLACAS_VEICULOS
        ]
    ).drop_duplicates()


# =========================================================
# HELPERS
# =========================================================
def normalizar_nome_coluna(col) -> str:
    texto = str(col).strip()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = texto.lower()
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto


def limpar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return re.sub(r"\s+", " ", str(valor).strip())


def normalizar_texto(valor) -> str:
    texto = limpar_texto(valor).lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def apenas_digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def normalizar_pedido_indicador(valor) -> str:
    """
    Normaliza Pedido/AWB para o indicador não contar duplicado.
    Exemplos tratados:
    - 12345678901 e 12345678901.0 viram a mesma chave;
    - AWB com espaço, hífen, ponto ou barra vira uma chave única;
    - REC-0046344 mantém REC0046344.
    """
    texto = limpar_texto(valor).upper()
    if not texto or texto in ["NAN", "NONE", "NULL"]:
        return ""

    # Remove .0 típico de número vindo do Excel.
    texto = re.sub(r"\.0$", "", texto)

    # Se vier em notação científica, tenta converter.
    if re.fullmatch(r"\d+(?:\.\d+)?E\+?\d+", texto):
        try:
            texto = str(int(float(texto)))
        except Exception:
            pass

    # Padroniza BAGREC/REC para a mesma chave.
    # Alguns arquivos podem trazer BAGREC-0046344 em uma origem
    # e REC-0046344 em outra, mas operacionalmente é o mesmo pedido.
    texto_sem_sep = re.sub(r"[^A-Z0-9]", "", texto)

    m_bagrec = re.match(r"^BAGREC0*(\d+)$", texto_sem_sep)
    if m_bagrec:
        return f"REC{m_bagrec.group(1)}"

    m_rec = re.match(r"^REC0*(\d+)$", texto_sem_sep)
    if m_rec:
        return f"REC{m_rec.group(1)}"

    # Para AWB numérica, mantém somente os dígitos.
    # Se vier com texto junto, exemplo AWB57787418365, pega a sequência numérica principal.
    somente_digitos = re.sub(r"\D", "", texto_sem_sep)
    if len(somente_digitos) >= 8 and texto_sem_sep.startswith(("AWB", "PEDIDO", "ORDER")):
        texto_sem_sep = somente_digitos

    if texto_sem_sep in ["", "NAN", "NONE", "NULL", "0"]:
        return ""

    return texto_sem_sep


def formatar_cnpj(valor) -> str:
    """Formata CNPJ com 14 dígitos no padrão 00.000.000/0000-00."""
    digitos = apenas_digitos(valor)
    if len(digitos) != 14:
        texto = limpar_texto(valor).upper()
        return texto if texto else "SEM CNPJ"
    return f"{digitos[:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:14]}"


def salvar_motoristas_cnpj_extra_df(df_extra: pd.DataFrame) -> None:
    df = df_extra.copy() if df_extra is not None else pd.DataFrame(columns=COLUNAS_MOTORISTAS_CNPJ_EXTRA)
    if "Nome Motorista" not in df.columns:
        df["Nome Motorista"] = ""
    if "CNPJ" not in df.columns:
        if "CNPJ/CPF" in df.columns:
            df["CNPJ"] = df["CNPJ/CPF"]
        else:
            df["CNPJ"] = ""
    df = df[["Nome Motorista", "CNPJ"]].copy()
    df["Nome Motorista"] = df["Nome Motorista"].astype(str).apply(limpar_texto).str.upper()
    df["CNPJ"] = df["CNPJ"].astype(str).apply(formatar_cnpj)
    df = df[(df["Nome Motorista"] != "") & (df["CNPJ"] != "") & (df["CNPJ"] != "SEM CNPJ")].copy()
    df["_nome_norm"] = df["Nome Motorista"].apply(normalizar_texto)
    df = df.drop_duplicates(subset=["_nome_norm"], keep="last")
    df = df.drop(columns=["_nome_norm"]).sort_values("Nome Motorista").reset_index(drop=True)
    df.to_csv(ARQUIVO_MOTORISTAS_CNPJ_EXTRA, index=False, encoding="utf-8-sig")

    db = carregar_db_motoristas()
    db["cnpj_extra"] = df.to_dict(orient="records")
    salvar_db_motoristas(db)


def carregar_motoristas_cnpj_excluidos_csv() -> pd.DataFrame:
    """Carrega motoristas removidos manualmente da base aplicada de CNPJ."""
    db = carregar_db_motoristas()
    if db.get("cnpj_excluidos"):
        df = pd.DataFrame(db.get("cnpj_excluidos", []))
    elif ARQUIVO_MOTORISTAS_CNPJ_EXCLUIDOS.exists():
        try:
            df = pd.read_csv(ARQUIVO_MOTORISTAS_CNPJ_EXCLUIDOS, dtype=str).fillna("")
        except Exception:
            df = pd.DataFrame(columns=COLUNAS_MOTORISTAS_CNPJ_EXCLUIDOS)
    else:
        df = pd.DataFrame(columns=COLUNAS_MOTORISTAS_CNPJ_EXCLUIDOS)

    if "Nome Motorista" not in df.columns:
        df["Nome Motorista"] = ""

    df = df[["Nome Motorista"]].copy()
    df["Nome Motorista"] = df["Nome Motorista"].astype(str).apply(limpar_texto).str.upper()
    df = df[df["Nome Motorista"] != ""].copy()
    df["_nome_norm"] = df["Nome Motorista"].apply(normalizar_texto)
    df = df.drop_duplicates(subset=["_nome_norm"], keep="last")
    return df.drop(columns=["_nome_norm"]).sort_values("Nome Motorista").reset_index(drop=True)


def salvar_motoristas_cnpj_excluidos(df_excluidos: pd.DataFrame) -> None:
    df = df_excluidos.copy() if df_excluidos is not None else pd.DataFrame(columns=COLUNAS_MOTORISTAS_CNPJ_EXCLUIDOS)
    if "Nome Motorista" not in df.columns:
        df["Nome Motorista"] = ""
    df = df[["Nome Motorista"]].copy()
    df["Nome Motorista"] = df["Nome Motorista"].astype(str).apply(limpar_texto).str.upper()
    df = df[df["Nome Motorista"] != ""].copy()
    df["_nome_norm"] = df["Nome Motorista"].apply(normalizar_texto)
    df = df.drop_duplicates(subset=["_nome_norm"], keep="last")
    df = df.drop(columns=["_nome_norm"]).sort_values("Nome Motorista").reset_index(drop=True)
    df.to_csv(ARQUIVO_MOTORISTAS_CNPJ_EXCLUIDOS, index=False, encoding="utf-8-sig")

    db = carregar_db_motoristas()
    db["cnpj_excluidos"] = df.to_dict(orient="records")
    salvar_db_motoristas(db)


def carregar_motoristas_cnpj_extra_csv() -> pd.DataFrame:
    """Carrega motoristas/CNPJ cadastrados manualmente na dashboard."""
    db = carregar_db_motoristas()
    if db.get("cnpj_extra"):
        df = pd.DataFrame(db.get("cnpj_extra", []))
    elif ARQUIVO_MOTORISTAS_CNPJ_EXTRA.exists():
        try:
            df = pd.read_csv(ARQUIVO_MOTORISTAS_CNPJ_EXTRA, dtype=str).fillna("")
        except Exception:
            df = pd.DataFrame(columns=COLUNAS_MOTORISTAS_CNPJ_EXTRA)
    else:
        df = pd.DataFrame(columns=COLUNAS_MOTORISTAS_CNPJ_EXTRA)

    if "Nome Motorista" not in df.columns:
        df["Nome Motorista"] = ""

    # Compatibilidade com arquivo antigo que usava a coluna CNPJ/CPF.
    if "CNPJ" not in df.columns:
        if "CNPJ/CPF" in df.columns:
            df["CNPJ"] = df["CNPJ/CPF"]
        else:
            df["CNPJ"] = ""

    df = df[["Nome Motorista", "CNPJ"]].copy()
    df["Nome Motorista"] = df["Nome Motorista"].astype(str).apply(limpar_texto).str.upper()
    df["CNPJ"] = df["CNPJ"].astype(str).apply(formatar_cnpj)
    df = df[(df["Nome Motorista"] != "") & (df["CNPJ"] != "") & (df["CNPJ"] != "SEM CNPJ")].copy()
    df["_nome_norm"] = df["Nome Motorista"].apply(normalizar_texto)
    df = df.drop_duplicates(subset=["_nome_norm"], keep="last")
    df = df.drop(columns=["_nome_norm"]).sort_values("Nome Motorista").reset_index(drop=True)
    return df


def salvar_motorista_cnpj_extra(nome_motorista: str, cnpj: str) -> Tuple[bool, str]:
    """Salva/atualiza motorista e CNPJ em arquivo local. O CNPJ precisa ter 14 dígitos."""
    nome = limpar_texto(nome_motorista).upper()
    digitos_cnpj = apenas_digitos(cnpj)

    if not nome:
        return False, "Informe o nome do motorista."
    if not digitos_cnpj:
        return False, "Informe o CNPJ do motorista."
    if len(digitos_cnpj) != 14:
        return False, "O CNPJ deve conter exatamente 14 dígitos. Verifique e tente novamente."

    cnpj_formatado = formatar_cnpj(digitos_cnpj)
    df = carregar_motoristas_cnpj_extra_csv()
    novo = pd.DataFrame([{"Nome Motorista": nome, "CNPJ": cnpj_formatado}])
    df = pd.concat([df, novo], ignore_index=True)
    df["_nome_norm"] = df["Nome Motorista"].apply(normalizar_texto)
    df = df.drop_duplicates(subset=["_nome_norm"], keep="last")
    df = df.drop(columns=["_nome_norm"]).sort_values("Nome Motorista").reset_index(drop=True)
    salvar_motoristas_cnpj_extra_df(df)

    # Se o motorista estava excluído, ao cadastrar novamente ele volta para a base aplicada.
    remover_motorista_da_lista_excluidos(nome)
    return True, f"Motorista {nome} cadastrado/atualizado com CNPJ {cnpj_formatado}."


def excluir_motorista_cnpj(nome_motorista: str) -> Tuple[bool, str]:
    """Exclui motorista da base aplicada de CNPJ.

    Se for cadastro manual, remove do CSV manual.
    Se estiver na base interna, grava em lista de excluídos para não aparecer na base aplicada.
    """
    nome = limpar_texto(nome_motorista).upper()
    nome_norm = normalizar_texto(nome)
    if not nome_norm:
        return False, "Informe ou selecione um motorista para excluir."

    removido_manual = False
    df_extra = carregar_motoristas_cnpj_extra_csv()
    if not df_extra.empty:
        qtd_antes = len(df_extra)
        df_extra = df_extra[df_extra["Nome Motorista"].apply(normalizar_texto) != nome_norm].copy()
        removido_manual = len(df_extra) < qtd_antes
        salvar_motoristas_cnpj_extra_df(df_extra)

    df_excluidos = carregar_motoristas_cnpj_excluidos_csv()
    if nome_norm not in set(df_excluidos["Nome Motorista"].apply(normalizar_texto)):
        df_excluidos = pd.concat([df_excluidos, pd.DataFrame([{"Nome Motorista": nome}])], ignore_index=True)
        salvar_motoristas_cnpj_excluidos(df_excluidos)

    if removido_manual:
        return True, f"Motorista {nome} removido do cadastro manual e bloqueado na base aplicada."
    return True, f"Motorista {nome} bloqueado/removido da base aplicada."


def remover_motorista_da_lista_excluidos(nome_motorista: str) -> None:
    nome_norm = normalizar_texto(nome_motorista)
    if not nome_norm or not ARQUIVO_MOTORISTAS_CNPJ_EXCLUIDOS.exists():
        return
    df = carregar_motoristas_cnpj_excluidos_csv()
    if df.empty:
        return
    df = df[df["Nome Motorista"].apply(normalizar_texto) != nome_norm].copy()
    salvar_motoristas_cnpj_excluidos(df)


def montar_base_cnpj_motoristas_final() -> pd.DataFrame:
    """Monta base final de CNPJ, juntando base interna + cadastros manuais e removendo excluídos."""
    base_interna = pd.DataFrame(
        [{"Nome Motorista": nome, "CNPJ": formatar_cnpj(cnpj), "Origem": "Base interna"} for nome, cnpj in BASE_CNPJ_MOTORISTAS.items()]
    )

    df_extra = carregar_motoristas_cnpj_extra_csv()
    if not df_extra.empty:
        df_extra = df_extra.copy()
        df_extra["Origem"] = "Cadastro manual"
        base = pd.concat([base_interna, df_extra], ignore_index=True)
    else:
        base = base_interna

    base["Nome Motorista"] = base["Nome Motorista"].astype(str).apply(limpar_texto).str.upper()
    base["CNPJ"] = base["CNPJ"].astype(str).apply(formatar_cnpj)
    base["_nome_norm"] = base["Nome Motorista"].apply(normalizar_texto)
    # O cadastro manual sobrescreve a base interna quando o nome for igual.
    base = base.drop_duplicates(subset=["_nome_norm"], keep="last")

    df_excluidos = carregar_motoristas_cnpj_excluidos_csv()
    if not df_excluidos.empty:
        excluidos_norm = set(df_excluidos["Nome Motorista"].apply(normalizar_texto))
        base = base[~base["_nome_norm"].isin(excluidos_norm)].copy()

    return base.drop(columns=["_nome_norm"]).sort_values("Nome Motorista").reset_index(drop=True)


def obter_cnpj_motorista(nome_motorista: str) -> str:
    """Busca o CNPJ do motorista na base manual e depois na base interna."""
    nome_norm = normalizar_texto(nome_motorista)
    if not nome_norm:
        return "SEM CNPJ"

    df_excluidos = carregar_motoristas_cnpj_excluidos_csv()
    if not df_excluidos.empty and nome_norm in set(df_excluidos["Nome Motorista"].apply(normalizar_texto)):
        return "SEM CNPJ"

    # Primeiro consulta o cadastro manual da dashboard.
    df_extra = carregar_motoristas_cnpj_extra_csv()
    if not df_extra.empty:
        for _, row in df_extra.iterrows():
            if normalizar_texto(row.get("Nome Motorista", "")) == nome_norm:
                cnpj = formatar_cnpj(row.get("CNPJ", ""))
                return cnpj if cnpj else "SEM CNPJ"

    # Depois consulta a base interna fixa.
    for nome_base, cnpj in BASE_CNPJ_MOTORISTAS.items():
        if normalizar_texto(nome_base) == nome_norm:
            return formatar_cnpj(cnpj)
    return "SEM CNPJ"


def renderizar_medidor_sucesso(percentual: float, titulo: str = "Sucesso nas entregas") -> str:
    """
    Renderiza um medidor semicircular em HTML/CSS.
    O ponteiro vai de -90 graus a +90 graus conforme o percentual de sucesso.
    """
    try:
        pct = float(percentual)
    except Exception:
        pct = 0.0

    pct = max(0.0, min(100.0, pct))
    angulo = -90 + (pct * 1.8)

    return f"""
    <div style="width:100%;display:flex;justify-content:center;">
        <div style="width:330px;background:#ffffff;border:1px solid #e5e7eb;border-radius:18px;padding:18px 18px 14px 18px;box-shadow:0 8px 20px rgba(15,23,42,0.05);">
            <div style="font-weight:850;color:#1f2937;text-align:center;margin-bottom:8px;font-size:16px;">{titulo}</div>
            <div style="position:relative;width:260px;height:135px;margin:0 auto;overflow:hidden;">
                <div style="position:absolute;left:0;top:0;width:260px;height:260px;border-radius:50%;
                    background:conic-gradient(from 270deg, #ef4444 0deg 55deg, #f59e0b 55deg 115deg, #84cc16 115deg 155deg, #22c55e 155deg 180deg, transparent 180deg 360deg);">
                </div>
                <div style="position:absolute;left:32px;top:32px;width:196px;height:196px;border-radius:50%;background:#ffffff;"></div>
                <div style="position:absolute;left:127px;top:18px;width:6px;height:100px;background:#111827;border-radius:8px;
                    transform-origin:50% 100%;transform:rotate({angulo}deg);box-shadow:0 2px 6px rgba(0,0,0,.25);">
                </div>
                <div style="position:absolute;left:116px;top:106px;width:28px;height:28px;border-radius:50%;background:#111827;border:4px solid #ffffff;box-shadow:0 2px 7px rgba(0,0,0,.25);"></div>
            </div>
            <div style="text-align:center;margin-top:4px;">
                <span style="font-size:34px;font-weight:900;color:#111827;">{pct:.1f}%</span>
                <span style="font-size:15px;font-weight:700;color:#64748b;"> de 100%</span>
            </div>
        </div>
    </div>
    """


def moeda(valor: float) -> str:
    try:
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def to_float(valor) -> float:
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    texto = texto.replace("R$", "").replace(" ", "")
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    texto = re.sub(r"[^0-9.\-]", "", texto)
    try:
        return float(texto)
    except Exception:
        return 0.0


def to_date(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return pd.NaT
    return pd.to_datetime(valor, dayfirst=True, errors="coerce")


def detectar_coluna(df: pd.DataFrame, candidatos: List[str]) -> Optional[str]:
    normalizadas = {normalizar_nome_coluna(c): c for c in df.columns}
    candidatos_norm = [normalizar_nome_coluna(c) for c in candidatos]

    for cand in candidatos_norm:
        if cand in normalizadas:
            return normalizadas[cand]

    for cand in candidatos_norm:
        for norm, original in normalizadas.items():
            if cand in norm or norm in cand:
                return original

    return None


def _engine_excel_mais_rapido(nome_arquivo: str):
    """Usa calamine quando instalado, pois costuma ser bem mais rápido que openpyxl."""
    try:
        import python_calamine  # noqa: F401
        return "calamine"
    except Exception:
        ext = str(nome_arquivo).lower().rsplit(".", 1)[-1]
        if ext == "xls":
            return "xlrd"
        return "openpyxl"


@st.cache_data(show_spinner=False)
def ler_excel_colunas_bytes(nome_arquivo: str, conteudo: bytes, sheet_name=0) -> List[str]:
    """Lê somente o cabeçalho para montar os campos do sidebar sem carregar a planilha inteira."""
    try:
        engine = _engine_excel_mais_rapido(nome_arquivo)
        df = pd.read_excel(
            io.BytesIO(conteudo),
            sheet_name=sheet_name,
            nrows=0,
            engine=engine,
        )
        return [limpar_texto(c) for c in df.columns]
    except Exception as e:
        st.warning(f"Não foi possível ler o cabeçalho de {nome_arquivo}: {e}")
        return []


@st.cache_data(show_spinner="Lendo somente as colunas necessárias do Excel...")
def ler_excel_bytes(nome_arquivo: str, conteudo: bytes, sheet_name=0, usecols_tuple: Tuple[str, ...] = ()) -> pd.DataFrame:
    """
    Leitura cacheada e otimizada do Excel.
    Importante: lê somente as colunas selecionadas no sidebar, em vez da planilha inteira.
    """
    try:
        engine = _engine_excel_mais_rapido(nome_arquivo)
        if usecols_tuple:
            colunas_desejadas = set(usecols_tuple)
            usecols = lambda c: limpar_texto(c) in colunas_desejadas
        else:
            usecols = None

        df = pd.read_excel(
            io.BytesIO(conteudo),
            sheet_name=sheet_name,
            engine=engine,
            usecols=usecols,
        )
        df = df.dropna(how="all").copy()
        df.columns = [limpar_texto(c) for c in df.columns]
        return df
    except Exception as e:
        st.warning(f"Não foi possível ler {nome_arquivo}: {e}")
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def listar_abas_bytes(nome_arquivo: str, conteudo: bytes) -> List[str]:
    # Para ganhar velocidade, não abrimos o workbook só para descobrir abas.
    # A dashboard usa a primeira aba por padrão.
    return ["Primeira aba"]


@st.cache_data(show_spinner="Lendo PDFs uma única vez...")
def extrair_texto_pdf_bytes(nome_arquivo: str, conteudo: bytes) -> str:
    if fitz is None:
        st.error("PyMuPDF não instalado. Rode: py -m pip install pymupdf")
        return ""

    try:
        doc = fitz.open(stream=conteudo, filetype="pdf")
        partes = []
        for i, page in enumerate(doc, start=1):
            texto = page.get_text("text") or ""
            partes.append(f"\n--- PAGINA {i} ---\n{texto}")
        doc.close()
        return "\n".join(partes)
    except Exception as e:
        st.warning(f"Não foi possível ler o PDF {nome_arquivo}: {e}")
        return ""


def ler_excel_upload(uploaded_file, sheet_name=None) -> pd.DataFrame:
    uploaded_file.seek(0)
    return ler_excel_bytes(uploaded_file.name, uploaded_file.read(), sheet_name=sheet_name)


def listar_abas(uploaded_file) -> List[str]:
    uploaded_file.seek(0)
    return listar_abas_bytes(uploaded_file.name, uploaded_file.read())


def extrair_texto_pdf(uploaded_file) -> str:
    uploaded_file.seek(0)
    return extrair_texto_pdf_bytes(uploaded_file.name, uploaded_file.read())


def carregar_excel_sistema_otimizado(excel_payloads: Tuple[Tuple[str, bytes], ...], colunas_necessarias: Tuple[str, ...]) -> pd.DataFrame:
    """Carrega todos os Excels usando somente as colunas necessárias para o fechamento."""
    frames = []
    for nome_arquivo, conteudo in excel_payloads:
        df = ler_excel_bytes(nome_arquivo, conteudo, sheet_name=0, usecols_tuple=colunas_necessarias)
        if not df.empty:
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()



def extrair_info_pdf(nome_arquivo: str, texto: str) -> Dict[str, object]:
    texto_limpo = re.sub(r"\s+", " ", texto)

    data = None
    m_data = re.search(r"(\d{2}/\d{2}/\d{4}|\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2})", texto_limpo)
    if m_data:
        data = pd.to_datetime(m_data.group(1), dayfirst=True, errors="coerce")

    rota = ""
    for padrao in [
        r"(?:manifesto|rota|route|numero da rota|n[ºo]\s*rota)\s*[:\-]?\s*([A-Za-z0-9\-_./]+)",
        r"\b(RD\d{6,})\b",
        r"\bRT\s*[:\-]?\s*([A-Za-z0-9\-_./]+)",
    ]:
        m = re.search(padrao, texto_limpo, flags=re.I)
        if m:
            rota = m.group(1).strip().upper()
            break

    placa = ""
    m_placa = re.search(r"\b([A-Z]{3}[0-9][A-Z0-9][0-9]{2})\b", texto_limpo.upper())
    if m_placa:
        placa = m_placa.group(1).strip().upper()

    motorista = ""
    # No PDF da lista de entregas, o motorista vem logo após o cabeçalho:
    # Motorista Placa Data e Horário de Saída\nNOME PLACA DATA
    if placa:
        padrao_motorista_placa = rf"Motorista\s+Placa\s+Data\s+e\s+Hor[aá]rio\s+de\s+Sa[ií]da\s+([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ ]{{4,90}})\s+{placa}"
        m = re.search(padrao_motorista_placa, texto_limpo, flags=re.I)
        if m:
            motorista = limpar_texto(m.group(1)).upper()

    if not motorista:
        for padrao in [
            r"(?:motorista|entregador|driver)\s*[:\-]?\s*([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ ]{4,80})",
            r"(?:nome)\s*[:\-]?\s*([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ ]{4,80})",
        ]:
            m = re.search(padrao, texto_limpo, flags=re.I)
            if m:
                motorista = limpar_texto(m.group(1)).upper()
                break

    return {
        "Arquivo PDF": nome_arquivo,
        "Texto PDF": texto,
        "Data PDF": data,
        "Rota PDF": rota,
        "Placa PDF": placa,
        "Motorista PDF": motorista,
    }


def extrair_itens_pdf(nome_arquivo: str, texto: str, info_pdf: Dict[str, object]) -> pd.DataFrame:
    """
    Extrai do PDF os dados por AWB, principalmente o P. Taxado.

    No PDF, a linha vem neste padrão aproximado:
    AWB Volumes Peso Real P. Taxado Origem Manu. Esp. Modalidade CEP ...
    Exemplo:
    57787418365 2 12,37 13,00 WINE1 NOR PX 04194260 ...
    """
    texto_flat = re.sub(r"\s+", " ", texto)

    padrao = re.compile(
        r"\b(?P<awb>\d{11})\s+"
        r"(?P<volumes>\d+)\s+"
        r"(?P<peso_real>\d{1,4}(?:[\.,]\d{1,3})?)\s+"
        r"(?P<p_taxado>\d{1,4}(?:[\.,]\d{1,3})?)\s+"
        r"(?P<origem>[A-Z0-9]{2,12})\s+"
        r"(?P<manu>[A-Z0-9]{2,6})\s+"
        r"(?P<modalidade>PX|PP|EXPRESSO|ECOMM|PREMIUM)\s+"
        r"(?P<cep>\d{8})\b",
        flags=re.I,
    )

    linhas = []
    def adicionar_linha_pdf(pedido, volumes, peso_real, p_taxado, cep):
        linhas.append({
            "Arquivo PDF": nome_arquivo,
            "Pedido": str(pedido).strip(),
            "Rota PDF": str(info_pdf.get("Rota PDF", "")).strip().upper(),
            "Data PDF": info_pdf.get("Data PDF"),
            "Placa PDF": str(info_pdf.get("Placa PDF", "")).strip().upper(),
            "Motorista PDF": str(info_pdf.get("Motorista PDF", "")).strip().upper(),
            "Volumes PDF": int(volumes),
            "Peso Real PDF": to_float(peso_real),
            "P Taxado PDF": to_float(p_taxado),
            "CEP PDF": re.sub(r"\D", "", str(cep)).zfill(8),
        })

    for m in padrao.finditer(texto_flat):
        adicionar_linha_pdf(
            m.group("awb").strip(),
            m.group("volumes"),
            m.group("peso_real"),
            m.group("p_taxado"),
            m.group("cep"),
        )

    # Alguns PDFs trazem item BAGREC/REC como entrega fechada no Excel.
    # Exemplo no PDF: BAGREC-0046344; no Excel: REC-0046344.
    padrao_bagrec = re.compile(
        r"\bBAGREC\s*[-\uFFFE]?\s*(?P<num>\d{5,})\s+"
        r"(?P<volumes>\d+)\s+"
        r"(?P<peso_real>\d{1,4}(?:[\.,]\d{1,3})?)\s+"
        r"(?P<p_taxado>\d{1,4}(?:[\.,]\d{1,3})?)\s+"
        r"(?P<origem>[A-Z0-9]{2,12})\s+"
        r"(?P<manu>[A-Z0-9]{2,6})\s+"
        r"(?P<modalidade>PX|PP|EXPRESSO|ECOMM|PREMIUM)\s+"
        r"(?P<cep>\d{8})\b",
        flags=re.I,
    )

    for m in padrao_bagrec.finditer(texto_flat):
        adicionar_linha_pdf(
            f"REC-{m.group('num').strip()}",
            m.group("volumes"),
            m.group("peso_real"),
            m.group("p_taxado"),
            m.group("cep"),
        )

    df = pd.DataFrame(linhas)
    # Não remover duplicidades aqui.
    # A quantidade do fechamento deve seguir as entregas do PDF/manifesto.
    # Se a mesma AWB aparecer mais de uma vez no manifesto, cada ocorrência válida conta como entrega.
    return df


@st.cache_data(show_spinner="Processando PDFs...")
def processar_pdf_cache(nome_arquivo: str, conteudo: bytes) -> Tuple[Dict[str, object], pd.DataFrame]:
    texto = extrair_texto_pdf_bytes(nome_arquivo, conteudo)
    info = extrair_info_pdf(nome_arquivo, texto)
    itens_pdf = extrair_itens_pdf(nome_arquivo, texto, info)
    return info, itens_pdf


@st.cache_data(show_spinner="Buscando motorista e placa nos PDFs...")
def extrair_motoristas_placas_dos_pdfs_cache(
    pdf_payloads_tuple: Tuple[Tuple[str, bytes], ...]
) -> Tuple[Tuple[str, str, str, str], ...]:
    """
    Lê os PDFs carregados e retorna uma base simples para consulta:
    Motorista, Placa, Rota e Arquivo PDF.

    Essa informação é usada para preencher automaticamente o nome do motorista
    na tela "Ver todos os motoristas da base", quando a placa existir no PDF.
    """
    registros = []
    for nome_arquivo, conteudo in pdf_payloads_tuple:
        try:
            info, _itens = processar_pdf_cache(nome_arquivo, conteudo)
        except Exception:
            continue

        placa = normalizar_placa(info.get("Placa PDF", ""))
        motorista = limpar_texto(info.get("Motorista PDF", "")).upper()
        rota = limpar_texto(info.get("Rota PDF", "")).upper()

        if placa and motorista:
            registros.append((motorista, placa, rota, nome_arquivo))

    if not registros:
        return tuple()

    df = pd.DataFrame(registros, columns=["Motorista PDF", "Placa", "Rota PDF", "Arquivo PDF"])
    df["Placa"] = df["Placa"].apply(normalizar_placa)
    df["Motorista PDF"] = df["Motorista PDF"].astype(str).str.strip().str.upper()
    df = df[(df["Placa"] != "") & (df["Motorista PDF"] != "")].copy()
    df = df.drop_duplicates(subset=["Placa"], keep="last").reset_index(drop=True)
    return tuple(df[["Motorista PDF", "Placa", "Rota PDF", "Arquivo PDF"]].itertuples(index=False, name=None))


def preparar_tabela_cep(df_cep: pd.DataFrame, col_cep, col_tipo, col_valor) -> pd.DataFrame:
    df = df_cep.copy()
    df["CEP Base"] = df[col_cep].astype(str).apply(lambda x: re.sub(r"\D", "", x).zfill(8))
    df["Tipo Veículo"] = df[col_tipo].astype(str).str.upper().str.strip()
    df["Valor CEP"] = df[col_valor].apply(to_float)
    return df[["CEP Base", "Tipo Veículo", "Valor CEP"]].drop_duplicates()


def preparar_tabela_placas(df_placas: pd.DataFrame, col_placa, col_tipo, col_motorista=None) -> pd.DataFrame:
    df = df_placas.copy()
    df["Placa"] = df[col_placa].astype(str).str.upper().str.replace("-", "", regex=False).str.strip()
    df["Tipo Veículo"] = df[col_tipo].astype(str).str.upper().str.strip()
    if col_motorista and col_motorista in df.columns:
        df["Motorista Cadastro"] = df[col_motorista].astype(str).str.strip()
    else:
        df["Motorista Cadastro"] = ""
    return df[["Placa", "Tipo Veículo", "Motorista Cadastro"]].drop_duplicates()


def preparar_planilha_sistema(
    df: pd.DataFrame,
    col_pedido: str,
    col_status: str,
    col_data: str,
    col_cep: str,
    col_rota: str,
    col_peso: str,
    col_placa: Optional[str] = None,
    col_motorista: Optional[str] = None,
) -> pd.DataFrame:
    out = df.copy()
    out["Pedido"] = out[col_pedido].astype(str).str.strip()
    out["Status"] = out[col_status].astype(str).str.strip()
    out["Status Normalizado"] = out["Status"].apply(normalizar_texto)
    out["Data Rota"] = out[col_data].apply(to_date).dt.date
    out["CEP"] = out[col_cep].astype(str).apply(lambda x: re.sub(r"\D", "", x).zfill(8))
    out["Rota"] = out[col_rota].astype(str).str.strip()
    out["Peso Taxado KG"] = out[col_peso].apply(to_float)
    out["Placa"] = out[col_placa].astype(str).str.upper().str.replace("-", "", regex=False).str.strip() if col_placa else ""
    out["Motorista"] = out[col_motorista].astype(str).str.strip() if col_motorista else ""
    return out


def preparar_planilha_status_cep(
    df: pd.DataFrame,
    col_pedido: str,
    col_status: str,
    col_motivo_indicador: Optional[str],
    col_cep: str,
    col_rota: Optional[str] = None,
) -> pd.DataFrame:
    """
    Prepara o Excel para validação do pagamento.

    Regra atual:
    - Excel valida Pedido/AWB + Status + CEP;
    - Quando existir coluna de rota/carga, ela também é usada para evitar que uma AWB
      fechada em outra rota entre no PDF errado.
    """
    out = pd.DataFrame()
    out["Pedido"] = df[col_pedido].astype(str).str.strip()
    out["Status"] = df[col_status].astype(str).str.strip()
    out["Status Normalizado"] = out["Status"].apply(normalizar_texto)
    out["CEP Excel"] = df[col_cep].astype(str).apply(lambda x: re.sub(r"\D", "", x).zfill(8))

    if col_rota and col_rota in df.columns:
        out["Rota Excel"] = df[col_rota].astype(str).str.strip().str.upper()
    else:
        out["Rota Excel"] = ""

    out = out[(out["Pedido"].astype(str).str.strip() != "") & (out["Pedido"].astype(str).str.lower() != "nan")]
    return out


def montar_conferencia_awb_pdf_excel(df_pdf_itens: pd.DataFrame, df_status_cep: pd.DataFrame) -> Dict[str, object]:
    """
    Confere se todas as AWBs/Pedidos presentes nos PDFs também existem no Excel carregado.
    Esta validação serve apenas como aviso visual para evitar fechamento com Excel incompleto.
    """
    resumo = {
        "total_awb_pdf": 0,
        "total_awb_excel_encontradas": 0,
        "total_awb_nao_encontradas": 0,
        "awbs_nao_encontradas": [],
        "df_awbs_nao_encontradas": pd.DataFrame(columns=["AWB não encontrada no Excel", "Arquivo PDF", "Rota PDF", "Motorista PDF"]),
    }

    if df_pdf_itens is None or df_pdf_itens.empty or "Pedido" not in df_pdf_itens.columns:
        return resumo

    pdf = df_pdf_itens.copy()
    pdf["AWB Normalizada"] = pdf["Pedido"].apply(normalizar_pedido_indicador)
    pdf = pdf[pdf["AWB Normalizada"].astype(str).str.strip() != ""].copy()
    pdf = pdf.drop_duplicates(subset=["AWB Normalizada"], keep="first")

    resumo["total_awb_pdf"] = int(len(pdf))

    if df_status_cep is None or df_status_cep.empty or "Pedido" not in df_status_cep.columns:
        faltantes = pdf.copy()
    else:
        excel = df_status_cep.copy()
        excel["AWB Normalizada"] = excel["Pedido"].apply(normalizar_pedido_indicador)
        awbs_excel = set(excel["AWB Normalizada"].dropna().astype(str).str.strip())
        awbs_excel.discard("")
        faltantes = pdf[~pdf["AWB Normalizada"].isin(awbs_excel)].copy()

    resumo["total_awb_nao_encontradas"] = int(len(faltantes))
    resumo["total_awb_excel_encontradas"] = int(max(0, resumo["total_awb_pdf"] - resumo["total_awb_nao_encontradas"]))

    if not faltantes.empty:
        faltantes_display = faltantes.copy()
        faltantes_display["AWB não encontrada no Excel"] = faltantes_display["Pedido"].astype(str).str.strip()
        colunas_faltantes = [
            "AWB não encontrada no Excel",
            "Arquivo PDF",
            "Rota PDF",
            "Motorista PDF",
        ]
        colunas_faltantes = [c for c in colunas_faltantes if c in faltantes_display.columns]
        faltantes_display = faltantes_display[colunas_faltantes].reset_index(drop=True)
        resumo["awbs_nao_encontradas"] = faltantes_display["AWB não encontrada no Excel"].astype(str).tolist()
        resumo["df_awbs_nao_encontradas"] = faltantes_display

    return resumo


def renderizar_aviso_conferencia_awb(conferencia_awb: Dict[str, object]) -> None:
    """Mostra no topo da dashboard o status da conferência AWB PDF x Excel."""
    if not conferencia_awb:
        return

    total_pdf = int(conferencia_awb.get("total_awb_pdf", 0) or 0)
    total_excel_encontradas = int(conferencia_awb.get("total_awb_excel_encontradas", 0) or 0)
    total_faltantes = int(conferencia_awb.get("total_awb_nao_encontradas", 0) or 0)
    df_faltantes = conferencia_awb.get("df_awbs_nao_encontradas", pd.DataFrame())

    if total_pdf <= 0:
        return

    if total_faltantes > 0:
        st.markdown(
            f"""
            <div style="background:#fff1f2;border:1px solid #fecdd3;border-left:8px solid #dc2626;border-radius:18px;padding:18px 20px;margin:0 0 22px 0;box-shadow:0 8px 20px rgba(15,23,42,0.05);">
                <div style="font-size:20px;font-weight:900;color:#991b1b;margin-bottom:10px;">🔴 ATENÇÃO — DIVERGÊNCIA ENTRE PDF E EXCEL</div>
                <div style="font-size:15px;color:#7f1d1d;line-height:1.7;">
                    <b>AWBs no PDF:</b> {total_pdf}<br>
                    <b>AWBs do PDF encontradas no Excel:</b> {total_excel_encontradas}<br>
                    <b>AWBs não encontradas no Excel:</b> {total_faltantes}
                </div>
                <div style="margin-top:10px;font-size:14px;color:#7f1d1d;">
                    ⚠️ O Excel carregado não contém todas as entregas do PDF. Isso pode gerar valores incorretos no fechamento.
                    Baixe/exporte o Excel completo antes de validar o pagamento.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if isinstance(df_faltantes, pd.DataFrame) and not df_faltantes.empty:
            with st.expander(f"Ver AWBs não encontradas no Excel ({total_faltantes})", expanded=True):
                st.dataframe(df_faltantes, use_container_width=True, hide_index=True, height=min(300, 80 + (35 * total_faltantes)))
    else:
        st.markdown(
            f"""
            <div style="background:#ecfdf5;border:1px solid #bbf7d0;border-left:8px solid #16a34a;border-radius:18px;padding:16px 20px;margin:0 0 22px 0;box-shadow:0 8px 20px rgba(15,23,42,0.05);">
                <div style="font-size:20px;font-weight:900;color:#166534;margin-bottom:8px;">🟢 CONFERÊNCIA DE DADOS OK</div>
                <div style="font-size:15px;color:#14532d;line-height:1.7;">
                    <b>AWBs no PDF:</b> {total_pdf}<br>
                    <b>AWBs do PDF encontradas no Excel:</b> {total_excel_encontradas}<br>
                    <b>AWBs não encontradas no Excel:</b> 0
                </div>
                <div style="margin-top:8px;font-size:14px;color:#14532d;">
                    ✅ Todas as AWBs dos PDFs foram localizadas no Excel carregado.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def filtrar_pedidos_pagos_excel(df: pd.DataFrame, status_entregue: List[str], status_ocorrencia: List[str]) -> pd.DataFrame:
    """
    Valida pagamento usando apenas o Excel:
    - Pedido precisa ter status de entregue;
    - Pedido não pode ter status de ocorrência/não pagar.
    """
    entregue_norm = [normalizar_texto(x) for x in status_entregue if str(x).strip()]
    ocorrencia_norm = [normalizar_texto(x) for x in status_ocorrencia if str(x).strip()]

    df = df.copy()
    df["É Entregue"] = df["Status Normalizado"].apply(lambda x: any(s in x for s in entregue_norm)) if entregue_norm else False
    df["É Ocorrência"] = df["Status Normalizado"].apply(lambda x: any(s in x for s in ocorrencia_norm)) if ocorrencia_norm else False

    chaves_grupo = ["Pedido"]
    if "Rota Excel" in df.columns and df["Rota Excel"].astype(str).str.strip().ne("").any():
        chaves_grupo = ["Rota Excel", "Pedido"]

    grupo = (
        df.groupby(chaves_grupo, dropna=False)
        .agg(
            Tem_Entregue=("É Entregue", "max"),
            Tem_Ocorrencia=("É Ocorrência", "max"),
            CEP_Excel=("CEP Excel", "last"),
            Status_Encontrados=("Status", lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip()])))),
        )
        .reset_index()
    )

    if "Rota Excel" not in grupo.columns:
        grupo["Rota Excel"] = ""
    grupo["Entrega Paga"] = grupo["Tem_Entregue"] & (~grupo["Tem_Ocorrencia"])
    return grupo[grupo["Entrega Paga"]].copy()


def montar_entregas_pagas_pdf(df_pedidos_pagos: pd.DataFrame, df_pdf_itens: pd.DataFrame) -> pd.DataFrame:
    """Monta a base final puxando rota/data/motorista/placa/peso dos PDFs e status/CEP do Excel."""
    if df_pdf_itens.empty or df_pedidos_pagos.empty:
        return pd.DataFrame()

    itens = df_pdf_itens.copy()
    itens["Pedido"] = itens["Pedido"].astype(str).str.strip()
    itens["Rota PDF"] = itens["Rota PDF"].astype(str).str.strip().str.upper()
    # Não remover duplicidade por Rota + Pedido.
    # A quantidade correta é a quantidade de linhas/entregas do PDF com status fechado no Excel.

    pedidos = df_pedidos_pagos.copy()
    pedidos["Pedido"] = pedidos["Pedido"].astype(str).str.strip()
    pedidos["Rota Excel"] = pedidos.get("Rota Excel", "").fillna("").astype(str).str.strip().str.upper()

    if pedidos["Rota Excel"].ne("").any():
        base = pedidos.merge(
            itens,
            how="inner",
            left_on=["Rota Excel", "Pedido"],
            right_on=["Rota PDF", "Pedido"],
        )
    else:
        base = pedidos.merge(itens, how="inner", on="Pedido")

    if base.empty:
        return pd.DataFrame()

    base["Data Rota"] = pd.to_datetime(base["Data PDF"], errors="coerce").dt.date
    base["Rota"] = base["Rota PDF"].fillna("").astype(str).str.strip()
    base["CEP"] = base["CEP PDF"].fillna("").astype(str).str.replace(r"\D", "", regex=True).str.zfill(8)
    cep_excel = base["CEP_Excel"].fillna("").astype(str).str.replace(r"\D", "", regex=True).str.zfill(8)
    base["CEP"] = base["CEP"].where((base["CEP"].ne("00000000")) & (base["CEP"].str.strip().ne("")), cep_excel)
    base["Placa"] = base["Placa PDF"].fillna("").astype(str).str.upper().str.replace("-", "", regex=False).str.strip()
    base["Motorista"] = base["Motorista PDF"].fillna("").astype(str).str.strip()
    base["Peso_Taxado_KG"] = base["P Taxado PDF"].fillna(0).astype(float)
    base["Peso Taxado Cálculo"] = base["Peso_Taxado_KG"]
    base["Fonte Peso Taxado"] = "PDF"
    base["Valor Excel Usado"] = "Pedido + Status + CEP"

    return base


def remover_awbs_duplicadas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Garante que cada AWB/Pedido seja pago somente uma vez.

    Regra aplicada:
    - remove Pedido vazio/nan;
    - normaliza o Pedido/AWB como texto;
    - se a mesma AWB aparecer em mais de um PDF, rota ou data, mantém somente a primeira ocorrência;
    - a primeira ocorrência é definida pela menor data e, em seguida, pela ordem de arquivo/rota.
    """
    if df.empty or "Pedido" not in df.columns:
        return df.copy()

    out = df.copy()
    out["Pedido"] = out["Pedido"].astype(str).str.strip()
    out = out[(out["Pedido"] != "") & (out["Pedido"].str.lower() != "nan")].copy()

    colunas_ordem = [c for c in ["Data Rota", "Arquivo PDF", "Rota", "Pedido"] if c in out.columns]
    if colunas_ordem:
        out = out.sort_values(colunas_ordem, kind="mergesort")

    out = out.drop_duplicates(subset=["Pedido"], keep="first").copy()
    return out


def filtrar_entregas_validas(df: pd.DataFrame, status_entregue: List[str], status_ocorrencia: List[str]) -> pd.DataFrame:
    entregue_norm = [normalizar_texto(x) for x in status_entregue if str(x).strip()]
    ocorrencia_norm = [normalizar_texto(x) for x in status_ocorrencia if str(x).strip()]

    df = df.copy()
    df["É Entregue"] = df["Status Normalizado"].apply(lambda x: any(s in x for s in entregue_norm)) if entregue_norm else False
    df["É Ocorrência"] = df["Status Normalizado"].apply(lambda x: any(s in x for s in ocorrencia_norm)) if ocorrencia_norm else False

    # Validação por ROTA x DIA x PEDIDO.
    grupo = (
        df.groupby(["Data Rota", "Rota", "Pedido"], dropna=False)
        .agg(
            Tem_Entregue=("É Entregue", "max"),
            Tem_Ocorrencia=("É Ocorrência", "max"),
            CEP=("CEP", "last"),
            Peso_Taxado_KG=("Peso Taxado KG", "max"),
            Placa=("Placa", "last"),
            Motorista=("Motorista", "last"),
            Status_Encontrados=("Status", lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip()])))),
        )
        .reset_index()
    )

    # Pago somente entregue sem ocorrência dentro da mesma rota/dia/pedido.
    grupo["Entrega Paga"] = grupo["Tem_Entregue"] & (~grupo["Tem_Ocorrencia"])
    return grupo[grupo["Entrega Paga"]].copy()


def enriquecer_com_pdf(
    df_validas: pd.DataFrame,
    df_pdf_info: pd.DataFrame,
    df_pdf_itens: pd.DataFrame,
) -> pd.DataFrame:
    df = df_validas.copy()
    pdf_info = df_pdf_info.copy()
    pdf_itens = df_pdf_itens.copy()

    if pdf_info.empty:
        df["Arquivo PDF"] = ""
        df["Motorista PDF"] = ""
        df["P Taxado PDF"] = 0.0
        df["Peso Taxado Cálculo"] = df["Peso_Taxado_KG"].fillna(0).astype(float)
        df["Fonte Peso Taxado"] = "Excel"
        return df

    if "Data PDF" in pdf_info.columns:
        pdf_info["Data PDF Date"] = pd.to_datetime(pdf_info["Data PDF"], errors="coerce").dt.date
    else:
        pdf_info["Data PDF Date"] = pd.NaT

    # Primeiro cruza informações gerais do PDF por Data + Rota.
    df = df.merge(
        pdf_info[["Arquivo PDF", "Data PDF Date", "Rota PDF", "Placa PDF", "Motorista PDF"]],
        how="left",
        left_on=["Data Rota", "Rota"],
        right_on=["Data PDF Date", "Rota PDF"],
    )

    # Depois cruza os itens do PDF por Rota + Pedido/AWB para puxar P. Taxado oficial do PDF.
    if not pdf_itens.empty:
        pdf_itens = pdf_itens.copy()
        pdf_itens["Pedido"] = pdf_itens["Pedido"].astype(str).str.strip()
        pdf_itens["Rota PDF"] = pdf_itens["Rota PDF"].astype(str).str.strip().str.upper()
        # Não remover duplicidade por Rota + Pedido.
        # Cada ocorrência válida no PDF representa uma entrega no fechamento.

        df = df.merge(
            pdf_itens[["Rota PDF", "Pedido", "P Taxado PDF", "Peso Real PDF", "CEP PDF"]],
            how="left",
            left_on=["Rota", "Pedido"],
            right_on=["Rota PDF", "Pedido"],
            suffixes=("", "_ItemPDF"),
        )
    else:
        df["P Taxado PDF"] = 0.0
        df["Peso Real PDF"] = 0.0
        df["CEP PDF"] = ""

    # Preenche placa/motorista vindo do PDF quando estiverem vazios.
    placa_atual = df["Placa"].fillna("").astype(str).str.strip()
    placa_pdf = df.get("Placa PDF", "").fillna("").astype(str).str.strip() if "Placa PDF" in df.columns else ""
    df["Placa"] = placa_atual.where(placa_atual.ne(""), placa_pdf)

    motorista_atual = df["Motorista"].fillna("").astype(str).str.strip()
    motorista_pdf = df.get("Motorista PDF", "").fillna("").astype(str).str.strip() if "Motorista PDF" in df.columns else ""
    df["Motorista"] = motorista_atual.where(motorista_atual.ne(""), motorista_pdf)

    # CEP oficial para pagamento: prioriza o CEP do PDF quando existir.
    cep_excel = df["CEP"].fillna("").astype(str).str.replace(r"\D", "", regex=True).str.zfill(8)
    cep_pdf = df.get("CEP PDF", "").fillna("").astype(str).str.replace(r"\D", "", regex=True).str.zfill(8) if "CEP PDF" in df.columns else ""
    usar_cep_pdf = cep_pdf.ne("") & cep_pdf.ne("00000000") if hasattr(cep_pdf, "ne") else False
    df["CEP"] = cep_excel.where(~usar_cep_pdf, cep_pdf)

    # PESO TAXADO OFICIAL: prioriza P. Taxado do PDF. Excel fica só como fallback.
    df["P Taxado PDF"] = df["P Taxado PDF"].fillna(0).astype(float)
    peso_excel = df["Peso_Taxado_KG"].fillna(0).astype(float)
    usar_pdf = df["P Taxado PDF"].gt(0)
    df["Peso Taxado Cálculo"] = peso_excel.where(~usar_pdf, df["P Taxado PDF"])
    df["Fonte Peso Taxado"] = "Excel"
    df.loc[usar_pdf, "Fonte Peso Taxado"] = "PDF"

    return df


def calcular_pagamento(df: pd.DataFrame, df_cep: pd.DataFrame, df_placas: pd.DataFrame) -> pd.DataFrame:
    base = df.copy()
    base["Placa"] = base["Placa"].astype(str).str.upper().str.replace("-", "", regex=False).str.strip()

    if not df_placas.empty:
        base = base.merge(df_placas, how="left", on="Placa")

        # Quando a placa não existe na base interna, antes o valor zerava.
        # Agora a dashboard assume CARRO para não deixar a entrega sem pagamento.
        base["Tipo Veículo Final"] = base["Tipo Veículo"].fillna("").astype(str).str.upper().str.strip()
        base["Tipo Veículo Assumido"] = base["Tipo Veículo Final"].eq("")
        base["Tipo Veículo Final"] = base["Tipo Veículo Final"].replace("", "CARRO")

        base["Motorista Final"] = base["Motorista"].where(
            base["Motorista"].astype(str).str.strip() != "",
            base["Motorista Cadastro"],
        )
    else:
        base["Tipo Veículo Final"] = "CARRO"
        base["Tipo Veículo Assumido"] = True
        base["Motorista Final"] = base["Motorista"]

    # A base de valores é pelo começo do CEP com 0 na frente.
    # Exemplo: CEP 01045-000 => prefixo 010.
    base["CEP Prefixo"] = base["CEP"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(8).str[:3]

    if not df_cep.empty:
        base = base.merge(
            df_cep,
            how="left",
            left_on=["CEP Prefixo", "Tipo Veículo Final"],
            right_on=["CEP Prefixo", "Tipo Veículo"],
            suffixes=("", "_CEP"),
        )
    else:
        base["Valor CEP"] = 0.0

    base["Valor CEP"] = base["Valor CEP"].fillna(0).astype(float)

    # REGRA OPERACIONAL DO KG EXCEDENTE:
    # - Só cobra o que passar de 10 kg do peso taxado;
    # - Qualquer fração deve arredondar sempre para cima.
    # Exemplos:
    #   Peso taxado 11,10 kg => excedente 1,10 kg => cobra 2 kg
    #   Peso taxado 10,01 kg => excedente 0,01 kg => cobra 1 kg
    #   Peso taxado 34,50 kg => excedente 24,50 kg => cobra 25 kg
    peso_taxado_calculo = base["Peso Taxado Cálculo"].fillna(0).astype(float)
    base["KG Excedente"] = (peso_taxado_calculo - 10).clip(lower=0).apply(math.ceil).astype(int)
    base["Valor Excedente KG"] = base["KG Excedente"] * 0.30
    base["Total Entrega"] = base["Valor CEP"] + base["Valor Excedente KG"]
    return base


def gerar_fechamento_diario(df_pagamento: pd.DataFrame) -> pd.DataFrame:
    if df_pagamento.empty:
        return pd.DataFrame()

    resumo = (
        df_pagamento.groupby(["Motorista Final", "Data Rota"], dropna=False)
        .agg(
            Rotas=("Rota", lambda x: ", ".join(sorted(set(map(str, x))))),
            Quantidade_Entregas=("Pedido", "size"),
            Valor_Entregas=("Valor CEP", "sum"),
            KG_Excedente_Calculado=("KG Excedente", "sum"),
            Valor_KG_Excedente=("Valor Excedente KG", "sum"),
            Total_Dia=("Total Entrega", "sum"),
        )
        .reset_index()
        .sort_values(["Motorista Final", "Data Rota"])
    )
    return resumo



def aplicar_rd_fechada_recibo(df_base: pd.DataFrame, rds_fechadas: List[str], valor_rd_fechada: float = 250.0) -> pd.DataFrame:
    """Substitui as rotas selecionadas por uma linha única de RD Fechada.

    Regra: quando uma RD é fechada, ela passa a valer somente o valor fixo informado.
    Não considera kg excedente nem valores por entrega daquela RD.
    """
    if df_base is None or df_base.empty or not rds_fechadas:
        return df_base.copy() if df_base is not None else pd.DataFrame()

    df = df_base.copy()
    if "Rota" not in df.columns:
        return df

    rds_set = {str(rd).strip().upper() for rd in rds_fechadas if str(rd).strip()}
    if not rds_set:
        return df

    df["Rota"] = df["Rota"].astype(str).str.strip().str.upper()
    df_normais = df[~df["Rota"].isin(rds_set)].copy()
    linhas_rd = []

    for rd in sorted(rds_set):
        df_rd = df[df["Rota"] == rd].copy()
        if df_rd.empty:
            continue

        base_row = df_rd.iloc[0].copy()
        base_row["Pedido"] = f"RD FECHADA - {rd}"
        base_row["Status_Encontrados"] = "RD Fechada"
        base_row["CEP"] = ""
        base_row["CEP Prefixo"] = ""
        base_row["Valor CEP"] = float(valor_rd_fechada)
        base_row["KG Excedente"] = 0
        base_row["Valor Excedente KG"] = 0.0
        base_row["Total Entrega"] = float(valor_rd_fechada)
        base_row["Peso Taxado Cálculo"] = 0.0
        base_row["Peso_Taxado_KG"] = 0.0
        base_row["P Taxado PDF"] = 0.0
        base_row["Peso Real PDF"] = 0.0
        base_row["Descrição Relatório"] = "RD Fechada"
        base_row["Quantidade RD Fechada"] = int(len(df_rd))
        base_row["Valor RD Fechada"] = float(valor_rd_fechada)
        linhas_rd.append(base_row)

    if linhas_rd:
        df_rd_final = pd.DataFrame(linhas_rd)
        df = pd.concat([df_normais, df_rd_final], ignore_index=True)
    else:
        df = df_normais

    if "Data Rota" in df.columns:
        df = df.sort_values(["Data Rota", "Rota", "Pedido"], kind="mergesort")
    return df.reset_index(drop=True)


def substituir_fechamento_diario_por_recibo(
    df_dia_base: pd.DataFrame,
    df_recibo_modificado: pd.DataFrame,
    motorista: str,
    data_inicio,
    data_fim,
) -> pd.DataFrame:
    """Atualiza no fechamento diário o período do motorista usando a base já ajustada com RD Fechada."""
    if df_dia_base is None or df_dia_base.empty or df_recibo_modificado is None or df_recibo_modificado.empty:
        return df_dia_base.copy() if df_dia_base is not None else pd.DataFrame()

    df_out = df_dia_base.copy()
    motorista_norm = str(motorista).upper().strip()
    inicio = pd.to_datetime(data_inicio, errors="coerce").date()
    fim = pd.to_datetime(data_fim, errors="coerce").date()

    datas_out = pd.to_datetime(df_out["Data Rota"], errors="coerce").dt.date
    mask_periodo = (
        (df_out["Motorista Final"].astype(str).str.upper().str.strip() == motorista_norm)
        & (datas_out >= inicio)
        & (datas_out <= fim)
    )
    df_out = df_out[~mask_periodo].copy()

    df_mod = df_recibo_modificado.copy()
    df_mod["Data Rota"] = pd.to_datetime(df_mod["Data Rota"], errors="coerce").dt.date
    resumo_mod = gerar_fechamento_diario(df_mod)

    if not resumo_mod.empty:
        df_out = pd.concat([df_out, resumo_mod], ignore_index=True)
        df_out = df_out.sort_values(["Motorista Final", "Data Rota"], kind="mergesort").reset_index(drop=True)
    return df_out

def preparar_base_bonus_excel(
    df_excel_raw: pd.DataFrame,
    col_pedido: str,
    col_status: str,
    col_data_excel: Optional[str],
    col_motorista_excel: Optional[str],
    status_entregue: List[str],
    status_ocorrencia: List[str],
    col_motivo_indicador: Optional[str] = None,
) -> pd.DataFrame:
    """
    Prepara base completa do Excel para cálculo dos bônus e das métricas do motorista.

    Agora esta base mantém TODOS os pedidos destinados ao motorista no Excel.
    Ela não filtra somente entregas pagas, pois o medidor precisa calcular:

    - Total de pedidos destinados ao motorista;
    - Entregas realizadas;
    - Entregas pendentes/ocorrência/não realizadas;
    - Percentual de sucesso real.

    A coluna "Entrega Paga" continua existindo para os bônus e para as métricas.
    """
    if df_excel_raw.empty or not col_data_excel or not col_motorista_excel:
        return pd.DataFrame()
    if col_data_excel not in df_excel_raw.columns or col_motorista_excel not in df_excel_raw.columns:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["Pedido"] = df_excel_raw[col_pedido].astype(str).str.strip() if col_pedido in df_excel_raw.columns else ""
    out["Status"] = df_excel_raw[col_status].astype(str).str.strip() if col_status in df_excel_raw.columns else ""

    # Status real do indicador: Motivo 1.
    # Se a coluna Motivo 1 existir, ela manda no cálculo do medidor.
    # Motivo 1 vazio = realizada; Motivo 1 preenchido = ocorrência/não realizada.
    if col_motivo_indicador and col_motivo_indicador in df_excel_raw.columns:
        out["Motivo 1"] = df_excel_raw[col_motivo_indicador].fillna("").astype(str).str.strip()
    elif "Motivo 1" in df_excel_raw.columns:
        out["Motivo 1"] = df_excel_raw["Motivo 1"].fillna("").astype(str).str.strip()
    else:
        out["Motivo 1"] = ""

    out["Status Indicador"] = out["Motivo 1"]
    out["Status Normalizado"] = out["Status Indicador"].apply(normalizar_texto)
    out["Data Entrega"] = df_excel_raw[col_data_excel].apply(to_date).dt.date
    out["Motorista Final"] = df_excel_raw[col_motorista_excel].astype(str).str.strip().str.upper()

    # O Motivo 1 pode trazer tanto confirmação de entrega
    # (ex.: Portaria, Algum parente, A própria pessoa)
    # quanto ocorrência real. Por isso NÃO podemos considerar
    # todo Motivo 1 preenchido como falha.
    # Regra correta: ocorrência somente quando o texto do Motivo 1
    # bater com a lista de ocorrências/não pagar configurada no sidebar.
    ocorrencia_norm = [normalizar_texto(x) for x in status_ocorrencia if str(x).strip()]
    out["É Ocorrência"] = out["Status Normalizado"].apply(
        lambda x: any(s in x for s in ocorrencia_norm)
    ) if ocorrencia_norm else False
    out["É Entregue"] = ~out["É Ocorrência"]
    out["Entrega Paga"] = out["É Entregue"] & (~out["É Ocorrência"])

    # Mantém todos os pedidos válidos destinados ao motorista.
    # Remove apenas registros sem pedido, sem data ou sem motorista.
    out = out[
        out["Data Entrega"].notna()
        & (out["Motorista Final"].astype(str).str.strip() != "")
        & (out["Pedido"].astype(str).str.strip() != "")
        & (out["Pedido"].astype(str).str.lower() != "nan")
    ].copy()

    # Evita contar a mesma entrega mais de uma vez se o Excel tiver linhas repetidas.
    # Mantém o último status encontrado para aquele pedido/data/motorista.
    out = out.drop_duplicates(subset=["Motorista Final", "Data Entrega", "Pedido"], keep="last")
    return out.reset_index(drop=True)


def listar_sabados_do_mes(ano: int, mes: int) -> List[object]:
    inicio = pd.Timestamp(year=int(ano), month=int(mes), day=1)
    fim = inicio + pd.offsets.MonthEnd(0)
    return [d.date() for d in pd.date_range(inicio, fim, freq="W-SAT")]


def calcular_bonus_sabados_excel(
    df_bonus_excel: pd.DataFrame,
    motorista: str,
    ano: int,
    mes: int,
    valor_por_entrega: float = 2.0,
) -> Tuple[float, int, bool, List[object], List[object]]:
    """Calcula bônus de sábado: só libera se trabalhou todos os sábados do mês."""
    if df_bonus_excel.empty:
        return 0.0, 0, False, [], []

    todos_sabados = listar_sabados_do_mes(ano, mes)
    if not todos_sabados:
        return 0.0, 0, False, [], []

    df = df_bonus_excel.copy()
    df["Data Entrega DT"] = pd.to_datetime(df["Data Entrega"], errors="coerce").dt.date
    df = df[
        (df["Motorista Final"].astype(str).str.upper().str.strip() == str(motorista).upper().strip())
        & (pd.to_datetime(df["Data Entrega"], errors="coerce").dt.month == int(mes))
        & (pd.to_datetime(df["Data Entrega"], errors="coerce").dt.year == int(ano))
    ].copy()

    if df.empty:
        return 0.0, 0, False, todos_sabados, []

    if "Entrega Paga" in df.columns:
        df = df[df["Entrega Paga"].fillna(False).astype(bool)].copy()

    if df.empty:
        return 0.0, 0, False, todos_sabados, []

    dias_trabalhados = sorted(set(df["Data Entrega DT"].dropna()))
    sabados_trabalhados = sorted([d for d in dias_trabalhados if d in set(todos_sabados)])
    veio_todos_sabados = set(todos_sabados).issubset(set(sabados_trabalhados))

    if not veio_todos_sabados:
        return 0.0, 0, False, todos_sabados, sabados_trabalhados

    qtd_entregas_sabado = int(df[df["Data Entrega DT"].isin(todos_sabados)].shape[0])
    bonus = qtd_entregas_sabado * max(0.0, to_float(valor_por_entrega))
    return float(bonus), qtd_entregas_sabado, True, todos_sabados, sabados_trabalhados


def detalhar_bonus_sabados_excel(
    df_bonus_excel: pd.DataFrame,
    motorista: str,
    ano: int,
    mes: int,
    valor_por_entrega: float = 2.0,
) -> Dict[str, object]:
    """Retorna os detalhes para conferência visual do bônus de sábado.

    Regra:
    - identifica todos os sábados do mês;
    - busca no Excel em quais sábados o motorista teve entrega paga;
    - conta as entregas feitas nos sábados;
    - libera o valor apenas se o motorista teve entrega em todos os sábados do mês.
    """
    todos_sabados = listar_sabados_do_mes(ano, mes)
    valor_unitario = max(0.0, to_float(valor_por_entrega))

    if df_bonus_excel.empty or not todos_sabados:
        return {
            "todos_sabados": todos_sabados,
            "sabados_trabalhados": [],
            "sabados_faltantes": todos_sabados,
            "qtd_entregas_sabado": 0,
            "veio_todos_sabados": False,
            "bonus_calculado": 0.0,
        }

    df = df_bonus_excel.copy()
    df["Data Entrega DT"] = pd.to_datetime(df["Data Entrega"], errors="coerce").dt.date
    data_entrega_convertida = pd.to_datetime(df["Data Entrega"], errors="coerce")

    df = df[
        (df["Motorista Final"].astype(str).str.upper().str.strip() == str(motorista).upper().strip())
        & (data_entrega_convertida.dt.month == int(mes))
        & (data_entrega_convertida.dt.year == int(ano))
    ].copy()

    if df.empty:
        return {
            "todos_sabados": todos_sabados,
            "sabados_trabalhados": [],
            "sabados_faltantes": todos_sabados,
            "qtd_entregas_sabado": 0,
            "veio_todos_sabados": False,
            "bonus_calculado": 0.0,
        }

    if "Entrega Paga" in df.columns:
        df = df[df["Entrega Paga"].fillna(False).astype(bool)].copy()

    if df.empty:
        return {
            "todos_sabados": todos_sabados,
            "sabados_trabalhados": [],
            "sabados_faltantes": todos_sabados,
            "qtd_entregas_sabado": 0,
            "veio_todos_sabados": False,
            "bonus_calculado": 0.0,
        }

    sabados_set = set(todos_sabados)
    df_sabados = df[df["Data Entrega DT"].isin(sabados_set)].copy()
    sabados_trabalhados = sorted(set(df_sabados["Data Entrega DT"].dropna()))
    sabados_faltantes = sorted([d for d in todos_sabados if d not in set(sabados_trabalhados)])
    veio_todos_sabados = len(sabados_faltantes) == 0 and len(todos_sabados) > 0
    qtd_entregas_sabado = int(len(df_sabados))
    bonus_calculado = qtd_entregas_sabado * valor_unitario if veio_todos_sabados else 0.0

    return {
        "todos_sabados": todos_sabados,
        "sabados_trabalhados": sabados_trabalhados,
        "sabados_faltantes": sabados_faltantes,
        "qtd_entregas_sabado": qtd_entregas_sabado,
        "veio_todos_sabados": veio_todos_sabados,
        "bonus_calculado": float(bonus_calculado),
    }


def parse_datas_feriados(texto_datas: str, ano_padrao: int) -> List[object]:
    """Aceita feriados digitados um por linha ou separados por vírgula/; no formato dd/mm ou dd/mm/aaaa."""
    datas = []
    partes = re.split(r"[\n,;]+", str(texto_datas or ""))
    for parte in partes:
        item = parte.strip()
        if not item:
            continue
        if re.fullmatch(r"\d{1,2}/\d{1,2}", item):
            item = f"{item}/{ano_padrao}"
        data = pd.to_datetime(item, dayfirst=True, errors="coerce")
        if pd.notna(data):
            datas.append(data.date())
    return sorted(set(datas))


def calcular_bonus_feriados_excel(
    df_bonus_excel: pd.DataFrame,
    motorista: str,
    datas_feriado: List[object],
    valor_por_entrega: float = 2.0,
) -> Tuple[float, int]:
    """Calcula bônus de feriado pelas entregas do Excel nas datas informadas."""
    if df_bonus_excel.empty or not datas_feriado:
        return 0.0, 0

    df = df_bonus_excel.copy()
    df["Data Entrega DT"] = pd.to_datetime(df["Data Entrega"], errors="coerce").dt.date
    df = df[
        (df["Motorista Final"].astype(str).str.upper().str.strip() == str(motorista).upper().strip())
        & (df["Data Entrega DT"].isin(set(datas_feriado)))
    ].copy()

    if "Entrega Paga" in df.columns:
        df = df[df["Entrega Paga"].fillna(False).astype(bool)].copy()

    qtd_entregas_feriado = int(len(df))
    bonus = qtd_entregas_feriado * max(0.0, to_float(valor_por_entrega))
    return float(bonus), qtd_entregas_feriado



def preparar_base_metricas_motorista_excel(
    df_excel_raw: pd.DataFrame,
    col_pedido: str,
    col_status: str,
    col_data_excel: Optional[str],
    col_motorista_excel: Optional[str],
    status_entregue: List[str],
    status_ocorrencia: List[str],
    col_motivo_indicador: Optional[str] = None,
) -> pd.DataFrame:
    """
    Prepara a base completa do Excel para o medidor de métricas do motorista.

    Regra do indicador:
    - Pedido/AWB, Data e Motorista vêm do Excel;
    - O status do indicador deve vir da coluna Motivo 1;
    - Motivo 1 vazio = entrega realizada;
    - Motivo 1 preenchido = ocorrência / não realizada.

    A coluna Status continua existindo apenas para conferência e para compatibilidade,
    mas não é usada no cálculo do indicador quando Motivo 1 estiver selecionada.
    """
    if df_excel_raw.empty or not col_data_excel or not col_motorista_excel:
        return pd.DataFrame()
    if col_data_excel not in df_excel_raw.columns or col_motorista_excel not in df_excel_raw.columns:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["Pedido"] = df_excel_raw[col_pedido].astype(str).str.strip() if col_pedido in df_excel_raw.columns else ""
    out["Status"] = df_excel_raw[col_status].astype(str).str.strip() if col_status in df_excel_raw.columns else ""

    # Status real do indicador: Motivo 1.
    # Se a coluna Motivo 1 existir, ela manda no cálculo do medidor.
    # Motivo 1 vazio = realizada; Motivo 1 preenchido = ocorrência/não realizada.
    if col_motivo_indicador and col_motivo_indicador in df_excel_raw.columns:
        out["Motivo 1"] = df_excel_raw[col_motivo_indicador].fillna("").astype(str).str.strip()
    elif "Motivo 1" in df_excel_raw.columns:
        out["Motivo 1"] = df_excel_raw["Motivo 1"].fillna("").astype(str).str.strip()
    else:
        out["Motivo 1"] = ""

    out["Status Indicador"] = out["Motivo 1"]
    out["Status Normalizado"] = out["Status Indicador"].apply(normalizar_texto)
    out["Data Entrega"] = df_excel_raw[col_data_excel].apply(to_date).dt.date
    out["Motorista Final"] = df_excel_raw[col_motorista_excel].astype(str).str.strip().str.upper()

    # O Motivo 1 pode trazer tanto confirmação de entrega
    # (ex.: Portaria, Algum parente, A própria pessoa)
    # quanto ocorrência real. Por isso NÃO podemos considerar
    # todo Motivo 1 preenchido como falha.
    # Regra correta: ocorrência somente quando o texto do Motivo 1
    # bater com a lista de ocorrências/não pagar configurada no sidebar.
    ocorrencia_norm = [normalizar_texto(x) for x in status_ocorrencia if str(x).strip()]
    out["É Ocorrência"] = out["Status Normalizado"].apply(
        lambda x: any(s in x for s in ocorrencia_norm)
    ) if ocorrencia_norm else False
    out["É Entregue"] = ~out["É Ocorrência"]
    out["Entrega Paga"] = out["É Entregue"] & (~out["É Ocorrência"])

    out = out[
        out["Data Entrega"].notna()
        & (out["Motorista Final"].astype(str).str.strip() != "")
        & (out["Pedido"].astype(str).str.strip() != "")
        & (out["Pedido"].astype(str).str.lower() != "nan")
    ].copy()

    out = out.drop_duplicates(subset=["Motorista Final", "Data Entrega", "Pedido"], keep="last")
    return out.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def criar_excel_fechamento(
    df_dia,
    df_entregas,
    df_pdf_info,
    df_relatorio_entregas: Optional[pd.DataFrame] = None,
    motorista_relatorio: str = "",
    data_inicio_relatorio=None,
    data_fim_relatorio=None,
    acareacao_relatorio: float = 0.0,
    vale_relatorio: float = 0.0,
    desconto_relatorio: float = 0.0,
    bonus_extra_relatorio: float = 0.0,
    bonus_sabados_relatorio: float = 0.0,
    bonus_feriado_relatorio: float = 0.0,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_dia_export = df_dia.copy()
        colunas_export = [
            "Motorista Final", "Data Rota", "Rotas", "Quantidade_Entregas",
            "Valor_Entregas", "KG_Excedente_Calculado",
            "Valor_KG_Excedente", "Total_Dia",
        ]
        colunas_export = [c for c in colunas_export if c in df_dia_export.columns]
        df_dia_export = df_dia_export[colunas_export]

        colunas_adicionais = [
            "Subtotal", "Acareação", "Bônus Extr", "Bônus Sáb",
            "Bônus Feri", "Vale", "Desconto", "Total",
        ]
        for col in colunas_adicionais:
            df_dia_export[col] = 0.0
            df_dia_export[col] = pd.to_numeric(df_dia_export[col], errors="coerce").fillna(0.0)

        if not df_dia_export.empty:
            if "Total_Dia" in df_dia_export.columns:
                df_dia_export["Subtotal"] = pd.to_numeric(
                    df_dia_export["Total_Dia"],
                    errors="coerce",
                ).fillna(0.0)

            df_periodo = df_dia_export.copy()
            if motorista_relatorio and "Motorista Final" in df_periodo.columns:
                df_periodo = df_periodo[
                    df_periodo["Motorista Final"].astype(str).str.upper().str.strip()
                    == str(motorista_relatorio).upper().strip()
                ].copy()

            if not df_periodo.empty and "Data Rota" in df_periodo.columns and data_inicio_relatorio is not None and data_fim_relatorio is not None:
                datas_periodo_tmp = pd.to_datetime(df_periodo["Data Rota"], errors="coerce").dt.date
                inicio_tmp = pd.to_datetime(data_inicio_relatorio, errors="coerce").date()
                fim_tmp = pd.to_datetime(data_fim_relatorio, errors="coerce").date()
                df_periodo = df_periodo[(datas_periodo_tmp >= inicio_tmp) & (datas_periodo_tmp <= fim_tmp)].copy()

            if df_periodo.empty:
                df_periodo = df_dia_export.copy()

            idx_primeiro = df_periodo.index[0]
            idx_ultimo = df_periodo.index[-1]

            df_dia_export.loc[idx_primeiro, "Acareação"] = to_float(acareacao_relatorio)
            df_dia_export.loc[idx_primeiro, "Bônus Extr"] = to_float(bonus_extra_relatorio)
            df_dia_export.loc[idx_primeiro, "Vale"] = to_float(vale_relatorio)
            df_dia_export.loc[idx_primeiro, "Desconto"] = to_float(desconto_relatorio)

            bonus_sabados_total = to_float(bonus_sabados_relatorio)
            if bonus_sabados_total > 0 and "Data Rota" in df_periodo.columns:
                datas_periodo = pd.to_datetime(df_periodo["Data Rota"], errors="coerce")
                idx_sabados = df_periodo.index[datas_periodo.dt.weekday == 5].tolist()
                if idx_sabados:
                    if df_relatorio_entregas is not None and not df_relatorio_entregas.empty and "Data Rota" in df_relatorio_entregas.columns:
                        df_bonus_sab = df_relatorio_entregas.copy()
                        df_bonus_sab["Data Rota DT"] = pd.to_datetime(df_bonus_sab["Data Rota"], errors="coerce").dt.date
                        df_bonus_sab = df_bonus_sab[pd.to_datetime(df_bonus_sab["Data Rota"], errors="coerce").dt.weekday == 5].copy()
                        qtd_sabados_total = int(len(df_bonus_sab))
                        valor_unitario_sabado = bonus_sabados_total / qtd_sabados_total if qtd_sabados_total > 0 else 0.0
                        for idx_sab in idx_sabados:
                            data_sab = pd.to_datetime(df_dia_export.loc[idx_sab, "Data Rota"], errors="coerce").date()
                            qtd_data = int((df_bonus_sab["Data Rota DT"] == data_sab).sum())
                            df_dia_export.loc[idx_sab, "Bônus Sáb"] = qtd_data * valor_unitario_sabado
                    else:
                        df_dia_export.loc[idx_sabados[0], "Bônus Sáb"] = bonus_sabados_total
                else:
                    df_dia_export.loc[idx_primeiro, "Bônus Sáb"] = bonus_sabados_total

            bonus_feriado_total = to_float(bonus_feriado_relatorio)
            if bonus_feriado_total > 0:
                df_dia_export.loc[idx_ultimo, "Bônus Feri"] = bonus_feriado_total

            df_dia_export["Total"] = (
                pd.to_numeric(df_dia_export["Subtotal"], errors="coerce").fillna(0.0)
                + pd.to_numeric(df_dia_export["Acareação"], errors="coerce").fillna(0.0)
                + pd.to_numeric(df_dia_export["Bônus Extr"], errors="coerce").fillna(0.0)
                + pd.to_numeric(df_dia_export["Bônus Sáb"], errors="coerce").fillna(0.0)
                + pd.to_numeric(df_dia_export["Bônus Feri"], errors="coerce").fillna(0.0)
                - pd.to_numeric(df_dia_export["Vale"], errors="coerce").fillna(0.0)
                - pd.to_numeric(df_dia_export["Desconto"], errors="coerce").fillna(0.0)
            )

        df_dia_export.to_excel(writer, index=False, sheet_name="Fechamento diario")
        df_entregas.to_excel(writer, index=False, sheet_name="Entregas pagas")
        df_pdf_info.drop(columns=["Texto PDF"], errors="ignore").to_excel(writer, index=False, sheet_name="PDFs lidos")

        workbook = writer.book
        money_fmt = workbook.add_format({"num_format": 'R$ #,##0.00'})
        header_fmt = workbook.add_format({
            "bold": True,
            "bg_color": "#0B2B69",
            "font_color": "white",
            "border": 1,
            "align": "center",
        })

        for sheet_name, df in {
            "Fechamento diario": df_dia_export,
            "Entregas pagas": df_entregas,
            "PDFs lidos": df_pdf_info.drop(columns=["Texto PDF"], errors="ignore"),
        }.items():
            ws = writer.sheets[sheet_name]
            for idx, col in enumerate(df.columns):
                ws.write(0, idx, col, header_fmt)
                width = max(14, min(45, max([len(str(col))] + [len(str(v)) for v in df[col].head(200).fillna("").astype(str)])))
                ws.set_column(idx, idx, width)
                if any(x in col.lower() for x in ["valor", "total", "subtotal", "acareação", "bonus", "bônus", "vale", "desconto"]):
                    ws.set_column(idx, idx, 16, money_fmt)

    return output.getvalue()



LOGO_GDS_BASE64 = """/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCAFJA0kDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9UKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoopM0ALRRmigAoopM0ALRSZ9jQGBGf50riTuLRSbge9MaeNM7nAx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Whhko0YKK8lY+SrV62IlzVZOT7t3EC/LjtTh0oorq2MOtwooopjCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKMUUUAGKMUUUAGKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKCQKM0AFFGRnGefSigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoopM9ev5UALRSZGcUtABRRRQAUUUUAFFFFABRRSFsf5zQAtFN3gfSmz3EVtE0s0ixRr95nYKB9SaAJKK52b4jeFLd9kvibR43/utfxA/+hVe07xTo2sHFhq1je/9e1zHJ/6CTQBqUUgYGgsACecDuRQAtFFFABRRRQAUUUUAFFITjrRuGf8ACgBaKzL/AMT6RpUhS91SztHHUT3CIR+ZqrD478N3LhItf0uVzwFS9iYn8A1AG7RTIp454w8brIh/iQ5H6U4tigBaKOtFABRRRQAUhYD+VLXH/FfxF4k8KeBNV1Hwh4Zfxf4kRAljpIuEt0kkY4BeRiAqLnc3fA4GaAOE/aY/a6+H/wCyp4etNQ8YX081/fMVstG01VlvLkD7zqhYAIvd2IXPGckCvm4f8Fo/guTgeF/HBJ6EWVpz/wCTVfJPxD/YC/ax/aC+Jd74o8a6XZvqmpygy3l7qsPkW0efljRFZisaDgKoOPqST93fsn/8E1Ph3+zqbLXNahTxt45iCyHU7+L/AEa0k/6d4TwMHo7ZbjI25wAD6E+C/wAV5/jH4Rj8SHwd4g8H6fdYazh8SRRQ3NzGeknlJI7Ip7b9pIwQCOa9AByAabtJHU8jvTqACiiigAoopNw3AZ5oAWiqeoaxY6TEZL28gs4x1e4lVAPzNYL/ABW8FRvsfxdoSv02nU4M/wDoVAHVUVnaZ4j0rW0D6fqdnfoejW1wkg/NSa0MigBaKTcP8il60AFFFFABRRmk3D60ALRSbgPXP0oyP8igBaKTcPf8jRuHuPwoAWik3D1pN4z0J7ZxQA6ijrRQ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"""



def gerar_relatorio_entregas_pdf(
    df_recibo: pd.DataFrame,
    motorista: str,
    data_inicio,
    data_fim,
    quinzena: str = "",
    acareacao: float = 0.0,
    vale: float = 0.0,
    desconto: float = 0.0,
    bonus_extra: float = 0.0,
    bonus_sabados: float = 0.0,
    bonus_feriado: float = 0.0,
    cnpj_motorista: str = "SEM CNPJ",
) -> bytes:
    """Gera relatório de entregas em PDF resumido por data, CEP, valor, quantidade e total, mantendo totais e demais informações."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        KeepTogether,
        Image,
    )

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=0.65 * cm,
        leftMargin=0.65 * cm,
        topMargin=0.45 * cm,
        bottomMargin=0.55 * cm,
    )

    styles = getSampleStyleSheet()
    preto = colors.black
    laranja_claro = colors.HexColor("#FBE4D5")
    laranja_linha = colors.HexColor("#F4B183")
    branco = colors.white

    style_title = ParagraphStyle(
        "TituloRelatorioGDS",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=12,
        textColor=preto,
        alignment=TA_CENTER,
        spaceAfter=4,
    )
    style_header = ParagraphStyle(
        "HeaderRelatorioGDS",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.2,
        textColor=preto,
        alignment=TA_CENTER,
    )
    style_normal = ParagraphStyle(
        "NormalRelatorioGDS",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.5,
        textColor=preto,
        leading=7.4,
        alignment=TA_CENTER,
    )
    style_small_center = ParagraphStyle(
        "SmallCenterRelatorioGDS",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.2,
        alignment=TA_CENTER,
        textColor=preto,
    )

    def fmt_data(v):
        try:
            return pd.to_datetime(v).strftime("%d/%m/%Y")
        except Exception:
            return str(v)

    def fmt_num(v):
        try:
            valor = float(v)
            if abs(valor - int(valor)) < 0.0001:
                return str(int(valor))
            return f"{valor:.2f}".replace(".", ",")
        except Exception:
            return "0"

    def p(txt, style=style_normal):
        return Paragraph(str(txt if txt is not None else ""), style)

    df = df_recibo.copy()
    if df.empty:
        df = pd.DataFrame(columns=["Data Rota", "Pedido", "CEP Prefixo", "Valor CEP", "KG Excedente", "Valor Excedente KG", "Total Entrega"])

    if "Data Rota" in df.columns:
        df["Data Rota"] = pd.to_datetime(df["Data Rota"], errors="coerce")
    else:
        df["Data Rota"] = pd.NaT

    if "CEP Prefixo" not in df.columns:
        if "CEP" in df.columns:
            df["CEP Prefixo"] = (
                df["CEP"]
                .astype(str)
                .str.replace(r"\D", "", regex=True)
                .str.zfill(8)
                .str[:3]
            )
        else:
            df["CEP Prefixo"] = ""

    for col in ["Valor CEP", "KG Excedente", "Valor Excedente KG", "Total Entrega"]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = df[col].fillna(0).astype(float)

    if "Pedido" not in df.columns:
        df["Pedido"] = ""
    if "Rota" not in df.columns:
        df["Rota"] = ""
    if "Status_Encontrados" not in df.columns:
        df["Status_Encontrados"] = ""

    df = df.sort_values(["Data Rota", "CEP Prefixo", "Pedido"], kind="mergesort")

    total_entregas = int(len(df))
    total_valor_entregas = float(df["Valor CEP"].sum())
    total_kg = float(df["KG Excedente"].sum())
    total_valor_kg = float(df["Valor Excedente KG"].sum())
    subtotal_base = float(df["Total Entrega"].sum())

    acareacao = max(0.0, to_float(acareacao))
    vale = max(0.0, to_float(vale))
    desconto = max(0.0, to_float(desconto))
    bonus_extra = max(0.0, to_float(bonus_extra))
    bonus_sabados = max(0.0, to_float(bonus_sabados))
    bonus_feriado = max(0.0, to_float(bonus_feriado))
    total_geral = subtotal_base + acareacao + bonus_extra + bonus_sabados + bonus_feriado - vale - desconto

    periodo_txt = f"{fmt_data(data_inicio)} a {fmt_data(data_fim)}"
    quinzena_txt = limpar_texto(quinzena) or "Não informado"
    cnpj_txt = limpar_texto(cnpj_motorista) or obter_cnpj_motorista(motorista)

    elementos = []

    try:
        logo_bytes = base64.b64decode(LOGO_GDS_BASE64)
        logo = Image(io.BytesIO(logo_bytes), width=3.7 * cm, height=1.05 * cm)
        logo.hAlign = "CENTER"
        elementos.append(logo)
        elementos.append(Spacer(1, 0.06 * cm))
    except Exception:
        pass

    elementos.append(Paragraph("Relatorio de entregas realizadas", style_title))

    info_table = Table(
        [
            [p("Nome do Entregador", style_header), p("CNPJ", style_header), p("Quinzena", style_header)],
            [p(str(motorista).upper()), p(cnpj_txt), p(quinzena_txt)],
            [p("Período", style_header), p("Total de Entregas", style_header), p("Total do Relatório", style_header)],
            [p(periodo_txt), p(str(total_entregas)), p(moeda(total_geral))],
        ],
        colWidths=[9.2 * cm, 5.0 * cm, 5.0 * cm],
    )
    info_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.65, laranja_linha),
        ("BACKGROUND", (0, 0), (-1, 0), laranja_claro),
        ("BACKGROUND", (0, 2), (-1, 2), laranja_claro),
        ("BACKGROUND", (0, 1), (-1, 1), branco),
        ("BACKGROUND", (0, 3), (-1, 3), branco),
        ("TEXTCOLOR", (0, 0), (-1, -1), preto),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 2), (-1, 2), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.8),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    elementos.append(info_table)
    elementos.append(Spacer(1, 0.16 * cm))

    # RELATÓRIO SIMPLIFICADO:
    # Agrupa por DATA + CEP + VALOR.
    # Exemplo:
    # 16/02/2026 | CEP 010 | Entregas realizadas | R$ 6,00 | 10 | R$ 60,00
    linhas = [["DATA", "DESCRIÇÃO", "CEP", "VALOR", "QUANTIDADE", "TOTAL"]]

    if not df.empty:
        df_resumo = df.copy()
        df_resumo["Data Agrupamento"] = pd.to_datetime(df_resumo["Data Rota"], errors="coerce").dt.date
        df_resumo["CEP Prefixo"] = df_resumo["CEP Prefixo"].astype(str).apply(normalizar_prefixo_cep)
        df_resumo["Valor CEP"] = df_resumo["Valor CEP"].fillna(0).astype(float)
        df_resumo["Total Entrega"] = df_resumo["Total Entrega"].fillna(0).astype(float)
        df_resumo["Valor Excedente KG"] = df_resumo["Valor Excedente KG"].fillna(0).astype(float)
        df_resumo["KG Excedente"] = df_resumo["KG Excedente"].fillna(0).astype(float)
        if "Descrição Relatório" not in df_resumo.columns:
            df_resumo["Descrição Relatório"] = "Entregas realizadas"
        df_resumo["Descrição Relatório"] = df_resumo["Descrição Relatório"].fillna("Entregas realizadas").astype(str)

        resumo_data_cep = (
            df_resumo.groupby(["Data Agrupamento", "Descrição Relatório", "CEP Prefixo", "Valor CEP"], dropna=False)
            .agg(
                Quantidade=("Pedido", "size"),
                Total=("Valor CEP", "sum"),
                KG_Excedente=("KG Excedente", "sum"),
                Total_KG=("Valor Excedente KG", "sum"),
                Total_Geral=("Total Entrega", "sum"),
            )
            .reset_index()
            .sort_values(["Data Agrupamento", "Descrição Relatório", "CEP Prefixo", "Valor CEP"], kind="mergesort")
        )

        for _, row in resumo_data_cep.iterrows():
            qtd = int(row.get("Quantidade", 0) or 0)
            descricao_linha = str(row.get("Descrição Relatório", "Entregas realizadas") or "Entregas realizadas")
            cep_prefixo = normalizar_prefixo_cep(row.get("CEP Prefixo", "")) if descricao_linha != "RD Fechada" else ""
            valor_cep = float(row.get("Valor CEP", 0) or 0)
            total_entregas_cep = float(row.get("Total", 0) or 0)
            kg_excedente = float(row.get("KG_Excedente", 0) or 0)
            total_kg = float(row.get("Total_KG", 0) or 0)
            total_linha = float(row.get("Total_Geral", 0) or 0)

            linhas.append([
                fmt_data(row.get("Data Agrupamento")),
                descricao_linha,
                cep_prefixo,
                moeda(valor_cep),
                str(qtd),
                moeda(total_entregas_cep),
            ])

            if kg_excedente > 0:
                linhas.append([
                    fmt_data(row.get("Data Agrupamento")),
                    "Kg excedente acima de 10kg",
                    cep_prefixo,
                    "R$ 0,30",
                    fmt_num(kg_excedente),
                    moeda(total_kg),
                ])

                # Linha de total do agrupamento quando houver kg excedente,
                # para deixar claro o total de entrega + adicional de kg.
                linhas.append([
                    fmt_data(row.get("Data Agrupamento")),
                    "Total com kg excedente",
                    cep_prefixo,
                    "",
                    "",
                    moeda(total_linha),
                ])

    linhas.append(["", "TOTAL DAS ENTREGAS", "", "", str(total_entregas), moeda(subtotal_base)])

    tabela = Table(
        linhas,
        colWidths=[2.4 * cm, 6.3 * cm, 2.0 * cm, 2.8 * cm, 2.6 * cm, 3.1 * cm],
        repeatRows=1,
    )
    tabela_style = TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.45, laranja_linha),
        ("BACKGROUND", (0, 0), (-1, 0), laranja_claro),
        ("TEXTCOLOR", (0, 0), (-1, -1), preto),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (-1, -2), branco),
        ("BACKGROUND", (0, -1), (-1, -1), laranja_claro),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.2),
        ("LEADING", (0, 0), (-1, -1), 7.0),
        ("TOPPADDING", (0, 0), (-1, -1), 1.4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.4),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.5),
    ])

    # Destaca linhas de KG e total com kg excedente
    for idx_linha, linha in enumerate(linhas):
        if idx_linha == 0:
            continue
        descricao = str(linha[1]).lower()
        if "kg excedente" in descricao or "total com kg" in descricao:
            tabela_style.add("BACKGROUND", (0, idx_linha), (-1, idx_linha), colors.HexColor("#F2F2F2"))
            tabela_style.add("FONTNAME", (0, idx_linha), (-1, idx_linha), "Helvetica")

    tabela.setStyle(tabela_style)
    elementos.append(tabela)
    elementos.append(Spacer(1, 0.14 * cm))

    linhas_totais_finais = [
        ["Total Entregas", moeda(total_valor_entregas)],
        ["Total Kg Excedente", moeda(total_valor_kg)],
        ["Subtotal", moeda(subtotal_base)],
        ["Acareação", moeda(acareacao)],
        ["Bônus Extra", moeda(bonus_extra)],
        ["Bônus Sábados", moeda(bonus_sabados)],
        ["Bônus Feriado", moeda(bonus_feriado)],
        ["Vale", f"- {moeda(vale)}" if vale > 0 else moeda(0)],
        ["Desconto", f"- {moeda(desconto)}" if desconto > 0 else moeda(0)],
        ["TOTAL DO RELATÓRIO", moeda(total_geral)],
    ]

    totais_finais = Table(linhas_totais_finais, colWidths=[4.8 * cm, 4.0 * cm], hAlign="CENTER")
    totais_finais.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.55, laranja_linha),
        ("BACKGROUND", (0, 0), (-1, -1), branco),
        ("BACKGROUND", (0, -1), (-1, -1), laranja_claro),
        ("TEXTCOLOR", (0, 0), (-1, -1), preto),
        ("FONTNAME", (0, 0), (-1, -2), "Helvetica"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.8),
        ("TOPPADDING", (0, 0), (-1, -1), 1.6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elementos.append(KeepTogether(totais_finais))
    elementos.append(Spacer(1, 0.95 * cm))

    assinatura_motorista = Paragraph(
        "<para align='center'>_________________________________________<br/>"
        "<b>Assinatura do Motorista</b></para>",
        style_normal,
    )
    elementos.append(assinatura_motorista)
    elementos.append(Spacer(1, 0.10 * cm))

    elementos.append(Paragraph(
        f"Relatório gerado automaticamente em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        style_small_center,
    ))

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()


def gerar_recibo_pdf(
    df_recibo: pd.DataFrame,
    motorista: str,
    data_inicio,
    data_fim,
    quinzena: str = "",
    acareacao: float = 0.0,
    vale: float = 0.0,
    desconto: float = 0.0,
    bonus_extra: float = 0.0,
    bonus_sabados: float = 0.0,
    bonus_feriado: float = 0.0,
    cnpj_motorista: str = "SEM CNPJ",
) -> bytes:
    """Gera recibo em PDF com logo GDS, cores claras e fechamento separado por dia e por valor real do CEP."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        KeepTogether,
        Image,
    )

    buffer = io.BytesIO()

    # Recibo em A4 RETRATO.
    # Margens reduzidas e fontes compactas para tentar manter tudo em uma folha
    # sempre que a quantidade de linhas permitir.
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=0.65 * cm,
        leftMargin=0.65 * cm,
        topMargin=0.45 * cm,
        bottomMargin=0.55 * cm,
    )

    styles = getSampleStyleSheet()
    preto = colors.black
    laranja_claro = colors.HexColor("#FBE4D5")
    laranja_linha = colors.HexColor("#F4B183")
    branco = colors.white

    style_title = ParagraphStyle(
        "TituloReciboGDS",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=12,
        textColor=preto,
        alignment=TA_CENTER,
        spaceAfter=4,
    )
    style_header = ParagraphStyle(
        "HeaderReciboGDS",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.2,
        textColor=preto,
        alignment=TA_CENTER,
    )
    style_normal = ParagraphStyle(
        "NormalReciboGDS",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.8,
        textColor=preto,
        leading=8.2,
        alignment=TA_CENTER,
    )
    style_small_right = ParagraphStyle(
        "SmallRightGDS",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.2,
        alignment=TA_CENTER,
        textColor=preto,
    )

    def fmt_data(v):
        try:
            return pd.to_datetime(v).strftime("%d/%m/%Y")
        except Exception:
            return str(v)

    def fmt_num(v):
        try:
            valor = float(v)
            if abs(valor - int(valor)) < 0.0001:
                return str(int(valor))
            return f"{valor:.2f}".replace(".", ",")
        except Exception:
            return "0"

    df = df_recibo.copy()
    df["Data Rota"] = pd.to_datetime(df["Data Rota"], errors="coerce")
    df = df.sort_values("Data Rota")

    # Garante que o recibo/relatório tenha o prefixo do CEP usado no cálculo.
    # Exemplo: CEP 01045-000 => prefixo 010.
    if "CEP Prefixo" not in df.columns:
        if "CEP" in df.columns:
            df["CEP Prefixo"] = (
                df["CEP"]
                .astype(str)
                .str.replace(r"\D", "", regex=True)
                .str.zfill(8)
                .str[:3]
            )
        else:
            df["CEP Prefixo"] = ""

    # O recibo correto precisa usar a base DETALHADA de entregas, não a média do dia.
    # Assim, quando no mesmo dia houver CEP de R$ 7,00 e CEP de R$ 8,50, o recibo mostra linhas separadas.
    usar_base_detalhada = "Valor CEP" in df.columns and "Pedido" in df.columns

    if usar_base_detalhada:
        df["Valor CEP"] = df["Valor CEP"].fillna(0).astype(float)
        df["KG Excedente"] = df.get("KG Excedente", 0).fillna(0).astype(float) if hasattr(df.get("KG Excedente", 0), "fillna") else 0.0
        df["Valor Excedente KG"] = df.get("Valor Excedente KG", 0).fillna(0).astype(float) if hasattr(df.get("Valor Excedente KG", 0), "fillna") else 0.0
        df["Total Entrega"] = df.get("Total Entrega", 0).fillna(0).astype(float) if hasattr(df.get("Total Entrega", 0), "fillna") else 0.0

        total_entregas = int(len(df))
        total_valor_entregas = float(df["Valor CEP"].sum())
        total_kg = float(df["KG Excedente"].sum())
        total_valor_kg = float(df["Valor Excedente KG"].sum())
        total_geral = float(df["Total Entrega"].sum())
    else:
        total_entregas = int(df["Quantidade_Entregas"].sum()) if "Quantidade_Entregas" in df.columns else 0
        total_valor_entregas = float(df["Valor_Entregas"].sum()) if "Valor_Entregas" in df.columns else 0.0
        total_kg = float(df["KG_Excedente_Calculado"].sum()) if "KG_Excedente_Calculado" in df.columns else 0.0
        total_valor_kg = float(df["Valor_KG_Excedente"].sum()) if "Valor_KG_Excedente" in df.columns else 0.0
        total_geral = float(df["Total_Dia"].sum()) if "Total_Dia" in df.columns else 0.0

    acareacao = max(0.0, to_float(acareacao))
    vale = max(0.0, to_float(vale))
    desconto = max(0.0, to_float(desconto))
    bonus_extra = max(0.0, to_float(bonus_extra))
    bonus_sabados = max(0.0, to_float(bonus_sabados))
    bonus_feriado = max(0.0, to_float(bonus_feriado))
    subtotal_base = float(total_geral)
    total_geral = subtotal_base + acareacao + bonus_extra + bonus_sabados + bonus_feriado - vale - desconto

    periodo_txt = f"{fmt_data(data_inicio)} a {fmt_data(data_fim)}"
    quinzena_txt = limpar_texto(quinzena) or "Não informado"
    cnpj_txt = limpar_texto(cnpj_motorista) or obter_cnpj_motorista(motorista)

    elementos = []

    try:
        logo_bytes = base64.b64decode(LOGO_GDS_BASE64)
        logo = Image(io.BytesIO(logo_bytes), width=3.7 * cm, height=1.05 * cm)
        logo.hAlign = "CENTER"
        elementos.append(logo)
        elementos.append(Spacer(1, 0.06 * cm))
    except Exception:
        pass

    elementos.append(Paragraph("Recibo de pagamento de entregas realizadas", style_title))

    info_table = Table(
        [
            [
                Paragraph("Nome do Entregador", style_header),
                Paragraph("CNPJ", style_header),
                Paragraph("Quinzena", style_header),
            ],
            [
                Paragraph(str(motorista).upper(), style_normal),
                Paragraph(cnpj_txt, style_normal),
                Paragraph(quinzena_txt, style_normal),
            ],
            [
                Paragraph("Período", style_header),
                Paragraph("Total de Entregas", style_header),
                Paragraph("Total a Pagar", style_header),
            ],
            [
                Paragraph(periodo_txt, style_normal),
                Paragraph(str(total_entregas), style_normal),
                Paragraph(moeda(total_geral), style_normal),
            ],
        ],
        colWidths=[9.2 * cm, 5.0 * cm, 5.0 * cm],
    )
    info_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.65, laranja_linha),
        ("BACKGROUND", (0, 0), (-1, 0), laranja_claro),
        ("BACKGROUND", (0, 2), (-1, 2), laranja_claro),
        ("BACKGROUND", (0, 1), (-1, 1), branco),
        ("BACKGROUND", (0, 3), (-1, 3), branco),
        ("TEXTCOLOR", (0, 0), (-1, -1), preto),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 2), (-1, 2), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.8),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    elementos.append(info_table)
    elementos.append(Spacer(1, 0.18 * cm))

    linhas = [["DESCRIÇÃO", "CEP", "VALOR", "QUANTIDADE", "TOTAL"]]
    linhas_estilo_total_dia = []
    linhas_estilo_kg = []

    if usar_base_detalhada:
        df_resumo = df.copy()
        if "CEP Prefixo" not in df_resumo.columns:
            if "CEP" in df_resumo.columns:
                df_resumo["CEP Prefixo"] = df_resumo["CEP"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(8).str[:3]
            else:
                df_resumo["CEP Prefixo"] = ""

        df_resumo["CEP Prefixo"] = df_resumo["CEP Prefixo"].astype(str).apply(normalizar_prefixo_cep)
        df_resumo["Valor CEP"] = df_resumo["Valor CEP"].fillna(0).astype(float)
        if "Descrição Relatório" not in df_resumo.columns:
            df_resumo["Descrição Relatório"] = "Entregas realizadas"
        df_resumo["Descrição Relatório"] = df_resumo["Descrição Relatório"].fillna("Entregas realizadas").astype(str)

        resumo_cep = (
            df_resumo.groupby(["Descrição Relatório", "CEP Prefixo", "Valor CEP"], dropna=False)
            .agg(
                Quantidade=("Pedido", "size"),
                Total_Entregas=("Valor CEP", "sum"),
            )
            .reset_index()
            .sort_values(["Descrição Relatório", "CEP Prefixo", "Valor CEP"])
        )

        for _, item in resumo_cep.iterrows():
            descricao_linha = str(item.get("Descrição Relatório", "Entregas realizadas") or "Entregas realizadas")
            cep_prefixo = str(item.get("CEP Prefixo", "") or "").strip()
            cep_prefixo = normalizar_prefixo_cep(cep_prefixo) if cep_prefixo and descricao_linha != "RD Fechada" else ""
            valor_cep = float(item.get("Valor CEP", 0) or 0)
            qtd = int(item.get("Quantidade", 0) or 0)
            total_linha = float(item.get("Total_Entregas", 0) or 0)
            linhas.append([
                descricao_linha,
                cep_prefixo,
                moeda(valor_cep),
                str(qtd),
                moeda(total_linha),
            ])

        kg_excedente = float(df_resumo["KG Excedente"].sum()) if "KG Excedente" in df_resumo.columns else 0.0
        valor_kg = float(df_resumo["Valor Excedente KG"].sum()) if "Valor Excedente KG" in df_resumo.columns else 0.0
        if kg_excedente > 0:
            linhas_estilo_kg.append(len(linhas))
            linhas.append([
                "Kg excedente acima de 10kg",
                "",
                "R$ 0,30",
                fmt_num(kg_excedente),
                moeda(valor_kg),
            ])
    else:
        qtd = int(df.get("Quantidade_Entregas", pd.Series(dtype=float)).sum()) if "Quantidade_Entregas" in df.columns else 0
        valor_entregas = float(df.get("Valor_Entregas", pd.Series(dtype=float)).sum()) if "Valor_Entregas" in df.columns else 0.0
        kg_excedente = float(df.get("KG_Excedente_Calculado", pd.Series(dtype=float)).sum()) if "KG_Excedente_Calculado" in df.columns else 0.0
        valor_kg = float(df.get("Valor_KG_Excedente", pd.Series(dtype=float)).sum()) if "Valor_KG_Excedente" in df.columns else 0.0

        linhas.append(["Entregas realizadas", "", "Conforme CEP", str(qtd), moeda(valor_entregas)])
        if kg_excedente > 0:
            linhas_estilo_kg.append(len(linhas))
            linhas.append(["Kg excedente acima de 10kg", "", "R$ 0,30", fmt_num(kg_excedente), moeda(valor_kg)])

    linhas_estilo_total_dia.append(len(linhas))
    linhas.append(["Total das entregas", "", "", str(total_entregas), moeda(subtotal_base)])

    tabela = Table(
        linhas,
        colWidths=[7.0 * cm, 2.2 * cm, 3.0 * cm, 3.0 * cm, 4.0 * cm],
        repeatRows=1,
    )

    tabela_style = TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.55, laranja_linha),
        ("BACKGROUND", (0, 0), (-1, 0), laranja_claro),
        ("TEXTCOLOR", (0, 0), (-1, -1), preto),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.5),
        ("LEADING", (0, 0), (-1, -1), 7.2),
        ("TOPPADDING", (0, 0), (-1, -1), 1.4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.4),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("BACKGROUND", (0, 1), (-1, -1), branco),
    ])

    for i in linhas_estilo_kg:
        tabela_style.add("FONTNAME", (1, i), (-1, i), "Helvetica-Bold")
        tabela_style.add("TEXTCOLOR", (0, i), (-1, i), preto)
        tabela_style.add("BACKGROUND", (0, i), (-1, i), branco)

    for i in linhas_estilo_total_dia:
        tabela_style.add("BACKGROUND", (0, i), (-1, i), laranja_claro)
        tabela_style.add("FONTNAME", (1, i), (-1, i), "Helvetica-Bold")
        tabela_style.add("TEXTCOLOR", (0, i), (-1, i), preto)

    tabela.setStyle(tabela_style)
    elementos.append(tabela)
    elementos.append(Spacer(1, 0.14 * cm))

    assinatura = Table(
        [[""], ["Assinatura do Entregador"]],
        colWidths=[10.0 * cm],
        hAlign="LEFT",
    )
    assinatura.setStyle(TableStyle([
        ("LINEABOVE", (0, 1), (0, 1), 0.8, preto),
        ("ALIGN", (0, 1), (0, 1), "CENTER"),
        ("FONTSIZE", (0, 1), (0, 1), 7.2),
        ("TEXTCOLOR", (0, 1), (0, 1), preto),
        ("TOPPADDING", (0, 1), (0, 1), 5),
        ("BOTTOMPADDING", (0, 0), (0, 0), 18),
    ]))

    linhas_totais_finais = [
        ["Total Entregas", moeda(total_valor_entregas)],
        ["Total Kg Excedente", moeda(total_valor_kg)],
        ["Subtotal", moeda(subtotal_base)],
        ["Acareação", moeda(acareacao)],
        ["Bônus Extra", moeda(bonus_extra)],
        ["Bônus Sábados", moeda(bonus_sabados)],
        ["Bônus Feriado", moeda(bonus_feriado)],
        ["Vale", f"- {moeda(vale)}" if vale > 0 else moeda(0)],
        ["Desconto", f"- {moeda(desconto)}" if desconto > 0 else moeda(0)],
        ["TOTAL DO RECIBO", moeda(total_geral)],
    ]

    totais_finais = Table(
        linhas_totais_finais,
        colWidths=[4.0 * cm, 3.7 * cm],
        hAlign="RIGHT",
    )
    totais_finais.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.55, laranja_linha),
        ("BACKGROUND", (0, 0), (-1, -1), branco),
        ("BACKGROUND", (0, -1), (-1, -1), laranja_claro),
        ("TEXTCOLOR", (0, 0), (-1, -1), preto),
        ("FONTNAME", (0, 0), (-1, -2), "Helvetica"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6.8),
        ("TOPPADDING", (0, 0), (-1, -1), 1.6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6),
        ("ALIGN", (0, 0), (0, -1), "RIGHT"),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    bloco_assinatura_totais = Table(
        [[assinatura, totais_finais]],
        colWidths=[10.8 * cm, 7.9 * cm],
        hAlign="CENTER",
    )
    bloco_assinatura_totais.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    elementos.append(KeepTogether(bloco_assinatura_totais))
    elementos.append(Spacer(1, 0.12 * cm))

    observacao = (
        "Declaro que conferi as informações acima referentes às entregas realizadas, "
        "incluindo valores por CEP e eventual adicional por kg excedente acima de 10kg."
    )
    elementos.append(Paragraph(observacao, style_normal))
    elementos.append(Spacer(1, 0.08 * cm))

    elementos.append(Paragraph(
        f"Recibo gerado automaticamente em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        style_small_right,
    ))

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()


@st.cache_data(show_spinner="Calculando fechamento...")
def processar_fechamento_cache(
    excel_payloads: Tuple[Tuple[str, bytes], ...],
    pdf_payloads: Tuple[Tuple[str, bytes], ...],
    col_pedido: str,
    col_status: str,
    col_motivo_indicador: Optional[str],
    col_cep: str,
    col_rota_excel: Optional[str],
    col_data_excel: Optional[str],
    col_motorista_excel: Optional[str],
    status_entregue_tuple: Tuple[str, ...],
    status_ocorrencia_tuple: Tuple[str, ...],
    motoristas_extra_tuple: Tuple[Tuple[str, str, str], ...] = tuple(),
    reajustes_cep_tuple: Tuple[Tuple[str, str, float], ...] = tuple(),
    placas_excluidas_tuple: Tuple[str, ...] = tuple(),
):
    df_cep_final = montar_base_cep_final(reajustes_cep_tuple)
    df_placas_final = montar_base_placas_final(motoristas_extra_tuple, placas_excluidas_tuple)

    # Agora o Excel é usado somente como conferência:
    # Pedido + Status + CEP. Todo o restante vem dos PDFs.
    colunas_necessarias = tuple(dict.fromkeys([
        c for c in [col_pedido, col_status, col_motivo_indicador, col_cep, col_rota_excel, col_data_excel, col_motorista_excel]
        if c and str(c).strip()
    ]))
    df_excel_raw = carregar_excel_sistema_otimizado(excel_payloads, colunas_necessarias)
    if df_excel_raw.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_bonus_excel = preparar_base_bonus_excel(
        df_excel_raw,
        col_pedido=col_pedido,
        col_status=col_status,
        col_data_excel=col_data_excel,
        col_motorista_excel=col_motorista_excel,
        status_entregue=list(status_entregue_tuple),
        status_ocorrencia=list(status_ocorrencia_tuple),
        col_motivo_indicador=col_motivo_indicador,
    )

    df_metricas_excel = preparar_base_metricas_motorista_excel(
        df_excel_raw,
        col_pedido=col_pedido,
        col_status=col_status,
        col_data_excel=col_data_excel,
        col_motorista_excel=col_motorista_excel,
        status_entregue=list(status_entregue_tuple),
        status_ocorrencia=list(status_ocorrencia_tuple),
        col_motivo_indicador=col_motivo_indicador,
    )

    pdf_infos = []
    pdf_itens_frames = []
    for nome_arquivo, conteudo in pdf_payloads:
        info, itens_pdf = processar_pdf_cache(nome_arquivo, conteudo)
        pdf_infos.append(info)
        if not itens_pdf.empty:
            pdf_itens_frames.append(itens_pdf)

    df_pdf_info = pd.DataFrame(pdf_infos)
    df_pdf_itens = pd.concat(pdf_itens_frames, ignore_index=True) if pdf_itens_frames else pd.DataFrame()

    if df_pdf_itens.empty:
        return df_pdf_info, df_pdf_itens, pd.DataFrame(), pd.DataFrame(), df_bonus_excel, df_metricas_excel

    df_status_cep = preparar_planilha_status_cep(
        df_excel_raw,
        col_pedido=col_pedido,
        col_status=col_status,
        col_motivo_indicador=col_motivo_indicador,
        col_cep=col_cep,
        col_rota=col_rota_excel,
    )
    df_pedidos_pagos = filtrar_pedidos_pagos_excel(
        df_status_cep,
        list(status_entregue_tuple),
        list(status_ocorrencia_tuple),
    )

    df_pagamento_base = montar_entregas_pagas_pdf(df_pedidos_pagos, df_pdf_itens)
    if df_pagamento_base.empty:
        return df_pdf_info, df_pdf_itens, pd.DataFrame(), pd.DataFrame(), df_bonus_excel, df_metricas_excel

    df_pagamento = calcular_pagamento(df_pagamento_base, df_cep_final, df_placas_final)

    # Regra principal desta versão:
    # contar a quantidade de entregas do PDF que possuem status fechado no Excel.
    # Não deduplicar AWB/Pedido, pois o controle manual considera a ocorrência da entrega no manifesto.

    df_dia = gerar_fechamento_diario(df_pagamento)

    return df_pdf_info, df_pdf_itens, df_pagamento, df_dia, df_bonus_excel, df_metricas_excel


def processar_fechamento_com_progresso(
    excel_payloads: Tuple[Tuple[str, bytes], ...],
    pdf_payloads: Tuple[Tuple[str, bytes], ...],
    col_pedido: str,
    col_status: str,
    col_motivo_indicador: Optional[str],
    col_cep: str,
    col_rota_excel: Optional[str],
    col_data_excel: Optional[str],
    col_motorista_excel: Optional[str],
    status_entregue_tuple: Tuple[str, ...],
    status_ocorrencia_tuple: Tuple[str, ...],
    motoristas_extra_tuple: Tuple[Tuple[str, str, str], ...] = tuple(),
    reajustes_cep_tuple: Tuple[Tuple[str, str, float], ...] = tuple(),
    placas_excluidas_tuple: Tuple[str, ...] = tuple(),
):
    progresso_box = st.empty()
    barra = st.progress(0)

    def atualizar(percentual: int, mensagem: str):
        barra.progress(percentual)
        progresso_box.info(mensagem)

    atualizar(10, "📄 Lendo planilha Excel...")

    df_cep_final = montar_base_cep_final(reajustes_cep_tuple)
    df_placas_final = montar_base_placas_final(motoristas_extra_tuple, placas_excluidas_tuple)

    # Agora o Excel é usado somente como conferência:
    # Pedido + Status + CEP. Todo o restante vem dos PDFs.
    colunas_necessarias = tuple(dict.fromkeys([
        c for c in [col_pedido, col_status, col_motivo_indicador, col_cep, col_rota_excel, col_data_excel, col_motorista_excel]
        if c and str(c).strip()
    ]))
    df_excel_raw = carregar_excel_sistema_otimizado(excel_payloads, colunas_necessarias)
    if df_excel_raw.empty:
        atualizar(100, "✅ Fechamento concluído!")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_bonus_excel = preparar_base_bonus_excel(
        df_excel_raw,
        col_pedido=col_pedido,
        col_status=col_status,
        col_data_excel=col_data_excel,
        col_motorista_excel=col_motorista_excel,
        status_entregue=list(status_entregue_tuple),
        status_ocorrencia=list(status_ocorrencia_tuple),
        col_motivo_indicador=col_motivo_indicador,
    )

    df_metricas_excel = preparar_base_metricas_motorista_excel(
        df_excel_raw,
        col_pedido=col_pedido,
        col_status=col_status,
        col_data_excel=col_data_excel,
        col_motorista_excel=col_motorista_excel,
        status_entregue=list(status_entregue_tuple),
        status_ocorrencia=list(status_ocorrencia_tuple),
        col_motivo_indicador=col_motivo_indicador,
    )

    atualizar(30, "📑 Processando PDFs...")

    pdf_infos = []
    pdf_itens_frames = []
    total_pdfs = max(1, len(pdf_payloads))
    for idx_pdf, (nome_arquivo, conteudo) in enumerate(pdf_payloads, start=1):
        info, itens_pdf = processar_pdf_cache(nome_arquivo, conteudo)
        pdf_infos.append(info)
        if not itens_pdf.empty:
            pdf_itens_frames.append(itens_pdf)
        progresso_pdf = 30 + int((idx_pdf / total_pdfs) * 20)
        atualizar(min(progresso_pdf, 50), f"📑 Processando PDFs... ({idx_pdf}/{len(pdf_payloads)})")

    df_pdf_info = pd.DataFrame(pdf_infos)
    df_pdf_itens = pd.concat(pdf_itens_frames, ignore_index=True) if pdf_itens_frames else pd.DataFrame()

    if df_pdf_itens.empty:
        atualizar(100, "✅ Fechamento concluído!")
        return df_pdf_info, df_pdf_itens, pd.DataFrame(), pd.DataFrame(), df_bonus_excel, df_metricas_excel, {}

    atualizar(55, "🔄 Cruzando dados...")

    df_status_cep = preparar_planilha_status_cep(
        df_excel_raw,
        col_pedido=col_pedido,
        col_status=col_status,
        col_motivo_indicador=col_motivo_indicador,
        col_cep=col_cep,
        col_rota=col_rota_excel,
    )

    conferencia_awb = montar_conferencia_awb_pdf_excel(df_pdf_itens, df_status_cep)

    df_pedidos_pagos = filtrar_pedidos_pagos_excel(
        df_status_cep,
        list(status_entregue_tuple),
        list(status_ocorrencia_tuple),
    )

    df_pagamento_base = montar_entregas_pagas_pdf(df_pedidos_pagos, df_pdf_itens)
    if df_pagamento_base.empty:
        atualizar(100, "✅ Fechamento concluído!")
        return df_pdf_info, df_pdf_itens, pd.DataFrame(), pd.DataFrame(), df_bonus_excel, df_metricas_excel, conferencia_awb

    atualizar(75, "💰 Calculando pagamentos...")

    df_pagamento = calcular_pagamento(df_pagamento_base, df_cep_final, df_placas_final)

    # Regra principal desta versão:
    # contar a quantidade de entregas do PDF que possuem status fechado no Excel.
    # Não deduplicar AWB/Pedido, pois o controle manual considera a ocorrência da entrega no manifesto.

    df_dia = gerar_fechamento_diario(df_pagamento)

    atualizar(95, "📊 Gerando dashboard...")
    atualizar(100, "✅ Fechamento concluído!")

    return df_pdf_info, df_pdf_itens, df_pagamento, df_dia, df_bonus_excel, df_metricas_excel, conferencia_awb


# =========================================================
# HEADER
# =========================================================
st.markdown(
    """
    <div class="main-header">
        <div class="main-header-title">🚚 Fechamento de Entregadores</div>
        <div class="main-header-subtitle">
            Consolidação por PDF: Excel valida Pedido, Status e CEP; PDFs trazem rota, data, placa, motorista e peso.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.header("📁 Arquivos")

if "upload_reset_counter" not in st.session_state:
    st.session_state["upload_reset_counter"] = 0

if st.sidebar.button("🧹 Limpar cache", use_container_width=True):
    st.cache_data.clear()
    novo_contador_upload = int(st.session_state.get("upload_reset_counter", 0)) + 1
    for chave in list(st.session_state.keys()):
        del st.session_state[chave]
    st.session_state["upload_reset_counter"] = novo_contador_upload
    st.session_state["cache_limpo_msg"] = True
    st.rerun()

if st.session_state.pop("cache_limpo_msg", False):
    st.sidebar.success("Cache e arquivos carregados foram limpos com sucesso!")

_upload_reset_counter = int(st.session_state.get("upload_reset_counter", 0))

arquivos_excel_sistema = st.sidebar.file_uploader(
    "Planilha da quinzena / sistema",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key=f"arquivos_excel_sistema_{_upload_reset_counter}",
)

arquivos_pdf_rotas = st.sidebar.file_uploader(
    "PDFs das rotas",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"arquivos_pdf_rotas_{_upload_reset_counter}",
)

st.sidebar.success("Base de valores por CEP e placas carregada internamente.")

# Lê os bytes dos arquivos uma vez.
# Esses payloads são reutilizados na consulta de motoristas via PDF e no processamento do fechamento.
excel_payloads = []
if arquivos_excel_sistema:
    for f in arquivos_excel_sistema:
        f.seek(0)
        excel_payloads.append((f.name, f.read()))

pdf_payloads = []
if arquivos_pdf_rotas:
    for f in arquivos_pdf_rotas:
        f.seek(0)
        pdf_payloads.append((f.name, f.read()))

pdf_motoristas_tuple = extrair_motoristas_placas_dos_pdfs_cache(tuple(pdf_payloads)) if pdf_payloads else tuple()

st.sidebar.markdown("---")
st.sidebar.header("👤 Motoristas / CNPJ")

with st.sidebar.expander("➕ Cadastrar motorista e CNPJ", expanded=False):
    novo_nome_motorista_cnpj = st.text_input(
        "Nome do motorista",
        key="novo_nome_motorista_cnpj",
        placeholder="Ex: JOÃO DA SILVA",
    )
    novo_documento_motorista = st.text_input(
        "CNPJ do motorista",
        key="novo_documento_motorista",
        placeholder="Ex: 00.000.000/0001-00",
    )

    if st.button("Salvar motorista/CNPJ", use_container_width=True):
        ok, msg = salvar_motorista_cnpj_extra(novo_nome_motorista_cnpj, novo_documento_motorista)
        if ok:
            st.success(msg)
            st.session_state.pop("resultado_fechamento", None)
            st.cache_data.clear()
            st.rerun()
        else:
            st.error(msg)

    df_motoristas_cnpj_extra = carregar_motoristas_cnpj_extra_csv()
    if not df_motoristas_cnpj_extra.empty:
        st.caption("Motoristas/CNPJ cadastrados manualmente:")
        st.dataframe(df_motoristas_cnpj_extra, use_container_width=True, hide_index=True, height=180)
    else:
        st.caption("Nenhum motorista/CNPJ extra cadastrado ainda.")

with st.sidebar.expander("🗑️ Excluir motorista/CNPJ", expanded=False):
    df_base_para_excluir = montar_base_cnpj_motoristas_final()
    opcoes_excluir = df_base_para_excluir["Nome Motorista"].tolist() if not df_base_para_excluir.empty else []

    motorista_excluir = st.selectbox(
        "Motorista para excluir",
        options=[""] + opcoes_excluir,
        key="motorista_excluir_cnpj",
    )
    nome_excluir_digitado = st.text_input(
        "Ou digite o nome do motorista",
        key="nome_excluir_motorista_cnpj",
        placeholder="Ex: JOÃO DA SILVA",
    )
    confirmar_exclusao = st.checkbox(
        "Confirmo que desejo excluir/bloquear este motorista da base aplicada de CNPJ",
        key="confirmar_exclusao_motorista_cnpj",
    )

    if st.button("Excluir motorista/CNPJ", use_container_width=True):
        nome_para_excluir = nome_excluir_digitado.strip() or motorista_excluir
        if not confirmar_exclusao:
            st.error("Marque a confirmação para excluir o motorista.")
        else:
            ok, msg = excluir_motorista_cnpj(nome_para_excluir)
            if ok:
                st.success(msg)
                st.session_state.pop("resultado_fechamento", None)
                st.cache_data.clear()
                st.rerun()
            else:
                st.error(msg)

if st.sidebar.button("📋 Ver motoristas/CNPJ cadastrados", use_container_width=True):
    st.session_state["mostrar_base_cnpj_motoristas"] = not st.session_state.get("mostrar_base_cnpj_motoristas", False)

st.sidebar.caption("Os motoristas/CNPJ cadastrados manualmente ficam salvos no arquivo motoristas_cnpj_extra.csv e sobrescrevem a base interna quando o nome for igual.")

st.sidebar.markdown("---")
st.sidebar.header("🚗 Placas / Tipo de Veículo")

df_motoristas_extra = carregar_motoristas_extra_csv()
df_placas_excluidas = carregar_placas_excluidas_csv()
motoristas_extra_tuple = preparar_motoristas_extra_para_cache(df_motoristas_extra)
placas_excluidas_tuple = preparar_placas_excluidas_para_cache(df_placas_excluidas)

with st.sidebar.expander("➕ Cadastrar placa/tipo", expanded=False):
    nova_placa = st.text_input("Placa do veículo", key="novo_motorista_placa", placeholder="Ex: ABC1D23")
    novo_tipo = st.selectbox("Tipo de veículo", ["CARRO", "MOTO"], key="novo_motorista_tipo")

    if st.button("Salvar placa", use_container_width=True):
        ok, msg = salvar_motorista_extra(nova_placa, novo_tipo)
        if ok:
            st.success(msg)
            st.session_state.pop("resultado_fechamento", None)
            st.cache_data.clear()
            st.rerun()
        else:
            st.error(msg)

    if not df_motoristas_extra.empty:
        st.caption("Placas manuais salvas:")
        st.dataframe(df_motoristas_extra[["Placa", "Tipo Veículo"]], use_container_width=True, hide_index=True, height=160)
    else:
        st.caption("Nenhuma placa extra cadastrada ainda.")

with st.sidebar.expander("🗑️ Excluir placa", expanded=False):
    base_para_excluir = montar_base_placas_final(motoristas_extra_tuple, placas_excluidas_tuple)
    placas_disponiveis = sorted(base_para_excluir["Placa"].dropna().astype(str).unique().tolist()) if not base_para_excluir.empty else []

    placa_select = st.selectbox(
        "Selecionar placa da base",
        [""] + placas_disponiveis,
        format_func=lambda x: "Digite manualmente ou selecione..." if x == "" else x,
        key="excluir_placa_select",
    )
    placa_manual_excluir = st.text_input(
        "Ou digite a placa para excluir",
        key="excluir_placa_manual",
        placeholder="Ex: ABC1D23",
    )
    placa_para_excluir = placa_manual_excluir.strip() or placa_select

    if st.button("Excluir placa da base aplicada", use_container_width=True):
        ok, msg = excluir_placa_base(placa_para_excluir)
        if ok:
            st.success(msg)
            st.session_state.pop("resultado_fechamento", None)
            st.cache_data.clear()
            st.rerun()
        else:
            st.error(msg)

    if not df_placas_excluidas.empty:
        st.caption("Placas atualmente excluídas/bloqueadas:")
        st.dataframe(df_placas_excluidas[["Placa"]], use_container_width=True, hide_index=True, height=120)

st.sidebar.caption("Os cadastros manuais ficam salvos no arquivo motoristas_placas_extra.csv. As placas excluídas ficam salvas em placas_excluidas.csv e não entram na base aplicada.")


if st.sidebar.button("📋 Ver todas as placas da base", use_container_width=True):
    st.session_state["mostrar_base_motoristas"] = not st.session_state.get("mostrar_base_motoristas", False)


st.sidebar.markdown("---")
st.sidebar.header("💰 Reajustes por CEP")

df_reajustes_cep = carregar_reajustes_cep_csv()
reajustes_cep_tuple = preparar_reajustes_cep_para_cache(df_reajustes_cep)

with st.sidebar.expander("➕ Cadastrar novo CEP", expanded=False):
    cadastro_prefixo_cep = st.text_input("Prefixo do CEP novo", key="cadastro_prefixo_cep", placeholder="Ex: 010, 045, 056")
    cadastro_valor_moto = st.text_input("Valor para MOTO", key="cadastro_valor_moto", placeholder="Ex: 7,00")
    cadastro_valor_carro = st.text_input("Valor para CARRO", key="cadastro_valor_carro", placeholder="Ex: 8,50")

    if st.button("Salvar novo CEP", use_container_width=True):
        ok, msg = salvar_novo_cep(cadastro_prefixo_cep, cadastro_valor_moto, cadastro_valor_carro)
        if ok:
            st.success(msg)
            st.session_state.pop("resultado_fechamento", None)
            st.cache_data.clear()
            st.rerun()
        else:
            st.error(msg)

with st.sidebar.expander("✏️ Reajustar valor por CEP", expanded=False):
    novo_prefixo_cep = st.text_input("Prefixo do CEP", key="novo_reajuste_prefixo_cep", placeholder="Ex: 010, 045, 056")
    novo_tipo_cep = st.selectbox("Tipo de veículo", ["CARRO", "MOTO"], key="novo_reajuste_tipo_veiculo")
    novo_valor_cep = st.text_input("Novo valor da entrega", key="novo_reajuste_valor_cep", placeholder="Ex: 8,50")

    if st.button("Salvar reajuste de CEP", use_container_width=True):
        ok, msg = salvar_reajuste_cep(novo_prefixo_cep, novo_tipo_cep, novo_valor_cep)
        if ok:
            st.success(msg)
            st.session_state.pop("resultado_fechamento", None)
            st.cache_data.clear()
            st.rerun()
        else:
            st.error(msg)

    if not df_reajustes_cep.empty:
        st.caption("CEPs/reajustes manuais salvos:")
        df_reajustes_view = df_reajustes_cep.copy()
        df_reajustes_view["Valor CEP"] = df_reajustes_view["Valor CEP"].apply(moeda)
        st.dataframe(df_reajustes_view, use_container_width=True, hide_index=True, height=180)
    else:
        st.caption("Nenhum CEP/reajuste cadastrado ainda.")

st.sidebar.caption("Os novos CEPs e reajustes ficam salvos no arquivo reajustes_valores_cep.csv e entram na base aplicada por CEP Prefixo + Tipo de Veículo.")


if st.sidebar.button("💵 Ver todos os valores aplicados por CEP", use_container_width=True):
    st.session_state["mostrar_base_valores_cep"] = not st.session_state.get("mostrar_base_valores_cep", False)

# =========================================================
# CONSULTAS RÁPIDAS DAS BASES INTERNAS / CADASTRADAS
# =========================================================
if st.session_state.get("mostrar_base_cnpj_motoristas", False):
    st.markdown('<div class="section-heading">Base completa de motoristas e CNPJ aplicada</div>', unsafe_allow_html=True)

    df_base_cnpj_view = montar_base_cnpj_motoristas_final().copy()
    total_manual_cnpj = int((df_base_cnpj_view["Origem"] == "Cadastro manual").sum()) if "Origem" in df_base_cnpj_view.columns else 0
    total_interna_cnpj = int((df_base_cnpj_view["Origem"] == "Base interna").sum()) if "Origem" in df_base_cnpj_view.columns else 0

    st.info(
        f"Total de motoristas na base aplicada: **{len(df_base_cnpj_view)}** | "
        f"Base interna: **{total_interna_cnpj}** | "
        f"Cadastro manual: **{total_manual_cnpj}**"
    )
    st.dataframe(df_base_cnpj_view, use_container_width=True, hide_index=True, height=420)
    st.download_button(
        "⬇️ Baixar base de motoristas/CNPJ em CSV",
        data=df_base_cnpj_view.to_csv(index=False, sep=";", encoding="utf-8-sig"),
        file_name="base_motoristas_cnpj_aplicada.csv",
        mime="text/csv",
        use_container_width=True,
    )

if st.session_state.get("mostrar_base_motoristas", False):
    st.markdown('<div class="section-heading">Base completa de placas e tipos de veículo aplicada</div>', unsafe_allow_html=True)

    df_base_motoristas_view = montar_base_placas_final(motoristas_extra_tuple, placas_excluidas_tuple).copy()
    df_base_motoristas_view["Placa"] = df_base_motoristas_view["Placa"].apply(normalizar_placa)
    df_base_motoristas_view["Base de Dados"] = df_base_motoristas_view.get("Origem", "Base interna")

    df_base_motoristas_view = df_base_motoristas_view[
        ["Placa", "Tipo Veículo", "Base de Dados"]
    ].sort_values(["Base de Dados", "Placa"]).reset_index(drop=True)

    total_manual = int((df_base_motoristas_view["Base de Dados"] == "Cadastro manual").sum())
    total_interna = int((df_base_motoristas_view["Base de Dados"] == "Base interna").sum())

    st.info(
        f"Total de placas na base aplicada: **{len(df_base_motoristas_view)}** | "
        f"Base interna: **{total_interna}** | "
        f"Cadastro manual: **{total_manual}**"
    )
    st.dataframe(df_base_motoristas_view, use_container_width=True, hide_index=True, height=420)
    st.download_button(
        "⬇️ Baixar base de placas em CSV",
        data=df_base_motoristas_view.to_csv(index=False, sep=";", encoding="utf-8-sig"),
        file_name="base_placas_tipos_aplicada.csv",
        mime="text/csv",
        use_container_width=True,
    )
    st.caption("Consulta simplificada: mostra somente placa, tipo de veículo e origem da base de dados aplicada no cálculo.")
    st.markdown("---")

if st.session_state.get("mostrar_base_valores_cep", False):
    st.markdown('<div class="section-heading">Valores por CEP aplicados no cálculo</div>', unsafe_allow_html=True)
    df_base_cep_view = montar_base_cep_final(reajustes_cep_tuple).copy()
    reajustes_chaves = set(
        zip(
            df_reajustes_cep.get("CEP Prefixo", pd.Series(dtype=str)).astype(str),
            df_reajustes_cep.get("Tipo Veículo", pd.Series(dtype=str)).astype(str),
        )
    )
    df_base_cep_view["Origem"] = df_base_cep_view.apply(
        lambda row: "Reajuste manual" if (str(row["CEP Prefixo"]), str(row["Tipo Veículo"])) in reajustes_chaves else "Base interna",
        axis=1,
    )
    df_base_cep_view = df_base_cep_view.sort_values(["CEP Prefixo", "Tipo Veículo"]).reset_index(drop=True)
    df_base_cep_display = df_base_cep_view.copy()
    df_base_cep_display["Valor CEP"] = df_base_cep_display["Valor CEP"].apply(moeda)
    st.info(f"Total de regras de CEP aplicadas: **{len(df_base_cep_view)}**")
    st.dataframe(df_base_cep_display, use_container_width=True, hide_index=True, height=420)
    st.download_button(
        "⬇️ Baixar valores por CEP em CSV",
        data=df_base_cep_view.to_csv(index=False, sep=";", encoding="utf-8-sig"),
        file_name="valores_por_cep_aplicados.csv",
        mime="text/csv",
        use_container_width=True,
    )
    st.markdown("---")


st.sidebar.markdown("---")
st.sidebar.header("⚙️ Configuração")

# Carregamento prévio das planilhas para selecionar colunas.
# Os bytes já foram lidos acima para permitir consulta de motoristas/placas dos PDFs.

# Antes, a tela carregava o Excel inteiro só para descobrir as colunas.
# Agora carrega apenas o cabeçalho; a planilha completa só é lida quando clicar em Processar.
cols_excel = []
if excel_payloads:
    cols_excel = ler_excel_colunas_bytes(excel_payloads[0][0], excel_payloads[0][1], sheet_name=0)

if not cols_excel:
    st.info("Carregue a planilha da quinzena para configurar as colunas e gerar o fechamento.")
    st.stop()

df_sistema_raw = pd.DataFrame(columns=cols_excel)

# Sugestões de colunas.
# Nesta versão, o Excel é usado somente para: Pedido + Status + CEP.
# Data, rota, placa, motorista e peso taxado vêm direto dos PDFs.
col_pedido_auto = detectar_coluna(df_sistema_raw, ["pedido", "numero pedido", "id pedido", "awb", "cte", "encomenda"])
col_status_auto = detectar_coluna(df_sistema_raw, ["status", "ocorrencia", "situacao"])
col_motivo_indicador_auto = detectar_coluna(df_sistema_raw, ["Motivo 1", "motivo 1", "motivo_1", "motivo", "ocorrencia", "ocorrência"])
col_cep_auto = detectar_coluna(df_sistema_raw, ["cep", "cep destino", "cep entrega"])
col_rota_auto = detectar_coluna(df_sistema_raw, ["carga", "rota", "route", "manifesto", "romaneio"])
col_data_excel_auto = detectar_coluna(df_sistema_raw, ["data", "data entrega", "data rota", "data baixa", "dt entrega", "dt baixa", "finalizado em"])
col_motorista_excel_auto = detectar_coluna(df_sistema_raw, ["motorista", "entregador", "driver", "nome motorista", "nome entregador", "prestador"])

cols = df_sistema_raw.columns.tolist()

def idx(col):
    return cols.index(col) if col in cols else 0

# =========================================================
# BLOQUEIO DAS CONFIGURAÇÕES DE COLUNAS
# =========================================================
SENHA_CONFIGURACAO = "2305"
if "configuracoes_desbloqueadas" not in st.session_state:
    st.session_state["configuracoes_desbloqueadas"] = False

config_bloqueada = not st.session_state.get("configuracoes_desbloqueadas", False)

col_pedido = st.sidebar.selectbox(
    "Coluna do Pedido/AWB",
    cols,
    index=idx(col_pedido_auto),
    disabled=config_bloqueada,
)
col_status = st.sidebar.selectbox(
    "Coluna do Status",
    cols,
    index=idx(col_status_auto),
    disabled=config_bloqueada,
)
col_motivo_indicador = st.sidebar.selectbox(
    "Coluna do status do indicador (Motivo 1)",
    cols,
    index=idx(col_motivo_indicador_auto),
    disabled=config_bloqueada,
    help="Use a coluna Motivo 1. Para o indicador: Motivo 1 vazio = realizada; Motivo 1 preenchido = ocorrência/não realizada.",
)
col_cep = st.sidebar.selectbox(
    "Coluna do CEP",
    cols,
    index=idx(col_cep_auto),
    disabled=config_bloqueada,
)
col_rota_excel = st.sidebar.selectbox(
    "Coluna da Rota/Carga",
    cols,
    index=idx(col_rota_auto),
    disabled=config_bloqueada,
)
col_data_excel = st.sidebar.selectbox(
    "Coluna da Data no Excel",
    cols,
    index=idx(col_data_excel_auto),
    disabled=config_bloqueada,
    help="Usada para calcular bônus de sábados e feriados pelo Excel.",
)
col_motorista_excel = st.sidebar.selectbox(
    "Coluna do Motorista no Excel",
    cols,
    index=idx(col_motorista_excel_auto),
    disabled=config_bloqueada,
    help="Usada para calcular bônus de sábados e feriados pelo Excel.",
)

st.sidebar.caption("O pagamento principal usa PDF + Excel. O indicador usa Data, Motorista, Pedido e a coluna Motivo 1: confirmações de entrega contam como realizadas; ocorrências configuradas contam como não realizadas.")

with st.sidebar.expander("🔐 Alterar filtros de colunas", expanded=False):
    if st.session_state.get("configuracoes_desbloqueadas", False):
        st.success("Configurações desbloqueadas.")
        if st.button("🔒 Bloquear novamente", use_container_width=True):
            st.session_state["configuracoes_desbloqueadas"] = False
            st.rerun()
    else:
        senha_digitada = st.text_input("Senha para liberar alteração", type="password")
        if st.button("🔓 Desbloquear configurações", use_container_width=True):
            if senha_digitada == SENHA_CONFIGURACAO:
                st.session_state["configuracoes_desbloqueadas"] = True
                st.success("Configurações liberadas.")
                st.rerun()
            else:
                st.error("Senha incorreta.")

st.sidebar.markdown("---")
st.sidebar.header("✅ Status pago")
status_entregue = st.sidebar.text_area(
    "Status considerados como entrega realizada",
    value="fechada\nentregue\nrealizada\nfinalizado\nbaixado",
    height=100,
).splitlines()

status_ocorrencia = st.sidebar.text_area(
    "Status considerados ocorrência/não pagar",
    value="devolvido\npedido não liberado\npedido nao liberado\nausente\nrecusado\ncancelado\ninsucesso\nocorrencia\nocorrência\nnão entregue\nnao entregue",
    height=110,
).splitlines()

st.sidebar.markdown("---")
processar = st.sidebar.button("🚀 Processar fechamento", use_container_width=True)


# =========================================================
# PROCESSAMENTO
# =========================================================
if processar:
    resultado_cache = processar_fechamento_com_progresso(
        excel_payloads=tuple(excel_payloads),
        pdf_payloads=tuple(pdf_payloads),
        col_pedido=col_pedido,
        col_status=col_status,
        col_motivo_indicador=col_motivo_indicador,
        col_cep=col_cep,
        col_rota_excel=col_rota_excel,
        col_data_excel=col_data_excel,
        col_motorista_excel=col_motorista_excel,
        status_entregue_tuple=tuple(status_entregue),
        status_ocorrencia_tuple=tuple(status_ocorrencia),
        motoristas_extra_tuple=motoristas_extra_tuple,
        reajustes_cep_tuple=reajustes_cep_tuple,
        placas_excluidas_tuple=placas_excluidas_tuple,
    )

    # Compatibilidade: versões anteriores retornavam 5 ou 6 itens; versões novas retornam 7.
    # Caso algum cache antigo/ajuste intermediário retorne mais itens, pegamos apenas os necessários.
    if len(resultado_cache) >= 7:
        df_pdf_info, df_pdf_itens, df_pagamento, df_dia, df_bonus_excel, df_metricas_excel, conferencia_awb = resultado_cache[:7]
    elif len(resultado_cache) >= 6:
        df_pdf_info, df_pdf_itens, df_pagamento, df_dia, df_bonus_excel, df_metricas_excel = resultado_cache[:6]
        conferencia_awb = {}
    elif len(resultado_cache) == 5:
        df_pdf_info, df_pdf_itens, df_pagamento, df_dia, df_bonus_excel = resultado_cache
        df_metricas_excel = df_bonus_excel.copy() if isinstance(df_bonus_excel, pd.DataFrame) else pd.DataFrame()
        conferencia_awb = {}
    else:
        raise ValueError(f"Retorno inesperado do processamento: {len(resultado_cache)} itens.")
    st.session_state["resultado_fechamento"] = {
        "df_pdf_info": df_pdf_info,
        "df_pdf_itens": df_pdf_itens,
        "df_pagamento": df_pagamento,
        "df_dia": df_dia,
        "df_bonus_excel": df_bonus_excel,
        "df_metricas_excel": df_metricas_excel,
        "conferencia_awb": conferencia_awb,
    }

if "resultado_fechamento" not in st.session_state:
    st.info("Configure os arquivos e clique em **Processar fechamento** para gerar a dashboard.")
    st.stop()

resultado = st.session_state["resultado_fechamento"]
df_pdf_info = resultado["df_pdf_info"]
df_pdf_itens = resultado["df_pdf_itens"]
df_pagamento = resultado["df_pagamento"]
df_dia = resultado["df_dia"]
df_bonus_excel = resultado.get("df_bonus_excel", pd.DataFrame())
df_metricas_excel = resultado.get("df_metricas_excel", df_bonus_excel if isinstance(df_bonus_excel, pd.DataFrame) else pd.DataFrame())
df_metricas_excel = resultado.get("df_metricas_excel", pd.DataFrame())
conferencia_awb = resultado.get("conferencia_awb", {})

# =========================================================
# DASHBOARD
# =========================================================
renderizar_aviso_conferencia_awb(conferencia_awb)

total_entregas = int(len(df_pagamento)) if not df_pagamento.empty else 0
total_motoristas = int(df_dia["Motorista Final"].nunique()) if not df_dia.empty else 0
total_rotas = int(df_pagamento["Rota"].nunique()) if not df_pagamento.empty else 0
valor_total = float(df_pagamento["Total Entrega"].sum()) if not df_pagamento.empty else 0.0

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">Entregas pagas</div><div class="kpi-value">{total_entregas}</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">Motoristas</div><div class="kpi-value">{total_motoristas}</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">Rotas</div><div class="kpi-value">{total_rotas}</div></div>', unsafe_allow_html=True)
with c4:
    total_a_pagar_placeholder = st.empty()
    total_a_pagar_placeholder.markdown(
        f'<div class="kpi-card"><div class="kpi-title">Total a pagar</div><div class="kpi-value">{moeda(valor_total)}</div></div>',
        unsafe_allow_html=True,
    )


# =========================================================
# =========================================================
# FECHAMENTO DIÁRIO
# =========================================================
if df_dia.empty:
    st.warning("Nenhuma entrega paga foi encontrada com as regras configuradas.")
    st.stop()

motoristas = sorted([
    m for m in df_dia["Motorista Final"].fillna("").astype(str).str.strip().unique()
    if m
])

st.markdown('<div class="section-heading">Fechamento diário por entregador</div>', unsafe_allow_html=True)

motoristas_disponiveis_fechamento = sorted(
    df_dia["Motorista Final"].dropna().astype(str).str.strip().unique()
) if "Motorista Final" in df_dia.columns else []
motoristas_disponiveis_fechamento = [m for m in motoristas_disponiveis_fechamento if m]

motorista_filtro = st.multiselect("Filtrar entregador", motoristas_disponiveis_fechamento, default=motoristas_disponiveis_fechamento)

df_dia_view = df_dia.copy()
if motorista_filtro:
    df_dia_view = df_dia_view[df_dia_view["Motorista Final"].isin(motorista_filtro)]

df_dia_view_display = df_dia_view.copy()
# Organiza as colunas do fechamento diário para mostrar o Peso Taxado usado no cálculo.
colunas_fechamento = [
    "Motorista Final", "Data Rota", "Rotas", "Quantidade_Entregas",
    "Valor_Entregas", "KG_Excedente_Calculado",
    "Valor_KG_Excedente", "Total_Dia",
]
colunas_fechamento = [c for c in colunas_fechamento if c in df_dia_view_display.columns]
df_dia_view_display = df_dia_view_display[colunas_fechamento]

for col in ["Valor_Entregas", "Valor_KG_Excedente", "Total_Dia"]:
    if col in df_dia_view_display.columns:
        df_dia_view_display[col] = df_dia_view_display[col].apply(moeda)

for col in ["KG_Excedente_Calculado"]:
    if col in df_dia_view_display.columns:
        df_dia_view_display[col] = df_dia_view_display[col].fillna(0).astype(float).round(2)

st.dataframe(df_dia_view_display, use_container_width=True, height=420)

st.markdown("---")
st.markdown('<div class="section-heading">Gerar relatório de entregas / recibo de pagamento</div>', unsafe_allow_html=True)

col_rec1, col_rec_q, col_rec2, col_rec3 = st.columns([2.0, 1.2, 1.3, 1.3])

with col_rec1:
    motorista_recibo = st.selectbox(
        "Motorista para o recibo",
        motoristas,
        index=0 if motoristas else None,
    )

with col_rec_q:
    quinzena_recibo = st.selectbox(
        "Quinzena",
        ["1ª Quinzena", "2ª Quinzena"],
        index=1,
    )

_datas_disponiveis = pd.to_datetime(df_dia["Data Rota"], errors="coerce").dropna()
_data_min = _datas_disponiveis.min().date() if not _datas_disponiveis.empty else datetime.now().date()
_data_max = _datas_disponiveis.max().date() if not _datas_disponiveis.empty else datetime.now().date()

with col_rec2:
    data_inicio_recibo = st.date_input(
        "Data inicial",
        value=_data_min,
        min_value=_data_min,
        max_value=_data_max,
    )

with col_rec3:
    data_fim_recibo = st.date_input(
        "Data final",
        value=_data_max,
        min_value=_data_min,
        max_value=_data_max,
    )

cnpj_motorista_recibo = obter_cnpj_motorista(motorista_recibo)
st.markdown(
    f"""
    <div style="background:#f8fafc;border:1px solid #e5e7eb;border-radius:14px;padding:12px 16px;margin:4px 0 16px 0;">
        <span style="color:#64748b;font-weight:700;">CNPJ do motorista:</span>
        <span style="color:#111827;font-weight:850;margin-left:8px;">{cnpj_motorista_recibo}</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# Opção de RD Fechada: substitui o cálculo normal da rota pelo valor fixo de R$ 250,00.
df_rd_opcoes = df_pagamento.copy()
df_rd_opcoes["Data Rota DT"] = pd.to_datetime(df_rd_opcoes["Data Rota"], errors="coerce").dt.date
df_rd_opcoes = df_rd_opcoes[
    (df_rd_opcoes["Motorista Final"].astype(str) == str(motorista_recibo))
    & (df_rd_opcoes["Data Rota DT"] >= data_inicio_recibo)
    & (df_rd_opcoes["Data Rota DT"] <= data_fim_recibo)
].copy()

rds_disponiveis_fechada = sorted([
    str(rd).strip().upper()
    for rd in df_rd_opcoes.get("Rota", pd.Series(dtype=str)).dropna().unique().tolist()
    if str(rd).strip()
])

st.markdown("**RD Fechada**")
col_rd1, col_rd2 = st.columns([1.0, 3.0])
with col_rd1:
    possui_rd_fechada = st.radio(
        "Motorista possui RD fechada?",
        ["Não", "Sim"],
        horizontal=True,
        index=0,
        help="Quando marcado como Sim, a RD selecionada passa a valer R$ 250,00 e ignora kg excedente e bônus daquela rota.",
    )

with col_rd2:
    rds_fechadas_recibo = []
    if possui_rd_fechada == "Sim":
        rds_fechadas_recibo = st.multiselect(
            "Selecione as RDs fechadas",
            options=rds_disponiveis_fechada,
            default=[],
            help="Cada RD selecionada será paga como RD Fechada no valor fixo de R$ 250,00.",
        )
        if rds_fechadas_recibo:
            st.success(f"RD Fechada: {len(rds_fechadas_recibo)} RD(s) × R$ 250,00 = {moeda(len(rds_fechadas_recibo) * 250.0)}")
        else:
            st.info("Selecione uma ou mais RDs para aplicar o valor fechado de R$ 250,00.")
    else:
        rds_fechadas_recibo = []

st.markdown("**Ajustes manuais e bônus do recibo**")
col_aj1, col_aj2, col_aj3, col_aj4 = st.columns(4)
with col_aj1:
    acareacao_recibo = st.number_input(
        "Acareação (+)",
        min_value=0.0,
        value=0.0,
        step=1.0,
        format="%.2f",
        help="Valor adicional pago ao motorista por acareação.",
    )
with col_aj2:
    vale_recibo = st.number_input(
        "Vale (-)",
        min_value=0.0,
        value=0.0,
        step=1.0,
        format="%.2f",
        help="Desconto de antecipação/vale já pago ao motorista.",
    )
with col_aj3:
    desconto_recibo = st.number_input(
        "Desconto (-)",
        min_value=0.0,
        value=0.0,
        step=1.0,
        format="%.2f",
        help="Desconto por perda, avaria ou outro desconto operacional.",
    )
with col_aj4:
    bonus_extra_recibo = st.number_input(
        "Bônus Extra (+)",
        min_value=0.0,
        value=0.0,
        step=1.0,
        format="%.2f",
        help="Valor extra acordado manualmente com o motorista.",
    )

st.markdown("**Bônus automáticos pagos somente na 2ª quinzena**")

col_b1, col_b3 = st.columns([1.1, 1.9])
with col_b1:
    valor_bonus_por_entrega = st.number_input(
        "Valor por entrega bônus",
        min_value=0.0,
        value=2.0,
        step=0.5,
        format="%.2f",
        help="Valor usado para bônus de sábado e feriado. Hoje: R$ 2,00 por entrega.",
    )
with col_b3:
    datas_feriado_txt = st.text_area(
        "Datas de feriado do mês",
        value="",
        height=78,
        placeholder="Exemplo:\n01/05/2026\n09/07/2026",
        help="Digite um feriado por linha. A dashboard buscará no Excel as entregas do motorista nessas datas.",
    )

# Para o recibo bater com o fechamento manual, usamos a base DETALHADA de entregas.
# Assim o recibo separa cada dia por valor real do CEP: R$ 7,00, R$ 8,50, R$ 8,35 etc.
df_recibo = df_pagamento.copy()
df_recibo["Data Rota DT"] = pd.to_datetime(df_recibo["Data Rota"], errors="coerce").dt.date
df_recibo = df_recibo[
    (df_recibo["Motorista Final"].astype(str) == str(motorista_recibo))
    & (df_recibo["Data Rota DT"] >= data_inicio_recibo)
    & (df_recibo["Data Rota DT"] <= data_fim_recibo)
].copy()

# Aplica RD Fechada no recibo/relatório: remove o cálculo normal da RD e substitui por R$ 250,00.
df_recibo = aplicar_rd_fechada_recibo(df_recibo, rds_fechadas_recibo, valor_rd_fechada=250.0)

# Dados usados para preencher os adicionais no Excel consolidado.
df_relatorio_entregas_excel = pd.DataFrame()
motorista_relatorio_excel = ""
data_inicio_relatorio_excel = None
data_fim_relatorio_excel = None
acareacao_relatorio_excel = 0.0
vale_relatorio_excel = 0.0
desconto_relatorio_excel = 0.0
bonus_extra_relatorio_excel = 0.0
bonus_sabados_relatorio_excel = 0.0
bonus_feriado_relatorio_excel = 0.0
df_dia_excel = df_dia.copy()

if df_recibo.empty:
    st.warning("Não há dados para gerar recibo com o motorista e período selecionados.")
else:
    subtotal_recibo = float(df_recibo["Total Entrega"].sum())
    acareacao_recibo = max(0.0, to_float(acareacao_recibo))
    vale_recibo = max(0.0, to_float(vale_recibo))
    desconto_recibo = max(0.0, to_float(desconto_recibo))
    bonus_extra_recibo = max(0.0, to_float(bonus_extra_recibo))
    bonus_sabados_recibo = 0.0
    bonus_feriado_recibo = 0.0
    qtd_entregas_sabado = 0
    qtd_entregas_feriado = 0
    veio_todos_sabados = False

    ano_ref = data_fim_recibo.year
    mes_ref = data_fim_recibo.month
    datas_feriado = parse_datas_feriados(datas_feriado_txt, ano_ref)

    detalhes_sabado = detalhar_bonus_sabados_excel(
        df_bonus_excel,
        motorista_recibo,
        ano_ref,
        mes_ref,
        valor_bonus_por_entrega,
    )

    todos_sabados_mes = detalhes_sabado.get("todos_sabados", [])
    sabados_trabalhados = detalhes_sabado.get("sabados_trabalhados", [])
    sabados_faltantes = detalhes_sabado.get("sabados_faltantes", [])
    qtd_entregas_sabado = int(detalhes_sabado.get("qtd_entregas_sabado", 0) or 0)
    veio_todos_sabados = bool(detalhes_sabado.get("veio_todos_sabados", False))
    bonus_sabado_calculado = float(detalhes_sabado.get("bonus_calculado", 0.0) or 0.0)
    bonus_sabado_forcado = qtd_entregas_sabado * max(0.0, to_float(valor_bonus_por_entrega))

    aplicar_bonus_sabado_sem_requisitos = st.radio(
        "Aplicar bônus de sábado mesmo sem todos os requisitos?",
        options=["Não", "Sim"],
        index=0,
        horizontal=True,
        help="Use Sim apenas quando quiser liberar manualmente o bônus de sábado mesmo faltando algum sábado do mês.",
    )

    bonus_sabado_exibicao = bonus_sabado_calculado
    if aplicar_bonus_sabado_sem_requisitos == "Sim" and quinzena_recibo == "2ª Quinzena":
        bonus_sabado_exibicao = bonus_sabado_forcado

    sabados_txt = ", ".join([pd.to_datetime(d).strftime("%d/%m") for d in todos_sabados_mes]) if todos_sabados_mes else "não identificado"
    sabados_trab_txt = ", ".join([pd.to_datetime(d).strftime("%d/%m") for d in sabados_trabalhados]) if sabados_trabalhados else "nenhum"
    sabados_faltantes_txt = ", ".join([pd.to_datetime(d).strftime("%d/%m") for d in sabados_faltantes]) if sabados_faltantes else "nenhum"
    feriados_txt = ", ".join([pd.to_datetime(d).strftime("%d/%m/%Y") for d in datas_feriado]) if datas_feriado else "nenhum informado"

    st.markdown("**Conferência do bônus de sábado**")
    conf1, conf2, conf3, conf4 = st.columns([1.4, 1.4, 1.1, 1.1])
    with conf1:
        st.caption("Sábados do mês")
        st.write(sabados_txt)
    with conf2:
        st.caption("Sábados com entrega no Excel")
        st.write(sabados_trab_txt)
    with conf3:
        st.caption("Entregas aos sábados")
        st.markdown(f"### {qtd_entregas_sabado}")
    with conf4:
        st.caption("Bônus calculado")
        st.markdown(f"### {moeda(bonus_sabado_exibicao)}")

    if quinzena_recibo == "2ª Quinzena":
        if veio_todos_sabados:
            bonus_sabados_recibo = bonus_sabado_calculado
            st.success(
                f"Bônus sábados aplicado: motorista teve entrega em todos os sábados do mês. "
                f"{qtd_entregas_sabado} entregas aos sábados × {moeda(valor_bonus_por_entrega)} = {moeda(bonus_sabados_recibo)}."
            )
        elif aplicar_bonus_sabado_sem_requisitos == "Sim":
            bonus_sabados_recibo = bonus_sabado_forcado
            st.success(
                f"Bônus sábados aplicado manualmente, mesmo sem todos os requisitos. "
                f"{qtd_entregas_sabado} entregas aos sábados × {moeda(valor_bonus_por_entrega)} = {moeda(bonus_sabados_recibo)}."
            )
        else:
            bonus_sabados_recibo = 0.0
            st.warning(
                f"Bônus sábados não aplicado. Para liberar, o motorista precisa ter entrega em todos os sábados do mês. "
                f"Sábados faltantes: {sabados_faltantes_txt}."
            )
    elif quinzena_recibo != "2ª Quinzena":
        st.info("Bônus de sábado foi conferido, mas só pode ser pago/aplicado na 2ª quinzena.")

    if quinzena_recibo == "2ª Quinzena":
        bonus_feriado_recibo, qtd_entregas_feriado = calcular_bonus_feriados_excel(
            df_bonus_excel,
            motorista_recibo,
            datas_feriado,
            valor_bonus_por_entrega,
        )

        if datas_feriado:
            st.info(
                f"Bônus feriado: datas {feriados_txt} | {qtd_entregas_feriado} entregas × "
                f"{moeda(valor_bonus_por_entrega)} = {moeda(bonus_feriado_recibo)}."
            )
    else:
        if datas_feriado_txt.strip():
            st.info("Bônus de feriado foi identificado apenas para pagamento na 2ª quinzena.")
        else:
            st.info("Bônus de feriados não entra na 1ª quinzena.")

    if rds_fechadas_recibo:
        rds_fechadas_set = {str(rd).strip().upper() for rd in rds_fechadas_recibo if str(rd).strip()}
        df_rd_bonus_bloqueado = df_pagamento.copy()
        df_rd_bonus_bloqueado["Data Rota DT"] = pd.to_datetime(df_rd_bonus_bloqueado["Data Rota"], errors="coerce").dt.date
        df_rd_bonus_bloqueado["Rota Ajuste"] = df_rd_bonus_bloqueado.get("Rota", pd.Series(dtype=str)).astype(str).str.strip().str.upper()
        df_rd_bonus_bloqueado = df_rd_bonus_bloqueado[
            (df_rd_bonus_bloqueado["Motorista Final"].astype(str) == str(motorista_recibo))
            & (df_rd_bonus_bloqueado["Data Rota DT"] >= data_inicio_recibo)
            & (df_rd_bonus_bloqueado["Data Rota DT"] <= data_fim_recibo)
            & (df_rd_bonus_bloqueado["Rota Ajuste"].isin(rds_fechadas_set))
        ].copy()

        if not df_rd_bonus_bloqueado.empty:
            if bonus_sabados_recibo > 0:
                qtd_entregas_sabado_rd_fechada = int(
                    (pd.to_datetime(df_rd_bonus_bloqueado["Data Rota"], errors="coerce").dt.weekday == 5).sum()
                )
                if qtd_entregas_sabado_rd_fechada > 0:
                    bonus_sabados_recibo = max(
                        0.0,
                        bonus_sabados_recibo - (qtd_entregas_sabado_rd_fechada * max(0.0, to_float(valor_bonus_por_entrega)))
                    )

            # Ajuste: o bônus de feriado informado/calculado deve continuar entrando no recibo,
            # no relatório e no Excel consolidado.
            # A RD Fechada altera somente o valor das RDs selecionadas para R$ 250,00;
            # as demais regras do recibo continuam sendo aplicadas normalmente no fechamento.

        st.info("RD Fechada aplicada somente nas RDs selecionadas. Essas RDs entram apenas com o valor fixo de R$ 250,00 e não recebem kg excedente, bônus de sábado ou bônus de feriado. As demais RDs seguem a regra normal de pagamento.")

    total_recibo = (
        subtotal_recibo
        + acareacao_recibo
        + bonus_extra_recibo
        + bonus_sabados_recibo
        + bonus_feriado_recibo
        - vale_recibo
        - desconto_recibo
    )

    # Atualiza o card superior "Total a pagar" conforme os ajustes do recibo.
    # Assim o total acompanha acareação, bônus, RD fechada, vale e desconto selecionados abaixo.
    try:
        total_a_pagar_placeholder.markdown(
            f'<div class="kpi-card"><div class="kpi-title">Total a pagar</div><div class="kpi-value">{moeda(total_recibo)}</div></div>',
            unsafe_allow_html=True,
        )
    except Exception:
        pass

    qtd_recibo = int(len(df_recibo))
    kg_recibo = float(df_recibo["KG Excedente"].sum())

    st.info(
        f"Recibo selecionado: **{motorista_recibo}** | "
        f"CNPJ: **{cnpj_motorista_recibo}** | "
        f"Quinzena: **{quinzena_recibo}** | "
        f"Entregas: **{qtd_recibo}** | "
        f"Kg excedente: **{kg_recibo:.2f}** | "
        f"Subtotal: **{moeda(subtotal_recibo)}** | "
        f"Acareação: **{moeda(acareacao_recibo)}** | "
        f"Bônus Extra: **{moeda(bonus_extra_recibo)}** | "
        f"Bônus Sábados: **{moeda(bonus_sabados_recibo)}** | "
        f"Bônus Feriado: **{moeda(bonus_feriado_recibo)}** | "
        f"Vale: **- {moeda(vale_recibo)}** | "
        f"Desconto: **- {moeda(desconto_recibo)}** | "
        f"Total: **{moeda(total_recibo)}**"
    )

    df_relatorio_entregas_excel = df_recibo.drop(columns=["Data Rota DT"], errors="ignore").copy()
    df_dia_excel = substituir_fechamento_diario_por_recibo(
        df_dia,
        df_relatorio_entregas_excel,
        motorista_recibo,
        data_inicio_recibo,
        data_fim_recibo,
    )
    motorista_relatorio_excel = motorista_recibo
    data_inicio_relatorio_excel = data_inicio_recibo
    data_fim_relatorio_excel = data_fim_recibo
    acareacao_relatorio_excel = acareacao_recibo
    vale_relatorio_excel = vale_recibo
    desconto_relatorio_excel = desconto_recibo
    bonus_extra_relatorio_excel = bonus_extra_recibo
    bonus_sabados_relatorio_excel = bonus_sabados_recibo
    bonus_feriado_relatorio_excel = bonus_feriado_recibo

    relatorio_pdf = gerar_relatorio_entregas_pdf(
        df_recibo.drop(columns=["Data Rota DT"], errors="ignore"),
        motorista_recibo,
        data_inicio_recibo,
        data_fim_recibo,
        quinzena_recibo,
        acareacao_recibo,
        vale_recibo,
        desconto_recibo,
        bonus_extra_recibo,
        bonus_sabados_recibo,
        bonus_feriado_recibo,
        cnpj_motorista=cnpj_motorista_recibo,
    )

    recibo_pdf = gerar_recibo_pdf(
        df_recibo.drop(columns=["Data Rota DT"], errors="ignore"),
        motorista_recibo,
        data_inicio_recibo,
        data_fim_recibo,
        quinzena_recibo,
        acareacao_recibo,
        vale_recibo,
        desconto_recibo,
        bonus_extra_recibo,
        bonus_sabados_recibo,
        bonus_feriado_recibo,
        cnpj_motorista=cnpj_motorista_recibo,
    )

    nome_motorista_arquivo = normalizar_nome_coluna(motorista_recibo) or "motorista"
    col_pdf_relatorio, col_pdf_recibo = st.columns(2)
    with col_pdf_relatorio:
        st.download_button(
            "📄 Gerar relatório de entregas",
            data=relatorio_pdf,
            file_name=f"relatorio_entregas_{nome_motorista_arquivo}_{normalizar_nome_coluna(quinzena_recibo)}_{data_inicio_recibo.strftime('%Y%m%d')}_{data_fim_recibo.strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with col_pdf_recibo:
        st.download_button(
            "🧾 Gerar recibo de pagamento",
            data=recibo_pdf,
            file_name=f"recibo_pagamento_{nome_motorista_arquivo}_{normalizar_nome_coluna(quinzena_recibo)}_{data_inicio_recibo.strftime('%Y%m%d')}_{data_fim_recibo.strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

st.markdown("---")

g1, g2 = st.columns(2)
with g1:
    fig = px.bar(df_dia_view, x="Data Rota", y="Total_Dia", color="Motorista Final", text="Total_Dia")
    fig.update_traces(texttemplate="R$ %{text:.2f}", textposition="outside")
    fig.update_layout(height=420, xaxis_title="", yaxis_title="Total dia")
    st.plotly_chart(fig, use_container_width=True)

with g2:
    por_motorista = df_dia_view.groupby("Motorista Final", as_index=False)["Total_Dia"].sum()
    fig2 = px.pie(por_motorista, names="Motorista Final", values="Total_Dia", hole=0.55)
    fig2.update_layout(height=420)
    st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")
st.markdown('<div class="section-heading">Entregas consideradas para pagamento</div>', unsafe_allow_html=True)

df_pagamento_display = df_pagamento.copy()
for col in ["Valor CEP", "Valor Excedente KG", "Total Entrega"]:
    if col in df_pagamento_display.columns:
        df_pagamento_display[col] = df_pagamento_display[col].apply(moeda)

cols_show = [
    "Motorista Final", "Data Rota", "Rota", "Pedido", "Status_Encontrados",
    "CEP", "CEP Prefixo", "Placa", "Tipo Veículo Final", "Tipo Veículo Assumido",
    "Fonte Peso Taxado", "KG Excedente", "Valor CEP", "Valor Excedente KG", "Total Entrega", "Arquivo PDF"
]
cols_show = [c for c in cols_show if c in df_pagamento_display.columns]

st.dataframe(df_pagamento_display[cols_show], use_container_width=True, height=520)


st.caption(
    "Regra ativa: quantidade de entregas baseada nas linhas válidas do PDF com status fechado no Excel. "
    "AWB repetida no manifesto não é removida, para bater com o controle manual."
)

st.markdown("---")
st.markdown('<div class="section-heading">Exportação</div>', unsafe_allow_html=True)

excel_bytes = criar_excel_fechamento(
    df_dia_excel,
    df_pagamento,
    df_pdf_info,
    df_relatorio_entregas_excel,
    motorista_relatorio_excel,
    data_inicio_relatorio_excel,
    data_fim_relatorio_excel,
    acareacao_relatorio_excel,
    vale_relatorio_excel,
    desconto_relatorio_excel,
    bonus_extra_relatorio_excel,
    bonus_sabados_relatorio_excel,
    bonus_feriado_relatorio_excel,
)

st.download_button(
    "📥 Baixar fechamento consolidado em Excel",
    data=excel_bytes,
    file_name=f"fechamento_entregadores_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.caption(
    "Regra aplicada: pagamento por entrega realizada sem ocorrência, validando Pedido x Status no Excel e dados da entrega no PDF. "
    "KG_Excedente_Calculado: somente o que passou de 10kg em cada entrega, calculado a R$ 0,30 por kg excedente."
)
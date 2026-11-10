"""Leitura e tratamento das planilhas de presença da Weekly Meeting.

Este módulo não tem interface: só lê as planilhas, padroniza as colunas e
classifica cada participante (presente, falta justificada, injustificada...).
É usado pelo app.py (Streamlit). As mesmas regras existem em assets/app.js,
que é a versão web publicada no GitHub Pages; se mudar uma regra aqui,
mude lá também.
"""
from __future__ import annotations

import re
import unicodedata
import warnings
from pathlib import Path
from typing import IO

import pandas as pd

# O openpyxl emite avisos sobre recursos do Excel que ele não lê (validação de dados,
# formatação condicional etc.). Eles não afetam a leitura, então são silenciados.
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Categorias de classificação de cada participante
PRESENTE = "Presente"
JUSTIFICADA = "Justificada"
JUSTIFICADA_OUTROS = "Justificada (Outros)"
INJUSTIFICADA = "Injustificada"
SEM_JUSTIFICATIVA = "Ausente sem justificativa base"
# Todas as categorias que contam como falta, na ordem em que aparecem nos gráficos.
CATEGORIAS_FALTA = [JUSTIFICADA, JUSTIFICADA_OUTROS, INJUSTIFICADA, SEM_JUSTIFICATIVA]

# Valor usado quando uma coluna de hierarquia/local está vazia ou não existe na planilha.
SEM_INFO = "(sem informação)"

# Nome canônico -> variações aceitas no cabeçalho (comparadas já normalizadas)
# Ex.: o cabeçalho "Líder Direto" vira "lider direto" após _norm() e é reconhecido como "lider".
# Para aceitar um novo nome de coluna, basta acrescentá-lo (em minúsculas e sem acento) na lista.
COLUNAS = {
    "nome": ["full name", "nome", "nome completo"],
    "email": ["e-mail", "email"],
    "branch": ["branch"],
    "office": ["office location"],
    "cargo": ["job title"],
    "lider": ["lider direto"],
    "supervisor": ["supervisor"],
    "gestor": ["gestor"],
    "gerente": ["gerente"],
    "report1": ["report1"],            # presença: "Yes" ou "Absence"
    "just_base": ["justificativas base", "justificativa base"],
    "report": ["report"],              # usado para separar "Outros" justificado de injustificado
    "outro": ["outro (detalhado)", "outros (detalhado)"],  # texto livre do motivo
}
# Sem estas colunas não é possível classificar as faltas; o arquivo é recusado com erro.
OBRIGATORIAS = ["nome", "report1", "just_base", "report"]


def _norm(texto) -> str:
    """minúsculas, sem acento e com espaços colapsados.

    Serve para comparar textos sem se preocupar com diferenças de digitação:
    "  Líder   Direto " e "lider direto" ficam iguais.
    """
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    # NFKD separa a letra do acento ("í" -> "i" + "´"); depois os acentos são descartados.
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().lower()


def _eh_outros(valor) -> bool:
    """True quando a Justificativa Base é a opção "Outros (detalhe ao lado)".

    Compara só o começo do texto, para aceitar pequenas variações na escrita da opção.
    """
    return _norm(valor).startswith("outros")


def semana_do_arquivo(nome_arquivo: str) -> str:
    """Extrai a semana do nome do arquivo: "WM43 - ATA 2026.xlsx" -> "WM43".

    O número sempre fica com 2 dígitos ("WM5" -> "WM05") para a ordenação funcionar.
    Se o nome não tiver "WM<número>", usa o próprio nome do arquivo sem a extensão.
    """
    m = re.search(r"WM\s*(\d{1,2})", nome_arquivo, re.IGNORECASE)
    return f"WM{int(m.group(1)):02d}" if m else Path(nome_arquivo).stem


def _achar_aba(xls: pd.ExcelFile) -> tuple[str, int]:
    """Retorna (aba, linha do cabeçalho) da aba que contém a coluna Report1.

    A planilha tem várias abas (MoM, Attendance, Report...) e o cabeçalho nem sempre
    está na primeira linha. Por isso o código procura, nas 15 primeiras linhas de
    cada aba, a célula "Report1" e usa aquela linha como cabeçalho.
    """
    for aba in xls.sheet_names:
        amostra = pd.read_excel(xls, sheet_name=aba, header=None, nrows=15)
        for i, linha in amostra.iterrows():
            if any(_norm(v) == "report1" for v in linha.values):
                return aba, i
    raise ValueError("nenhuma aba com a coluna 'Report1' foi encontrada")


def ler_arquivo(origem: str | Path | IO[bytes], nome_arquivo: str) -> pd.DataFrame:
    """Lê uma planilha e devolve uma tabela padronizada, com uma linha por participante.

    `origem` pode ser um caminho ou um arquivo em memória (upload do Streamlit).
    `nome_arquivo` é usado para descobrir a semana e para identificar a origem dos dados.
    """
    xls = pd.ExcelFile(origem)
    aba, linha_cab = _achar_aba(xls)
    bruto = pd.read_excel(xls, sheet_name=aba, header=linha_cab)

    # Troca os cabeçalhos originais pelos nomes canônicos (ver COLUNAS).
    # Se duas colunas corresponderem ao mesmo nome canônico, vale a primeira.
    renomear = {}
    for col in bruto.columns:
        n = _norm(col)
        for canonico, variacoes in COLUNAS.items():
            if n in variacoes and canonico not in renomear.values():
                renomear[col] = canonico
    # Mantém só as colunas reconhecidas; as demais da planilha são descartadas.
    df = bruto.rename(columns=renomear)[list(renomear.values())]

    faltando = [c for c in OBRIGATORIAS if c not in df.columns]
    if faltando:
        raise ValueError(f"colunas obrigatórias ausentes em '{aba}': {faltando}")

    df = df[df["nome"].notna()].copy()  # descarta linhas sem nome (linhas em branco, totais etc.)
    # Colunas opcionais que não existem na planilha são criadas vazias,
    # para o resto do código não precisar checar se elas existem.
    for c in COLUNAS:
        if c not in df.columns:
            df[c] = pd.NA
    df["arquivo"] = nome_arquivo
    df["semana"] = semana_do_arquivo(nome_arquivo)
    df["classificacao"] = df.apply(classificar, axis=1)
    # Remove espaços duplicados/sobrando para que textos iguais sejam contados juntos.
    df["just_base"] = df["just_base"].map(lambda v: re.sub(r"\s+", " ", str(v)).strip() if pd.notna(v) else v)
    df["outro"] = df["outro"].map(lambda v: re.sub(r"\s+", " ", str(v)).strip() if pd.notna(v) else v)
    df["outro_agrupado"] = df["outro"].map(agrupar_detalhe)
    # Colunas de hierarquia e local: célula vazia vira "(sem informação)", que aparece como um grupo próprio.
    for c in ["lider", "supervisor", "gestor", "gerente", "branch", "office"]:
        df[c] = df[c].map(lambda v: str(v).strip() if pd.notna(v) and str(v).strip() else SEM_INFO)
    return df


def classificar(linha) -> str:
    """Regras:
    - Report1 = Yes -> Presente
    - Report1 = Absence e Justificativa Base preenchida e != Outros -> Justificada
    - Report1 = Absence, Justificativa Base = Outros e Report = Absence -> Injustificada
    - Report1 = Absence, Justificativa Base = Outros e Report != Absence -> Justificada (Outros)
    - Report1 = Absence e Justificativa Base vazia -> Ausente sem justificativa base
    """
    if _norm(linha["report1"]) != "absence":
        return PRESENTE
    if _eh_outros(linha["just_base"]):
        return INJUSTIFICADA if _norm(linha["report"]) == "absence" else JUSTIFICADA_OUTROS
    if _norm(linha["just_base"]):
        return JUSTIFICADA
    return SEM_JUSTIFICATIVA


def agrupar_detalhe(valor):
    """Agrupa variações de escrita ('DAY OFF', 'Day-off', 'Day off') num só rótulo.

    Converte para minúsculas, troca hífen por espaço e deixa só a primeira letra maiúscula.
    Os acentos são mantidos, para o texto continuar legível nos gráficos.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)) or not str(valor).strip():
        return pd.NA
    chave = str(valor).lower().replace("-", " ")
    chave = re.sub(r"\s+", " ", chave).strip()
    return chave[:1].upper() + chave[1:]


def carregar(arquivos: list[tuple[str, str | Path | IO[bytes]]]) -> tuple[pd.DataFrame, list[str]]:
    """arquivos: lista de (nome, caminho_ou_buffer). Retorna (dados, erros).

    Lê todas as planilhas e junta tudo numa tabela só. Um arquivo com problema
    não impede os outros de serem carregados: o erro vai para a lista `erros`,
    que o app mostra na barra lateral.
    """
    partes, erros = [], []
    for nome, origem in arquivos:
        try:
            partes.append(ler_arquivo(origem, nome))
        except Exception as e:  # noqa: BLE001 - erro mostrado ao usuário
            erros.append(f"{nome}: {e}")
    if not partes:
        return pd.DataFrame(), erros
    return pd.concat(partes, ignore_index=True), erros


def faltas_por(df: pd.DataFrame, coluna: str) -> pd.DataFrame:
    """Tabela pessoa-ausente por nível hierárquico x classificação.

    Para cada valor de `coluna` (ex.: cada líder direto), conta as faltas de cada tipo
    e calcula:
      - "Total de faltas": soma de todas as faltas;
      - "Equipe": quantas pessoas daquele líder estavam na lista (presentes + ausentes);
      - "% faltas": Total de faltas / Equipe.
    Ordena do maior para o menor número de faltas.
    """
    aus = df[df["classificacao"] != PRESENTE]
    tab = pd.crosstab(aus[coluna], aus["classificacao"])  # linhas = pessoas, colunas = tipos de falta
    # Garante que todas as categorias existam como coluna, mesmo com zero faltas.
    for c in CATEGORIAS_FALTA:
        if c not in tab.columns:
            tab[c] = 0
    # A coluna "sem justificativa base" só aparece se houver algum caso.
    tab = tab[[c for c in CATEGORIAS_FALTA if c != SEM_JUSTIFICATIVA or tab[c].sum() > 0]]
    tab["Total de faltas"] = tab.sum(axis=1)
    equipe = df.groupby(coluna).size()  # tamanho da equipe usa todos, não só os ausentes
    tab["Equipe"] = equipe.reindex(tab.index).fillna(0).astype(int)
    tab["% faltas"] = (tab["Total de faltas"] / tab["Equipe"] * 100).round(1)
    return tab.sort_values("Total de faltas", ascending=False)

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
    "report": ["report"],              # modelo atual: separa "Outros" justificado de injustificado; antigo: presença e motivo
    "outro": ["outro (detalhado)", "outros (detalhado)"],  # texto livre do motivo
    "obs": ["injustificada"],          # layout antigo: comentário sobre a falta injustificada
}

# Os dois modelos de planilha aceitos (o usuário escolhe na barra lateral do app):
#   atual:  colunas Report1 (presença) + Justificativa Base + Report + Outro (detalhado)
#   antigo: só a coluna Report, com "Yes", "Absence" ou o motivo da falta
#           ("Justified", "Vacation", "Medical Absence"...) + coluna opcional "Injustificada"
#   obrig: colunas sem as quais a planilha é recusada
#   cats:  tipos de falta que o modelo consegue identificar
MODELOS = {
    "atual": {"rotulo": "Modelo atual", "obrig": ["nome", "report1", "just_base", "report"],
              "cats": CATEGORIAS_FALTA},
    "antigo": {"rotulo": "Layout antigo", "obrig": ["nome", "report"], "cats": [JUSTIFICADA, INJUSTIFICADA]},
}
# Nome das colunas obrigatórias como aparecem na planilha (usado nas mensagens de erro).
NOME_COLUNA = {"nome": "Full Name", "report1": "Report1", "just_base": "Justificativa Base", "report": "Report"}
# Mensagem para arquivo que o Windows não deixa ler (normalmente porque está aberto no Excel).
MSG_ARQUIVO_TRAVADO = ("Não foi possível ler o arquivo. Ele provavelmente está aberto no Excel ou em outro programa: "
                       "feche-o e tente novamente.")


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


def semana_do_arquivo(nome_arquivo: str) -> str | None:
    """Extrai a semana do nome do arquivo: "WM43 - ATA 2026.xlsx" -> "WM43".

    O número sempre fica com 2 dígitos ("WM5" -> "WM05") para a ordenação funcionar.
    Devolve None se o nome não tiver "WM" seguido de um número de semana válido (1 a 53).
    """
    m = re.search(r"WM\s*(\d{1,2})(?!\d)", nome_arquivo, re.IGNORECASE)
    if not m or not 1 <= int(m.group(1)) <= 53:
        return None
    return f"WM{int(m.group(1)):02d}"


class ModeloDiferente(ValueError):
    """Planilha carregada no modelo errado. `certo` diz em qual modelo ela deve ser aberta."""

    def __init__(self, certo: str):
        if certo == "antigo":
            msg = ('Esta planilha está no layout antigo (só com a coluna "Report", sem "Report1" e "Justificativa Base"). '
                   'Selecione "Layout antigo" em "Modelo de planilha", na barra lateral.')
        else:
            msg = ('Esta planilha está no modelo atual (com as colunas "Report1" e "Justificativa Base"). '
                   'Selecione "Modelo atual" em "Modelo de planilha", na barra lateral.')
        super().__init__(msg)
        self.certo = certo


def _modelo_do_cabecalho(valores) -> str | None:
    """Diz se uma linha é cabeçalho do modelo atual ("Report1") ou do layout antigo
    ("Report" junto com "Full Name" ou "Líder Direto", sem "Report1").

    A exigência do nome/líder evita confundir com a aba "Report" exportada do Teams.
    """
    cel = {_norm(v) for v in valores}
    if "report1" in cel:
        return "atual"
    if "report" in cel and ({"full name", "lider direto"} & cel):
        return "antigo"
    return None


def _achar_aba(xls: pd.ExcelFile, modelo: str) -> tuple[str, int]:
    """Retorna (aba, linha do cabeçalho) da primeira aba do modelo escolhido.

    A planilha tem várias abas (MoM, Attendance, Report...) e o cabeçalho nem sempre
    está na primeira linha. Por isso o código procura o cabeçalho nas 15 primeiras
    linhas de cada aba. Se só houver abas do outro modelo, gera ModeloDiferente.
    """
    outro = None  # modelo encontrado em alguma aba, quando diferente do escolhido
    for aba in xls.sheet_names:
        amostra = pd.read_excel(xls, sheet_name=aba, header=None, nrows=15)
        for i, linha in amostra.iterrows():
            achado = _modelo_do_cabecalho(linha.values)
            if achado == modelo:
                return aba, i
            if achado:
                outro = outro or achado
                break  # esta aba é do outro modelo; passa para a próxima
    if outro:
        raise ModeloDiferente(outro)
    if modelo == "antigo":
        raise ValueError('Nenhuma aba da planilha tem as colunas "Full Name" e "Report". '
                         'Confira se este é o arquivo de presença da Weekly Meeting (com a aba "Attendance").')
    raise ValueError('Nenhuma aba da planilha tem a coluna "Report1". '
                     'Confira se este é o arquivo de presença da Weekly Meeting (com a aba "Attendance").')


def ler_arquivo(origem: str | Path | IO[bytes], nome_arquivo: str, modelo: str = "atual") -> pd.DataFrame:
    """Lê uma planilha e devolve uma tabela padronizada, com uma linha por participante.

    `origem` pode ser um caminho ou um arquivo em memória (upload do Streamlit).
    `nome_arquivo` é usado para descobrir a semana e para identificar a origem dos dados.
    `modelo` é "atual" ou "antigo" (ver MODELOS).
    Qualquer problema gera um ValueError com uma mensagem explicando o que corrigir.
    """
    # 1. O nome do arquivo precisa indicar a semana.
    semana = semana_do_arquivo(nome_arquivo)
    if semana is None:
        raise ValueError('O nome do arquivo precisa indicar a semana: "WM" seguido do número da semana (1 a 53), '
                         'por exemplo "WM43 - ATA 2026.xlsx". Renomeie o arquivo e carregue novamente.')

    # 2. O arquivo precisa abrir como planilha (falha se estiver travado, corrompido ou protegido por senha).
    try:
        xls = pd.ExcelFile(origem)
    except PermissionError as e:
        raise ValueError(MSG_ARQUIVO_TRAVADO) from e
    except Exception as e:  # noqa: BLE001 - qualquer outra falha de abertura vira a mesma mensagem
        raise ValueError("Não foi possível abrir o arquivo. Ele pode estar corrompido, protegido por senha "
                         "ou não ser uma planilha válida. Tente abri-lo no Excel e salvá-lo novamente como .xlsx.") from e

    # 3. Procura a aba com os dados de presença.
    aba, linha_cab = _achar_aba(xls, modelo)
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

    # 4. As colunas obrigatórias precisam existir.
    faltando = [c for c in MODELOS[modelo]["obrig"] if c not in df.columns]
    if faltando:
        rotulo = "as colunas obrigatórias" if len(faltando) > 1 else "a coluna obrigatória"
        nomes = ", ".join(f'"{NOME_COLUNA[c]}"' for c in faltando)
        raise ValueError(f'A aba "{aba}" não tem {rotulo} {nomes}. Confira se os títulos das colunas não foram alterados.')

    df = df[df["nome"].notna()].copy()  # descarta linhas sem nome (linhas em branco, totais etc.)
    # 5. Precisa haver ao menos um participante.
    if df.empty:
        raise ValueError(f'A aba "{aba}" não tem nenhum participante preenchido abaixo do cabeçalho.')
    # Colunas opcionais que não existem na planilha são criadas vazias,
    # para o resto do código não precisar checar se elas existem.
    for c in COLUNAS:
        if c not in df.columns:
            df[c] = pd.NA
    df["arquivo"] = nome_arquivo
    df["semana"] = semana
    df["classificacao"] = df.apply(classificar, axis=1, modelo=modelo)
    # Remove espaços duplicados/sobrando para que textos iguais sejam contados juntos.
    for c in ["just_base", "report", "outro", "obs"]:
        df[c] = df[c].map(lambda v: re.sub(r"\s+", " ", str(v)).strip() if pd.notna(v) else v)
    # Motivo da falta justificada usado no gráfico "Por motivo": a Justificativa Base no
    # modelo atual; o próprio valor de Report ("Vacation", "Justified"...) no layout antigo.
    df["motivo"] = df["report"] if modelo == "antigo" else df["just_base"]
    df["outro_agrupado"] = df["outro"].map(agrupar_detalhe)
    # Colunas de hierarquia e local: célula vazia vira "(sem informação)", que aparece como um grupo próprio.
    for c in ["lider", "supervisor", "gestor", "gerente", "branch", "office"]:
        df[c] = df[c].map(lambda v: str(v).strip() if pd.notna(v) and str(v).strip() else SEM_INFO)
    return df


def classificar(linha, modelo: str = "atual") -> str:
    """Regras do modelo atual:
    - Report1 = Yes -> Presente
    - Report1 = Absence e Justificativa Base preenchida e != Outros -> Justificada
    - Report1 = Absence, Justificativa Base = Outros e Report = Absence -> Injustificada
    - Report1 = Absence, Justificativa Base = Outros e Report != Absence -> Justificada (Outros)
    - Report1 = Absence e Justificativa Base vazia -> Ausente sem justificativa base

    Regras do layout antigo (só a coluna Report):
    - Report = Yes (ou vazio) -> Presente
    - Report = Absence -> Injustificada
    - Report = qualquer outro motivo (Justified, Vacation, Medical Absence...) -> Justificada
    """
    if modelo == "antigo":
        valor = _norm(linha["report"])
        if not valor or valor == "yes":
            return PRESENTE
        return INJUSTIFICADA if valor == "absence" else JUSTIFICADA
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


def carregar(arquivos: list[tuple[str, str | Path | IO[bytes]]],
             modelo: str = "atual") -> tuple[pd.DataFrame, list[str], list[str]]:
    """arquivos: lista de (nome, caminho_ou_buffer). Retorna (dados, erros, outro_modelo).

    Lê todas as planilhas no `modelo` escolhido e junta tudo numa tabela só. Um arquivo
    com problema não impede os outros de serem carregados: o erro vai para a lista `erros`
    (no formato "arquivo: motivo"), que o app mostra no topo da página.
    `outro_modelo` lista os arquivos recusados por serem do outro modelo (eles também
    estão em `erros`); o app usa isso para não reclamar dos arquivos da pasta do projeto.
    Dois arquivos diferentes da mesma semana não são aceitos (as faltas seriam contadas em dobro).
    """
    partes, erros, outro_modelo = [], [], []
    semanas_vistas: dict[str, str] = {}  # semana -> arquivo que a carregou
    for nome, origem in arquivos:
        try:
            if Path(nome).name.startswith("~$"):
                raise ValueError("Este é um arquivo temporário que o Excel cria enquanto a planilha está aberta. "
                                 'Carregue o arquivo original (sem "~$" no início do nome).')
            df = ler_arquivo(origem, nome, modelo)
            semana = df["semana"].iat[0]
            if semana in semanas_vistas:
                raise ValueError(f'A semana {semana} já foi carregada pelo arquivo "{semanas_vistas[semana]}". '
                                 "Deixe apenas um arquivo por semana.")
            semanas_vistas[semana] = nome
            partes.append(df)
        except Exception as e:  # noqa: BLE001 - erro mostrado ao usuário
            erros.append(f"{nome}: {e}")
            if isinstance(e, ModeloDiferente):
                outro_modelo.append(nome)
    if not partes:
        return pd.DataFrame(), erros, outro_modelo
    return pd.concat(partes, ignore_index=True), erros, outro_modelo


def faltas_por(df: pd.DataFrame, coluna: str, categorias: list[str] = CATEGORIAS_FALTA) -> pd.DataFrame:
    """Tabela pessoa-ausente por nível hierárquico x classificação.

    Para cada valor de `coluna` (ex.: cada líder direto), conta as faltas de cada tipo
    e calcula:
      - "Total de faltas": soma de todas as faltas;
      - "Equipe": quantas pessoas daquele líder estavam na lista (presentes + ausentes);
      - "% faltas": Total de faltas / Equipe.
    `categorias` são os tipos de falta do modelo de planilha (ver MODELOS).
    Ordena do maior para o menor número de faltas.
    """
    aus = df[df["classificacao"] != PRESENTE]
    tab = pd.crosstab(aus[coluna], aus["classificacao"])  # linhas = pessoas, colunas = tipos de falta
    # Garante que todas as categorias existam como coluna, mesmo com zero faltas.
    for c in categorias:
        if c not in tab.columns:
            tab[c] = 0
    # A coluna "sem justificativa base" só aparece se houver algum caso.
    tab = tab[[c for c in categorias if c != SEM_JUSTIFICATIVA or tab[c].sum() > 0]]
    tab["Total de faltas"] = tab.sum(axis=1)
    equipe = df.groupby(coluna).size()  # tamanho da equipe usa todos, não só os ausentes
    tab["Equipe"] = equipe.reindex(tab.index).fillna(0).astype(int)
    tab["% faltas"] = (tab["Total de faltas"] / tab["Equipe"] * 100).round(1)
    return tab.sort_values("Total de faltas", ascending=False)

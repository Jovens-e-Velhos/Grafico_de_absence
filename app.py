"""Painel de acompanhamento de faltas da Weekly Meeting (versão local em Python/Streamlit).

Rodar:  streamlit run app.py   (ou dar duplo clique em iniciar.bat)

Fluxo geral do arquivo:
  1. Carrega as planilhas (enviadas pela barra lateral e/ou .xlsx da pasta do projeto).
  2. Deixa o usuário escolher qual coluna representa cada nível hierárquico.
  3. Aplica os filtros (Semana, Branch, Office).
  4. Mostra KPIs, faltas por hierarquia, justificadas, injustificadas e evolução semanal.

As regras de leitura e classificação das faltas ficam em processamento.py;
aqui fica só a interface (filtros, gráficos e tabelas).
"""
import sys
from pathlib import Path

# Garante que o processamento.py (que está na mesma pasta) seja encontrado no import,
# mesmo quando o Streamlit é iniciado a partir de outra pasta.
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import processamento as P

# Pasta onde o app procura planilhas .xlsx automaticamente (a própria pasta do projeto).
PASTA_DADOS = Path(__file__).parent

# Cor de cada classificação de falta nos gráficos.
# Cada categoria tem sempre a mesma cor, independente do filtro aplicado.
CORES = {  # paleta categórica validada (slots 1-3) + neutro
    P.JUSTIFICADA: "#2a78d6",
    P.JUSTIFICADA_OUTROS: "#1baf7a",
    P.INJUSTIFICADA: "#eb6834",
    P.SEM_JUSTIFICATIVA: "#898781",
}
COR_UNICA = "#2a78d6"  # cor dos gráficos de série única (barras simples)

# Níveis hierárquicos exibidos no painel: chave interna -> nome mostrado na tela.
NIVEIS = {"lider": "Líder Direto", "supervisor": "Supervisor", "gerente": "Gerente"}

st.set_page_config(page_title="Faltas Weekly Meeting", page_icon="📊", layout="wide")


# ---------------------------------------------------------------- dados
@st.cache_data(show_spinner="Lendo planilhas…")
def carregar_cache(itens: tuple[tuple[str, bytes], ...]):
    """Lê as planilhas e devolve (dados, erros).

    O resultado fica em cache: enquanto os mesmos arquivos estiverem carregados,
    mexer nos filtros não faz o app reler as planilhas do zero.
    Recebe uma tupla de (nome do arquivo, conteúdo em bytes) porque o cache
    precisa de argumentos imutáveis para comparar.
    """
    import io
    return P.carregar([(nome, io.BytesIO(conteudo)) for nome, conteudo in itens])


# Barra lateral: envio de arquivos.
with st.sidebar:
    st.header("Arquivos")
    enviados = st.file_uploader("Envie um ou mais .xlsx", type="xlsx", accept_multiple_files=True)
    # Por padrão, usa os .xlsx da pasta do projeto só quando nada foi enviado.
    usar_pasta = st.checkbox("Incluir .xlsx da pasta do projeto", value=not enviados)

# Monta a lista de arquivos a processar: primeiro os enviados, depois os da pasta.
itens = []
if enviados:
    itens += [(f.name, f.getvalue()) for f in enviados]
if usar_pasta:
    nomes_enviados = {n for n, _ in itens}
    for p in sorted(PASTA_DADOS.glob("*.xlsx")):
        # Ignora arquivos temporários do Excel (~$...) e os que já vieram pelo upload,
        # para não contar a mesma semana duas vezes.
        if not p.name.startswith("~$") and p.name not in nomes_enviados:
            itens.append((p.name, p.read_bytes()))

if not itens:
    st.info("Envie ao menos um arquivo .xlsx na barra lateral.")
    st.stop()  # interrompe a execução: não há nada para mostrar

dados, erros = carregar_cache(tuple(itens))
for e in erros:  # ex.: planilha sem a aba esperada ou sem colunas obrigatórias
    st.sidebar.error(e)
if dados.empty:
    st.stop()

# ---------------------------------------------------------------- hierarquia
# Cada nível (Líder Direto, Supervisor, Gerente) pode ser ligado a qualquer coluna de
# hierarquia que exista na planilha. Por padrão, cada nível usa a coluna de mesmo nome.
with st.sidebar:
    st.header("Hierarquia")
    st.caption("Escolha qual coluna da planilha representa cada nível.")
    # Colunas de hierarquia que o processamento.py reconhece: chave interna -> nome exibido.
    opcoes = {"lider": "Líder Direto", "supervisor": "Supervisor", "gestor": "Gestor", "gerente": "Gerente"}
    # Só oferece as colunas que têm algum valor preenchido nas planilhas carregadas.
    disponiveis = [c for c in opcoes if (dados[c] != P.SEM_INFO).any()]

    def _seletor(nivel, padrao):
        """Caixa de seleção da coluna usada por um nível.

        Já vem com a coluna `padrao` selecionada, se ela existir; senão, vem "(não usar)",
        e o nível não aparece no painel.
        """
        lista = disponiveis + ["(não usar)"]
        idx = lista.index(padrao) if padrao in lista else len(lista) - 1
        return st.selectbox(NIVEIS[nivel], lista, index=idx, format_func=lambda c: opcoes.get(c, c), key=f"niv_{nivel}")

    # Resultado: nível -> coluna escolhida (ou "(não usar)").
    mapa_niveis = {
        "lider": _seletor("lider", "lider"),
        "supervisor": _seletor("supervisor", "supervisor"),
        "gerente": _seletor("gerente", "gerente"),
    }

# ---------------------------------------------------------------- filtros
# Cada arquivo corresponde a uma semana (ex.: "WM43", lido do nome do arquivo).
semanas = sorted(dados["semana"].unique())
c1, c2, c3 = st.columns(3)
sel_semanas = c1.multiselect("Semana", semanas, default=semanas[-1:])  # padrão: semana mais recente
sel_branch = c2.multiselect("Branch", sorted(dados["branch"].unique()))
sel_office = c3.multiselect("Office Location", sorted(dados["office"].unique()))

# Aplica os filtros. Filtro vazio = sem restrição (todas as semanas/branches/offices).
df = dados[dados["semana"].isin(sel_semanas or semanas)]
if sel_branch:
    df = df[df["branch"].isin(sel_branch)]
if sel_office:
    df = df[df["office"].isin(sel_office)]

# Todos os que não estavam presentes, seja qual for o tipo de falta.
ausentes = df[df["classificacao"] != P.PRESENTE]

st.title("Acompanhamento de faltas — Weekly Meeting")
st.caption(f"Arquivos: {', '.join(sorted(df['arquivo'].unique()))}")

# ---------------------------------------------------------------- KPIs
# Números-resumo do recorte filtrado.
n_total, n_aus = len(df), len(ausentes)
cont = ausentes["classificacao"].value_counts()  # quantidade de faltas por classificação
k = st.columns(5)
k[0].metric("Participantes esperados", n_total)
k[1].metric("Presentes", n_total - n_aus)
k[2].metric("Faltas", n_aus, f"{n_aus / n_total:.1%} do total" if n_total else None, delta_color="off")
# "Justificadas" soma as justificadas normais e as justificadas via "Outros".
k[3].metric("Justificadas", int(cont.get(P.JUSTIFICADA, 0) + cont.get(P.JUSTIFICADA_OUTROS, 0)),
            f"{int(cont.get(P.JUSTIFICADA_OUTROS, 0))} via 'Outros'", delta_color="off")
k[4].metric("Injustificadas", int(cont.get(P.INJUSTIFICADA, 0)))


# ---------------------------------------------------------------- helpers de gráfico
def _estilo(fig: go.Figure, altura: int) -> go.Figure:
    """Aplica o visual padrão a todos os gráficos: margens, legenda no topo,
    grade discreta, eixo começando em zero e números inteiros no eixo."""
    fig.update_layout(
        height=altura, margin=dict(l=10, r=30, t=10, b=10), bargap=0.25,
        legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0, title=None),
        xaxis=dict(title=None, showgrid=True, gridwidth=1, zeroline=False, tickformat="d", rangemode="tozero"),
        yaxis=dict(title=None, showgrid=False),
        hoverlabel=dict(font_size=13),
    )
    fig.update_traces(marker_line_width=0)
    return fig


def barras_empilhadas(tab: pd.DataFrame, rotulo: str) -> go.Figure:
    """Barras horizontais empilhadas: uma barra por pessoa da hierarquia (ex.: cada líder),
    dividida pelo tipo de falta. O total aparece escrito no fim de cada barra.

    `tab` vem de P.faltas_por(): uma linha por pessoa, uma coluna por classificação,
    mais "Total de faltas", "Equipe" e "% faltas".
    """
    tab = tab.sort_values("Total de faltas")  # crescente: o plotly desenha de baixo para cima, então o maior fica no topo
    fig = go.Figure()
    # Só desenha as categorias que têm ao menos uma falta (evita itens vazios na legenda).
    cats = [c for c in P.CATEGORIAS_FALTA if c in tab.columns and tab[c].sum() > 0]
    for i, c in enumerate(cats):
        fig.add_bar(
            y=tab.index, x=tab[c], name=c, orientation="h", marker_color=CORES[c],
            # Arredonda só a ponta do último segmento, para a barra empilhada parecer uma peça única.
            marker_cornerradius=4 if i == len(cats) - 1 else 0,
            # Dados extras usados no texto que aparece ao passar o mouse.
            customdata=tab[["Total de faltas", "Equipe", "% faltas"]].values,
            hovertemplate=f"<b>%{{y}}</b><br>{c}: %{{x}}<br>Total: %{{customdata[0]}} de %{{customdata[1]}} "
                          f"(%{{customdata[2]}}%)<extra></extra>",
        )
    fig.add_scatter(  # total ao fim da barra (rótulo direto)
        y=tab.index, x=tab["Total de faltas"], mode="text", text=tab["Total de faltas"],
        textposition="middle right", showlegend=False, hoverinfo="skip", textfont=dict(size=12),
    )
    fig.update_layout(barmode="stack")
    # A altura cresce com o número de barras, para os nomes não ficarem espremidos.
    return _estilo(fig, max(260, 28 * len(tab) + 80))


def barras_simples(serie: pd.Series) -> go.Figure:
    """Barras horizontais de uma única cor, com o valor escrito ao lado.
    Usado nas contagens por Justificativa Base e por motivo de 'Outros'."""
    serie = serie.sort_values()
    fig = px.bar(x=serie.values, y=serie.index, orientation="h", text=serie.values)
    fig.update_traces(marker_color=COR_UNICA, marker_cornerradius=4, textposition="outside",
                      hovertemplate="<b>%{y}</b><br>%{x} pessoa(s)<extra></extra>")
    return _estilo(fig, max(220, 34 * len(serie) + 60))


def tabela(d: pd.DataFrame, nome_csv: str):
    """Mostra uma tabela e um botão para baixá-la em CSV.

    O CSV usa ';' como separador e UTF-8 com BOM, para abrir certinho
    (com acentos e colunas separadas) no Excel em português.
    """
    st.dataframe(d, use_container_width=True, hide_index=True)
    st.download_button("Baixar CSV", d.to_csv(index=False, sep=";").encode("utf-8-sig"),
                       file_name=nome_csv, mime="text/csv", key=f"dl_{nome_csv}")


# Colunas exibidas nas tabelas de pessoas: nome interno -> título na tabela.
COLS_PESSOA = {"semana": "Semana", "nome": "Nome", "lider": "Líder Direto", "supervisor": "Supervisor",
               "gestor": "Gestor", "gerente": "Gerente", "branch": "Branch", "office": "Office",
               "just_base": "Justificativa Base", "report": "Report", "outro": "Outro (detalhado)",
               "classificacao": "Classificação"}


def pessoas(d: pd.DataFrame) -> pd.DataFrame:
    """Prepara uma lista de colaboradores para exibição.

    As colunas Supervisor, Gestor e Gerente só aparecem se tiverem algum valor
    preenchido (evita colunas inteiras com "(sem informação)").
    """
    cols = [c for c in COLS_PESSOA if c not in ("gestor", "gerente", "supervisor") or (d[c] != P.SEM_INFO).any()]
    return d[cols].rename(columns=COLS_PESSOA)


# ---------------------------------------------------------------- 1. faltas por hierarquia
st.subheader("Faltas por nível hierárquico")
# Descarta os níveis marcados como "(não usar)" na barra lateral.
niveis_ativos = {n: c for n, c in mapa_niveis.items() if c != "(não usar)"}
if n_aus == 0:
    st.success("Nenhuma falta no recorte selecionado.")
elif niveis_ativos:
    # Uma aba por nível. Se o nível estiver ligado a uma coluna de outro nome,
    # o nome da coluna aparece entre parênteses, ex.: "Gerente (Supervisor)".
    abas = st.tabs([f"{NIVEIS[n]} ({opcoes[c]})" if opcoes[c] != NIVEIS[n] else NIVEIS[n] for n, c in niveis_ativos.items()])
    for aba, (nivel, coluna) in zip(abas, niveis_ativos.items()):
        with aba:
            tab = P.faltas_por(df, coluna)  # faltas agrupadas pela coluna escolhida
            st.plotly_chart(barras_empilhadas(tab, NIVEIS[nivel]), use_container_width=True, key=f"g_{nivel}")
            with st.expander("Ver tabela"):
                tabela(tab.reset_index().rename(columns={coluna: NIVEIS[nivel]}), f"faltas_por_{nivel}.csv")

# ---------------------------------------------------------------- 2. justificadas
st.subheader("Faltas justificadas")
col_a, col_b = st.columns(2)
with col_a:
    # Quantas vezes cada opção de "Justificativa Base" foi escolhida (sem contar "Outros").
    st.markdown("**Por Justificativa Base** (exceto 'Outros')")
    just = ausentes[ausentes["classificacao"] == P.JUSTIFICADA]["just_base"].value_counts()
    if just.empty:
        st.caption("Nenhuma.")
    else:
        st.plotly_chart(barras_simples(just), use_container_width=True, key="g_just")
with col_b:
    # "Outros" com Report diferente de "Absence" conta como justificada; aqui aparece
    # o motivo escrito em "Outro (detalhado)". "outro_agrupado" é esse texto em minúsculas,
    # para que "Férias" e "férias" sejam contados juntos.
    st.markdown("**Justificadas via 'Outros'** — motivo informado em 'Outro (detalhado)'")
    outros_j = ausentes[ausentes["classificacao"] == P.JUSTIFICADA_OUTROS]["outro_agrupado"].value_counts()
    if outros_j.empty:
        st.caption("Nenhuma.")
    else:
        st.plotly_chart(barras_simples(outros_j), use_container_width=True, key="g_outros_j")
with st.expander("Ver colaboradores com falta justificada"):
    tabela(pessoas(ausentes[ausentes["classificacao"].isin([P.JUSTIFICADA, P.JUSTIFICADA_OUTROS])]),
           "faltas_justificadas.csv")

# ---------------------------------------------------------------- 3. injustificadas
st.subheader("Faltas injustificadas")
st.caption("Justificativa Base = 'Outros (detalhe ao lado)' e Report = 'Absence'.")
inj = ausentes[ausentes["classificacao"] == P.INJUSTIFICADA]
if inj.empty:
    st.success("Nenhuma falta injustificada.")
else:
    # Tabela enxuta: quem faltou, a quem responde e o que foi escrito em "Outro (detalhado)".
    # Supervisor/Gestor/Gerente só entram se existirem nos dados.
    tabela(pessoas(inj)[["Semana", "Nome", "Líder Direto"] +
                        [c for c in ["Supervisor", "Gestor", "Gerente"] if c in pessoas(inj).columns] +
                        ["Branch", "Outro (detalhado)"]], "faltas_injustificadas.csv")

# Ausências com "Justificativa Base" em branco: não dá para classificar, então são listadas à parte.
sem_base = ausentes[ausentes["classificacao"] == P.SEM_JUSTIFICATIVA]
if not sem_base.empty:
    st.warning(f"{len(sem_base)} ausência(s) sem 'Justificativa Base' preenchida.")
    tabela(pessoas(sem_base), "ausentes_sem_justificativa.csv")

# ---------------------------------------------------------------- 4. evolução semanal
# Só faz sentido com mais de uma semana carregada.
if dados["semana"].nunique() > 1:
    st.subheader("Evolução semanal")
    # Usa todas as semanas (ignora o filtro de Semana), mas respeita Branch e Office.
    base = dados.copy()
    if sel_branch:
        base = base[base["branch"].isin(sel_branch)]
    if sel_office:
        base = base[base["office"].isin(sel_office)]
    # Conta as faltas por semana e por classificação -> uma linha no gráfico para cada classificação.
    evo = (base[base["classificacao"] != P.PRESENTE]
           .groupby(["semana", "classificacao"]).size().rename("faltas").reset_index())
    fig = px.line(evo, x="semana", y="faltas", color="classificacao", markers=True,
                  color_discrete_map=CORES, category_orders={"classificacao": P.CATEGORIAS_FALTA})
    fig.update_traces(line_width=2, marker_size=8,
                      hovertemplate="%{x}<br>%{fullData.name}: %{y}<extra></extra>")
    fig = _estilo(fig, 320)
    # Ao passar o mouse sobre uma semana, mostra todas as classificações juntas.
    fig.update_layout(hovermode="x unified", xaxis=dict(dtick=None, showgrid=False),
                      yaxis=dict(showgrid=True, rangemode="tozero"))
    st.plotly_chart(fig, use_container_width=True, key="g_evo")

# ---------------------------------------------------------------- dados completos
# Lista completa de ausentes do recorte, para conferência ou exportação.
with st.expander("Todos os ausentes do recorte"):
    tabela(pessoas(ausentes), "todos_ausentes.csv")

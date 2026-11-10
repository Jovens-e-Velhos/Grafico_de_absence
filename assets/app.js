/* Painel de faltas da Weekly Meeting — roda 100% no navegador (GitHub Pages).
   Lê planilhas .xlsx com SheetJS e desenha gráficos com ECharts.

   Nenhum dado sai do computador de quem usa: os arquivos são lidos pelo próprio
   navegador e somem ao fechar a página.

   Fluxo geral:
     1. O usuário escolhe ou arrasta arquivos (planilhas ou .zip)   -> carregarArquivos()
     2. Cada planilha vira uma lista de linhas padronizadas        -> lerPlanilha()
     3. Os filtros (Semana, Branch, Office) definem o recorte      -> recorte()
     4. O painel é redesenhado: KPIs, gráficos e tabelas           -> renderPainel()

   As regras de classificação são as mesmas do processamento.py (versão Python);
   se mudar uma regra aqui, mude lá também. */

// Tudo fica dentro desta função para não criar variáveis globais na página.
(() => {
  'use strict';

  // ------------------------------------------------------------ constantes
  // Classificações possíveis de cada participante.
  const CAT = {
    PRES: 'Presente',
    J: 'Justificada',
    JO: 'Justificada (Outros)',
    INJ: 'Injustificada',
    SEM: 'Ausente sem justificativa base',
  };
  // Classificações que contam como falta, na ordem usada nos gráficos e legendas.
  const FALTAS = [CAT.J, CAT.JO, CAT.INJ, CAT.SEM];
  // Paleta DSV (secundárias) validada para daltonismo: azul, azul claro, laranja; cinza = neutro
  const CORES = { [CAT.J]: '#0A34A1', [CAT.JO]: '#4B87E0', [CAT.INJ]: '#C64801', [CAT.SEM]: '#83878D' };
  const AZUL_DSV = '#002664';
  // Valor usado quando uma coluna de hierarquia/local está vazia ou não existe na planilha.
  const SEM_INFO = '(sem informação)';
  // Fonte da identidade DSV; se não estiver instalada, cai para Arial.
  const FONTE = '"Foundry Sterling", Arial, "Helvetica Neue", sans-serif';
  // Respeita a configuração do sistema "reduzir movimento": desliga as animações.
  const REDUZ_MOV = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  // Nome interno de cada coluna -> variações aceitas no cabeçalho da planilha.
  // As variações são comparadas em minúsculas e sem acento (ver norm()), então
  // "Líder Direto" e "LIDER DIRETO" são reconhecidos como "lider direto".
  // Para aceitar um novo nome de coluna, basta acrescentá-lo na lista.
  const COLUNAS = {
    nome: ['full name', 'nome', 'nome completo'],
    email: ['e-mail', 'email'],
    branch: ['branch'],
    office: ['office location'],
    cargo: ['job title'],
    lider: ['lider direto'],
    supervisor: ['supervisor'],
    gestor: ['gestor'],
    gerente: ['gerente'],
    report1: ['report1'],              // presença: "Yes" ou "Absence"
    just_base: ['justificativas base', 'justificativa base'],
    report: ['report'],                // separa "Outros" justificado de injustificado
    outro: ['outro (detalhado)', 'outros (detalhado)'], // texto livre do motivo
  };
  // Sem estas colunas não dá para classificar as faltas; a planilha é recusada com erro.
  const OBRIGATORIAS = ['nome', 'report1', 'just_base', 'report'];
  // Níveis hierárquicos exibidos nas abas: chave interna -> nome na tela.
  const NIVEIS = { lider: 'Líder Direto', supervisor: 'Supervisor', gerente: 'Gerente' };
  // Coluna usada por padrão em cada nível (a de mesmo nome).
  const MAPA_PADRAO = { lider: 'lider', supervisor: 'supervisor', gerente: 'gerente' };
  // Colunas de hierarquia que podem ser escolhidas em "Configurar hierarquia".
  const COLS_HIER = { lider: 'Líder Direto', supervisor: 'Supervisor', gestor: 'Gestor', gerente: 'Gerente' };
  // Valor que significa "este nível não é exibido".
  const NAO_USAR = '';

  // ------------------------------------------------------------ utilidades
  // Atalho para document.querySelector.
  const $ = (s, el = document) => el.querySelector(s);
  // norm: minúsculas, sem acento e sem espaços duplicados — para comparar textos digitados de formas diferentes.
  // limpa: tira espaços sobrando; texto vazio vira null.
  const norm = v => v == null ? '' : String(v).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ').trim().toLowerCase();
  const limpa = v => { if (v == null) return null; const t = String(v).replace(/\s+/g, ' ').trim(); return t || null; };
  // Cria um elemento HTML: el('div', { class: 'x', onclick: fn }, 'texto', outroElemento).
  // Atributos que começam com "on" viram eventos. Texto é inserido como texto puro
  // (nunca como HTML), o que evita que conteúdo da planilha seja interpretado como código.
  const el = (tag, attrs = {}, ...filhos) => {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'class') n.className = v;
      else if (k === 'style') n.style.cssText = v;
      else if (k.startsWith('on')) n.addEventListener(k.slice(2), v);
      else if (v !== false && v != null) n.setAttribute(k, v === true ? '' : v);
    }
    for (const f of filhos.flat()) if (f != null) n.append(f instanceof Node ? f : String(f));
    return n;
  };
  // Conta quantos itens caem em cada grupo: contar(lista, r => r.lider) -> Map { 'Ana' => 3, ... }.
  const contar = (arr, fn) => arr.reduce((m, x) => { const k = fn(x); m.set(k, (m.get(k) || 0) + 1); return m; }, new Map());
  // Percentual no formato brasileiro: pct(3, 40) -> "7,5%".
  const pct = (a, b) => b ? (a / b * 100).toFixed(1).replace('.', ',') + '%' : '–';

  // Extrai a semana do nome do arquivo: "WM43 - ATA 2026.xlsx" -> "WM43".
  // O número fica com 2 dígitos ("WM5" -> "WM05") para a ordenação funcionar.
  // Sem "WM<número>" no nome, usa o próprio nome do arquivo.
  function semanaDoArquivo(nome) {
    const m = nome.match(/WM\s*(\d{1,2})/i);
    return m ? 'WM' + String(+m[1]).padStart(2, '0') : nome.replace(/\.xlsx?$/i, '');
  }

  // ------------------------------------------------------------ leitura e regras
  // Regras de classificação de cada participante:
  //   Report1 ≠ Absence                                   -> Presente
  //   Justificativa Base = "Outros" e Report = Absence    -> Injustificada
  //   Justificativa Base = "Outros" e Report ≠ Absence    -> Justificada (Outros)
  //   Justificativa Base preenchida com outra opção       -> Justificada
  //   Justificativa Base vazia                            -> Ausente sem justificativa base
  function classificar(r) {
    if (norm(r.report1) !== 'absence') return CAT.PRES;
    if (norm(r.just_base).startsWith('outros')) return norm(r.report) === 'absence' ? CAT.INJ : CAT.JO;
    if (norm(r.just_base)) return CAT.J;
    return CAT.SEM;
  }

  // Junta variações de escrita do motivo em "Outro (detalhado)":
  // "DAY OFF", "Day-off" e "day off" viram "Day off". Os acentos são mantidos.
  function agruparDetalhe(v) {
    if (!v) return null;
    const t = v.toLowerCase().replace(/-/g, ' ').replace(/\s+/g, ' ').trim();
    return t.charAt(0).toUpperCase() + t.slice(1);
  }

  // Lê uma planilha (conteúdo em bytes) e devolve uma lista de participantes já padronizados.
  // A planilha tem várias abas (MoM, Attendance, Report...) e o cabeçalho nem sempre está
  // na primeira linha: usa a primeira aba que tiver a célula "Report1" nas 15 primeiras linhas.
  function lerPlanilha(buffer, nomeArquivo) {
    const wb = XLSX.read(buffer, { type: 'array' });
    for (const aba of wb.SheetNames) {
      // Cada linha vira um array de células (header: 1), mantendo as vazias como null.
      const linhas = XLSX.utils.sheet_to_json(wb.Sheets[aba], { header: 1, defval: null, blankrows: true });
      const iCab = linhas.slice(0, 15).findIndex(l => Array.isArray(l) && l.some(v => norm(v) === 'report1'));
      if (iCab < 0) continue; // esta aba não tem os dados de presença; tenta a próxima

      // idx: nome interno da coluna -> posição dela na linha. Se duas colunas
      // corresponderem ao mesmo nome, vale a primeira.
      const idx = {};
      linhas[iCab].forEach((c, j) => {
        const n = norm(c);
        for (const [k, vars] of Object.entries(COLUNAS)) if (vars.includes(n) && !(k in idx)) idx[k] = j;
      });
      const faltando = OBRIGATORIAS.filter(k => !(k in idx));
      if (faltando.length) throw new Error(`colunas obrigatórias ausentes na aba "${aba}": ${faltando.join(', ')}`);

      const semana = semanaDoArquivo(nomeArquivo);
      const linhasOk = [];
      for (const l of linhas.slice(iCab + 1)) {
        if (!Array.isArray(l)) continue;
        const get = k => (k in idx ? l[idx[k]] : null);
        const nome = limpa(get('nome'));
        if (!nome) continue; // linha sem nome (em branco, totais etc.) é ignorada
        const r = { arquivo: nomeArquivo, semana, nome };
        for (const k of ['email', 'cargo', 'report1', 'just_base', 'report', 'outro']) r[k] = limpa(get(k));
        // Hierarquia e local vazios viram "(sem informação)", que aparece como um grupo próprio.
        for (const k of ['lider', 'supervisor', 'gestor', 'gerente', 'branch', 'office']) r[k] = limpa(get(k)) ?? SEM_INFO;
        r.classificacao = classificar(r);
        r.outro_agrupado = agruparDetalhe(r.outro);
        linhasOk.push(r);
      }
      return linhasOk;
    }
    throw new Error('nenhuma aba com a coluna "Report1" foi encontrada');
  }

  // Agrupa as faltas por uma coluna de hierarquia (ex.: 'lider').
  // Devolve uma linha por pessoa com: faltas de cada tipo, total, tamanho da equipe
  // (presentes + ausentes) e % de faltas. Ordena do maior total para o menor.
  function faltasPor(linhas, coluna) {
    const equipe = contar(linhas, r => r[coluna]); // equipe conta todos, não só os ausentes
    const mapa = new Map(); // pessoa -> { Justificada: n, Injustificada: n, ... }
    for (const r of linhas) {
      if (r.classificacao === CAT.PRES) continue;
      const k = r[coluna];
      if (!mapa.has(k)) mapa.set(k, Object.fromEntries(FALTAS.map(c => [c, 0])));
      mapa.get(k)[r.classificacao]++;
    }
    return [...mapa.entries()].map(([nome, c]) => {
      const total = FALTAS.reduce((s, x) => s + c[x], 0);
      const eq = equipe.get(nome) || 0;
      return { nome, ...c, total, equipe: eq, pct: eq ? total / eq * 100 : 0 };
    }).sort((a, b) => b.total - a.total || b.pct - a.pct || a.nome.localeCompare(b.nome)); // empate: maior %, depois ordem alfabética
  }

  // ------------------------------------------------------------ estado
  // Tudo o que o usuário carregou ou escolheu fica aqui; a tela é sempre desenhada a partir deste objeto.
  const estado = {
    arquivos: new Map(), // nome -> linhas
    erros: [],           // mensagens exibidas no topo do painel
    semanas: null,       // Set selecionado (null = padrão: última semana)
    branch: new Set(),   // vazio = todas as branches
    office: new Set(),   // vazio = todos os offices
    mapa: { ...MAPA_PADRAO }, // nível -> coluna da planilha
    nivel: 'lider',      // aba de hierarquia aberta
  };
  // Todas as linhas de todos os arquivos carregados, numa lista só.
  const todas = () => [...estado.arquivos.values()].flat();

  // ------------------------------------------------------------ upload
  // Botão "Adicionar planilhas". O valor é limpo depois da leitura para permitir
  // escolher o mesmo arquivo de novo (ex.: depois de corrigi-lo).
  const input = $('#arquivos');
  input.addEventListener('change', () => { carregarArquivos([...input.files]); input.value = ''; });

  // Arrastar e soltar: destaca a área de soltar enquanto o arquivo está sobre a página.
  // preventDefault impede o navegador de simplesmente abrir o arquivo.
  const drop = $('#drop');
  ['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => {
    e.preventDefault(); drop.classList.add('ativo');
  }));
  ['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => {
    e.preventDefault(); if (ev === 'drop' || e.target === document.documentElement) drop.classList.remove('ativo');
  }));
  // Extensões aceitas.
  const EXT_PLANILHA = /\.(xlsx|xlsm|xls|ods|csv)$/i;
  const EXT_ZIP = /\.zip$/i;

  // Ao soltar arquivos, carrega só os de formato aceito; se nenhum servir, mostra um aviso.
  document.addEventListener('drop', e => {
    const fs = [...(e.dataTransfer?.files || [])].filter(f => EXT_PLANILHA.test(f.name) || EXT_ZIP.test(f.name));
    if (fs.length) carregarArquivos(fs);
    else if (e.dataTransfer?.files?.length) { estado.erros.push('Formato não suportado: use .xlsx, .xlsm, .xls, .ods, .csv ou .zip'); renderErros(); }
  });

  // Converte a seleção em uma lista de planilhas { nome, buffer }, abrindo arquivos .zip
  async function expandir(files) {
    const planilhas = [];
    for (const f of files) {
      const base = f.name.split(/[\\/]/).pop(); // só o nome, sem a pasta
      if (base.startsWith('~$')) continue;      // arquivo temporário do Excel (planilha aberta)
      if (EXT_PLANILHA.test(base)) { planilhas.push({ nome: base, buffer: await f.arrayBuffer() }); continue; }
      if (!EXT_ZIP.test(base)) { estado.erros.push(`${base}: formato não suportado`); continue; }
      try {
        const zip = await JSZip.loadAsync(await f.arrayBuffer());
        // Dentro do .zip, pega só planilhas; ignora pastas, arquivos ocultos,
        // temporários do Excel e a pasta __MACOSX que o Mac cria ao compactar.
        const entradas = Object.values(zip.files).filter(z => {
          const n = z.name.split('/').pop();
          return !z.dir && !z.name.startsWith('__MACOSX/') && !n.startsWith('~$') && !n.startsWith('.') && EXT_PLANILHA.test(n);
        });
        if (!entradas.length) { estado.erros.push(`${base}: nenhuma planilha encontrada dentro do .zip`); continue; }
        for (const z of entradas) planilhas.push({ nome: z.name.split('/').pop(), buffer: await z.async('arraybuffer'), origem: base });
      } catch (err) {
        estado.erros.push(`${base}: não foi possível abrir o .zip (${err.message})`);
      }
    }
    return planilhas;
  }

  // Lê os arquivos escolhidos e redesenha o painel.
  // Um arquivo com o mesmo nome de outro já carregado substitui o anterior.
  // Um arquivo com problema não impede os outros de serem carregados.
  async function carregarArquivos(files) {
    for (const p of await expandir(files)) {
      try {
        estado.arquivos.set(p.nome, lerPlanilha(p.buffer, p.nome));
        estado.erros = estado.erros.filter(e => !e.startsWith(p.nome)); // apaga erro antigo deste arquivo
      } catch (err) {
        estado.erros.push(`${p.nome}${p.origem ? ` (dentro de ${p.origem})` : ''}: ${err.message}`);
      }
    }
    estado.semanas = null; // volta o filtro de semana para "a mais recente"
    ajustarMapaPadrao();
    renderTudo();
  }

  // Decide quais níveis hierárquicos aparecem, conforme as colunas preenchidas nos dados.
  function ajustarMapaPadrao() {
    const dados = todas();
    const tem = c => dados.some(r => r[c] !== SEM_INFO);
    // cada nível usa a coluna de mesmo nome; nível sem coluna na planilha fica oculto até ela aparecer
    for (const n of Object.keys(estado.mapa)) {
      if (!estado.mapa[n] && tem(MAPA_PADRAO[n])) estado.mapa[n] = MAPA_PADRAO[n];
      else if (estado.mapa[n] && !tem(estado.mapa[n])) estado.mapa[n] = NAO_USAR;
    }
    // Se a aba aberta deixou de existir, abre a primeira disponível.
    if (!estado.mapa[estado.nivel]) estado.nivel = Object.keys(estado.mapa).find(n => estado.mapa[n]) || 'lider';
  }

  // ------------------------------------------------------------ recorte
  // Linhas que passam nos filtros. Filtro de Branch/Office vazio = sem restrição.
  // ignorarSemana: usado na evolução semanal, que mostra todas as semanas.
  function recorte({ ignorarSemana = false } = {}) {
    return todas().filter(r =>
      (ignorarSemana || estado.semanas.has(r.semana)) &&
      (!estado.branch.size || estado.branch.has(r.branch)) &&
      (!estado.office.size || estado.office.has(r.office)));
  }

  // ------------------------------------------------------------ render geral
  // Redesenha a página inteira: filtros, configuração e painel.
  // É chamado quando arquivos são adicionados ou removidos.
  function renderTudo() {
    const dados = todas();
    const temDados = dados.length > 0;
    $('#vazio').hidden = temDados;   // tela inicial com a área de soltar arquivos
    $('#painel').hidden = !temDados;
    renderErros();
    if (!temDados) return;

    // Semanas selecionadas: por padrão só a mais recente. Remove semanas de arquivos
    // que foram retirados; se não sobrar nenhuma, volta para a mais recente.
    const semanas = [...new Set(dados.map(r => r.semana))].sort();
    if (!estado.semanas) estado.semanas = new Set(semanas.slice(-1));
    estado.semanas = new Set([...estado.semanas].filter(s => semanas.includes(s)));
    if (!estado.semanas.size) estado.semanas = new Set(semanas.slice(-1));

    renderChips();
    renderFiltros(dados, semanas);
    renderConfig(dados);
    renderPainel();
    observarRevela();
  }

  // Mostra os avisos de arquivos que não puderam ser lidos.
  function renderErros() {
    const box = $('#erros');
    box.replaceChildren(...estado.erros.map(e => el('div', { class: 'alerta' }, '⚠ ', e)));
  }

  // Etiquetas dos arquivos carregados ("WM43 · 398 pessoas"), com botão × para remover.
  function renderChips() {
    const box = $('#chips');
    box.replaceChildren(...[...estado.arquivos.entries()].sort().map(([nome, linhas]) =>
      el('span', { class: 'chip', title: nome },
        el('b', {}, semanaDoArquivo(nome)), ` ${linhas.length} pessoas`,
        el('button', { type: 'button', 'aria-label': `Remover ${nome}`, onclick: () => { estado.arquivos.delete(nome); renderTudo(); } }, '×'))));
  }

  // filtro multi-seleção (details + checkboxes)
  //   rotulo: nome do filtro ("Semana")
  //   opcoes: valores possíveis
  //   selecionados: Set alterado diretamente quando o usuário marca/desmarca
  //   vazioSignificaTodos: true em Branch/Office (nada marcado = todos);
  //                        false em Semana (sempre precisa de ao menos uma marcada)
  //   aoMudar: função chamada a cada mudança (redesenha o painel)
  function filtroMulti(rotulo, opcoes, selecionados, vazioSignificaTodos, aoMudar) {
    // Texto curto mostrado no botão do filtro: "Todos", "WM43, WM44", "3 selecionados"...
    const resumo = () => {
      if (!selecionados.size) return vazioSignificaTodos ? 'Todos' : 'Nenhum';
      if (selecionados.size === opcoes.length) return 'Todos';
      if (selecionados.size <= 2) return [...selecionados].join(', ');
      return `${selecionados.size} selecionados`;
    };
    const b = el('b', {}, resumo());
    const det = el('details', { class: 'filtro' }, el('summary', {}, el('span', {}, rotulo + ': '), b));
    const caixas = opcoes.map(o => {
      const cb = el('input', { type: 'checkbox', value: o });
      cb.checked = selecionados.has(o);
      cb.addEventListener('change', () => {
        cb.checked ? selecionados.add(o) : selecionados.delete(o);
        // Em Semana, não deixa desmarcar a última opção marcada.
        if (!vazioSignificaTodos && !selecionados.size) { selecionados.add(o); cb.checked = true; }
        b.textContent = resumo(); aoMudar();
      });
      return el('label', {}, cb, o);
    });
    // Botões "Marcar todos" (v = true) e "Limpar"/"Só a última" (v = false).
    const marcar = v => () => {
      selecionados.clear();
      if (v || !vazioSignificaTodos) opcoes.forEach(o => selecionados.add(o));
      if (!v && !vazioSignificaTodos) { selecionados.clear(); selecionados.add(opcoes[opcoes.length - 1]); }
      caixas.forEach(l => { l.firstChild.checked = selecionados.has(l.firstChild.value); });
      b.textContent = resumo(); aoMudar();
    };
    det.append(el('div', { class: 'filtro__opcoes' },
      el('div', { class: 'filtro__acoes' },
        el('button', { type: 'button', onclick: marcar(true) }, 'Marcar todos'),
        el('button', { type: 'button', onclick: marcar(false) }, vazioSignificaTodos ? 'Limpar' : 'Só a última')),
      caixas));
    return det;
  }
  // Fecha qualquer filtro aberto ao clicar fora dele.
  document.addEventListener('click', e => {
    document.querySelectorAll('details.filtro[open]').forEach(d => { if (!d.contains(e.target)) d.open = false; });
  });

  // Monta os filtros Semana, Branch e Office a partir dos valores existentes nos dados.
  function renderFiltros(dados, semanas) {
    const unicos = c => [...new Set(dados.map(r => r[c]))].sort((a, b) => a.localeCompare(b));
    $('#filtros').replaceChildren(
      filtroMulti('Semana', semanas, estado.semanas, false, renderPainel),
      filtroMulti('Branch', unicos('branch'), estado.branch, true, renderPainel),
      filtroMulti('Office', unicos('office'), estado.office, true, renderPainel),
    );
  }

  // Painel "Configurar hierarquia": uma caixa de seleção por nível para escolher
  // qual coluna da planilha ele usa. Só lista colunas que têm algum valor preenchido.
  function renderConfig(dados) {
    const disp = Object.keys(COLS_HIER).filter(c => dados.some(r => r[c] !== SEM_INFO));
    $('#config-niveis').replaceChildren(...Object.entries(NIVEIS).map(([nivel, rot]) => {
      const sel = el('select', {},
        ...disp.map(c => el('option', { value: c }, COLS_HIER[c])),
        el('option', { value: NAO_USAR }, '(não usar)'));
      sel.value = disp.includes(estado.mapa[nivel]) ? estado.mapa[nivel] : NAO_USAR;
      sel.addEventListener('change', () => {
        estado.mapa[nivel] = sel.value;
        if (!estado.mapa[estado.nivel]) estado.nivel = Object.keys(estado.mapa).find(n => estado.mapa[n]) || 'lider';
        renderPainel();
      });
      return el('label', {}, rot, sel);
    }));
  }

  // ------------------------------------------------------------ painel
  // Redesenha KPIs, gráficos e tabelas para o recorte atual.
  // É chamado a cada mudança de filtro ou de configuração.
  function renderPainel() {
    const df = recorte();
    const aus = df.filter(r => r.classificacao !== CAT.PRES); // todos os ausentes
    const cont = contar(aus, r => r.classificacao);
    const n = c => cont.get(c) || 0; // nº de faltas de uma classificação

    // KPIs do topo. "Justificadas" soma as justificadas normais e as via "Outros".
    $('#recorte').textContent = `· ${[...estado.semanas].sort().join(', ')}`;
    animarNumero('total', df.length);
    animarNumero('presentes', df.length - aus.length);
    animarNumero('faltas', aus.length);
    animarNumero('justificadas', n(CAT.J) + n(CAT.JO));
    animarNumero('injustificadas', n(CAT.INJ));
    $('[data-kpi-pct="faltas"]').textContent = `${pct(aus.length, df.length)} do total`;
    $('[data-kpi-extra="justificadas"]').textContent = `${n(CAT.JO)} via “Outros”`;

    renderHierarquia(df);
    renderJustificadas(aus);
    renderInjustificadas(aus);
    renderEvolucao();
    tabela($('#t-todos'), colunasPessoa(aus), aus, 'todos_ausentes.csv');
  }

  // KPI com contagem animada
  // O número "corre" do valor anterior até o novo em 0,9 s, desacelerando no final.
  function animarNumero(chave, alvo) {
    const n = $(`[data-kpi="${chave}"]`);
    const de = +n.dataset.v || 0; // valor mostrado antes
    n.dataset.v = alvo;
    if (REDUZ_MOV || de === alvo) { n.textContent = alvo.toLocaleString('pt-BR'); return; }
    const t0 = performance.now(), dur = 900;
    const passo = t => {
      const p = Math.min(1, (t - t0) / dur), e = 1 - Math.pow(1 - p, 3); // p = progresso 0..1; e = progresso suavizado
      n.textContent = Math.round(de + (alvo - de) * e).toLocaleString('pt-BR');
      // Para se um novo valor chegou no meio da animação (outra animação assume).
      if (p < 1 && n.dataset.v == alvo) requestAnimationFrame(passo);
    };
    requestAnimationFrame(passo);
  }

  // ---------- hierarquia
  // Abas (Líder Direto / Supervisor / Gerente), gráfico de barras empilhadas e tabela
  // do nível selecionado.
  function renderHierarquia(df) {
    const ativos = Object.entries(estado.mapa).filter(([, c]) => c); // níveis que não estão em "(não usar)"
    // Se o nível usa uma coluna de outro nome, ela aparece entre parênteses: "Gerente (Supervisor)".
    $('#abas').replaceChildren(...ativos.map(([nivel, col]) => {
      const rot = NIVEIS[nivel];
      return el('button', {
        class: 'aba', role: 'tab', type: 'button', 'aria-selected': String(nivel === estado.nivel),
        onclick: () => { estado.nivel = nivel; renderHierarquia(recorte()); },
      }, rot, COLS_HIER[col] !== rot ? el('small', {}, ` (${COLS_HIER[col]})`) : null);
    }));

    const col = estado.mapa[estado.nivel];
    if (!col) { grafico('g-hier', null); $('#t-hier').replaceChildren(); return; }
    const tab = faltasPor(df, col);
    if (!tab.length) {
      grafico('g-hier', null, 120, 'Nenhuma falta no recorte selecionado. 🎉');
    } else {
      grafico('g-hier', opcaoEmpilhada(tab), Math.max(280, tab.length * 30 + 80)); // altura cresce com o nº de pessoas
    }
    const rot = NIVEIS[estado.nivel];
    // A coluna "sem justificativa base" só aparece na tabela se houver algum caso.
    const cats = FALTAS.filter(c => c !== CAT.SEM || tab.some(t => t[c]));
    tabela($('#t-hier'), [
      { k: 'nome', t: rot }, ...cats.map(c => ({ k: c, t: c, num: true })),
      { k: 'total', t: 'Total de faltas', num: true }, { k: 'equipe', t: 'Equipe', num: true },
      { k: 'pct', t: '% faltas', num: true, fmt: v => v.toFixed(1).replace('.', ',') + '%' },
    ], tab, `faltas_por_${estado.nivel}.csv`);
  }

  // Configuração do ECharts para as barras horizontais empilhadas (uma barra por pessoa,
  // dividida pelo tipo de falta).
  function opcaoEmpilhada(tab) {
    const linhas = [...tab].reverse(); // maior no topo
    const cats = FALTAS.filter(c => c !== CAT.SEM || linhas.some(r => r[c]));
    const series = cats.map((c, i) => ({
      name: c, type: 'bar', stack: 'total', barMaxWidth: 20,
      data: linhas.map(r => r[c] || null), // null em vez de 0: não desenha segmento vazio
      itemStyle: { color: CORES[c], borderColor: '#fff', borderWidth: 1 }, // borda branca separa os segmentos
      emphasis: { focus: 'series' },       // ao passar o mouse, destaca o tipo de falta
      // As barras entram em cascata, de cima para baixo, e cada tipo de falta um pouco depois do anterior.
      animationDelay: idx => (linhas.length - idx) * 22 + i * 140,
    }));
    // Série invisível de valor 0 empilhada no fim, só para escrever o total ao lado da barra.
    series.push({ // rótulo de total ao fim da barra
      name: 'Total', type: 'bar', stack: 'total', data: linhas.map(() => 0), silent: true,
      itemStyle: { color: 'transparent' }, tooltip: { show: false },
      label: { show: true, position: 'right', formatter: p => linhas[p.dataIndex].total, color: '#2F3135', fontWeight: 700, fontFamily: FONTE },
      animationDelay: idx => (linhas.length - idx) * 22 + cats.length * 140,
    });
    return {
      ...base(),
      legend: { data: cats, top: 0, left: 0, icon: 'rect', itemWidth: 12, itemHeight: 12, textStyle: { color: '#2F3135' } },
      grid: { left: 4, right: 44, top: 36, bottom: 4, containLabel: true },
      // Caixa ao passar o mouse: faltas de cada tipo + total, equipe e %.
      tooltip: {
        ...base().tooltip, trigger: 'axis', axisPointer: { type: 'shadow', shadowStyle: { color: 'rgba(0,38,100,.06)' } },
        formatter: ps => {
          const r = linhas[ps[0].dataIndex];
          const itens = cats.filter(c => r[c]).map(c => `${marcador(CORES[c])}${c}: <b>${r[c]}</b>`).join('<br>');
          return `<b>${esc(r.nome)}</b><br>${itens}<br><span style="color:#646768">Total ${r.total} de ${r.equipe} (${pct(r.total, r.equipe)})</span>`;
        },
      },
      xAxis: { type: 'value', minInterval: 1, splitLine: { lineStyle: { color: '#EFF0F0' } }, axisLabel: { color: '#83878D' } },
      yAxis: {
        type: 'category', data: linhas.map(r => r.nome), axisTick: { show: false },
        axisLine: { lineStyle: { color: '#CDD0D3' } },
        axisLabel: { color: '#2F3135', width: 180, overflow: 'truncate' }, // nomes longos são cortados com "…"
      },
      series,
    };
  }

  // ---------- justificadas
  // Rosca com a divisão das faltas por tipo, barras por Justificativa Base,
  // barras por motivo de "Outros" e a lista de quem teve falta justificada.
  function renderJustificadas(aus) {
    const cont = contar(aus, r => r.classificacao);
    const dadosDonut = FALTAS.filter(c => cont.get(c)).map(c => ({ name: c, value: cont.get(c), itemStyle: { color: CORES[c] } }));
    grafico('g-donut', aus.length ? opcaoDonut(dadosDonut, aus.length) : null, 300, 'Sem faltas.');

    // Quantas vezes cada opção de "Justificativa Base" foi escolhida (exceto "Outros").
    const just = contar(aus.filter(r => r.classificacao === CAT.J), r => r.just_base);
    grafico('g-just', just.size ? opcaoBarras(just, CORES[CAT.J]) : null, Math.max(240, just.size * 40 + 40), 'Nenhuma.');

    // Motivos escritos em "Outro (detalhado)" nas faltas justificadas via "Outros".
    const outros = contar(aus.filter(r => r.classificacao === CAT.JO), r => r.outro_agrupado || '(não informado)');
    grafico('g-outros', outros.size ? opcaoBarras(outros, CORES[CAT.JO]) : null, Math.max(240, outros.size * 40 + 40), 'Nenhuma.');

    const linhas = aus.filter(r => r.classificacao === CAT.J || r.classificacao === CAT.JO);
    tabela($('#t-just'), colunasPessoa(linhas), linhas, 'faltas_justificadas.csv');
  }

  // Configuração do gráfico de rosca. O total de faltas fica escrito no centro.
  function opcaoDonut(dados, total) {
    return {
      ...base(),
      tooltip: { ...base().tooltip, trigger: 'item', formatter: p => `${marcador(p.color)}${p.name}: <b>${p.value}</b> (${p.percent.toFixed(1).replace('.', ',')}%)` },
      legend: { bottom: 0, left: 'center', icon: 'rect', itemWidth: 12, itemHeight: 12, textStyle: { color: '#2F3135' } },
      graphic: [{ // texto no centro da rosca: número + "faltas"
        type: 'group', left: 'center', top: '36%',
        children: [
          { type: 'text', style: { text: String(total), fontSize: 34, fontWeight: 300, fill: AZUL_DSV, fontFamily: FONTE, textAlign: 'center' }, left: 'center' },
          { type: 'text', top: 42, style: { text: 'faltas', fontSize: 13, fill: '#646768', fontFamily: FONTE, textAlign: 'center' }, left: 'center' },
        ],
      }],
      series: [{
        type: 'pie', radius: ['55%', '76%'], center: ['50%', '44%'], data: dados,
        itemStyle: { borderColor: '#fff', borderWidth: 2 },
        label: { show: true, formatter: '{c}', color: '#2F3135', fontWeight: 700 },
        labelLine: { length: 6, length2: 6 },
        emphasis: { scale: true, scaleSize: 6 },
        animationType: 'expansion', animationDuration: 1100, // a rosca "abre" girando
      }],
    };
  }

  // Configuração de barras horizontais de uma cor só, com o valor ao lado de cada barra.
  // `mapa` vem de contar(): rótulo -> quantidade.
  function opcaoBarras(mapa, cor) {
    // Ordem crescente porque o ECharts desenha de baixo para cima: o maior fica no topo.
    const pares = [...mapa.entries()].sort((a, b) => a[1] - b[1] || b[0].localeCompare(a[0]));
    return {
      ...base(),
      grid: { left: 4, right: 36, top: 8, bottom: 4, containLabel: true },
      tooltip: { ...base().tooltip, trigger: 'item', formatter: p => `${esc(p.name)}<br><b>${p.value}</b> pessoa(s)` },
      xAxis: { type: 'value', minInterval: 1, splitLine: { lineStyle: { color: '#EFF0F0' } }, axisLabel: { color: '#83878D' } },
      yAxis: {
        type: 'category', data: pares.map(p => p[0]), axisTick: { show: false },
        axisLine: { lineStyle: { color: '#CDD0D3' } }, axisLabel: { color: '#2F3135', width: 170, overflow: 'truncate' },
      },
      series: [{
        type: 'bar', data: pares.map(p => p[1]), barMaxWidth: 22,
        itemStyle: { color: cor },
        emphasis: { itemStyle: { color: AZUL_DSV } },
        label: { show: true, position: 'right', color: '#2F3135', fontWeight: 700 },
        animationDelay: idx => (pares.length - idx) * 70,
      }],
    };
  }

  // ---------- injustificadas
  // Tabela das faltas injustificadas e, à parte, as ausências sem Justificativa Base.
  function renderInjustificadas(aus) {
    const inj = aus.filter(r => r.classificacao === CAT.INJ);
    const box = $('#t-inj');
    if (!inj.length) {
      box.replaceChildren(el('p', { class: 'ok-msg' }, '✓ Nenhuma falta injustificada no recorte.'));
    } else {
      // Tabela enxuta: quem faltou, a quem responde e o que foi escrito em "Outro (detalhado)".
      const cols = colunasPessoa(inj).filter(c => ['semana', 'nome', 'lider', 'supervisor', 'gestor', 'gerente', 'branch', 'outro'].includes(c.k));
      tabela(box, cols, inj, 'faltas_injustificadas.csv');
    }
    // Ausências com "Justificativa Base" em branco não podem ser classificadas; ficam num aviso separado.
    const sem = aus.filter(r => r.classificacao === CAT.SEM);
    const boxSem = $('#sem-base');
    boxSem.replaceChildren();
    if (sem.length) {
      boxSem.append(el('div', { class: 'alerta' }, `${sem.length} ausência(s) sem “Justificativa Base” preenchida.`));
      const c = el('div', { class: 'cartao', style: 'margin-top:12px' });
      boxSem.append(c);
      tabela(c, colunasPessoa(sem), sem, 'ausentes_sem_justificativa.csv');
    }
  }

  // ---------- evolução
  // Gráfico de linhas com as faltas por semana, uma linha por tipo de falta.
  // Só aparece com 2 semanas ou mais. Ignora o filtro de Semana, mas respeita Branch e Office.
  function renderEvolucao() {
    const semanas = [...new Set(todas().map(r => r.semana))].sort();
    const sec = $('#sec-evo');
    sec.hidden = semanas.length < 2;
    if (sec.hidden) return;
    const aus = recorte({ ignorarSemana: true }).filter(r => r.classificacao !== CAT.PRES);
    const cats = FALTAS.filter(c => aus.some(r => r.classificacao === c));
    grafico('g-evo', {
      ...base(),
      legend: { top: 0, left: 0, icon: 'rect', itemWidth: 12, itemHeight: 12, textStyle: { color: '#2F3135' } },
      grid: { left: 4, right: 24, top: 40, bottom: 4, containLabel: true },
      tooltip: { ...base().tooltip, trigger: 'axis', axisPointer: { type: 'line', lineStyle: { color: '#83878D' } } },
      xAxis: { type: 'category', data: semanas, boundaryGap: false, axisLine: { lineStyle: { color: '#CDD0D3' } }, axisLabel: { color: '#2F3135' } },
      yAxis: { type: 'value', minInterval: 1, splitLine: { lineStyle: { color: '#EFF0F0' } }, axisLabel: { color: '#83878D' } },
      series: cats.map(c => ({
        name: c, type: 'line', data: semanas.map(s => aus.filter(r => r.semana === s && r.classificacao === c).length),
        symbol: 'rect', symbolSize: 8, lineStyle: { width: 2 }, itemStyle: { color: CORES[c] },
        emphasis: { focus: 'series' },
        endLabel: { show: true, formatter: '{c}', color: '#2F3135', fontWeight: 700 }, // valor da última semana no fim da linha
        animationDuration: 1400, animationEasing: 'cubicOut',
      })),
    }, 340);
  }

  // ------------------------------------------------------------ gráficos (render ao entrar na tela)
  // Configurações comuns a todos os gráficos: fonte DSV, animações e estilo da caixa de dica (tooltip).
  function base() {
    return {
      textStyle: { fontFamily: FONTE },
      animation: !REDUZ_MOV,
      animationDuration: 900, animationEasing: 'cubicOut',
      animationDurationUpdate: 700, animationEasingUpdate: 'cubicInOut',
      tooltip: {
        backgroundColor: '#fff', borderColor: '#CDD0D3', borderWidth: 1, padding: [8, 12],
        textStyle: { color: '#000', fontFamily: FONTE, fontSize: 13 },
        extraCssText: 'box-shadow:0 6px 18px rgba(0,38,100,.14);border-radius:0;',
      },
    };
  }
  // Quadradinho colorido usado dentro das dicas para identificar a série.
  const marcador = cor => `<span style="display:inline-block;width:10px;height:10px;background:${cor};margin-right:6px"></span>`;
  // Escapa caracteres especiais antes de inserir texto da planilha no HTML das dicas.
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  // Os gráficos só são desenhados quando aparecem na tela (ao rolar a página), para que
  // a animação de entrada seja vista. Até lá, a configuração fica guardada em "pendente".
  const graficos = new Map(); // id -> { chart, pendente, visivel }
  // Observa quando cada gráfico entra na tela (15% visível) e então o desenha.
  const obsGraf = new IntersectionObserver(ents => {
    for (const e of ents) {
      if (!e.isIntersecting) continue;
      const g = graficos.get(e.target.id);
      if (g && !g.visivel) { g.visivel = true; aplicar(e.target.id); }
      obsGraf.unobserve(e.target);
    }
  }, { threshold: 0.15 });

  // Define o conteúdo de um gráfico.
  //   id: id do elemento HTML; opcao: configuração do ECharts (null = sem dados)
  //   altura: em pixels; msgVazio: texto exibido quando não há dados
  // Se o gráfico já está visível, desenha na hora; senão, espera ele aparecer.
  function grafico(id, opcao, altura = 260, msgVazio = '') {
    const box = document.getElementById(id);
    let g = graficos.get(id);
    if (!g) { g = { chart: null, pendente: null, visivel: false }; graficos.set(id, g); obsGraf.observe(box); }
    box.style.height = altura + 'px';
    if (!opcao) {
      if (g.chart) { g.chart.dispose(); g.chart = null; }
      box.replaceChildren(el('p', { class: 'vazio-msg' }, msgVazio));
      g.pendente = null;
      return;
    }
    g.pendente = opcao;
    if (g.visivel) aplicar(id);
  }

  // Desenha de fato a configuração pendente de um gráfico.
  // Cria o gráfico na primeira vez; nas seguintes, reaproveita-o (o ECharts anima a transição).
  function aplicar(id) {
    const g = graficos.get(id);
    if (!g.pendente) return;
    const box = document.getElementById(id);
    if (!g.chart) { box.replaceChildren(); g.chart = echarts.init(box, null, { renderer: 'canvas' }); }
    else g.chart.resize(); // a altura pode ter mudado
    g.chart.setOption(g.pendente, true); // true = substitui a configuração anterior inteira
    g.pendente = null;
  }

  // Redimensiona os gráficos quando a janela muda de tamanho. Espera 120 ms
  // depois do último evento para não redesenhar dezenas de vezes seguidas.
  let tRes;
  window.addEventListener('resize', () => {
    clearTimeout(tRes);
    tRes = setTimeout(() => graficos.forEach(g => g.chart && g.chart.resize()), 120);
  });

  // revela seções ao rolar
  // Seções com a classe "revela" ganham "visivel" ao entrar na tela; o CSS faz o efeito de surgir.
  const obsRevela = new IntersectionObserver(ents => ents.forEach(e => {
    if (e.isIntersecting) { e.target.classList.add('visivel'); obsRevela.unobserve(e.target); }
  }), { threshold: 0.08 });
  function observarRevela() { document.querySelectorAll('.revela:not(.visivel)').forEach(s => obsRevela.observe(s)); }

  // ------------------------------------------------------------ tabelas
  // Colunas das tabelas de pessoas. Supervisor, Gestor e Gerente só aparecem
  // se tiverem algum valor preenchido nas linhas exibidas.
  //   k: campo do dado; t: título; tag: mostra a classificação com o quadradinho colorido
  function colunasPessoa(linhas) {
    const tem = c => linhas.some(r => r[c] !== SEM_INFO);
    return [
      { k: 'semana', t: 'Semana' }, { k: 'nome', t: 'Nome' }, { k: 'lider', t: 'Líder Direto' },
      tem('supervisor') && { k: 'supervisor', t: 'Supervisor' },
      tem('gestor') && { k: 'gestor', t: 'Gestor' },
      tem('gerente') && { k: 'gerente', t: 'Gerente' },
      { k: 'branch', t: 'Branch' }, { k: 'office', t: 'Office' },
      { k: 'just_base', t: 'Justificativa Base' }, { k: 'report', t: 'Report' },
      { k: 'outro', t: 'Outro (detalhado)' }, { k: 'classificacao', t: 'Classificação', tag: true },
    ].filter(Boolean);
  }

  // Monta uma tabela dentro de `box`, com contador de registros e botão "Baixar CSV".
  //   cols: lista de { k, t, num?, fmt?, tag? }
  //     num: alinha à direita (números); fmt: função que formata o valor
  //   linhas: os dados; arquivoCsv: nome do arquivo baixado
  function tabela(box, cols, linhas, arquivoCsv) {
    if (!linhas.length) { box.replaceChildren(el('p', { class: 'vazio-msg' }, 'Nenhum registro.')); return; }
    const fmt = (c, r) => { const v = r[c.k]; return v == null ? '' : c.fmt ? c.fmt(v) : String(v); };
    // As linhas surgem em cascata (até a 20ª; as demais aparecem juntas).
    const corpo = linhas.map((r, i) => el('tr', { style: `animation-delay:${Math.min(i, 20) * 18}ms` },
      cols.map(c => el('td', { class: c.num ? 'num' : '' },
        c.tag ? el('span', { class: 'tag', style: `--c:${CORES[r[c.k]] || '#83878D'}` }, fmt(c, r)) : fmt(c, r)))));
    // Gera o CSV no próprio navegador e dispara o download.
    // Usa ";" como separador e começa com o BOM (marca UTF-8) para abrir certo,
    // com acentos e colunas separadas, no Excel em português.
    const csv = () => {
      const q = v => `"${String(v).replace(/"/g, '""')}"`;
      const txt = [cols.map(c => q(c.t)).join(';'), ...linhas.map(r => cols.map(c => q(fmt(c, r))).join(';'))].join('\r\n');
      const a = el('a', { href: URL.createObjectURL(new Blob(['﻿' + txt], { type: 'text/csv;charset=utf-8' })), download: arquivoCsv });
      document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    };
    box.replaceChildren(
      el('div', { class: 'tabela-topo' },
        el('span', {}, `${linhas.length} registro(s)`),
        el('button', { type: 'button', class: 'btn btn--azul btn--mini', onclick: csv }, 'Baixar CSV')),
      el('div', { class: 'tabela-rolagem' },
        el('table', {}, el('thead', {}, el('tr', {}, cols.map(c => el('th', {}, c.t)))), el('tbody', {}, corpo))));
  }

  // Primeira renderização: mostra a tela inicial (nenhum arquivo carregado).
  renderTudo();
})();

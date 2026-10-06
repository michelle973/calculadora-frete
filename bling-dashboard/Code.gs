/**
 * Dashboard de vendas Bling -> Google Planilhas
 *
 * - Aba Dashboard: top 5 produtos (pai) por canal, 7 e 30 dias vs períodos anteriores.
 * - Página web (Implantar > App da Web): dashboards de celular Meta Marketplaces
 *   (?d=marketplaces) e Meta Grupo (?d=grupo), com os números do mês atual.
 *
 * Uso: menu "Bling" na planilha -> "1. Configurar" -> "2. Atualizar agora"
 *      -> "3. Ativar atualização automática".
 */

// ===== Configuração =====

const API = 'https://api.bling.com.br/Api/v3';
const TZ = 'America/Sao_Paulo';

// Código da loja no Bling -> nome do canal no dashboard
const CANAIS = {
  204574266: 'Shopify + Matriz', // Shopify
  204787717: 'Shopify + Matriz', // Matriz
  204787719: 'Corporativo',
  204859624: 'Mercado Livre',
  205453078: 'Shopee',
  205507010: 'Amazon',
};
const ORDEM_CANAIS = ['Shopify + Matriz', 'Corporativo', 'Mercado Livre', 'Shopee', 'Amazon'];

// Situações que NÃO entram na conta (padrão do Bling: 12 = Cancelado, 21 = Em digitação)
const SITUACOES_EXCLUIDAS = [12, 21];

// Produtos com estes termos no nome ficam fora do ranking (sem diferenciar maiúsculas/acentos)
const TERMOS_EXCLUIDOS = ['personalizacao']; // não pega "personalizado/personalizada"

// Dashboard Meta Marketplaces: faturamento líquido = Total Venda - Frete
// Metas ficam na aba "Metas" da planilha; estes valores só preenchem a aba na primeira vez
const META_MARKETPLACES = 50000;
const MARKETPLACES = [
  { nome: 'Mercado Livre', lojas: [204859624], legenda: 'Vendas sem vendedor · líquido' },
  { nome: 'Shopee', lojas: [205453078], legenda: 'Total Venda - Frete' },
  { nome: 'Amazon', lojas: [205507010], legenda: 'Total Venda - Frete' },
];

// Dashboard Meta Grupo: base = Total Venda, por vendedor
// Meta total do grupo = soma das metas dos grupos
const GRUPOS = [
  { nome: 'Hursula', meta: 50000 },
  { nome: 'Carlos', meta: 40000 },
  { nome: 'Mitcha', meta: 25000 },
  { nome: 'Outros', meta: 18000 },
  { nome: 'Loja', meta: 57000 },
];
// Nome do vendedor no Bling (sem acento, minúsculo) -> grupo. Vendedor fora da lista é ignorado.
// A aba "Vendedores" guarda o grupo de cada vendedor e pode ser corrigida à mão.
const VENDEDOR_GRUPO = {
  'hursula ramos': 'Hursula',
  'hursula shopify': 'Hursula',
  'carlos': 'Carlos',
  'mitchaelle': 'Mitcha',
  'michelle s dos santos moritz': 'Outros',
  'luana puel': 'Outros',
  'michelle': 'Outros',
  'andre': 'Outros',
  'atendimento': 'Outros',
  'shopify': 'Loja',
};

const TOP_N = 5;
const DIAS_HISTORICO = 60; // 30 dias + 30 anteriores
const LIMITE_EXECUCAO_MS = 4 * 60 * 1000; // Apps Script corta em 6 min; sobra tempo para gravar
const INTERVALO_COMPLETA_MS = 6 * 3600 * 1000; // releitura completa dos 60 dias a cada 6h

const COLS_PEDIDOS = ['pedidoId', 'data', 'lojaId', 'situacaoId', 'vendedorId', 'total', 'frete', 'detalhe'];
const COLS_ITENS = ['pedidoId', 'data', 'canal', 'produtoPaiId', 'produto', 'quantidade', 'valor'];
const N_ITENS = COLS_ITENS.length;

// ===== Menu =====

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Bling')
    .addItem('1. Configurar', 'configurar')
    .addItem('2. Atualizar agora', 'atualizar')
    .addItem('3. Ativar atualização automática', 'ativarAutomatico')
    .addToUi();
}

function configurar() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const campos = [
    ['CLIENT_ID', 'Client ID do app no Bling'],
    ['CLIENT_SECRET', 'Client Secret do app no Bling'],
    ['REFRESH_TOKEN', 'Refresh token'],
  ];
  for (const [chave, rotulo] of campos) {
    const r = ui.prompt('Configurar Bling', rotulo + ':', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    props.setProperty(chave, r.getResponseText().trim());
  }
  props.deleteProperty('ACCESS_TOKEN');
  props.deleteProperty('ACCESS_EXPIRA');
  obterToken(); // valida na hora
  ui.alert('Configurado com sucesso! Agora use "2. Atualizar agora".');
}

function ativarAutomatico() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'atualizar')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('atualizar').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getUi().alert('Pronto! Os dashboards vão se atualizar sozinhos a cada 15 minutos.');
}

// Se a carga não terminou, agenda outra rodada (a primeira carga leva várias)
function agendarContinuacao(pendentes, minutos) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'continuarCarga')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (pendentes) ScriptApp.newTrigger('continuarCarga').timeBased().after((minutos || 1) * 60 * 1000).create();
}

function continuarCarga() { atualizar(); }

// ===== Autenticação =====

function obterToken() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('ACCESS_TOKEN');
  const expira = Number(props.getProperty('ACCESS_EXPIRA') || 0);
  if (token && Date.now() < expira) return token;

  const id = props.getProperty('CLIENT_ID');
  const secret = props.getProperty('CLIENT_SECRET');
  const refresh = props.getProperty('REFRESH_TOKEN');
  if (!id || !secret || !refresh) throw new Error('Use o menu Bling > 1. Configurar primeiro.');

  const resp = UrlFetchApp.fetch(API + '/oauth/token', {
    method: 'post',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(id + ':' + secret),
      Accept: '1.0',
    },
    payload: { grant_type: 'refresh_token', refresh_token: refresh },
    muteHttpExceptions: true,
  });
  const json = JSON.parse(resp.getContentText());
  if (!json.access_token) {
    throw new Error('Falha ao renovar token do Bling (refazer autorização): ' + resp.getContentText());
  }
  // O Bling troca o refresh_token a cada renovação: salvar o novo
  props.setProperties({
    ACCESS_TOKEN: json.access_token,
    ACCESS_EXPIRA: String(Date.now() + (json.expires_in - 300) * 1000),
    REFRESH_TOKEN: json.refresh_token,
  });
  return json.access_token;
}

// opcional = true: devolve null se o app não tiver permissão (403) em vez de dar erro
function apiGet(caminho, opcional) {
  for (let tentativa = 0; tentativa < 5; tentativa++) {
    Utilities.sleep(350); // limite do Bling: 3 requisições por segundo
    const resp = UrlFetchApp.fetch(API + caminho, {
      headers: { Authorization: 'Bearer ' + obterToken(), Accept: 'application/json' },
      muteHttpExceptions: true,
    });
    const code = resp.getResponseCode();
    if (code === 200) return JSON.parse(resp.getContentText());
    if (code === 429) { Utilities.sleep(2000 * (tentativa + 1)); continue; }
    if (code === 401) { PropertiesService.getScriptProperties().deleteProperty('ACCESS_TOKEN'); continue; }
    if (code === 404 || (opcional && code === 403)) return null;
    throw new Error('Bling ' + code + ' em ' + caminho + ': ' + resp.getContentText());
  }
  throw new Error('Bling não respondeu após várias tentativas: ' + caminho);
}

// ===== Atualização =====

function atualizar() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // já tem uma atualização rodando
  try {
    atualizarInterno();
  } finally {
    lock.releaseLock();
  }
}

function atualizarInterno() {
  const inicio = Date.now();
  // Rede de segurança: se esta rodada for interrompida, outra começa em 8 minutos
  agendarContinuacao(true, 8);
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SS_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  lerMetas(SpreadsheetApp.getActiveSpreadsheet()); // cria a aba Metas se ainda não existir
  const hoje = new Date();
  const inicioMes = fmt(hoje).slice(0, 8) + '01';

  // Releitura completa (60 dias) a cada 6h; nas outras rodadas só o mês atual e as 2 últimas semanas
  const completa = Date.now() - Number(props.getProperty('ULTIMA_COMPLETA') || 0) > INTERVALO_COMPLETA_MS;
  const periodoIni = completa
    ? fmt(addDias(hoje, -DIAS_HISTORICO))
    : [inicioMes, fmt(addDias(hoje, -14))].sort()[0];
  const periodoFim = fmt(hoje);

  // 1. Lista os pedidos do período (a listagem já traz loja, situação e total)
  const listados = {};
  for (let pagina = 1; ; pagina++) {
    const r = apiGet('/pedidos/vendas?limite=100&pagina=' + pagina +
      '&dataInicial=' + periodoIni + '&dataFinal=' + periodoFim);
    if (!r || !r.data || r.data.length === 0) break;
    for (const p of r.data) {
      listados[p.id] = { data: p.data, lojaId: p.loja ? p.loja.id : 0, situacao: p.situacao.id, total: p.total };
    }
  }

  // 2. Atualiza a situação dos pedidos e busca o detalhe dos que faltam
  //    (continua na próxima execução se faltar tempo)
  const abaPed = aba('Pedidos', COLS_PEDIDOS);
  const pedidos = lerPedidos(abaPed);
  const abaItens = aba('Itens', COLS_ITENS);
  abaItens.getRange(1, 1, 1, N_ITENS).setValues([COLS_ITENS]); // planilhas antigas não tinham "valor"
  const comItens = new Set(abaItens.getLastRow() > 1
    ? abaItens.getRange(2, 1, abaItens.getLastRow() - 1, 1).getValues().map(l => String(l[0]))
    : []);
  const cacheProdutos = carregarCacheProdutos();
  const vendedores = carregarVendedores();
  const novasLinhas = [];
  let pendentes = 0;

  for (const id of Object.keys(listados)) {
    const l = listados[id];
    const reg = pedidos[id] || (pedidos[id] = { id: id, vendedorId: '', frete: '', detalhe: 0 });
    reg.data = l.data;
    reg.lojaId = l.lojaId;
    reg.situacao = l.situacao;
    if (!reg.detalhe) reg.total = l.total;

    const canal = CANAIS[l.lojaId];
    const precisaItens = canal && !comItens.has(id);
    const precisaInfo = l.data >= inicioMes && !reg.detalhe; // vendedor e frete só para o mês atual
    if (SITUACOES_EXCLUIDAS.indexOf(l.situacao) !== -1 || !(precisaItens || precisaInfo)) continue;
    if (Date.now() - inicio > LIMITE_EXECUCAO_MS) { pendentes++; continue; }

    const det = apiGet('/pedidos/vendas/' + id);
    if (!det || !det.data) continue;
    const d = det.data;
    reg.vendedorId = d.vendedor && d.vendedor.id ? d.vendedor.id : '';
    reg.total = Number(d.total) || 0;
    reg.frete = Number(d.transporte && d.transporte.frete) || 0;
    reg.detalhe = 1;
    if (reg.vendedorId) garantirVendedor(reg.vendedorId, vendedores);

    if (precisaItens) {
      for (const item of d.itens || []) {
        const pai = produtoPai(item, cacheProdutos);
        const qtd = Number(item.quantidade) || 0;
        const unitario = (Number(item.valor) || 0) * (1 - (Number(item.desconto) || 0) / 100);
        novasLinhas.push([id, l.data, canal, pai.id, pai.nome, qtd, qtd * unitario]);
      }
      comItens.add(id);
    }
  }
  if (novasLinhas.length) {
    abaItens.getRange(abaItens.getLastRow() + 1, 1, novasLinhas.length, N_ITENS).setValues(novasLinhas);
  }
  salvarCacheProdutos(cacheProdutos);
  salvarVendedores(vendedores);
  gravarPedidos(abaPed, pedidos);
  if (completa && !pendentes) props.setProperty('ULTIMA_COMPLETA', String(Date.now()));
  agendarContinuacao(pendentes);
  props.setProperty('PENDENTES', String(pendentes));
  props.setProperty('ATUALIZADO_EM', String(Date.now()));

  // 3. Monta o ranking só com pedidos válidos no momento (se um pedido for cancelado depois, some da conta)
  const validos = {};
  Object.keys(pedidos).forEach(id => {
    if (SITUACOES_EXCLUIDAS.indexOf(pedidos[id].situacao) === -1) validos[id] = true;
  });
  montarDashboard(abaItens, validos, pendentes);
}

// ===== Pedidos =====

function lerPedidos(abaPed) {
  const pedidos = {};
  if (abaPed.getLastRow() < 2) return pedidos;
  abaPed.getRange(2, 1, abaPed.getLastRow() - 1, COLS_PEDIDOS.length).getValues().forEach(l => {
    const id = String(l[0]);
    pedidos[id] = {
      id: id, data: l[1] instanceof Date ? fmt(l[1]) : String(l[1]), lojaId: Number(l[2]),
      situacao: Number(l[3]), vendedorId: l[4], total: Number(l[5]) || 0, frete: Number(l[6]) || 0,
      detalhe: Number(l[7]) || 0,
    };
  });
  return pedidos;
}

function gravarPedidos(abaPed, pedidos) {
  const limite = fmt(addDias(new Date(), -(DIAS_HISTORICO + 5)));
  const linhas = Object.values(pedidos)
    .filter(p => p.data >= limite)
    .map(p => [p.id, p.data, p.lojaId, p.situacao, p.vendedorId, p.total, p.frete, p.detalhe]);
  if (abaPed.getLastRow() > 1) abaPed.getRange(2, 1, abaPed.getLastRow() - 1, COLS_PEDIDOS.length).clearContent();
  if (linhas.length) abaPed.getRange(2, 1, linhas.length, COLS_PEDIDOS.length).setValues(linhas);
}

// ===== Vendedores =====

function carregarVendedores() {
  const a = aba('Vendedores', ['vendedorId', 'nome', 'grupo']);
  const vendedores = { _novos: [] };
  if (a.getLastRow() > 1) {
    a.getRange(2, 1, a.getLastRow() - 1, 3).getValues()
      .forEach(([id, nome, grupo]) => { vendedores[id] = { nome: nome, grupo: grupo }; });
  }
  return vendedores;
}

function garantirVendedor(id, vendedores) {
  if (vendedores[id]) return;
  const r = apiGet('/vendedores/' + id, true);
  const nome = r && r.data && r.data.contato ? r.data.contato.nome : '';
  const grupo = VENDEDOR_GRUPO[normalizar(nome)] || '';
  vendedores[id] = { nome: nome, grupo: grupo };
  vendedores._novos.push([id, nome, grupo]);
}

function salvarVendedores(vendedores) {
  if (!vendedores._novos.length) return;
  const a = aba('Vendedores', ['vendedorId', 'nome', 'grupo']);
  a.getRange(a.getLastRow() + 1, 1, vendedores._novos.length, 3).setValues(vendedores._novos);
}

// ===== Produto pai =====

function carregarCacheProdutos() {
  const a = aba('Produtos', ['produtoId', 'produtoPaiId', 'nomePai']);
  const cache = {};
  if (a.getLastRow() > 1) {
    a.getRange(2, 1, a.getLastRow() - 1, 3).getValues()
      .forEach(([id, paiId, nome]) => { cache[id] = { id: paiId, nome: nome }; });
  }
  cache._novos = [];
  return cache;
}

function salvarCacheProdutos(cache) {
  if (!cache._novos.length) return;
  const a = aba('Produtos', ['produtoId', 'produtoPaiId', 'nomePai']);
  a.getRange(a.getLastRow() + 1, 1, cache._novos.length, 3).setValues(cache._novos);
}

function produtoPai(item, cache) {
  const prodId = item.produto && item.produto.id;
  if (!prodId) return { id: 'sem-cadastro:' + item.descricao, nome: item.descricao };
  if (cache[prodId]) return cache[prodId];

  let pai = { id: prodId, nome: item.descricao };
  const p = apiGet('/produtos/' + prodId);
  if (p && p.data) {
    pai = { id: prodId, nome: p.data.nome };
    const paiId = p.data.variacao && p.data.variacao.produtoPai && p.data.variacao.produtoPai.id;
    if (paiId) {
      const pp = cache[paiId] || (() => {
        const r = apiGet('/produtos/' + paiId);
        return { id: paiId, nome: r && r.data ? r.data.nome : pai.nome };
      })();
      pai = { id: paiId, nome: pp.nome };
    }
  }
  cache[prodId] = pai;
  cache._novos.push([prodId, pai.id, pai.nome]);
  return pai;
}

// ===== Dashboard =====

function montarDashboard(abaItens, pedidosValidos, pendentes) {
  const hoje = new Date();
  // Cada período: janela atual (até ontem) e a janela anterior de mesmo tamanho
  const periodos = [7, 30].map(dias => ({
    dias: dias,
    atualIni: fmt(addDias(hoje, -dias)), atualFim: fmt(addDias(hoje, -1)),
    antIni: fmt(addDias(hoje, -2 * dias)), antFim: fmt(addDias(hoje, -dias - 1)),
  }));

  // soma[canal][produtoPaiId] = { paiId, nome, a7, p7, a30, p30, valor30, qtdValor30 }
  const soma = {};
  ORDEM_CANAIS.forEach(c => { soma[c] = {}; });
  if (abaItens.getLastRow() > 1) {
    for (const [pedidoId, dataRaw, canal, paiId, nome, qtd, valor] of
      abaItens.getRange(2, 1, abaItens.getLastRow() - 1, N_ITENS).getValues()) {
      if (!pedidosValidos[pedidoId] || !soma[canal] || produtoExcluido(nome)) continue;
      const data = dataRaw instanceof Date ? fmt(dataRaw) : String(dataRaw);
      const reg = soma[canal][paiId] || (soma[canal][paiId] =
        { paiId: paiId, nome: nome, a7: 0, p7: 0, a30: 0, p30: 0, valor30: 0, qtdValor30: 0 });
      for (const p of periodos) {
        if (data >= p.atualIni && data <= p.atualFim) reg['a' + p.dias] += qtd;
        else if (data >= p.antIni && data <= p.antFim) reg['p' + p.dias] += qtd;
      }
      // preço médio: só linhas que têm valor (itens carregados antes desta versão não têm)
      if (valor !== '' && data >= periodos[1].atualIni && data <= periodos[1].atualFim) {
        reg.valor30 += Number(valor) || 0;
        reg.qtdValor30 += qtd;
      }
    }
  }

  // Top N de cada canal/período e estoque disponível dos produtos que aparecem
  const tops = {};
  const paisVisiveis = new Set();
  for (const canal of ORDEM_CANAIS) {
    tops[canal] = {};
    for (const p of periodos) {
      const k = 'a' + p.dias;
      tops[canal][p.dias] = Object.values(soma[canal]).filter(r => r[k] > 0)
        .sort((a, b) => b[k] - a[k]).slice(0, TOP_N);
      tops[canal][p.dias].forEach(r => paisVisiveis.add(String(r.paiId)));
    }
  }
  const estoque = estoqueDisponivel(Array.from(paisVisiveis));

  const d = aba('Dashboard');
  d.clear();
  d.clearFormats(); // o layout mudou de colunas; não deixar formato antigo (%) sobrar
  d.getRange('A1').setValue('Top ' + TOP_N + ' produtos por canal (unidades vendidas)')
    .setFontSize(16).setFontWeight('bold');
  d.getRange('A2').setValue(
    periodos.map(p => p.dias + ' dias: ' + br(p.atualIni) + ' a ' + br(p.atualFim) +
      ' vs ' + br(p.antIni) + ' a ' + br(p.antFim)).join('   |   ') +
    '   |   Atualizado em ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') +
    (pendentes ? '   |   ⚠ ' + pendentes + ' pedidos ainda sendo carregados' : ''))
    .setFontColor('#666666');

  // 7 dias nas colunas A-F, 30 dias nas colunas H-N
  const colunas = { 7: 1, 30: 8 };
  let linha = 4;
  for (const canal of ORDEM_CANAIS) {
    d.getRange(linha, 1).setValue(canal).setFontSize(13).setFontWeight('bold');
    linha++;
    let altura = 0;
    for (const p of periodos) {
      altura = Math.max(altura, tabelaTop(d, linha, colunas[p.dias], tops[canal][p.dias], p.dias, estoque));
    }
    linha += altura + 2;
  }
  for (const [col, nCols] of [[1, 6], [8, 7]]) {
    d.setColumnWidth(col, 30);
    d.setColumnWidth(col + 1, 320);
    d.setColumnWidths(col + 2, nCols - 2, 105);
    d.getRange(4, col + 2, Math.max(linha - 4, 1), nCols - 2).setHorizontalAlignment('center');
  }
  d.setColumnWidth(7, 30);

  limparItensAntigos(abaItens);
}

// Escreve a tabela Top N de um período e devolve quantas linhas de dados ocupou.
// 7 dias: # | Produto | atual | anterior | Variação | Estoque
// 30 dias: idem + Preço médio antes do Estoque
function tabelaTop(d, linha, col, lista, dias, estoque) {
  const atual = 'a' + dias, ant = 'p' + dias;
  const comPreco = dias === 30;
  const cab = ['#', 'Produto', 'Últimos ' + dias + ' dias', dias + ' dias anteriores', 'Variação']
    .concat(comPreco ? ['Preço médio'] : [], ['Estoque disp.']);
  d.getRange(linha, col, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#efefef');
  if (!lista.length) {
    d.getRange(linha + 1, col + 1).setValue('Sem vendas no período').setFontColor('#999999');
    return 1;
  }
  const valores = lista.map((r, i) => {
    const est = estoque[String(r.paiId)];
    return [i + 1, r.nome, r[atual], r[ant], r[ant] ? (r[atual] - r[ant]) / r[ant] : 'novo']
      .concat(comPreco ? [r.qtdValor30 ? r.valor30 / r.qtdValor30 : '—'] : [], [est === undefined ? '—' : est]);
  });
  d.getRange(linha + 1, col, valores.length, cab.length).setValues(valores);
  d.getRange(linha + 1, col + 2, valores.length, 2).setNumberFormat('0');
  d.getRange(linha + 1, col + 4, valores.length, 1).setNumberFormat('+0%;-0%;0%');
  d.getRange(linha + 1, cab.length + col - 1, valores.length, 1).setNumberFormat('0');
  if (comPreco) d.getRange(linha + 1, col + 5, valores.length, 1).setNumberFormat('"R$" #,##0.00');
  valores.forEach((v, i) => {
    if (typeof v[4] === 'number') {
      d.getRange(linha + 1 + i, col + 4).setFontColor(v[4] > 0 ? '#188038' : v[4] < 0 ? '#d93025' : '#000000');
    }
  });
  return valores.length;
}

// ===== Estoque =====

// Saldo virtual (disponível) de cada produto pai, somando todas as variações.
// Devolve { paiId: saldo }; produtos sem permissão/sem cadastro ficam de fora.
function estoqueDisponivel(paiIds) {
  const ids = paiIds.filter(id => /^\d+$/.test(id));
  if (!ids.length) return {};

  // Variações de cada pai (guardadas na aba "Variacoes" para não consultar toda vez)
  const abaVar = aba('Variacoes', ['produtoPaiId', 'variacoesIds']);
  const variacoes = {};
  if (abaVar.getLastRow() > 1) {
    abaVar.getRange(2, 1, abaVar.getLastRow() - 1, 2).getValues()
      .forEach(([pai, lista]) => { variacoes[String(pai)] = String(lista).split(',').filter(x => x); });
  }
  const novas = [];
  for (const pai of ids) {
    if (variacoes[pai]) continue;
    const r = apiGet('/produtos/' + pai, true);
    const lista = r && r.data && r.data.variacoes && r.data.variacoes.length
      ? r.data.variacoes.map(v => String(v.id))
      : [pai]; // produto simples: o saldo é dele mesmo
    variacoes[pai] = lista;
    novas.push([pai, lista.join(',')]);
  }
  if (novas.length) abaVar.getRange(abaVar.getLastRow() + 1, 1, novas.length, 2).setValues(novas);

  // Saldos em lotes de 50 produtos
  const todos = [];
  ids.forEach(pai => variacoes[pai].forEach(v => todos.push(v)));
  const saldo = {};
  for (let i = 0; i < todos.length; i += 50) {
    const lote = todos.slice(i, i + 50);
    const r = apiGet('/estoques/saldos?' + lote.map(id => 'idsProdutos%5B%5D=' + id).join('&'), true);
    if (!r || !r.data) return {}; // sem permissão de estoque no app
    r.data.forEach(e => { saldo[String(e.produto.id)] = Number(e.saldoVirtualTotal) || 0; });
  }
  const resultado = {};
  ids.forEach(pai => {
    resultado[pai] = variacoes[pai].reduce((s, v) => s + (saldo[v] || 0), 0);
  });
  return resultado;
}

function limparItensAntigos(abaItens) {
  if (abaItens.getLastRow() < 2) return;
  const limite = fmt(addDias(new Date(), -(DIAS_HISTORICO + 5)));
  const linhas = abaItens.getRange(2, 1, abaItens.getLastRow() - 1, N_ITENS).getValues();
  const manter = linhas.filter(l => (l[1] instanceof Date ? fmt(l[1]) : String(l[1])) >= limite);
  if (manter.length === linhas.length) return;
  abaItens.getRange(2, 1, linhas.length, N_ITENS).clearContent();
  if (manter.length) abaItens.getRange(2, 1, manter.length, N_ITENS).setValues(manter);
}

// ===== Página web: dashboards Meta Marketplaces e Meta Grupo =====

const MESES = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
  'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];

function doGet(e) {
  const tipo = e && e.parameter && e.parameter.d === 'marketplaces' ? 'marketplaces' : 'grupo';
  const dados = calcularMetas();
  const html = tipo === 'marketplaces' ? paginaMarketplaces(dados) : paginaGrupo(dados);
  return HtmlService.createHtmlOutput(html)
    .setTitle(tipo === 'marketplaces' ? 'Amouh · Meta Marketplaces' : 'Amouh · Meta Grupo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function calcularMetas() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SS_ID');
  const ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  const hoje = new Date();
  const hojeStr = fmt(hoje);
  const inicioMes = hojeStr.slice(0, 8) + '01';
  const [ano, mes, dia] = hojeStr.split('-').map(Number);
  const diasMes = new Date(ano, mes, 0).getDate();

  const abaPed = ss.getSheetByName('Pedidos');
  const pedidos = abaPed ? Object.values(lerPedidos(abaPed)) : [];
  const vendedores = {};
  const abaVend = ss.getSheetByName('Vendedores');
  if (abaVend && abaVend.getLastRow() > 1) {
    abaVend.getRange(2, 1, abaVend.getLastRow() - 1, 3).getValues()
      .forEach(([id, , grupo]) => { vendedores[String(id)] = String(grupo).trim(); });
  }

  const porMarketplace = {};
  MARKETPLACES.forEach(m => { porMarketplace[m.nome] = 0; });
  const porGrupo = {};
  GRUPOS.forEach(g => { porGrupo[g.nome] = 0; });

  for (const p of pedidos) {
    if (p.data < inicioMes || p.data > hojeStr) continue;
    if (SITUACOES_EXCLUIDAS.indexOf(p.situacao) !== -1 || !p.detalhe) continue;
    const mkt = MARKETPLACES.find(m => m.lojas.indexOf(p.lojaId) !== -1);
    if (mkt) porMarketplace[mkt.nome] += p.total - p.frete;
    const grupo = vendedores[String(p.vendedorId)];
    if (grupo && porGrupo[grupo] !== undefined) porGrupo[grupo] += p.total;
  }

  const atualizado = Number(props.getProperty('ATUALIZADO_EM') || 0);
  return {
    data: br(hojeStr),
    hora: atualizado ? Utilities.formatDate(new Date(atualizado), TZ, 'HH:mm') : '',
    mes: MESES[mes - 1],
    dia: dia,
    diasMes: diasMes,
    pendentes: Number(props.getProperty('PENDENTES') || 0),
    porMarketplace: porMarketplace,
    porGrupo: porGrupo,
    metas: lerMetas(ss),
  };
}

// Lê a aba "Metas" (criada com os valores padrão se não existir). Devolve { Marketplaces, Hursula, ... }
function lerMetas(ss) {
  const padrao = [['Marketplaces', META_MARKETPLACES]].concat(GRUPOS.map(g => [g.nome, g.meta]));
  let a = ss.getSheetByName('Metas');
  if (!a) {
    a = ss.insertSheet('Metas');
    a.getRange(1, 1, 1, 2).setValues([['meta', 'valor (R$)']]).setFontWeight('bold');
    a.getRange(2, 1, padrao.length, 2).setValues(padrao);
  }
  const metas = {};
  padrao.forEach(([nome, valor]) => { metas[nome] = valor; });
  if (a.getLastRow() > 1) {
    a.getRange(2, 1, a.getLastRow() - 1, 2).getValues().forEach(([nome, valor]) => {
      const v = Number(String(valor).replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
      if (metas[String(nome).trim()] !== undefined && v > 0) metas[String(nome).trim()] = v;
    });
  }
  return metas;
}

function resumoMeta(total, meta, dados) {
  const projecao = total / dados.dia * dados.diasMes;
  return {
    total: total, meta: meta,
    pct: meta ? total / meta : 0,
    falta: Math.max(meta - total, 0),
    projecao: projecao,
    projecaoPct: meta ? projecao / meta : 0,
  };
}

function paginaMarketplaces(dados) {
  const total = MARKETPLACES.reduce((s, m) => s + dados.porMarketplace[m.nome], 0);
  const r = resumoMeta(total, dados.metas.Marketplaces, dados);
  const itens = MARKETPLACES.map(m => {
    const v = dados.porMarketplace[m.nome];
    return '<div class="item"><div><div class="item-nome">' + m.nome + '</div>' +
      '<div class="item-sub">' + m.legenda + '</div></div>' +
      '<div class="item-dir"><div class="item-valor">' + brl(v) + '</div>' +
      '<div class="item-pct">' + pct(total ? v / total : 0) + ' do total</div></div></div>';
  }).join('');
  return paginaBase(dados, 'META MARKETPLACES', 'FATURAMENTO LÍQUIDO ACUMULADO', 'Meta do mês', r,
    'Resultado por canal', itens, 'líquido = Total Venda - Frete');
}

function paginaGrupo(dados) {
  const total = GRUPOS.reduce((s, g) => s + dados.porGrupo[g.nome], 0);
  const metaTotal = GRUPOS.reduce((s, g) => s + dados.metas[g.nome], 0);
  const r = resumoMeta(total, metaTotal, dados);
  const itens = GRUPOS.map(g => {
    const v = dados.porGrupo[g.nome];
    const meta = dados.metas[g.nome];
    const p = meta ? v / meta : 0;
    return '<div class="item item-meta"><div class="item-linha"><div><div class="item-nome">' + g.nome + '</div>' +
      '<div class="item-sub">Meta: ' + brl(meta) + '</div></div>' +
      '<div class="item-dir"><div class="item-valor">' + brl(v) + '</div>' +
      '<div class="item-pct">' + pct(p) + '</div></div></div>' +
      '<div class="barra fina"><div class="barra-fill ' + (p >= 0.8 ? 'verde' : 'rosa') + '" style="width:' +
      Math.min(p * 100, 100).toFixed(1) + '%"></div></div></div>';
  }).join('');
  return paginaBase(dados, 'META GRUPO', 'FATURAMENTO ACUMULADO DO GRUPO', 'Meta total', r,
    'Acompanhamento por meta', itens, 'base: Total Venda');
}

function paginaBase(dados, titulo, rotulo, rotuloMeta, r, secao, itens, rodape) {
  return '<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">' +
    '<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&family=Playfair+Display:wght@600;700;800&display=swap" rel="stylesheet">' +
    '<style>' + CSS_DASHBOARD + '</style></head><body><div class="tela">' +
    '<header><div class="topo"><div class="logo">AMOUH</div><div class="data">' + dados.data +
    (dados.hora ? '<span class="hora">atualizado ' + dados.hora + '</span>' : '') + '</div></div>' +
    '<div class="subtitulo">' + titulo + ' · ' + dados.mes + '</div></header>' +
    '<main>' +
    '<section class="card-principal"><div class="rotulo">' + rotulo + '</div>' +
    '<div class="linha-total"><div class="total">' + brl(r.total) + '</div><div class="pct-total">' + pct(r.pct) + '</div></div>' +
    '<div class="meta">' + rotuloMeta + ': <b>' + brl(r.meta) + '</b></div>' +
    '<div class="barra"><div class="barra-fill verde" style="width:' + Math.min(r.pct * 100, 100).toFixed(1) + '%"></div></div></section>' +
    '<section class="cards">' +
    '<div class="mini azul"><div class="mini-rotulo">FALTA PARA A META</div><div class="mini-valor">' + brl(r.falta) + '</div></div>' +
    '<div class="mini dourado"><div class="mini-rotulo">PROJEÇÃO MÊS</div><div class="mini-valor">' + brl(r.projecao) + '</div></div>' +
    '<div class="mini rosa"><div class="mini-rotulo">PROJEÇÃO DA META</div><div class="mini-valor">' + pct(r.projecaoPct) + '</div></div>' +
    '</section>' +
    '<h2>' + secao + '</h2>' + itens +
    (dados.pendentes ? '<div class="aviso">Carregando pedidos do Bling (' + dados.pendentes + ' restantes). Números ainda parciais.</div>' : '') +
    '</main>' +
    '<footer><div><b>AMOUH</b> · Dashboard comercial</div><div>' + dados.dia + ' dias registrados · ' + rodape + '</div></footer>' +
    '</div></body></html>';
}

const CSS_DASHBOARD = [
  ':root{--azul:#132C45;--rosa:#F2C6CE;--verde:#2A5135;--dourado:#C9A84C;--fundo:#F6F4EF;--borda:#E8E3DA;--cinza:#6F7780;--trilho:#ECE9E2}',
  '*{box-sizing:border-box;margin:0;padding:0}',
  'body{background:var(--fundo);font-family:Montserrat,sans-serif;color:var(--azul)}',
  '.tela{max-width:390px;min-height:692px;margin:0 auto;background:var(--fundo);display:flex;flex-direction:column}',
  'header{background:linear-gradient(180deg,#1a3552 0%,var(--azul) 35%);padding:30px 26px 52px}',
  '.topo{display:flex;justify-content:space-between;align-items:flex-start}',
  '.logo{font-family:"Playfair Display",serif;font-weight:700;font-size:40px;color:#fff;letter-spacing:1px;line-height:1}',
  '.data{font-size:12px;color:rgba(255,255,255,.75);text-align:right;padding-top:6px}',
  '.hora{display:block;font-size:9px;opacity:.7;margin-top:3px}',
  '.subtitulo{margin-top:12px;font-size:11px;font-weight:600;letter-spacing:2.5px;color:var(--dourado)}',
  'main{padding:0 16px;flex:1}',
  '.card-principal{background:#fff;border-radius:22px;padding:20px 22px 18px;margin-top:14px}',
  '.rotulo{font-size:9.5px;font-weight:600;letter-spacing:2px;color:var(--cinza)}',
  '.linha-total{display:flex;justify-content:space-between;align-items:center;margin-top:6px}',
  '.total{font-family:"Playfair Display",serif;font-weight:700;font-size:32px;color:var(--azul)}',
  '.pct-total{font-size:20px;font-weight:700;color:var(--verde)}',
  '.meta{font-size:11.5px;color:var(--cinza);margin-top:2px}.meta b{font-weight:600}',
  '.barra{height:9px;background:var(--trilho);border-radius:9px;margin-top:12px;overflow:hidden}',
  '.barra.fina{height:6px;margin-top:8px}',
  '.barra-fill{height:100%;border-radius:9px}.barra-fill.verde{background:var(--verde)}.barra-fill.rosa{background:var(--rosa)}',
  '.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-top:10px}',
  '.mini{border-radius:14px;padding:12px 10px 18px}',
  '.mini.azul{background:var(--azul);color:#fff}.mini.dourado{background:var(--dourado);color:var(--azul)}.mini.rosa{background:var(--rosa);color:var(--azul)}',
  '.mini-rotulo{font-size:7.5px;font-weight:700;letter-spacing:1.2px}',
  '.mini-valor{font-family:"Playfair Display",serif;font-weight:700;font-size:15px;margin-top:4px;white-space:nowrap}',
  'h2{font-family:"Playfair Display",serif;font-weight:700;font-size:19px;margin:18px 0 8px}',
  '.item{background:#fff;border:1.5px solid var(--borda);border-radius:14px;padding:12px 16px;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center}',
  '.item-meta{display:block}.item-linha{display:flex;justify-content:space-between;align-items:center}',
  '.item-nome{font-size:12.5px;font-weight:700}',
  '.item-sub{font-size:10px;color:var(--cinza);margin-top:2px}',
  '.item-dir{text-align:right}',
  '.item-valor{font-family:"Playfair Display",serif;font-weight:700;font-size:18px}',
  '.item-pct{font-size:9.5px;font-weight:700;color:var(--verde);margin-top:1px}',
  '.aviso{font-size:10px;color:var(--cinza);text-align:center;margin:6px 0}',
  'footer{display:flex;justify-content:space-between;border-top:1px solid #DDD8CE;margin:16px 16px 0;padding:10px 0 14px;font-size:9px;color:var(--cinza)}',
  'footer b{color:var(--azul)}',
].join('');

function brl(v) {
  const [inteiro, dec] = Math.abs(v).toFixed(2).split('.');
  return (v < 0 ? '-' : '') + 'R$ ' + inteiro.replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + dec;
}

function pct(v) { return (v * 100).toFixed(2).replace('.', ',') + '%'; }

// ===== Utilitários =====

function aba(nome, cabecalho) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let a = ss.getSheetByName(nome);
  if (!a) {
    a = ss.insertSheet(nome);
    if (cabecalho) {
      a.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight('bold');
      a.getRange('B:B').setNumberFormat('@'); // datas como texto AAAA-MM-DD
    }
  }
  return a;
}

function normalizar(texto) {
  return String(texto).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function produtoExcluido(nome) {
  const n = normalizar(nome);
  return TERMOS_EXCLUIDOS.some(t => n.indexOf(t) !== -1);
}

function addDias(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function fmt(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function br(s) { return s.split('-').reverse().join('/'); }

/**
 * Dashboard de vendas Bling -> Google Planilhas
 *
 * Mostra, para cada canal, os 5 produtos (pai) mais vendidos em quantidade
 * nos últimos 7 dias até ontem, comparando com os 7 dias anteriores.
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

const TOP_N = 5;
const LIMITE_EXECUCAO_MS = 5 * 60 * 1000; // Apps Script corta em 6 min

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
  ScriptApp.newTrigger('atualizar').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Pronto! O dashboard vai se atualizar sozinho a cada hora.');
}

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

function apiGet(caminho) {
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
    if (code === 404) return null;
    throw new Error('Bling ' + code + ' em ' + caminho + ': ' + resp.getContentText());
  }
  throw new Error('Bling não respondeu após várias tentativas: ' + caminho);
}

// ===== Atualização =====

function atualizar() {
  const inicio = Date.now();
  const hoje = new Date();
  const periodoIni = fmt(addDias(hoje, -14));
  const periodoFim = fmt(addDias(hoje, -1));

  // 1. Lista os pedidos dos últimos 14 dias (a listagem já traz loja e situação)
  const pedidos = {};
  for (let pagina = 1; ; pagina++) {
    const r = apiGet('/pedidos/vendas?limite=100&pagina=' + pagina +
      '&dataInicial=' + periodoIni + '&dataFinal=' + periodoFim);
    if (!r || !r.data || r.data.length === 0) break;
    for (const p of r.data) {
      const lojaId = p.loja ? p.loja.id : 0;
      if (!CANAIS[lojaId]) continue;
      if (SITUACOES_EXCLUIDAS.indexOf(p.situacao.id) !== -1) continue;
      pedidos[p.id] = { data: p.data, canal: CANAIS[lojaId] };
    }
  }

  // 2. Busca os itens dos pedidos ainda não salvos (continua na próxima execução se faltar tempo)
  const abaItens = aba('Itens', ['pedidoId', 'data', 'canal', 'produtoPaiId', 'produto', 'quantidade']);
  const salvos = new Set(abaItens.getLastRow() > 1
    ? abaItens.getRange(2, 1, abaItens.getLastRow() - 1, 1).getValues().map(l => String(l[0]))
    : []);
  const cacheProdutos = carregarCacheProdutos();
  const novasLinhas = [];
  let pendentes = 0;

  for (const id of Object.keys(pedidos)) {
    if (salvos.has(id)) continue;
    if (Date.now() - inicio > LIMITE_EXECUCAO_MS) { pendentes++; continue; }
    const det = apiGet('/pedidos/vendas/' + id);
    if (!det || !det.data) continue;
    for (const item of det.data.itens || []) {
      const pai = produtoPai(item, cacheProdutos);
      novasLinhas.push([id, pedidos[id].data, pedidos[id].canal, pai.id, pai.nome, Number(item.quantidade) || 0]);
    }
  }
  if (novasLinhas.length) {
    abaItens.getRange(abaItens.getLastRow() + 1, 1, novasLinhas.length, 6).setValues(novasLinhas);
  }
  salvarCacheProdutos(cacheProdutos);

  // 3. Monta o dashboard só com pedidos válidos no momento (se um pedido for cancelado depois, some da conta)
  montarDashboard(abaItens, pedidos, pendentes);
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
  const atualIni = fmt(addDias(hoje, -7)), atualFim = fmt(addDias(hoje, -1));
  const antIni = fmt(addDias(hoje, -14)), antFim = fmt(addDias(hoje, -8));

  // soma[canal][produtoPaiId] = { nome, atual, anterior }
  const soma = {};
  ORDEM_CANAIS.forEach(c => { soma[c] = {}; });
  if (abaItens.getLastRow() > 1) {
    for (const [pedidoId, dataRaw, canal, paiId, nome, qtd] of
      abaItens.getRange(2, 1, abaItens.getLastRow() - 1, 6).getValues()) {
      if (!pedidosValidos[pedidoId] || !soma[canal]) continue;
      const data = dataRaw instanceof Date ? fmt(dataRaw) : String(dataRaw);
      const reg = soma[canal][paiId] || (soma[canal][paiId] = { nome: nome, atual: 0, anterior: 0 });
      if (data >= atualIni && data <= atualFim) reg.atual += qtd;
      else if (data >= antIni && data <= antFim) reg.anterior += qtd;
    }
  }

  const d = aba('Dashboard');
  d.clear();
  d.getRange('A1').setValue('Top ' + TOP_N + ' produtos por canal (unidades vendidas)')
    .setFontSize(16).setFontWeight('bold');
  d.getRange('A2').setValue(
    'Últimos 7 dias: ' + br(atualIni) + ' a ' + br(atualFim) +
    '   |   Semana anterior: ' + br(antIni) + ' a ' + br(antFim) +
    '   |   Atualizado em ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') +
    (pendentes ? '   |   ⚠ ' + pendentes + ' pedidos ainda sendo carregados' : ''))
    .setFontColor('#666666');

  let linha = 4;
  for (const canal of ORDEM_CANAIS) {
    const lista = Object.values(soma[canal])
      .filter(r => r.atual > 0)
      .sort((a, b) => b.atual - a.atual)
      .slice(0, TOP_N);

    d.getRange(linha, 1).setValue(canal).setFontSize(13).setFontWeight('bold');
    linha++;
    d.getRange(linha, 1, 1, 5)
      .setValues([['#', 'Produto', 'Últimos 7 dias', 'Semana anterior', 'Variação']])
      .setFontWeight('bold').setBackground('#efefef');
    linha++;

    if (!lista.length) {
      d.getRange(linha, 2).setValue('Sem vendas no período').setFontColor('#999999');
      linha += 2;
      continue;
    }
    const valores = lista.map((r, i) => [
      i + 1, r.nome, r.atual, r.anterior,
      r.anterior ? (r.atual - r.anterior) / r.anterior : 'novo',
    ]);
    d.getRange(linha, 1, valores.length, 5).setValues(valores);
    d.getRange(linha, 5, valores.length, 1).setNumberFormat('+0%;-0%;0%');
    valores.forEach((v, i) => {
      if (typeof v[4] === 'number') {
        d.getRange(linha + i, 5).setFontColor(v[4] > 0 ? '#188038' : v[4] < 0 ? '#d93025' : '#000000');
      }
    });
    linha += valores.length + 1;
  }
  d.setColumnWidth(1, 40);
  d.setColumnWidth(2, 380);
  d.setColumnWidths(3, 3, 130);
  d.getRange(4, 3, Math.max(linha - 4, 1), 3).setHorizontalAlignment('center');

  limparItensAntigos(abaItens);
}

function limparItensAntigos(abaItens) {
  if (abaItens.getLastRow() < 2) return;
  const limite = fmt(addDias(new Date(), -30));
  const linhas = abaItens.getRange(2, 1, abaItens.getLastRow() - 1, 6).getValues();
  const manter = linhas.filter(l => (l[1] instanceof Date ? fmt(l[1]) : String(l[1])) >= limite);
  if (manter.length === linhas.length) return;
  abaItens.getRange(2, 1, linhas.length, 6).clearContent();
  if (manter.length) abaItens.getRange(2, 1, manter.length, 6).setValues(manter);
}

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

function addDias(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function fmt(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function br(s) { return s.split('-').reverse().join('/'); }

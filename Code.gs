/**
 * SISTEMA PDV TOP FASHION
 * Banco de dados: esta planilha do Google Sheets.
 * Tela: arquivo Index.html (publicado como App da Web).
 *
 * Primeira vez: rode a função "configurar" uma vez (menu Executar).
 */

var SCHEMA = {
  Config:     ['Chave', 'Valor'],
  Produtos:   ['SKU', 'Referencia', 'Nome', 'Categoria', 'Cor', 'Tamanho', 'Custo', 'Preco', 'Estoque', 'EstoqueMinimo', 'Ativo', 'CriadoEm'],
  Movimentos: ['Data', 'Tipo', 'SKU', 'Produto', 'Qtd', 'CustoUnit', 'Documento', 'Usuario', 'Obs'],
  Vendas:     ['ID', 'Data', 'Cliente', 'ClienteID', 'Qtd', 'Subtotal', 'Desconto', 'Total', 'Pagamentos', 'Custo', 'Lucro', 'Status', 'Usuario', 'Obs'],
  VendaItens: ['VendaID', 'Data', 'SKU', 'Produto', 'Qtd', 'Preco', 'Custo', 'Total'],
  Financeiro: ['ID', 'Data', 'Tipo', 'Categoria', 'Descricao', 'Valor', 'Forma', 'Referencia', 'Usuario'],
  Clientes:   ['ID', 'Nome', 'Telefone', 'Obs', 'CriadoEm'],
  Crediario:  ['ID', 'Data', 'ClienteID', 'Cliente', 'VendaID', 'Valor', 'Pago', 'Saldo', 'Vencimento', 'Status']
};

var CONFIG_PADRAO = [
  ['NOME_LOJA', 'TOP FASHION MS'],
  ['PIN_DONO', '1234'],
  ['PIN_CAIXA', '4321'],
  ['ETQ_LARGURA_MM', '40'],
  ['ETQ_ALTURA_MM', '25'],
  ['ETQ_COLUNAS', '1'],
  ['ETQ_ESPACO_MM', '2'],
  ['ETQ_MOSTRAR_PRECO', 'SIM'],
  ['PROX_CODIGO', '10000001'],
  ['PROX_REF', '1'],
  ['PROX_VENDA', '1'],
  ['PROX_FIN', '1'],
  ['PROX_CLIENTE', '1'],
  ['PROX_CRED', '1']
];

var COLUNAS_TEXTO = { Config: ['Valor'], Produtos: ['SKU', 'Referencia'], Movimentos: ['SKU'], VendaItens: ['SKU'] };

/* ======================= INSTALAÇÃO ======================= */

function configurar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (nome) {
    var s = ss.getSheetByName(nome);
    if (!s) s = ss.insertSheet(nome);
    if (s.getLastRow() === 0) s.appendRow(SCHEMA[nome]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, SCHEMA[nome].length).setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
    (COLUNAS_TEXTO[nome] || []).forEach(function (col) {
      var c = SCHEMA[nome].indexOf(col) + 1;
      s.getRange(1, c, s.getMaxRows(), 1).setNumberFormat('@');
    });
  });
  var atuais = {};
  rows_('Config').forEach(function (r) { atuais[r.Chave] = true; });
  CONFIG_PADRAO.forEach(function (kv) { if (!atuais[kv[0]]) append_('Config', { Chave: kv[0], Valor: kv[1] }); });
  var padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 1 && padrao.getLastRow() === 0) ss.deleteSheet(padrao);
  return 'Pronto! Planilha configurada.';
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Top Fashion PDV')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/**
 * Porta de entrada para o site no GitHub Pages.
 * Recebe {fn, args} em JSON e devolve {ok, data} ou {ok:false, erro}.
 */
function doPost(e) {
  var API = {
    apiLogin: apiLogin, apiDados: apiDados, apiCriarProduto: apiCriarProduto, apiEntrada: apiEntrada,
    apiEditarReferencia: apiEditarReferencia, apiAdicionarVariacoes: apiAdicionarVariacoes, apiAjustarEstoque: apiAjustarEstoque,
    apiHistorico: apiHistorico, apiVender: apiVender, apiVendas: apiVendas, apiCancelarVenda: apiCancelarVenda,
    apiLancar: apiLancar, apiExcluirLancamento: apiExcluirLancamento, apiFinanceiro: apiFinanceiro, apiResumo: apiResumo,
    apiSalvarCliente: apiSalvarCliente, apiCrediario: apiCrediario, apiReceberCrediario: apiReceberCrediario, apiSalvarConfig: apiSalvarConfig
  };
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!API.hasOwnProperty(req.fn)) throw new Error('Função inválida.');
    out = { ok: true, data: API[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, erro: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ======================= UTILITÁRIOS ======================= */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sh_(nome) {
  var s = ss_().getSheetByName(nome);
  if (!s) { configurar(); s = ss_().getSheetByName(nome); }
  return s;
}

function rows_(nome) {
  var v = sh_(nome).getDataRange().getValues();
  var h = v.shift() || [];
  var out = [];
  for (var i = 0; i < v.length; i++) {
    if (v[i].join('') === '') continue;
    var o = { _row: i + 2 };
    for (var j = 0; j < h.length; j++) o[h[j]] = v[i][j];
    out.push(o);
  }
  return out;
}

function append_(nome, obj) { appendMany_(nome, [obj]); }

function appendMany_(nome, objs) {
  if (!objs.length) return;
  var s = sh_(nome), h = SCHEMA[nome];
  var data = objs.map(function (o) { return h.map(function (k) { return o[k] === undefined || o[k] === null ? '' : o[k]; }); });
  s.getRange(s.getLastRow() + 1, 1, data.length, h.length).setValues(data);
}

function set_(nome, row, col, val) {
  sh_(nome).getRange(row, SCHEMA[nome].indexOf(col) + 1).setValue(val);
}

function tz_() { return Session.getScriptTimeZone() || 'America/Campo_Grande'; }
function ymd_(d) { return d instanceof Date ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : String(d || '').slice(0, 10); }
function dataHora_(d) { return d instanceof Date ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd HH:mm') : String(d || ''); }
function num_(v) { var n = Number(String(v === undefined || v === null ? '' : v).replace(',', '.')); return isNaN(n) ? 0 : n; }
function r2_(n) { return Math.round(n * 100) / 100; }
function txt_(v) { return String(v === undefined || v === null ? '' : v).trim(); }
function noIntervalo_(d, de, ate) { var x = ymd_(d); return (!de || x >= de) && (!ate || x <= ate); }
function hoje_() { return ymd_(new Date()); }

function comTrava_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function cfg_() {
  var c = {};
  rows_('Config').forEach(function (r) { c[r.Chave] = txt_(r.Valor); });
  return c;
}

function setCfg_(chave, valor) {
  var r = rows_('Config').filter(function (x) { return x.Chave === chave; })[0];
  if (r) set_('Config', r._row, 'Valor', String(valor));
  else append_('Config', { Chave: chave, Valor: String(valor) });
}

function contador_(chave, prefixo, digitos) {
  var c = cfg_();
  var n = parseInt(c[chave] || '1', 10);
  setCfg_(chave, n + 1);
  if (!prefixo && !digitos) return String(n);
  var s = String(n);
  while (s.length < digitos) s = '0' + s;
  return prefixo + s;
}

/** Reserva n números seguidos de um contador e devolve o primeiro. */
function contadorBloco_(chave, n) {
  var c = cfg_();
  var ini = parseInt(c[chave] || '1', 10);
  setCfg_(chave, ini + n);
  return ini;
}

function auth_(pin, precisaDono) {
  // Proteção contra quem tenta adivinhar o PIN: 10 erros em 15 minutos bloqueiam por 15 minutos.
  var cache = CacheService.getScriptCache();
  var falhas = parseInt(cache.get('falhas_pin') || '0', 10);
  if (falhas >= 10) throw new Error('Muitas tentativas erradas. Aguarde 15 minutos.');
  var c = cfg_(), p = txt_(pin), papel = null;
  if (p && p === c.PIN_DONO) papel = 'dono';
  else if (p && p === c.PIN_CAIXA) papel = 'caixa';
  if (!papel) { cache.put('falhas_pin', String(falhas + 1), 900); throw new Error('PIN inválido.'); }
  if (precisaDono && papel !== 'dono') throw new Error('Somente o dono pode fazer isso.');
  return papel;
}

function nomeProduto_(p) {
  return [p.Nome, p.Cor, p.Tamanho].filter(function (x) { return txt_(x); }).join(' ');
}

function produtoOut_(p, papel) {
  return {
    sku: String(p.SKU), ref: String(p.Referencia), nome: txt_(p.Nome), categoria: txt_(p.Categoria),
    cor: txt_(p.Cor), tamanho: txt_(p.Tamanho), custo: papel === 'dono' ? num_(p.Custo) : null,
    preco: num_(p.Preco), estoque: num_(p.Estoque), minimo: num_(p.EstoqueMinimo),
    ativo: txt_(p.Ativo) !== 'NAO'
  };
}

function mapaProdutos_() {
  var m = {};
  rows_('Produtos').forEach(function (p) { m[String(p.SKU)] = p; });
  return m;
}

function movEstoque_(p, qtd, tipo, doc, usuario, obs, custoUnit) {
  var novo = num_(p.Estoque) + qtd;
  set_('Produtos', p._row, 'Estoque', novo);
  p.Estoque = novo;
  return {
    Data: new Date(), Tipo: tipo, SKU: String(p.SKU), Produto: nomeProduto_(p), Qtd: qtd,
    CustoUnit: custoUnit === undefined ? num_(p.Custo) : custoUnit, Documento: doc || '', Usuario: usuario, Obs: obs || ''
  };
}

/* ======================= LOGIN / DADOS ======================= */

function apiLogin(pin) { return dados_(auth_(pin)); }

function apiDados(pin) { return dados_(auth_(pin)); }

function dados_(papel) {
  var c = cfg_();
  var config = {
    loja: c.NOME_LOJA, etqLargura: num_(c.ETQ_LARGURA_MM) || 40, etqAltura: num_(c.ETQ_ALTURA_MM) || 25,
    etqColunas: num_(c.ETQ_COLUNAS) || 1, etqEspaco: num_(c.ETQ_ESPACO_MM), etqPreco: c.ETQ_MOSTRAR_PRECO !== 'NAO'
  };
  if (papel === 'dono') { config.pinDono = c.PIN_DONO; config.pinCaixa = c.PIN_CAIXA; }
  var produtos = rows_('Produtos').map(function (p) { return produtoOut_(p, papel); });
  var clientes = rows_('Clientes').map(function (r) { return { id: String(r.ID), nome: txt_(r.Nome), telefone: txt_(r.Telefone), obs: txt_(r.Obs) }; });
  return { papel: papel, config: config, produtos: produtos, clientes: clientes };
}

/* ======================= PRODUTOS / ESTOQUE ======================= */

/** Cadastra um produto novo com grade de cores x tamanhos e já dá entrada no estoque. */
function apiCriarProduto(pin, d) {
  var papel = auth_(pin);
  return comTrava_(function () {
    var nome = txt_(d.nome);
    if (!nome) throw new Error('Informe o nome do produto.');
    var custo = r2_(num_(d.custo)), preco = r2_(num_(d.preco));
    if (preco <= 0) throw new Error('Informe o preço de venda.');
    var cores = (d.cores && d.cores.length) ? d.cores : [''];
    var tams = (d.tamanhos && d.tamanhos.length) ? d.tamanhos : [''];
    var ref = contador_('PROX_REF', 'R', 4);
    var prox = contadorBloco_('PROX_CODIGO', cores.length * tams.length);
    var agora = new Date(), novos = [], movs = [], entradaTotal = 0, criados = [];
    cores.forEach(function (cor) {
      tams.forEach(function (tam) {
        var qtd = Math.max(0, Math.round(num_((d.qtd || {})[cor + '|' + tam])));
        var sku = String(prox++);
        var p = {
          SKU: sku, Referencia: ref, Nome: nome, Categoria: txt_(d.categoria), Cor: txt_(cor), Tamanho: txt_(tam),
          Custo: custo, Preco: preco, Estoque: qtd, EstoqueMinimo: Math.max(0, num_(d.minimo)), Ativo: 'SIM', CriadoEm: agora
        };
        novos.push(p);
        criados.push({ sku: sku, qtd: qtd });
        if (qtd > 0) {
          entradaTotal += qtd * custo;
          movs.push({ Data: agora, Tipo: 'ENTRADA', SKU: sku, Produto: nomeProduto_(p), Qtd: qtd, CustoUnit: custo, Documento: txt_(d.fornecedor), Usuario: papel, Obs: 'Cadastro' });
        }
      });
    });
    appendMany_('Produtos', novos);
    appendMany_('Movimentos', movs);
    if (d.lancarGasto && entradaTotal > 0) {
      append_('Financeiro', {
        ID: contador_('PROX_FIN', 'F', 5), Data: agora, Tipo: 'SAIDA', Categoria: 'Compra de mercadoria',
        Descricao: nome + (d.fornecedor ? ' - ' + txt_(d.fornecedor) : ''), Valor: r2_(entradaTotal),
        Forma: txt_(d.forma) || 'Pix', Referencia: ref, Usuario: papel
      });
    }
    return { ref: ref, criados: criados };
  });
}

/** Entrada de mercadoria de produtos já cadastrados. itens: [{sku, qtd, custo}] */
function apiEntrada(pin, d) {
  var papel = auth_(pin);
  return comTrava_(function () {
    var mapa = mapaProdutos_(), movs = [], total = 0;
    (d.itens || []).forEach(function (it) {
      var p = mapa[String(it.sku)];
      var q = Math.round(num_(it.qtd));
      if (!p) throw new Error('Produto não encontrado: ' + it.sku);
      if (q <= 0) return;
      var custo = (papel === 'dono' && it.custo !== '' && it.custo !== null && it.custo !== undefined) ? r2_(num_(it.custo)) : num_(p.Custo);
      if (papel === 'dono' && custo !== num_(p.Custo)) { set_('Produtos', p._row, 'Custo', custo); p.Custo = custo; }
      total += q * custo;
      movs.push(movEstoque_(p, q, 'ENTRADA', txt_(d.fornecedor), papel, txt_(d.obs), custo));
    });
    if (!movs.length) throw new Error('Nenhuma quantidade informada.');
    appendMany_('Movimentos', movs);
    if (d.lancarGasto && total > 0) {
      append_('Financeiro', {
        ID: contador_('PROX_FIN', 'F', 5), Data: new Date(), Tipo: 'SAIDA', Categoria: 'Compra de mercadoria',
        Descricao: 'Reposição' + (d.fornecedor ? ' - ' + txt_(d.fornecedor) : ''), Valor: r2_(total),
        Forma: txt_(d.forma) || 'Pix', Referencia: '', Usuario: papel
      });
    }
    return { itens: movs.length, total: r2_(total) };
  });
}

/** Edita dados de uma referência inteira (nome, categoria, preço, custo, mínimo) */
function apiEditarReferencia(pin, ref, d) {
  var papel = auth_(pin, true);
  return comTrava_(function () {
    var n = 0;
    rows_('Produtos').forEach(function (p) {
      if (String(p.Referencia) !== String(ref)) return;
      if (d.nome !== undefined) set_('Produtos', p._row, 'Nome', txt_(d.nome));
      if (d.categoria !== undefined) set_('Produtos', p._row, 'Categoria', txt_(d.categoria));
      if (d.preco !== undefined && d.preco !== '') set_('Produtos', p._row, 'Preco', r2_(num_(d.preco)));
      if (d.custo !== undefined && d.custo !== '') set_('Produtos', p._row, 'Custo', r2_(num_(d.custo)));
      if (d.minimo !== undefined && d.minimo !== '') set_('Produtos', p._row, 'EstoqueMinimo', num_(d.minimo));
      if (d.ativo !== undefined) set_('Produtos', p._row, 'Ativo', d.ativo ? 'SIM' : 'NAO');
      n++;
    });
    return { alterados: n, papel: papel };
  });
}

/** Adiciona novas variações (cor/tamanho) a uma referência existente */
function apiAdicionarVariacoes(pin, ref, d) {
  var papel = auth_(pin);
  return comTrava_(function () {
    var base = rows_('Produtos').filter(function (p) { return String(p.Referencia) === String(ref); });
    if (!base.length) throw new Error('Referência não encontrada.');
    var b = base[0], agora = new Date(), novos = [], movs = [], criados = [];
    var existe = {};
    base.forEach(function (p) { existe[txt_(p.Cor) + '|' + txt_(p.Tamanho)] = true; });
    var validas = (d.variacoes || []).filter(function (v) {
      var chave = txt_(v.cor) + '|' + txt_(v.tamanho);
      if (existe[chave]) return false;
      existe[chave] = true;
      return true;
    });
    if (!validas.length) throw new Error('Essas variações já existem.');
    var prox = contadorBloco_('PROX_CODIGO', validas.length);
    validas.forEach(function (v) {
      var q = Math.max(0, Math.round(num_(v.qtd)));
      var sku = String(prox++);
      var p = { SKU: sku, Referencia: String(ref), Nome: b.Nome, Categoria: b.Categoria, Cor: txt_(v.cor), Tamanho: txt_(v.tamanho),
        Custo: num_(b.Custo), Preco: num_(b.Preco), Estoque: q, EstoqueMinimo: num_(b.EstoqueMinimo), Ativo: 'SIM', CriadoEm: agora };
      novos.push(p); criados.push({ sku: sku, qtd: q });
      if (q > 0) movs.push({ Data: agora, Tipo: 'ENTRADA', SKU: sku, Produto: nomeProduto_(p), Qtd: q, CustoUnit: num_(b.Custo), Documento: '', Usuario: papel, Obs: 'Nova variação' });
    });
    appendMany_('Produtos', novos);
    appendMany_('Movimentos', movs);
    return { criados: criados };
  });
}

/** Acerto de estoque (contagem). Define a quantidade real. */
function apiAjustarEstoque(pin, sku, novaQtd, motivo) {
  var papel = auth_(pin, true);
  return comTrava_(function () {
    var p = mapaProdutos_()[String(sku)];
    if (!p) throw new Error('Produto não encontrado.');
    var dif = Math.round(num_(novaQtd)) - num_(p.Estoque);
    if (dif === 0) return { ok: true };
    append_('Movimentos', movEstoque_(p, dif, 'AJUSTE', '', papel, txt_(motivo) || 'Contagem'));
    return { ok: true, diferenca: dif };
  });
}

function apiHistorico(pin, sku) {
  auth_(pin);
  return rows_('Movimentos').filter(function (m) { return String(m.SKU) === String(sku); })
    .slice(-40).reverse().map(function (m) {
      return { data: dataHora_(m.Data), tipo: m.Tipo, qtd: num_(m.Qtd), doc: txt_(m.Documento), obs: txt_(m.Obs) };
    });
}

/* ======================= VENDAS ======================= */

/**
 * v = { itens:[{sku,qtd,preco}], desconto, pagamentos:[{forma, valor}], clienteId, vencimento, obs }
 * Forma "Crediário" gera conta a receber para o cliente.
 */
function apiVender(pin, v) {
  var papel = auth_(pin);
  return comTrava_(function () {
    if (!v.itens || !v.itens.length) throw new Error('Carrinho vazio.');
    var mapa = mapaProdutos_(), agora = new Date();
    var subtotal = 0, custo = 0, qtd = 0, linhas = [];
    v.itens.forEach(function (it) {
      var p = mapa[String(it.sku)];
      if (!p) throw new Error('Produto não encontrado: ' + it.sku);
      var q = Math.round(num_(it.qtd));
      if (q <= 0) return;
      var preco = (it.preco === undefined || it.preco === null || it.preco === '') ? num_(p.Preco) : r2_(num_(it.preco));
      subtotal += q * preco; custo += q * num_(p.Custo); qtd += q;
      linhas.push({ p: p, q: q, preco: preco });
    });
    if (!linhas.length) throw new Error('Carrinho vazio.');
    subtotal = r2_(subtotal); custo = r2_(custo);
    var desconto = r2_(Math.max(0, num_(v.desconto)));
    var total = r2_(subtotal - desconto);
    if (total < 0) throw new Error('Desconto maior que o total.');

    var pags = (v.pagamentos || []).filter(function (x) { return num_(x.valor) > 0; })
      .map(function (x) { return { forma: txt_(x.forma), valor: r2_(num_(x.valor)) }; });
    var soma = r2_(pags.reduce(function (a, x) { return a + x.valor; }, 0));
    if (Math.abs(soma - total) > 0.01) throw new Error('A soma dos pagamentos (' + soma.toFixed(2) + ') não bate com o total (' + total.toFixed(2) + ').');

    var cliente = null;
    if (v.clienteId) cliente = rows_('Clientes').filter(function (c) { return String(c.ID) === String(v.clienteId); })[0] || null;
    var temCred = pags.some(function (x) { return x.forma === 'Crediário'; });
    if (temCred && !cliente) throw new Error('Escolha o cliente para vender no crediário.');

    var id = contador_('PROX_VENDA', 'V', 5);
    var itensOut = [], movs = [], vi = [];
    linhas.forEach(function (l) {
      var nome = nomeProduto_(l.p);
      vi.push({ VendaID: id, Data: agora, SKU: String(l.p.SKU), Produto: nome, Qtd: l.q, Preco: l.preco, Custo: num_(l.p.Custo), Total: r2_(l.q * l.preco) });
      movs.push(movEstoque_(l.p, -l.q, 'VENDA', id, papel, ''));
      itensOut.push({ nome: nome, qtd: l.q, preco: l.preco, total: r2_(l.q * l.preco), estoque: num_(l.p.Estoque) });
    });

    append_('Vendas', {
      ID: id, Data: agora, Cliente: cliente ? txt_(cliente.Nome) : '', ClienteID: cliente ? String(cliente.ID) : '', Qtd: qtd,
      Subtotal: subtotal, Desconto: desconto, Total: total,
      Pagamentos: pags.map(function (x) { return x.forma + '=' + x.valor.toFixed(2); }).join('; '),
      Custo: custo, Lucro: r2_(total - custo), Status: 'OK', Usuario: papel, Obs: txt_(v.obs)
    });
    appendMany_('VendaItens', vi);
    appendMany_('Movimentos', movs);

    var fin = [];
    pags.forEach(function (x) {
      if (x.forma === 'Crediário') {
        append_('Crediario', {
          ID: contador_('PROX_CRED', 'C', 5), Data: agora, ClienteID: String(cliente.ID), Cliente: txt_(cliente.Nome), VendaID: id,
          Valor: x.valor, Pago: 0, Saldo: x.valor, Vencimento: txt_(v.vencimento), Status: 'ABERTO'
        });
      } else {
        fin.push({ ID: contador_('PROX_FIN', 'F', 5), Data: agora, Tipo: 'ENTRADA', Categoria: 'Venda', Descricao: 'Venda ' + id,
          Valor: x.valor, Forma: x.forma, Referencia: id, Usuario: papel });
      }
    });
    appendMany_('Financeiro', fin);

    return { id: id, data: dataHora_(agora), total: total, subtotal: subtotal, desconto: desconto, pagamentos: pags,
      itens: itensOut, cliente: cliente ? { nome: txt_(cliente.Nome), telefone: txt_(cliente.Telefone) } : null };
  });
}

function apiVendas(pin, de, ate) {
  var papel = auth_(pin);
  if (papel !== 'dono') { de = hoje_(); ate = hoje_(); }
  var itens = {};
  rows_('VendaItens').forEach(function (i) {
    if (!noIntervalo_(i.Data, de, ate)) return;
    (itens[i.VendaID] = itens[i.VendaID] || []).push({ nome: txt_(i.Produto), qtd: num_(i.Qtd), preco: num_(i.Preco), total: num_(i.Total) });
  });
  return rows_('Vendas').filter(function (v) { return noIntervalo_(v.Data, de, ate); }).reverse().map(function (v) {
    return { id: String(v.ID), data: dataHora_(v.Data), cliente: txt_(v.Cliente), qtd: num_(v.Qtd), subtotal: num_(v.Subtotal),
      desconto: num_(v.Desconto), total: num_(v.Total), pagamentos: txt_(v.Pagamentos), status: txt_(v.Status),
      lucro: papel === 'dono' ? num_(v.Lucro) : null, itens: itens[v.ID] || [] };
  });
}

/** Cancela a venda: devolve as peças ao estoque e estorna o dinheiro. */
function apiCancelarVenda(pin, id, motivo) {
  var papel = auth_(pin, true);
  return comTrava_(function () {
    var v = rows_('Vendas').filter(function (x) { return String(x.ID) === String(id); })[0];
    if (!v) throw new Error('Venda não encontrada.');
    if (txt_(v.Status) === 'CANCELADA') throw new Error('Essa venda já foi cancelada.');
    var mapa = mapaProdutos_(), movs = [];
    rows_('VendaItens').forEach(function (i) {
      if (String(i.VendaID) !== String(id)) return;
      var p = mapa[String(i.SKU)];
      if (p) movs.push(movEstoque_(p, num_(i.Qtd), 'DEVOLUCAO', id, papel, txt_(motivo) || 'Venda cancelada'));
    });
    appendMany_('Movimentos', movs);
    var agora = new Date(), fin = [];
    rows_('Financeiro').forEach(function (f) {
      if (String(f.Referencia) === String(id) && f.Tipo === 'ENTRADA' && f.Categoria === 'Venda') {
        fin.push({ ID: contador_('PROX_FIN', 'F', 5), Data: agora, Tipo: 'SAIDA', Categoria: 'Estorno de venda',
          Descricao: 'Cancelamento ' + id, Valor: num_(f.Valor), Forma: f.Forma, Referencia: id, Usuario: papel });
      }
    });
    rows_('Crediario').forEach(function (c) {
      if (String(c.VendaID) !== String(id) || c.Status === 'CANCELADO') return;
      if (num_(c.Pago) > 0) {
        fin.push({ ID: contador_('PROX_FIN', 'F', 5), Data: agora, Tipo: 'SAIDA', Categoria: 'Estorno de venda',
          Descricao: 'Cancelamento ' + id + ' (crediário já pago)', Valor: num_(c.Pago), Forma: 'Dinheiro', Referencia: id, Usuario: papel });
      }
      set_('Crediario', c._row, 'Saldo', 0);
      set_('Crediario', c._row, 'Status', 'CANCELADO');
    });
    appendMany_('Financeiro', fin);
    set_('Vendas', v._row, 'Status', 'CANCELADA');
    set_('Vendas', v._row, 'Obs', (txt_(v.Obs) + ' Cancelada: ' + (txt_(motivo) || '-')).trim());
    return { ok: true };
  });
}

/* ======================= FINANCEIRO ======================= */

function apiLancar(pin, d) {
  var papel = auth_(pin, true);
  return comTrava_(function () {
    var valor = r2_(num_(d.valor));
    if (valor <= 0) throw new Error('Informe o valor.');
    var tipo = d.tipo === 'ENTRADA' ? 'ENTRADA' : 'SAIDA';
    var data = d.data ? new Date(d.data + 'T12:00:00') : new Date();
    append_('Financeiro', { ID: contador_('PROX_FIN', 'F', 5), Data: data, Tipo: tipo, Categoria: txt_(d.categoria) || 'Outros',
      Descricao: txt_(d.descricao), Valor: valor, Forma: txt_(d.forma) || 'Dinheiro', Referencia: '', Usuario: papel });
    return { ok: true };
  });
}

function apiExcluirLancamento(pin, id) {
  auth_(pin, true);
  return comTrava_(function () {
    var f = rows_('Financeiro').filter(function (x) { return String(x.ID) === String(id); })[0];
    if (!f) throw new Error('Lançamento não encontrado.');
    if (f.Categoria === 'Venda' || f.Categoria === 'Estorno de venda' || f.Categoria === 'Recebimento crediário')
      throw new Error('Esse lançamento é automático. Para desfazer, cancele a venda.');
    sh_('Financeiro').deleteRow(f._row);
    return { ok: true };
  });
}

function apiFinanceiro(pin, de, ate) {
  auth_(pin, true);
  return rows_('Financeiro').filter(function (f) { return noIntervalo_(f.Data, de, ate); }).reverse().map(function (f) {
    return { id: String(f.ID), data: ymd_(f.Data), tipo: f.Tipo, categoria: txt_(f.Categoria), descricao: txt_(f.Descricao),
      valor: num_(f.Valor), forma: txt_(f.Forma) };
  });
}

/** Resumo do período. O caixa só vê o dia de hoje e sem custos. */
function apiResumo(pin, de, ate) {
  var papel = auth_(pin);
  if (papel !== 'dono') { de = hoje_(); ate = hoje_(); }
  var r = { de: de, ate: ate, papel: papel, vendas: 0, pecas: 0, faturamento: 0, descontos: 0, custo: 0, lucroBruto: 0,
    porForma: {}, entradas: 0, saidas: 0, despesas: 0, comprasMercadoria: 0, despesasPorCategoria: {}, entradasPorForma: {},
    topProdutos: [], canceladas: 0 };
  rows_('Vendas').forEach(function (v) {
    if (!noIntervalo_(v.Data, de, ate)) return;
    if (txt_(v.Status) === 'CANCELADA') { r.canceladas++; return; }
    r.vendas++; r.pecas += num_(v.Qtd); r.faturamento += num_(v.Total); r.descontos += num_(v.Desconto);
    r.custo += num_(v.Custo); r.lucroBruto += num_(v.Lucro);
    txt_(v.Pagamentos).split(';').forEach(function (s) {
      var kv = s.split('='); if (kv.length < 2) return;
      var f = kv[0].trim(); r.porForma[f] = r2_((r.porForma[f] || 0) + num_(kv[1]));
    });
  });
  var canceladas = {};
  rows_('Vendas').forEach(function (v) { if (txt_(v.Status) === 'CANCELADA') canceladas[v.ID] = true; });
  var top = {};
  rows_('VendaItens').forEach(function (i) {
    if (!noIntervalo_(i.Data, de, ate) || canceladas[i.VendaID]) return;
    var k = txt_(i.Produto);
    top[k] = top[k] || { nome: k, qtd: 0, total: 0 };
    top[k].qtd += num_(i.Qtd); top[k].total += num_(i.Total);
  });
  r.topProdutos = Object.keys(top).map(function (k) { return top[k]; })
    .sort(function (a, b) { return b.qtd - a.qtd; }).slice(0, 10)
    .map(function (t) { t.total = r2_(t.total); return t; });

  if (papel === 'dono') {
    rows_('Financeiro').forEach(function (f) {
      if (!noIntervalo_(f.Data, de, ate)) return;
      var val = num_(f.Valor);
      if (f.Tipo === 'ENTRADA') {
        r.entradas += val;
        r.entradasPorForma[f.Forma] = r2_((r.entradasPorForma[f.Forma] || 0) + val);
      } else {
        r.saidas += val;
        if (f.Categoria === 'Compra de mercadoria') r.comprasMercadoria += val;
        else if (f.Categoria !== 'Estorno de venda') {
          r.despesas += val;
          r.despesasPorCategoria[f.Categoria] = r2_((r.despesasPorCategoria[f.Categoria] || 0) + val);
        }
      }
    });
    var est = { pecas: 0, custo: 0, venda: 0, baixo: 0, zerado: 0 };
    rows_('Produtos').forEach(function (p) {
      if (txt_(p.Ativo) === 'NAO') return;
      var q = num_(p.Estoque);
      if (q > 0) { est.pecas += q; est.custo += q * num_(p.Custo); est.venda += q * num_(p.Preco); }
      if (q <= 0) est.zerado++; else if (q <= num_(p.EstoqueMinimo)) est.baixo++;
    });
    est.custo = r2_(est.custo); est.venda = r2_(est.venda);
    r.estoque = est;
    var aReceber = 0, vencido = 0, h = hoje_();
    rows_('Crediario').forEach(function (c) {
      if (c.Status !== 'ABERTO') return;
      aReceber += num_(c.Saldo);
      if (ymd_(c.Vencimento) && ymd_(c.Vencimento) < h) vencido += num_(c.Saldo);
    });
    r.aReceber = r2_(aReceber); r.vencido = r2_(vencido);
  } else {
    r.custo = null; r.lucroBruto = null; r.topProdutos = r.topProdutos.map(function (t) { return { nome: t.nome, qtd: t.qtd }; });
  }
  ['faturamento', 'descontos', 'custo', 'lucroBruto', 'entradas', 'saidas', 'despesas', 'comprasMercadoria'].forEach(function (k) {
    if (r[k] !== null) r[k] = r2_(r[k]);
  });
  r.lucroLiquido = papel === 'dono' ? r2_(r.lucroBruto - r.despesas) : null;
  r.saldo = papel === 'dono' ? r2_(r.entradas - r.saidas) : null;
  return r;
}

/* ======================= CLIENTES / CREDIÁRIO ======================= */

function apiSalvarCliente(pin, c) {
  auth_(pin);
  return comTrava_(function () {
    var nome = txt_(c.nome);
    if (!nome) throw new Error('Informe o nome do cliente.');
    if (c.id) {
      var r = rows_('Clientes').filter(function (x) { return String(x.ID) === String(c.id); })[0];
      if (!r) throw new Error('Cliente não encontrado.');
      set_('Clientes', r._row, 'Nome', nome);
      set_('Clientes', r._row, 'Telefone', txt_(c.telefone));
      set_('Clientes', r._row, 'Obs', txt_(c.obs));
      return { id: String(c.id), nome: nome, telefone: txt_(c.telefone), obs: txt_(c.obs) };
    }
    var id = contador_('PROX_CLIENTE', 'CL', 4);
    append_('Clientes', { ID: id, Nome: nome, Telefone: txt_(c.telefone), Obs: txt_(c.obs), CriadoEm: new Date() });
    return { id: id, nome: nome, telefone: txt_(c.telefone), obs: txt_(c.obs) };
  });
}

function apiCrediario(pin) {
  auth_(pin);
  var h = hoje_(), por = {};
  rows_('Crediario').forEach(function (c) {
    if (c.Status === 'CANCELADO') return;
    var k = String(c.ClienteID);
    por[k] = por[k] || { clienteId: k, cliente: txt_(c.Cliente), saldo: 0, vencido: 0, contas: [] };
    var venc = ymd_(c.Vencimento);
    if (c.Status === 'ABERTO') {
      por[k].saldo += num_(c.Saldo);
      if (venc && venc < h) por[k].vencido += num_(c.Saldo);
    }
    por[k].contas.push({ id: String(c.ID), data: ymd_(c.Data), venda: String(c.VendaID), valor: num_(c.Valor), pago: num_(c.Pago),
      saldo: num_(c.Saldo), vencimento: venc, status: c.Status });
  });
  return Object.keys(por).map(function (k) {
    var x = por[k]; x.saldo = r2_(x.saldo); x.vencido = r2_(x.vencido); x.contas.reverse(); return x;
  }).sort(function (a, b) { return b.saldo - a.saldo; });
}

/** Recebe um valor do cliente e abate das contas mais antigas. */
function apiReceberCrediario(pin, clienteId, valor, forma) {
  var papel = auth_(pin);
  return comTrava_(function () {
    var resta = r2_(num_(valor));
    if (resta <= 0) throw new Error('Informe o valor recebido.');
    var abertas = rows_('Crediario').filter(function (c) { return String(c.ClienteID) === String(clienteId) && c.Status === 'ABERTO'; });
    var devido = r2_(abertas.reduce(function (a, c) { return a + num_(c.Saldo); }, 0));
    if (resta > devido + 0.001) throw new Error('Valor maior que o saldo devedor (R$ ' + devido.toFixed(2) + ').');
    var nome = abertas.length ? txt_(abertas[0].Cliente) : '';
    abertas.forEach(function (c) {
      if (resta <= 0) return;
      var paga = Math.min(resta, num_(c.Saldo));
      var novoSaldo = r2_(num_(c.Saldo) - paga);
      set_('Crediario', c._row, 'Pago', r2_(num_(c.Pago) + paga));
      set_('Crediario', c._row, 'Saldo', novoSaldo);
      if (novoSaldo <= 0) set_('Crediario', c._row, 'Status', 'PAGO');
      resta = r2_(resta - paga);
    });
    append_('Financeiro', { ID: contador_('PROX_FIN', 'F', 5), Data: new Date(), Tipo: 'ENTRADA', Categoria: 'Recebimento crediário',
      Descricao: 'Crediário - ' + nome, Valor: r2_(num_(valor)), Forma: txt_(forma) || 'Dinheiro', Referencia: String(clienteId), Usuario: papel });
    return { ok: true, saldoRestante: r2_(devido - num_(valor)) };
  });
}

/* ======================= CONFIGURAÇÕES ======================= */

function apiSalvarConfig(pin, c) {
  auth_(pin, true);
  return comTrava_(function () {
    var atual = cfg_();
    var pd = txt_(c.pinDono) || atual.PIN_DONO, pc = txt_(c.pinCaixa) || atual.PIN_CAIXA;
    if (pd.length < 4 || pc.length < 4) throw new Error('O PIN precisa ter pelo menos 4 números.');
    if (pd === pc) throw new Error('Os PINs do dono e do caixa precisam ser diferentes.');
    if (c.loja !== undefined) setCfg_('NOME_LOJA', txt_(c.loja));
    if (pd !== atual.PIN_DONO) setCfg_('PIN_DONO', pd);
    if (pc !== atual.PIN_CAIXA) setCfg_('PIN_CAIXA', pc);
    if (c.etqLargura) setCfg_('ETQ_LARGURA_MM', num_(c.etqLargura));
    if (c.etqAltura) setCfg_('ETQ_ALTURA_MM', num_(c.etqAltura));
    if (c.etqColunas) setCfg_('ETQ_COLUNAS', Math.max(1, Math.min(4, Math.round(num_(c.etqColunas)))));
    if (c.etqEspaco !== undefined) setCfg_('ETQ_ESPACO_MM', num_(c.etqEspaco));
    if (c.etqPreco !== undefined) setCfg_('ETQ_MOSTRAR_PRECO', c.etqPreco ? 'SIM' : 'NAO');
    return { ok: true };
  });
}

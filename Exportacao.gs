/***********************************************************************************************************************************
 *
 * NOME DO ARQUIVO: Exportacao.gs
 *
 ***********************************************************************************************************************************/

/**
 * @scriptName Lógica de Exportação de Arquivos (somente MANUAL)
 * @version 1.3 (pasta por ANO/MÊS-ptBR + check rápido de duplicidades + log detalhado por item + atualização do painel)
 * @description
 * - Lê "Cadastro Petiko" e coleta as linhas com a caixa F (Cadastro Manual) marcada.
 * - Valida se SKU/NCM/GTIN existem e não são "FALTANDO".
 * - Check rápido na própria aba "Cadastro Petiko" para:
 *      (1) SKU duplicado com Nome diferente  → BLOQUEAR
 *      (2) EAN/GTIN duplicado               → BLOQUEAR
 *      (3) Nome (Descrição) duplicado com SKU diferente → AVISAR (não bloqueia)
 * - Prepara palcos (Omie_Produtos / Omie_Produtos_2) e exporta por grupo (Petiko, Innova, Paws).
 * - Salva .xlsx em pastas específicas (CONFIG.ID_PASTA_EXPORTACAO) agora organizadas como:
 *      EMPRESA / ANO / "MM-NomeMes" / Arquivo.xlsx
 * - Registra no "LogDeArquivos" (A: Timestamp, B: Nome do Arquivo, C: URL, D: SKU, E: Nome, F: GTIN) **uma linha por item exportado**.
 *
 * Dependências:
 * - CONFIG (Config.gs)
 * - Utilitarios.gs: _encontrarUltimaLinhaNaColuna, _escreverStatus, _atualizarPainel
 */

// ======================================================================
// Entrada (manual): disparada pela janela Log.html
// ======================================================================

function executarExportacaoManual() {
  var linhasSelecionadas = [];
  var ss, abaInfo;

  try {
    _escreverStatus('[INICIANDO] Lendo dados da planilha...');
    ss = SpreadsheetApp.getActiveSpreadsheet();
    abaInfo = ss.getSheetByName(CONFIG.PLANILHA_DESTINO.INFORMACOES.NOME);
    if (!abaInfo) throw new Error("Aba 'Informações' não encontrada.");

    var idx = CONFIG.PLANILHA_DESTINO.INFORMACOES;
    var baseRow = idx.PRIMEIRA_LINHA_DADOS;
    var ultimaLinha = _encontrarUltimaLinhaNaColuna(abaInfo, 1);
    if (ultimaLinha < baseRow) throw new Error('Nenhum produto encontrado na planilha.');

    // Ler até a coluna F (checkbox manual) — A..F
    var largura = Math.max(6, idx.COL_CADASTRO_MANUAL_INDICE_0 + 1);
    var dados = abaInfo.getRange(baseRow, 1, (ultimaLinha - baseRow + 1), largura).getValues();

    // Selecionados (F = true)
    var selecionados = [];
    for (var i = 0; i < dados.length; i++) {
      if (dados[i][idx.COL_CADASTRO_MANUAL_INDICE_0] === true) {
        selecionados.push(dados[i]);
        linhasSelecionadas.push(baseRow + i); // guarda a linha absoluta correspondente
      }
    }
    if (selecionados.length === 0) {
      throw new Error("Nenhum produto foi selecionado na coluna 'Gerar Manualmente' (F).");
    }
    _escreverStatus('[OK] ' + selecionados.length + ' produto(s) selecionado(s). Validando...');

    // Validação essencial (A: SKU, C: NCM, D: GTIN)
    var validos = [];
    var linhasValidos = []; // paralelo a "validos"
    for (i = 0; i < selecionados.length; i++) {
      var r = selecionados[i];
      var sku  = r[idx.COL_SKU_INDICE_0];
      var ncm  = r[idx.COL_NCM_INDICE_0];
      var gtin = r[idx.COL_GTIN_INDICE_0];
      if (_campoOK(sku) && _campoOK(ncm) && _campoOK(gtin)) {
        validos.push(r);
        linhasValidos.push(linhasSelecionadas[i]);
      }
    }
    if (validos.length === 0) {
      throw new Error('Nenhum dos produtos selecionados passou na validação (SKU, NCM, GTIN).');
    }
    if (validos.length !== selecionados.length) {
      _escreverStatus('[AVISO] ' + (selecionados.length - validos.length) + ' produto(s) ignorado(s) por falta de dados.');
    }

    // === CHECK RÁPIDO DE DUPLICIDADES NA PRÓPRIA ABA "Cadastro Petiko" ===
    _escreverStatus('[EM ANDAMENTO] Checkando duplicidades (SKU/EAN/NOME) na aba "Cadastro Petiko"...');
    var resultadoCheck = _filtrarDuplicidadesRapido(validos, linhasValidos, dados, baseRow, idx);
    var validosFiltrados = resultadoCheck.validosFiltrados;
    var linhasValidosFiltradas = resultadoCheck.linhasValidosFiltradas;

    // resumo do check
    _escreverStatus('[OK] Checkup: ' + validosFiltrados.length + ' exportado(s), ' + resultadoCheck.barrados + ' barrado(s) por duplicidade.');

    if (validosFiltrados.length === 0) {
      throw new Error('Todos os itens selecionados foram barrados por duplicidade (SKU/EAN).');
    }

    // === Separar por grupo (com a lista já filtrada) ===
    _escreverStatus('[EM ANDAMENTO] Separando por grupo...');
    var industriaIdx = idx.COL_INDUSTRIA_INDICE_0;
    var grupoInnova = [], grupoPaws = [], grupoPetiko = [];
    for (i = 0; i < validosFiltrados.length; i++) {
      var p = validosFiltrados[i];
      if (p[industriaIdx] === 'Innova') grupoInnova.push(p);
      if (p[industriaIdx] === 'Paws')   grupoPaws.push(p);
      grupoPetiko.push(p);
    }

    if (grupoPetiko.length > 0) {
      _prepararAbaParaGrupo(CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS.NOME, grupoPetiko);
      _gerarArquivoParaGrupo('Petiko', 'MANUAL');
    }
    if (grupoInnova.length > 0) {
      _prepararAbaParaGrupo(CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS_2.NOME, grupoInnova);
      _gerarArquivoParaGrupo('Innova', 'MANUAL');
    }
    if (grupoPaws.length > 0) {
      _prepararAbaParaGrupo(CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS_2.NOME, grupoPaws);
      _gerarArquivoParaGrupo('Paws', 'MANUAL');
    }

    Utilities.sleep(500);

    // Atualiza o painel com sucesso (mensagem final) e mantém retorno para o Log.html
    var finalMsg = validosFiltrados.length + ' item(ns) exportado(s).';
    _escreverStatus('[OK] ' + finalMsg);
    _atualizarPainel('Sucesso ✅', finalMsg);

    return 'Processo concluído! ' + validosFiltrados.length + ' produto(s) foram processados. Os arquivos foram gerados e estão prontos para importação no Omie.';

  } catch (e) {
    _escreverStatus('[ERRO] ' + (e && e.message ? e.message : String(e)));
    _atualizarPainel('Erro ❌', (e && e.message) ? e.message : String(e));
    throw e;
  } finally {
    try {
      // Desmarca SEMPRE (sucesso ou erro) — e funciona com linhas não contíguas
      if (!abaInfo) {
        var ssSafe = SpreadsheetApp.getActiveSpreadsheet();
        abaInfo = ssSafe.getSheetByName(CONFIG.PLANILHA_DESTINO.INFORMACOES.NOME);
      }
      if (abaInfo && linhasSelecionadas.length) {
        _escreverStatus('[EM ANDAMENTO] Limpando caixas de seleção...');
        _limparCheckboxF(abaInfo, linhasSelecionadas);
      }
    } catch (e2) {
      // silencioso
    }
    Utilities.sleep(300);
    PropertiesService.getScriptProperties().deleteProperty('export_status');
  }
}



// ======================================================================
// Palco + Export
// ======================================================================

function _prepararAbaParaGrupo(nomeAba, dadosDoGrupo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) throw new Error('Aba de preparação "' + nomeAba + '" não foi encontrada.');

  var primeiraLinha = CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS.PRIMEIRA_LINHA_DADOS;
  var colInicialSKU = CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS.COL_SKU; // base 1 (C)

  // Limpa colunas C, D, E, G (palco padrão)
  var colunasParaLimpar = ['C', 'D', 'E', 'G'];
  var rangesParaLimpar = colunasParaLimpar.map(function(c){ return c + primeiraLinha + ':' + c; });
  aba.getRangeList(rangesParaLimpar).clearContent();

  // Monta [SKU, Nome, NCM, '', GTIN] começando na COLUNA C
  var idx = CONFIG.PLANILHA_DESTINO.INFORMACOES;
  var dadosParaCopiar = dadosDoGrupo.map(function(l) {
    var sku  = l[idx.COL_SKU_INDICE_0];
    var nome = l[idx.COL_NOME_INDICE_0];
    var ncm  = l[idx.COL_NCM_INDICE_0];
    var gtin = l[idx.COL_GTIN_INDICE_0];
    return [sku, nome, ncm, '', gtin];
  });

  if (dadosParaCopiar.length > 0) {
    aba.getRange(primeiraLinha, colInicialSKU, dadosParaCopiar.length, dadosParaCopiar[0].length).setValues(dadosParaCopiar);
    _escreverStatus('[OK] ' + dadosParaCopiar.length + ' produtos preparados na aba "' + nomeAba + '".');
  } else {
    _escreverStatus('[OK] Nenhum dado para preparar na aba "' + nomeAba + '".');
  }
}

function _gerarArquivoParaGrupo(nomeGrupo, tipoExecucao) {
  _escreverStatus('[EM ANDAMENTO] Gerando arquivo para o grupo ' + nomeGrupo + '...');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var idTemplate = CONFIG.PLANILHA_DESTINO.ID_TEMPLATE_EXPORTACAO;

  var idPastaDestino, nomeAbaOrigem;
  if (nomeGrupo.toUpperCase() === 'PETIKO') {
    idPastaDestino = CONFIG.ID_PASTA_EXPORTACAO.PETIKO;
    nomeAbaOrigem = CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS.NOME;
  } else if (nomeGrupo.toUpperCase() === 'INNOVA') {
    idPastaDestino = CONFIG.ID_PASTA_EXPORTACAO.INNOVA;
    nomeAbaOrigem = CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS_2.NOME;
  } else if (nomeGrupo.toUpperCase() === 'PAWS') {
    idPastaDestino = CONFIG.ID_PASTA_EXPORTACAO.PAWS;
    nomeAbaOrigem = CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS_2.NOME;
  } else {
    throw new Error('Grupo de exportação desconhecido: ' + nomeGrupo);
  }

  var abaOrigem = ss.getSheetByName(nomeAbaOrigem);
  if (!abaOrigem) throw new Error('Aba de origem ' + nomeAbaOrigem + ' não encontrada.');

  var primeiraLinha = CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS.PRIMEIRA_LINHA_DADOS;
  var colInicial = CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS.COL_SKU; // C (base 1)
  var ultimaLinha = _encontrarUltimaLinhaNaColuna(abaOrigem, colInicial);

  if (ultimaLinha < primeiraLinha) {
    _escreverStatus('[OK] Nenhum dado para exportar para o grupo ' + nomeGrupo + '. Pulando.');
    return;
  }

  var ultimaColuna = abaOrigem.getLastColumn();
  var numCols = (ultimaColuna - colInicial + 1);
  var rangeDados = abaOrigem.getRange(primeiraLinha, colInicial, (ultimaLinha - primeiraLinha + 1), numCols);
  var dadosParaExportar = rangeDados.getDisplayValues();

  _escreverStatus('[EM ANDAMENTO] Populando template...');
  var planilhaTemplate = SpreadsheetApp.openById(idTemplate);
  var abaDestinoProdutos = planilhaTemplate.getSheetByName(CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS.NOME);
  if (!abaDestinoProdutos) throw new Error('Aba de produtos não encontrada no template.');

  // Limpa e cola no template
  abaDestinoProdutos
    .getRange(primeiraLinha, colInicial, abaDestinoProdutos.getMaxRows() - primeiraLinha + 1, numCols)
    .clearContent();
  abaDestinoProdutos
    .getRange(primeiraLinha, colInicial, dadosParaExportar.length, dadosParaExportar[0].length)
    .setValues(dadosParaExportar);
  SpreadsheetApp.flush();

  // ===== NOVA ESTRUTURA DE PASTAS: EMPRESA / ANO / "MM-NomeMes" =====
  var tz = 'America/Sao_Paulo';
  var agora = new Date();

  var ano = Utilities.formatDate(agora, tz, 'yyyy'); // "2025"
  var mesNum = Utilities.formatDate(agora, tz, 'MM'); // "10"
  var mesNome = _nomeMesPtBR(parseInt(mesNum, 10));  // "Outubro"
  var nomeMesPasta = mesNum + '-' + mesNome;         // "10-Outubro"

  var pastaEmpresa = DriveApp.getFolderById(idPastaDestino);
  var pastaAno = _obterOuCriarSubpasta(pastaEmpresa, ano);
  var pastaMes = _obterOuCriarSubpasta(pastaAno, nomeMesPasta);

  // ===== Nome do arquivo (mantido)
  var dataFormatada = Utilities.formatDate(agora, tz, 'dd-MM-yyyy');
  var horaFormatada = Utilities.formatDate(agora, tz, 'HH-mm-ss');

  // Sequencial por grupo/dia dentro da pasta do mês
  var arquivosExistentes = pastaMes.getFiles();
  var contador = 0;
  var prefixo = (tipoExecucao === 'MANUAL') ? 'MANUAL_' : '';
  while (arquivosExistentes.hasNext()) {
    var arq = arquivosExistentes.next();
    if (arq.getName().startsWith(prefixo) && arq.getName().includes('_' + nomeGrupo + '_' + dataFormatada + '_')) {
      contador++;
    }
  }
  var sequencia = (contador + 1).toString().padStart(2, '0');
  var nomeArquivo = prefixo + sequencia + '_' + nomeGrupo + '_' + dataFormatada + '_' + horaFormatada + '.xlsx';

  _escreverStatus('[EM ANDAMENTO] Salvando arquivo ' + nomeArquivo + ' no Drive...');
  var url = 'https://docs.google.com/spreadsheets/d/' + idTemplate + '/export?format=xlsx';
  var token = ScriptApp.getOAuthToken();
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Falha ao exportar o template. Código: ' + resp.getResponseCode() + '. Mensagem: ' + resp.getContentText());
  }

  var arquivoSalvo = pastaMes.createFile(resp.getBlob()).setName(nomeArquivo);
  _escreverStatus('[OK] Arquivo para "' + nomeGrupo + '" salvo com sucesso.');

  // === Log detalhado: uma linha por item exportado (A:Timestamp, B:Arquivo, C:URL, D:SKU, E:Nome, F:GTIN) ===
  _registrarArquivoGeradoDetalhado(arquivoSalvo, dadosParaExportar);
}

// ======================================================================
// Helpers locais
// ======================================================================

function _campoOK(x) {
  var s = (x == null) ? '' : String(x).trim();
  return s !== '' && s.toUpperCase() !== 'FALTANDO';
}

function _limparCheckboxF(aba, linhas) {
  if (!linhas || !linhas.length) return;
  var colF = CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_CADASTRO_MANUAL_INDICE_0 + 1; // base 1
  var uniq = Array.from(new Set(linhas)).sort(function(a, b) { return a - b; });
  for (var i = 0; i < uniq.length; i++) {
    aba.getRange(uniq[i], colF, 1, 1).setValue(false);
  }
}

/**
 * Escreve no LogDeArquivos uma linha por item exportado.
 * Espera que o cabeçalho da aba seja:
 * A1: Timestamp | B1: Nome do Arquivo | C1: URL | D1: SKU | E1: Nome | F1: GTIN
 * @param {GoogleAppsScript.Drive.File} arquivo
 * @param {string[][]} dadosExportados - linhas do palco (a partir da coluna C): [SKU, Nome, NCM, '', GTIN, ...]
 */
function _registrarArquivoGeradoDetalhado(arquivo, dadosExportados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaLog = ss.getSheetByName('LogDeArquivos');
  if (!abaLog) {
    throw new Error("Aba 'LogDeArquivos' não encontrada. Não foi possível registrar o arquivo.");
  }
  var ts = new Date();
  var nomeArquivo = arquivo.getName();
  var url = arquivo.getUrl();

  var linhas = [];
  for (var i = 0; i < dadosExportados.length; i++) {
    var sku  = dadosExportados[i][0] || '';
    var nome = dadosExportados[i][1] || '';
    var gtin = dadosExportados[i][4] || '';
    linhas.push([ts, nomeArquivo, url, sku, nome, gtin]);
  }

  if (linhas.length === 0) return;

  // Modelo "append" no final
  abaLog.getRange(abaLog.getLastRow() + 1, 1, linhas.length, 6).setValues(linhas);
}

/** Cria (ou retorna) subpasta por nome dentro de uma pasta pai. */
function _obterOuCriarSubpasta(pastaPai, nome) {
  var it = pastaPai.getFoldersByName(nome);
  return it.hasNext() ? it.next() : pastaPai.createFolder(nome);
}

/** Nome do mês em pt-BR. @param {number} m 1..12 */
function _nomeMesPtBR(m) {
  var meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  return meses[m-1] || String(m);
}

/**
 * Check rápido de duplicidades (somente dentro da aba "Cadastro Petiko").
 * BLOQUEIA: (1) SKU duplicado com Nome diferente; (2) GTIN/EAN duplicado
 * AVISA:    (3) Nome igual com SKU diferente
 */
function _filtrarDuplicidadesRapido(validos, linhasValidos, dados, baseRow, idx) {
  var colSKU = idx.COL_SKU_INDICE_0;   // A -> 0
  var colNome = idx.COL_NOME_INDICE_0; // B -> 1
  var colGTIN = idx.COL_GTIN_INDICE_0; // D -> 3

  var mapSku = new Map();
  var mapGtin = new Map();
  var mapNome = new Map();

  for (var i = 0; i < dados.length; i++) {
    var r = dados[i];
    var abs = baseRow + i;
    var sku = r[colSKU];
    var nome = r[colNome];
    var gtin = r[colGTIN];

    if (sku !== '' && sku != null) {
      if (!mapSku.has(sku)) mapSku.set(sku, []);
      mapSku.get(sku).push(abs);
    }
    if (gtin !== '' && gtin != null) {
      if (!mapGtin.has(gtin)) mapGtin.set(gtin, []);
      mapGtin.get(gtin).push(abs);
    }
    if (nome !== '' && nome != null) {
      if (!mapNome.has(nome)) mapNome.set(nome, []);
      mapNome.get(nome).push(abs);
    }
  }

  var barradosSet = new Set();
  var barradosCount = 0;

  for (var j = 0; j < validos.length; j++) {
    var row = validos[j];
    var absLine = linhasValidos[j];
    var skuC = row[colSKU] || '';
    var nomeC = row[colNome] || '';
    var gtinC = row[colGTIN] || '';

    var bloqueado = false;

    var skuHits = (mapSku.get(skuC) || []).filter(function(n){ return n !== absLine; });
    if (skuHits.length > 0) {
      var conflitos = [];
      for (var k = 0; k < skuHits.length; k++) {
        var idxRow = skuHits[k] - baseRow;
        var nomeOutro = (dados[idxRow] && dados[idxRow][colNome]) || '';
        if (nomeOutro !== nomeC) conflitos.push(skuHits[k]);
      }
      if (conflitos.length > 0) {
        _escreverStatus('[FALHA] "' + nomeC + '" (SKU ' + skuC + ') barrado: SKU duplicado com nome diferente (linha(s) ' + conflitos.join(', ') + ').');
        bloqueado = true;
      }
    }

    var gtinHits = (mapGtin.get(gtinC) || []).filter(function(n){ return n !== absLine; });
    if (gtinHits.length > 0) {
      _escreverStatus('[FALHA] "' + nomeC + '" (SKU ' + skuC + ') barrado: EAN ' + gtinC + ' duplicado (linha(s) ' + gtinHits.join(', ') + ').');
      bloqueado = true;
    }

    var nomeHits = (mapNome.get(nomeC) || []).filter(function(n){ return n !== absLine; });
    if (nomeHits.length > 0) {
      var difSkuRows = [];
      for (var t = 0; t < nomeHits.length; t++) {
        var idxRow2 = nomeHits[t] - baseRow;
        var skuOutro = (dados[idxRow2] && dados[idxRow2][colSKU]) || '';
        if (skuOutro !== skuC) difSkuRows.push(nomeHits[t]);
      }
      if (difSkuRows.length > 0) {
        var plural = (difSkuRows.length > 1) ? 'linhas ' : 'linha ';
        _escreverStatus('[AVISO] "' + nomeC + '" (SKU ' + skuC + '): Descrição igual a outra com SKU diferente (' + plural + difSkuRows.join(', ') + ').');
      }
    }

    if (bloqueado) {
      if (!barradosSet.has(j)) {
        barradosSet.add(j);
        barradosCount++;
      }
    }
  }

  var validosFiltrados = [];
  var linhasValidosFiltradas = [];
  for (var m = 0; m < validos.length; m++) {
    if (!barradosSet.has(m)) {
      validosFiltrados.push(validos[m]);
      linhasValidosFiltradas.push(linhasValidos[m]);
    }
  }

  return {
    validosFiltrados: validosFiltrados,
    linhasValidosFiltradas: linhasValidosFiltradas,
    barrados: barradosCount
  };
}

/***********************************************************************************************************************************
 *
 * SUGESTÃO DE NOME PARA O ARQUIVO: InterfaceBackend.gs
 *
 ***********************************************************************************************************************************/
/**
 * @version 2.4 — Gerenciador “bottom-first” + dedupe por arquivo + downloadUrl
 * - UI mínima: Painel, Log e Gerenciador
 * - Adiciona obterUltimosArquivosParaPainel(qtd)
 */





function onOpen() {
  if (typeof iniciarInterfacePrincipal === 'function') {
    iniciarInterfacePrincipal();
  }
  if (typeof createMetabaseMenu === 'function') {
    createMetabaseMenu();
  }
}






function iniciarInterfacePrincipal() {
  SpreadsheetApp.getUi()
    .createMenu('▶️ Painel de Controle')
    .addItem('Abrir Painel', 'exibirPainelDeControle')
    .addSeparator()
    .addItem('🔄 Sincronizar Manualmente', 'importarDadosEConsultarAbas')
    .addItem('🗂️ Gerenciar Arquivos', 'abrirGerenciadorDeArquivos')
    .addToUi();
  exibirPainelDeControle();
}

function exibirPainelDeControle() {
  const html = HtmlService.createTemplateFromFile('PainelDeControle.html')
    .evaluate()
    .setTitle('Painel de Controle OMIE');
  SpreadsheetApp.getUi().showSidebar(html);
}

function incluir(nomeArquivo) {
  return HtmlService.createHtmlOutputFromFile(nomeArquivo).getContent();
}

function iniciarGeracaoManual() {
  const html = HtmlService.createTemplateFromFile('Log.html')
    .evaluate()
    .setWidth(700)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Log da Exportação Manual');
}

function abrirGerenciadorDeArquivos() {
  const html = HtmlService.createHtmlOutputFromFile('GerenciadorDeArquivos.html')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gerenciador de Arquivos Exportados');
}

// -------- Painel: dados simples (sem escrever em células) ----------
function obterDadosDoPainel() {
  try {
    const props = PropertiesService.getScriptProperties();
    const status = props.getProperty('panel_status') || '--';
    const resultado = props.getProperty('panel_result') || '--';
    const timestamp = props.getProperty('panel_ts') || '--';
    return { status, resultado, timestamp };
  } catch (e) {
    return { error: e.message };
  }
}

// -------- Log: polling ----------
function obterStatusDaExportacao() {
  try {
    return PropertiesService.getScriptProperties().getProperty('export_status');
  } catch (e) {
    return '[ERRO] ' + (e && e.message ? e.message : String(e));
  }
}

// -------- Gerenciador (lê do final: últimos N únicos, dedupe por arquivo) ----------
function obterDadosGerenciador() {
  var result = {
    linksPastas: { petiko: '', innova: '', paws: '' },
    arquivosRecentes: []
  };

  // 1) Links das pastas
  try {
    var idsPastas = CONFIG.ID_PASTA_EXPORTACAO;
    result.linksPastas = {
      petiko: DriveApp.getFolderById(idsPastas.PETIKO).getUrl(),
      innova: DriveApp.getFolderById(idsPastas.INNOVA).getUrl(),
      paws:   DriveApp.getFolderById(idsPastas.PAWS).getUrl()
    };
  } catch (e) {
    result.error = 'Falha ao obter links das pastas: ' + e.message;
    return result;
  }

  // 2) Últimos arquivos no FINAL da aba LogDeArquivos (modelo append)
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('LogDeArquivos');
    if (!sh) return result;

    var last = sh.getLastRow();
    if (last <= 1) return result; // só cabeçalho

    var COUNT_UNIQUE = 10;                 // queremos os 10 ARQUIVOS únicos mais recentes
    var SCAN_RANGE   = Math.min(400, last - 1); // varre no máx. 400 linhas do fim
    var numCols      = Math.min(6, sh.getLastColumn());

    var start = Math.max(2, last - SCAN_RANGE + 1);
    var rows  = last - start + 1;

    var values   = sh.getRange(start, 1, rows, numCols).getValues();           // A..F (bruto)
    var displays = sh.getRange(start, 1, rows, numCols).getDisplayValues();    // A..F (display)
    var rich     = sh.getRange(start, 3, rows, 1).getRichTextValues();         // C (rich link)
    var formulas = sh.getRange(start, 3, rows, 1).getFormulas();               // C (fórmulas)

    var fromHyperlinkFormula = function (raw) {
      var s = String(raw || '');
      var m = s.match(/HYPERLINK\(\s*"([^"]+)"\s*[,;]/i);
      return m ? m[1] : null;
    };
    var firstHttp = function (s) {
      s = String(s || '').trim();
      return s.startsWith('http') ? s : '';
    };
    var extractId = function (url) {
      var m = String(url || '').match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (m) return m[1];
      m = String(url || '').match(/[?&]id=([a-zA-Z0-9_-]+)/);
      return m ? m[1] : null;
    };
    var toDl = function (url) {
      var id = extractId(url);
      return id ? ('https://drive.google.com/uc?export=download&id=' + id) : url;
    };

    var coletados = [];
    var vistos = new Set(); // chave: fileId (ou, se não tiver, a própria URL)

    // Itera de baixo pra cima: mais recentes primeiro
    for (var i = rows - 1; i >= 0 && coletados.length < COUNT_UNIQUE; i--) {
      var rowDisplays = displays[i];
      var rowValues   = values[i];
      var rt = rich[i] && rich[i][0];
      var fm = formulas[i] && formulas[i][0];

      var ts   = rowDisplays[0]; // A
      var nome = rowDisplays[1]; // B

      var url = (rt && rt.getLinkUrl())
             || fromHyperlinkFormula(fm)
             || firstHttp(rowDisplays[2])
             || firstHttp(rowValues[2]);

      if (!url) continue;

      var id = extractId(url);
      var key = id ? id : url; // preferimos dedupe por fileId
      if (vistos.has(key)) continue;
      vistos.add(key);

      coletados.push({
        timestamp: ts,
        nome: nome,
        url: url,
        downloadUrl: toDl(url)
      });
    }

    result.arquivosRecentes = coletados;
    return result;
  } catch (e2) {
    result.error = 'Falha ao obter arquivos recentes: ' + e2.message;
    return result;
  }
}

/**
 * NOVO: Retorna os últimos N ARQUIVOS únicos para o Painel (nome curto + download).
 * Estrutura: { arquivos: [{nomeCurto, url, downloadUrl, timestamp}] }
 */
function obterUltimosArquivosParaPainel(qtd) {
  var res = { arquivos: [] };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('LogDeArquivos');
    if (!sh) return res;

    var last = sh.getLastRow();
    if (last <= 1) return res;

    var COUNT_UNIQUE = Math.max(1, Math.min(Number(qtd) || 5, 10));
    var SCAN_RANGE   = Math.min(400, last - 1);
    var numCols      = Math.min(6, sh.getLastColumn());

    var start = Math.max(2, last - SCAN_RANGE + 1);
    var rows  = last - start + 1;

    var values   = sh.getRange(start, 1, rows, numCols).getValues();
    var displays = sh.getRange(start, 1, rows, numCols).getDisplayValues();
    var rich     = sh.getRange(start, 3, rows, 1).getRichTextValues();
    var formulas = sh.getRange(start, 3, rows, 1).getFormulas();

    var fromHyperlinkFormula = function (raw) {
      var s = String(raw || '');
      var m = s.match(/HYPERLINK\(\s*"([^"]+)"\s*[,;]/i);
      return m ? m[1] : null;
    };
    var firstHttp = function (s) {
      s = String(s || '').trim();
      return s.startsWith('http') ? s : '';
    };
    var extractId = function (url) {
      var m = String(url || '').match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (m) return m[1];
      m = String(url || '').match(/[?&]id=([a-zA-Z0-9_-]+)/);
      return m ? m[1] : null;
    };
    var toDl = function (url) {
      var id = extractId(url);
      return id ? ('https://drive.google.com/uc?export=download&id=' + id) : url;
    };
    var nomeCurto = function (nomeCheio) {
      var base = String(nomeCheio || '').replace(/\.[^.]+$/, '');
      var m = base.match(/^([^_]+)_([^_]+)_([^_]+)/); // MANUAL_02_Paws_...
      if (m) return m[1] + '_' + m[2] + '_' + m[3];
      // fallback: corta até ~24 chars
      return base.length > 24 ? (base.slice(0, 24) + '…') : base;
    };

    var itens = [];
    var vistos = new Set();

    for (var i = rows - 1; i >= 0 && itens.length < COUNT_UNIQUE; i--) {
      var rowDisplays = displays[i];
      var rowValues   = values[i];
      var rt = rich[i] && rich[i][0];
      var fm = formulas[i] && formulas[i][0];

      var ts   = rowDisplays[0];  // A
      var nome = rowDisplays[1];  // B

      var url = (rt && rt.getLinkUrl())
             || fromHyperlinkFormula(fm)
             || firstHttp(rowDisplays[2])
             || firstHttp(rowValues[2]);
      if (!url) continue;

      var id = extractId(url);
      var key = id ? id : url;
      if (vistos.has(key)) continue;
      vistos.add(key);

      itens.push({
        nomeCurto: nomeCurto(nome),
        url: url,
        downloadUrl: toDl(url),
        timestamp: ts
      });
    }

    res.arquivos = itens;
    return res;
  } catch (e) {
    res.error = e && e.message ? e.message : String(e);
    return res;
  }
}

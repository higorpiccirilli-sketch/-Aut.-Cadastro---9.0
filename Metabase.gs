/**
 * Metabase Updater — Versão enxuta (somente 1328 e 1317)
 * Menu único: "Atualizar relatórios Metabase"
 * Abas: "Rel 1328" (card 1328) e "Qtd Box" (card 1317)
 * - Lê credenciais de Script Properties (MB_URL, MB_USER, MB_PASS, ALERT_EMAIL)
 * - Se faltar algo, usa DEFAULT_SECRETS abaixo como fallback (para não travar)
 * - Cria a aba antes da consulta e registra erro em A1 se falhar
 * - Cache de sessão (30 min) + retry em 401
 */

// =================== RELATÓRIOS ===================
const RELATORIOS = [
  { cardId: 1328, sheetName: 'Dados Box' },
  { cardId: 1317, sheetName: 'Qtd Box' }
];

// =================== FALLBACK DE SEGREDOS (apenas p/ evitar travar) ===================
const DEFAULT_SECRETS = {
  MB_URL: 'https://bi.petiko.com.br',
  MB_USER: '',
  MB_PASS: '',
  ALERT_EMAIL: ''
};

// =================== SEGREDOS (Script Properties) ===================
function getCfg_() {
  const p = PropertiesService.getScriptProperties();
  // lê propriedades (se existirem)
  let url  = (p.getProperty('MB_URL')  || '').trim();
  let user = (p.getProperty('MB_USER') || '').trim();
  let pass = (p.getProperty('MB_PASS') || '').trim();
  let alertEmail = (p.getProperty('ALERT_EMAIL') || '').trim();

  // se faltar algo, usa fallback
  const faltas = [];
  if (!url)  { url  = DEFAULT_SECRETS.MB_URL;  faltas.push('MB_URL'); }
  if (!user) { user = DEFAULT_SECRETS.MB_USER; faltas.push('MB_USER'); }
  if (!pass) { pass = DEFAULT_SECRETS.MB_PASS; faltas.push('MB_PASS'); }
  if (!alertEmail) { alertEmail = DEFAULT_SECRETS.ALERT_EMAIL; }

  if (faltas.length) {
    Logger.log('Aviso: usando DEFAULT_SECRETS para: ' + faltas.join(', '));
    try {
      SpreadsheetApp.getActive().toast('Usando credenciais padrão (DEFAULT_SECRETS). Recomendo salvar em Script Properties.');
    } catch (_) {}
  }

  return {
    baseUrl: url.replace(/\/+$/, ''),
    user: user,
    pass: pass,
    alertEmail: alertEmail
  };
}

/**
 * OPCIONAL: Rode UMA VEZ para gravar os segredos nas Script Properties.
 * (Você pode editar os valores aqui antes de rodar.)
 */
function configurarSegredos() {
  PropertiesService.getScriptProperties().setProperties({
    MB_URL: DEFAULT_SECRETS.MB_URL,
    MB_USER: DEFAULT_SECRETS.MB_USER,
    MB_PASS: DEFAULT_SECRETS.MB_PASS,
    ALERT_EMAIL: DEFAULT_SECRETS.ALERT_EMAIL
  }, true);
  SpreadsheetApp.getActive().toast('Segredos salvos nas Script Properties.');
}

// =================== MENU ===================
function createMetabaseMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Metabase')
    .addItem('Atualizar relatórios Metabase', 'atualizarRelatorios')   // sua execução direta (sem log)
    .addItem('Atualizar (com Log)', 'abrirLogMetabase')                // NOVO: abre o modal com polling
    .addToUi();
}


// =================== EXECUÇÃO ===================
function atualizarRelatorios() {
  const ss = SpreadsheetApp.getActive();

  RELATORIOS.forEach(r => {
    let sh = ss.getSheetByName(r.sheetName);
    if (!sh) sh = ss.insertSheet(r.sheetName); // garante a aba antes de consultar

    try {
      buscarDadosDoMetabase(r.cardId, r.sheetName, null);
      Logger.log(`OK: ${r.sheetName} ← card ${r.cardId}`);
    } catch (e) {
      Logger.log(`Erro em ${r.sheetName} ← card ${r.cardId}: ${e}`);
      sh.clear();
      sh.getRange('A1').setValue(
        `Erro ao atualizar "${r.sheetName}" (card ${r.cardId}).\n` +
        (e && e.message ? e.message : String(e))
      ).setFontWeight('bold');
      sendAlert_(r.sheetName, r.cardId, e); // opcional
    }
  });

  SpreadsheetApp.getActive().toast('Atualização finalizada.');
}

// =================== METABASE: SESSÃO + CONSULTA ===================
function getMbToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('mb_token');
  if (cached) return cached;

  const cfg = getCfg_();
  const res = UrlFetchApp.fetch(cfg.baseUrl + '/api/session', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ username: cfg.user, password: cfg.pass }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() >= 300) {
    throw new Error('Falha no login do Metabase: ' + res.getContentText());
  }

  const token = JSON.parse(res.getContentText()).id;
  cache.put('mb_token', token, 1800); // 30 min
  return token;
}

function fetchCardJson_(cardId, params) {
  const cfg = getCfg_();
  const doQuery = (token) => UrlFetchApp.fetch(`${cfg.baseUrl}/api/card/${cardId}/query/json`, {
    method: 'post',
    headers: { 'X-Metabase-Session': token },
    contentType: 'application/json',
    payload: params ? JSON.stringify({ parameters: params }) : JSON.stringify({}),
    muteHttpExceptions: true
  });

  let token = getMbToken_();
  let res = doQuery(token);
  if (res.getResponseCode() === 401) {
    CacheService.getScriptCache().remove('mb_token'); // expirada → refaz login
    token = getMbToken_();
    res = doQuery(token);
  }
  if (res.getResponseCode() >= 300) {
    throw new Error(`Erro ao consultar card ${cardId}: ${res.getContentText()}`);
  }
  return JSON.parse(res.getContentText());
}

/**
 * Busca dados do card e escreve na aba.
 * params (opcional): objeto com parâmetros do Metabase (se o card usa parâmetros).
 */
function buscarDadosDoMetabase(cardId, nomeDaAba, params) {
  const dados = fetchCardJson_(cardId, params);
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(nomeDaAba);
  if (!sh) sh = ss.insertSheet(nomeDaAba);
  sh.clear();

  if (!dados || !dados.length) {
    sh.getRange('A1').setValue(`Nenhum dado retornado (Card ${cardId}).`);
    return;
  }

  const headers = Object.keys(dados[0]);
  const out = [headers];
  dados.forEach(row => out.push(headers.map(h => row[h])));

  sh.getRange(1, 1, out.length, headers.length).setValues(out);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sh.getRange('A1').setNote(`Atualizado em ${new Date().toLocaleString()}`);
  sh.autoResizeColumns(1, headers.length);
  SpreadsheetApp.flush();
}

// =================== UTILITÁRIOS ===================
function sendAlert_(nomeAba, cardId, errorObj) {
  const cfg = getCfg_();
  if (!cfg.alertEmail) return;
  const assunto = `[ALERTA] Falha ao atualizar "${nomeAba}" (card ${cardId})`;
  const corpo = `Planilha: ${SpreadsheetApp.getActive().getName()}
Aba: ${nomeAba}
CardId: ${cardId}

Erro: ${errorObj && errorObj.message ? errorObj.message : String(errorObj)}

Verifique os logs (Executions) no Apps Script.`;
  MailApp.sendEmail(cfg.alertEmail, assunto, corpo);
}

function resetCacheSessao() {
  CacheService.getScriptCache().remove('mb_token');
  SpreadsheetApp.getActive().toast('Cache de sessão do Metabase limpo.');
}

// Diagnóstico rápido (opcional)
function diagnosticoMb() {
  const ui = SpreadsheetApp.getUi();
  try {
    const cfg = getCfg_();
    const res = UrlFetchApp.fetch(cfg.baseUrl + '/api/session', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ username: cfg.user, password: cfg.pass }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code >= 300) {
      ui.alert('Diagnóstico',
        'Falha no login do Metabase (' + code + '):\n' + res.getContentText(),
        ui.ButtonSet.OK);
    } else {
      ui.alert('Diagnóstico', 'Login OK. Pode rodar "Atualizar relatórios Metabase".', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Diagnóstico', 'Erro: ' + (e.message || String(e)), ui.ButtonSet.OK);
  }
}

/** ===== LOG METABASE: helpers ===== */
function _mbWriteStatus(msg) {
  console.log(msg);
  PropertiesService.getScriptProperties().setProperty('mb_status', String(msg || ''));
}
function obterStatusMetabase() {
  return PropertiesService.getScriptProperties().getProperty('mb_status');
}

/** Abre o modal de log e dispara a atualização com log (não mexe no painel) */
function abrirLogMetabase() {
  const html = HtmlService.createHtmlOutputFromFile('MetabaseLog.html')
    .setWidth(700)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Log de Atualização (Metabase)');
}

/**
 * Wrapper que executa a atualização com mensagens de status.
 * (Mantém sua lógica em atualizarRelatorios(); aqui só adicionamos o log.)
 */
function atualizarRelatoriosComLog() {
  try {
    _mbWriteStatus('[INICIANDO] Preparando atualização de relatórios do Metabase...');
    const ss = SpreadsheetApp.getActive();

    // Passo 1: autenticar sessão
    _mbWriteStatus('[EM ANDAMENTO] Autenticando no Metabase...');
    const token = getMbToken_();
    _mbWriteStatus('[OK] Sessão autenticada.');

    // Passo 2: processar cada relatório
    for (var i = 0; i < RELATORIOS.length; i++) {
      const r = RELATORIOS[i];
      _mbWriteStatus(`[EM ANDAMENTO] Atualizando "${r.sheetName}" (card ${r.cardId})...`);
      try {
        // Garante a aba e atualiza
        let sh = ss.getSheetByName(r.sheetName);
        if (!sh) sh = ss.insertSheet(r.sheetName);
        buscarDadosDoMetabase(r.cardId, r.sheetName, null);
        _mbWriteStatus(`[OK] ${r.sheetName} atualizado.`);
      } catch (eItem) {
        _mbWriteStatus(`[ERRO] Falha em "${r.sheetName}" (card ${r.cardId}): ${eItem && eItem.message ? eItem.message : String(eItem)}`);
        // mantém o comportamento do seu atualizarRelatorios: limpa a aba e escreve A1
        let sh = ss.getSheetByName(r.sheetName);
        if (!sh) sh = ss.insertSheet(r.sheetName);
        sh.clear();
        sh.getRange('A1').setValue(
          `Erro ao atualizar "${r.sheetName}" (card ${r.cardId}).\n` +
          (eItem && eItem.message ? eItem.message : String(eItem))
        ).setFontWeight('bold');
        // alerta opcional
        try { sendAlert_(r.sheetName, r.cardId, eItem); } catch (_){}
      }
    }

    // Passo 3: final
    const msgFinal = 'Atualização concluída! Os relatórios foram gerados e estão prontos para uso.';
    _mbWriteStatus('[OK] ' + msgFinal);
    return msgFinal;
  } catch (e) {
    _mbWriteStatus('[ERRO FATAL] ' + (e && e.message ? e.message : String(e)));
    throw e;
  } finally {
    Utilities.sleep(300);
    // limpa a propriedade para encerrar o polling no HTML (mesma lógica do nosso Log)
    PropertiesService.getScriptProperties().deleteProperty('mb_status');
  }
}

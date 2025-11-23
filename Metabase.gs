/**
 * Metabase Updater — Versão Segura
 * Menu único: "Atualizar relatórios Metabase"
 * Abas: "Rel 1328" (card 1328) e "Qtd Box" (card 1317)
 * 
 * SEGURANÇA: As credenciais devem estar salvas nas Script Properties.
 * Não escreva senhas diretamente neste código.
 */

// =================== RELATÓRIOS ===================
const RELATORIOS = [
  { cardId: 1328, sheetName: 'Dados Box' },
  { cardId: 1317, sheetName: 'Qtd Box' }
];

// =================== CONFIGURAÇÃO SEGURA ===================
function getCfg_() {
  const p = PropertiesService.getScriptProperties();
  
  // Lê do cofre
  const url  = (p.getProperty('MB_URL')  || '').trim();
  const user = (p.getProperty('MB_USER') || '').trim();
  const pass = (p.getProperty('MB_PASS') || '').trim();
  const alertEmail = (p.getProperty('ALERT_EMAIL') || '').trim();

  // Validação de segurança
  const faltas = [];
  if (!url)  faltas.push('MB_URL');
  if (!user) faltas.push('MB_USER');
  if (!pass) faltas.push('MB_PASS');

  if (faltas.length > 0) {
    const msg = 'ERRO DE SEGURANÇA: As seguintes credenciais não foram encontradas nas Propriedades do Script: ' + faltas.join(', ');
    throw new Error(msg);
  }

  return {
    baseUrl: url.replace(/\/+$/, ''),
    user: user,
    pass: pass,
    alertEmail: alertEmail
  };
}

// =================== MENU ===================
function createMetabaseMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Metabase')
    .addItem('Atualizar relatórios Metabase', 'atualizarRelatorios')
    .addItem('Atualizar (com Log)', 'abrirLogMetabase')
    .addToUi();
}

// =================== EXECUÇÃO ===================
function atualizarRelatorios() {
  const ss = SpreadsheetApp.getActive();
  RELATORIOS.forEach(r => {
    let sh = ss.getSheetByName(r.sheetName);
    if (!sh) sh = ss.insertSheet(r.sheetName);

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
      sendAlert_(r.sheetName, r.cardId, e);
    }
  });
  SpreadsheetApp.getActive().toast('Atualização finalizada.');
}

// =================== METABASE: SESSÃO + CONSULTA ===================
function getMbToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('mb_token');
  if (cached) return cached;

  const cfg = getCfg_(); // Agora puxa do cofre seguro
  const res = UrlFetchApp.fetch(cfg.baseUrl + '/api/session', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ username: cfg.user, password: cfg.pass }),
    muteHttpExceptions: true
  });
  
  if (res.getResponseCode() >= 300) {
    throw new Error('Falha no login do Metabase (Verifique Script Properties): ' + res.getContentText());
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
    CacheService.getScriptCache().remove('mb_token');
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
  try {
    const cfg = getCfg_();
    if (!cfg.alertEmail) return;
    const assunto = `[ALERTA] Falha ao atualizar "${nomeAba}" (card ${cardId})`;
    const corpo = `Planilha: ${SpreadsheetApp.getActive().getName()}
Aba: ${nomeAba}
CardId: ${cardId}

Erro: ${errorObj && errorObj.message ? errorObj.message : String(errorObj)}`;
    MailApp.sendEmail(cfg.alertEmail, assunto, corpo);
  } catch(e) {
    console.error("Erro ao enviar alerta de email: " + e.message);
  }
}

function resetCacheSessao() {
  CacheService.getScriptCache().remove('mb_token');
  SpreadsheetApp.getActive().toast('Cache de sessão do Metabase limpo.');
}

// Diagnóstico rápido
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
      ui.alert('Diagnóstico', 'Falha no login (' + code + '):\n' + res.getContentText(), ui.ButtonSet.OK);
    } else {
      ui.alert('Diagnóstico', 'Login OK. Credenciais seguras funcionando.', ui.ButtonSet.OK);
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

/** Abre o modal de log e dispara a atualização com log */
function abrirLogMetabase() {
  const html = HtmlService.createHtmlOutputFromFile('MetabaseLog.html')
    .setWidth(700)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Log de Atualização (Metabase)');
}

function atualizarRelatoriosComLog() {
  try {
    _mbWriteStatus('[INICIANDO] Preparando atualização de relatórios do Metabase...');
    const ss = SpreadsheetApp.getActive();

    _mbWriteStatus('[EM ANDAMENTO] Autenticando no Metabase...');
    // Apenas chama getMbToken_, que já usa getCfg_ segura
    getMbToken_(); 
    _mbWriteStatus('[OK] Sessão autenticada.');

    for (var i = 0; i < RELATORIOS.length; i++) {
      const r = RELATORIOS[i];
      _mbWriteStatus(`[EM ANDAMENTO] Atualizando "${r.sheetName}" (card ${r.cardId})...`);
      try {
        let sh = ss.getSheetByName(r.sheetName);
        if (!sh) sh = ss.insertSheet(r.sheetName);
        buscarDadosDoMetabase(r.cardId, r.sheetName, null);
        _mbWriteStatus(`[OK] ${r.sheetName} atualizado.`);
      } catch (eItem) {
        _mbWriteStatus(`[ERRO] Falha em "${r.sheetName}": ${eItem.message}`);
        let sh = ss.getSheetByName(r.sheetName);
        if (!sh) sh = ss.insertSheet(r.sheetName);
        sh.clear();
        sh.getRange('A1').setValue(`Erro: ${eItem.message}`).setFontWeight('bold');
        sendAlert_(r.sheetName, r.cardId, eItem);
      }
    }

    const msgFinal = 'Atualização concluída! Relatórios seguros e gerados.';
    _mbWriteStatus('[OK] ' + msgFinal);
    return msgFinal;
  } catch (e) {
    _mbWriteStatus('[ERRO FATAL] ' + e.message);
    throw e;
  } finally {
    Utilities.sleep(300);
    PropertiesService.getScriptProperties().deleteProperty('mb_status');
  }
}

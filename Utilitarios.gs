/***********************************************************************************************************************************
 *
 * NOME DO ARQUIVO: Utilitarios.gs
 *
 ***********************************************************************************************************************************/

/**
 * @scriptName Funções Utilitárias Globais
 * @version 1.1 — logger append no final
 */

function _encontrarUltimaLinhaNaColuna(aba, indiceColuna) {
  const todos = aba.getRange(1, indiceColuna, aba.getMaxRows()).getValues();
  for (let i = todos.length - 1; i >= 0; i--) {
    if (todos[i][0] !== "") return i + 1;
  }
  return 1;
}

function _escreverStatus(mensagem) {
  console.log(mensagem);
  PropertiesService.getScriptProperties().setProperty('export_status', mensagem);
}

/**
 * Logger simples (A:ts, B:nome, C:url) — agora APPEND no final e mantém 200 dados.
 */
function _registrarArquivoGerado(arquivo) {
  try {
    const abaLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LogDeArquivos');
    if (!abaLog) {
      console.error("Aba 'LogDeArquivos' não encontrada. Não foi possível registrar o arquivo.");
      return;
    }
    const timestamp = new Date();
    const nome = arquivo.getName();
    const url = arquivo.getUrl();

    // APPEND no final
    const start = abaLog.getLastRow() + 1;
    abaLog.getRange(start, 1, 1, 3).setValues([[timestamp, nome, url]]);

    // Manter no máx. 1 cabeçalho + 200 dados
    const maxTotal = 201;
    const total = abaLog.getLastRow();
    if (total > maxTotal) {
      const excedente = total - maxTotal;
      // remove do topo (mais antigos), preservando cabeçalho
      abaLog.deleteRows(2, excedente);
    }
  } catch(e) {
    console.error("Falha ao registrar arquivo no LogDeArquivos: " + e.message);
  }
}

function _atualizarPainel(status, resultado) {
  var tz = 'America/Sao_Paulo';
  var ts = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    panel_status: status,
    panel_result: resultado,
    panel_ts: ts
  }, true);
}

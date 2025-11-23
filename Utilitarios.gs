/***********************************************************************************************************************************
 * ARQUIVO: Utilitarios.gs
 ***********************************************************************************************************************************/

/**
 * Encontra a última linha com dados de forma otimizada.
 * Ignora formatações ou checkboxes vazios, foca no valor real.
 */
function _encontrarUltimaLinhaNaColuna(aba, indiceColuna) {
  const maxRow = aba.getLastRow();
  if (maxRow === 0) return 1;
  
  // Pega apenas a coluna desejada para verificar
  const dados = aba.getRange(1, indiceColuna, maxRow, 1).getValues();
  
  // Varre de baixo para cima até achar algo escrito
  for (let i = dados.length - 1; i >= 0; i--) {
    if (dados[i][0] !== "" && dados[i][0] != null) {
      return i + 1;
    }
  }
  return 1;
}

function _escreverStatus(mensagem) {
  console.log(mensagem);
  PropertiesService.getScriptProperties().setProperty('export_status', mensagem);
}

/**
 * Registra logs na aba 'Log'.
 * Correção: Usa a Coluna A (1) para definir onde colar, evitando pular linhas com checkbox preenchido.
 */
function _registrarArquivoGeradoDetalhado(arquivo, dadosExportados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaLog = ss.getSheetByName('Log'); 
    
    if (!abaLog) {
      console.warn("Aba 'Log' não encontrada.");
      return; 
    }

    const ts = new Date();
    const nomeArquivo = arquivo.getName();
    const url = arquivo.getUrl();

    // Monta matriz para escrita
    const linhas = dadosExportados.map(linha => {
      // [0:SKU, 1:Nome, 2:NCM, 3:vazio, 4:GTIN]
      return [
        ts,           // A: Timestamp
        nomeArquivo,  // B: Nome Arquivo
        url,          // C: URL
        linha[0],     // D: SKU
        linha[1],     // E: Nome
        linha[4],     // F: GTIN
        linha[2]      // G: NCM
      ];
    });

    if (linhas.length > 0) {
      // === CORREÇÃO AQUI ===
      // Busca a última linha baseada na Coluna A (Data), ignorando a Coluna H (Checkbox)
      const startRow = _encontrarUltimaLinhaNaColuna(abaLog, 1) + 1;
      
      // Escreve colunas A até G (7 colunas)
      // A Coluna H (Checkbox) você já deixou preenchida, então o script não precisa tocar nela.
      abaLog.getRange(startRow, 1, linhas.length, 7).setValues(linhas);
      
      // Limpeza automática (mantém ~400 logs para não pesar)
      const maxLinhas = 401;
      const ultimaLinhaReal = _encontrarUltimaLinhaNaColuna(abaLog, 1);
      if (ultimaLinhaReal > maxLinhas + 50) {
         // Deleta linhas antigas (isso vai apagar seus checkboxes dessas linhas também, o que é bom para limpeza)
         abaLog.deleteRows(2, ultimaLinhaReal - maxLinhas);
      }
    }
  } catch(e) {
    console.error("Erro no log: " + e.message);
  }
}

function _atualizarPainel(status, resultado) {
  const tz = 'America/Sao_Paulo';
  const ts = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  PropertiesService.getScriptProperties().setProperties({
    panel_status: status,
    panel_result: resultado,
    panel_ts: ts
  }, true);
}

function _nomeMesPtBR(m) {
  const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  return meses[m-1] || String(m);
}

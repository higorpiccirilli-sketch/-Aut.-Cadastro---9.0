/***********************************************************************************************************************************
 * ARQUIVO: Exportacao.gs
 ***********************************************************************************************************************************/

// =================================================================================
// BLOCO 1: EXPORTAÇÃO DE PRODUTOS (Coluna F) - MANTIDO
// =================================================================================

function executarExportacaoManual() {
  let abaInfo, linhasSelecionadas = [];

  try {
    _escreverStatus('[1/5] Lendo seleção (Produtos - Col F)...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    abaInfo = ss.getSheetByName(CONFIG.PLANILHA_DESTINO.INFORMACOES.NOME);
    if (!abaInfo) throw new Error("Aba 'Informações' não encontrada.");

    const idxF = CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_CADASTRO_MANUAL_INDICE_0;
    const resultado = _coletarLinhasSelecionadas(abaInfo, idxF);
    
    linhasSelecionadas = resultado.linhasAbsolutas;
    const dadosBrutos = resultado.dados;

    if (dadosBrutos.length === 0) throw new Error("Nenhum produto selecionado na Coluna F.");
    Utilities.sleep(300);

    _escreverStatus(`[2/5] Validando ${dadosBrutos.length} itens...`);
    const validacao = _validarDadosComRegraGtin(dadosBrutos);
    if (validacao.validos.length === 0) throw new Error("Todos os itens falharam na validação.");

    if (validacao.itensSemGtin > 0) {
      const ui = SpreadsheetApp.getUi();
      const resposta = ui.alert('Atenção: GTINs Vazios', `Existem ${validacao.itensSemGtin} produtos sem GTIN. Deseja continuar?`, ui.ButtonSet.YES_NO);
      if (resposta !== ui.Button.YES) throw new Error("Cancelado pelo usuário.");
    }
    Utilities.sleep(300);

    _escreverStatus('[3/5] Verificando duplicidades...');
    const itensUnicos = _filtrarDuplicidades(validacao.validos);
    if (itensUnicos.length === 0) throw new Error("Todos os itens barrados por duplicidade.");
    Utilities.sleep(300);

    _escreverStatus('[4/5] Iniciando geração de arquivos...');
    
    _gerarProcessoCompleto('Petiko', itensUnicos, CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS.NOME);

    const itensPaws = itensUnicos.filter(i => String(i[CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_INDUSTRIA_INDICE_0]).trim().toUpperCase() === 'PAWS');
    if (itensPaws.length > 0) {
      _gerarProcessoCompleto('Paws', itensPaws, CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS_2.NOME);
    }

    const itensInnova = itensUnicos.filter(i => String(i[CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_INDUSTRIA_INDICE_0]).trim().toUpperCase() === 'INNOVA');
    if (itensInnova.length > 0) {
      _gerarProcessoCompleto('Innova', itensInnova, CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS_2.NOME);
    }

    const totalArquivos = 1 + (itensPaws.length>0?1:0) + (itensInnova.length>0?1:0);
    const msgFinal = `Processo concluído! ${totalArquivos} arquivo(s) de PRODUTOS gerado(s).`;
    
    _escreverStatus(`[OK] ${msgFinal}`);
    _atualizarPainel('Sucesso ✅', msgFinal);
    
    return msgFinal;

  } catch (e) {
    _escreverStatus(`[ERRO] ${e.message}`);
    _atualizarPainel('Erro ❌', e.message);
    throw e;
  } finally {
    if (abaInfo && linhasSelecionadas.length > 0) {
      const idxF = CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_CADASTRO_MANUAL_INDICE_0;
      _limparCheckboxEspecifico(abaInfo, linhasSelecionadas, idxF);
    }
    Utilities.sleep(500);
    PropertiesService.getScriptProperties().deleteProperty('export_status');
  }
}

function _gerarProcessoCompleto(nomeGrupo, dados, nomeAbaPalcoLocal) {
  _escreverStatus(`>> Gerando arquivo: ${nomeGrupo} (${dados.length} itens)...`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPalcoLocal = ss.getSheetByName(nomeAbaPalcoLocal);
  if (!abaPalcoLocal) throw new Error(`Aba palco local '${nomeAbaPalcoLocal}' não encontrada.`);

  abaPalcoLocal.getRange("B6:G200").clearContent();

  const idx = CONFIG.PLANILHA_DESTINO.INFORMACOES;
  const dadosInput = dados.map(r => [
    r[idx.COL_SKU_INDICE_0],
    r[idx.COL_NOME_INDICE_0],
    r[idx.COL_NCM_INDICE_0],
    '',
    r[idx.COL_GTIN_INDICE_0]
  ]);
  
  if (dadosInput.length > 0) {
    abaPalcoLocal.getRange(6, 3, dadosInput.length, 5).setValues(dadosInput);
  }
  SpreadsheetApp.flush(); 

  const valoresFinais = abaPalcoLocal.getRange("B6:AN200").getValues();

  const idTemplate = CONFIG.PLANILHA_DESTINO.ID_TEMPLATE_EXPORTACAO;
  const templateSheet = SpreadsheetApp.openById(idTemplate);
  const abaDestinoTemplate = templateSheet.getSheetByName(CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS.NOME);
  
  // Limpeza Cruzada
  const abaCaracteristicasTemplate = templateSheet.getSheetByName(CONFIG.PLANILHA_DESTINO.OMIE_CARACTERISTICAS.NOME);
  if (abaCaracteristicasTemplate) {
    abaCaracteristicasTemplate.getRange("B6:D200").clearContent();
  }

  abaDestinoTemplate.getRange("B6:AN200").clearContent();
  SpreadsheetApp.flush();
  abaDestinoTemplate.getRange("B6:AN200").setValues(valoresFinais);
  SpreadsheetApp.flush();

  _salvarNoDrive(nomeGrupo, idTemplate, dadosInput); 
}

// =================================================================================
// BLOCO 2: EXPORTAÇÃO DE CARACTERÍSTICAS (Coluna H)
// =================================================================================

function obterItensSelecionadosParaModal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaInfo = ss.getSheetByName(CONFIG.PLANILHA_DESTINO.INFORMACOES.NOME);
  const idxH = CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_CADASTRO_CARACTERISTICAS_INDICE_0;
  
  const resultado = _coletarLinhasSelecionadas(abaInfo, idxH);
  if (resultado.dados.length === 0) return [];
  
  const idx = CONFIG.PLANILHA_DESTINO.INFORMACOES;
  return resultado.dados.map(r => ({
    sku: r[idx.COL_SKU_INDICE_0],
    nome: r[idx.COL_NOME_INDICE_0],
  }));
}

function executarGeracaoCaracteristicas(dadosDoModal) {
  let abaInfo, linhasSelecionadas = [];
  
  try {
    _escreverStatus('[1/4] Processando dados do formulário...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    abaInfo = ss.getSheetByName(CONFIG.PLANILHA_DESTINO.INFORMACOES.NOME);
    
    const idxH = CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_CADASTRO_CARACTERISTICAS_INDICE_0;
    const sel = _coletarLinhasSelecionadas(abaInfo, idxH);
    linhasSelecionadas = sel.linhasAbsolutas;

    const dadosExplodidos = [];
    
    dadosDoModal.forEach(item => {
      const sku = item.sku;
      const nomeProd = String(item.nome || '');
      const tema = String(item.tema || '').trim();

      dadosExplodidos.push([sku, 'linha-comercial', 'Sim']);
      dadosExplodidos.push([sku, 'classificacao', 'BETTER']);
      dadosExplodidos.push([sku, 'custos-adicionais', '0,65']);

      const match = nomeProd.match(/\s-\s([A-Z0-9]+)\s-\s/);
      if (match) {
        const tamanho = match[1]; 
        if (tamanho !== 'U') { 
          dadosExplodidos.push([sku, 'tamanho', tamanho]);
        }
      }

      if (tema !== '') {
        dadosExplodidos.push([sku, 'Tema', tema]);
      }
    });

    if (dadosExplodidos.length === 0) throw new Error("Nenhuma característica gerada.");

    _escreverStatus(`[2/4] Gerando ${dadosExplodidos.length} linhas de características...`);
    
    const abaPalcoChar = ss.getSheetByName(CONFIG.PLANILHA_DESTINO.OMIE_CARACTERISTICAS.NOME);
    if (!abaPalcoChar) throw new Error("Aba local 'Omie_Produtos_Caracteristicas' não encontrada.");
    
    abaPalcoChar.getRange("B6:D500").clearContent(); 
    abaPalcoChar.getRange(6, 2, dadosExplodidos.length, 3).setValues(dadosExplodidos);
    SpreadsheetApp.flush();

    const idTemplate = CONFIG.PLANILHA_DESTINO.ID_TEMPLATE_EXPORTACAO;
    const templateSheet = SpreadsheetApp.openById(idTemplate);
    const abaDestinoTemplate = templateSheet.getSheetByName(CONFIG.PLANILHA_DESTINO.OMIE_CARACTERISTICAS.NOME);
    if (!abaDestinoTemplate) throw new Error("Aba Template Caracteristicas não encontrada.");

    // Limpeza Cruzada
    const abaProdutosTemplate = templateSheet.getSheetByName(CONFIG.PLANILHA_DESTINO.OMIE_PRODUTOS.NOME);
    if (abaProdutosTemplate) abaProdutosTemplate.getRange("B6:AN200").clearContent();

    abaDestinoTemplate.getRange("B6:D500").clearContent();
    SpreadsheetApp.flush();
    abaDestinoTemplate.getRange(6, 2, dadosExplodidos.length, 3).setValues(dadosExplodidos);
    SpreadsheetApp.flush();

    _escreverStatus('[3/4] Salvando arquivo...');
    
    // === CORREÇÃO: ESTRUTURA PARA O LOG (SKU, NOME, NCM, VAZIO, GTIN) ===
    // O utilitário espera: [0:SKU, 1:Nome, 2:NCM, 3:vazio, 4:GTIN]
    const dadosLog = dadosDoModal.map(d => [
      d.sku,                  // 0: SKU
      d.nome + ' (CARAC)',    // 1: Nome (Indicativo)
      '-',                    // 2: NCM (Não tem)
      '',                     // 3: Vazio
      '-'                     // 4: GTIN (Não tem)
    ]);
    
    _salvarNoDrive("Petiko_Caracteristicas", idTemplate, dadosLog);

    const msg = `Sucesso! Arquivo de Características gerado.`;
    _escreverStatus(`[OK] ${msg}`);
    return msg;

  } catch (e) {
    _escreverStatus(`[ERRO] ${e.message}`);
    throw e;
  } finally {
    if (abaInfo && linhasSelecionadas.length > 0) {
      const idxH = CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_CADASTRO_CARACTERISTICAS_INDICE_0;
      _limparCheckboxEspecifico(abaInfo, linhasSelecionadas, idxH);
    }
    Utilities.sleep(500);
    PropertiesService.getScriptProperties().deleteProperty('export_status');
  }
}


// =================================================================================
// BLOCO 3: HELPER COMPARTILHADO DE SALVAMENTO (ATUALIZADO PASTA)
// =================================================================================

function _salvarNoDrive(nomeGrupo, idTemplate, dadosParaLog) {
  const isCarac = (nomeGrupo === "Petiko_Caracteristicas");
  const chavePasta = isCarac ? "PETIKO" : nomeGrupo.toUpperCase();
  
  const idPastaDestino = CONFIG.ID_PASTA_EXPORTACAO[chavePasta];
  if (!idPastaDestino) throw new Error(`Pasta não encontrada para ${nomeGrupo}`);

  const tz = 'America/Sao_Paulo';
  const agora = new Date();
  const ano = Utilities.formatDate(agora, tz, 'yyyy');
  const mesNome = `${Utilities.formatDate(agora, tz, 'MM')}-${_nomeMesPtBR(parseInt(Utilities.formatDate(agora, tz, 'MM'),10))}`;

  const pastaEmpresa = DriveApp.getFolderById(idPastaDestino);
  const pastaAno = _obterOuCriarSubpasta(pastaEmpresa, ano);
  let pastaMes = _obterOuCriarSubpasta(pastaAno, mesNome);

  // === CORREÇÃO: CRIA SUBPASTA "Caracteristica" SE NECESSÁRIO ===
  if (isCarac) {
    pastaMes = _obterOuCriarSubpasta(pastaMes, "Caracteristica");
  }

  const hojeStr = Utilities.formatDate(agora, tz, 'dd-MM-yyyy');
  const hora = Utilities.formatDate(agora, tz, 'HH-mm-ss');
  
  // === CORREÇÃO: NOMES ===
  let nomeArquivo = "";
  if (isCarac) {
    // Padrão novo: Caracteristica - Data Hora
    nomeArquivo = `Caracteristica - ${hojeStr} ${hora}.xlsx`;
  } else {
    // Padrão produto: 01_Grupo_Data_Hora
    let contador = 0;
    const arquivos = pastaMes.getFiles();
    while (arquivos.hasNext()) {
      const nome = arquivos.next().getName();
      if (nome.includes(nomeGrupo) && nome.includes(hojeStr)) {
        contador++;
      }
    }
    const seq = (contador + 1).toString().padStart(2, '0');
    nomeArquivo = `${seq}_${nomeGrupo}_${hojeStr}_${hora}.xlsx`;
  }

  const url = `https://docs.google.com/spreadsheets/d/${idTemplate}/export?format=xlsx`;
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  
  if (resp.getResponseCode() !== 200) throw new Error('Erro no download: ' + resp.getContentText());

  const arquivoSalvo = pastaMes.createFile(resp.getBlob()).setName(nomeArquivo);
  
  _registrarArquivoGeradoDetalhado(arquivoSalvo, dadosParaLog);
}

// === UTILS === (MANTIDOS)
function _coletarLinhasSelecionadas(aba, indiceColunaCheckbox) {
  const idx = CONFIG.PLANILHA_DESTINO.INFORMACOES;
  const baseRow = idx.PRIMEIRA_LINHA_DADOS;
  const ultimaLinha = _encontrarUltimaLinhaNaColuna(aba, 1); 
  
  if (ultimaLinha < baseRow) return { dados: [], linhasAbsolutas: [] };

  const colCheck = (indiceColunaCheckbox !== undefined) ? indiceColunaCheckbox : idx.COL_CADASTRO_MANUAL_INDICE_0;
  const larguraLeitura = Math.max(colCheck, idx.COL_CADASTRO_MANUAL_INDICE_0) + 1;
  
  const range = aba.getRange(baseRow, 1, ultimaLinha - baseRow + 1, larguraLeitura);
  const valores = range.getValues();
  
  const selecionados = [];
  const linhas = [];

  for (let i = 0; i < valores.length; i++) {
    if (valores[i][colCheck] === true) {
      selecionados.push(valores[i]);
      linhas.push(baseRow + i);
    }
  }
  return { dados: selecionados, linhasAbsolutas: linhas };
}

function _limparCheckboxEspecifico(aba, linhas, indiceColuna) {
  if (!linhas || linhas.length === 0) return;
  let letra = 'F';
  if (indiceColuna === 7) letra = 'H';
  const rangeList = linhas.map(l => `${letra}${l}`);
  aba.getRangeList(rangeList).uncheck();
}

function _validarDadosComRegraGtin(itens) {
  const idx = CONFIG.PLANILHA_DESTINO.INFORMACOES;
  const validos = [];
  let qtdSemGtin = 0;
  itens.forEach(row => {
    const sku = row[idx.COL_SKU_INDICE_0];
    const ncm = row[idx.COL_NCM_INDICE_0];
    const gtin = row[idx.COL_GTIN_INDICE_0];
    if (!_campoPreenchido(sku) || !_campoPreenchido(ncm)) return;
    const itemProc = [...row];
    if (!_campoPreenchido(gtin)) {
      qtdSemGtin++;
      itemProc[idx.COL_GTIN_INDICE_0] = ""; 
    }
    validos.push(itemProc);
  });
  return { validos: validos, itensSemGtin: qtdSemGtin };
}

function _campoPreenchido(val) {
  if (val == null) return false;
  const s = String(val).trim();
  return s !== "" && s.toUpperCase() !== "FALTANDO";
}

function _filtrarDuplicidades(itens) {
  const skus = new Set();
  const unicos = [];
  const idx = CONFIG.PLANILHA_DESTINO.INFORMACOES;
  itens.forEach(row => {
    const sku = String(row[idx.COL_SKU_INDICE_0]).trim();
    if (!skus.has(sku)) {
      skus.add(sku);
      unicos.push(row);
    }
  });
  return unicos;
}

function _obterOuCriarSubpasta(pai, nome) {
  var it = pai.getFoldersByName(nome);
  return it.hasNext() ? it.next() : pai.createFolder(nome);
}

function _limparCheckboxF(aba, linhas) {
  const idxF = CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_CADASTRO_MANUAL_INDICE_0;
  _limparCheckboxEspecifico(aba, linhas, idxF);
}

function _nomeMesPtBR(m) {
  const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  return meses[m-1] || String(m);
}

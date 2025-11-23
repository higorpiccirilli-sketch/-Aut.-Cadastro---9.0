/***********************************************************************************************************************************
 * ARQUIVO: InterfaceBackend.gs
 ***********************************************************************************************************************************/

function onOpen() {
  if (typeof iniciarInterfacePrincipal === 'function') iniciarInterfacePrincipal();
  if (typeof createMetabaseMenu === 'function') createMetabaseMenu();
}

function iniciarInterfacePrincipal() {
  SpreadsheetApp.getUi()
    .createMenu('▶️ Painel de Controle')
    .addItem('Abrir Painel', 'exibirPainelDeControle')
    .addSeparator()
    .addItem('🔄 Sincronizar Manualmente', 'importarDadosEConsultarAbas')
    .addItem('🗂️ Gerenciar Arquivos', 'abrirGerenciadorDeArquivos')
    .addSeparator()
    .addItem('▶️ Gerar Arquivo de PRODUTOS', 'iniciarGeracaoManual')
    .addItem('▶️ Gerar Arquivo de CARACTERÍSTICAS', 'iniciarGeracaoCaracteristicas')
    .addSeparator()
    .addItem('🗑️ Excluir Arquivos Selecionados (Log)', 'menuExcluirArquivos')
    .addToUi();
  exibirPainelDeControle();
}

function iniciarGeracaoCaracteristicas() {
  try {
    const dados = obterItensSelecionadosParaModal(); 
    if (!dados || dados.length === 0) {
      SpreadsheetApp.getUi().alert('Aviso', 'Selecione pelo menos um produto (Coluna H) na aba "Cadastro Petiko".', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    const html = HtmlService.createTemplateFromFile('ModalCaracteristicas.html');
    html.dadosIniciais = JSON.stringify(dados); 
    SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(700).setHeight(600), 'Definir Características');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro: ' + e.message);
  }
}

function processarGeracaoCaracteristicasDoFrontend(dadosDoFormulario) {
  return executarGeracaoCaracteristicas(dadosDoFormulario);
}

function menuExcluirArquivos() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Lixeira do Drive', 'Tem certeza que deseja mover para a lixeira os arquivos marcados na aba "Log"?\n\n(A linha será riscada).', ui.ButtonSet.YES_NO);
  if (resp === ui.Button.YES) {
    try {
      const msg = excluirArquivosSelecionados();
      ui.alert('Resultado', msg, ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Erro', e.message, ui.ButtonSet.OK);
    }
  }
}

function exibirPainelDeControle() {
  const html = HtmlService.createTemplateFromFile('PainelDeControle.html').evaluate().setTitle('Painel de Controle OMIE');
  SpreadsheetApp.getUi().showSidebar(html);
}

function incluir(nomeArquivo) {
  return HtmlService.createHtmlOutputFromFile(nomeArquivo).getContent();
}

function iniciarGeracaoManual() {
  const html = HtmlService.createTemplateFromFile('Log.html').evaluate().setWidth(700).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Log da Exportação Manual');
}

function abrirGerenciadorDeArquivos() {
  const html = HtmlService.createHtmlOutputFromFile('GerenciadorDeArquivos.html').setWidth(600).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gerenciador de Arquivos Exportados');
}

// === API DO PAINEL ===

function obterDadosDoPainel() {
  const p = PropertiesService.getScriptProperties();
  return {
    status: p.getProperty('panel_status') || '--',
    resultado: p.getProperty('panel_result') || '--',
    timestamp: p.getProperty('panel_ts') || '--'
  };
}

function obterStatusDaExportacao() {
  return PropertiesService.getScriptProperties().getProperty('export_status');
}

function obterDadosGerenciador() {
  const ids = CONFIG.ID_PASTA_EXPORTACAO;
  const links = {
    petiko: DriveApp.getFolderById(ids.PETIKO).getUrl(),
    innova: DriveApp.getFolderById(ids.INNOVA).getUrl(),
    paws: DriveApp.getFolderById(ids.PAWS).getUrl()
  };
  const arqs = _lerLogDeArquivos(10); 
  return { linksPastas: links, arquivosRecentes: arqs };
}

function obterUltimosArquivosParaPainel(qtd) {
  return { arquivos: _lerLogDeArquivos(qtd || 5) };
}

// === CORREÇÃO AQUI: IGNORAR CHECKBOX VAZIO ===
function _lerLogDeArquivos(qtdLimite) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Log'); 
    if (!sh) return [];
    
    // CORREÇÃO: Usa a coluna A (1) para achar a ultima linha REAL
    // Em vez de sh.getLastRow() que pegaria até a linha 1000 por causa dos checkboxes
    const last = _encontrarUltimaLinhaNaColuna(sh, 1);
    
    if (last <= 1) return [];

    const start = Math.max(2, last - 200); 
    const rows = last - start + 1;
    const data = sh.getRange(start, 1, rows, 3).getValues(); // Lê A, B, C
    
    const itens = [];
    const vistos = new Set();
    
    for (let i = data.length - 1; i >= 0 && itens.length < qtdLimite; i--) {
      const row = data[i];
      const ts = row[0];
      const nome = String(row[1] || '');
      const url = String(row[2] || '');
      
      if (!url || !url.startsWith('http')) continue;
      if (vistos.has(url)) continue;
      vistos.add(url);

      let nomeCurto = nome;
      // Ajuste para exibir "Caracteristica" ou "05_Grupo" bonitinho
      const match = nome.match(/^([0-9]+)_([^_]+)/); 
      
      if (match) {
        nomeCurto = `${match[1]}_${match[2]}`; 
      } else if (nome.toLowerCase().includes('caracteristica')) {
        nomeCurto = "Caracteristica";
      } else {
        if (nome.length > 20) nomeCurto = nome.substring(0, 20) + '...';
      }
      
      nomeCurto = nomeCurto.replace('.xlsx', '');

      let dlUrl = url;
      const idMatch = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (idMatch) dlUrl = `https://drive.google.com/uc?export=download&id=${idMatch[1]}`;

      itens.push({
        timestamp: _formatarDataSimples(ts),
        nome: nome,
        nomeCurto: nomeCurto,
        url: url,
        downloadUrl: dlUrl
      });
    }
    return itens;
  } catch (e) {
    console.error(e);
    return [];
  }
}

function _formatarDataSimples(dateObj) {
  try {
    return Utilities.formatDate(new Date(dateObj), 'America/Sao_Paulo', 'dd/MM HH:mm');
  } catch(e) { return '--'; }
}

function excluirArquivosSelecionados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('Log');
  if (!aba) throw new Error("Aba 'Log' não encontrada.");
  
  // Aqui tbm usamos a correção para não ler linhas vazias desnecessárias
  const lastRow = _encontrarUltimaLinhaNaColuna(aba, 1);
  if (lastRow < 2) return "O Log está vazio.";
  
  const rangeDados = aba.getRange(2, 1, lastRow - 1, 8);
  const valores = rangeDados.getValues();
  let contExcluidos = 0;
  const linhasParaRiscar = []; 
  const celulasCheckbox = [];  
  for (let i = 0; i < valores.length; i++) {
    const isChecked = valores[i][7]; 
    const url = valores[i][2];       
    const linhaAbsoluta = i + 2;
    if (isChecked === true) {
      try {
        const fileId = _extrairIdParaExclusao(url);
        if (fileId) {
          DriveApp.getFileById(fileId).setTrashed(true);
          contExcluidos++;
          linhasParaRiscar.push(`A${linhaAbsoluta}:H${linhaAbsoluta}`);
          celulasCheckbox.push(`H${linhaAbsoluta}`);
        }
      } catch (e) {
        console.warn(`Erro ao excluir: ${e.message}`);
        if (String(e.message).includes("not found") || String(e.message).includes("não encontrado")) {
          contExcluidos++;
          linhasParaRiscar.push(`A${linhaAbsoluta}:H${linhaAbsoluta}`);
          celulasCheckbox.push(`H${linhaAbsoluta}`);
        }
      }
    }
  }
  if (linhasParaRiscar.length > 0) {
    aba.getRangeList(linhasParaRiscar).setFontLine("line-through").setFontColor("#e06c75");     
    aba.getRangeList(celulasCheckbox).setDataValidation(null).setValue("LIXEIRA");
  }
  if (contExcluidos === 0) return "Nenhum arquivo selecionado ou encontrado.";
  return `${contExcluidos} arquivo(s) movido(s) para a lixeira.`;
}

function _extrairIdParaExclusao(url) {
  const match = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

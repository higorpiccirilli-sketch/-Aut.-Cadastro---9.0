/***********************************************************************************************************************************
 *
 * NOME DO ARQUIVO: Sincronizacao.gs
 *
 ***********************************************************************************************************************************/

/**
 * @scriptName Lógica de Sincronização de Dados (manual)
 * @version 1.0 enxuto
 * @description
 * Lê a planilha de origem (CONFIG.URL_ORIGEM) e sincroniza com a aba "Cadastro Petiko":
 * - Adiciona produtos novos (A:SKU, B:Nome, C:NCM, D:GTIN) — não escreve F/G.
 * - Atualiza dados faltantes (SKU/NCM/GTIN) de produtos já existentes.
 *
 * Observações:
 * - NÃO grava nada em A2/B2/C2 (ou A3/B3/C3).
 * - NÃO toca em F (Cadastro Manual) nem G (Reservado/Características).
 * - “Indústria” (coluna E) fica a critério do usuário (não é alterada aqui).
 *
 * Dependências:
 * - CONFIG (Config.gs)
 * - Utilitarios.gs: _encontrarUltimaLinhaNaColuna, _atualizarPainel
 * - Serviços do Google: SpreadsheetApp
 */

// =================================================================================
// Bloco 1: Wrappers (entrada manual)
// =================================================================================

/**
 * Ponto de entrada para a sincronização manual.
 * Exibe confirmação e um resumo ao final.
 */
function importarDadosEConsultarAbas() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Confirmação de Execução',
    'Deseja iniciar a sincronização de produtos? (Isso irá adicionar novos e atualizar os incompletos)',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) {
    ui.alert('A execução foi cancelada pelo usuário.');
    return;
  }

  try {
    const resumo = _executarLogicaDeSincronizacao();
    const msg =
      'Sincronização Concluída!\n\n' +
      `- Produtos Novos Adicionados: ${resumo.adicionados}\n` +
      `- Produtos Existentes Atualizados: ${resumo.atualizados}`;

    // Atualiza o painel
    _atualizarPainel('Sucesso ✅', `Sincronização: ${resumo.adicionados} novo(s), ${resumo.atualizados} atualizado(s).`);

    ui.alert(msg);
  } catch (e) {
    // Atualiza o painel com erro
    _atualizarPainel('Erro ❌', e && e.message ? e.message : String(e));
    ui.alert(`Ocorreu um erro durante a execução: ${e.name} - ${e.message}`);
  }
}

// =================================================================================
/**
 * Faz a sincronização com base nos nomes listados em ABAS_ORIGEM.ACOMPANHAMENTO.
 * @returns {{adicionados:number, atualizados:number}}
 */
function _executarLogicaDeSincronizacao() {
  // Abre planilha de origem e destino
  const planilhaOrigem = SpreadsheetApp.openByUrl(CONFIG.URL_ORIGEM);
  const planilhaDestino = SpreadsheetApp.getActiveSpreadsheet();
  const abaInfo = planilhaDestino.getSheetByName(CONFIG.PLANILHA_DESTINO.INFORMACOES.NOME);
  if (!abaInfo) throw new Error(`A aba de destino "${CONFIG.PLANILHA_DESTINO.INFORMACOES.NOME}" não foi encontrada.`);

  // Constrói mapa com dados (nome-> {nome, sku, ncm, gtin})
  const mapaDadosOrigem = _construirMapaDeDados(planilhaOrigem);

  // Lê o que já existe na aba de destino
  const primeiraLinha = CONFIG.PLANILHA_DESTINO.INFORMACOES.PRIMEIRA_LINHA_DADOS;
  const ultimaLinhaInfo = _encontrarUltimaLinhaNaColuna(abaInfo, 1); // coluna A
  const larguraLeitura = 4; // A..D (SKU, Nome, NCM, GTIN)

  const dadosExistentes =
    (ultimaLinhaInfo >= primeiraLinha)
      ? abaInfo.getRange(
          primeiraLinha,
          1,
          (ultimaLinhaInfo - primeiraLinha + 1),
          larguraLeitura
        ).getValues()
      : [];

  // Mapa de existentes por nome (chave: nome normalizado)
  const mapaExistentes = new Map();
  dadosExistentes.forEach((linha, index) => {
    const nomeProduto = linha[CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_NOME_INDICE_0]; // B (0-based)
    if (nomeProduto) {
      mapaExistentes.set(
        String(nomeProduto).trim().toLowerCase(),
        {
          linha,
          numeroLinha: CONFIG.PLANILHA_DESTINO.INFORMACOES.PRIMEIRA_LINHA_DADOS + index
        }
      );
    }
  });

  // Lista de nomes para sincronizar (Acompanhamento!A3:A1001)
  const nomesParaSincronizar = planilhaOrigem
    .getSheetByName(CONFIG.ABAS_ORIGEM.ACOMPANHAMENTO.NOME)
    .getRange(CONFIG.ABAS_ORIGEM.ACOMPANHAMENTO.INTERVALO_NOMES)
    .getValues()
    .flat()
    .map(n => String(n || '').trim())
    .filter(n => n);

  const produtosParaAdicionar = [];
  const produtosParaAtualizar = [];

  // Decide adicionar/atualizar
  for (const nome of nomesParaSincronizar) {
    const dadosOrigem = mapaDadosOrigem.get(nome); // chave exatamente como no mapa (nome “limpo” com trim)
    if (!dadosOrigem) {
      console.warn(`Produto "${nome}" listado no Acompanhamento mas não encontrado nas abas de dados.`);
      continue;
    }

    const chave = String(nome).trim().toLowerCase();
    if (mapaExistentes.has(chave)) {
      // atualizar somente campos “FALTANDO”
      const existente = mapaExistentes.get(chave);
      const linha = existente.linha;

      const precisaAtualizar = (
        (String(linha[CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_SKU_INDICE_0]  || '').toUpperCase()  === 'FALTANDO' && String(dadosOrigem.sku  || '').toUpperCase()  !== 'FALTANDO') ||
        (String(linha[CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_NCM_INDICE_0]  || '').toUpperCase()  === 'FALTANDO' && String(dadosOrigem.ncm  || '').toUpperCase()  !== 'FALTANDO') ||
        (String(linha[CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_GTIN_INDICE_0] || '').toUpperCase()  === 'FALTANDO' && String(dadosOrigem.gtin || '').toUpperCase() !== 'FALTANDO')
      );

      if (precisaAtualizar) {
        produtosParaAtualizar.push({
          numeroLinha: existente.numeroLinha,
          dadosNovos: dadosOrigem
        });
      }
    } else {
      // adicionar novo
      produtosParaAdicionar.push(dadosOrigem);
    }
  }

  // Atualiza existentes (apenas os campos que estavam “FALTANDO”)
  if (produtosParaAtualizar.length > 0) {
    produtosParaAtualizar.forEach(item => {
      const r = item.dadosNovos;
      const row = item.numeroLinha;
      // A: SKU
      if (String(r.sku || '').toUpperCase() !== 'FALTANDO') {
        abaInfo.getRange(row, CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_SKU_INDICE_0 + 1).setValue(r.sku);
      }
      // C: NCM
      if (String(r.ncm || '').toUpperCase() !== 'FALTANDO') {
        abaInfo.getRange(row, CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_NCM_INDICE_0 + 1).setValue(r.ncm);
      }
      // D: GTIN
      if (String(r.gtin || '').toUpperCase() !== 'FALTANDO') {
        abaInfo.getRange(row, CONFIG.PLANILHA_DESTINO.INFORMACOES.COL_GTIN_INDICE_0 + 1).setValue(r.gtin);
      }
    });
  }

  // Adiciona novos (somente A..D). Não escreve F/G.
  if (produtosParaAdicionar.length > 0) {
    const proximaLinha = _encontrarUltimaLinhaNaColuna(abaInfo, 1) + 1;
    const dadosAteD = produtosParaAdicionar.map(d => [d.sku, d.nome, d.ncm, d.gtin]); // A..D
    abaInfo.getRange(proximaLinha, 1, dadosAteD.length, 4).setValues(dadosAteD);
  }

  return {
    adicionados: produtosParaAdicionar.length,
    atualizados: produtosParaAtualizar.length
  };
}

// =================================================================================
// Bloco 3: Leitura da planilha de origem (consolidado)
// =================================================================================

/**
 * Lê todas as abas listadas em CONFIG.ABAS_ORIGEM.DADOS.NOMES e consolida em um mapa por “Nome” (coluna A).
 * @param {Spreadsheet} planilhaOrigem
 * @returns {Map<string, {nome:string, sku:string, ncm:string, gtin:string}>}
 */
function _construirMapaDeDados(planilhaOrigem) {
  const mapa = new Map();
  const abasDeDados = CONFIG.ABAS_ORIGEM.DADOS.NOMES;

  for (const nomeAba of abasDeDados) {
    const aba = planilhaOrigem.getSheetByName(nomeAba);
    if (!aba) continue;

    const nomes = aba.getRange(CONFIG.ABAS_ORIGEM.DADOS.INTERVALO_NOMES_PRODUTOS).getValues().flat();
    const skus  = aba.getRange(CONFIG.ABAS_ORIGEM.DADOS.INTERVALO_SKU).getValues().flat();
    const ncms  = aba.getRange(CONFIG.ABAS_ORIGEM.DADOS.INTERVALO_NCM).getValues().flat();
    const gtins = aba.getRange(CONFIG.ABAS_ORIGEM.DADOS.INTERVALO_GTIN).getValues().flat();

    for (let i = 0; i < nomes.length; i++) {
      const nome = String(nomes[i] || '').trim();
      if (!nome) continue;
      if (!mapa.has(nome)) {
        mapa.set(nome, {
          nome: nome,
          sku:  skus[i]  || 'FALTANDO',
          ncm:  ncms[i]  || 'FALTANDO',
          gtin: gtins[i] || 'FALTANDO'
        });
      }
    }
  }
  return mapa;
}

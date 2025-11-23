/***********************************************************************************************************************************
 * ARQUIVO: Sincronizacao.gs
 ***********************************************************************************************************************************/

function importarDadosEConsultarAbas() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Sincronização de Produtos',
    'Deseja buscar novos produtos e atualizar dados faltantes da planilha de origem?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    _atualizarPainel('Em Andamento ⏳', 'Lendo dados da origem...');
    const resumo = _executarLogicaDeSincronizacaoOtimizada();
    
    const msg = `Sincronização Concluída!\n\n🆕 Novos adicionados: ${resumo.adicionados}\n🔄 Existentes atualizados: ${resumo.atualizados}`;
    _atualizarPainel('Sucesso ✅', msg);
    ui.alert(msg);
    
  } catch (e) {
    _atualizarPainel('Erro ❌', e.message);
    ui.alert('Erro: ' + e.message);
  }
}

function _executarLogicaDeSincronizacaoOtimizada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaInfo = ss.getSheetByName(CONFIG.PLANILHA_DESTINO.INFORMACOES.NOME);
  if (!abaInfo) throw new Error("Aba de destino não encontrada.");

  // 1. Lê a Origem
  const planilhaOrigem = SpreadsheetApp.openByUrl(CONFIG.URL_ORIGEM);
  const mapaDadosOrigem = _construirMapaDeDados(planilhaOrigem);

  // 2. Lê o Destino (CORREÇÃO AQUI: Usa a coluna B para definir o fim real, ignorando fórmulas em H:AN)
  const startRow = CONFIG.PLANILHA_DESTINO.INFORMACOES.PRIMEIRA_LINHA_DADOS;
  
  // Usa a função utilitária para achar a última linha real baseada na Coluna 2 (Nome)
  // Se houver fórmulas na linha 2000, mas nome só até a 50, ele retorna 50.
  const lastRowReal = _encontrarUltimaLinhaNaColuna(abaInfo, 2); 
  
  let valoresDestino = [];
  
  if (lastRowReal >= startRow) {
    // Lê colunas A até D (SKU, Nome, NCM, GTIN) somente até onde tem dados reais
    valoresDestino = abaInfo.getRange(startRow, 1, lastRowReal - startRow + 1, 4).getValues();
  }

  // 3. Prepara Listas
  const nomesParaSincronizar = planilhaOrigem
    .getSheetByName(CONFIG.ABAS_ORIGEM.ACOMPANHAMENTO.NOME)
    .getRange(CONFIG.ABAS_ORIGEM.ACOMPANHAMENTO.INTERVALO_NOMES)
    .getValues()
    .flat()
    .map(n => String(n || '').trim())
    .filter(n => n);

  const novosItens = [];
  let contAtualizados = 0;
  
  // Índices base-0 da matriz lida (A=0, B=1, C=2, D=3)
  const I_SKU = 0; 
  const I_NOME = 1; 
  const I_NCM = 2; 
  const I_GTIN = 3;

  // Mapa para busca rápida no destino (Nome -> Índice do Array)
  const mapaDestinoIndices = new Map();
  valoresDestino.forEach((row, idx) => {
    const nomeKey = String(row[I_NOME] || '').trim().toLowerCase();
    if (nomeKey) mapaDestinoIndices.set(nomeKey, idx);
  });

  // 4. Processamento Lógico
  nomesParaSincronizar.forEach(nome => {
    const dadosNovos = mapaDadosOrigem.get(nome);
    if (!dadosNovos) return;

    const chave = String(nome).trim().toLowerCase();
    
    if (mapaDestinoIndices.has(chave)) {
      // --- ATUALIZAR ---
      const idxArr = mapaDestinoIndices.get(chave);
      const linha = valoresDestino[idxArr];
      let mudou = false;

      const ehFaltando = (v) => String(v || '').toUpperCase() === 'FALTANDO' || v === '';
      const temValor = (v) => String(v || '').toUpperCase() !== 'FALTANDO' && v !== '';

      if (ehFaltando(linha[I_SKU]) && temValor(dadosNovos.sku)) { linha[I_SKU] = dadosNovos.sku; mudou = true; }
      if (ehFaltando(linha[I_NCM]) && temValor(dadosNovos.ncm)) { linha[I_NCM] = dadosNovos.ncm; mudou = true; }
      if (ehFaltando(linha[I_GTIN]) && temValor(dadosNovos.gtin)) { linha[I_GTIN] = dadosNovos.gtin; mudou = true; }

      if (mudou) contAtualizados++;
    } else {
      // --- ADICIONAR NOVO ---
      novosItens.push([dadosNovos.sku, dadosNovos.nome, dadosNovos.ncm, dadosNovos.gtin]);
    }
  });

  // 5. Gravação em Lote
  
  // A) Atualizações (Reescreve o bloco existente)
  if (contAtualizados > 0 && valoresDestino.length > 0) {
    abaInfo.getRange(startRow, 1, valoresDestino.length, 4).setValues(valoresDestino);
  }

  // B) Novos (Append IMEDIATAMENTE após a última linha real de dados)
  if (novosItens.length > 0) {
    const nextRow = (lastRowReal < startRow) ? startRow : lastRowReal + 1;
    abaInfo.getRange(nextRow, 1, novosItens.length, 4).setValues(novosItens);
  }

  return { adicionados: novosItens.length, atualizados: contAtualizados };
}

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

/***********************************************************************************************************************************
 *
 * SUGESTÃO DE NOME PARA O ARQUIVO: Config.gs (MANTER)
 *
 ***********************************************************************************************************************************/

/**
 * @scriptName Configuração Global do Projeto
 * @version 1.1
 * @description
 * Centraliza URLs, IDs e nomes de abas/índices usados pelos demais scripts.
 */

const CONFIG = {
  // URL da planilha de ORIGEM (produtos)
  URL_ORIGEM: 'https://docs.google.com/spreadsheets/d/1s2rk-WBSPLV1Qf8X6Bl-oPW3rj4Vyv9RWCxdvPIsBX8/edit',

  // Pastas de exportação no Drive (por indústria)
  ID_PASTA_EXPORTACAO: {
    PETIKO: '1eQgDScdVZPYe6-Tg2yMZ45G7gsD2V7BY',
    INNOVA: '1x5Os57g76gdhT5N4FUNC3WF6yVvuyIMw',
    PAWS:   '1c1M0ndtRiW8OASMkdZVl6qjPlKYEuuhL'
  },

  // Planilha de origem (onde buscamos nomes e dados)
  ABAS_ORIGEM: {
    ACOMPANHAMENTO: {
      NOME: 'Acompanhamento',
      INTERVALO_NOMES: 'A3:A1001'
    },
    DADOS: {
      NOMES: ['COR.', 'PELU.', 'INTERA.', 'GATI.', 'ACE. KIT'],
      INTERVALO_NOMES_PRODUTOS: 'A4:A500',
      INTERVALO_SKU:  'G4:G500',
      INTERVALO_NCM:  'J4:J500',
      INTERVALO_GTIN: 'K4:K500'
    }
  },

  // Planilha de destino (onde gerimos e exportamos)
  PLANILHA_DESTINO: {
    ID_TEMPLATE_EXPORTACAO: '1ZxvOAuyhwkuzDnNI0BAuWyLaG4Se3LuHUAU9DMoTFOE',

    INFORMACOES: {
      NOME: 'Cadastro Petiko',
      PRIMEIRA_LINHA_DADOS: 5,

      // Índices base 0 (A=0, B=1, ...)
      COL_SKU_INDICE_0: 0,           // A
      COL_NOME_INDICE_0: 1,          // B
      COL_NCM_INDICE_0: 2,           // C
      COL_GTIN_INDICE_0: 3,          // D
      COL_INDUSTRIA_INDICE_0: 4,     // E

      // Exportação manual: checkbox fica na COLUNA F (base-0 = 5)
      COL_CADASTRO_MANUAL_INDICE_0: 5
      // Observação: a coluna G pode ser usada pelo gerador de características (seleção).
    },

    // Palcos de exportação (produtos)
    OMIE_PRODUTOS: {
      NOME: 'Omie_Produtos',
      PRIMEIRA_LINHA_DADOS: 6,
      COL_SKU: 3 // base 1 (C)
    },
    OMIE_PRODUTOS_2: {
      NOME: 'Omie_Produtos_2',
      PRIMEIRA_LINHA_DADOS: 6,
      COL_SKU: 3 // base 1 (C)
    },

    // Palco de características (somente colagem B:C:D a partir da linha 6)
    OMIE_CARACTERISTICAS: {
      NOME: 'Omie_Produtos_Caracteristicas',
      PRIMEIRA_LINHA_DADOS: 6,
      COL_CODIGO: 2 // base-1 (B)
    }
  }
};

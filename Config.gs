/***********************************************************************************************************************************
 * ARQUIVO: Config.gs
 ***********************************************************************************************************************************/

const CONFIG = {
  URL_ORIGEM: 'https://docs.google.com/spreadsheets/d/1s2rk-WBSPLV1Qf8X6Bl-oPW3rj4Vyv9RWCxdvPIsBX8/edit',

  ID_PASTA_EXPORTACAO: {
    PETIKO: '1eQgDScdVZPYe6-Tg2yMZ45G7gsD2V7BY',
    INNOVA: '1x5Os57g76gdhT5N4FUNC3WF6yVvuyIMw',
    PAWS:   '1c1M0ndtRiW8OASMkdZVl6qjPlKYEuuhL'
  },

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

  PLANILHA_DESTINO: {
    ID_TEMPLATE_EXPORTACAO: '1ZxvOAuyhwkuzDnNI0BAuWyLaG4Se3LuHUAU9DMoTFOE',

    INFORMACOES: {
      NOME: 'Cadastro Petiko',
      PRIMEIRA_LINHA_DADOS: 5,
      COL_SKU_INDICE_0: 0,        // A
      COL_NOME_INDICE_0: 1,       // B
      COL_NCM_INDICE_0: 2,        // C
      COL_GTIN_INDICE_0: 3,       // D
      COL_INDUSTRIA_INDICE_0: 4,  // E
      COL_CADASTRO_MANUAL_INDICE_0: 5, // F (Para Produtos)
      COL_CADASTRO_CARACTERISTICAS_INDICE_0: 7 // H (Para Características) <--- NOVO
    },

    OMIE_PRODUTOS: {
      NOME: 'Omie_Produtos',
      PRIMEIRA_LINHA_DADOS: 6,
      COL_SKU: 3
    },
    OMIE_PRODUTOS_2: {
      NOME: 'Omie_Produtos_2',
      PRIMEIRA_LINHA_DADOS: 6,
      COL_SKU: 3
    },
    OMIE_CARACTERISTICAS: {
      NOME: 'Omie_Produtos_Caracteristicas',
      PRIMEIRA_LINHA_DADOS: 6,
      COL_CODIGO_SKU: 2
    }
  }
};

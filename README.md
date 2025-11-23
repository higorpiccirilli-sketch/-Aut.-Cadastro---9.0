# 🚀 Automação de Gestão e Importação de Produtos - Omie ERP

Este projeto é uma solução robusta desenvolvida em **Google Apps Script** para automatizar o ciclo de vida do cadastro de produtos, desde a sincronização de dados brutos até a geração de arquivos `.xlsx` validados para importação no **ERP Omie** (sistema de gestão brasileiro).

O sistema resolve o problema de manipulação manual de planilhas, garantindo integridade de dados (SKU, NCM, GTIN), evitando duplicidades e organizando automaticamente os arquivos no Google Drive.

## 🎯 Objetivo Principal

Gerar planilhas modelos de importação para o **ERP Omie** de forma automática, segmentada por marca (Petiko, Paws, Innova) e validada, reduzindo drasticamente o tempo operacional e erros humanos no cadastro de produtos.

---

## 🛠️ Funcionalidades Principais

### 1. Sincronização Inteligente de Dados
- **Atualização Incremental:** O script lê uma planilha de "Origem", compara com a base local e identifica novos produtos.
- **Preenchimento de Lacunas:** Se um produto já existe localmente mas possui dados faltantes (ex: NCM "FALTANDO"), o script atualiza apenas esse campo, preservando edições manuais anteriores.
- **Performance:** Utiliza processamento em lote (Batch Processing) para ler e escrever milhares de linhas em segundos, minimizando chamadas à API do Google Sheets.

### 2. Exportação para Omie (.xlsx)
O sistema gera arquivos Excel formatados especificamente para o layout de importação do Omie.
- **Seleção Manual:** O usuário seleciona os itens desejados via Checkbox na planilha.
- **Validação Rígida:** Impede a geração se campos obrigatórios (SKU, NCM) estiverem vazios.
- **Validação Flexível:** Alerta o usuário caso existam itens sem GTIN, permitindo autorização manual.
- **Lógica Multi-Marca:**
  - **Arquivo Mestre (Petiko):** Contém *todos* os itens selecionados.
  - **Arquivos Segmentados (Paws/Innova):** Gera arquivos adicionais apenas se houverem produtos dessas marcas no lote.
- **Uso de Template:** Utiliza uma planilha "Molde" oculta para garantir que formatações, cabeçalhos e fórmulas complexas do Excel sejam preservados na exportação.

### 3. Organização Automática no Drive
- Cria automaticamente uma estrutura de pastas: `Empresa > Ano (YYYY) > Mês (MM-Nome)`.
- Nomenclatura padronizada: `SEQUENCIAL_MARCA_DATA_HORA.xlsx` (ex: `05_Innova_23-11-2025_14-30-00.xlsx`).

### 4. Interface e Gestão (Frontend)
- **Painel Lateral (Sidebar):** Controle central para disparar sincronizações e exportações.
- **Logs Detalhados:** Registro histórico de cada item exportado (incluindo NCM e Link direto).
- **Gerenciador de Arquivos (Lixeira):** Funcionalidade personalizada no Menu para mover arquivos gerados para a Lixeira do Drive e marcar visualmente (riscado) no log da planilha.

---

## 🧩 Arquitetura do Projeto

O código segue princípios de **Clean Code** e **Separação de Responsabilidades**:

| Arquivo | Responsabilidade |
| :--- | :--- |
| `Config.gs` | Centraliza IDs (Planilhas, Drive), URLs e mapeamento de colunas. Nenhuma configuração fica "hardcoded" na lógica. |
| `Sincronizacao.gs` | Lógica de leitura da origem, comparação de dados em memória e atualização em lote da base local. |
| `Exportacao.gs` | "Coração" do sistema. Valida dados, gerencia duplicidades, manipula o Template externo e salva no Drive. |
| `InterfaceBackend.gs` | Camada de comunicação entre o HTML (Sidebar/Modais) e o Google Apps Script. |
| `Utilitarios.gs` | Funções helpers reutilizáveis (busca de última linha otimizada, logs, formatação de data). |
| `PainelDeControle.html` | Interface gráfica do usuário (Sidebar). |

---

## ⚙️ Fluxo de Exportação (Deep Dive)

Para garantir que o arquivo final funcione no Omie e mantenha as fórmulas auxiliares, o script executa o seguinte pipeline:

1.  **Staging Local:** Limpa abas auxiliares (`Omie_Produtos`) na planilha atual e cola os dados brutos (SKU, Nome, etc.).
2.  **Cálculo:** Força o Google Sheets a recalcular fórmulas nessas abas auxiliares (ex: concatenações ou tratamentos de string necessários para o ERP).
3.  **Template Externo:** Abre uma planilha Template separada (ID fixo).
4.  **Deep Clean:** Limpa completamente a área de dados do Template.
5.  **Transferência:** Copia os valores calculados (Value-only) do Staging Local para o Template.
6.  **Exportação:** Usa a API de Drive (`UrlFetchApp`) para baixar o Template preenchido como binário `.xlsx`.
7.  **Salvamento:** Salva o arquivo na pasta correta do Drive e registra no Log.

---

## 💻 Tecnologias Utilizadas

*   **Google Apps Script (GAS):** Backend Serverless (V8 Runtime).
*   **Google Sheets API:** Manipulação avançada de células e abas.
*   **Google Drive API:** Gestão de sistema de arquivos e permissões.
*   **HTML5 / CSS3:** Construção do Painel Lateral e Modais de alerta.

---

## ⚠️ Requisitos

*   Conta Google Workspace ou Gmail.
*   Acesso às planilhas de Origem e Destino configuradas no `Config.gs`.
*   Planilha Template de Importação Omie hospedada no Google Drive.

---

**Autor:** [Seu Nome]

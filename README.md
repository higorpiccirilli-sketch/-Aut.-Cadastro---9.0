# Automação de Gestão de Produtos e Exportação Omie (Google Apps Script)

Este projeto é um sistema de automação desenvolvido em **Google Apps Script** para gerenciar o cadastro de produtos em uma planilha Google Sheets e automatizar a geração de arquivos de importação (`.xlsx`) compatíveis com o ERP **Omie**.

## 🚀 Funcionalidades Principais

### 1. Sincronização de Dados
- **Importação Automática:** Conecta-se a uma planilha de origem externa.
- **Detecção de Novos Produtos:** Identifica produtos novos listados na origem e os adiciona à base de dados local.
- **Atualização Inteligente:** Atualiza apenas campos específicos (SKU, NCM, GTIN) que estejam marcados como "FALTANDO", preservando dados já preenchidos manualmente.

### 2. Exportação para Omie (.xlsx)
- **Seleção Manual:** O usuário seleciona quais produtos deseja exportar através de caixas de seleção (Checkboxes) na planilha.
- **Validação de Dados:** Verifica integridade de SKU, NCM e GTIN antes da exportação.
- **Check de Duplicidade:** Impede a exportação de SKUs ou EANs duplicados.
- **Geração de Arquivo:** Utiliza um *Template* auxiliar para gerar arquivos Excel limpos e formatados.
- **Organização no Drive:** Salva os arquivos automaticamente em pastas organizadas por **Empresa > Ano > Mês**.

### 3. Integração com Metabase
- **Conexão API:** Conecta-se à API do Metabase para extrair relatórios atualizados (ex: Dados Box, Quantidade Box).
- **Gestão de Sessão:** Implementa cache de token de autenticação para otimizar chamadas à API.

### 4. Interface de Usuário (Frontend)
- **Painel Lateral:** Sidebar HTML para controle rápido das funções.
- **Logs em Tempo Real:** Modal para acompanhamento visual do progresso das execuções.
- **Gerenciador de Arquivos:** Interface para listar e baixar os últimos arquivos gerados diretamente da planilha.

---

## 🛠️ Arquitetura do Projeto

O código está modularizado para facilitar a manutenção e seguir boas práticas (Separation of Concerns):

*   `Config.gs`: Centraliza IDs de planilhas, pastas do Drive, URLs e mapeamento de colunas. Nenhuma configuração "hardcoded" fica nos scripts lógicos.
*   `Sincronizacao.gs`: Lógica para ler a planilha de origem e atualizar a base local.
*   `Exportacao.gs`: Lógica de validação, preparação de dados e geração do arquivo `.xlsx` via Template.
*   `Metabase.gs`: Cliente HTTP para autenticação e consulta à API do Metabase.
*   `InterfaceBackend.gs`: Controladores que ligam o HTML ao Google Apps Script.
*   `Utilitarios.gs`: Funções auxiliares (logs, formatação de data, busca de última linha).

---

## ⚙️ Configuração

Para rodar este projeto, é necessário configurar as **Script Properties** (Propriedades do Script) no editor do Google Apps Script com as seguintes chaves (para segurança):

*   `MB_URL`: URL base do Metabase.
*   `MB_USER`: Usuário do Metabase.
*   `MB_PASS`: Senha do Metabase.
*   `ALERT_EMAIL`: E-mail para receber alertas de erro.

Além disso, o arquivo `Config.gs` deve ser ajustado com os IDs das suas planilhas e pastas do Google Drive.

## 💻 Tecnologias Utilizadas

*   **Google Apps Script (GAS):** Backend Serverless baseado em JavaScript (V8 Runtime).
*   **Google Sheets API:** Manipulação de células e abas.
*   **Google Drive API:** Criação e organização de pastas/arquivos.
*   **UrlFetchApp:** Requisições HTTP externas (API Metabase e Download de Blob).
*   **HTML Service:** Criação de interfaces gráficas dentro do Sheets.

---

## 📝 Como Usar

1.  Abra a planilha de gestão.
2.  Acesse o menu customizado **"▶️ Painel de Controle"**.
3.  **Para Sincronizar:** Clique em "Sincronizar Manualmente" para puxar novos produtos.
4.  **Para Exportar:**
    *   Marque a caixa de seleção (Coluna F) dos produtos desejados.
    *   No painel, clique em "Gerar Arquivos Manuais".
    *   Aguarde o processamento e o link de download aparecerá no Log.

---

**Autor:** [Seu Nome]

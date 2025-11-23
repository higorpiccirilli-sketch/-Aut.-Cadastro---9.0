# 🚀 Automação de Gestão de Produtos e Integração Omie ERP

Este projeto é uma solução **Enterprise-grade** desenvolvida em **Google Apps Script** para orquestrar o ciclo de vida do cadastro de produtos. Ele atua como um middleware entre planilhas de gestão, dados externos e o ERP **Omie**, automatizando sincronização, validação, transformação de dados e gestão de arquivos no Google Drive.

O sistema elimina erros manuais, garante integridade relacional (Produtos vs. Características) e oferece uma interface gráfica robusta diretamente no Google Sheets.

---

## 🎯 Objetivos e Soluções

1.  **Eliminação de Erro Humano:** Validações rígidas de SKU, NCM e GTIN antes da geração de arquivos.
2.  **Padronização para ERP:** Gera arquivos `.xlsx` estritamente formatados para importação no Omie.
3.  **Transformação de Dados (ETL):** Converte linhas únicas de produtos em múltiplas linhas de atributos (Características) automaticamente.
4.  **Gestão de Documentos:** Organiza arquivos no Drive e permite exclusão (Lixeira) diretamente pela interface da planilha.

---

## 🛠️ Funcionalidades Principais

### 1. Sincronização Inteligente (Data Sync)
- **Atualização Incremental:** Conecta-se a uma base de dados externa (Planilha de Origem) e identifica novos SKUs.
- **Enriquecimento de Dados:** Preenche lacunas de dados locais (ex: NCM ou GTIN marcados como "FALTANDO") sem sobrescrever edições manuais existentes.
- **Otimização:** Utiliza leitura em lote (Batch Processing) para comparar milhares de linhas em milissegundos.

### 2. Exportação de Produtos (Arquivo Mestre)
Gera planilhas de cadastro de produtos para o Omie.
- **Lógica Multi-Marca:**
  - **Petiko:** Arquivo mestre contendo *todos* os itens selecionados.
  - **Innova / Paws:** Arquivos segmentados gerados automaticamente apenas se houver produtos dessas marcas no lote.
- **Template Engine:** Utiliza uma planilha "Molde" oculta, realiza limpeza higiênica de dados antigos (`Deep Clean`) e injeta os novos dados preservando fórmulas complexas.

### 3. Exportação de Características (Lógica 1:N)
Funcionalidade avançada para cadastro de atributos no ERP.
- **Explosão de Dados:** Transforma 1 linha de Produto em N linhas de Características (Tamanho, Linha Comercial, Classificação, etc.).
- **Extração via Regex:** Identifica automaticamente o tamanho do produto (ex: "P", "M", "G") a partir da descrição, ignorando padrões inválidos.
- **Interface Modal:** Abre um formulário HTML flutuante para que o usuário defina o "Tema" dos produtos em lote antes da geração.
- **Higiene Cruzada:** Garante que, ao gerar características, as abas de produtos do template sejam limpas (e vice-versa), evitando contaminação de dados na importação.

### 4. Gestão de Arquivos e Logs
- **Log Inteligente:** Registra cada exportação com Timestamp, Link direto, SKU e NCM. O script ignora checkboxes vazios para calcular a posição correta de inserção.
- **Lixeira Integrada:** Permite ao usuário excluir arquivos do Google Drive marcando uma caixa de seleção na planilha. O sistema move o arquivo para a Lixeira e risca visualmente a linha no log.
- **Estrutura de Pastas:**
  - `Empresa > Ano > Mês > Arquivos de Produto`
  - `Empresa > Ano > Mês > Caracteristica > Arquivos de Característica`

### 5. Segurança e Integração BI
- **Cofre de Senhas:** Credenciais do Metabase salvas em `Script Properties` (não expostas no código).
- **Conexão API:** Atualiza relatórios de BI automaticamente via requisições HTTP autenticadas.

---

## 🧩 Arquitetura do Projeto

O código é modular e segue princípios de **Clean Code**, facilitando manutenção e escalabilidade.

| Módulo | Responsabilidade |
| :--- | :--- |
| `Config.gs` | Centraliza IDs, URLs e mapeamento de colunas. Único ponto de alteração para manutenção básica. |
| `Exportacao.gs` | Core do sistema. Gerencia as regras de negócio de Produtos e Características, manipulação de Templates e API do Drive. |
| `Sincronizacao.gs` | Motor de comparação de dados. Executa lógica de "Merge" inteligente entre origem e destino. |
| `InterfaceBackend.gs` | Controlador (Controller). Gerencia o Menu, Sidebar, Modais e comunicação Client-Server. |
| `Utilitarios.gs` | Helpers globais. Inclui algoritmos otimizados de busca de última linha (`getLastRow` inteligente). |
| `Metabase.gs` | Cliente API seguro para atualização de dados de Business Intelligence. |
| `Frontend (HTML)` | `PainelDeControle`, `ModalCaracteristicas`, `Log` - Interfaces de usuário responsivas. |

---

## ⚙️ Fluxo Técnico de Geração (.xlsx)

Para garantir a integridade dos arquivos Omie (que possuem fórmulas ocultas e validações), o script executa o seguinte pipeline:

1.  **Staging Local:** Os dados brutos são colados em abas auxiliares (`Omie_Produtos`) na planilha ativa.
2.  **Cálculo Server-side:** O Google Sheets recalcula fórmulas nessas abas (tratamento de strings, concatenações).
3.  **Abertura do Template:** O script acessa a planilha Template oculta via ID.
4.  **Limpeza Cruzada:**
    *   Se gerando Produtos: Apaga a aba de Características do Template.
    *   Se gerando Características: Apaga a aba de Produtos do Template.
5.  **Injeção de Dados:** Copia os valores calculados do Staging Local para o Template limpo.
6.  **Download & Save:** Baixa o blob binário e salva na pasta correta do Drive com nomenclatura padronizada.

---

## 💻 Stack Tecnológico

*   **Google Apps Script (V8 Engine):** Lógica de backend serverless.
*   **Google Drive API:** Manipulação de arquivos, pastas e lixeira.
*   **Google Sheets API:** Leitura/Escrita de células e formatação condicional.
*   **HTML5 / CSS3:** Interfaces de usuário (Sidebar e Modais).
*   **JSON:** Troca de dados entre Frontend (Modal) e Backend (Script).

---

## ⚠️ Configuração

Este projeto requer a configuração de **Propriedades do Script** para segurança:
*   `MB_URL`, `MB_USER`, `MB_PASS`: Credenciais do Metabase.

IDs de pastas e planilhas devem ser configurados no objeto `CONFIG` em `Config.gs`.

---

**Autor:** [Seu Nome]

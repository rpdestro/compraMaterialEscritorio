# 📦 Controle de Materiais e Suprimentos

Uma aplicação web Front-End para controle de pedidos de materiais, projetada com foco em performance, usabilidade e nos princípios de **Clean Code**.

## 🚀 Principais Funcionalidades

* **Gestão Integrada de Arquivos:** Importação, mesclagem e exportação direta de planilhas Excel (`.xlsx`, `.xls`, `.csv`) com preservação de estilos e formatação nativa.
* **Dashboard Interativo:** Acompanhamento dinâmico e em tempo real dos status de suprimentos (Total, Pendentes, Solicitados, Empenhados e Entregues).
* **Data Grid Avançado:** Tabela de alta performance com paginação, ordenação, reordenação de colunas via *drag-and-drop* e filtros compostos (busca global e por pedidos de datas).
* **Operações CRUD \& em Massa:** Inserção, clonagem, edição e exclusão individual ou em lote (Bulk Actions).
* **Resiliência de Dados:** Sistema de *Auto-save* automático no `localStorage` para prevenção contra perda de dados em caso de fechamento acidental da aba.
* **Experiência do Usuário (UX):** Suporte nativo à alternância entre Tema Claro e Tema Escuro (*Dark Mode*).

## 🛠️ Arquitetura e Tecnologias

O projeto adota uma arquitetura modularizada, baseada em **JavaScript Vanilla (ES6+)**, garantindo alta escalabilidade e fácil manutenção.

* **Linguagens:** HTML5 e CSS3 (variáveis CSS, layouts fluidos sem dependência de frameworks).
* **Engine JavaScript:**

  * `Store`: Gerenciamento de estado global (*State Management*).
  * `FileIO`: Manipulação de arquivos (Parsing/Export).
  * `UI / Modals`: Renderização reativa do DOM e interface.
  * `Logic / Utils`: Tratamento de regras de negócios e algoritmos de filtragem/ordenação.
* **Bibliotecas de Terceiros:** [xlsx-js-style](https://www.npmjs.com/package/xlsx-js-style) via CDN para processamento e estilização inteligente de dados Excel.
* **Pronto para QA:** Componentes e ações mapeados com atributos `data-testid`, facilitando automação de testes End-to-End (E2E) com **Playwright**.

## 📂 Estrutura de Arquivos

* `index.html`: Marcação semântica principal e modais do sistema.
* `style.css`: Folha de estilos centralizada e otimizada.
* `script.js`: Motor de regras de negócio, renderização e eventos da aplicação.

## ⚙️ Como Executar Localmente

Sendo uma aplicação focada no *Client-Side* (Front-End puro), a execução é imediata e não exige configuração de servidores ou processos de *build*.

1. Clone o repositório:

```bash
   git clone https://github.com/seu-usuario/controle-materiais.git
   ```

2. Navegue até o diretório do projeto:

```bash
   cd controle-materiais
   ```

3. Abra o arquivo `index.html` diretamente em um navegador moderno, ou utilize extensões como **Live Server** no VS Code para auto-reload durante o desenvolvimento.

\---

Desenvolvido por **Renato Pinheiro Destro**


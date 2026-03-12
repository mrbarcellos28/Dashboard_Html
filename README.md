# 📊 Dashboard Operacional Multi-Bases (Google Apps Script)

Este projeto é um **Web App construído em Google Apps Script (GAS)** que atua como um dashboard gerencial completo. Ele extrai, processa e consolida dados de múltiplas abas de uma planilha do Google Sheets, exibindo indicadores (KPIs), gráficos e tabelas em uma interface web responsiva e em *dark mode*.

## 🚀 Funcionalidades

O painel é dividido em visões estratégicas para diferentes áreas da empresa:

* **Visão Geral:** Resumo de faturamento, pedidos represados e status de clientes.
* **Vendas:** Progresso de metas, análise de margem comercial, pedidos cadastrados no OMIE e painel de ativação de clientes (Churn/Follow-up).
* **Logística:** Controle de entregas, OTIF, saving de fretes, custo logístico e gestão de pallets pendentes.
* **Produção & Suprimentos:** Volume produzido vs meta, qualidade (devoluções), giro de estoque e MRP (Planejado vs Realizado).
* **Diretoria (Acesso Restrito):** Visão executiva com cascata de faturamento, detratores/clientes ouro, projeção de entrada de caixa e ranking de vendedores.
* **Modo Carrossel:** Rotação automática das telas, ideal para exibição em TVs corporativas.
* **Auto-Refresh:** Atualização assíncrona dos dados a cada 5 minutos em *background*.

## 📂 Estrutura de Arquivos

* `DashboardPage.html`: Contém todo o Front-end (HTML, CSS e JavaScript). Renderiza a interface, os componentes visuais e os gráficos (via Chart.js ou canvas nativo).
* `Code.gs`: Contém o Back-end (Google Apps Script). Executa a função `doGet` para servir a página web e possui as lógicas de leitura, filtro e cálculos de cada aba da planilha (`Base-faturamento`, `Base-pedidos`, `Base-OMIE`, etc.).

## ⚙️ Instalação e Configuração

Para rodar este dashboard na sua própria conta do Google / Google Workspace:

1. Abra a sua planilha do Google Sheets que contém as bases de dados.
2. No menu superior, clique em **Extensões > Apps Script**.
3. Crie um arquivo de script (`.gs`) e cole o conteúdo de `Code.gs`.
4. Crie um arquivo HTML chamado **exatamente** `DashboardPage.html` e cole o conteúdo correspondente.
5. Salve o projeto.

### Para uso interno na planilha:
1. Feche o editor de script e atualize a planilha (F5).
2. Um novo menu chamado **"📊 Abrir Dashboard"** aparecerá na barra superior.
3. Clique nele para abrir o painel em uma janela modal sobre a planilha.

### Para uso como Web App (Link externo ou TV):
1. No editor do Apps Script, clique no botão azul **Implantar > Nova implantação** (Deploy > New deployment).
2. Selecione o tipo **App da Web** (Web App).
3. Preencha a descrição, escolha executar como "Você" e libere o acesso para "Qualquer pessoa" (ou apenas para sua organização).
4. Clique em **Implantar**. 
5. Copie a "URL do App da Web" gerada. Esse é o link que você pode abrir em qualquer navegador ou TV.

## 🛠️ Tecnologias Utilizadas
* Google Apps Script (ES5/ES6)
* HTML5 / CSS3 (Variáveis CSS, Flexbox, CSS Grid)
* JavaScript Vanilla (Manipulação de DOM, Canvas API)

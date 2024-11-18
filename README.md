# Desafio-BCP-SECURITIES

## Visão Geral
Este projeto foi desenvolvido como parte do desafio técnico para o processo seletivo de estágio em Análise de Dados na BCP Securities.

## Coleta e Análise de Dados de Debêntures

Este repositório contém um script em Python que automatiza a **coleta**, o **processamento** e a **visualização** de dados de debêntures, com o objetivo de facilitar a análise de indicadores financeiros.

---

### Funcionalidades

#### 1. **Coleta Automática de Dados**
- O script faz o download dos dados diários de preços de debêntures disponíveis no site da ANBIMA para os últimos cinco dias úteis. Caso algum dia específico esteja indisponível devido a problemas no site, o código prossegue com os demais dias úteis disponíveis e registra as mensagens de erro em um arquivo log.txt, garantindo o processamento dos dados sem interrupções.
- Os arquivos são automaticamente salvos na pasta `Daily Prices`, criada no mesmo diretório do script.

#### 2. **Processamento e Consolidação**
- Cada aba dos arquivos Excel baixados é processada para selecionar informações relevantes, como:
  - **Código**
  - **PU**
  - **Taxa de Compra**
  - **Taxa de Venda**
  - **Taxa Indicativa**
  - **Indexador**
  - **Date**

Os dados são consolidados em dois arquivos Excel, ambos podem ser importados para ferramentas de análise de dados:

-consolidated_data.xlsx: Neste arquivo, os dados estão organizados em abas separadas, cada uma contendo informações específicas sobre a data e o indicador correspondente.

-consolidated_data_single_sheet.xlsx: Neste arquivo, todos os dados são consolidados em uma única tabela, facilitando a visualização e a análise integrada.

Ambos os arquivos podem ser importados para o Power BI para análises adicionais.

#### 3. **Visualização de Dados**
- O script gera gráficos que mostram a **Taxa Indicativa Média por Data** para diferentes indexadores (ex.: **IPCA +**, **% do DI**, **DI +**).
- Os gráficos são salvos no formato PNG e podem ser utilizados em relatórios ou apresentações.

---

### Como Usar

#### **Pré-requisitos**
Certifique-se de ter o Python instalado, juntamente com as seguintes bibliotecas:

- `pandas`
- `matplotlib`
- `requests`
- `openpyxl`
- `xlrd`

Para instalar as dependências, execute:

```bash
pip install pandas matplotlib requests openpyxl xlrd


## Dashboard Interativo (*Extra*)

Este projeto inclui um **dashboard interativo**, desenvolvido com **Streamlit**, como uma funcionalidade adicional para visualização e análise dos dados de taxas indicativas. Embora não tenha sido solicitado como parte do escopo original do desafio, ele foi criado para enriquecer a experiência do usuário. Além disso, este script depende do código principal teste_desafioBCP_ThaisAngelo.py para ser executado, pois utiliza como base o datasheet consolidated_data.xlsx gerado por ele. Certifique-se de rodar o código principal antes de utilizar este script.

### Funcionalidades do Dashboard

- Carregamento de dados consolidados de um arquivo Excel.
- Filtro interativo de datas para explorar os dados.
- Exibição dos dados filtrados diretamente no dashboard.
- Geração de gráficos que mostram as taxas indicativas médias por data e indexador.

### Como Executar o Dashboard

1. Certifique-se de que as dependências estão instaladas:
    ```bash
    pip install pandas matplotlib streamlit
    ```

2. Coloque o arquivo `consolidated_data.xlsx` no mesmo diretório do script.

3. Execute o dashboard com o comando:
    ```bash
    streamlit run nome_do_script.py
    ```

4. Acesse o dashboard no navegador pelo link exibido no terminal, geralmente em `http://localhost:8501`.

### Código do Dashboard

O código do dashboard está localizado no arquivo `dashboard_desafioBCP_ThaisAngelo.py.py`. Ele inclui:

- Funções para carregar e filtrar dados do Excel.
- Geração de gráficos com o Matplotlib.
- Interface interativa criada com Streamlit.

### Observação

> **Nota:** Este dashboard não faz parte do escopo obrigatório do projeto e foi incluído como uma melhoria para fins de apresentação e exploração de dados.

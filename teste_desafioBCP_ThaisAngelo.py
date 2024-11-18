import os
import pandas as pd
import matplotlib.pyplot as plt
import requests
from datetime import datetime, timedelta

# Obtém o diretório atual onde o script está localizado
current_directory = os.path.dirname(os.path.abspath(__file__))
# Define o caminho da pasta onde os arquivos serão salvos, dentro do diretório atual
folder_path = os.path.join(current_directory, 'Daily Prices')

# Verifica se a pasta 'Daily Prices' existe, e cria-a se não existir
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# URL base para download dos arquivos
base_url = 'https://www.anbima.com.br/informacoes/merc-sec-debentures/arqs/'

############################################
# Parte do código que baixa os arquivos e os salva na pasta 'Daily Prices'

# Função para obter os últimos 5 dias úteis, excluindo o dia atual
def get_last_weekdays(num_days=5):
    weekdays = []  # Lista para armazenar os dias úteis
    today = datetime.today() - timedelta(days=1)  # Começar a contar do dia anterior ao atual

    # Itera até obter o número necessário de dias úteis
    while len(weekdays) < num_days:
        if today.weekday() < 5:  # Segunda a Sexta são considerados dias úteis
            weekdays.append(today)
        today -= timedelta(days=1)  # Voltar um dia no tempo
    return weekdays

# Função para gerar os links de download dos arquivos
def generate_download_links(weekdays):
    download_links = []  # Lista para armazenar os links de download
    for day in weekdays:
        day_str = day.strftime('%y%b%d').lower()  # Formata o dia, mês e ano no formato correto

        # Cria o link de download para cada dia
        link = f"{base_url}d{day_str}.xls"
        download_links.append(link)
    return download_links

# Função para gerar o nome do arquivo com base na data
def format_file_name(day):
    file_name = f"{day.strftime('%Y%m%d').lower()}.xls"
    return file_name

# Função para baixar os arquivos e salvá-los na pasta 'Daily Prices'
def download_files():
    weekdays = get_last_weekdays()  # Obtém os últimos 5 dias úteis
    download_links = generate_download_links(weekdays)  # Gera os links de download
    session = requests.Session()  # Cria uma sessão para os downloads

    for day, link in zip(weekdays, download_links):
        file_name = format_file_name(day)  # Gera o nome do arquivo com base na data
        file_path = os.path.join(folder_path, file_name)  # Define o caminho completo do arquivo

        # Tentar baixar o arquivo
        file_response = session.get(link)

        # Verifica se o download foi bem-sucedido
        if file_response.status_code == 200:
            with open(file_path, 'wb') as file:
                file.write(file_response.content)
            print(f"Arquivo salvo: {file_path}")
        else:  # Caso ocorra um erro, imprime a mensagem e registra no log
            error_message = f"Erro ao acessar o link: {link} - Código de Status: {file_response.status_code}"
            print(error_message)
            print('Erro salvo no arquivo log.txt')
            # Escreve a mensagem no arquivo log.txt
            with open('log.txt', 'a') as log_file:  # 'a' para adicionar sem sobrescrever
                log_file.write(error_message + '\n')

############################################

# Função para determinar o indexador com base no nome da aba (sheet)
def determine_indexador(sheet_name):
    if 'ipca_spread' in sheet_name.lower():
        return 'IPCA +'
    elif 'di_percentual' in sheet_name.lower():
        return '% do DI'
    elif 'di_spread' in sheet_name.lower():
        return 'DI +'
    elif 'vencidos_antecipadamente' in sheet_name.lower():
        return 'Vencidos Antecipadamente'
    else:
        return 'Desconhecido'

# Função para processar cada aba de um arquivo Excel
def process_sheet(file_path, sheet_name):
    # Carregar o arquivo Excel e a aba específica
    df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=7, engine='xlrd')  # Pular as primeiras 7 linhas irrelevantes

    # Verificar se há linhas em branco na coluna "Código"
    if not df['Código'].isna().empty:
        first_blank_index = df[df['Código'].isna()].index[1]  # Encontrar a segunda linha em branco

        # Remover todas as linhas após a primeira linha em branco
        df = df.iloc[:first_blank_index]

    # Verificar se as colunas necessárias existem no DataFrame
    if all(col in df.columns for col in ['Código', 'Nome', 'PU', 'Taxa de Compra', 'Taxa de Venda', 'Taxa Indicativa']):
        df = df.iloc[1:] # remove a primeira linha em branco
        # Adicionar uma coluna com a data (baseada no nome do arquivo)
        df['Date'] = os.path.basename(file_path).split('.')[0]  # Extrai 'aaaammdd' do nome do arquivo
        df.loc[0, 'Date'] = None  # Define a primeira linha como NaN (vazia)

        # Selecionar somente as colunas de interesse
        df_filtered = df[['Código', 'Nome', 'PU', 'Taxa de Compra', 'Taxa de Venda', 'Taxa Indicativa', 'Date']].copy()

        # Adicionar a coluna de indexador usando .loc para evitar o warning
        df_filtered.loc[:, 'Indexador'] = determine_indexador(sheet_name)
        
        # Remover as linhas onde a célula na coluna 'Código' está vazia
        df_filtered = df_filtered.dropna(subset=['Código'])

        return df_filtered
    else:
        print(f"Aba {sheet_name} no arquivo {os.path.basename(file_path)} não contém as colunas necessárias.")
        return None

# Função para processar o arquivo Excel completo
def process_file(file_path):
    # Carregar o arquivo Excel com todas as abas (sheets)
    xls = pd.ExcelFile(file_path, engine='xlrd')

    # Dicionário para armazenar DataFrames tratados por aba
    treated_sheets = {}

    # Iterar pelas abas do arquivo Excel
    for sheet_name in xls.sheet_names:
        df = process_sheet(file_path, sheet_name)
        if df is not None:
            sheet_key = f"{os.path.basename(file_path).split('.')[0]}_{sheet_name}"  # Nome para a aba combinada
            treated_sheets[sheet_key] = df

    return treated_sheets

# Função para salvar todas as abas tratadas em um único arquivo Excel
def save_all_sheets(treated_sheets):
    file_path = 'consolidated_data.xlsx'  # Datasheet completo, com todas as informações de todas as datas
    file_path_concatenado = 'consolidated_data_single_sheet.xlsx'
    
    consolidated_df = pd.concat(treated_sheets.values(), ignore_index=True)
    # Salvar todas as abas tratadas em um único arquivo Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for sheet_key, df in treated_sheets.items():
            df.to_excel(writer, sheet_name=sheet_key[:31], index=False)  # Limitar o nome da aba a 31 caracteres

    with pd.ExcelWriter(file_path_concatenado, engine='openpyxl') as writer:
        consolidated_df.to_excel(writer, sheet_name='Consolidated', index=False)

    return file_path, file_path_concatenado

# def save_all_sheets(treated_sheets):
#     file_path = 'consolidated_data_single_sheet.xlsx'
    
#     consolidated_df = pd.concat(treated_sheets.values(), ignore_index=True)
    
#     with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
#         consolidated_df.to_excel(writer, sheet_name='Consolidated', index=False)
    
#     return file_path, saved_file_path_concatenado

# Função para plotar os gráficos de taxa indicativa média por data
def plot_indicative_rate_by_date(dataframe):
    # Converter a coluna 'Taxa Indicativa' para float
    dataframe['Taxa Indicativa'] = pd.to_numeric(dataframe['Taxa Indicativa'], errors='coerce')

    # Converter a coluna 'Date' para o formato datetime
    dataframe['Date'] = pd.to_datetime(dataframe['Date'], format='%Y%m%d')

    # Agrupar os dados por 'Date' e 'Indexador' e calcular a média da 'Taxa Indicativa'
    grouped_data = dataframe.groupby(['Date', 'Indexador'])['Taxa Indicativa'].mean().reset_index()

    # Lista de indexadores únicos
    indexadores = grouped_data['Indexador'].unique()

    for indexador in indexadores:
        if indexador == 'Vencidos Antecipadamente':
            continue  # Ignorar este indexador

        # Filtrar os dados para o indexador atual
        data_indexador = grouped_data[grouped_data['Indexador'] == indexador]

        # Plotar o gráfico
        plt.figure(figsize=(10, 6))
        plt.plot(data_indexador['Date'], data_indexador['Taxa Indicativa'], marker='o', linestyle='-', label=indexador)
        plt.xlabel('Data')
        plt.ylabel('Taxa Indicativa Média')
        plt.title(f'Taxa Indicativa Média por Data ({indexador})')
        plt.legend()
        plt.grid(True)

        # Formatar datas no eixo x para o formato dd-mm-yyyy
        date_format = data_indexador['Date'].dt.strftime('%d-%m-%Y')
        plt.xticks(data_indexador['Date'], date_format, rotation=45)

        # Salvar o gráfico com o nome do arquivo no formato aaaammdd
        file_name = f'indicative_rate_{indexador.replace(" ", "_")}.png'
        plt.savefig(file_name)
        plt.show()

    # Retornar a lista de arquivos salvos
    return [f'indicative_rate_{indexador.replace(" ", "_")}.png' for indexador in indexadores if indexador != 'Vencidos Antecipadamente']



def main():
    mensagens = download_files()
    # Dicionário para armazenar todos os DataFrames tratados de todos os arquivos
    all_treated_sheets = {}
    
    # Iterar pelos arquivos na pasta
    for file in os.listdir(folder_path):
        if file.endswith('.xls'):  # Verifica arquivos Excel
            file_path = os.path.join(folder_path, file)
            treated_sheets = process_file(file_path)
            if treated_sheets:
                all_treated_sheets.update(treated_sheets)
    
    if all_treated_sheets:
        saved_file_path, saved_file_path_concatenado = save_all_sheets(all_treated_sheets)  #arquivo salvo com todas as abas/sheets dos 5 arquivos em um só, que possa ser importado para power bi
        print(f"Dataset consolidado salvo como '{saved_file_path}'")
        
        # Carregar o dataset consolidado
        consolidated_data = pd.read_excel(saved_file_path, sheet_name=None)
        
        # Concatenar todas as abas em um único DataFrame
        all_data_frames = []
        for sheet_name, df in consolidated_data.items():
            all_data_frames.append(df)
        consolidated_df = pd.concat(all_data_frames, ignore_index=True)
        
        # Plotar os gráficos
        plot_files = plot_indicative_rate_by_date(consolidated_df)
        for plot_file in plot_files:
            print(f"Gráfico salvo como '{plot_file}'")
    else:
        print("Nenhum arquivo contém as colunas necessárias ou os dados não foram processados corretamente.")

if __name__ == "__main__":
    main()



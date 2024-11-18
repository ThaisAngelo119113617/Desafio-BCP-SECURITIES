import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# Função para carregar os dados consolidados do arquivo Excel
def load_data(file_path):
    # Carregar todas as abas do arquivo Excel em um dicionário de DataFrames
    consolidated_data = pd.read_excel(file_path, sheet_name=None)
    # Concatenar todas as abas em um único DataFrame
    all_data_frames = [df for df in consolidated_data.values()]
    return pd.concat(all_data_frames, ignore_index=True)

# Função para plotar os gráficos de taxa indicativa média por data
def plot_indicative_rate_by_date(dataframe, selected_dates):
    # Converter a coluna 'Taxa Indicativa' para tipo float
    dataframe['Taxa Indicativa'] = pd.to_numeric(dataframe['Taxa Indicativa'], errors='coerce')
    # Filtrar os dados pelas datas selecionadas
    dataframe = dataframe[dataframe['Date'].isin(selected_dates)]
    
    # Converter a coluna 'Date' para o formato datetime
    dataframe['Date'] = pd.to_datetime(dataframe['Date'], format='%Y%m%d')
    
    # Agrupar os dados por 'Date' e 'Indexador' e calcular a média da 'Taxa Indicativa'
    grouped_data = dataframe.groupby(['Date', 'Indexador'])['Taxa Indicativa'].mean().reset_index()
    
    # Lista de indexadores únicos
    indexadores = grouped_data['Indexador'].unique()
    
    plots = []
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
        plots.append(file_name)
        plt.show()
    
    return plots

# Função principal do Streamlit para criar o dashboard
def main():
    st.title('Debêntures Dashboard')
    st.write('Este dashboard permite visualizar as taxas indicativas médias por data.')
    
    # Carregar os dados do arquivo Excel consolidado
    file_path = 'consolidated_data.xlsx'
    df = load_data(file_path)
    
    # Selecionar as datas únicas no DataFrame e remover NaN
    unique_dates = df['Date'].dropna().unique()
    
    # Seleção de datas no Streamlit para filtragem
    selected_dates = st.multiselect('Selecione as datas:', unique_dates, default=list(unique_dates))
    
    # Exibir os dados filtrados
    st.write('Dados filtrados:')
    filtered_data = df[df['Date'].isin(selected_dates)]
    st.dataframe(filtered_data)
    
    # Plotar os gráficos ao clicar no botão
    if st.button('Gerar Gráficos'):
        plot_files = plot_indicative_rate_by_date(filtered_data, selected_dates)
        for plot_file in plot_files:
            st.image(plot_file)

# Ponto de entrada do script
if __name__ == "__main__":
    main()

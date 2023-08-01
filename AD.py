import pandas as pd


def ler_csv_maquina(csv_maquina):
    lista_csv = csv_maquina.split('?')
    lista_csv.pop()
    dados_completos = pd.DataFrame(columns=['Name', 'Creation Date', 'Disabled', 'DNS Host Name',
                                            'Last logon date', 'Operating System', 'Parent Container'])
    for data in lista_csv:
        dados = pd.read_csv(data, encoding='ISO-8859-1', names=['Name', 'Creation Date', 'Disabled', 'DNS Host Name',
                                                                'Last logon date', 'Operating System',
                                                                'Parent Container'])
        dados.drop(0, axis=0, inplace=True)
        dados = converter_dados(dados)

        dados_completos = pd.concat([dados_completos, dados], ignore_index=True)

    dados_completos.loc[dados_completos['Operating System'].str.contains('Server|SERVER|server'),
                        'Type of device'] = 'Server'
    dados_completos.loc[~dados_completos['Operating System'].str.contains('Server|SERVER|server'),
                        'Type of device'] = 'Client'
    dados_completos.loc[(dados_completos['Operating System'].isnull()) | (dados_completos['Operating System'] == ' ') |
                        (dados_completos['Operating System'] == ''), 'Type of device'] = 'unknown'
    dados_completos = dados_completos[['Name', 'Type of device', 'Creation Date', 'Disabled', 'DNS Host Name',
                                       'Last logon date', 'Operating System', 'Parent Container']]

    return dados_completos


def ler_csv_usuario(csv_usuario):
    lista_csv = csv_usuario.split('?')
    lista_csv.pop()
    dados_completos = pd.DataFrame(columns=['Name', 'Creation Date', 'Disabled',
                                            'Display Name', 'Email Address',
                                            'First Name', 'Last logon date', 'Last Name',
                                            'Parent Container'])
    for data in lista_csv:
        dados = pd.read_csv(data, encoding='ISO-8859-1', names=['Name', 'Creation Date', 'Disabled',
                                                                'Display Name', 'Email Address',
                                                                'First Name', 'Last logon date', 'Last Name',
                                                                'Parent Container'])
        dados.drop(0, axis=0, inplace=True)
        dados = converter_dados(dados)
        dados_completos = pd.concat([dados_completos, dados], ignore_index=True)

    return dados_completos


def converter_dados(dados):
    try:
        dados['Last logon date'].replace({' AM': '', ' PM': ''}, regex=True, inplace=True)
        dados['Creation Date'].replace({' AM': '', ' PM': ''}, regex=True, inplace=True)
        dados['Last logon date'] = pd.to_datetime(dados['Last logon date'])
        dados['Creation Date'] = pd.to_datetime(dados['Creation Date'])
        dados['Last logon date'] = pd.to_datetime(dados['Last logon date'],
                                                  format='%d/%m/%Y %H:%M:%S', errors='coerce')
        dados['Creation Date'] = pd.to_datetime(dados['Creation Date'],
                                                format='%d/%m/%Y %H:%M:%S', errors='coerce')
        return dados
    except:
        dados['Last logon date'] = pd.to_datetime(dados['Last logon date'],
                                                  format='%d/%m/%Y %H:%M:%S', errors='coerce')
        dados['Creation Date'] = pd.to_datetime(dados['Creation Date'],
                                                format='%d/%m/%Y %H:%M:%S', errors='coerce')
        return dados


def get_numero_linhas(guia, caminho):
    guia = pd.read_excel(caminho, sheet_name=guia)
    numero_linhas = len(guia.index)
    return int(numero_linhas)


def get_numero_colunas(guia, caminho):
    guia = pd.read_excel(caminho, sheet_name=guia)
    numero_colunas = len(guia.columns)
    return numero_colunas


def get_valor_celula(guia, caminho):
    guia = pd.read_excel(caminho, sheet_name=guia)
    valor = guia['Name']
    return valor

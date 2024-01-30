import pandas as pd

def extract_data(file_path):
    try:
        print('Initializind extract info on data base ...')
        header = pd.read_excel(file_path, nrows=1, header=None).iloc[0].tolist()
        df = pd.read_excel(file_path, header=None, names=header, skiprows=1)
        df['ID'] = df.index
        data_list = [['ID', 'Titulo', 'Conteudo', 'Classe']]
        
        for row_id, row in df.iterrows():
            title = str(row['Titulo']).replace('\n', '')
            content = str(row['Noticia']).replace('\n', '')
            classe = int(row['Classe']) if not pd.isna(row['Classe']) and row['Classe'] != '' else None
            data_list.append([row_id, title, content, classe])

        print('Finish extract info on data base\n')
        return data_list

    except FileNotFoundError:
        raise FileNotFoundError("Erro: O arquivo Excel não foi encontrado. Verifique o caminho e o nome do arquivo.")

    except pd.errors.EmptyDataError:
        raise pd.errors.EmptyDataError("Erro: O arquivo Excel está vazio ou não contém a planilha especificada.")

    except pd.errors.ParserError:
        raise pd.errors.ParserError("Erro: Há um problema na leitura do arquivo Excel.")

    except Exception as e:
        raise Exception(f"Ocorreu um erro não tratado: {e}")

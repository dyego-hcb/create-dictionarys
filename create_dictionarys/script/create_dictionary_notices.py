import csv
import sys
import os
import pandas as pd

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..'))
sys.path.append(project_root)

from utils.extract_info_notices.script.extract_info_notices import extract_data
from utils.words_tokenize.script.words_tokenize import tokenize_words
from utils.words_lowercase.script.words_lowercase import convert_words_lower_case
from utils.remove_accentuation.script.remove_accentuation import remover_accentuation
from utils.remove_punctuation.script.remove_punctuation import remover_ponctuation
from utils.remove_stopwords.script.remove_stopwords import remove_stopwords_in_list
from utils.words_stemmer.script.words_stemmer import stemmize_words
from unidecode import unidecode 

def add_notice_on_dict(dict_notice, id_dict_notice, id_notice, title_notice, content_notice, classe_notice):
    words_without_ponctuation = remover_ponctuation(content_notice)
    words_without_accentuation = remover_accentuation(words_without_ponctuation)
    words_tokenized = tokenize_words(words_without_accentuation)
    words_lower_case = convert_words_lower_case(words_tokenized)
    words_without_stopwords = remove_stopwords_in_list(words_lower_case)
    words_stemmed = stemmize_words(words_without_stopwords)

    dict_notice[id_dict_notice] = {}
    dict_notice[id_dict_notice]['id_notice'] = id_notice
    dict_notice[id_dict_notice]['title_notice'] = title_notice
    dict_notice[id_dict_notice]['notice_content'] = words_lower_case
    dict_notice[id_dict_notice]['notice_content_without_stopwords'] = words_without_stopwords
    dict_notice[id_dict_notice]['notice_content_stemm_without_stopwords'] = words_stemmed
    dict_notice[id_dict_notice]['classe_notice'] = classe_notice
    dict_notice[id_dict_notice]['notice_words_total_with_stopwords'] = len(words_lower_case)
    dict_notice[id_dict_notice]['notice_words_total_without_stopwords'] = len(words_without_stopwords)

    return dict_notice


def create_dictionary_notices(path_bd, classe_notice):
    data_list = extract_data(path_bd)

    print('Starting create dict of notices ...')
    dict_notice: dict = {}
    id_dict = 0
    words_total_with_stopwords_group = 0
    words_total_without_stopwords_group = 0

    for data in data_list[1:]:
        id_notice = data[0]
        title = data[1]
        content = data[2]
        classe = data[3]

        if classe == classe_notice:
            id_dict += 1
            dict_notice = add_notice_on_dict(dict_notice, id_dict, id_notice, title, content, classe)
            words_total_with_stopwords_group += dict_notice[id_dict]['notice_words_total_with_stopwords']
            words_total_without_stopwords_group += dict_notice[id_dict]['notice_words_total_without_stopwords']

    print('Finish create dict of notices\n')

    return dict_notice, words_total_with_stopwords_group, words_total_without_stopwords_group

def save_dict_notice_to_xlsx(file_path, dict_notice, group_name, words_with_stopwords, words_without_stopwords):
    print('Starting save dict of notices ...')

    df = pd.DataFrame.from_dict(dict_notice, orient='index')

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=f'Dicionario de Noticias {group_name}', index=False)

        worksheet = writer.sheets[f'Dicionario de Noticias {group_name}']

        (max_row, max_col) = df.shape

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value)

        for row_num, row_data in enumerate(df.itertuples(index=False)):
            for col_num, value in enumerate(row_data):
                if isinstance(value, list):
                    df.at[row_num, df.columns[col_num]] = '\n'.join(value)
                    value = '\n'.join(value)
                worksheet.write(row_num, col_num, value)

        last_col = len(df.columns)
        df.insert(last_col, 'info_dict', '')

        info_dict_list = [
            f"- Dicionario de Noticias {group_name}",
            f"  Esse dicionario possui {len(dict_notice)} noticias",
            f"  Esse dicionario possui {words_without_stopwords} palavras (STOP-WORDS REMOVIDAS)",
            f"  Esse dicionario possui {words_with_stopwords} palavras (STOP-WORDS NO TEXTO)",
            "  Estrutura do dicionario:",
            "  ID - ID NOTICE - TITLE NOTICE - NOTICE ALL WORDS - CLASSE NOTICE - NOTICE WORDS TOTAL WITH STOPWORDS -  NOTICE WORDS TOTAL WITHOUT STOPWORDS"
        ]

        for i in range(len(info_dict_list)):
            df.at[i, 'info_dict'] = info_dict_list[i]

        df.to_excel(writer, sheet_name=f'Dicionario de Noticias {group_name}', index=False)
    
    print('Finish save dict of notices\n')

def save_dict_notice_to_csv(file_path, dict_notice, group_name, words_with_stopwords, words_without_stopwords):
    print('Starting save dict of notices ...')

    pd.set_option('display.max_colwidth', None)

    df = pd.DataFrame.from_dict(dict_notice, orient='index')

    df['title_notice'] = df['title_notice'].apply(unidecode)

    df = df.apply(lambda col: col.apply(lambda x: '\n'.join(map(str, x)) if isinstance(x, list) else x))

    df['info_dict'] = ''

    info_dict_list = [
        f"- Dicionario de Noticias {group_name}",
        f"  Esse dicionario possui {len(dict_notice)} noticias",
        f"  Esse dicionario possui {words_without_stopwords} palavras (STOP-WORDS REMOVIDAS)",
        f"  Esse dicionario possui {words_with_stopwords} palavras (STOP-WORDS NO TEXTO)",
        "  Estrutura do dicionario:",
        "  ID - ID NOTICE - TITLE NOTICE - NOTICE ALL WORDS - CLASSE NOTICE - NOTICE WORDS TOTAL WITH STOPWORDS -  NOTICE WORDS TOTAL WITHOUT STOPWORDS"
    ]

    df.at[1, 'info_dict'] = info_dict_list[0]
    for i in range(1, len(info_dict_list)):
        df.at[i+1, 'info_dict'] = info_dict_list[i]

    df.to_csv(file_path, index=False, encoding='utf-8', quoting=csv.QUOTE_NONNUMERIC)

    print('Finish save dict of notices\n')
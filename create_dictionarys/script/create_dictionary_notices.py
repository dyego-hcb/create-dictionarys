import csv
import sys
import os
import openpyxl
import pandas as pd

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..'))
sys.path.append(project_root)

from unidecode import unidecode
from openpyxl import Workbook
from utils.words_stemmer.script.words_stemmer import stemmize_words
from utils.remove_stopwords.script.remove_stopwords import remove_stopwords_in_list
from utils.remove_punctuation.script.remove_punctuation import remover_ponctuation
from utils.remove_accentuation.script.remove_accentuation import remover_accentuation
from utils.words_lowercase.script.words_lowercase import convert_words_lower_case
from utils.words_tokenize.script.words_tokenize import tokenize_words
from utils.extract_info_notices.script.extract_info_notices import extract_data

def add_notice_on_dict(dict_notice, id_dict_notice, id_notice, title_notice, content_notice, classe_notice):
    words_without_ponctuation = remover_ponctuation(content_notice)
    words_without_accentuation = remover_accentuation(
        words_without_ponctuation)
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
    dict_notice[id_dict_notice]['notice_words_total_with_stopwords'] = len(
        words_lower_case)
    dict_notice[id_dict_notice]['notice_words_total_without_stopwords'] = len(
        words_without_stopwords)

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
            dict_notice = add_notice_on_dict(
                dict_notice, id_dict, id_notice, title, content, classe)
            words_total_with_stopwords_group += dict_notice[id_dict]['notice_words_total_with_stopwords']
            words_total_without_stopwords_group += dict_notice[id_dict]['notice_words_total_without_stopwords']

    print('Finish create dict of notices\n')

    return dict_notice, words_total_with_stopwords_group, words_total_without_stopwords_group


def load_dict_notices_xlsx(folder_path, file_name, dict_notice):
    print('Starting loading dict of notices ...')
    file_path = os.path.join(folder_path, file_name)

    df = pd.read_excel(file_path)

    for row_id, row in df.iterrows():
        id_dict_notice = row_id
        id_notice = row['id_notice'] if not pd.isna(row['id_notice']) else ''
        title_notice = row['title_notice']
        notice_content = row['notice_content']
        notice_content_without_stopwords = row['notice_content_without_stopwords']
        notice_content_stemm_without_stopwords = row['notice_content_stemm_without_stopwords']
        classe_notice = row['classe_notice']
        notice_words_total_with_stopwords = row['notice_words_total_with_stopwords'] if not pd.isna(row['notice_words_total_with_stopwords']) else ''
        notice_words_total_without_stopwords = row['notice_words_total_without_stopwords'] if not pd.isna(row['notice_words_total_without_stopwords']) else ''

        print("id_dict_notice: " + str(id_dict_notice))
        print("id_notice: " + str(id_notice))
        print("title_notice: " + str(title_notice))
        print("notice_content: " + str(notice_content))
        print("notice_content_without_stopwords: " + str(notice_content_without_stopwords))
        print("notice_content_stemm_without_stopwords: " + str(notice_content_stemm_without_stopwords))
        print("classe_notice: " + str(classe_notice))
        print("notice_words_total_with_stopwords: " + str(notice_words_total_with_stopwords))
        print("notice_words_total_without_stopwords: " + str(notice_words_total_without_stopwords))

        dict_notice[id_dict_notice] = {}
        dict_notice[id_dict_notice]['id_notice'] = id_notice
        dict_notice[id_dict_notice]['title_notice'] = title_notice
        dict_notice[id_dict_notice]['notice_content'] = notice_content
        dict_notice[id_dict_notice]['notice_content_without_stopwords'] = notice_content_without_stopwords
        dict_notice[id_dict_notice]['notice_content_stemm_without_stopwords'] = notice_content_stemm_without_stopwords
        dict_notice[id_dict_notice]['classe_notice'] = classe_notice
        dict_notice[id_dict_notice]['notice_words_total_with_stopwords'] = notice_words_total_with_stopwords
        dict_notice[id_dict_notice]['notice_words_total_without_stopwords'] = notice_words_total_without_stopwords

    print('Finish load dict of notices\n')

    return dict_notice


def save_dict_notices_to_xlsx(file_path, dict_notice, group_name, words_with_stopwords, words_without_stopwords):
    print('Starting save dict of notices ...')

    df = pd.DataFrame.from_dict(dict_notice, orient='index')

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(
            writer, sheet_name=f'Dicionario de Noticias {group_name}', index=False)

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

        df.to_excel(
            writer, sheet_name=f'Dicionario de Noticias {group_name}', index=False)

    print('Finish save dict of notices\n')


def save_dict_notices_to_csv(file_path, dict_notice, group_name, words_with_stopwords, words_without_stopwords):
    print('Starting save dict of notices ...')

    pd.set_option('display.max_colwidth', None)

    df = pd.DataFrame.from_dict(dict_notice, orient='index')

    df['title_notice'] = df['title_notice'].apply(unidecode)

    df = df.apply(lambda col: col.apply(lambda x: '\n'.join(
        map(str, x)) if isinstance(x, list) else x))

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

    df.to_csv(file_path, index=False, encoding='utf-8',
              quoting=csv.QUOTE_NONNUMERIC)

    print('Finish save dict of notices\n')


def create_dictionary_notices_relevant_info(dict_notice_relevant_info, dict_notice):
    print('Starting create dict of notices with relevant info ...')

    for id_dict_notice in dict_notice:
        notice_info = dict_notice[id_dict_notice]
        dict_notice_relevant_info[id_dict_notice] = {}
        dict_notice_relevant_info[id_dict_notice]['id_notice'] = notice_info.get('id_notice')
        dict_notice_relevant_info[id_dict_notice]['title_notice'] = notice_info.get('title_notice') 
        dict_notice_relevant_info[id_dict_notice]['real_words_strongs_in_notice'] = []
        dict_notice_relevant_info[id_dict_notice]['fake_words_strongs_in_notice'] = []
        dict_notice_relevant_info[id_dict_notice]['notice_strong_words_real_total'] = 0
        dict_notice_relevant_info[id_dict_notice]['notice_strong_words_fake_total'] = 0
        dict_notice_relevant_info[id_dict_notice]['notice_strong_words_total'] = 0
        dict_notice_relevant_info[id_dict_notice]['classe_notice'] = notice_info.get('classe_notice')
        dict_notice_relevant_info[id_dict_notice]['notice_words_total_with_stopwords'] = notice_info.get('notice_words_total_with_stopwords')
        dict_notice_relevant_info[id_dict_notice]['notice_words_total_without_stopwords'] = notice_info.get('notice_words_total_without_stopwords')

    print('Finish create dict of notices with relevant info ...\n')

    return dict_notice_relevant_info

def update_dictionary_notices_relevant_info(dict_notice_relevant_info, dict_notice, dict_strong_word, class_strong_word):
    print('Starting update dict of notices with relevant info ...')

    for id_dict_notice, notice_info in dict_notice.items():
        real_words_strongs_in_notice = list(dict_notice_relevant_info[id_dict_notice]['real_words_strongs_in_notice'])
        fake_words_strongs_in_notice = list(dict_notice_relevant_info[id_dict_notice]['fake_words_strongs_in_notice'])

        notice_words = notice_info['notice_content_stemm_without_stopwords']
        
        words_strongs_in_notice = [word for word in notice_words if word in {entry['word'] for entry in dict_strong_word.values()}]

        if class_strong_word == 1:
            real_words_strongs_in_notice.extend(words_strongs_in_notice)
        elif class_strong_word == 0:
            fake_words_strongs_in_notice.extend(words_strongs_in_notice)

        dict_notice_relevant_info[id_dict_notice]['real_words_strongs_in_notice'] = real_words_strongs_in_notice
        dict_notice_relevant_info[id_dict_notice]['fake_words_strongs_in_notice'] = fake_words_strongs_in_notice
        dict_notice_relevant_info[id_dict_notice]['notice_strong_words_real_total'] += len(
            real_words_strongs_in_notice)
        dict_notice_relevant_info[id_dict_notice]['notice_strong_words_fake_total'] += len(
            fake_words_strongs_in_notice)
        dict_notice_relevant_info[id_dict_notice]['notice_strong_words_total'] += (dict_notice_relevant_info[id_dict_notice]['notice_strong_words_real_total'] + dict_notice_relevant_info[id_dict_notice]['notice_strong_words_fake_total'] )

    print('Finish update dict of notices with relevant info ...\n')
    return dict_notice_relevant_info


def load_dict_notices_relevant_info_xlsx(folder_path, file_name, dict_notice_relevant_info):
    print('Starting loading dict of notices relevant info ...')
    file_path = os.path.join(folder_path, file_name)

    df = pd.read_excel(file_path)

    for row_id, row in df.iterrows():
        id_dict_notice = row_id
        id_notice = row['id_notice']
        title_notice = row['title_notice']
        real_words_strongs_in_notice = row['real_words_strongs_in_notice']
        fake_words_strongs_in_notice = row['fake_words_strongs_in_notice']
        notice_strong_words_real_total = row['notice_strong_words_real_total']
        notice_strong_words_fake_total = row['notice_strong_words_fake_total']
        notice_words_total_with_stopwords = row['notice_words_total_with_stopwords']
        notice_strong_words_total = row['notice_strong_words_total']
        classe_notice = row['classe_notice']
        notice_words_total_with_stopwords = row['notice_words_total_with_stopwords']
        notice_words_total_without_stopwords = row['notice_words_total_without_stopwords']

        dict_notice_relevant_info[id_dict_notice] = {}
        dict_notice_relevant_info[id_dict_notice]['id_notice'] = id_notice
        dict_notice_relevant_info[id_dict_notice]['title_notice'] = title_notice
        dict_notice_relevant_info[id_dict_notice]['real_words_strongs_in_notice'] = real_words_strongs_in_notice
        dict_notice_relevant_info[id_dict_notice]['fake_words_strongs_in_notice'] = fake_words_strongs_in_notice
        dict_notice_relevant_info[id_dict_notice]['notice_strong_words_real_total'] = notice_strong_words_real_total
        dict_notice_relevant_info[id_dict_notice]['notice_strong_words_fake_total'] = notice_strong_words_fake_total
        dict_notice_relevant_info[id_dict_notice]['notice_strong_words_total'] = notice_strong_words_total
        dict_notice_relevant_info[id_dict_notice]['classe_notice'] = classe_notice
        dict_notice_relevant_info[id_dict_notice]['notice_words_total_with_stopwords'] = notice_words_total_with_stopwords
        dict_notice_relevant_info[id_dict_notice]['notice_words_total_without_stopwords'] = notice_words_total_without_stopwords

    print('Finish load dict of notices relevant info\n')

    return dict_notice_relevant_info


def save_dict_notices_relevant_info_to_xlsx(file_path, dict_notice_relevant_info, group_name):
    print('Starting save dict of notices relevant info ...')

    df = pd.DataFrame.from_dict(dict_notice_relevant_info, orient='index')

    wb = Workbook()
    ws = wb.active
    ws.title = f'InfoRelDict de Noticias {group_name}'

    headers = list(dict_notice_relevant_info[next(
        iter(dict_notice_relevant_info))].keys())
    for col_num, value in enumerate(headers, start=1):
        ws.cell(row=1, column=col_num, value=value)

    for row_num, (id_notice, row_data) in enumerate(df.iterrows(), start=2):
        for col_num, value in enumerate(row_data, start=1):
            if isinstance(value, list):
                value = '\n'.join(value)
            ws.cell(row=row_num, column=col_num, value=value)

    info_dict_list = [
        f"- Dicionario de Noticias {group_name}",
        f"  Esse dicionario possui {len(dict_notice_relevant_info)} noticias",
        f"  Esse dicionario possui {dict_notice_relevant_info.get('notice_strong_words_real_total')} palavras fortes reais",
        f"  Esse dicionario possui {dict_notice_relevant_info.get('notice_strong_words_fake_total')} palavras fortes fakes",
        f"  Esse dicionario possui {dict_notice_relevant_info.get('notice_strong_words_total')} palavras fortes totais",
        "  Estrutura do dicionario:",
        "  ID - ID NOTICE - TITLE NOTICE - NOTICE ALL WORDS - CLASSE NOTICE - NOTICE WORDS TOTAL WITH STOPWORDS -  NOTICE WORDS TOTAL WITHOUT STOPWORDS"
    ]

    for i, info in enumerate(info_dict_list, start=len(df) + 2):
        ws.cell(row=i, column=1, value=info)

    wb.save(file_path)

    print('Finish save dict of notices relevant info\n')


def save_dict_notices_relevant_info_to_csv(file_path, dict_notice_relevant_info, group_name):
    print('Starting save dict of notices relevant info ...')

    pd.set_option('display.max_colwidth', None)

    df = pd.DataFrame.from_dict(dict_notice_relevant_info, orient='index')

    df['title_notice'] = df['title_notice'].astype(str)
    df['title_notice'] = df['title_notice'].apply(unidecode)

    info_dict_list = [
        f"- Dicionario de Noticias {group_name}",
        f"  Esse dicionario possui {len(dict_notice_relevant_info)} noticias",
        f"  Esse dicionario possui {dict_notice_relevant_info.get('notice_strong_words_real_total')} palavras fortes reais",
        f"  Esse dicionario possui {dict_notice_relevant_info.get('notice_strong_words_fake_total')} palavras fortes fakes",
        "  Estrutura do dicionario:",
        "  ID - ID NOTICE - TITLE NOTICE - NOTICE ALL WORDS - CLASSE NOTICE - NOTICE WORDS TOTAL WITH STOPWORDS -  NOTICE WORDS TOTAL WITHOUT STOPWORDS"
    ]

    header_df = pd.DataFrame({'info_dict': info_dict_list})
    df = pd.concat([header_df, df], ignore_index=True)

    df.to_csv(file_path, index=False, encoding='utf-8',
              quoting=csv.QUOTE_NONNUMERIC)

    print('Finish save dict of notices relevant info\n')

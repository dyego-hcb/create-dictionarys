import csv
import sys
import os
import pandas as pd

from unidecode import unidecode 

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..'))
sys.path.append(project_root)


def add_words_on_dict_group(dict_words_group, id_dict_words_group, word):

    if not any(entry['word'] == word for entry in dict_words_group.values()):
        id_dict_words_group += 1
        dict_words_group[id_dict_words_group] = {}
        dict_words_group[id_dict_words_group]['word'] = word
        dict_words_group[id_dict_words_group]['notices_appear_total'] = 0
        dict_words_group[id_dict_words_group]['ids_notice_appear'] = []
        dict_words_group[id_dict_words_group]['titles_notice_appear'] = []
        dict_words_group[id_dict_words_group]['classe_notice_word_appear'] = []
        dict_words_group[id_dict_words_group]['words_total_appear_in_notice'] = []
        dict_words_group[id_dict_words_group]['words_total_in_notice_without_stop_words'] = []
        dict_words_group[id_dict_words_group]['words_total_in_notice_with_stop_words'] = []
        dict_words_group[id_dict_words_group]['words_total_appear_in_group'] = 0
        dict_words_group[id_dict_words_group]['words_total_in_group_without_stop_words'] = 0
        dict_words_group[id_dict_words_group]['words_total_in_group_with_stop_words'] = 0

    return dict_words_group, id_dict_words_group


def create_dictionary_words_group(dict_words_group, dict_notice):
    print('Starting create dict of words group ...')
    id_dict_words_group = 0

    for id_dict_notice, notice_info in dict_notice.items():
        notice_words = notice_info['notice_content_stemm_without_stopwords']
        for word in notice_words:
            dict_words_group, id_dict_words_group = add_words_on_dict_group(
                dict_words_group, id_dict_words_group, word)

    dict_words_group_sorted = dict(
        sorted(dict_words_group.items(), key=lambda x: x[1]['word']))
    print('Finish create dict of words group\n')
    return dict_words_group_sorted


def update_dictionary_words_group(dict_words_group, dict_notice, total_words_group_with_stop_word, total_words_group_without_stop_word):
    print('Starting update dict of words group ...')

    for id_dict_words_group, word_info in dict_words_group.items():
        word = word_info['word']
        words_total_appear_in_group = 0
        word_append_on_notices_total = []
        words_total_in_notice_without_stop_words = []
        words_total_in_notice_with_stop_words = []
        ids_notice_appear = []
        titles_notice_appear = []
        classe_notice_word_appear = []

        for id_dict_notice, notice_info in dict_notice.items():
            notice_words = notice_info['notice_content_stemm_without_stopwords']
            words_total_appear_in_notice = notice_words.count(word)

            if words_total_appear_in_notice > 0:
                words_total_appear_in_group += words_total_appear_in_notice
                ids_notice_appear.append(notice_info['id_notice'])
                titles_notice_appear.append(notice_info['title_notice'])
                classe_notice_word_appear.append(notice_info['classe_notice'])
                word_append_on_notices_total.append(
                    words_total_appear_in_notice)
                words_total_in_notice_with_stop_words.append(
                    notice_info['notice_words_total_with_stopwords'])
                words_total_in_notice_without_stop_words.append(
                    notice_info['notice_words_total_without_stopwords'])

        dict_words_group[id_dict_words_group]['notices_appear_total'] = len(
            ids_notice_appear)
        dict_words_group[id_dict_words_group]['ids_notice_appear'] = ids_notice_appear
        dict_words_group[id_dict_words_group]['titles_notice_appear'] = titles_notice_appear
        dict_words_group[id_dict_words_group]['classe_notice_word_appear'] = classe_notice_word_appear
        dict_words_group[id_dict_words_group]['words_total_appear_in_notice'] = word_append_on_notices_total
        dict_words_group[id_dict_words_group]['words_total_in_notice_without_stop_words'] = words_total_in_notice_without_stop_words
        dict_words_group[id_dict_words_group]['words_total_in_notice_with_stop_words'] = words_total_in_notice_with_stop_words
        dict_words_group[id_dict_words_group]['words_total_appear_in_group'] += words_total_appear_in_group
        dict_words_group[id_dict_words_group]['words_total_in_group_without_stop_words'] = total_words_group_without_stop_word
        dict_words_group[id_dict_words_group]['words_total_in_group_with_stop_words'] = total_words_group_with_stop_word

    print('Finish update dict of words group\n')

    return dict_words_group


def save_dict_words_group_to_xlsx(file_path, dict_word, group_name):
    print('Starting save dict words group ...')

    df = pd.DataFrame.from_dict(dict_word, orient='index')

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df = df.sort_values(by='word')
        df.reset_index(drop=True, inplace=True)
        df.to_excel(
            writer, sheet_name=f'Dicionario de Palavras {group_name}', index=False)

        worksheet = writer.sheets[f'Dicionario de Palavras {group_name}']

        (max_row, max_col) = df.shape

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value)

        for row_num, row_data in enumerate(df.itertuples(index=False)):
            for col_num, value in enumerate(row_data):
                if isinstance(value, list):
                    df.at[row_num, df.columns[col_num]
                          ] = '\n'.join(map(str, value))
                    value = '\n'.join(map(str, value))
                worksheet.write(row_num, col_num, value)

        last_col = len(df.columns)
        df.insert(last_col, 'info_dict', '')

        info_dict_list = [
            f"- Dicionario de Palavras das Noticias Classificadas como {group_name}",
            f"  Esse dicionario possui {len(dict_word)} palavras",
            "  Estrutura do dicionario:",
            "  ID - " +
            "WORD - " +
            "NUMBER NOTICE'S ON WORD APPEAR - " +
            "ID'S NOTICES ON WORD APPEAR - " +
            "TITLE'S NOTICES ON WORD APPEAR -"
            "CLASSIFICATION NOTICE'S ON WORD APPEAR - " +
            "NUMBER ON WORD APPEAR ON NOTICE -  " +
            "NUMBER WORDS ON NOTICE WITH STOPWORDS - " +
            "NUMBER WORDS ON NOTICE WITHOUT STOPWORDS - " +
            "NUMBER ON WORD APPEAR IN GROUP - " +
            "NUMBER WORDS ON GROUP WITH STOPWORDS - " +
            "NUMBER WORDS ON GROUP WITHOUT STOPWORDS"
        ]

        for i in range(len(info_dict_list)):
            df.at[i, 'info_dict'] = info_dict_list[i]

        df.to_excel(
            writer, sheet_name=f'Dicionario de Palavras {group_name}', index=False)

    print('Finish save dict words group\n')


def save_dict_words_group_relevants_info_to_csv(file_path, dict_word, group_name):
    print('Starting save dict words group ...')

    pd.set_option('display.max_colwidth', None)

    selected_columns = ['word', 'words_total_appear_in_group', 'words_total_in_group_without_stop_words', 'words_total_in_group_with_stop_words']

    df = pd.DataFrame.from_dict(dict_word, orient='index')

    if selected_columns:
        df = df[selected_columns]

    df['info_dict'] = ''

    df['word'] = df['word'].astype(str)
    df = df.sort_values(by='word').reindex(columns=df.columns)

    info_dict_list = [
        f"- Dicionario de Palavras das Noticias Classificadas como {group_name}",
        f"  Esse dicionario possui {len(dict_word)} palavras",
        "  Estrutura do dicionario:",
        "  ID - " +
        "WORD - " +
        "NUMBER ON WORD APPEAR IN GROUP - " +
        "NUMBER WORDS ON GROUP WITH STOPWORDS - " +
        "NUMBER WORDS ON GROUP WITHOUT STOPWORDS"
    ]

    df.at[1, 'info_dict'] = info_dict_list[0]
    for i in range(1, len(info_dict_list)):
        df.at[i+1, 'info_dict'] = info_dict_list[i]
    
    df.to_csv(file_path, index=False, encoding='utf-8', quoting=csv.QUOTE_NONNUMERIC)

    print('Finish save dict words group\n')

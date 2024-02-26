import csv
import sys
import os
import openpyxl
import pandas as pd

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..'))
sys.path.append(project_root)

from unidecode import unidecode

def add_words_on_dict_words(dict_words, id_dict_words, word):

    if not any(entry['word'] == word for entry in dict_words.values()):
        id_dict_words += 1
        dict_words[id_dict_words] = {}
        dict_words[id_dict_words]['word'] = word
        dict_words[id_dict_words]['notices_appear_total'] = 0
        dict_words[id_dict_words]['ids_notice_appear'] = []
        dict_words[id_dict_words]['classe_notice_word_appear'] = []
        dict_words[id_dict_words]['words_total_appear_in_notice'] = []
        dict_words[id_dict_words]['words_total_in_notice_without_stop_words'] = []
        dict_words[id_dict_words]['words_total_in_notice_with_stop_words'] = []
        dict_words[id_dict_words]['words_total_in_group_real_without_stop_words'] = 0
        dict_words[id_dict_words]['words_total_in_group_real_with_stop_words'] = 0
        dict_words[id_dict_words]['words_total_in_group_fake_without_stop_words'] = 0
        dict_words[id_dict_words]['words_total_in_group_fake_with_stop_words'] = 0
        dict_words[id_dict_words]['words_total_appear_in_both_group'] = 0
        dict_words[id_dict_words]['words_total_appear_in_group_real'] = 0
        dict_words[id_dict_words]['words_total_appear_in_group_fake'] = 0
        dict_words[id_dict_words]['percet_strong_word_in_group_real'] = 0
        dict_words[id_dict_words]['percet_strong_word_in_group_fake'] = 0

    return dict_words, id_dict_words


def create_dictionary_words(dict_words, dict_words_group):
    print('Starting create dict of words ...')
    id_dict_words = len(dict_words)

    for id_group, word_group_info in dict_words_group.items():
        word = word_group_info['word']
        dict_words, id_dict_words = add_words_on_dict_words(
            dict_words, id_dict_words, word)

    dict_words_sorted = dict(
        sorted(dict_words.items(), key=lambda x: x[1]['word']))
    print('Finish create dict of words\n')
    return dict_words_sorted


def update_dictionary_words(dict_words, dict_words_group, classe_group):
    print('Starting update dict of words ...')

    for id_dict_words, word_dict_info in dict_words.items():
        word = word_dict_info['word']
        for id_group, word_group_info in dict_words_group.items():
            word_group = word_group_info['word']

            if word in word_group:
                dict_words[id_dict_words]['notices_appear_total'] += len(
                    dict_words_group[id_group]['ids_notice_appear'])
                dict_words[id_dict_words]['ids_notice_appear'].extend(
                    dict_words_group[id_group]['ids_notice_appear'])
                dict_words[id_dict_words]['classe_notice_word_appear'].extend(
                    dict_words_group[id_group]['classe_notice_word_appear'])
                dict_words[id_dict_words]['words_total_appear_in_notice'].extend(
                    dict_words_group[id_group]['words_total_appear_in_notice'])
                dict_words[id_dict_words]['words_total_in_notice_without_stop_words'].extend(
                    dict_words_group[id_group]['words_total_in_notice_without_stop_words'])
                dict_words[id_dict_words]['words_total_in_notice_with_stop_words'].extend(
                    dict_words_group[id_group]['words_total_in_notice_with_stop_words'])

                if (classe_group == 1):
                    words_total_appear_in_group_real = dict_words_group[
                        id_group]['words_total_appear_in_group']
                    words_total_in_group_real_without_stop_words = dict_words_group[
                        id_group]['words_total_in_group_without_stop_words']
                    words_total_in_group_real_with_stop_words = dict_words_group[
                        id_group]['words_total_in_group_with_stop_words']
                    dict_words[id_dict_words]['words_total_in_group_real_without_stop_words'] = words_total_in_group_real_without_stop_words
                    dict_words[id_dict_words]['words_total_in_group_real_with_stop_words'] = words_total_in_group_real_with_stop_words
                    dict_words[id_dict_words]['words_total_appear_in_group_real'] = words_total_appear_in_group_real
                    dict_words[id_dict_words]['words_total_appear_in_both_group'] += words_total_appear_in_group_real
                else:
                    words_total_appear_in_group_fake = dict_words_group[
                        id_group]['words_total_appear_in_group']
                    words_total_in_group_fake_without_stop_words = dict_words_group[
                        id_group]['words_total_in_group_without_stop_words']
                    words_total_in_group_fake_with_stop_words = dict_words_group[
                        id_group]['words_total_in_group_with_stop_words']
                    dict_words[id_dict_words]['words_total_in_group_fake_without_stop_words'] = words_total_in_group_fake_without_stop_words
                    dict_words[id_dict_words]['words_total_in_group_fake_with_stop_words'] = words_total_in_group_fake_with_stop_words
                    dict_words[id_dict_words]['words_total_appear_in_group_fake'] = words_total_appear_in_group_fake
                    dict_words[id_dict_words]['words_total_appear_in_both_group'] += words_total_appear_in_group_fake

    print('Finish update dict of words\n')

    return dict_words


def calculate_percent_to_strong_word(dict_words):
    for id_dict_words, word_dict_info in dict_words.items():

        percet_strong_word_in_group_fake = (
            word_dict_info['words_total_appear_in_group_fake'] / word_dict_info['words_total_appear_in_both_group']) * 100
        percet_strong_word_in_group_real = (
            word_dict_info['words_total_appear_in_group_real'] / word_dict_info['words_total_appear_in_both_group']) * 100

        dict_words[id_dict_words]['percet_strong_word_in_group_fake'] = percet_strong_word_in_group_fake
        dict_words[id_dict_words]['percet_strong_word_in_group_real'] = percet_strong_word_in_group_real

    return dict_words


def load_dict_words_xlsx(folder_path, file_name, dict_words):
    print('Starting loading dict words ...')
    file_path = os.path.join(folder_path, file_name)

    df = pd.read_excel(file_path)

    for row_id, row in df.iterrows():
        id_dict_words = row_id
        word = row['word']
        notices_appear_total = row['notices_appear_total']
        ids_notice_appear = row['ids_notice_appear']
        classe_notice_word_appear = row['classe_notice_word_appear']
        words_total_appear_in_notice = row['words_total_appear_in_notice']
        words_total_in_notice_without_stop_words = row['words_total_in_notice_without_stop_words']
        words_total_in_notice_with_stop_words = row['words_total_in_notice_with_stop_words']
        words_total_in_group_real_without_stop_words = row['words_total_in_group_real_without_stop_words']
        words_total_in_group_real_with_stop_words = row['words_total_in_group_real_with_stop_words']
        words_total_in_group_real_with_stop_words = row['words_total_in_group_real_with_stop_words']
        words_total_in_group_fake_without_stop_words = row['words_total_in_group_fake_without_stop_words']
        words_total_in_group_fake_with_stop_words = row['words_total_in_group_fake_with_stop_words']
        words_total_appear_in_both_group = row['words_total_appear_in_both_group']
        words_total_appear_in_group_real = row['words_total_appear_in_group_real']
        words_total_appear_in_group_fake = row['words_total_appear_in_group_fake']
        percet_strong_word_in_group_fake = row['percet_strong_word_in_group_fake']
        percet_strong_word_in_group_real = row['percet_strong_word_in_group_real']

        dict_words[id_dict_words] = {}
        dict_words[id_dict_words]['word'] = word
        dict_words[id_dict_words]['notices_appear_total'] = notices_appear_total
        dict_words[id_dict_words]['ids_notice_appear'] = ids_notice_appear
        dict_words[id_dict_words]['classe_notice_word_appear'] = classe_notice_word_appear
        dict_words[id_dict_words]['words_total_appear_in_notice'] = words_total_appear_in_notice
        dict_words[id_dict_words]['words_total_in_notice_without_stop_words'] = words_total_in_notice_without_stop_words
        dict_words[id_dict_words]['words_total_in_notice_with_stop_words'] = words_total_in_notice_with_stop_words
        dict_words[id_dict_words]['words_total_in_group_real_without_stop_words'] = words_total_in_group_real_without_stop_words
        dict_words[id_dict_words]['words_total_in_group_real_with_stop_words'] = words_total_in_group_real_with_stop_words
        dict_words[id_dict_words]['words_total_in_group_fake_without_stop_words'] = words_total_in_group_fake_without_stop_words
        dict_words[id_dict_words]['words_total_in_group_fake_with_stop_words'] = words_total_in_group_fake_with_stop_words
        dict_words[id_dict_words]['words_total_appear_in_both_group'] = words_total_appear_in_both_group
        dict_words[id_dict_words]['words_total_appear_in_group_real'] = words_total_appear_in_group_real
        dict_words[id_dict_words]['words_total_appear_in_group_fake'] = words_total_appear_in_group_fake
        dict_words[id_dict_words]['percet_strong_word_in_group_real'] = percet_strong_word_in_group_real
        dict_words[id_dict_words]['percet_strong_word_in_group_fake'] = percet_strong_word_in_group_fake

    print('Finish load dict of words\n')

    return dict_words


def save_dict_words_to_xlsx(file_path, dict_word, group_name):
    print('Starting save dict words ...')

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
            "CLASSIFICATION NOTICE'S ON WORD APPEAR - " +
            "NUMBER ON WORD APPEAR ON NOTICE -  " +
            "NUMBER WORDS ON NOTICE REAL WITH STOPWORDS - " +
            "NUMBER WORDS ON NOTICE REAL WITHOUT STOPWORDS - " +
            "NUMBER WORDS ON NOTICE FAKE WITH STOPWORDS - " +
            "NUMBER WORDS ON NOTICE FAKE WITHOUT STOPWORDS - " +
            "NUMBER ON WORD APPEAR IN GROUP REAL - " +
            "NUMBER ON WORD APPEAR IN GROUP FAKE - " +
            "PERCENTAGE OF BEING A STRONG WORD IN GROUP OF REAL WORDS - " +
            "PERCENTAGE OF BEING A STRONG WORD IN GROUP OF FAKE WORDS"
        ]

        for i in range(len(info_dict_list)):
            df.at[i, 'info_dict'] = info_dict_list[i]

        df.to_excel(
            writer, sheet_name=f'Dicionario de Palavras {group_name}', index=False)

    print('Finish save dict words\n')


def save_dict_words_relevants_info_to_csv(file_path, dict_word, group_name):
    print('Starting save dict words ...')

    pd.set_option('display.max_colwidth', None)

    selected_columns = ['word', 'words_total_appear_in_both_group', 'words_total_appear_in_group_real', 'words_total_appear_in_group_fake',
                        'percet_strong_word_in_group_real', 'percet_strong_word_in_group_fake']

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
        "NUMBER TOTAL ON WORD APPEAR IN BOTH GROUP - " +
        "NUMBER ON WORD APPEAR IN GROUP REAL - " +
        "NUMBER ON WORD APPEAR IN GROUP FAKE - " +
        "PERCENTAGE OF BEING A STRONG WORD IN GROUP OF REAL WORDS - " +
        "PERCENTAGE OF BEING A STRONG WORD IN GROUP OF FAKE WORDS"
    ]

    df.at[1, 'info_dict'] = info_dict_list[0]
    for i in range(1, len(info_dict_list)):
        df.at[i+1, 'info_dict'] = info_dict_list[i]

    df.to_csv(file_path, index=False, encoding='utf-8',
              quoting=csv.QUOTE_NONNUMERIC)

    print('Finish save dict words\n')

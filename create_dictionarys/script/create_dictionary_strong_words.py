import sys
import os
import csv
import pandas as pd

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..'))
sys.path.append(project_root)

def add_words_on_dict_strong_words(dict_strong_words, id_dict_strong_words, word, words_total_appear_in_both_group, words_total_appear_in_group_real, percet_strong_word_in_group_real, words_total_appear_in_group_fake, percet_strong_word_in_group_fake):

    if not any(entry['word'] == word for entry in dict_strong_words.values()):
        id_dict_strong_words += 1
        dict_strong_words[id_dict_strong_words] = {}
        dict_strong_words[id_dict_strong_words]['word'] = word
        dict_strong_words[id_dict_strong_words]['words_total_appear_in_both_group'] = words_total_appear_in_both_group
        dict_strong_words[id_dict_strong_words]['words_total_appear_in_group_real'] = words_total_appear_in_group_real
        dict_strong_words[id_dict_strong_words]['percet_strong_word_in_group_real'] = percet_strong_word_in_group_real
        dict_strong_words[id_dict_strong_words]['words_total_appear_in_group_fake'] = words_total_appear_in_group_fake
        dict_strong_words[id_dict_strong_words]['percet_strong_word_in_group_fake'] = percet_strong_word_in_group_fake

    return dict_strong_words, id_dict_strong_words


def create_dictionary_strong_words(dict_strong_words, dict_words, class_dict):
    print('Starting create dict of strong words ...')
    id_dict_strong_words = len(dict_strong_words)

    for id_word, words_info in dict_words.items():
        word = words_info['word']
        words_total_appear_in_both_group = words_info['words_total_appear_in_both_group']
        words_total_appear_in_group_real = words_info['words_total_appear_in_group_real']
        percet_strong_word_in_group_real = words_info['percet_strong_word_in_group_real']
        words_total_appear_in_group_fake = words_info['words_total_appear_in_group_fake']
        percet_strong_word_in_group_fake = words_info['percet_strong_word_in_group_fake']

        if (class_dict == 1):
            if (percet_strong_word_in_group_real >= 70):
                dict_strong_words, id_dict_strong_words = add_words_on_dict_strong_words(
                    dict_strong_words, id_dict_strong_words, word, words_total_appear_in_both_group, words_total_appear_in_group_real, percet_strong_word_in_group_real, words_total_appear_in_group_fake, percet_strong_word_in_group_fake)
        else:
            if (percet_strong_word_in_group_fake >= 70):
                dict_strong_words, id_dict_strong_words = add_words_on_dict_strong_words(
                    dict_strong_words, id_dict_strong_words, word, words_total_appear_in_both_group, words_total_appear_in_group_real, percet_strong_word_in_group_real, words_total_appear_in_group_fake, percet_strong_word_in_group_fake)

    dict_strong_words_sorted = dict(
        sorted(dict_strong_words.items(), key=lambda x: x[1]['word']))
    print('Finish create dict of strong words\n')

    return dict_strong_words_sorted


def load_dict_strong_wrods_xlsx(folder_path, file_name, dict_strong_words):
    print('Starting loading dict strong words ...')

    file_path = os.path.join(folder_path, file_name)

    df = pd.read_excel(file_path)

    for row_id, row in df.iterrows():
        id_dict_strong_words = row_id
        word = row['word']
        words_total_appear_in_both_group = row['words_total_appear_in_both_group']
        words_total_appear_in_group_real = row['words_total_appear_in_group_real']
        percet_strong_word_in_group_real = row['percet_strong_word_in_group_real']
        words_total_appear_in_group_fake = row['words_total_appear_in_group_fake']
        percet_strong_word_in_group_fake = row['percet_strong_word_in_group_fake']

        if id_dict_strong_words not in dict_strong_words:
            dict_strong_words[id_dict_strong_words] = {}

        dict_strong_words[id_dict_strong_words]['word'] = row['word']
        dict_strong_words[id_dict_strong_words]['words_total_appear_in_both_group'] = row['words_total_appear_in_both_group']
        dict_strong_words[id_dict_strong_words]['words_total_appear_in_group_real'] = row['words_total_appear_in_group_real']
        dict_strong_words[id_dict_strong_words]['percet_strong_word_in_group_real'] = row['percet_strong_word_in_group_real']
        dict_strong_words[id_dict_strong_words]['words_total_appear_in_group_fake'] = row['words_total_appear_in_group_fake']
        dict_strong_words[id_dict_strong_words]['percet_strong_word_in_group_fake'] = row['percet_strong_word_in_group_fake']

    print('Finish load dict of strong words\n')

    return dict_strong_words


def save_dict_strong_words_to_xlsx(file_path, dict_strong_word, group_name):
    print('Starting save dict strong words ...')

    df = pd.DataFrame.from_dict(dict_strong_word, orient='index')

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df = df.sort_values(by='word')
        df = pd.concat([pd.DataFrame([df.columns], columns=df.columns), df])
        df.reset_index(drop=True, inplace=True)
        df.to_excel(
            writer, sheet_name=f'Dicionario de Palavras Fortes {group_name}', index=False)

        worksheet = writer.sheets[f'Dicionario de Palavras Fortes {group_name}']

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
            f"- Dicionario de Palavras Fortes {group_name}",
            f"  Esse dicionario possui {len(dict_strong_word)} palavras",
            "  Estrutura do dicionario:",
            "  ID - " +
            "WORD - " +
            "NUMBER ON WORD APPEAR IN GROUP REAL - " +
            "PERCENTAGE OF BEING A STRONG WORD IN GROUP OF REAL WORDS - " +
            "NUMBER ON WORD APPEAR IN GROUP FAKE - " +
            "PERCENTAGE OF BEING A STRONG WORD IN GROUP OF FAKE WORDS"
        ]

        for i in range(len(info_dict_list)):
            df.at[i, 'info_dict'] = info_dict_list[i]

        df.to_excel(
            writer, sheet_name=f'Dicionario de Palavras {group_name}', index=False)

    print('Finish save dict strong words\n')


def save_dict_strong_words_to_csv(file_path, dict_strong_word, group_name):
    print('Starting save dict strong words ...')

    pd.set_option('display.max_colwidth', None)

    df = pd.DataFrame.from_dict(dict_strong_word, orient='index')

    df['info_dict'] = ''

    df['word'] = df['word'].astype(str)
    df = df.sort_values(by='word').reindex(columns=df.columns)

    info_dict_list = [
        f"- Dicionario de Palavras Fortes {group_name}",
        f"  Esse dicionario possui {len(dict_strong_word)} palavras",
        "  Estrutura do dicionario:",
        "  ID - " +
        "WORD - " +
        "NUMBER ON WORD APPEAR IN GROUP REAL - " +
        "PERCENTAGE OF BEING A STRONG WORD IN GROUP OF REAL WORDS - " +
        "NUMBER ON WORD APPEAR IN GROUP FAKE - " +
        "PERCENTAGE OF BEING A STRONG WORD IN GROUP OF FAKE WORDS"
    ]

    df.at[1, 'info_dict'] = info_dict_list[0]
    for i in range(1, len(info_dict_list)):
        df.at[i+1, 'info_dict'] = info_dict_list[i]

    df.to_csv(file_path, index=False, encoding='utf-8',
              quoting=csv.QUOTE_NONNUMERIC)

    print('Finish save dict strong words\n')

print("\nStarting execute ORI methods\n")
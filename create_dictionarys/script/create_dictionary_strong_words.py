import sys
import os
import csv
import openpyxl
import pandas as pd

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..'))
sys.path.append(project_root)

from create_dictionary_notices import create_dictionary_notices, create_dictionary_notices_relevant_info, update_dictionary_notices_relevant_info, load_dict_notices_xlsx, save_dict_notices_to_xlsx, save_dict_notices_to_csv, load_dict_notices_relevant_info_xlsx, save_dict_notices_relevant_info_to_xlsx, save_dict_notices_relevant_info_to_csv
from script.create_dictionary_words_group import create_dictionary_words_group, update_dictionary_words_group, load_dict_words_group_xlsx, save_dict_words_group_to_xlsx, save_dict_words_group_relevants_info_to_csv
from script.create_dictionary_words import create_dictionary_words, update_dictionary_words, calculate_percent_to_strong_word, load_dict_words_xlsx, save_dict_words_to_xlsx, save_dict_words_relevants_info_to_csv

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

        dict_strong_words[id_dict_strong_words] = {}
        dict_strong_words[id_dict_strong_words]['word'] = word
        dict_strong_words[id_dict_strong_words]['words_total_appear_in_both_group'] = words_total_appear_in_both_group
        dict_strong_words[id_dict_strong_words]['words_total_appear_in_group_real'] = words_total_appear_in_group_real
        dict_strong_words[id_dict_strong_words]['percet_strong_word_in_group_real'] = percet_strong_word_in_group_real
        dict_strong_words[id_dict_strong_words]['words_total_appear_in_group_fake'] = words_total_appear_in_group_fake
        dict_strong_words[id_dict_strong_words]['percet_strong_word_in_group_fake'] = percet_strong_word_in_group_fake

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


path_bd = os.path.join(project_root, 'base_data', 'FakeRecogna.xlsx')
path_load_dicts = os.path.join(project_root, 'output')

dict_notice_boths: dict = {}
# words_total_with_stopwords_both_group = 0
# words_total_without_stopwords_both_group = 0

# dict_notice_boths, words_total_with_stopwords_both_group, words_total_without_stopwords_both_group = create_dictionary_notices(path_bd, dict_notice_boths, 1)
# dict_notice_boths, words_total_with_stopwords_both_group, words_total_without_stopwords_both_group = create_dictionary_notices(path_bd, dict_notice_boths, 0)

# outpat_notice = os.path.join(project_root, 'output', 'dict_notice_boths.csv')
# save_dict_notices_to_csv(outpat_notice, dict_notice_boths, 'Boths', words_total_with_stopwords_both_group, words_total_without_stopwords_both_group)

# outpat_notice = os.path.join(project_root, 'output', 'dict_notice_boths.xlsx')
# save_dict_notices_to_xlsx(outpat_notice, dict_notice_boths, 'Boths', words_total_with_stopwords_both_group, words_total_without_stopwords_both_group)

dict_notice_reais: dict = {}
words_total_with_stopwords_group_reais = 0
words_total_without_stopwords_group_reais = 0

# dict_notice_reais = load_dict_notices_xlsx(
#     path_load_dicts, "dict_notice_reais.xlsx", dict_notice_reais)

dict_notice_reais, words_total_with_stopwords_group_reais, words_total_without_stopwords_group_reais = create_dictionary_notices(path_bd, dict_notice_reais, 1)

# outpat_notice = os.path.join(project_root, 'output', 'dict_notice_reais.csv')
# save_dict_notices_to_csv(outpat_notice, dict_notice_reais, 'Reais', words_total_with_stopwords_group_reais, words_total_without_stopwords_group_reais)

# outpat_notice = os.path.join(project_root, 'output', 'dict_notice_reais.xlsx')
# save_dict_notices_to_xlsx(outpat_notice, dict_notice_reais, 'Reais', words_total_with_stopwords_group_reais, words_total_without_stopwords_group_reais)

dict_words_reais: dict = {}
# words_dict_reais_total = 0

# dict_words_reais = load_dict_words_group_xlsx(
#     path_load_dicts, "dict_words_reais.xlsx", dict_words_reais)

# outpat_words = os.path.join(project_root, 'output', 'teste_load_dict_words_reais.xlsx')
# save_dict_words_group_to_xlsx(outpat_words, dict_words_reais, 'Reais')

# dict_words_reais = create_dictionary_words_group(dict_words_reais, dict_notice_reais)
# dict_words_reais = update_dictionary_words_group(dict_words_reais, dict_notice_reais, words_total_with_stopwords_group_reais, words_total_without_stopwords_group_reais)

# words_dict_reais_total = len(dict_words_reais)

# outpat_words = os.path.join(project_root, 'output', 'dict_words_reais_relevant_info.csv')
# save_dict_words_group_relevants_info_to_csv(outpat_words, dict_words_reais, 'Reais')

# outpat_words = os.path.join(project_root, 'output', 'dict_words_reais.xlsx')
# save_dict_words_group_to_xlsx(outpat_words, dict_words_reais, 'Reais')

dict_notice_fakes: dict = {}
words_total_with_stopwords_group_fakes = 0
words_total_without_stopwords_group_fakes = 0

# dict_notice_fakes = load_dict_notices_xlsx(
#     path_load_dicts, "dict_notice_fakes.xlsx", dict_notice_fakes)

dict_notice_fakes, words_total_with_stopwords_group_fakes, words_total_without_stopwords_group_fakes = create_dictionary_notices(path_bd, dict_notice_fakes, 0)

# outpat_notice = os.path.join(project_root, 'output', 'dict_notice_fakes.csv')
# save_dict_notices_to_csv(outpat_notice, dict_notice_fakes, 'Fakes', words_total_with_stopwords_group_fakes, words_total_without_stopwords_group_fakes)

# outpat_notice = os.path.join(project_root, 'output', 'dict_notice_fakes.xlsx')
# save_dict_notices_to_xlsx(outpat_notice, dict_notice_fakes, 'Fakes', words_total_with_stopwords_group_fakes, words_total_without_stopwords_group_fakes)

dict_words_fakes: dict = {}
# words_dict_fakes_total = 0

# dict_words_fakes = load_dict_words_group_xlsx(
#     path_load_dicts, "dict_words_fakes.xlsx", dict_words_fakes)

# dict_words_fakes = create_dictionary_words_group(dict_words_fakes, dict_notice_fakes)
# dict_words_fakes = update_dictionary_words_group(dict_words_fakes, dict_notice_fakes, words_total_with_stopwords_group_fakes, words_total_without_stopwords_group_fakes)

# words_dict_fakes_total = len(dict_words_fakes)

# outpat_words = os.path.join(project_root, 'output', 'dict_words_fakes_relavant_info.csv')
# save_dict_words_group_relevants_info_to_csv(outpat_words, dict_words_fakes, 'Fakes')

# outpat_words = os.path.join(project_root, 'output', 'dict_words_fakes.xlsx')
# save_dict_words_group_to_xlsx(outpat_words, dict_words_fakes, 'Fakes')

dict_words: dict = {}
# word_dict_words_total = 0

# dict_words = load_dict_words_xlsx(path_load_dicts, "dict_words.xlsx", dict_words)

# dict_words = create_dictionary_words(dict_words, dict_words_reais)
# dict_words = create_dictionary_words(dict_words, dict_words_fakes)
# dict_words = update_dictionary_words(dict_words, dict_words_reais, 1)
# dict_words = update_dictionary_words(dict_words, dict_words_fakes, 0)
# dict_words = calculate_percent_to_strong_word(dict_words);

# word_dict_words_total = len(dict_words)

# outpat_words = os.path.join(project_root, 'output', 'dict_words_relevant_info.csv')
# save_dict_words_relevants_info_to_csv(outpat_words, dict_words, 'Geral')

# outpat_words = os.path.join(project_root, 'output', 'dict_words.xlsx')
# save_dict_words_to_xlsx(outpat_words, dict_words, 'Geral')

dict_strong_words_reais: dict = {}
# strong_words_dict_reais_total = 0

dict_strong_words_reais = load_dict_strong_wrods_xlsx(
    path_load_dicts, "dict_strong_words_reais_pre_processado_25.xlsx", dict_strong_words_reais)

outpat_strong_words = os.path.join(project_root, 'output', 'teste_load_dict_strong_words_25_reais.xlsx')
save_dict_strong_words_to_xlsx(outpat_strong_words, dict_strong_words_reais, 'R')

# dict_strong_words_reais = create_dictionary_strong_words(dict_strong_words_reais, dict_words, 1)

# strong_words_dict_reais_total = len(dict_strong_words_reais)

# outpat_strong_words = os.path.join(project_root, 'output', 'dict_strong_words_reais.xlsx')
# save_dict_strong_words_to_xlsx(outpat_strong_words, dict_strong_words_reais, 'R')

# outpat_strong_words = os.path.join(project_root, 'output', 'dict_strong_words_reais_relevant_info.csv')
# save_dict_strong_words_to_csv(outpat_strong_words, dict_strong_words_reais, 'Reais')

dict_strong_words_fakes: dict = {}
# strong_words_dict_fakes_total = 0

dict_strong_words_fakes = load_dict_strong_wrods_xlsx(
    path_load_dicts, "dict_strong_words_fakes_pre_processado_25.xlsx", dict_strong_words_fakes)

outpat_strong_words = os.path.join(project_root, 'output', 'teste_load_dict_strong_25_words_fakes.xlsx')
save_dict_strong_words_to_xlsx(outpat_strong_words, dict_strong_words_fakes, 'F')

# dict_strong_words_fakes = create_dictionary_strong_words(dict_strong_words_fakes, dict_words, 0)
# strong_words_dict_fakes_total = len(dict_strong_words_fakes)

# outpat_strong_words = os.path.join(project_root, 'output', 'dict_strong_words_fakes.xlsx')
# save_dict_strong_words_to_xlsx(outpat_strong_words, dict_strong_words_fakes, 'F')

# outpat_strong_words = os.path.join(project_root, 'output', 'dict_strong_words_fakes_relevant_info.csv')
# save_dict_strong_words_to_csv(outpat_strong_words, dict_strong_words_fakes, 'Fakes')

dict_notice_boths_relevant_info: dict = {}

dict_notice_boths_relevant_info = create_dictionary_notices_relevant_info(
    dict_notice_boths_relevant_info, dict_notice_reais)
dict_notice_boths_relevant_info = create_dictionary_notices_relevant_info(
    dict_notice_boths_relevant_info, dict_notice_fakes)

dict_notice_boths_relevant_info = update_dictionary_notices_relevant_info(
    dict_notice_boths_relevant_info, dict_notice_reais, dict_strong_words_reais, 1)
dict_notice_boths_relevant_info = update_dictionary_notices_relevant_info(
    dict_notice_boths_relevant_info, dict_notice_reais, dict_strong_words_fakes, 0)

dict_notice_boths_relevant_info = update_dictionary_notices_relevant_info(
    dict_notice_boths_relevant_info, dict_notice_fakes, dict_strong_words_reais, 1)
dict_notice_boths_relevant_info = update_dictionary_notices_relevant_info(
    dict_notice_boths_relevant_info, dict_notice_fakes, dict_strong_words_fakes, 0)

outpat_notice_reais_relevant_info_xlsx = os.path.join(
    project_root, 'output', 'dict_notice_relevant_info_boths.xlsx')
save_dict_notices_relevant_info_to_xlsx(
    outpat_notice_reais_relevant_info_xlsx, dict_notice_boths_relevant_info, 'Boths')

outpat_notice_reais_relevant_info_csv = os.path.join(
    project_root, 'output', 'dict_notice_relevant_info_boths.csv')
save_dict_notices_relevant_info_to_csv(
    outpat_notice_reais_relevant_info_csv, dict_notice_boths_relevant_info, 'Boths')

dict_notice_reais_relevant_info: dict = {}
dict_notice_reais_relevant_info = create_dictionary_notices_relevant_info(
    dict_notice_reais_relevant_info, dict_notice_reais)
dict_notice_reais_relevant_info = update_dictionary_notices_relevant_info(
    dict_notice_reais_relevant_info, dict_notice_reais, dict_strong_words_reais, 1)
dict_notice_reais_relevant_info = update_dictionary_notices_relevant_info(
    dict_notice_reais_relevant_info, dict_notice_reais, dict_strong_words_fakes, 0)

outpat_notice_reais_relevant_info_xlsx = os.path.join(
    project_root, 'output', 'dict_notice_relevant_info_reais.xlsx')
save_dict_notices_relevant_info_to_xlsx(
    outpat_notice_reais_relevant_info_xlsx, dict_notice_reais_relevant_info, 'Reais')

outpat_notice_reais_relevant_info_csv = os.path.join(
    project_root, 'output', 'dict_notice_relevant_info_reais.csv')
save_dict_notices_relevant_info_to_csv(
    outpat_notice_reais_relevant_info_csv, dict_notice_reais_relevant_info, 'Reais')

dict_notice_fakes_relevant_info: dict = {}
dict_notice_fakes_relevant_info = create_dictionary_notices_relevant_info(
    dict_notice_fakes_relevant_info, dict_notice_fakes)
dict_notice_fakes_relevant_info = update_dictionary_notices_relevant_info(
    dict_notice_fakes_relevant_info, dict_notice_fakes, dict_strong_words_reais, 1)
dict_notice_fakes_relevant_info = update_dictionary_notices_relevant_info(
    dict_notice_fakes_relevant_info, dict_notice_fakes, dict_strong_words_fakes, 0)

outpat_notice_fakes_relevant_info_xlsx = os.path.join(
    project_root, 'output', 'dict_notice_relevant_info_fakes.xlsx')
save_dict_notices_relevant_info_to_xlsx(
    outpat_notice_fakes_relevant_info_xlsx, dict_notice_fakes_relevant_info, 'Fakes')

outpat_notice_fakes_relevant_info_csv = os.path.join(
    project_root, 'output', 'dict_notice_relevant_info_fakes.csv')
save_dict_notices_relevant_info_to_xlsx(
    outpat_notice_fakes_relevant_info_csv, dict_notice_fakes_relevant_info, 'Fakes')

print("\nFinish execute ORI methods\n")

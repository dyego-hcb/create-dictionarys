import sys
import os
import pandas as pd

from create_dictionary_notices import create_dictionary_notices, save_dict_notice_to_xlsx

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..'))
sys.path.append(project_root)

def add_words_on_dict(dict_words, id_dict_words, word):
    
    if not any(entry['word'] == word for entry in dict_words.values()):
        id_dict_words +=1
        dict_words[id_dict_words] = {}
        dict_words[id_dict_words]['word'] = word
        dict_words[id_dict_words]['notices_appear_total'] = 0
        dict_words[id_dict_words]['ids_notice_appear'] = []
        dict_words[id_dict_words]['titles_notice_appear'] = []
        dict_words[id_dict_words]['classe_notice_word_appear'] = []
        dict_words[id_dict_words]['words_total_appear_in_notice'] = []
        dict_words[id_dict_words]['words_total_in_notice_without_stop_words'] = []
        dict_words[id_dict_words]['words_total_in_notice_with_stop_words'] = []
        dict_words[id_dict_words]['words_total_appear_in_group'] = 0
        dict_words[id_dict_words]['words_total_in_group_without_stop_words'] = 0
        dict_words[id_dict_words]['words_total_in_group_with_stop_words'] = 0

    return dict_words, id_dict_words

def update_dictionary_words(dict_words, dict_notice, total_words_group_with_stop_word, total_words_group_without_stop_word):
    print('Starting update dict of words ...')

    for id_word, word_info in dict_words.items():
        word = word_info['word']
        words_total_appear_in_group = 0
        word_append_on_notices_total = []
        words_total_in_notice_without_stop_words = []
        words_total_in_notice_with_stop_words = []
        ids_notice_appear = []
        titles_notice_appear = []
        classe_notice_word_appear = []

        for id_notice, notice_info in dict_notice.items():
            notice_words = notice_info['notice_content_stemm_without_stopwords']
            words_total_appear_in_notice = notice_words.count(word)

            if words_total_appear_in_notice > 0:
                words_total_appear_in_group += words_total_appear_in_notice
                ids_notice_appear.append(notice_info['id_notice'])
                titles_notice_appear.append(notice_info['title_notice'])
                classe_notice_word_appear.append(notice_info['classe_notice'])
                word_append_on_notices_total.append(words_total_appear_in_notice)
                words_total_in_notice_with_stop_words.append(notice_info['notice_words_total_with_stopwords'])
                words_total_in_notice_without_stop_words.append(notice_info['notice_words_total_without_stopwords'])                
        
        notices_appear_total = len(ids_notice_appear)

        dict_words[id_word]['notices_appear_total'] = notices_appear_total
        dict_words[id_word]['ids_notice_appear'] = ids_notice_appear
        dict_words[id_word]['titles_notice_appear'] = titles_notice_appear
        dict_words[id_word]['classe_notice_word_appear'] = classe_notice_word_appear
        dict_words[id_word]['words_total_appear_in_notice'] = word_append_on_notices_total
        dict_words[id_word]['words_total_in_notice_without_stop_words'] = words_total_in_notice_without_stop_words
        dict_words[id_word]['words_total_in_notice_with_stop_words'] = words_total_in_notice_with_stop_words
        dict_words[id_word]['words_total_appear_in_group'] += words_total_appear_in_group
        dict_words[id_word]['words_total_in_group_without_stop_words'] = total_words_group_without_stop_word
        dict_words[id_word]['words_total_in_group_with_stop_words'] = total_words_group_with_stop_word
    
    print('Finish update dict of words')

    return dict_words

def create_dictionary_words(dict_words, dict_notice):
    print('Starting create dict of words ...')
    id_dict_words = 0

    for id_notice, notice_info in dict_notice.items():
        notice_words = notice_info['notice_content_stemm_without_stopwords']
        for word in notice_words:
            dict_words, id_dict_words = add_words_on_dict(dict_words, id_dict_words, word)
    
    dict_words_sorted = dict(sorted(dict_words.items(), key=lambda x: x[1]['word']))
    print('Finish create dict of words')
    return dict_words_sorted

def save_dict_words_to_xlsx(file_path, dict_word, group_name):
    print('Starting save dict ...')

    df = pd.DataFrame.from_dict(dict_word, orient='index')

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df = df.sort_values(by='word')
        df.reset_index(drop=True, inplace=True)
        df.to_excel(writer, sheet_name=f'Dicionario de Palavras {group_name}', index=False)

        worksheet = writer.sheets[f'Dicionario de Palavras {group_name}']

        (max_row, max_col) = df.shape

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value)

        for row_num, row_data in enumerate(df.itertuples(index=False)):
            for col_num, value in enumerate(row_data):
                if isinstance(value, list):
                    df.at[row_num, df.columns[col_num]] = '\n'.join(map(str, value))
                    value = '\n'.join(map(str, value))
                worksheet.write(row_num, col_num, value)

        last_col = len(df.columns)
        df.insert(last_col, 'info_dict', '')

        info_dict_list = [
            f"- Dicionario de Palavras das Noticias Classificadas como {group_name}",
            f"  Esse dicionario possui {len(dict_word)} palavras",
            "  Estrutura do dicionario:",
            "  ID - WORD - NUMBER NOTICE'S ON WORD APPEAR - ID'S NOTICES ON WORD APPEAR - NOTICE'S TITLE'S ON WORD APPEAR - CLASSIFICATION NOTICE'S ON WORD APPEAR -  NUMBER ON WORD APPEAR ON NOTICE - NUMBER WORDS ON NOTICE WITH STOPWORDS - NUMBER WORDS ON NOTICE WITHOUT STOPWORDS - NUMBER ON WORD APPEAR IN GROUP - NUMBER WORDS ON GROUP WITH STOPWORDS - NUMBER WORDS ON GROUP WITHOUT STOPWORDS"
        ]

        for i in range(len(info_dict_list)):
            df.at[i, 'info_dict'] = info_dict_list[i]

        df.to_excel(writer, sheet_name=f'Dicionario de Palavras {group_name}', index=False)

    print('Finish save dict ')


path_bd = os.path.join(project_root, 'base_data', 'FakeRecogna.xlsx')

dict_notice_real: dict = {}
words_total_with_stopwords_group_real = 0
words_total_without_stopwords_group_real = 0

dict_notice_real, words_total_with_stopwords_group_real, words_total_without_stopwords_group_real = create_dictionary_notices(path_bd, 1)

dict_words_real: dict = {}
words_dict_real_total = 0

dict_words_real = create_dictionary_words(dict_words_real, dict_notice_real)
dict_words_real = update_dictionary_words(dict_words_real, dict_notice_real, words_total_with_stopwords_group_real, words_total_without_stopwords_group_real)

words_dict_real_total = len(dict_words_real)

outpat_notice = os.path.join(project_root, 'output', 'dict_notice_real.xlsx')
save_dict_notice_to_xlsx(outpat_notice, dict_notice_real, 'Reais', words_total_with_stopwords_group_real, words_total_without_stopwords_group_real)

outpat_words = os.path.join(project_root, 'output', 'dict_words_real.xlsx')
save_dict_words_to_xlsx(outpat_words, dict_words_real, 'Reais')

dict_notice_fake: dict = {}
words_total_with_stopwords_group_fake = 0
words_total_without_stopwords_group_fake = 0

dict_notice_fake, words_total_with_stopwords_group_fake, words_total_without_stopwords_group_fake = create_dictionary_notices(path_bd, 0)

dict_words_fake: dict = {}
words_dict_real_fake = 0

dict_words_fake = create_dictionary_words(dict_words_fake, dict_notice_fake)
dict_words_fake = update_dictionary_words(dict_words_fake, dict_notice_fake, words_total_with_stopwords_group_fake, words_total_without_stopwords_group_fake)

words_dict_fake_total = len(dict_words_fake)

outpat_notice = os.path.join(project_root, 'output', 'dict_notice_fake.xlsx')
save_dict_notice_to_xlsx(outpat_notice, dict_notice_fake, 'Fakes', words_total_with_stopwords_group_fake, words_total_without_stopwords_group_fake)

outpat_words = os.path.join(project_root, 'output', 'dict_words_fake.xlsx')
save_dict_words_to_xlsx(outpat_words, dict_words_fake, 'Fakes')
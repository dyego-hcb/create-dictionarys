import sys
import os

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
    dict_notice[id_dict_notice]['notice_all_words'] = words_stemmed
    dict_notice[id_dict_notice]['classe_notice'] = classe_notice
    dict_notice[id_dict_notice]['notice_words_total_with_stopwords'] = len(words_lower_case)
    dict_notice[id_dict_notice]['notice_words_total_without_stopwords'] = len(words_without_stopwords)

    return  dict_notice

def create_dictionary_notices(path_bd):
    data_list = extract_data(path_bd)

    print('Starting create dict of notices ...')
    dict_notice_real: dict = {}
    dict_notice_fake: dict = {}
    id_dict_real = 0
    words_total_with_stopwords_group_real = 0
    words_total_without_stopwords_group_real = 0
    id_dict_fake = 0
    words_total_with_stopwords_group_fake = 0
    words_total_without_stopwords_group_fake = 0

    for data in data_list[1:]:
        id_dict = data[0]
        title = data[1]
        content = data[2]
        classe = data[3]

        if classe == 1:
            current_dict = dict_notice_real
            id_dict_real += 1
            id_dict_group = id_dict_real
            words_total_with_stopwords_group_real += len(content.split())
            words_total_without_stopwords_group_real += len(remove_stopwords_in_list(content.split()))
        elif classe == 0:
            current_dict = dict_notice_fake
            id_dict_fake += 1
            id_dict_group = id_dict_fake
            words_total_with_stopwords_group_fake += len(content.split())
            words_total_without_stopwords_group_fake += len(remove_stopwords_in_list(content.split()))

        current_dict = add_notice_on_dict(
            current_dict, id_dict_group, id_dict, title, content, classe
        )

    print('Finish create dict of notices')

    return dict_notice_real, words_total_with_stopwords_group_real, words_total_without_stopwords_group_real, dict_notice_fake, words_total_with_stopwords_group_fake, words_total_without_stopwords_group_fake

# def save_dict_to_txt(file_path, dict_notice, group_name, words_with_stopwords, words_without_stopwords):
#     with open(file_path, 'w', encoding='utf-8') as file:
#         file.write(f"- Dicionario de Noticias {group_name}\n")
#         file.write(f"  Esse dicionario possui {words_without_stopwords} palavras (STOP-WORDS REMOVIDAS)\n")
#         file.write(f"  Esse dicionario possui {words_with_stopwords} palavras (STOP-WORDS NO TEXTO)\n")
#         file.write("  Estrutura do dicionario: \n")
#         file.write("  ID - ID NOTICE - TITLE NOTICE - NOTICE ALL WORDS - CLASSE NOTICE - NOTICE WORDS TOTAL WITH STOPWORDS -  NOTICE WORDS TOTAL WITHOUT STOPWORDS \n")        
#         file.write("- Dados do Dicionario:\n")
#         for key, value in dict_notice.items():
#             file.write(str(key) + ' - ' + str(value) + '\n')

# dict_notice_real: dict = {}
# dict_notice_fake: dict = {}
# words_total_with_stopwords_group_real = 0
# words_total_without_stopwords_group_real = 0
# words_total_with_stopwords_group_fake = 0
# words_total_without_stopwords_group_fake = 0

# path_bd = os.path.join(project_root, 'base_data', 'FakeRecogna.xlsx')
# dict_notice_real, words_total_with_stopwords_group_real, words_total_without_stopwords_group_real, dict_notice_fake, words_total_with_stopwords_group_fake, words_total_without_stopwords_group_fake = create_dictionary_notices(path_bd)
# outpat = os.path.join(project_root, 'output', 'dict_notice_real.txt')
# save_dict_to_txt(outpat, dict_notice_real, 'REAL', words_total_with_stopwords_group_real, words_total_without_stopwords_group_real)
# outpat = os.path.join(project_root, 'output', 'dict_notice_fake.txt')
# save_dict_to_txt(outpat, dict_notice_fake, 'FAKE', words_total_with_stopwords_group_real, words_total_without_stopwords_group_real)
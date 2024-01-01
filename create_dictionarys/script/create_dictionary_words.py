import sys
import os

from create_dictionary_notices import create_dictionary_notices

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
            notice_words = notice_info['notice_all_words']
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
        notice_words = notice_info['notice_all_words']
        for word in notice_words:
            dict_words, id_dict_words = add_words_on_dict(dict_words, id_dict_words, word)
    
    dict_words_sorted = dict(sorted(dict_words.items(), key=lambda x: x[1]['word']))
    print('Finish create dict of words')
    return dict_words_sorted

def save_dict_to_txt(file_path, dict_words, group_name, words_dict):
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(f"- Dicionario de Noticias {group_name}\n")
        file.write(f"  Esse dicionario possui {words_dict} palavras\n")
        file.write("  Estrutura do dicionario: \n")
        file.write("  ID - WORD - NOTICES APPEAR TOTAL - (ID NOTICE WORD APPEAR - TITLE NOTICE WORD APPEAR - CLASS NOTICE WORD APPEAR ; WORDS TOTAL APPEAR IN NOTICE (WORDS TOTAL IN NOTICE WITHOUT STOP WORDS)[WORDS TOTAL IN NOTICE WIT STOP WORDS] - NUM WORD APPEAR IN GROUP/(WORDS TOTAL GROUP WITHOUT STOPWORDS)[WORDS TOTAL GROUP WITH STOPWORDS])\n")        
        file.write("- Dados do Dicionario:\n")
        for key, value in dict_words.items():
            file.write(str(key) + ' - ' 
                       + value['word'] + ' - ' 
                       + str(value['notices_appear_total']) + ' - ( ') 
            for i in range(len(value['ids_notice_appear'])):
                file.write(str(value['ids_notice_appear'][i]) + ' - ' 
                + str(value['titles_notice_appear'][i]) + ' - ' 
                + str(value['classe_notice_word_appear'][i]) + ' ; ' 
                + str(value['words_total_appear_in_notice'][i]) + ' / ( ' 
                + str(value['words_total_in_notice_without_stop_words'][i]) + ' )[ ' 
                + str(value['words_total_in_notice_with_stop_words'][i]) + ' ] - ') 
            
            file.write(str(value['words_total_appear_in_group']) + ' / ( ' 
                + str(value['words_total_in_group_without_stop_words']) + ' )[' 
                + str(value['words_total_in_group_with_stop_words']) + '])\n')


path_bd = os.path.join(project_root, 'base_data', 'FakeRecogna.xlsx')

dict_notice_real: dict = {}
dict_notice_fake: dict = {}
words_total_with_stopwords_group_real = 0
words_total_without_stopwords_group_real = 0
words_total_with_stopwords_group_fake = 0
words_total_without_stopwords_group_fake = 0

dict_notice_real, words_total_with_stopwords_group_real, words_total_without_stopwords_group_real, dict_notice_fake, words_total_with_stopwords_group_fake, words_total_without_stopwords_group_fake = create_dictionary_notices(path_bd)

dict_words_real: dict = {}
dict_words_fake: dict = {}
words_dict_real_total = 0
words_dict_real_fake = 0

dict_words_real = create_dictionary_words(dict_words_real, dict_notice_real)
dict_words_real = update_dictionary_words(dict_words_real, dict_notice_real, words_total_with_stopwords_group_real, words_total_without_stopwords_group_real)

dict_words_fake = create_dictionary_words(dict_words_fake, dict_notice_fake)
dict_words_fake = update_dictionary_words(dict_words_fake, dict_notice_fake, words_total_with_stopwords_group_fake, words_total_without_stopwords_group_fake)

words_dict_real_total = len(dict_words_real)
words_dict_real_fake = len(dict_words_fake)

outpat = os.path.join(project_root, 'output', 'dict_words_real.txt')
save_dict_to_txt(outpat, dict_words_real, 'REAL', words_dict_real_total)
outpat = os.path.join(project_root, 'output', 'dict_words_fake.txt')
save_dict_to_txt(outpat, dict_words_fake, 'FAKE', words_dict_real_fake)
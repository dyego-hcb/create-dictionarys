import sys
import os
import pandas as pd

from create_dictionary_notices import create_dictionary_notices, create_dictionary_notices_relevant_info, update_dictionary_notices_relevant_info, load_dict_notices_xlsx, save_dict_notices_to_xlsx, save_dict_notices_to_csv, load_dict_notices_relevant_info_xlsx, save_dict_notices_relevant_info_to_xlsx, save_dict_notices_relevant_info_to_csv, create_dictionary_notices_adapter_to_weka, update_dictionary_notices_adapter_to_weka, remove_notices_not_appear_strong_words, load_dict_notices_adapter_to_weka_xlsx, save_dict_notices_adapter_to_weka_to_xlsx, save_dict_notices_adapter_to_weka_to_csv
from create_dictionary_strong_words import load_dict_strong_wrods_xlsx, save_dict_strong_words_to_xlsx
from create_dictionary_words_group import create_dictionary_words_group, update_dictionary_words_group, load_dict_words_group_xlsx, save_dict_words_group_to_xlsx, save_dict_words_group_relevants_info_to_csv
from create_dictionary_words import create_dictionary_words, update_dictionary_words, calculate_percent_to_strong_word, load_dict_words_xlsx, save_dict_words_to_xlsx, save_dict_words_relevants_info_to_csv

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..'))
sys.path.append(project_root)

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(current_dir, '..'))
    sys.path.append(project_root)

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
        path_load_dicts, "dict_strong_words_reais_with_100_words.xlsx", dict_strong_words_reais)

    # dict_strong_words_reais = create_dictionary_strong_words(dict_strong_words_reais, dict_words, 1)

    strong_words_dict_reais_total = len(dict_strong_words_reais)

    # outpat_strong_words = os.path.join(project_root, 'output', 'dict_strong_words_reais.xlsx')
    # save_dict_strong_words_to_xlsx(outpat_strong_words, dict_strong_words_reais, 'R')

    # outpat_strong_words = os.path.join(project_root, 'output', 'dict_strong_words_reais_relevant_info.csv')
    # save_dict_strong_words_to_csv(outpat_strong_words, dict_strong_words_reais, 'Reais')

    dict_strong_words_fakes: dict = {}
    # strong_words_dict_fakes_total = 0

    dict_strong_words_fakes = load_dict_strong_wrods_xlsx(
        path_load_dicts, "dict_strong_words_fakes_with_100_words.xlsx", dict_strong_words_fakes)

    # dict_strong_words_fakes = create_dictionary_strong_words(dict_strong_words_fakes, dict_words, 0)
    
    strong_words_dict_fakes_total = len(dict_strong_words_fakes)

    # outpat_strong_words = os.path.join(project_root, 'output', 'dict_strong_words_fakes.xlsx')
    # save_dict_strong_words_to_xlsx(outpat_strong_words, dict_strong_words_fakes, 'F')

    # outpat_strong_words = os.path.join(project_root, 'output', 'dict_strong_words_fakes_relevant_info.csv')
    # save_dict_strong_words_to_csv(outpat_strong_words, dict_strong_words_fakes, 'Fakes')

    # dict_notice_boths_relevant_info: dict = {}

    # dict_notice_boths_relevant_info = create_dictionary_notices_relevant_info(
    #     dict_notice_boths_relevant_info, dict_notice_reais)
    # dict_notice_boths_relevant_info = create_dictionary_notices_relevant_info(
    #     dict_notice_boths_relevant_info, dict_notice_fakes)

    # dict_notice_boths_relevant_info = update_dictionary_notices_relevant_info(
    #     dict_notice_boths_relevant_info, dict_notice_reais, dict_strong_words_reais, 1)
    # dict_notice_boths_relevant_info = update_dictionary_notices_relevant_info(
    #     dict_notice_boths_relevant_info, dict_notice_reais, dict_strong_words_fakes, 0)

    # dict_notice_boths_relevant_info = update_dictionary_notices_relevant_info(
    #     dict_notice_boths_relevant_info, dict_notice_fakes, dict_strong_words_reais, 1)
    # dict_notice_boths_relevant_info = update_dictionary_notices_relevant_info(
    #     dict_notice_boths_relevant_info, dict_notice_fakes, dict_strong_words_fakes, 0)

    # outpat_notice_reais_relevant_info_xlsx = os.path.join(
    #     project_root, 'output', 'dict_notice_relevant_info_boths.xlsx')
    # save_dict_notices_relevant_info_to_xlsx(
    #     outpat_notice_reais_relevant_info_xlsx, dict_notice_boths_relevant_info, 'Boths')

    # outpat_notice_reais_relevant_info_csv = os.path.join(
    #     project_root, 'output', 'dict_notice_relevant_info_boths.csv')
    # save_dict_notices_relevant_info_to_csv(
    #     outpat_notice_reais_relevant_info_csv, dict_notice_boths_relevant_info, 'Boths')

    # dict_notice_reais_relevant_info: dict = {}
    # dict_notice_reais_relevant_info = create_dictionary_notices_relevant_info(
    #     dict_notice_reais_relevant_info, dict_notice_reais)
    
    # dict_notice_reais_relevant_info = update_dictionary_notices_relevant_info(
    #     dict_notice_reais_relevant_info, dict_notice_reais, dict_strong_words_reais, 1)
    # dict_notice_reais_relevant_info = update_dictionary_notices_relevant_info(
    #     dict_notice_reais_relevant_info, dict_notice_reais, dict_strong_words_fakes, 0)

    # outpat_notice_reais_relevant_info_xlsx = os.path.join(
    #     project_root, 'output', 'dict_notice_relevant_info_reais.xlsx')
    # save_dict_notices_relevant_info_to_xlsx(
    #     outpat_notice_reais_relevant_info_xlsx, dict_notice_reais_relevant_info, 'Reais')

    # outpat_notice_reais_relevant_info_csv = os.path.join(
    #     project_root, 'output', 'dict_notice_relevant_info_reais.csv')
    # save_dict_notices_relevant_info_to_csv(
    #     outpat_notice_reais_relevant_info_csv, dict_notice_reais_relevant_info, 'Reais')

    # dict_notice_fakes_relevant_info: dict = {}
    # dict_notice_fakes_relevant_info = create_dictionary_notices_relevant_info(
    #     dict_notice_fakes_relevant_info, dict_notice_fakes)
    
    # dict_notice_fakes_relevant_info = update_dictionary_notices_relevant_info(
    #     dict_notice_fakes_relevant_info, dict_notice_fakes, dict_strong_words_reais, 1)
    # dict_notice_fakes_relevant_info = update_dictionary_notices_relevant_info(
    #     dict_notice_fakes_relevant_info, dict_notice_fakes, dict_strong_words_fakes, 0)

    # outpat_notice_fakes_relevant_info_xlsx = os.path.join(
    #     project_root, 'output', 'dict_notice_relevant_info_fakes.xlsx')
    # save_dict_notices_relevant_info_to_xlsx(
    #     outpat_notice_fakes_relevant_info_xlsx, dict_notice_fakes_relevant_info, 'Fakes')

    # outpat_notice_fakes_relevant_info_csv = os.path.join(
    #     project_root, 'output', 'dict_notice_relevant_info_fakes.csv')
    # save_dict_notices_relevant_info_to_xlsx(
    #     outpat_notice_fakes_relevant_info_csv, dict_notice_fakes_relevant_info, 'Fakes')
    
    strong_words_boths_total = strong_words_dict_fakes_total + strong_words_dict_reais_total

    dict_notice_boths_adapter_to_weka: dict = {}

    dict_notice_boths_adapter_to_weka = create_dictionary_notices_adapter_to_weka(
        dict_notice_boths_adapter_to_weka, dict_notice_reais, dict_strong_words_reais, 0)
    dict_notice_boths_adapter_to_weka = create_dictionary_notices_adapter_to_weka(
        dict_notice_boths_adapter_to_weka, dict_notice_fakes, dict_strong_words_fakes, 0)

    dict_notice_boths_adapter_to_weka = update_dictionary_notices_adapter_to_weka(
        dict_notice_boths_adapter_to_weka, dict_notice_reais, dict_strong_words_reais)
    dict_notice_boths_adapter_to_weka = update_dictionary_notices_adapter_to_weka(
        dict_notice_boths_adapter_to_weka, dict_notice_reais, dict_strong_words_fakes)

    dict_notice_boths_adapter_to_weka = update_dictionary_notices_adapter_to_weka(
        dict_notice_boths_adapter_to_weka, dict_notice_fakes, dict_strong_words_reais)
    dict_notice_boths_adapter_to_weka = update_dictionary_notices_adapter_to_weka(
        dict_notice_boths_adapter_to_weka, dict_notice_fakes, dict_strong_words_fakes)
    
    dict_notice_boths_adapter_to_weka = remove_notices_not_appear_strong_words(dict_notice_boths_adapter_to_weka, strong_words_boths_total)

    # outpat_notice_reais_adapter_to_weka_xlsx = os.path.join(
    #     project_root, 'output', 'dict_notice_adapter_to_weka_boths.xlsx')
    # save_dict_notices_adapter_to_weka_to_xlsx(
    #     outpat_notice_reais_adapter_to_weka_xlsx, dict_notice_boths_adapter_to_weka, 'Boths')

    outpat_notice_reais_adapter_to_weka_csv = os.path.join(
        project_root, 'output', 'dict_notice_adapter_to_weka_boths.csv')
    save_dict_notices_adapter_to_weka_to_csv(
        outpat_notice_reais_adapter_to_weka_csv, dict_notice_boths_adapter_to_weka, 'Boths')

    dict_notice_reais_adapter_to_weka: dict = {}
    dict_notice_reais_adapter_to_weka = create_dictionary_notices_adapter_to_weka(
        dict_notice_reais_adapter_to_weka, dict_notice_reais, dict_strong_words_reais, 0)
    dict_notice_reais_adapter_to_weka = create_dictionary_notices_adapter_to_weka(
        dict_notice_reais_adapter_to_weka, dict_notice_reais, dict_strong_words_fakes, 1)
    
    dict_notice_reais_adapter_to_weka = update_dictionary_notices_adapter_to_weka(
        dict_notice_reais_adapter_to_weka, dict_notice_reais, dict_strong_words_reais)
    dict_notice_reais_adapter_to_weka = update_dictionary_notices_adapter_to_weka(
        dict_notice_reais_adapter_to_weka, dict_notice_reais, dict_strong_words_fakes)
    
    dict_notice_reais_adapter_to_weka = remove_notices_not_appear_strong_words(dict_notice_reais_adapter_to_weka, strong_words_boths_total)

    # outpat_notice_reais_adapter_to_weka_xlsx = os.path.join(
    #     project_root, 'output', 'dict_notice_adapter_to_weka_reais.xlsx')
    # save_dict_notices_adapter_to_weka_to_xlsx(
    #     outpat_notice_reais_adapter_to_weka_xlsx, dict_notice_reais_adapter_to_weka, 'Reais')

    outpat_notice_reais_adapter_to_weka_csv = os.path.join(
        project_root, 'output', 'dict_notice_adapter_to_weka_reais.csv')
    save_dict_notices_adapter_to_weka_to_csv(
        outpat_notice_reais_adapter_to_weka_csv, dict_notice_reais_adapter_to_weka, 'Reais')

    dict_notice_fakes_adapter_to_weka: dict = {}
    dict_notice_fakes_adapter_to_weka = create_dictionary_notices_adapter_to_weka(
        dict_notice_fakes_adapter_to_weka, dict_notice_fakes, dict_strong_words_fakes, 0)
    dict_notice_fakes_adapter_to_weka = create_dictionary_notices_adapter_to_weka(
        dict_notice_fakes_adapter_to_weka, dict_notice_fakes, dict_strong_words_reais, 1)
    
    dict_notice_fakes_adapter_to_weka = update_dictionary_notices_adapter_to_weka(
        dict_notice_fakes_adapter_to_weka, dict_notice_fakes, dict_strong_words_reais)
    dict_notice_fakes_adapter_to_weka = update_dictionary_notices_adapter_to_weka(
        dict_notice_fakes_adapter_to_weka, dict_notice_fakes, dict_strong_words_fakes)
    
    dict_notice_fakes_adapter_to_weka = remove_notices_not_appear_strong_words(dict_notice_fakes_adapter_to_weka, strong_words_boths_total)

    # outpat_notice_fakes_adapter_to_weka_xlsx = os.path.join(
    #     project_root, 'output', 'dict_notice_adapter_to_weka_fakes.xlsx')
    # save_dict_notices_adapter_to_weka_to_xlsx(
    #     outpat_notice_fakes_adapter_to_weka_xlsx, dict_notice_fakes_adapter_to_weka, 'Fakes')

    outpat_notice_fakes_adapter_to_weka_csv = os.path.join(
        project_root, 'output', 'dict_notice_adapter_to_weka_fakes.csv')
    save_dict_notices_adapter_to_weka_to_csv(
        outpat_notice_fakes_adapter_to_weka_csv, dict_notice_fakes_adapter_to_weka, 'Fakes')

    print("\nFinish execute ORI methods\n")

if __name__ == "__main__":
    main()
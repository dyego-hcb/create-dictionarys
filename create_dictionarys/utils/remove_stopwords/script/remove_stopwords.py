import nltk
from nltk.corpus import stopwords

nltk.download('stopwords')
stop_words_pt = set(stopwords.words('portuguese'))

def remove_stopwords_in_list(word_list):
    try:
        filtered_words = [word for word in word_list if word.lower() not in stop_words_pt]
        return filtered_words

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None
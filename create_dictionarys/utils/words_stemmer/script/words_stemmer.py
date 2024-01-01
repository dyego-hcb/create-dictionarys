import nltk
from nltk.stem import PorterStemmer

nltk.download('punkt')

stemmer = PorterStemmer()

def stemmize_words(word_list):
    try:        
        palavras_stemizadas = [stemmer.stem(palavra) for palavra in word_list]

        return sorted(palavras_stemizadas)

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None
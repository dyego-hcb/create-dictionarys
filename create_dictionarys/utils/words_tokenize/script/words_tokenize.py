import nltk
from nltk.tokenize import word_tokenize

nltk.download('punkt')

def tokenize_words(text):
    try:
        tokens = word_tokenize(text)
        return tokens
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None

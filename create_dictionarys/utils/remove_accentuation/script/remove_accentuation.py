from unidecode import unidecode

def remove_accentuation_in_list(input_words):
    try:
        texto = ' '.join(input_words)
        texto_sem_acentuacoes = remover_accentuation(texto)
        palavras_sem_acentuacoes = texto_sem_acentuacoes.split()
        return palavras_sem_acentuacoes

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None

def remover_accentuation(texto):
    return unidecode(texto)

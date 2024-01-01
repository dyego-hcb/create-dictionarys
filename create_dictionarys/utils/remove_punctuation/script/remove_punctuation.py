import string

def remove_ponctuation_in_list(input_words):
    try:
        texto = ' '.join(input_words)

        texto_sem_pontuacao = remover_ponctuation(texto)

        palavras_sem_pontuacao = texto_sem_pontuacao.split()

        return palavras_sem_pontuacao

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None

def remover_ponctuation(texto):
    caracteres_especiais = ['``', "''", '`', '“', '”', '‘', '’', '-', '–', '!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '_', '+', '=', '{', '}', '[', ']', '|', '\\', ':', ';', '<', '>', ',', '.', '?', '/', '"']

    for char in caracteres_especiais:
        texto = texto.replace(char, '')

    texto_sem_pontuacao = ''.join(c if c.isalnum() or c.isspace() else ' ' for c in texto)

    return texto_sem_pontuacao

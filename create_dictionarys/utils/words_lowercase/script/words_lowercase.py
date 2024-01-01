def convert_words_lower_case(input_words):
    try:
        palavras_lower_case = [palavra.lower() for palavra in input_words]
        
        return palavras_lower_case

    except Exception as e:
        print(f"Ocorreu um erro não tratado: {e}")
        return None

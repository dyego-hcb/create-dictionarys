# Criador de Dicionários - Processamento de Notícias

Bem-vindo ao repositório de Criador de Dicionários - Processamento de Notícias! Este projeto é dedicado ao processamento de dados de notícias em língua portuguesa, incluindo a criação de dicionários informativos e análises linguísticas. Abaixo, fornecerei uma descrição detalhada dos principais scripts e suas funcionalidades.

## Estrutura do Projeto:

- A pasta **[Código Fonte de Redução de Palavras](./create_dictionarys/)** contém os inputs e outputs utilizados no projeto, assim como o próprio código-fonte.
- Na pasta **[Script](./create_dictionarys/script/)**, você encontrará o código-fonte do projeto.

## Arquivos do Projeto:

1. **`create_dictionary_notices.py`:**
   Este script é responsável por criar dicionários de notícias. Ele utiliza uma variedade de módulos, incluindo tokenização, remoção de acentuações, pontuações e stopwords, além de realizar stemming nas palavras. O dicionário resultante é organizado por classes (real e fake), contendo informações essenciais sobre cada notícia.

   - **Funções Principais:**
     - `add_notice_on_dict`: Adiciona informações de uma notícia ao dicionário.
     - `create_dictionary_notices`: Cria dicionários de notícias com base em um arquivo Excel.

   - **Instruções de Uso:**
     - Execute o script fornecendo o caminho do arquivo Excel como argumento.

2. **`create_dictionary_words.py`:**
   Este script cria dicionários de palavras com base nos dicionários de notícias gerados anteriormente. Ele contabiliza a ocorrência de palavras em diferentes contextos, incluindo sua presença em notícias específicas e no grupo como um todo.

   - **Funções Principais:**
     - `add_words_on_dict`: Adiciona informações de uma palavra ao dicionário.
     - `update_dictionary_words`: Atualiza o dicionário de palavras com base nas notícias.
     - `create_dictionary_words`: Cria dicionários de palavras a partir dos dicionários de notícias.

   - **Instruções de Uso:**
     - Execute o script após a execução bem-sucedida do `create_dictionary_notices.py`.

3. **`extract_info_notices.py`:**
   Este script extrai dados relevantes de um arquivo Excel contendo informações sobre notícias. Ele lida com exceções comuns e fornece uma estrutura de dados limpa para processamento posterior.

   - **Funções Principais:**
     - `extract_data`: Extrai dados do arquivo Excel.

   - **Instruções de Uso:**
     - Forneça o caminho do arquivo Excel como argumento.

4. **Módulos Auxiliares:**
   - Vários módulos, como `remove_accentuation`, `remove_punctuation`, `remove_stopwords`, `words_lowercase`, `words_stemmer`, e `words_tokenize`, são utilizados para realizar operações específicas durante o pré-processamento das notícias e palavras.

## Como Usar:

1. **Configuração:**
   - Certifique-se de ter as dependências instaladas, conforme indicado no script `create_dictionary_notices.py`.

2. **Execução dos Scripts:**
   - Execute `create_dictionary_notices.py` antes de executar `create_dictionary_words.py`.

3. **Saída:**
   - Os resultados são salvos no diretório `output` com arquivos de texto detalhando os dicionários de notícias e palavras.

## Notas Importantes:

- Certifique-se de ter o arquivo Excel de notícias no diretório `base_data`.
- Consulte os comentários e docstrings nos scripts para obter mais informações sobre cada funcionalidade.
- Se desejar, descomente a seção no final de `create_dictionary_notices.py` para salvar os dicionários em arquivos de texto.

Espero que este guia detalhado facilite a compreensão e utilização deste sistema de processamento de notícias. Sinta-se à vontade para entrar em contato em caso de dúvidas ou sugestões. Boa codificação!

***

# Dictionary Creator - News Processing

Welcome to the Dictionary Creator - News Processing repository! This project is dedicated to processing news data in the Portuguese language, including the creation of informative dictionaries and linguistic analysis. Below, I will provide a detailed description of the main scripts and their functionalities.

## Project Structure:

- The **[Dictionary Creator Source Code](./create_dictionaries/)** folder contains the inputs and outputs used in the project, as well as the source code itself.
- In the **[Script](./create_dictionaries/script/)** folder, you will find the source code for the project.

## Project Files:

1. **`create_dictionary_notices.py`:**
   This script is responsible for creating news dictionaries. It uses various modules, including tokenization, accent removal, punctuation and stopword removal, and word stemming. The resulting dictionary is organized by classes (real and fake), containing essential information about each news item.

   - **Main Functions:**
     - `add_notice_on_dict`: Adds information about a news item to the dictionary.
     - `create_dictionary_notices`: Creates news dictionaries based on an Excel file.

   - **Usage Instructions:**
     - Execute the script providing the path to the Excel file as an argument.

2. **`create_dictionary_words.py`:**
   This script creates word dictionaries based on the previously generated news dictionaries. It counts the occurrence of words in different contexts, including their presence in specific news and the group as a whole.

   - **Main Functions:**
     - `add_words_on_dict`: Adds information about a word to the dictionary.
     - `update_dictionary_words`: Updates the word dictionary based on the news.
     - `create_dictionary_words`: Creates word dictionaries from news dictionaries.

   - **Usage Instructions:**
     - Execute the script after the successful execution of `create_dictionary_notices.py`.

3. **`extract_info_notices.py`:**
   This script extracts relevant data from an Excel file containing information about news. It handles common exceptions and provides a clean data structure for further processing.

   - **Main Functions:**
     - `extract_data`: Extracts data from the Excel file.

   - **Usage Instructions:**
     - Provide the path to the Excel file as an argument.

4. **Auxiliary Modules:**
   - Various modules, such as `remove_accentuation`, `remove_punctuation`, `remove_stopwords`, `words_lowercase`, `words_stemmer`, and `words_tokenize`, are used to perform specific operations during the preprocessing of news and words.

## How to Use:

1. **Setup:**
   - Ensure you have the dependencies installed, as indicated in the `create_dictionary_notices.py` script.

2. **Running the Scripts:**
   - Execute `create_dictionary_notices.py` before running `create_dictionary_words.py`.

3. **Output:**
   - The results are saved in the `output` directory with text files detailing the news and word dictionaries.

## Important Notes:

- Ensure you have the news Excel file in the `base_data` directory.
- Refer to the comments and docstrings in the scripts for more information about each functionality.
- If desired, uncomment the section at the end of `create_dictionary_notices.py` to save the dictionaries to text files.

I hope this detailed guide facilitates understanding and using this news processing system. Feel free to reach out in case of questions or suggestions. Happy coding!

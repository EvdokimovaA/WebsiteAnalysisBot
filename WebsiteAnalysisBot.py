import os
import time

import pandas
import pandas as pd
import requests
from bs4 import BeautifulSoup
import telebot

API_TOKEN = os.getenv('API_TOKEN')
bot = telebot.TeleBot(API_TOKEN)

API_KEY = None
FOLDER_ID = None
request_link = 'https://yandex.ru/search/xml'

request_formulation_file_path = 'request_formulation.xlsx'
domain_list_file_path = 'domain_list.xlsx'
output_file_path = 'results.xlsx'


# Функция для нормализации доменного имени
def normalize_domain(domain):
    return domain[len('www.'):].strip() if domain.startswith('www.') else domain


# Функция чтения Excel файла
def read_excel(file_path):
    with pd.ExcelFile(file_path) as xls:
        data = pd.read_excel(xls, header=None, nrows=10)
    rows_list = data.apply(lambda row: ' '.join(row.astype(str)), axis=1).tolist()
    return rows_list


# Функция обработки XML ответа3цй
def parsing_xml_response(xml_code, domain_list, request_text):
    soup = BeautifulSoup(xml_code, 'lxml-xml')
    domain_tags = soup.find_all('domain')
    domains = [normalize_domain(domain_tag.text) for domain_tag in domain_tags]

    domain_indices = {domain: [] for domain in domain_list}

    for index, domain in enumerate(domains):
        if domain in domain_indices:
            domain_indices[domain].append(index + 1)

    domain_names = list(domain_indices.keys())
    domain_values = []

    for domain in domain_names:
        if domain_indices[domain]:
            domain_values.append(' '.join(map(str, domain_indices[domain])))
        else:
            domain_values.append("отсутствует")

    domain_names.insert(0, request_text)
    domain_values.insert(0, "позиция:")

    df = pd.DataFrame([domain_names, domain_values])

    if not os.path.exists(output_file_path):
        df.to_excel(output_file_path, header=False, index=False)
    else:
        with pandas.ExcelWriter(output_file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, header=False, index=False, startrow=writer.sheets['Sheet1'].max_row)


# Функция для выполнения запроса к Яндекс API
def yandex_search_api_req(rows_list, domain_list):
    for request_text in rows_list:
        params = {
            'folderid': FOLDER_ID,
            'apikey': API_KEY,
            'query': request_text
        }
        response = requests.get(request_link, params=params)

        if response.status_code == 200:
            parsing_xml_response(response.text, domain_list, request_text)
        else:
            print(f"Ошибка при выполнении запроса '{request_text}': {response.status_code}")
        time.sleep(1)


@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Пожалуйста, отправьте ваш API_KEY.")


@bot.message_handler(func=lambda message: API_KEY is None)
def get_api_key(message):
    global API_KEY
    API_KEY = message.text.strip()
    bot.send_message(message.chat.id, "API_KEY принят. Пожалуйста, введите ваш FOLDER_ID")


@bot.message_handler(func=lambda message: FOLDER_ID is None)
def get_folder_id(message):
    global API_KEY, FOLDER_ID
    if API_KEY is None:
        bot.send_message(message.chat.id, "Сначала отправьте ваш API_KEY.")
        return
    FOLDER_ID = message.text.strip()
    bot.send_message(
        message.chat.id,
        "FOLDER_ID принят. Пожалуйста, загрузите первый файл со списком поисковых запросов (.xlsx)",
    )


@bot.message_handler(content_types=['document'])
def handle_file(message):
    global API_KEY, FOLDER_ID
    if API_KEY is None:
        bot.send_message(message.chat.id, "Сначала отправьте ваш API_KEY.")
        return
    if FOLDER_ID is None:
        bot.send_message(message.chat.id, "Сначала отправьте ваш FOLDER_ID.")
        return

    if 'request_formulation.xlsx' not in os.listdir():

        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        with open('request_formulation.xlsx', 'wb') as new_file:
            new_file.write(downloaded_file)
        bot.send_message(
            message.chat.id,
            "Первый файл загружен. Пожалуйста, загрузите второй файл со списком доменных имён (.xlsx)",
        )
    elif 'domain_list.xlsx' not in os.listdir():
        bot.send_message(message.chat.id, "Второй файл загружен. Ожидайте, идёт анализ...")
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        with open('domain_list.xlsx', 'wb') as new_file:
            new_file.write(downloaded_file)

        requests_list = read_excel(request_formulation_file_path)
        domain_list = read_excel(domain_list_file_path)
        yandex_search_api_req(requests_list, domain_list)

        with open(output_file_path, 'rb') as result_file:
            bot.send_document(message.chat.id, result_file, caption="Результаты")

        os.remove(request_formulation_file_path)
        os.remove(domain_list_file_path)
        os.remove(output_file_path)
        API_KEY = None
        FOLDER_ID = None

        bot.send_message(message.chat.id, "Для повторной проверки введите команду /start")


if __name__ == "__main__":
    bot.polling(none_stop=True)

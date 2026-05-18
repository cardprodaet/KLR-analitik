#!/usr/bin/env python3
"""
Создаёт все нужные листы в Google Таблице.
Запускать ОДИН РАЗ перед первым запуском аналитики.
"""

import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID   = '1SOLbCBhGcnsrwiW9JSiw0wH3YMQZcVMWYh_sNovXQjI'
CREDENTIALS_FILE = 'credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

SHEETS = [
    'Настройки',
    'Продажи',
    'Заказы',
    'Остатки',
    'Финансы',
    'РК День',
    'РК Неделя',
    'РК 14 Дней',
    'РК Месяц',
    'НМ Отчёт',
    'Воронка День',
    'Воронка Неделя',
    'Воронка 14 Дней',
    'Воронка Месяц',
    'Цены',
    'Карточки',
    'Поставки',
]

def main():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    ss    = gspread.authorize(creds).open_by_key(SPREADSHEET_ID)

    existing = {ws.title for ws in ss.worksheets()}
    created  = 0

    for name in SHEETS:
        if name in existing:
            print(f'  уже есть: {name}')
        else:
            ss.add_worksheet(title=name, rows=1000, cols=50)
            print(f'  ✅ создан: {name}')
            created += 1

    # Настройки — добавляем заголовок только если A1 пустая (не трогаем B2 с API ключом)
    try:
        ws = ss.worksheet('Настройки')
        if not ws.acell('A1').value:
            ws.update(values=[['Параметр', 'Значение'], ['API ключ WB', '']], range_name='A1')
            print('  Настройки: добавлены заголовки (B2 не тронут)')
    except Exception as e:
        print(f'  Настройки: {e}')

    print(f'\nГотово! Создано листов: {created}, уже существовало: {len(SHEETS) - created}')
    print('Теперь вставь API ключ в лист Настройки → ячейка B2')

if __name__ == '__main__':
    main()

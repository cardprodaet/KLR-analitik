#!/usr/bin/env python3
"""
KLR Analytics — main1.py
Продажи, Заказы, Остатки, Финансовый отчёт
"""

from __future__ import annotations

import logging
import time
from datetime import datetime, timedelta

import gspread
import requests
from google.oauth2.service_account import Credentials

SPREADSHEET_ID   = '1SOLbCBhGcnsrwiW9JSiw0wH3YMQZcVMWYh_sNovXQjI'
CREDENTIALS_FILE = 'credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

STAT_BASE   = 'https://statistics-api.wildberries.ru'
WRITE_BATCH = 500
DAYS_BACK   = 30  # глубина выгрузки по умолчанию

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)-8s  %(message)s',
    datefmt='%H:%M:%S',
)
log = logging.getLogger(__name__)

# ── Google Sheets ──────────────────────────────────────────────────────────────

def get_spreadsheet() -> gspread.Spreadsheet:
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)

def get_api_key(ss: gspread.Spreadsheet) -> str:
    return ss.worksheet('Настройки').acell('B2').value.strip()

def set_status(ss: gspread.Spreadsheet, name: str, status: str) -> None:
    try:
        now = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
        ss.worksheet('Настройки').update(
            values=[['Последнее обновление:', name, now, status]], range_name='D2'
        )
        log.info('%s — %s', name, status)
    except Exception as exc:
        log.warning('set_status error: %s', exc)

def write_sheet(ss: gspread.Spreadsheet, name: str, rows: list[list]) -> None:
    ws = ss.worksheet(name)
    ws.clear()
    time.sleep(2)
    for i in range(0, len(rows), WRITE_BATCH):
        chunk = rows[i : i + WRITE_BATCH]
        if i == 0:
            ws.update(values=chunk, range_name='A1')
        else:
            ws.append_rows(chunk)
        time.sleep(2)
    log.info('%s → %d rows written', name, len(rows) - 1)

# ── HTTP ───────────────────────────────────────────────────────────────────────

def wb_request(method: str, url: str, api_key: str, max_retries: int = 5, **kwargs):
    headers = {'Authorization': api_key}
    if 'json' in kwargs:
        headers['Content-Type'] = 'application/json'
    for attempt in range(1, max_retries + 1):
        try:
            resp = getattr(requests, method)(url, headers=headers, timeout=60, **kwargs)
            if resp.status_code == 200:
                return resp
            if resp.status_code == 429:
                wait = 60 * attempt
                log.warning('429 rate limit (attempt %d/%d) — sleeping %ds', attempt, max_retries, wait)
                time.sleep(wait)
                continue
            log.error('HTTP %d: %s', resp.status_code, resp.text[:300])
            time.sleep(30)
        except requests.RequestException as exc:
            log.error('Request error (attempt %d/%d): %s', attempt, max_retries, exc)
            time.sleep(30)
    log.error('All retries exhausted for %s', url)
    return None

# ── Продажи ────────────────────────────────────────────────────────────────────

def load_sales(api_key: str, ss: gspread.Spreadsheet) -> None:
    set_status(ss, 'Продажи', '🔄 Загружается...')
    date_from = (datetime.now() - timedelta(days=DAYS_BACK)).strftime('%Y-%m-%dT00:00:00')
    url  = f'{STAT_BASE}/api/v1/supplier/sales?dateFrom={date_from}&flag=0'
    resp = wb_request('get', url, api_key)
    if not resp:
        set_status(ss, 'Продажи', '❌ Нет данных'); return

    data = resp.json()
    if not data:
        set_status(ss, 'Продажи', '❌ Нет данных'); return

    headers = [
        'Дата', 'Последнее изменение', 'Склад', 'Тип склада',
        'Страна', 'Округ', 'Регион',
        'Артикул продавца', 'Артикул WB', 'Баркод',
        'Категория', 'Предмет', 'Бренд', 'Размер',
        'Кол-во', 'Цена розничная', 'Скидка, %', 'СПП',
        'К перечислению', 'Итоговая цена', 'Цена со скидкой',
        'Отмена', 'Дата отмены', 'Тип заказа', 'Номер заказа', 'srid',
    ]
    rows = [headers]
    for s in data:
        rows.append([
            s.get('date', ''),            s.get('lastChangeDate', ''),
            s.get('warehouseName', ''),   s.get('warehouseType', ''),
            s.get('countryName', ''),     s.get('oblastOkrugName', ''),
            s.get('regionName', ''),      s.get('supplierArticle', ''),
            s.get('nmId', ''),            s.get('barcode', ''),
            s.get('category', ''),        s.get('subject', ''),
            s.get('brand', ''),           s.get('techSize', ''),
            s.get('quantity', 0),         s.get('totalPrice', 0),
            s.get('discountPercent', 0),  s.get('spp', 0),
            s.get('forPay', 0),           s.get('finishedPrice', 0),
            s.get('priceWithDisc', 0),    s.get('isCancel', False),
            s.get('cancelDate', ''),      s.get('orderType', ''),
            s.get('gNumber', ''),         s.get('srid', ''),
        ])
    write_sheet(ss, 'Продажи', rows)
    set_status(ss, 'Продажи', f'✅ Готово — {len(data)} записей')

# ── Заказы ─────────────────────────────────────────────────────────────────────

def load_orders(api_key: str, ss: gspread.Spreadsheet) -> None:
    set_status(ss, 'Заказы', '🔄 Загружается...')
    date_from = (datetime.now() - timedelta(days=DAYS_BACK)).strftime('%Y-%m-%dT00:00:00')
    url  = f'{STAT_BASE}/api/v1/supplier/orders?dateFrom={date_from}&flag=0'
    resp = wb_request('get', url, api_key)
    if not resp:
        set_status(ss, 'Заказы', '❌ Нет данных'); return

    data = resp.json()
    if not data:
        set_status(ss, 'Заказы', '❌ Нет данных'); return

    headers = [
        'Дата', 'Последнее изменение', 'Склад', 'Тип склада',
        'Страна', 'Округ', 'Регион',
        'Артикул продавца', 'Артикул WB', 'Баркод',
        'Категория', 'Предмет', 'Бренд', 'Размер',
        'Цена розничная', 'Скидка, %', 'СПП', 'Цена со скидкой',
        'Отмена', 'Дата отмены', 'Тип заказа', 'Номер заказа', 'srid',
    ]
    rows = [headers]
    for o in data:
        rows.append([
            o.get('date', ''),            o.get('lastChangeDate', ''),
            o.get('warehouseName', ''),   o.get('warehouseType', ''),
            o.get('countryName', ''),     o.get('oblastOkrugName', ''),
            o.get('regionName', ''),      o.get('supplierArticle', ''),
            o.get('nmId', ''),            o.get('barcode', ''),
            o.get('category', ''),        o.get('subject', ''),
            o.get('brand', ''),           o.get('techSize', ''),
            o.get('totalPrice', 0),       o.get('discountPercent', 0),
            o.get('spp', 0),              o.get('priceWithDisc', 0),
            o.get('isCancel', False),     o.get('cancelDate', ''),
            o.get('orderType', ''),       o.get('gNumber', ''),
            o.get('srid', ''),
        ])
    write_sheet(ss, 'Заказы', rows)
    set_status(ss, 'Заказы', f'✅ Готово — {len(data)} записей')

# ── Остатки ────────────────────────────────────────────────────────────────────

def load_stocks(api_key: str, ss: gspread.Spreadsheet) -> None:
    set_status(ss, 'Остатки', '🔄 Загружается...')
    date_from = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%dT00:00:00')
    url  = f'{STAT_BASE}/api/v1/supplier/stocks?dateFrom={date_from}'
    resp = wb_request('get', url, api_key)
    if not resp:
        set_status(ss, 'Остатки', '❌ Нет данных'); return

    data = resp.json()
    if not data:
        set_status(ss, 'Остатки', '❌ Нет данных'); return

    headers = [
        'Дата изменения', 'Склад',
        'Артикул продавца', 'Артикул WB', 'Баркод',
        'Категория', 'Предмет', 'Бренд', 'Размер',
        'Доступно', 'В пути к клиенту', 'В пути от клиента', 'Всего',
        'Цена', 'Скидка, %',
    ]
    rows = [headers]
    for s in data:
        rows.append([
            s.get('lastChangeDate', ''), s.get('warehouseName', ''),
            s.get('supplierArticle', ''), s.get('nmId', ''), s.get('barcode', ''),
            s.get('category', ''), s.get('subject', ''), s.get('brand', ''), s.get('techSize', ''),
            s.get('quantity', 0), s.get('inWayToClient', 0), s.get('inWayFromClient', 0),
            s.get('quantityFull', 0), s.get('Price', 0), s.get('Discount', 0),
        ])
    write_sheet(ss, 'Остатки', rows)
    set_status(ss, 'Остатки', f'✅ Готово — {len(data)} записей')

# ── Финансовый отчёт ───────────────────────────────────────────────────────────

def load_finances(api_key: str, ss: gspread.Spreadsheet) -> None:
    set_status(ss, 'Финансы', '🔄 Загружается...')
    date_from = (datetime.now() - timedelta(days=DAYS_BACK)).strftime('%Y-%m-%d')
    date_to   = datetime.now().strftime('%Y-%m-%d')

    all_rows: list[dict] = []
    rrdid = 0
    page  = 1

    while True:
        url  = (f'{STAT_BASE}/api/v5/supplier/reportDetailByPeriod'
                f'?dateFrom={date_from}&dateTo={date_to}&rrdid={rrdid}&limit=100000')
        resp = wb_request('get', url, api_key)
        if not resp:
            break
        data = resp.json()
        if not data:
            break
        all_rows.extend(data)
        log.info('Финансы страница %d: %d записей всего', page, len(all_rows))
        if len(data) < 100000:
            break
        rrdid = data[-1].get('rrd_id', 0)
        page += 1
        time.sleep(10)

    if not all_rows:
        set_status(ss, 'Финансы', '❌ Нет данных'); return

    headers = [
        'Номер отчёта', 'Период с', 'Период по', 'Дата заказа', 'Дата продажи',
        'Артикул WB', 'Артикул продавца', 'Предмет', 'Бренд', 'Баркод', 'Размер',
        'Тип документа', 'Операция', 'Офис', 'Склад WB',
        'Кол-во', 'Цена розн.', 'Сумма розн.',
        'Согл. скидка, %', 'Комиссия WB, %',
        'Цена с учётом скидки',
        'Доставок, шт', 'Возвратов, шт', 'Стоимость доставки',
        'Скидка пост. покупателя', 'Промо продавца',
        'К выплате продавцу',
        'Штраф', 'Доп. оплата',
        'Хранение', 'Удержания', 'Приёмка',
        'Стоимость логистики (пересчёт)', 'ID строки',
    ]
    rows = [headers]
    for r in all_rows:
        rows.append([
            r.get('realizationreport_id', ''), r.get('date_from', ''), r.get('date_to', ''),
            r.get('order_dt', ''), r.get('sale_dt', ''),
            r.get('nm_id', ''), r.get('sa_name', ''), r.get('subject_name', ''),
            r.get('brand_name', ''), r.get('barcode', ''), r.get('ts_name', ''),
            r.get('doc_type_name', ''), r.get('supplier_oper_name', ''),
            r.get('office_name', ''), r.get('ppvz_office_name', ''),
            r.get('quantity', 0), r.get('retail_price', 0), r.get('retail_sum', 0),
            r.get('sale_percent', 0), r.get('commission_percent', 0),
            r.get('retail_price_withdisc_rub', 0),
            r.get('delivery_amount', 0), r.get('return_amount', 0), r.get('delivery_rub', 0),
            r.get('product_discount_for_report', 0), r.get('supplier_promo', 0),
            r.get('ppvz_for_pay', 0),
            r.get('penalty', 0), r.get('additional_payment', 0),
            r.get('storage_fee', 0), r.get('deduction', 0), r.get('acceptance', 0),
            r.get('rebill_logistic_cost', 0), r.get('rrd_id', 0),
        ])
    write_sheet(ss, 'Финансы', rows)
    set_status(ss, 'Финансы', f'✅ Готово — {len(all_rows)} записей')

# ── Точка входа ────────────────────────────────────────────────────────────────

def main() -> None:
    log.info('=== main1 started ===')
    ss      = get_spreadsheet()
    api_key = get_api_key(ss)

    load_sales(api_key, ss);    time.sleep(5)
    load_orders(api_key, ss);   time.sleep(5)
    load_stocks(api_key, ss);   time.sleep(5)
    load_finances(api_key, ss)

    log.info('=== main1 complete ===')

if __name__ == '__main__':
    main()

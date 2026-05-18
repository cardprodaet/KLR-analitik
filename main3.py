#!/usr/bin/env python3
"""
KLR Analytics — main3.py
Воронка по периодам, Цены, Карточки товаров, Поставки
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

ANALYTICS_BASE = 'https://seller-analytics-api.wildberries.ru'
PRICES_BASE    = 'https://discounts-prices-api.wb.ru'
CONTENT_BASE   = 'https://suppliers-api.wildberries.ru'
MARKET_BASE    = 'https://marketplace-api.wildberries.ru'

PAGE_SLEEP    = 20
FUNNEL_SLEEP  = 90
WRITE_BATCH   = 500

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

# ── Воронка продаж ─────────────────────────────────────────────────────────────

def load_funnel_period(api_key: str, date_from: str, date_to: str,
                       ss: gspread.Spreadsheet, sheet_name: str) -> None:
    log.info('load_funnel_period [%s]: %s → %s', sheet_name, date_from, date_to)
    set_status(ss, sheet_name, '🔄 Загружается...')

    url = f'{ANALYTICS_BASE}/api/analytics/v3/sales-funnel/products'
    all_products: list[dict] = []
    offset, limit, page = 0, 1000, 1

    while True:
        body = {
            'selectedPeriod': {'start': date_from, 'end': date_to},
            'nmIds': [], 'brandNames': [], 'subjectIds': [], 'tagIds': [],
            'skipDeletedNm': False, 'limit': limit, 'offset': offset,
        }
        resp = wb_request('post', url, api_key, json=body)
        if not resp:
            break
        products = resp.json().get('data', {}).get('products', [])
        if not products:
            break
        all_products.extend(products)
        log.info('Products loaded: %d', len(all_products))
        if len(products) < limit:
            break
        offset += limit
        page   += 1
        time.sleep(PAGE_SLEEP)

    if not all_products:
        set_status(ss, sheet_name, '❌ Нет данных'); return

    headers = [
        'Артикул продавца', 'Артикул WB', 'Название', 'Предмет', 'Бренд',
        'Переходы в карточку', 'Переходы (пред.)',
        'В корзину, шт',      'В корзину (пред.)',
        'Конв. в корзину, %', 'Конв. в корзину (пред., %)',
        'В отложенные, шт',   'В отложенные (пред.)',
        'Заказали, шт',       'Заказали (пред.)',
        'Конв. в заказ, %',   'Конв. в заказ (пред., %)',
        'Выкупили, шт',       'Выкупили (пред.)',
        '% выкупа',           '% выкупа (пред.)',
        'Отменили, шт',       'Отменили (пред.)',
        'Заказали на сумму',  'Заказали сумма (пред.)',
        'Выкупили на сумму',  'Выкупили сумма (пред.)',
        'Средняя цена',       'Средняя цена (пред.)',
        'Остатки WB', 'Рейтинг товара', 'Рейтинг отзывов',
        'Время доставки, ч',  'Время доставки (пред.), ч',
    ]
    rows = [headers]
    for item in all_products:
        prod = item.get('product', {})
        s    = item.get('statistic', {}).get('selected', {})
        p    = item.get('statistic', {}).get('past',     {})
        sc   = s.get('conversions', {})
        pc   = p.get('conversions', {})
        st   = s.get('timeToReady', {})
        pt   = p.get('timeToReady', {})
        stk  = prod.get('stocks',   {})
        rows.append([
            prod.get('vendorCode', ''),    prod.get('nmId', ''),
            prod.get('title', ''),         prod.get('subjectName', ''),
            prod.get('brandName', ''),
            s.get('openCount', 0),         p.get('openCount', 0),
            s.get('cartCount', 0),         p.get('cartCount', 0),
            sc.get('addToCartPercent', 0), pc.get('addToCartPercent', 0),
            s.get('addToWishlist', 0),     p.get('addToWishlist', 0),
            s.get('orderCount', 0),        p.get('orderCount', 0),
            sc.get('cartToOrderPercent', 0), pc.get('cartToOrderPercent', 0),
            s.get('buyoutCount', 0),       p.get('buyoutCount', 0),
            sc.get('buyoutPercent', 0),    pc.get('buyoutPercent', 0),
            s.get('cancelCount', 0),       p.get('cancelCount', 0),
            s.get('orderSum', 0),          p.get('orderSum', 0),
            s.get('buyoutSum', 0),         p.get('buyoutSum', 0),
            s.get('avgPrice', 0),          p.get('avgPrice', 0),
            stk.get('wb', 0),              prod.get('productRating', 0),
            prod.get('feedbackRating', 0),
            st.get('days', 0) * 24 + st.get('hours', 0),
            pt.get('days', 0) * 24 + pt.get('hours', 0),
        ])
    write_sheet(ss, sheet_name, rows)
    set_status(ss, sheet_name, f'✅ Готово — {len(all_products)} товаров')

# ── Цены и скидки ──────────────────────────────────────────────────────────────

def load_prices(api_key: str, ss: gspread.Spreadsheet) -> None:
    set_status(ss, 'Цены', '🔄 Загружается...')
    url = f'{PRICES_BASE}/api/v2/list/goods/filter'
    all_goods: list[dict] = []
    limit, offset = 1000, 0

    while True:
        resp = wb_request('get', f'{url}?limit={limit}&offset={offset}', api_key)
        if not resp:
            break
        goods = resp.json().get('data', {}).get('listGoods', [])
        if not goods:
            break
        all_goods.extend(goods)
        log.info('Цены: загружено %d товаров', len(all_goods))
        if len(goods) < limit:
            break
        offset += limit
        time.sleep(5)

    if not all_goods:
        set_status(ss, 'Цены', '❌ Нет данных'); return

    headers = [
        'Артикул WB', 'Артикул продавца', 'Размер',
        'Цена до скидки, ₽', 'Цена со скидкой, ₽', 'Скидка продавца, %',
        'Клубная цена, ₽', 'Клубная скидка, %',
    ]
    rows = [headers]
    for good in all_goods:
        nm_id  = good.get('nmID', '')
        vendor = good.get('vendorCode', '')
        for size in good.get('sizes', []):
            rows.append([
                nm_id, vendor,
                size.get('techSizeName', ''),
                size.get('price', 0),
                size.get('discountedPrice', 0),
                size.get('discount', 0),
                size.get('clubPrice', 0),
                size.get('clubDiscount', 0),
            ])
    write_sheet(ss, 'Цены', rows)
    set_status(ss, 'Цены', f'✅ Готово — {len(all_goods)} товаров')

# ── Карточки товаров ───────────────────────────────────────────────────────────

def load_cards(api_key: str, ss: gspread.Spreadsheet) -> None:
    set_status(ss, 'Карточки', '🔄 Загружается...')
    url = f'{CONTENT_BASE}/content/v2/get/cards/list'
    all_cards: list[dict] = []
    cursor = {'limit': 100, 'updatedAt': '', 'nmID': 0}

    while True:
        body = {
            'settings': {
                'cursor': cursor,
                'filter': {'withPhoto': -1},
            }
        }
        resp = wb_request('post', url, api_key, json=body)
        if not resp:
            break
        data   = resp.json().get('data', {})
        cards  = data.get('cards', [])
        if not cards:
            break
        all_cards.extend(cards)
        log.info('Карточки: загружено %d', len(all_cards))
        new_cursor = data.get('cursor', {})
        if not new_cursor.get('updatedAt') or len(cards) < cursor['limit']:
            break
        cursor = {
            'limit':     cursor['limit'],
            'updatedAt': new_cursor.get('updatedAt', ''),
            'nmID':      new_cursor.get('nmID', 0),
        }
        time.sleep(3)

    if not all_cards:
        set_status(ss, 'Карточки', '❌ Нет данных'); return

    headers = [
        'Артикул WB', 'Артикул продавца', 'Бренд', 'Категория', 'Название',
        'Создана', 'Обновлена',
        'Размеры (ДxШxВ)', 'Вес, кг',
        'Фото, шт', 'Видео',
    ]
    rows = [headers]
    for card in all_cards:
        dim  = card.get('dimensions', {})
        rows.append([
            card.get('nmID', ''),          card.get('vendorCode', ''),
            card.get('brand', ''),         card.get('subjectName', ''),
            card.get('title', ''),
            card.get('createdAt', ''),     card.get('updatedAt', ''),
            f"{dim.get('length',0)}x{dim.get('width',0)}x{dim.get('height',0)}",
            dim.get('weight', 0),
            len(card.get('photos', [])),
            1 if card.get('video') else 0,
        ])
    write_sheet(ss, 'Карточки', rows)
    set_status(ss, 'Карточки', f'✅ Готово — {len(all_cards)} карточек')

# ── Поставки ───────────────────────────────────────────────────────────────────

def load_supplies(api_key: str, ss: gspread.Spreadsheet) -> None:
    set_status(ss, 'Поставки', '🔄 Загружается...')
    all_supplies: list[dict] = []
    next_id = 0

    while True:
        url  = f'{MARKET_BASE}/api/v3/supplies?limit=1000&next={next_id}'
        resp = wb_request('get', url, api_key)
        if not resp:
            break
        data     = resp.json()
        supplies = data.get('supplies', [])
        if not supplies:
            break
        all_supplies.extend(supplies)
        log.info('Поставки: загружено %d', len(all_supplies))
        next_id = data.get('next', 0)
        if not next_id or len(supplies) < 1000:
            break
        time.sleep(3)

    if not all_supplies:
        set_status(ss, 'Поставки', '❌ Нет данных'); return

    CARGO_TYPES = {1: 'Монопаллет', 2: 'Суперсейф', 3: 'QR-поставка'}
    headers = ['ID поставки', 'Название', 'Статус', 'Тип груза', 'Создана', 'Закрыта']
    rows    = [headers]
    for s in all_supplies:
        rows.append([
            s.get('id', ''),
            s.get('name', ''),
            'Закрыта' if s.get('done') else 'Открыта',
            CARGO_TYPES.get(s.get('cargoType', 0), '—'),
            s.get('createdAt', ''),
            s.get('closedAt', ''),
        ])
    write_sheet(ss, 'Поставки', rows)
    set_status(ss, 'Поставки', f'✅ Готово — {len(all_supplies)} поставок')

# ── Точка входа ────────────────────────────────────────────────────────────────

def main() -> None:
    log.info('=== main3 started ===')
    ss      = get_spreadsheet()
    api_key = get_api_key(ss)

    today       = datetime.now()
    yesterday   = (today - timedelta(days=1)).strftime('%Y-%m-%d')
    week_from   = (today - timedelta(days=7)).strftime('%Y-%m-%d')
    days14_from = (today - timedelta(days=14)).strftime('%Y-%m-%d')
    month_from  = today.replace(day=1).strftime('%Y-%m-%d')

    for sheet_name, df, dt in [
        ('Воронка День',    yesterday,   yesterday),
        ('Воронка Неделя',  week_from,   yesterday),
        ('Воронка 14 Дней', days14_from, yesterday),
        ('Воронка Месяц',   month_from,  yesterday),
    ]:
        load_funnel_period(api_key, df, dt, ss, sheet_name)
        time.sleep(FUNNEL_SLEEP)

    load_prices(api_key, ss);    time.sleep(10)
    load_cards(api_key, ss);     time.sleep(10)
    load_supplies(api_key, ss)

    set_status(ss, 'Все данные', f'✅ Завершено — {datetime.now().strftime("%d.%m.%Y %H:%M")}')
    log.info('=== main3 complete ===')

if __name__ == '__main__':
    main()

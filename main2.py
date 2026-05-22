#!/usr/bin/env python3
"""
KLR Analytics — main2.py
РК по периодам + НМ-отчёт (детальная аналитика по артикулам)
"""

from __future__ import annotations

import logging
import time
from datetime import datetime, timedelta

import gspread
import requests
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

SPREADSHEET_ID   = '1SOLbCBhGcnsrwiW9JSiw0wH3YMQZcVMWYh_sNovXQjI'
CREDENTIALS_FILE = 'credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

ADV_BASE       = 'https://advert-api.wildberries.ru'
ANALYTICS_BASE = 'https://seller-analytics-api.wildberries.ru'

CAMP_CHUNK  = 50
ADV_SLEEP   = 90
WRITE_BATCH = 500

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
    val = ss.worksheet('Настройки').acell('B2').value
    if not val or not val.strip():
        raise RuntimeError('API ключ WB не найден в ячейке B2 листа Настройки')
    return val.strip()

def set_status(ss: gspread.Spreadsheet, name: str, status: str) -> None:
    try:
        now = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
        ss.worksheet('Настройки').update(
            values=[['Последнее обновление:', name, now, status]], range_name='D2'
        )
        log.info('%s — %s', name, status)
    except Exception as exc:
        log.warning('set_status error: %s', exc)

def format_header(ws: gspread.Worksheet, num_cols: int) -> None:
    last_cell = rowcol_to_a1(1, num_cols)
    ws.format(f'A1:{last_cell}', {
        'backgroundColor': {'red': 0.122, 'green': 0.306, 'blue': 0.475},
        'textFormat': {
            'bold': True,
            'foregroundColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0},
        },
        'horizontalAlignment': 'CENTER',
    })
    ws.freeze(rows=1)

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
    if rows:
        format_header(ws, len(rows[0]))
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
            if resp.status_code == 401:
                log.error('HTTP 401 — token scope not allowed, skipping')
                return None
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

def safe_div(num: float, den: float, scale: float = 1, decimals: int = 2) -> float:
    return round(num / den * scale, decimals) if den else 0.0

# ── Кампании ───────────────────────────────────────────────────────────────────

def get_campaigns(api_key: str) -> list[int]:
    resp = wb_request('get', f'{ADV_BASE}/adv/v1/promotion/count', api_key)
    if not resp:
        return []
    all_ids: list[int] = [
        int(advert['advertId'])
        for group  in resp.json().get('adverts', [])
        if group.get('status') != -1
        for advert in group.get('advert_list', [])
        if advert.get('advertId')
    ]
    log.info('Campaigns found: %d', len(all_ids))
    return all_ids

# ── fullstats ──────────────────────────────────────────────────────────────────

def fetch_fullstats(api_key: str, campaign_ids: list[int], date_from: str, date_to: str) -> list[dict]:
    all_stats: list[dict] = []
    for i in range(0, len(campaign_ids), CAMP_CHUNK):
        chunk   = campaign_ids[i : i + CAMP_CHUNK]
        ids_str = ','.join(map(str, chunk))
        url     = f'{ADV_BASE}/adv/v3/fullstats?ids={ids_str}&beginDate={date_from}&endDate={date_to}'
        resp    = wb_request('get', url, api_key)
        if resp:
            data = resp.json()
            if isinstance(data, list):
                all_stats.extend(data)
        if i + CAMP_CHUNK < len(campaign_ids):
            time.sleep(ADV_SLEEP)
    log.info('Fullstats: %d campaign records', len(all_stats))
    return all_stats

# ── РК периоды ─────────────────────────────────────────────────────────────────

def write_rk_period(
    month_stats: list[dict],
    date_from:   str,
    date_to:     str,
    ss:          gspread.Spreadsheet,
    sheet_name:  str,
) -> None:
    log.info('write_rk_period [%s]: %s → %s', sheet_name, date_from, date_to)
    set_status(ss, sheet_name, '🔄 Записываем...')

    dt_from = datetime.strptime(date_from, '%Y-%m-%d')
    dt_to   = datetime.strptime(date_to,   '%Y-%m-%d')

    nm_stats: dict[tuple, dict] = {}
    for camp in month_stats:
        adv_id = camp.get('advertId')
        for day in camp.get('days', []):
            day_str = (day.get('date') or '')[:10]
            if not day_str:
                continue
            try:
                day_dt = datetime.strptime(day_str, '%Y-%m-%d')
            except ValueError:
                continue
            if not (dt_from <= day_dt <= dt_to):
                continue
            for app in day.get('apps', []):
                for nm in app.get('nms', []):
                    nm_id = nm.get('nmId')
                    if not nm_id:
                        continue
                    key = (nm_id, str(adv_id))
                    if key not in nm_stats:
                        nm_stats[key] = {
                            'nmId': nm_id, 'name': nm.get('name', '—'),
                            'advId': adv_id,
                            'views': 0, 'clicks': 0, 'sum': 0, 'orders': 0, 'orderSum': 0,
                        }
                    s = nm_stats[key]
                    s['views']    += nm.get('views',     0)
                    s['clicks']   += nm.get('clicks',    0)
                    s['sum']      += nm.get('sum',       0)
                    s['orders']   += nm.get('orders',    0)
                    s['orderSum'] += nm.get('sum_price', 0)

    if not nm_stats:
        set_status(ss, sheet_name, '❌ Нет данных'); return

    headers = [
        'Артикул WB', 'Название', 'ID кампании',
        'Показы', 'Клики', 'CTR, %', 'CPC, ₽',
        'Расход, ₽', 'Заказы', 'Сумма заказов, ₽', 'ДРР, %',
    ]
    rows = [headers]
    for s in nm_stats.values():
        rows.append([
            s['nmId'], s['name'], s['advId'],
            s['views'], s['clicks'],
            safe_div(s['clicks'], s['views'],    scale=100),
            safe_div(s['sum'],    s['clicks']),
            s['sum'], s['orders'], s['orderSum'],
            safe_div(s['sum'],    s['orderSum'], scale=100, decimals=1),
        ])
    rows[1:] = sorted(rows[1:], key=lambda r: r[7], reverse=True)
    write_sheet(ss, sheet_name, rows)
    set_status(ss, sheet_name, f'✅ Готово — {len(nm_stats)} артикулов')

# ── НМ-отчёт ───────────────────────────────────────────────────────────────────

def load_nm_report(api_key: str, ss: gspread.Spreadsheet) -> None:
    set_status(ss, 'НМ Отчёт', '🔄 Загружается...')

    today     = datetime.now()
    date_from = (today - timedelta(days=30)).strftime('%Y-%m-%d 00:00:00')
    date_to   = today.strftime('%Y-%m-%d 23:59:59')
    url       = f'{ANALYTICS_BASE}/api/v2/nm-report/detail'

    all_cards: list[dict] = []
    page = 1

    while True:
        body = {
            'brandNames': [], 'objectIDs': [], 'tagIDs': [], 'nmIDs': [],
            'timezone': 'Europe/Moscow',
            'period': {'begin': date_from, 'end': date_to},
            'orderBy': {'field': 'openCard', 'mode': 'desc'},
            'page': page,
        }
        resp = wb_request('post', url, api_key, json=body)
        if not resp:
            break
        data = resp.json().get('data', {})
        cards = data.get('cards', [])
        if not cards:
            break
        all_cards.extend(cards)
        log.info('НМ отчёт стр %d: %d артикулов', page, len(all_cards))
        if not data.get('isNextPage', False):
            break
        page += 1
        time.sleep(5)

    if not all_cards:
        set_status(ss, 'НМ Отчёт', '❌ Нет данных'); return

    headers = [
        'Артикул WB', 'Артикул продавца', 'Бренд', 'Предмет',
        # Текущий период
        'Переходы в карточку', 'Добавили в корзину', 'Заказали, шт', 'Заказали, ₽',
        'Выкупили, шт', 'Выкупили, ₽', '% выкупа',
        'Конв. клик→корзина, %', 'Конв. корзина→заказ, %',
        'Отменили, шт', 'Отменили, ₽',
        # Предыдущий период
        'Переходы (пред.)', 'В корзину (пред.)', 'Заказали шт (пред.)', 'Заказали ₽ (пред.)',
        'Выкупили шт (пред.)', 'Выкупили ₽ (пред.)', '% выкупа (пред.)',
        # Остатки
        'Остатки WB', 'Остатки продавца',
    ]
    rows = [headers]

    for card in all_cards:
        s  = card.get('statistics', {}).get('selectedPeriod', {})
        p  = card.get('statistics', {}).get('previousPeriod', {})
        st = card.get('stocks', {})
        rows.append([
            card.get('nmID', ''),         card.get('vendorCode', ''),
            card.get('brandName', ''),    card.get('object', {}).get('name', ''),
            s.get('openCardCount', 0),    s.get('addToCartCount', 0),
            s.get('ordersCount', 0),      s.get('ordersSumRub', 0),
            s.get('buyoutsCount', 0),     s.get('buyoutsSumRub', 0),
            s.get('buyoutPercent', 0),
            s.get('addToCartConversion', 0), s.get('cartToOrderConversion', 0),
            s.get('cancelCount', 0),      s.get('cancelSumRub', 0),
            p.get('openCardCount', 0),    p.get('addToCartCount', 0),
            p.get('ordersCount', 0),      p.get('ordersSumRub', 0),
            p.get('buyoutsCount', 0),     p.get('buyoutsSumRub', 0),
            p.get('buyoutPercent', 0),
            st.get('stocksMp', 0),        st.get('stocksWb', 0),
        ])

    write_sheet(ss, 'НМ Отчёт', rows)
    set_status(ss, 'НМ Отчёт', f'✅ Готово — {len(all_cards)} артикулов')

# ── Точка входа ────────────────────────────────────────────────────────────────

def main() -> None:
    log.info('=== main2 started ===')
    ss      = get_spreadsheet()
    api_key = get_api_key(ss)

    today       = datetime.now()
    yesterday   = today - timedelta(days=1)
    month_first = today.replace(day=1)
    month_from  = min(month_first, yesterday).strftime('%Y-%m-%d')
    yesterday   = yesterday.strftime('%Y-%m-%d')
    week_from   = (today - timedelta(days=7)).strftime('%Y-%m-%d')
    days14_from = (today - timedelta(days=14)).strftime('%Y-%m-%d')

    campaign_ids = get_campaigns(api_key)
    if not campaign_ids:
        log.warning('Нет кампаний — РК пропущены')
    else:
        log.info('Загружаем fullstats за месяц...')
        month_stats = fetch_fullstats(api_key, campaign_ids, month_from, yesterday)
        time.sleep(5)
        days14_to = (today - timedelta(days=8)).strftime('%Y-%m-%d')
        for sheet_name, df, dt in [
            ('РК День',    yesterday,   yesterday),
            ('РК Неделя',  week_from,   yesterday),
            ('РК 14 Дней', days14_from, days14_to),
            ('РК Месяц',   month_from,  yesterday),
        ]:
            write_rk_period(month_stats, df, dt, ss, sheet_name)
            time.sleep(5)

    time.sleep(10)
    load_nm_report(api_key, ss)

    log.info('=== main2 complete ===')

if __name__ == '__main__':
    main()

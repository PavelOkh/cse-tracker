"""Read public DPD HTML without customer credentials. No access-control bypass."""
import re
import time
from datetime import datetime
from urllib.parse import quote


class PublicPageError(ValueError):
    pass


def parse_public_html(number, html):
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    for tag in soup(['script', 'style', 'noscript']):
        tag.decompose()
    text = soup.get_text(' ', strip=True)
    if 'Forbidden' in text or 'If you are not a bot' in text:
        raise PublicPageError('Сайт DPD ограничил автоматический доступ')
    # Anchor to the order information block, avoiding navigation, ads and unrelated dates.
    headings = soup.find_all(string=re.compile(r'^\s*Информация по заказу\s*$'))
    if len(headings) != 1:
        raise PublicPageError('Не найдена однозначная информация по заказу; возможны загрузка JavaScript, отсутствие номера или изменение страницы')
    scope = headings[0].parent
    while scope and not (scope.find(string=re.compile(r'Номер заказа')) and scope.find(string=re.compile(r'Дата доставки'))):
        scope = scope.parent
    if scope is None:
        raise PublicPageError('Не найдены поля заказа и даты доставки')

    def field(label):
        matches = scope.find_all(string=re.compile(r'^\s*' + re.escape(label) + r'\s*:?\s*$'))
        if len(matches) != 1:
            raise PublicPageError('Не найдено однозначное поле: ' + label)
        parent = matches[0].parent
        # Screenshot structure: item-heading and item-data in the same item container.
        for _ in range(3):
            parent = parent.parent
            if parent is None:
                break
            value = parent.get_text(' ', strip=True)
            value = re.sub(r'^\s*' + re.escape(label) + r'\s*:?\s*', '', value).strip()
            if value:
                return value
        raise PublicPageError('Пустое поле: ' + label)

    actual = field('Номер заказа')
    if actual != str(number):
        raise PublicPageError('Номер в ответе не совпадает с запрошенным')
    date = field('Дата доставки')
    if date in ('-', '—', '–'):
        date = ''
    elif re.fullmatch(r'\d{2}\.\d{2}\.\d{4}', date):
        datetime.strptime(date, '%d.%m.%Y')
    else:
        raise PublicPageError('Неожиданный формат даты доставки')
    statuses = re.findall(r'(?:^|\s)Статус\s*:\s*(.+?)(?=\s*(?:Посмотреть историю заказа|Детальная информация|$))', text)
    if len(statuses) != 1:
        raise PublicPageError('Не найден однозначный статус отправления')
    status = statuses[0].strip()
    # Date alone is insufficient evidence of completed delivery.
    delivered = bool(re.match(r'^Доставлен(?:[аоы])?(?:\s|\(|$)', status, re.I))
    return {'Накладная':str(number), 'Статус':status,
            'Дата доставки':date if delivered else '',
            'Дата доставки на сайте':date,
            'Источник':'Публичная страница DPD',
            'Ссылка':f'https://dpd.ru/search/{quote(str(number), safe="")}', 'Ошибка':''}


def check_many_public(numbers, progress=None):
    import requests
    numbers = list(dict.fromkeys(str(n).strip() for n in numbers if str(n).strip()))
    rows = []
    with requests.Session() as session:
        session.headers.update({'User-Agent':'DeliveryTracker/1.1', 'Accept':'text/html', 'Accept-Language':'ru'})
        blocked = False
        for i, number in enumerate(numbers):
            row = {'Накладная':number, 'Статус':'Не проверено', 'Дата доставки':'', 'Ошибка':''}
            if blocked:
                row['Ошибка'] = 'Опрос остановлен после HTTP 403/429; повторите позже или используйте официальный API'
            else:
                try:
                    response = session.get(f'https://dpd.ru/search/{quote(number, safe="")}', timeout=(10,25))
                    if response.status_code in (403,429):
                        blocked = True
                    response.raise_for_status()
                    row = parse_public_html(number,response.text)
                except requests.RequestException as exc:
                    code = getattr(getattr(exc,'response',None),'status_code',None)
                    row['Ошибка'] = f'HTTP {code}' if code else type(exc).__name__
                except (PublicPageError,ValueError) as exc:
                    row['Ошибка'] = str(exc)
            rows.append(row)
            if progress:
                progress((i+1)/len(numbers))
            if not blocked and i+1 < len(numbers):
                time.sleep(1)
    return rows

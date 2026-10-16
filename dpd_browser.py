"""Visible Playwright worker. Run separately from Streamlit (Windows event loops)."""
import json
import re
import sys
import shutil
import time
from urllib.parse import quote
from dpd_public import parse_public_html, PublicPageError


def blocked_page(status, text):
    if status in (403,429):
        return f'HTTP {status}: сайт ограничил доступ'
    if re.search(r'if you are not a bot|access denied|forbidden|подтвердите.{0,30}(?:робот|человек)|проверка.{0,20}браузера|captcha', text, re.I):
        return 'Сайт показал проверку доступа / CAPTCHA'
    return ''


def check_many_browser(numbers, progress=None, delay=2, factory=None, mode="local"):
    """No persistent profiles, stealth, proxy rotation or CAPTCHA solving."""
    if factory is None:
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise RuntimeError('Установите requirements-browser.txt и выполните python -m playwright install chromium') from exc
        factory = sync_playwright
    numbers = list(dict.fromkeys(str(n).strip() for n in numbers if str(n).strip()))
    rows = []
    if not numbers:
        return rows
    stop_reason = ''
    with factory() as pw:
        launch_options = {'headless':False}
        if mode == 'cloud':
            executable = shutil.which('chromium') or shutil.which('chromium-browser')
            if not executable:
                raise RuntimeError('Chromium не установлен на сервере. Добавьте packages.txt из архива в корень GitHub-репозитория и пересоберите приложение Streamlit.')
            launch_options = {'headless':True, 'executable_path':executable}
        elif mode != 'local':
            raise RuntimeError('Неизвестный режим запуска браузера')
        browser = pw.chromium.launch(**launch_options)
        try:
            context = browser.new_context(locale='ru-RU')
            page = context.new_page()
            page.set_default_timeout(12000)
            for i, number in enumerate(numbers):
                row = {'Накладная':number,'Статус':'Не проверено','Дата доставки':'','Ошибка':''}
                if stop_reason:
                    row['Ошибка'] = 'Опрос остановлен: ' + stop_reason
                else:
                    try:
                        response = page.goto(f'https://dpd.ru/search/{quote(number, safe="")}',
                                             wait_until='domcontentloaded', timeout=30000)
                        status = response.status if response else None
                        stop_reason = blocked_page(status, '')
                        if not stop_reason:
                            text = page.locator('body').inner_text(timeout=5000)
                            stop_reason = blocked_page(status, text)
                        if stop_reason:
                            row['Ошибка'] = stop_reason
                        elif status is not None and status >= 400:
                            row['Ошибка'] = f'HTTP {status}'
                        else:
                            # Wait for the rendered order block, not analytics/network idle.
                            page.get_by_text('Информация по заказу',exact=True).wait_for(state='visible',timeout=12000)
                            # Parse repeatedly during hydration, bounded to 5 seconds.
                            deadline = time.monotonic()+5
                            while True:
                                try:
                                    stop_reason = blocked_page(None,page.locator('body').inner_text(timeout=1000))
                                    if stop_reason:
                                        row['Ошибка'] = stop_reason
                                        break
                                    row = parse_public_html(number,page.content())
                                    row['Источник'] = 'Playwright / Chromium — ' + ('сервер Streamlit' if mode == 'cloud' else 'локальный ПК')
                                    break
                                except PublicPageError:
                                    if time.monotonic() >= deadline:
                                        raise
                                    page.wait_for_timeout(500)
                    except Exception as exc:
                        if not browser.is_connected() or page.is_closed():
                            stop_reason = 'Окно браузера закрыто'
                            row['Ошибка'] = stop_reason
                        elif isinstance(exc,PublicPageError):
                            row['Ошибка'] = str(exc)
                        else:
                            # Recheck access after navigation/selector timeout.
                            try:
                                stop_reason = blocked_page(None,page.locator('body').inner_text(timeout=1000))
                            except Exception:
                                pass
                            row['Ошибка'] = stop_reason or f'{type(exc).__name__}: не удалось загрузить или разобрать страницу'
                rows.append(row)
                if progress:
                    progress(i+1,len(numbers),row)
                if not stop_reason and i+1 < len(numbers):
                    time.sleep(max(1,delay))
        finally:
            try:
                browser.close()
            except Exception:
                pass  # A window already closed by the user must not discard result rows.
    return rows


def main():
    def emit(value):
        print(json.dumps(value,ensure_ascii=True),flush=True)
    try:
        payload = json.loads(sys.stdin.read())
        check_many_browser(payload['numbers'], mode=payload.get('mode','local'),
            progress=lambda done,total,row: emit({'type':'row','done':done,'total':total,'row':row}))
        emit({'type':'done'})
    except Exception as exc:
        # Operational errors, no page contents/cookies in logs.
        if isinstance(exc,RuntimeError):
            message = str(exc)
        else:
            message = f'{type(exc).__name__}: ошибка браузера. Для Streamlit Cloud проверьте установку chromium из packages.txt; для локального запуска выполните setup_windows.bat.'
        emit({'type':'error','message':message})
        return 1
    return 0

if __name__ == '__main__':
    raise SystemExit(main())

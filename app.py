from pathlib import Path
import io
import os
from datetime import datetime
from uuid import uuid4
from contextlib import closing
import tempfile
from tracker_utils import parse_numbers, unique_numbers, save_csv, excel_safe
import runpy
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Проверка доставки CSE / DPD', page_icon='📦', layout='wide')
st.sidebar.caption('Версия 1.3 — Streamlit Cloud / локально')
carrier = st.sidebar.selectbox('Транспортная компания', ['CSE', 'DPD'])
if carrier == 'CSE':
    runpy.run_path(str(Path(__file__).with_name('CSE_app.py')), run_name='__main__')
else:
    from dpd_tracker import DPDTracker
    st.title('DPD — проверка списка накладных')
    st.caption('Публичный сайт: результат по заказу. Официальный API: по каждому месту и направлению возврата.')
    def secret(key):
        try:
            return str(st.secrets.get(key, ''))
        except FileNotFoundError:
            return ''
    local = os.environ.get('DPD_LOCAL_BROWSER') == '1'
    modes = ['Playwright — сервер Streamlit', 'Playwright — видимое окно (локально)', 'Публичный сайт — без ключа', 'Официальный API']
    mode = st.sidebar.radio('Способ проверки DPD', modes, index=1 if local else 0)
    client, key, test, year = '', '', False, 0
    if mode == 'Официальный API':
        client = st.sidebar.text_input('Номер клиента DPD', value=secret('DPD_CLIENT_NUMBER'))
        key = st.sidebar.text_input('API-ключ DPD', value=secret('DPD_CLIENT_KEY'), type='password')
        test = st.sidebar.checkbox('Тестовый сервер DPD')
        year = st.number_input('Год приёма отправлений (0 — не уточнять)', min_value=0, max_value=2100, value=0)
    if mode == 'Playwright — сервер Streamlit':
        st.info('Браузер работает на сервере; результат появится в таблице. API-ключ не нужен. При блокировке DPD опрос остановится.')
    elif mode == 'Playwright — видимое окно (локально)':
        st.info('Запустите приложение на своём ПК через start_windows.bat. Откроется отдельное окно Chromium. Не закрывайте его до окончания проверки. При блокировке опрос остановится.')
    text = st.text_area('Накладные: по одной в строке, либо через запятую / точку с запятой')
    uploaded = st.file_uploader('Или загрузите CSV / XLSX', type=['csv','xlsx'])
    numbers = parse_numbers(text)
    if uploaded:
        try:
            data = pd.read_excel(uploaded, dtype=str, keep_default_na=False) if uploaded.name.lower().endswith('.xlsx') else pd.read_csv(uploaded, dtype=str, keep_default_na=False, sep=None, engine='python')
            col = st.selectbox('Столбец с накладными',data.columns)
            numbers += data[col].tolist()
        except Exception:
            st.error('Не удалось прочитать файл. Проверьте формат и строку заголовков.')
    st.caption('Числовые номера в Excel храните как текст: утраченные ведущие нули восстановить нельзя.')
    numbers = unique_numbers(numbers)
    st.write(f'Уникальных номеров: {len(numbers)}')
    if st.button('Проверить', disabled=not numbers):
        st.session_state.pop('dpd_result', None)
        st.session_state.pop('dpd_checkpoint', None)
        if mode.startswith('Playwright'):
            browser_mode = 'cloud' if mode == 'Playwright — сервер Streamlit' else 'local'
            if browser_mode == 'local' and not local:
                st.error('Режим работает только при локальном запуске. Распакуйте архив на ПК и запустите start_windows.bat. В Streamlit Cloud видимое окно на вашем ПК открыть нельзя.')
            else:
                bar = st.progress(0)
                rows = []
                table = st.empty()
                from browser_job import browser_events
                filename = 'dpd_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid4().hex[:6] + '.csv'
                if browser_mode == 'local':
                    checkpoint = Path(__file__).parent / 'results' / filename
                else:
                    if 'dpd_temp_dir' not in st.session_state:
                        st.session_state['dpd_temp_dir'] = tempfile.TemporaryDirectory(prefix='dpd_results_')
                    checkpoint = Path(st.session_state['dpd_temp_dir'].name) / filename
                st.session_state['dpd_checkpoint'] = str(checkpoint)
                save_failed = False
                try:
                    with closing(browser_events(numbers, mode=browser_mode)) as events:
                        for event in events:
                            if event['type'] == 'row':
                                rows.append(event['row'])
                                st.session_state['dpd_result'] = pd.DataFrame(rows)
                                try:
                                    save_csv(rows, checkpoint)
                                except OSError:
                                    if not save_failed:
                                        st.warning('Не удалось сохранить CSV на диск. Скачайте результат кнопкой под таблицей.')
                                        save_failed = True
                                bar.progress(event['done']/event['total'])
                                table.dataframe(st.session_state['dpd_result'], hide_index=True)
                            elif event['type'] == 'error':
                                st.error(event['message'])
                    checked = {row['Накладная'] for row in rows}
                    for number in numbers:
                        if number not in checked:
                            rows.append({'Накладная':number,'Статус':'Не проверено','Дата доставки':'',
                                         'Ошибка':'Проверка прервана до получения результата'})
                    st.session_state['dpd_result'] = pd.DataFrame(rows)
                    try:
                        save_csv(rows, checkpoint)
                    except OSError:
                        save_failed = True
                except Exception as exc:
                    st.error(f'Ошибка запуска браузера ({type(exc).__name__}). Уже полученные строки доступны ниже.')
                table.empty()
                if checkpoint.exists() and not save_failed:
                    if browser_mode == 'local':
                        st.caption(f'Копия результата сохранена на компьютере: {checkpoint}')
                    else:
                        st.caption('Результат временно сохранён на сервере. Скачайте XLSX или CSV до закрытия сеанса.')

        elif mode == 'Публичный сайт — без ключа':
            from dpd_public import check_many_public
            st.session_state['dpd_result'] = pd.DataFrame(check_many_public(numbers, st.progress(0).progress))
        elif not client or not key:
            st.error('Введите номер клиента и API-ключ DPD.')
        else:
            bar = st.progress(0)
            try:
                tracker = DPDTracker(client,key,test)
                st.session_state['dpd_result'] = pd.DataFrame(tracker.check_many(numbers,year or None,bar.progress))
            except Exception as exc:
                st.error(f'Не удалось подключиться к DPD ({type(exc).__name__}). Проверьте доступность HTTPS/WSDL и реквизиты.')
    if 'dpd_result' in st.session_state:
        result = st.session_state['dpd_result']
        st.dataframe(result, hide_index=True, width='stretch')
        buf = io.BytesIO()
        # Force text cells, preventing formula interpretation of imported/user values.
        export = result.copy()
        export = export.map(excel_safe)
        export.to_excel(buf,index=False,engine='openpyxl')
        st.download_button('Скачать результат XLSX',buf.getvalue(),'dpd_results.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        st.download_button('Скачать результат CSV',export.to_csv(index=False,sep=';').encode('utf-8-sig'),'dpd_results.csv',mime='text/csv')

from pathlib import Path
import io
import runpy
import re
import pandas as pd
import streamlit as st

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
    mode = st.sidebar.radio('Способ проверки DPD', ['Публичный сайт — без ключа', 'Официальный API'])
    client = st.sidebar.text_input('Номер клиента DPD', value=secret('DPD_CLIENT_NUMBER'))
    key = st.sidebar.text_input('API-ключ DPD', value=secret('DPD_CLIENT_KEY'), type='password')
    test = st.sidebar.checkbox('Тестовый сервер DPD')
    year = st.number_input('Год приёма отправлений (0 — не уточнять)', min_value=0, max_value=2100, value=0)
    text = st.text_area('Накладные: по одной в строке, либо через запятую / точку с запятой')
    uploaded = st.file_uploader('Или загрузите CSV / XLSX', type=['csv','xlsx'])
    numbers = re.split(r'[\s,;]+',text.strip()) if text.strip() else []
    if uploaded:
        try:
            data = pd.read_excel(uploaded, dtype=str, keep_default_na=False) if uploaded.name.lower().endswith('.xlsx') else pd.read_csv(uploaded, dtype=str, keep_default_na=False, sep=None, engine='python')
            col = st.selectbox('Столбец с накладными',data.columns)
            numbers += data[col].tolist()
        except Exception:
            st.error('Не удалось прочитать файл. Проверьте формат и строку заголовков.')
    st.caption('Числовые номера в Excel храните как текст: утраченные ведущие нули восстановить нельзя.')
    numbers = list(dict.fromkeys(n.strip() for n in numbers if n.strip()))
    st.write(f'Уникальных номеров: {len(numbers)}')
    if st.button('Проверить', disabled=not numbers):
        st.session_state.pop('dpd_result', None)
        if mode == 'Публичный сайт — без ключа':
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
        st.dataframe(result, hide_index=True, use_container_width=True)
        buf = io.BytesIO()
        # Force text cells, preventing formula interpretation of imported/user values.
        export = result.copy()
        export = export.map(lambda v: "'"+v if isinstance(v,str) and v.startswith(('=','+','-','@')) else v)
        export.to_excel(buf,index=False,engine='openpyxl')
        st.download_button('Скачать результат XLSX',buf.getvalue(),'dpd_results.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

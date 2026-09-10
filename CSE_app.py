import requests
import pandas as pd
import time
import io
import streamlit as st
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl
from openpyxl.styles import PatternFill

# Настройка страницы браузера
st.set_page_config(page_title="CSE Трекинг", page_icon="📦", layout="centered")

def process_single_number(number):
    url = f'https://lk.cse.ru/api/new-track/{number}'
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Origin': 'https://www.cse.ru',
        'Referer': 'https://www.cse.ru/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 YaBrowser/26.8.0.0 Safari/537.36'
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()
        
        found_items = data.get('found', [])
        if not found_items:
            return {
                "Накладная": number, "Статус": "Данные не найдены", "Дата доставки": "",
                "Дата посл. статуса": "", "Новый номер": "", "Статус нового номера": "",
                "Дата доставки нового": "", "Дата статуса нового": "", "Тип подсветки": "желтый"
            }
            
        track_data = found_items[0]
        
        state = track_data.get('State', 'Статус не указан')
        info = track_data.get('Info', '')
        
        current_status = f"{state}. {info}" if info else state
        delivery_date = ""
        last_status_date = ""
        new_waybill_number = ""
        new_state = ""
        new_delivery_date = ""
        new_last_status_date = ""
        
        highlight_type = "желтый"

        history_dict = track_data.get('History', {})
        waybill_info = history_dict.get('waybill_info', [])
        
        if waybill_info:
            last_status_date = waybill_info[0].get('EventDate', '')

        for event in waybill_info:
            event_name = event.get('EventName', '')
            event_name_en = event.get('EventNameEn', '')
            event_date = event.get('EventDate', '')
            
            doc = event.get('Document')
            if doc and isinstance(doc, dict):
                doc_number = doc.get('Number', '')
                doc_state = doc.get('State', '')
                if doc_number:
                    new_waybill_number = doc_number
                    new_state = doc_state
                    
                    doc_history = doc.get('History', [])
                    if doc_history:
                        new_last_status_date = doc_history[0].get('EventDate', '')
                        
                    for sub_event in doc_history:
                        if "Доставка завершена" in sub_event.get('EventName', ''):
                            new_delivery_date = sub_event.get('EventDate', '')
                            break
                    break

            if "Доставка завершена" in event_name or event_name_en == "Delivery completed":
                current_status = "Доставка завершена"
                delivery_date = event_date
                last_status_date = ""
                highlight_type = "нет"
                break

        if "возврат" in state.lower() or "возвращается" in current_status.lower():
            highlight_type = "красный"
            if not new_waybill_number and "497-" in current_status:
                parts = current_status.split()
                for p in parts:
                    if "497-" in p:
                        new_waybill_number = p.strip(".,;")
                        break
        elif new_waybill_number or "добавочная" in current_status.lower() or "смена" in current_status.lower():
            highlight_type = "фиолетовый"

        if current_status == "Доставка завершена":
            highlight_type = "нет"

        return {
            "Накладная": number,
            "Статус": current_status,
            "Дата доставки": delivery_date,
            "Дата посл. статуса": last_status_date if current_status != "Доставка завершена" else "",
            "Новый номер": new_waybill_number,
            "Статус нового номера": new_state,
            "Дата доставки нового": new_delivery_date,
            "Дата статуса нового": new_last_status_date if (new_state and new_state != "Доставка завершена") else "",
            "Тип подсветки": highlight_type
        }

    except Exception as e:
        return {
            "Накладная": number, "Статус": "Ошибка обработки", "Дата доставки": "",
            "Дата посл. статуса": "", "Новый номер": "", "Статус нового номера": "",
            "Дата доставки нового": "", "Дата статуса нового": "", "Тип подсветки": "желтый"
        }

def generate_colored_excel(df_data):
    # Создаем Excel в памяти через openpyxl для раскраски
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_data.to_excel(writer, index=False, sheet_name='Результат')
    
    output.seek(0)
    wb = openpyxl.load_workbook(output)
    ws = wb.active
    
    yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    purple_fill = PatternFill(start_color="E1D5E7", end_color="E1D5E7", fill_type="solid")
    red_fill = PatternFill(start_color="F8CECC", end_color="F8CECC", fill_type="solid")
    
    headers = [cell.value for cell in ws[1]]
    type_col_idx = headers.index("Тип подсветки") + 1
    
    for row_idx in range(2, ws.max_row + 1):
        highlight_type = ws.cell(row=row_idx, column=type_col_idx).value
        target_fill = None
        if highlight_type == "желтый":
            target_fill = yellow_fill
        elif highlight_type == "фиолетовый":
            target_fill = purple_fill
        elif highlight_type == "красный":
            target_fill = red_fill
            
        if target_fill:
            for col_idx in range(1, len(headers)):
                ws.cell(row=row_idx, column=col_idx).fill = target_fill
                
    ws.delete_cols(type_col_idx)
    
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output.getvalue()

def main():
    st.title("📦 Массовая проверка накладных CSE")
    st.write("Загрузите Excel-файл с номерами накладных в первом столбце. Сохраняется исходный порядок и количество строк.")

    # Информационный блок с описанием цветовой индикации
    with st.expander("🎨 Справка по цветовой индикации в скачанном Excel-файле"):
        st.markdown("""
        В готовом отчете строки автоматически подсвечиваются цветом для быстрого визуального контроля:
        * 🟨 **Желтый фон** — доставка еще в процессе / не завершена.
        * 🟪 **Фиолетовый фон** — по отправлению произошла смена номера накладной (досыл / добавочная накладная).
        * 🟥 **Красный фон** — оформлен возврат отправления отправителю.
        * *Без заливки* — доставка успешно завершена.
        """)

    uploaded_file = st.file_uploader("Выберите файл .xlsx", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            df_input = pd.read_excel(uploaded_file)
            original_tracking_list = df_input.iloc[:, 0].tolist()
            
            original_list = [str(num).strip() if pd.notna(num) else "" for num in original_tracking_list]
            unique_numbers = list(dict.fromkeys([num for num in original_list if num]))
            
            st.info(f"Всего строк в файле: **{len(original_list)}** | Уникальных номеров для API: **{len(unique_numbers)}**")
            
            if st.button("Начать проверку", type="primary"):
                start_time = time.time()
                my_bar = st.progress(0, text="Идет опрос API CSE. Пожалуйста, подождите...")
                
                unique_results = []
                with ThreadPoolExecutor(max_workers=10) as executor:
                    futures = [executor.submit(process_single_number, num) for num in unique_numbers]
                    for i, future in enumerate(as_completed(futures)):
                        unique_results.append(future.result())
                        progress = (i + 1) / len(unique_numbers)
                        my_bar.progress(progress, text=f"Обработано {i + 1} из {len(unique_numbers)} уникальных номеров...")
                
                results_dict = {res["Накладная"]: res for res in unique_results}
                
                final_results = []
                for num in original_list:
                    if num == "":
                        final_results.append({
                            "Накладная": "", "Статус": "Пустая строка", "Дата доставки": "",
                            "Дата посл. статуса": "", "Новый номер": "", "Статус нового номера": "",
                            "Дата доставки нового": "", "Дата статуса нового": "", "Тип подсветки": "нет"
                        })
                    else:
                        final_results.append(results_dict.get(num, {
                            "Накладная": num, "Статус": "Ошибка кэша", "Дата доставки": "",
                            "Дата посл. статуса": "", "Новый номер": "", "Статус нового номера": "",
                            "Дата доставки нового": "", "Дата статуса нового": "", "Тип подсветки": "желтый"
                        }))
                
                elapsed_time = round(time.time() - start_time, 1)
                st.success(f"✅ Проверка завершена за {elapsed_time} сек.")
                
                df_output = pd.DataFrame(final_results)
                
                # Показываем таблицу на экране (без технической колонки типа подсветки)
                st.dataframe(df_output.drop(columns=["Тип подсветки"]), use_container_width=True)
                
                # Генерируем цветной Excel в памяти
                excel_data = generate_colored_excel(df_output)
                
                current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                original_name = uploaded_file.name.rsplit('.', 1)[0]
                download_name = f"{original_name}_Результат_{current_time}.xlsx"
                
                st.download_button(
                    label="📥 Скачать результат с раскраской (Excel)",
                    data=excel_data,
                    file_name=download_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"Произошла ошибка при обработке файла: {e}")

if __name__ == "__main__":
    main()

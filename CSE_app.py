import requests
import pandas as pd
import time
import io
import streamlit as st
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# Настройка страницы браузера
st.set_page_config(page_title="CSE Трекинг", page_icon="📦", layout="centered")

def process_single_number(number):
    url = f'https://lk.cse.ru/api/new-track/{number}'
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Origin': 'https://www.cse.ru',
        'Referer': 'https://www.cse.ru/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36'
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()
        
        found_items = data.get('found', [])
        if not found_items:
            return {"Накладная": number, "Статус": "Данные не найдены", "Дата": ""}
            
        track_data = found_items[0]
        
        state = track_data.get('State', 'Статус не указан')
        info = track_data.get('Info', '')
        
        if info:
            current_status = f"{state}. {info}"
        else:
            current_status = state

        delivery_date = ""

        if track_data.get('StateEn') == "Delivery completed" or "Доставка завершена" in state:
            current_status = "Доставка завершена"
            
            history_dict = track_data.get('History', {})
            waybill_info = history_dict.get('waybill_info', [])
            
            for event in waybill_info:
                event_name = event.get('EventName', '')
                event_name_en = event.get('EventNameEn', '')
                if "Доставка завершена" in event_name or event_name_en == "Delivery completed":
                    delivery_date = event.get('EventDate', '')
                    break
        
        return {"Накладная": number, "Статус": current_status, "Дата": delivery_date}

    except requests.exceptions.JSONDecodeError:
        return {"Накладная": number, "Статус": "Ошибка ответа сервера", "Дата": ""}
    except Exception as e:
        return {"Накладная": number, "Статус": "Ошибка обработки", "Дата": ""}

def main():
    st.title("📦 Массовая проверка накладных CSE")
    st.write("Загрузите Excel-файл с номерами накладных в первом столбце. Скрипт сохранит исходный порядок и количество строк.")

    # Виджет загрузки файла
    uploaded_file = st.file_uploader("Выберите файл .xlsx", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            # Читаем загруженный файл
            df_input = pd.read_excel(uploaded_file)
            original_tracking_list = df_input.iloc[:, 0].tolist()
            
            # Формируем исходный список точь-в-точь как в файле
            original_list = [str(num).strip() if pd.notna(num) else "" for num in original_tracking_list]
            
            # Выделяем только уникальные и непустые номера
            unique_numbers = list(dict.fromkeys([num for num in original_list if num]))
            
            st.info(f"Всего строк в файле: **{len(original_list)}** | Уникальных номеров для API: **{len(unique_numbers)}**")
            
            # Кнопка запуска
            if st.button("Начать проверку", type="primary"):
                start_time = time.time()
                
                # Прогресс-бар в интерфейсе
                my_bar = st.progress(0, text="Идет опрос API CSE. Пожалуйста, подождите...")
                
                unique_results = []
                
                # Запускаем потоки для уникальных номеров
                with ThreadPoolExecutor(max_workers=10) as executor:
                    # Запускаем все задачи
                    futures = [executor.submit(process_single_number, num) for num in unique_numbers]
                    
                    # Обновляем прогресс-бар по мере завершения каждого потока
                    for i, future in enumerate(as_completed(futures)):
                        unique_results.append(future.result())
                        progress = (i + 1) / len(unique_numbers)
                        my_bar.progress(progress, text=f"Обработано {i + 1} из {len(unique_numbers)} уникальных номеров...")
                
                # Создаем словарь для сопоставления
                results_dict = {res["Накладная"]: res for res in unique_results}
                
                # Восстанавливаем финальный список под размер исходного файла
                final_results = []
                for num in original_list:
                    if num == "":
                        final_results.append({"Накладная": "", "Статус": "Пустая строка", "Дата": ""})
                    else:
                        final_results.append(results_dict.get(num, {"Накладная": num, "Статус": "Неизвестная ошибка", "Дата": ""}))
                
                elapsed_time = round(time.time() - start_time, 1)
                st.success(f"✅ Проверка завершена за {elapsed_time} сек.")
                
                # Показываем результат на экране
                df_output = pd.DataFrame(final_results)
                st.dataframe(df_output, use_container_width=True)
                
                # Подготавливаем файл для скачивания (в памяти, не засоряя сервер)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_output.to_excel(writer, index=False, sheet_name='Результат')
                processed_data = output.getvalue()
                
                current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                original_name = uploaded_file.name.rsplit('.', 1)[0]
                download_name = f"{original_name}_Результат_{current_time}.xlsx"
                
                # Кнопка для скачивания готового Excel
                st.download_button(
                    label="📥 Скачать результат (Excel)",
                    data=processed_data,
                    file_name=download_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"Произошла ошибка при обработке файла: {e}")

if __name__ == "__main__":
    main()
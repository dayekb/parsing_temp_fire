import pandas as pd
import glob
import re

def main():
    """Объединяет Excel файлы в CSV - версия 3 для Курганской области"""
    
    excel_files = glob.glob('*.xls') + glob.glob('*.xlsx')
    print(f"Найдено файлов: {len(excel_files)}")
    
    all_data = []
    
    for file in excel_files:
        print(f"\nОбрабатываю: {file}")
        
        # Читаем файл
        df = pd.read_excel(file)
        print(f"Размер: {df.shape}")
        
        # Ищем дату
        date = None
        month_names = {
            'января': '01', 'февраля': '02', 'марта': '03', 'апреля': '04',
            'мая': '05', 'июня': '06', 'июля': '07', 'августа': '08',
            'сентября': '09', 'октября': '10', 'ноября': '11', 'декабря': '12'
        }
        
        for i in range(len(df)):
            for col_idx in range(len(df.columns)):
                cell = str(df.iloc[i, col_idx])
                if 'за' in cell and 'г.' in cell:
                    # Ищем паттерн "за X месяца Y г."
                    match = re.search(r'за\s+(\d{1,2})\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})\s+г\.', cell, re.IGNORECASE)
                    if match:
                        day = match.group(1).zfill(2)
                        month_name = match.group(2).lower()
                        year = match.group(3)
                        month = month_names[month_name]
                        date = f"{year}-{month}-{day}"
                        break
            if date:
                break
        
        print(f"Дата: {date}")
        
        # Берем данные с 8 строки (данные начинаются с строки 8)
        data = df.iloc[8:].copy()
        data = data.dropna(how='all')
        
        # Удаляем служебные строки и комментарии
        clean_data = []
        for i in range(len(data)):
            row = data.iloc[i]
            first_cell = str(row.iloc[0])
            
            # Проверяем, что это не служебная строка
            if not any(word in first_cell.lower() for word in [
                'среднее', 'итого', 'примечание', 'примечания', 'таблица содержит',
                'метеорологические данные', 'гидрометцентр', 'показатели пожароопасности',
                'рассчитанные на основании', 'утвержденным методикам', '-------',
                '............', 'курганская', 'челябинская', 'кусинская'
            ]):
                # Проверяем, что строка содержит реальные данные (есть числовые значения)
                has_numeric_data = False
                for col_idx in range(2, min(8, len(row))):  # Проверяем колонки с данными
                    cell_value = str(row.iloc[col_idx])
                    if cell_value.replace('.', '').replace('-', '').isdigit():
                        has_numeric_data = True
                        break
                
                if has_numeric_data:
                    clean_data.append(row)
        
        if clean_data:
            clean_df = pd.DataFrame(clean_data)
            
            # Создаем правильные заголовки
            headers = ['Авиаотделение', 'Метеостанция', 'Температура', 'Точка_росы', 
                      'Время_измерения', 'Осадки', 'Высота_снега', 'КППО_1', 'Класс_ПО_1',
                      'КППО_2', 'Класс_ПО_2', 'КППО_3', 'Класс_ПО_3', 'Пред_день', 'Ночь', 'Признак_снега']
            
            # Устанавливаем заголовки
            clean_df.columns = headers[:len(clean_df.columns)]
            
            # Добавляем колонку с датой в начало
            clean_df.insert(0, 'Дата', date)
            
            all_data.append(clean_df)
            print(f"Добавлено строк: {len(clean_df)}")
        else:
            print(f"Предупреждение: файл {file} не содержит данных")
    
    if all_data:
        # Объединяем
        result = pd.concat(all_data, ignore_index=True)
        
        # Сохраняем
        result.to_csv('combined_meteo_data_v3.csv', index=False, encoding='utf-8-sig')
        print(f"\nГотово! Строк: {len(result)}, колонок: {len(result.columns)}")
        
        # Показываем заголовки
        print(f"\nЗаголовки колонок:")
        for i, col in enumerate(result.columns):
            print(f"  {i+1}. {col}")
        
        # Показываем первые строки
        print(f"\nПервые 3 строки:")
        print(result.head(3))
        
        # Показываем статистику по датам
        print(f"\nСтатистика по датам:")
        date_counts = result['Дата'].value_counts().sort_index()
        for date, count in date_counts.items():
            print(f"  {date}: {count} строк")
        
        # Показываем статистику по авиаотделениям
        print(f"\nСтатистика по авиаотделениям:")
        avia_counts = result['Авиаотделение'].value_counts()
        for avia, count in avia_counts.items():
            if pd.notna(avia) and avia.strip():
                print(f"  {avia}: {count} строк")
        
    else:
        print("Нет данных для объединения!")

if __name__ == "__main__":
    main()

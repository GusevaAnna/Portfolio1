
import pandas as pd
import matplotlib.pyplot as plt

# Загружаем данные из CSV файла
try:
    data = pd.read_csv('sales_data.csv', delimiter=';', encoding='utf-8')
except UnicodeDecodeError:
    data = pd.read_csv('sales_data.csv', delimiter=';', encoding='cp1251')

# Проверяем, были ли данные загружены правильно
if len(data.columns) == 1:
    data = data.iloc[:, 0].str.split(',', expand=True)

# Исправляем названия столбцов
data.columns = [
    'Дата продажи',
    'ID заказа',
    'ID клиента',
    'Наименование товара',
    'Категория товара',
    'Количество',
    'Цена за единицу'
]

# Преобразуем столбцы в нужные типы данных
data['Дата продажи'] = pd.to_datetime(data['Дата продажи'])
data['Количество'] = data['Количество'].astype(int)
data['Цена за единицу'] = data['Цена за единицу'].astype(int)

# Рассчитываем ежедневную выручку
data['Выручка'] = data['Количество'] * data['Цена за единицу']
daily_revenue = data.groupby('Дата продажи')['Выручка'].sum().reset_index()

# Определяем день с максимальной выручкой
max_revenue_row = daily_revenue.loc[daily_revenue['Выручка'].idxmax()]
max_date = max_revenue_row['Дата продажи']
max_revenue = max_revenue_row['Выручка']

# Выводим результат
print(f"День с максимальной выручкой: {{max_date.date()}}, Выручка: {{max_revenue}} руб.")

# Визуализация
plt.figure(figsize=(10, 5))
plt.plot(daily_revenue['Дата продажи'], daily_revenue['Выручка'], marker='o', color='blue')
plt.title('Ежедневная выручка')
plt.xlabel('Дата')
plt.ylabel('Выручка (руб.)')
plt.grid()
plt.xticks(rotation=45)
plt.tight_layout()

# Сохраняем график в файл
plt.savefig('revenue_plot.png')
plt.show()

# Сохранение результатов в Excel с использованием XlsxWriter
with pd.ExcelWriter('sales_analysis.xlsx', engine='xlsxwriter') as writer:
    daily_revenue.to_excel(writer, sheet_name='Daily Revenue', index=False)

    # Получаем доступ к объекту книги и листа
    workbook = writer.book
    worksheet = writer.sheets['Daily Revenue']

    # Вставляем график в Excel
    worksheet.insert_image('H2', 'revenue_plot.png')  # Вставляем график

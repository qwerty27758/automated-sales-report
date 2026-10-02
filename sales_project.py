import requests
import pandas as pd
import time
import os

from dotenv import load_dotenv
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")

def create_report():
    start_time = time.time()

    def send_telegram_message(message):
        url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"

        response = requests.post(
            url, json={"chat_id": CHAT_ID, "text": message}, timeout=10
        )

        response.raise_for_status()
    
    # Загружаем данные о продажах из CSV
    df = pd.read_csv("sales.csv")
 
    df["Выручка"] = df["Количество"] * df["Цена"]
    df["Расходы"] = df["Количество"] * df["Себестоимость"]
    df["Прибыль"] = df["Выручка"] - df["Расходы"]
    df["Маржа"] = df["Прибыль"] / df["Выручка"]

    manager_summary = (
    df.groupby("Менеджер")
      .agg(
          Заказов=("Товар", "count"),
          Выручка=("Выручка", "sum"),
          Расходы=("Расходы", "sum"),
          Прибыль=("Прибыль", "sum")
      )
      .reset_index()
    )

    manager_summary["Маржа"] = (
        manager_summary["Прибыль"] / manager_summary["Выручка"]
    )

    # Находим заказы с убытками
    loss_orders = df[df["Прибыль"] < 0]

    # Сортируем заказы с убытками по возрастанию
    loss_orders_sorted = loss_orders.sort_values(by="Прибыль", ascending=True)

    # Метод loss_orders_sorted.iloc[0] возвращает первую строку, которая соответствует заказу с наибольшим убытком
    if not loss_orders_sorted.empty:
        worst_order = loss_orders_sorted.iloc[0]
    else:
        worst_order = None

    if worst_order is not None:
        print("Самый убыточный заказ:")
        print(f"Клиент: {worst_order['Клиент']}")
        print(f"Товар: {worst_order['Товар']}")
        print(f"Выручка: {worst_order['Выручка']}")
        print(f"Расходы: {worst_order['Расходы']}")
        print(f"Прибыль: {worst_order['Прибыль']}")
        print(f"Маржа: {worst_order['Маржа']:.2%}", "\n")
    else:
        print("Убыточных заказов нет.", "\n")

    total_revenue = df["Выручка"].sum()
    total_expenses = df["Расходы"].sum()
    total_profit = df["Прибыль"].sum()
    average_margin = df["Маржа"].mean()  # .mean() используется для вычисления среднего арифметического значения по столбцу
    total_margin = total_profit / total_revenue  # Общая маржа всех заказов.

    print("Общая выручка:", total_revenue)
    print("Общие расходы:", total_expenses)
    print("Общая прибыль:", total_profit)
    print("Средняя маржа:", f"{average_margin:.2%}"
    )  # f"{переменная:.2% округляет значение до двух знаков после запятой и добавляет знак процента
    print("Общая маржа:", f"{total_margin:.2%}")

    # Создаем DataFrame для итогов
    summary = pd.DataFrame(
        {
            "Показатель": [
                "Общая выручка",
                "Общие расходы",
                "Общая прибыль",
                "Общая маржа",
                "Количество убыточных заказов",
            ],
            "Значение": [
                total_revenue,
                total_expenses,
                total_profit,
                total_margin,
                len(loss_orders),
            ],
        }
    )

    # Сохраняем результаты в Excel с несколькими листами
    with pd.ExcelWriter("sales_report.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Все заказы", index=False)
        loss_orders_sorted.to_excel(writer, sheet_name="Убыточные заказы", index=False)
        summary.to_excel(writer, sheet_name="Итоги", index=False)

        sheet = writer.book["Итоги"]

        # Формат чисел
        sheet["B2"].number_format = '#,##0" ₽"'
        sheet["B3"].number_format = '#,##0" ₽"'
        sheet["B4"].number_format = '#,##0" ₽"'
        sheet["B5"].number_format = "0.00%"
        sheet["B6"].number_format = "0"

        # Ширина столбцов
        sheet.column_dimensions["A"].width = 32
        sheet.column_dimensions["B"].width = 20

        # Оформление заголовков
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

        loss_sheet = writer.book["Убыточные заказы"]

        # Ширина столбцов
        for column in loss_sheet.columns:
            column_letter = column[0].column_letter

            # Ограничиваем ширину столбца
            loss_sheet.column_dimensions[column_letter].width = 20

            # Включаем перенос текста
            for cell in column:
                cell.alignment = Alignment(wrap_text=True)
                loss_sheet.row_dimensions[cell.row].height = 45

        # Жирный заголовок
        for cell in loss_sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

        # Форматируем данные на листе "Убыточные заказы"
        for row in range(2, loss_sheet.max_row + 1):
            loss_sheet[f"E{row}"].number_format = '#,##0" ₽"'
            loss_sheet[f"F{row}"].number_format = '#,##0" ₽"'
            loss_sheet[f"I{row}"].number_format = '#,##0" ₽"'
            loss_sheet[f"J{row}"].number_format = '#,##0" ₽"'
            loss_sheet[f"K{row}"].number_format = '#,##0" ₽"'
            loss_sheet[f"L{row}"].number_format = '0.00%'

        # Выделяем отрицательную прибыль
        # Создаем заливку ячеек
        negative_fill = PatternFill(
            fill_type="solid", fgColor="FFC7CE"
        )  # Проходим по всем строкам
        for row in range(2, loss_sheet.max_row + 1):
            profit_cell = loss_sheet[f"K{row}"]
            # Если прибыль отрицательная - применяется заливка
            if profit_cell.value < 0:
                profit_cell.fill = negative_fill

        # Выделяем убыточные заказы на листе "Все заказы"
        all_sheet = writer.book["Все заказы"]

        for row in range(2, all_sheet.max_row + 1):
            profit_cell = all_sheet[f"K{row}"]

            if profit_cell.value < 0:
                profit_cell.fill = negative_fill

        for column in all_sheet.columns:
            column_letter = column[0].column_letter

            # Ограничиваем ширину столбца
            all_sheet.column_dimensions[column_letter].width = 20

            # Включаем перенос текста
            for cell in column:
                cell.alignment = Alignment(wrap_text=True)
                all_sheet.row_dimensions[cell.row].height = 45

        # Жирный заголовок
        for cell in loss_sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
        for row in range(2, all_sheet.max_row + 1):
            # Цена
            all_sheet[f"E{row}"].number_format = '#,##0" ₽"'

            # Себестоимость
            all_sheet[f"F{row}"].number_format = '#,##0" ₽"'

            # Выручка
            all_sheet[f"I{row}"].number_format = '#,##0" ₽"'

            # Расходы
            all_sheet[f"J{row}"].number_format = '#,##0" ₽"'

            # Прибыль
            all_sheet[f"K{row}"].number_format = '#,##0" ₽"'

            # Маржа
            all_sheet[f"L{row}"].number_format = '0.00%'

        # Оформление заголовков
        for cell in all_sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

        # Создаём лист для дашборда
        dashboard = writer.book.create_sheet("Дашборд")

        # Создаём скрытый лист с техническими данными для графиков
        chart_data = writer.book.create_sheet("Данные для графиков")
        chart_data.sheet_state = "hidden"

        # Убираем сетку
        dashboard.sheet_view.showGridLines = False

        # Заголовок
        dashboard["A1"] = "АНАЛИТИЧЕСКИЙ ДАШБОРД"
        dashboard["A1"].font = Font(bold=True, size=18)
        dashboard.row_dimensions[1].height = 30

        # KPI-блоки
        dashboard["A3"] = "Общая выручка"
        dashboard["A4"] = total_revenue

        dashboard["D3"] = "Общие расходы"
        dashboard["D4"] = total_expenses

        dashboard["G3"] = "Общая прибыль"
        dashboard["G4"] = total_profit

        dashboard["J3"] = "Общая маржа"
        dashboard["J4"] = total_margin

        dashboard["M3"] = "Убыточных заказов"
        dashboard["M4"] = len(loss_orders)

        # Формат денег
        for cell in ["A4", "D4", "G4"]:
            dashboard[cell].number_format = '#,##0" ₽"'

        # Формат маржи
        dashboard["J4"].number_format = '0.00%'
        dashboard["M4"].number_format = "0"

        # Формат шрифтов
        for cell in ["A3", "D3", "G3", "J3", "M3"]:
            dashboard[cell].font = Font(bold=True, size=11)

        for cell in ["A4", "D4", "G4", "J4", "M4"]:
            dashboard[cell].font = Font(bold=True, size=14)

        # Рамки KPI
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        for cell in ["A3", "A4", "D3", "D4", "G3", "G4", "J3", "J4", "M3", "M4"]:
            dashboard[cell].border = thin_border

        # Размеры строк KPI
        dashboard.row_dimensions[3].height = 22
        dashboard.row_dimensions[4].height = 28

        # Оформление KPI-блоков
        kpi_headers = ["A3", "D3", "G3", "J3", "M3"]
        kpi_values = ["A4", "D4", "G4", "J4", "M4"]

        header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

        for cell in kpi_headers:
            dashboard[cell].fill = header_fill
            dashboard[cell].alignment = Alignment(
                horizontal="center", vertical="center"
            )  # центрируем текст по горизонтали и вертикали

        for cell in kpi_values:
            dashboard[cell].alignment = Alignment(
                horizontal="center", vertical="center"
            )

        # Ширина столбцов
        for column in ["A", "D", "G", "J"]:
            dashboard.column_dimensions[column].width = 20

        for column in ["B", "E", "H", "K"]:
            dashboard.column_dimensions[column].width = 18
            dashboard.column_dimensions["M"].width = 20
            from openpyxl.chart import BarChart, Reference

        # График прибыли по заказам
        # Создаем столбчатую диаграмму
        chart = BarChart()

        chart.title = "Прибыль по заказам"
        chart.y_axis.title = "Прибыль, ₽"
        chart.x_axis.title = "Клиент"

        data = Reference(
            all_sheet,
            min_col=11,  # берём 11-й столбец J, то есть прибыль
            min_row=1,
            max_row=all_sheet.max_row,
        )

        categories = Reference(
            all_sheet,
            min_col=2,  # берём столбец B для горизонтальной оси
            min_row=2,
            max_row=all_sheet.max_row,
        )

        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)

        # Убираем легенду, так как она не нужна
        chart.legend = None

        chart.height = 12
        chart.width = 21

        # размещает график на дашборде начиная с ячейки A10
        dashboard.add_chart(chart, "A10")

        # Второй график — выручка и расходы по заказам
        chart2 = BarChart()

        chart2.type = "col"
        chart2.title = "Выручка и расходы по заказам"
        chart2.y_axis.title = "Сумма, ₽"
        chart2.x_axis.title = "Клиент"

        # Данные по выручке
        revenue_data = Reference(
             all_sheet, min_col=9, min_row=1, max_row=all_sheet.max_row
        )

        # Данные по расходам
        expenses_data = Reference(
            all_sheet, min_col=10, min_row=1, max_row=all_sheet.max_row
        )

        # Добавляем обе серии данных
        chart2.add_data(revenue_data, titles_from_data=True)

        chart2.add_data(expenses_data, titles_from_data=True)

        # ID заказов по оси X
        chart2.set_categories(categories)

        # Размер графика
        chart2.height = 12
        chart2.width = 21

        # Размещаем на дашборде
        dashboard.add_chart(chart2, "H10")

        # Создаем круговую диаграмму
        from openpyxl.chart import PieChart

       # Технические данные для круговой диаграммы
        chart_data["A1"] = "Тип заказа"
        chart_data["B1"] = "Количество"

        chart_data["A2"] = "Прибыльные"
        chart_data["B2"] = (df["Прибыль"] >= 0).sum()

        chart_data["A3"] = "Убыточные"
        chart_data["B3"] = (df["Прибыль"] < 0).sum()


        pie = PieChart()

        pie.title = "Прибыльные и убыточные заказы"

        data = Reference(
            chart_data, min_col=2, min_row=1, max_row=3
        )   # B — Количество

        labels = Reference(
            chart_data, min_col=1, min_row=2, max_row=3
        )  # A — Тип заказа

        pie.add_data(data, titles_from_data=True)
        pie.set_categories(labels)

        pie.height = 12
        pie.width = 18

        dashboard.add_chart(pie, "A36")

    # Отправка отчета в Telegram
    if worst_order is not None:
        worst_order_text = (
            f"🔻 Самый убыточный заказ:\n"
            f"Клиент: {worst_order['Клиент']}\n"
            f"Товар: {worst_order['Товар']}\n"
            f"Выручка: {worst_order['Выручка']:.2f} ₽\n"
            f"Расходы: {worst_order['Расходы']:.2f} ₽\n"
            f"Прибыль: {worst_order['Прибыль']:.2f} ₽\n"
            f"Маржа: {worst_order['Маржа']:.2%}\n"
        )
    else:
        worst_order_text = "🔻 Убыточных заказов нет.\n"


    telegram_message = (
        f"📊 Отчёт по продажам\n\n"
        f"Количество заказов: {len(df)}\n"
        f"Общая выручка: {total_revenue:.2f} ₽\n"
        f"Общие расходы: {total_expenses:.2f} ₽\n"
        f"Общая прибыль: {total_profit:.2f} ₽\n"
        f"Средняя маржа: {average_margin:.2%}\n\n"
        f"{worst_order_text}"
    )

    send_telegram_message(telegram_message)

    with open("sales_report.xlsx", "rb") as file:
        response = requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument",
            data={"chat_id": CHAT_ID},
            files={"document": file},
            timeout=30,
        )

    response.raise_for_status()

    print(f"\nВремя выполнения: {time.time() - start_time:.2f} сек.")


create_report()

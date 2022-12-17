import time
from openpyxl import load_workbook


def add_additional_price(start_price):
    if start_price < 800:
        percent = 0.55
    elif 800 <= start_price < 1200:
        percent = 0.5
    elif 1200 <= start_price < 1500:
        percent = 0.45
    elif 1500 <= start_price < 2000:
        percent = 0.4
    elif 2000 <= start_price < 3000:
        percent = 0.35
    elif 3000 <= start_price < 5000:
        percent = 0.3
    else:
        percent = 0.25
    additional_price = start_price * percent
    final_price = start_price + additional_price
    return final_price


def get_first_data():
    first_data = dict()
    workbook = load_workbook("1 документ.xlsx")
    current_page = workbook.active
    for index in range(13, current_page.max_row + 1):
        product_id = current_page[f"J{index}"].value
        start_price = current_page[f"O{index}"].value
        price = add_additional_price(start_price=start_price)
        first_data[product_id] = price
    return first_data


def check_data(first_data):
    first_product_ids = list(first_data.keys())
    workbook = load_workbook("2 документ.xlsx")
    current_page = workbook.active
    for index in range(2, current_page.max_row + 1):
        product_id = current_page[f"B{index}"].value
        if product_id in first_product_ids:
            current_page[f"C{index}"].value = first_data[product_id]
            current_page[f"I{index}"].value = "Активно"
            print(f"[INFO] Статус товара с артикулом {product_id} - Активно")
        else:
            current_page[f"C{index}"].value = "-"
            current_page[f"I{index}"].value = "В архиве"
            print(f"[INFO] Статус товара с артикулом {product_id} - В архиве")
    workbook.save("2 документ.xlsx")


def main():
    start_time = time.time()
    print("[INFO] Программа запущена")
    print("[INFO] Получаем данные из первой таблицы")
    first_data = get_first_data()
    print("[INFO] Данные из первой таблицы получены")
    print("[INFO] Сравниваем данные первой и второй таблицы")
    check_data(first_data=first_data)
    print("[INFO] Данные обработаны и перезаписаны")
    print("[INFO] Программа успешно завершена")
    stop_time = time.time()
    print(f"[INFO] Время на работы программы: {stop_time - start_time}")


if __name__ == "__main__":
    main()

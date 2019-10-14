from rutermextract import TermExtractor  # библиотека для извлечения ключевых слов на русском языке
import xlwt  # библиотека для работы с Excel
import os


#  функция для вывода данных в Excel
def output(term_extract, data):
    book = xlwt.Workbook(encoding="utf-8")  # создание файла
    sheet = book.add_sheet("OutputData")  # создание страницы
    sheet.write(0, 0, "Ключевое слово")
    sheet.write(0, 1, "Количество повторений")
    tmp = 1
    for term in term_extract(data):
        sheet.write(tmp, 0, term.normalized)  # первый столбец - ключевое слово
        sheet.write(tmp, 1, term.count)  # второй стобец - количество повторений
        tmp += 1
    book.save("OutputData.xls")


def main():
    file_path = str(input("Введите путь к текстовому файлу по следующему формату C:\\Users...\\FileName.txt:\n"))
    if not os.path.exists(file_path):
        print("Указанный файл не существует")
    else:
        with open(file_path, "r") as file:
            content = file.read()  # считывание содержимого файла
        term_extractor = TermExtractor()  # использование библиотеки rutermextract - деление текста на слова
        # приведение в нормальную форму, вычисление ключевых слов
        output(term_extractor, content)  # вывод данных


if __name__ == "__main__":
    main()

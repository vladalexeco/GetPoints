# Get Points
Приложение для операционной системы Windows, позволяющее осуществлять различные манипуляции с наборомами данных, предназначеннми для работы с географическими координатами.
Программа позволяет работать с файлами, полученными из различных GPS навигаторов, а также с файлами системы автоматизированного проектирования Autocad. Основная функция приложения
это возможность удобного конверирования (с учетом существующих настроек) одного формата в другой, а также подготовки данных к формированию технического отчета или протокола.
### Общий вид рабочей области программы
![image](https://github.com/vladalexeco/GetPoints/assets/27238541/b0186523-b30f-4bde-9dc8-00c0807ab898)
### Окно настроек программы
![image](https://github.com/vladalexeco/GetPoints/assets/27238541/34003703-3dda-4a00-a1c3-55a0e30d6d08)
# Инструкция по эксплуатации
Основной экран приложения состоит из двух основных блоков:
- таблица с рабочими файлами
- панель управления
### Таблица с рабочими файлами
Само окно таблицы состоим из двух столбцов: столбец с наименованием файлов, столбец с датой и временем создания файлов. Чтобы начать работу, нужно перетащить один или несколько
файлов (drag and drop) в рабочую область. Программа имеет возможность работать с такими форматами как: gpx, kml, txt, xls. Чтобы очистить таблицу, нужно нажать на клавишу del 
на клавиатуре.
### Панель управления
Панель представляет из себя набор функциональных кнопок:
- Добавить
- Удалить
- В Excel
- МСК
- GPX to KML
- KML to GPX
- Компоновка
- Параметры
### Добавить 
Клавиша открывает окно Windows для добавления файла
### Удалить
Клавиша удаляет отмеченный файл из таблицы
### В Excel
При нажатии на клавишу формируется сводная таблица с данными на основе ранее добавленых в таблицу файлов формата gpx и kml
### МСК
При нажатии на клавишу происходит конвертация файлов, ранее занесенных в таблицу, в документ (формат txt), содержащий в себе координаты в формате местных систем координат (МСК).
Выбор требуемых параметров МСК согласно региону можно выбрать из выпадающего меню в верхней части окна приложения.
### GPX to KML
При нажатии на клавишу происходит конвертация файлов gpx в формат kml
### KML to GPX
При нажатии на клавишу происходит конвертация файлов kml в формат gpx
### В журнал
При нажатии на клавишу происходит конвертирование журнала гамма-съемки в формате xls в gpx формат
### МСК в Excel
При нажатии на клавишу происходит конвертирование файла xls с координатами МСК в gpx формат
### Компоновка
При нажатии на клавишу происходит создание файла xls, который представляет собой компоновку двух ранее созданных файлов: документ с географическими координатами
и журнал с результатами измерений
### Параметры
При нажатии на клавишу открывается дополнительное окно с параметрами, c помощью которых можно произвести ряд настроек, таких как: порядок вывода данных, формат вывода
данных и т.д.
# Системные требования
Python 3, а также установленные библиотеки:
- pyinstaller==5.9.0
- pyinstaller-hooks-contrib==2023.1
- pyproj==2.1.3
- PyQt5==5.11.3
- PyQt5_sip==4.19.19
- xlrd==1.2.0
- xlwt==1.3.0
- lxml==4.4.1

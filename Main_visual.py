from PyQt5.QtWidgets import (QLineEdit, QPushButton, QLabel, QWidget, QApplication, QGridLayout,
                             QFileDialog, QMessageBox)
from PyQt5.QtGui import QFont, QPixmap, QIcon
from PyQt5.QtCore import QDateTime
import openpyxl
from openpyxl import load_workbook
import time
import datetime

class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()
        self.setWindowTitle("Тестовая программа: Аналитика v1.0")
        self.main_maket = QGridLayout()
        self.main_maket.addWidget(self.hello_label())
        self.main_maket.addWidget(self.book_ost_label(), 1, 0)
        self.main_maket.addWidget(self.book_ost_line_edit(), 1, 1)
        self.main_maket.addWidget(self.button_browseFiles_book_ost(), 1, 2)
        self.main_maket.addWidget(self.book_label(), 2, 0)
        self.main_maket.addWidget(self.book_line_edit(), 2, 1)
        self.main_maket.addWidget(self.button_browseFiles_book(), 2, 2)
        self.main_maket.addWidget(self.book_price_label(), 3, 0)
        self.main_maket.addWidget(self.book_price_line_edit(), 3, 1)
        self.main_maket.addWidget(self.button_browseFiles_book_price(), 3, 2)
        self.main_maket.addWidget(self.book_const_label(), 4, 0)
        self.main_maket.addWidget(self.book_const_line_edit(), 4, 1)
        self.main_maket.addWidget(self.button_browseFiles_book_const(), 4, 2)
        self.main_maket.addWidget(self.b_create(), 5, 0)
        self.main_maket.addWidget(self.b_close(), 5, 1)
        self.setLayout(self.main_maket)

    def hello_label(self):
        hello_label = QLabel("Здравствуйте, ув. пользователь!")
        return hello_label

    def book_label(self):
        book_label = QLabel("Выберите файл, содержащий остатки из 1С:")
        return book_label

    def book_line_edit(self):
        self.book_line_edit = QLineEdit(self)
        self.book_line_edit.setMinimumSize(30, 30)
        serchOtherIcon = self.book_line_edit.addAction(QIcon("search.png"), QLineEdit.TrailingPosition)
        serchOtherIcon.triggered.connect(self.browse_files_book)
        return self.book_line_edit

    def book_ost_label(self):
        book_ost_label = QLabel("Выберите файл, содержащий остатки из базы:")
        return book_ost_label

    def book_ost_line_edit(self):
        self.book_ost_line_edit = QLineEdit(self)
        self.book_ost_line_edit.setMinimumSize(30, 30)
        return self.book_ost_line_edit

    def book_price_label(self):
        book_price = QLabel("Выберите файл, содержащий прайс:")
        return book_price

    def book_price_line_edit(self):
        self.book_price_line_edit = QLineEdit(self)
        self.book_price_line_edit.setMinimumSize(30, 30)
        return self.book_price_line_edit

    def book_const_label(self):
        book_const = QLabel("Выберите файл, содержащий константу:")
        return book_const

    def book_const_line_edit(self):
        self.book_const_line_edit = QLineEdit(self)
        self.book_const_line_edit.setMinimumSize(30, 30)
        return self.book_const_line_edit

    def button_browseFiles_book(self):
        button_browseFiles_book = QPushButton("Обзор")
        button_browseFiles_book.setMinimumSize(30, 30)
        button_browseFiles_book.clicked.connect(self.browse_files_book)
        return button_browseFiles_book

    def browse_files_book(self):
        browse_files_book = QFileDialog.getOpenFileName(self, 'Open file', 'C:\Program Files', 'xlsx files (*.xlsx)')
        self.book_line_edit.setText(browse_files_book[0])

    def button_browseFiles_book_ost(self):
        button_browseFiles_book_ost = QPushButton("Обзор")
        button_browseFiles_book_ost.setMinimumSize(30, 30)
        button_browseFiles_book_ost.clicked.connect(self.browse_files_book_ost)
        return button_browseFiles_book_ost

    def browse_files_book_ost(self):
        browse_files_book_ost = QFileDialog.getOpenFileName(self, 'Open file', 'C:\Program Files', 'xlsx files (*.xlsx)')
        self.book_ost_line_edit.setText(browse_files_book_ost[0])

    def button_browseFiles_book_price(self):
        button_browseFiles_book_price = QPushButton("Обзор")
        button_browseFiles_book_price.setMinimumSize(30, 30)
        button_browseFiles_book_price.clicked.connect(self.browse_files_book_price)
        return button_browseFiles_book_price

    def browse_files_book_price(self):
        browse_files_book_price = QFileDialog.getOpenFileName(self, 'Open file', 'C:\Program Files', 'xlsx files (*.xlsx)')
        self.book_price_line_edit.setText(browse_files_book_price[0])

    def button_browseFiles_book_const(self):
        button_browseFiles_book_const = QPushButton("Обзор")
        button_browseFiles_book_const.setMinimumSize(30, 30)
        button_browseFiles_book_const.clicked.connect(self.browse_files_book_const)
        return button_browseFiles_book_const

    def browse_files_book_const(self):
        browse_files_book_const = QFileDialog.getOpenFileName(self, 'Open file', 'C:\Program Files', 'xlsx files (*.xlsx)')
        self.book_const_line_edit.setText(browse_files_book_const[0])

    def b_create(self):
        b_create = QPushButton("Выполнить обработку • Enter", self)
        b_create.setMinimumSize(10, 40)
        b_create.setFont(QFont('Century Gothic', 8, QFont.Normal))
        b_create.setToolTip('Нажмите, для добавления нового ярлыка')
        b_create.setStyleSheet("""
                                                   QPushButton:hover { background-color: green;
                                                   border-radius: 10px;
                                                   border-style: ridge;
                                                   border-color: dark;
                                                   border-width: 2px; }
                               QPushButton:!hover { background-color: white;
                                                   border-style: ridge;
                                                   border-width: 2px;
                                                   border-radius: 10px;
                                                   border-color: dark;   }
                               QPushButton:pressed { background-color: rgb(0, 255, 0);
                                                   border-radius: 17px;}
                                   """)

        b_create.clicked.connect(self.start_analytics)
        return b_create

    def b_close(self):
        b_close = QPushButton("Закрыть • Esc", self)
        b_close.setMinimumSize(10, 40)
        b_close.setFont(QFont('Century Gothic', 8, QFont.Normal))
        b_close.setShortcut('Esc')
        b_close.setToolTip('Нажмите, чтобы отменить изменения и выйти')
        b_close.setStyleSheet("""
                                           QPushButton:hover { background-color: rgba(139, 0, 0);
                                                               border-radius: 10px;
                                                               border-style: ridge;
                                                               border-color: dark;
                                                               border-width: 2px; }
                                           QPushButton:!hover { background-color: white;
                                                                border-style: ridge;
                                                                border-width: 2px;
                                                                border-radius: 10px;
                                                                border-color: dark;   }
                                           QPushButton:pressed { background-color: rgb(255, 0, 0);
                                                                 border-radius: 17px;}
                                       """)

        b_close.clicked.connect(self.start_close)
        return b_close

    def start_close(self):
        self.close()

    def start_analytics(self):
        self.base_line_edit = [self.book_line_edit, self.book_price_line_edit, self.book_const_line_edit, self.book_ost_line_edit]
        for line_edit in self.base_line_edit:
            if len(line_edit.text()) == 0:
                self.showMessageBox('Внимание!',
                                    '<center style=font-size:11pt><FONT FACE="Century Gothic"><b><u>Вы не заполнили поля!</center></u></b>'
                                    '<center style=font-size:7pt><FONT FACE="Century Gothic">(нажмите <b>ОК</b> и попробуйте снова)</center>')
                return
        start_time = time.time()
        self.ostatok_base_44_44dealer_46alyans()
        self.ostatok_46_46dealer_46BOSCH()
        print(f'отработла за {int(time.time() - start_time)} секунд')

    def showMessageBox(self, title, message):
        msgBox = QMessageBox()
        msgBox.setWindowTitle(title)
        msgBox.setIconPixmap(QPixmap("warning.png"))
        msgBox.setText(message)
        msgBox.setStandardButtons(QMessageBox.Ok)
        msgBox.exec_()



    def displayTime(self):
        self.b_create().setText(QDateTime.currentDateTime().toString())
        self.b_create().adjustSize()

    def ostatok_base_44_44dealer_46alyans(self):
        book = load_workbook(self.book_line_edit.text())
        book_price = load_workbook(self.book_price_line_edit.text())
        book_const = load_workbook(self.book_const_line_edit.text())
        book_rez = openpyxl.Workbook()
        book_al = openpyxl.Workbook()
        book_44_dealer = openpyxl.Workbook()
        sheet = book.active
        sheet_price = book_price.active
        sheet_44_dealer = book_44_dealer.active
        sheet_rez = book_rez.active
        sheet_al = book_al.active
        sheet_const = book_const.active

        sheet_rez.append(['Наименование', 'Производитель', 'Номер', 'Остаток', 'Кратность', 'Цена'])
        sheet_al.append(['Наименование', 'Производитель', 'Номер', 'Остаток', 'Кратность', 'Цена'])
        sheet_44_dealer.append(['Наименование', 'Производитель', 'Номер', 'Остаток', 'Кратность', 'Цена'])
        dict_ost = {}
        for row in sheet.iter_rows(min_row=11, max_row=None):
            if row[2].value != 0 and row[1].value:
                dict_ost[str(row[1].value)] = row[2].value
        dict_price = {}
        for row in sheet_price.iter_rows(min_row=2, max_row=None):
            a = dict(
                number=str(row[0].value),
                name=row[1].value,
                brand=row[2].value,
                partia=row[3].value,
                price_opt=row[4].value,
                price_uch=row[5].value,
                price_rrp=row[6].value,
            )
            dict_price[str(row[0].value)] = a

        for k, v in dict_ost.items():  # 44 и 44дилер
            k = str(k)
            if dict_price.get(k) and dict_price[k]['brand'] != 'BOSCH':
                a = [dict_price[k]['name'],
                     dict_price[k]['brand'],
                     dict_price[k]['number'], v + 2,
                     dict_price[k]['partia'],
                     dict_price[k]['price_rrp']]
                b = [dict_price[k]['name'],
                     dict_price[k]['brand'],
                     dict_price[k]['number'], v,
                     dict_price[k]['partia'],
                     dict_price[k]['price_opt']]
                sheet_44_dealer.append(a)
                sheet_rez.append(b)
            elif dict_price.get(k.rjust(10, '0')) and dict_price[k.rjust(10, '0')]['brand'] == 'BOSCH' and v <= 500:
                k = k.rjust(10, '0')
                a = [dict_price[k]['name'],
                     dict_price[k]['brand'],
                     dict_price[k]['number'], v + 2,
                     dict_price[k]['partia'],
                     dict_price[k]['price_opt']]
                b = [dict_price[k]['name'],
                     dict_price[k]['brand'],
                     dict_price[k]['number'], v,
                     dict_price[k]['partia'],
                     dict_price[k]['price_opt']]
                sheet_44_dealer.append(a)
                sheet_rez.append(b)
            elif dict_price.get(k.rjust(10, '0')) and dict_price[k.rjust(10, '0')]['brand'] == 'BOSCH':
                k = k.rjust(10, '0')
                b = [dict_price[k]['name'],
                     dict_price[k]['brand'],
                     dict_price[k]['number'], 1000,
                     dict_price[k]['partia'],
                     dict_price[k]['price_opt']]
                sheet_44_dealer.append(b)
                sheet_rez.append(b)

        for row in range(1, sheet_const.max_row):  # добаква константы
            name = sheet_const[row][0].value
            brand = sheet_const[row][1].value
            art = str(sheet_const[row][2].value)
            ost = int(sheet_const[row][3].value)
            partia = int(sheet_const[row][4].value)
            price_opt = sheet_const[row][5].value
            a = [name, brand, art, ost, partia, price_opt]
            sheet_rez.append(a)
            sheet_44_dealer.append(a)

        for k, v in dict_ost.items():  # 46 alyans
            k = str(k)
            if dict_price.get(k) and dict_price[k]['brand'] == 'SWF' or dict_price.get(k) and dict_price[k][
                'brand'] == 'MOTUL' or dict_price.get(k) and dict_price[k]['brand'] == 'MANDO' or dict_price.get(k) and \
                    dict_price[k]['brand'] == 'BSG':
                b = [dict_price[k]['name'],
                     dict_price[k]['brand'],
                     dict_price[k]['number'], v,
                     dict_price[k]['partia'],
                     dict_price[k]['price_uch']]
                sheet_al.append(b)
            elif dict_price.get(k.rjust(10, '0')) and dict_price[k.rjust(10, '0')]['brand'] == 'BOSCH':
                k = k.rjust(10, '0')
                b = [dict_price[k]['name'],
                     dict_price[k]['brand'],
                     dict_price[k]['number'], v,
                     dict_price[k]['partia'],
                     dict_price[k]['price_uch']]
                sheet_al.append(b)

        book_rez.save(f'44/Остатки {datetime.date.today().strftime("%d.%m.%y")}.xlsx')
        book_rez.save(f'ost_base.xlsx')
        book_al.save(f'46учетАльянс/Остатки {datetime.date.today().strftime("%d.%m.%y")}.xlsx')
        book_44_dealer.save(f'44дилер/Остатки {datetime.date.today().strftime("%d.%m.%y")}.xlsx')
        book.close()
        book_rez.close()
        book_al.close()
        book_price.close()
        book_44_dealer.close()

    def ostatok_46_46dealer_46BOSCH(self):
        book_ost = load_workbook(self.book_ost_line_edit.text())
        book_price = load_workbook(self.book_price_line_edit.text())
        book_46rez = openpyxl.Workbook()
        book_46_bosch_rez = openpyxl.Workbook()
        book_46_dealer_rez = openpyxl.Workbook()
        book_44_bosch_rez = openpyxl.Workbook()
        sheet_46rez = book_46rez.active
        sheet_44_bosch_rez = book_44_bosch_rez.active
        sheet_ost = book_ost.active
        sheet_price = book_price.active # доступ к листу с ценами
        sheet_46_dealer_rez = book_46_dealer_rez.active # доступ к листу с дилерами
        sheet_46_bosch_rez = book_46_bosch_rez.active # доступ к листу с резервами

        sheet_46rez.append(['Наименование', 'Производитель', 'Номер', 'Остаток', 'Кратность', 'Цена'])
        sheet_46_bosch_rez.append(['Наименование', 'Производитель', 'Номер', 'Остаток', 'Кратность', 'Цена'])
        sheet_46_dealer_rez.append(['Наименование', 'Производитель', 'Номер', 'Остаток', 'Кратность', 'Цена'])
        sheet_44_bosch_rez.append(['Наименование', 'Производитель', 'Номер', 'Остаток', 'Кратность', 'Цена'])

        dict_ost = {}
        for row in sheet_ost.iter_rows(min_row=2, max_row=None):
            a = dict(
                number=row[2].value,
                name=row[0].value,
                brand=row[1].value,
                partia=row[4].value,
                price_opt=row[5].value,
                ost=row[3].value
            )
            dict_ost[row[2].value] = a
        dict_price = {}
        for row in sheet_price.iter_rows(min_row=2, max_row=None):
            a = dict(
                number=row[0].value,
                name=row[1].value,
                brand=row[2].value,
                partia=row[3].value,
                price_opt=row[4].value,
                price_uch=row[5].value,
                price_rrp=row[6].value
            )
            dict_price[row[0].value] = a

        for v in dict_ost.values():  # 46 и 46бош и 46дилер
            if v['ost'] != 'Остаток':
                if v['brand'] != 'MOTUL':
                    d = [v['name'], v['brand'], v['number'], v['ost'], v['partia'], v['price_opt']]
                    sheet_44_bosch_rez.append(d)
                if v['brand'] == 'MOTUL':
                    a = [v['name'], v['brand'], v['number'], '36', v['partia'], v['price_opt']]
                    b = [v['name'], v['brand'], v['number'], '22', v['partia'], v['price_opt']]
                    c = [v['name'], v['brand'], v['number'], '24', v['partia'], v['price_opt']]
                    sheet_46rez.append(a)
                    sheet_46_bosch_rez.append(b)
                    sheet_46_dealer_rez.append(c)

                else:
                    if int(v['ost']) <= 8:
                        a = [v['name'], v['brand'], v['number'], '8', v['partia'], v['price_opt']]
                        b = [v['name'], v['brand'], v['number'], '6', v['partia'], v['price_opt']]
                        c = [v['name'], v['brand'], v['number'], '10', v['partia'], v['price_opt']]
                        sheet_46rez.append(a)
                        sheet_46_bosch_rez.append(b)
                        sheet_46_dealer_rez.append(c)
                    elif int(v['ost']) > 8 and int(v['ost']) <= 20:
                        a = [v['name'], v['brand'], v['number'], '20', v['partia'], v['price_opt']]
                        b = [v['name'], v['brand'], v['number'], '40', v['partia'], v['price_opt']]
                        c = [v['name'], v['brand'], v['number'], '30', v['partia'], v['price_opt']]
                        sheet_46rez.append(a)
                        sheet_46_bosch_rez.append(b)
                        sheet_46_dealer_rez.append(c)
                    elif int(v['ost']) > 20 and int(v['ost']) <= 50:
                        a = [v['name'], v['brand'], v['number'], '50', v['partia'], v['price_opt']]
                        b = [v['name'], v['brand'], v['number'], '40', v['partia'], v['price_opt']]
                        c = [v['name'], v['brand'], v['number'], '40', v['partia'], v['price_opt']]
                        sheet_46rez.append(a)
                        sheet_46_bosch_rez.append(b)
                        sheet_46_dealer_rez.append(c)
                    elif int(v['ost']) > 50 and int(v['ost']) <= 100:
                        a = [v['name'], v['brand'], v['number'], '100', v['partia'], v['price_opt']]
                        b = [v['name'], v['brand'], v['number'], '120', v['partia'], v['price_opt']]
                        c = [v['name'], v['brand'], v['number'], '90', v['partia'], v['price_opt']]
                        sheet_46rez.append(a)
                        sheet_46_bosch_rez.append(b)
                        sheet_46_dealer_rez.append(c)
                    elif int(v['ost']) > 100 and int(v['ost']) <= 500:
                        a = [v['name'], v['brand'], v['number'], '500', v['partia'], v['price_opt']]
                        b = [v['name'], v['brand'], v['number'], '400', v['partia'], v['price_opt']]
                        c = [v['name'], v['brand'], v['number'], '300', v['partia'], v['price_opt']]
                        sheet_46rez.append(a)
                        sheet_46_bosch_rez.append(b)
                        sheet_46_dealer_rez.append(c)
                    else:
                        a = [v['name'], v['brand'], v['number'], '1000', v['partia'], v['price_opt']]
                        b = [v['name'], v['brand'], v['number'], '1000', v['partia'], v['price_opt']]
                        c = [v['name'], v['brand'], v['number'], '1200', v['partia'], v['price_opt']]
                        sheet_46rez.append(a)
                        sheet_46_bosch_rez.append(b)
                        sheet_46_dealer_rez.append(c)

        book_46rez.save(f'46/Остатки {datetime.date.today().strftime("%d.%m.%y")}.xlsx')
        book_46_bosch_rez.save(f'46/BOSCH.xlsx')
        book_46_dealer_rez.save(f'46дилер/Остатки {datetime.date.today().strftime("%d.%m.%y")}.xlsx')
        book_44_bosch_rez.save(f'44/BOSCH.xlsx')
        book_46rez.close()
        book_46_bosch_rez.close()
        book_46_dealer_rez.close()
        book_44_bosch_rez.close()

if __name__ == ('__main__'):
    import sys
    app = QApplication(sys.argv)
    w = Main()
    w.show()
    sys.exit(app.exec_())